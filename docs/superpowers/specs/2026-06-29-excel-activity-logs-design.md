# Excel Activity Logs Design

Date: 2026-06-29

Status: design approved in chat, pending spec review and implementation plan.

## Context

The corporate Excel add-in `Аудиторам` should track which workbooks are opened, saved, and still active, without making Excel slower or less reliable.

An earlier centralized workbook approach is risky for roughly 100 users because many Excel instances can try to write to the same `.xlsx` at the same time. The previous logger also relied on delayed `Application.OnTime` work around close events, which can leave Excel hanging if the application exits before the timer chain finishes.

The add-in runtime must obey the existing safety rules:

- no internet, Git, shell, PowerShell, command prompt, or external hidden process calls;
- no VBA delete or move operations such as `Kill`, `RmDir`, `DeleteFile`, `DeleteFolder`, `MoveFile`, `MoveFolder`, or `Name old As new`;
- background services must fail fast and silently when paths are missing or unavailable;
- user-facing Excel must keep working even when logging cannot write.

## Goals

- Keep a historical event journal for workbook activity.
- Avoid simultaneous writes to one central workbook.
- Support about 100 users with one user per computer.
- Let the add-in write logs safely while Excel is in normal use.
- Provide an admin-friendly central workbook for viewing, filtering, deduplication, and reporting.
- Support dev testing on this machine with the same storage shape as company/work.

## Non-Goals

- Do not physically delete, move, archive, or rotate old logs from the add-in.
- Do not add visible logging controls to the normal `Аудиторам` Ribbon.
- Do not block workbook open, save, or close while waiting for the log store.
- Do not guarantee exact close timestamps; close time is inferred from heartbeat state.

## Recommended Architecture

Use per-computer append logs plus a separate central admin reader.

Each add-in runtime writes activity events to its own local spool and, when possible, appends unsynced rows to a network TSV file dedicated to that computer and month. Since each computer writes only its own file, there is no normal contention between users.

The central workbook is not part of the user runtime path. When an admin opens or refreshes it, it reads all available per-computer TSV files, imports raw events, deduplicates by `EventId`, and builds reporting tables.

This separates the two concerns:

- user add-in: collect events with minimal risk;
- central workbook: analyze history when someone intentionally opens it.

## Storage Layout

Company/work network store:

```text
NetworkAssetsRoot\_system\activity-logs\computers\
```

Expected company root from the current runbook:

```text
Z:\Project\DAVK\Оборудование\Addins\assets\_system\activity-logs\computers\
```

Dev simulation on this computer:

```text
%APPDATA%\AuditAddin\assets\_system\activity-logs\computers\
```

Local spool on each computer:

```text
%APPDATA%\AuditAddin\activity-log\spool\
```

Local sync state:

```text
%APPDATA%\AuditAddin\activity-log\state\
```

Per-computer monthly network file name:

```text
<MachineName>__YYYY-MM.tsv
```

Example:

```text
AUDIT-PC-17__2026-06.tsv
```

Because there is one user per computer, `MachineName` is the primary file partition. User fields still stay in the rows for reporting and diagnostics.

## Retention

The add-in does not delete or move old log files.

Monthly files remain in the folder until a human removes them manually. The central workbook imports all available TSV files by default. Period filtering is a viewer/reporting concern, not an enforced retention rule.

## Event Types

The event stream should use these primary event types:

- `SESSION_START`: the add-in logging service started in this Excel session.
- `OPEN`: a saved user workbook was seen open.
- `SAVE`: a user workbook was saved successfully.
- `SEEN`: heartbeat confirming a workbook was still open.
- `MISSING`: the logger previously tracked the workbook but no longer sees it.
- `SESSION_END_ATTEMPT`: Excel or the add-in began shutdown handling.
- `FLUSH_OK`: local rows were appended to the network per-computer log.
- `FLUSH_SKIP`: flush was skipped because the network store was unavailable or unsafe.

The central reports can hide technical event types by default while keeping them in `RawEvents`.

## Event Fields

Each row should be a TSV record with a stable header:

```text
EventId
EventAt
EventType
ExcelSessionId
WorkbookSessionId
WorkbookName
WorkbookFullName
WorkbookFolder
LastWriteTime
FileSize
WindowsUser
ExcelUser
MachineName
AddinVersion
HeartbeatIntervalSeconds
SourcePath
Notes
```

`EventId` must be unique and stable enough for central deduplication. A safe shape is:

```text
MachineName + ExcelSessionId + sequence number
```

`WorkbookSessionId` identifies one observed open period for one workbook. It should include the Excel session, normalized workbook path, and first observed open time.

Before writing a TSV row, all text fields must be sanitized in one helper:

- replace tabs with a single space;
- replace CR/LF line breaks with a single space;
- trim only control characters that would break the row structure;
- keep Russian text and normal file path characters intact.

This keeps the file append-only and parser-friendly without introducing a CSV quoting dialect.

## Runtime Logging Behavior

On startup:

- create an `ExcelSessionId`;
- initialize in-memory tracking;
- append `SESSION_START` to the local spool;
- schedule lightweight heartbeat only if there are trackable user workbooks.

On workbook open:

- ignore the add-in, the central viewer, log files, and unsaved workbooks with no path;
- append `OPEN` to the local spool;
- remember the workbook in memory.

On successful save:

- append `SAVE` to the local spool;
- update the in-memory last write snapshot.

On heartbeat:

- run rarely, with a default interval of 5 minutes;
- enumerate currently open user workbooks;
- append `SEEN` for tracked open workbooks;
- append `MISSING` for previously tracked workbooks that are no longer visible;
- attempt a short non-blocking flush from local spool to that computer's monthly network TSV.

On workbook close or add-in shutdown:

- do not open, create, save, or modify any central workbook;
- do not do heavy network work;
- cancel scheduled timers where possible;
- append only a local lightweight event if it can be done safely;
- allow Excel to close even if logging fails.

## Network Flush Behavior

The local spool is the durable source for rows produced by that machine. The network monthly TSV is the shared collection point for central reporting.

Flush should:

- resolve the network log folder from `NetworkAssetsRoot` in company/work;
- use the dev simulation folder when running in dev on this computer;
- create the target folder if it does not exist;
- append only rows that have not already been flushed according to local sync state;
- write only to the current computer's current month file;
- skip quickly if the folder is unavailable, the file is read-only, or the path is not configured;
- record `FLUSH_OK` or `FLUSH_SKIP` locally for diagnostics.

Flush must not:

- delete, move, rename, archive, or truncate any log file;
- open or save a central workbook;
- show `MsgBox`;
- block Excel shutdown.

## Central Admin Workbook

Create a separate admin `.xlsm` workbook for reading and reporting logs. It should not be a visible command for normal users on the `Аудиторам` Ribbon.

On open or manual refresh, it should:

- scan `NetworkAssetsRoot\_system\activity-logs\computers\` for `.tsv` files;
- import valid rows into `RawEvents`;
- deduplicate by `EventId`;
- ignore malformed rows, including a partially appended final row, and report them in `Diagnostics`;
- rebuild derived tables and reports.

The importer should treat the header as schema validation. Unknown extra columns may be ignored, but missing required columns should make that source file appear in `Diagnostics` rather than breaking refresh.

Suggested sheets:

- `Dashboard`: high-level filters and summary metrics.
- `RawEvents`: imported event stream.
- `Sessions`: one row per inferred workbook open session.
- `Files`: activity grouped by workbook path.
- `Users`: activity grouped by Windows and Excel user.
- `Computers`: freshness and event counts by machine.
- `Diagnostics`: import status, skipped files, malformed rows, duplicate counts.

## Session Calculations

The central workbook calculates approximate open duration from the event stream:

- `OpenedAt`: first `OPEN` for a `WorkbookSessionId`.
- `LastSeenAt`: last `SEEN`, `SAVE`, or `OPEN` timestamp for the session.
- `MissingAt`: first `MISSING` after the last seen state, when available.
- `ApproxClosedAt`: `MissingAt` if present, otherwise `LastSeenAt`.
- `OpenDurationMinutes`: `ApproxClosedAt - OpenedAt`.

Because heartbeat is every 5 minutes by default, close time is approximate. This is intentional: it avoids risky close-time writes.

## Reports and Macros

The central workbook should provide admin macros for:

- refresh all logs;
- show activity for a selected workbook path;
- show activity by user for a selected period;
- show files opened from a selected project folder;
- calculate open duration by workbook session;
- list files saved by multiple users;
- list very long workbook sessions;
- list computers whose logs have not updated recently;
- export filtered results to a new `.xlsx`.

Useful default views:

- "Кто открывал этот файл?"
- "Что открывал пользователь?"
- "Сколько времени книга была открыта?"
- "Последние активные книги по проекту"
- "Файлы с несколькими пользователями"
- "Диагностика установки логгера"

## Failure Handling

The runtime logger should treat logging as best-effort:

- if local spool write fails, write only to `Debug.Print` and continue;
- if network flush fails, keep local rows and retry later;
- if a TSV is being read while another Excel appends to it, the central workbook should ignore a partial final row and import it on the next refresh;
- if the network root is missing, the add-in should skip quickly;
- no normal user workflow should wait on log reporting.

## Testing Strategy

Source-level checks:

- no forbidden runtime markers in logger and viewer-support modules;
- no delete/move/rename operations in runtime logging code;
- `WorkbookBeforeClose` path does not open or save central workbooks;
- timers use `ThisWorkbook.FullName` macro targets when needed;
- activity logger is not exposed as normal user Ribbon controls.

Runtime/COM smoke checks:

- open a saved workbook and verify local `OPEN`;
- save workbook and verify local `SAVE`;
- run heartbeat and verify `SEEN`;
- close workbook and verify shutdown does not hang;
- simulate unavailable network store and verify Excel continues;
- flush to dev simulation folder and verify monthly per-computer TSV;
- open central viewer and verify import, deduplication, sessions, and diagnostics.

Scale-oriented checks:

- create sample TSV files for many computers and months;
- verify central import performance remains acceptable;
- verify duplicate `EventId` rows do not duplicate reports;
- verify malformed/partial rows go to diagnostics instead of breaking refresh.

## Open Decisions

- Exact central viewer file name and location.
- Whether the central viewer should be shipped as part of add-in assets or kept as an admin-only workbook outside the normal user package.
- Exact hidden admin procedures for enabling/disabling runtime logging during rollout.
- Whether heartbeat should be fixed at 5 minutes or configurable through a hidden admin setting.
