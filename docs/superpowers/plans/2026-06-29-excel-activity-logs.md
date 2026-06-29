# Excel Activity Logs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Реализовать безопасную систему логов Excel для надстройки `Аудиторам`: пользовательский runtime пишет per-computer TSV-журналы, отдельная admin-книга `!Журнал активности Excel.xlsm` по паролю открывается и сама строит отчеты.

**Architecture:** Runtime-надстройка не пишет в общую книгу и не делает тяжелую работу при закрытии Excel. Она пишет локальный spool, периодически дописывает строки в месячный TSV своего компьютера и использует heartbeat раз в 5 минут. Admin-книга живет отдельно, при открытии читает все TSV, дедуплицирует события и строит листы `RawEvents`, `Sessions`, `Files`, `Users`, `Computers`, `Diagnostics`, `Dashboard`.

**Tech Stack:** Excel VBA modules/classes, Excel COM build scripts in PowerShell, TSV text files through `FileSystemObject.OpenTextFile`, existing add-in build/release scripts, `addin/tests/Invoke-AddinTests.ps1`.

---

## Важные Контекстные Правила

- Перед реализацией читать и соблюдать `_publish/audit-macros-private/docs/addin-agent-instructions.md`.
- Код runtime надстройки не должен использовать `Shell`, `WScript.Shell`, PowerShell/cmd из VBA, интернет, Git, HTTP, `Kill`, `RmDir`, `DeleteFile`, `DeleteFolder`, `MoveFile`, `MoveFolder`, `Name old As new`.
- Не добавлять пользовательские кнопки логгера на вкладку `Аудиторам`.
- Не собирать company-release без прямой команды пользователя.
- После завершения изменений в надстройке: быстрые тесты, `publish-addin-release.ps1`, локальная dev-установка через `Install-AuditAddin.ps1`, commit/push dev-пакета.
- Нужен пароль открытия admin-книги. Без него можно реализовать runtime и build-script, но нельзя финально собрать `!Журнал активности Excel.xlsm`.

## File Structure

### Runtime add-in

- Modify: `_publish/audit-macros-private/addin/modAddinAssets.bas`
  - Добавить чтение `NetworkAssetsRoot` и `DeploymentMode`.
  - Добавить helpers для activity-log root:
    - `AddinAssets_GetActivityLogComputerStoreRoot`
    - `AddinAssets_GetActivityLogViewerWorkbookPath`
    - `AddinAssets_GetLocalActivityLogDevRoot`
  - Сохранить старый `AddinAssets_GetActivityLogWorkbookPath` для совместимости тестов, но новый логгер должен использовать новые helpers.

- Create: `_publish/audit-macros-private/addin/modActivityLogStore.bas`
  - Низкоуровневая append-only запись TSV.
  - Создание локального spool и sync-state.
  - Санитизация TSV-полей.
  - Генерация путей локального spool и network/dev computer store.
  - Неблокирующая выгрузка локальных строк в `<MachineName>__YYYY-MM.tsv`.
  - Никаких delete/move/rename/truncate.

- Modify: `_publish/audit-macros-private/addin/modActivityLogger.bas`
  - Перестроить orchestration вокруг нового `modActivityLogStore`.
  - Убрать старую запись в центральный workbook из runtime-пути.
  - Включить logger через hidden setting/default, но оставить возможность admin-выключения.
  - `WorkbookBeforeClose` делает только легкую локальную запись и отмену таймеров.
  - Heartbeat фиксированный: 5 минут.

- Modify: `_publish/audit-macros-private/addin/CActivityLogAppEvents.cls`
  - Оставить Application events: `WorkbookOpen`, `WorkbookAfterSave`, `WorkbookBeforeClose`.
  - Убедиться, что `WorkbookBeforeSave` не пишет тяжелые события и не нужен для центрального лога.

- Modify: `_publish/audit-macros-private/addin/build/ThisWorkbook.addin.txt`
  - Добавить отложенный старт `ActivityLogger_WorkbookOpen` через `RuntimeMacroTarget`.
  - В `Workbook_BeforeClose` отменять activity logger init timer и вызывать `StopActivityLogger`.
  - Не возвращать updater в runtime `ThisWorkbook`.

- Modify: `_publish/audit-macros-private/addin/build/build-audit-addin.ps1`
  - Подключить новый `modActivityLogStore.bas` в сборку `.xlam`.
  - Сохранять ANSI import path через существующий `New-VbeImportCopy`.

### Admin viewer workbook

- Create: `_publish/audit-macros-private/addin/admin_activity_log/modActivityLogViewer.bas`
  - Импорт TSV-файлов.
  - Проверка заголовка.
  - Игнорирование битых/частичных строк.
  - Дедупликация по `EventId`.
  - Построение листов `RawEvents`, `Sessions`, `Files`, `Users`, `Computers`, `Diagnostics`, `Dashboard`.

- Create: `_publish/audit-macros-private/addin/admin_activity_log/ThisWorkbook.activity-log-viewer.txt`
  - `Workbook_Open` вызывает `ActivityLogViewer_WorkbookOpen`.
  - Никаких кнопок обновления; обновление стартует при открытии.

- Create: `_publish/audit-macros-private/addin/build/build-activity-log-viewer.ps1`
  - Параметры:
    - `-OutputPath`
    - `-LogRoot`
    - `-OpenPassword`
  - Создает `!Журнал активности Excel.xlsm`.
  - Сохраняет книгу с password-to-open.
  - Импортирует viewer module и ThisWorkbook code через Excel COM.
  - Для dev по умолчанию кладет файл в `%APPDATA%\AuditAddin\assets\_system\activity-logs\!Журнал активности Excel.xlsm`.

### Tests and docs

- Create: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`
  - Source-level checks по runtime logger/store/viewer/build-script.

- Modify: `_publish/audit-macros-private/addin/tests/Invoke-AddinTests.ps1`
  - Добавить feature `activity-logs`.
  - Alias `activity-logger` должен включать `source-policy`, `assets-activity`, `activity-logs`.
  - `ChangedFiles` должен выбирать `activity-logs` для новых модулей и build-script viewer-а.

- Modify: `_publish/audit-macros-private/addin/tests/features/assets-activity.ps1`
  - Обновить старые проверки, которые сейчас ожидают отключенный логгер и старую центральную workbook-схему.

- Create: `_publish/audit-macros-private/addin/tests/activity-log-runtime-smoke-test.ps1`
  - Targeted Excel COM smoke для open/save/heartbeat/dev-store.
  - Запускать вручную после source-level тестов, потому что это Excel/COM integration.

- Modify: `_publish/audit-macros-private/docs/system-overview.md`
  - Добавить краткий раздел про новую систему activity logs, потому что меняется архитектура фонового runtime-сервиса и поток данных.

- Modify: `docs/superpowers/specs/2026-06-29-excel-activity-logs-design.md`
  - После реализации обновить статус и фактические имена процедур, если они отличаются от плана.

---

### Task 1: Source Tests For New Logging Architecture

**Files:**
- Create: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`
- Modify: `_publish/audit-macros-private/addin/tests/Invoke-AddinTests.ps1`

- [ ] **Step 1: Write failing source-level tests**

Create `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1` with checks for:

```powershell
. (Join-Path $PSScriptRoot "..\lib\AddinTestUtils.ps1")

$assets = Get-AddinTestText "addin\modAddinAssets.bas"
$logger = Get-AddinTestText "addin\modActivityLogger.bas"
$events = Get-AddinTestText "addin\CActivityLogAppEvents.cls"
$thisWorkbook = Get-AddinTestText "addin\build\ThisWorkbook.addin.txt"
$build = Get-AddinTestText "addin\build\build-audit-addin.ps1"
$storePath = Join-Path (Get-AddinTestRepoRoot) "addin\modActivityLogStore.bas"
$viewerModulePath = Join-Path (Get-AddinTestRepoRoot) "addin\admin_activity_log\modActivityLogViewer.bas"
$viewerWorkbookPath = Join-Path (Get-AddinTestRepoRoot) "addin\admin_activity_log\ThisWorkbook.activity-log-viewer.txt"
$viewerBuildPath = Join-Path (Get-AddinTestRepoRoot) "addin\build\build-activity-log-viewer.ps1"

foreach ($path in @($storePath, $viewerModulePath, $viewerWorkbookPath, $viewerBuildPath)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required activity log file missing: $path" }
}

$store = Get-Content -LiteralPath $storePath -Encoding UTF8 -Raw
$viewer = Get-Content -LiteralPath $viewerModulePath -Encoding UTF8 -Raw
$viewerWorkbook = Get-Content -LiteralPath $viewerWorkbookPath -Encoding UTF8 -Raw
$viewerBuild = Get-Content -LiteralPath $viewerBuildPath -Encoding UTF8 -Raw

foreach ($requiredText in @(
    "AddinAssets_GetActivityLogComputerStoreRoot",
    "AddinAssets_GetActivityLogViewerWorkbookPath",
    "NetworkAssetsRoot",
    "DeploymentMode"
)) {
    Assert-TextContains -Text $assets -Expected $requiredText -Label "activity log assets"
}

foreach ($requiredText in @(
    "ActivityLogStore_AppendEvent",
    "ActivityLogStore_FlushPendingRows",
    "ActivityLogStore_SanitizeTsvField",
    "ActivityLogStore_GetLocalSpoolPath",
    "ActivityLogStore_GetNetworkComputerLogPath",
    "<MachineName>__YYYY-MM.tsv"
)) {
    Assert-TextContains -Text $store -Expected $requiredText -Label "activity log store"
}

foreach ($forbiddenText in @("Kill", "RmDir", "DeleteFile", "DeleteFolder", "MoveFile", "MoveFolder", "Name old As new")) {
    Assert-TextNotContains -Text $store -Forbidden $forbiddenText -Label "activity log store forbidden file mutation"
    Assert-TextNotContains -Text $logger -Forbidden $forbiddenText -Label "activity logger forbidden file mutation"
}

foreach ($requiredText in @(
    "HEARTBEAT_INTERVAL_SECONDS As Long = 300",
    "ActivityLogger_RunHeartbeat",
    "ActivityLogStore_AppendEvent",
    "ActivityLogStore_FlushPendingRows"
)) {
    Assert-TextContains -Text $logger -Expected $requiredText -Label "activity logger runtime"
}

Assert-TextContains -Text $thisWorkbook -Expected 'RuntimeMacroTarget("ActivityLogger_WorkbookOpen")' -Label "runtime startup"
Assert-TextContains -Text $build -Expected "modActivityLogStore.bas" -Label "addin build"

foreach ($requiredText in @(
    "ActivityLogViewer_WorkbookOpen",
    "ActivityLogViewer_RefreshAll",
    "RawEvents",
    "Sessions",
    "Diagnostics",
    "EventId"
)) {
    Assert-TextContains -Text $viewer -Expected $requiredText -Label "activity log viewer"
}

Assert-TextContains -Text $viewerWorkbook -Expected "ActivityLogViewer_WorkbookOpen" -Label "activity viewer workbook open"
Assert-TextContains -Text $viewerBuild -Expected "!Журнал активности Excel.xlsm" -Label "activity viewer build"
Assert-TextContains -Text $viewerBuild -Expected "OpenPassword" -Label "activity viewer password"

Assert-NoForbiddenRuntimeMarkers -Paths @(
    $storePath,
    (Join-Path (Get-AddinTestRepoRoot) "addin\modActivityLogger.bas"),
    (Join-Path (Get-AddinTestRepoRoot) "addin\CActivityLogAppEvents.cls")
)

Write-Output "ADDIN_FEATURE_ACTIVITY_LOGS_OK"
```

- [ ] **Step 2: Wire the test into the dispatcher**

Modify `_publish/audit-macros-private/addin/tests/Invoke-AddinTests.ps1`:

```powershell
$featureScripts = [ordered]@{
    ...
    "activity-logs" = "features\activity-logs.ps1"
}

$aliases = @{
    ...
    "activity-logger" = @("source-policy", "assets-activity", "activity-logs")
}
```

In `Add-FeaturesForPath`, add `activity-logs` when paths match:

```powershell
if ($lower -match 'addin\\modactivitylogstore\.bas|addin\\admin_activity_log\\|addin\\build\\build-activity-log-viewer\.ps1') {
    Add-FeatureName -Target $Target -Name "activity-logs"
}
```

- [ ] **Step 3: Run failing test**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logs
```

Expected: FAIL because `modActivityLogStore.bas`, viewer files, and dispatcher wiring are not implemented yet.

- [ ] **Step 4: Commit test scaffolding**

Commit only test/dispatcher changes after verifying they fail for the expected reason:

```powershell
git add addin/tests/features/activity-logs.ps1 addin/tests/Invoke-AddinTests.ps1
git commit -m "test: add activity log architecture checks"
```

---

### Task 2: Asset Path Helpers

**Files:**
- Modify: `_publish/audit-macros-private/addin/modAddinAssets.bas`
- Test: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`
- Test: `_publish/audit-macros-private/addin/tests/features/assets-activity.ps1`

- [ ] **Step 1: Add config keys and path constants**

Add constants:

```vb
Private Const KEY_NETWORK_ASSETS_ROOT As String = "NetworkAssetsRoot"
Private Const KEY_DEPLOYMENT_MODE As String = "DeploymentMode"
Private Const DEPLOYMENT_MODE_DEV As String = "Dev"
Private Const ACTIVITY_LOG_ROOT_FOLDER As String = "activity-logs"
Private Const ACTIVITY_LOG_COMPUTERS_FOLDER As String = "computers"
Private Const ACTIVITY_LOG_VIEWER_WORKBOOK_NAME As String = "!Журнал активности Excel.xlsm"
```

- [ ] **Step 2: Add public helpers**

Add:

```vb
Public Function AddinAssets_GetNetworkAssetsRoot() As String
    Dim configuredRoot As String
    configuredRoot = Trim$(ReadConfigValue(KEY_NETWORK_ASSETS_ROOT))
    If Len(configuredRoot) > 0 And FolderExistsSafe(configuredRoot) Then
        AddinAssets_GetNetworkAssetsRoot = configuredRoot
    End If
End Function

Public Function AddinAssets_GetDeploymentMode() As String
    AddinAssets_GetDeploymentMode = Trim$(ReadConfigValue(KEY_DEPLOYMENT_MODE))
End Function

Public Function AddinAssets_GetLocalActivityLogDevRoot() As String
    Dim assetsRoot As String
    assetsRoot = AddinAssets_GetAssetsRoot()
    If Len(assetsRoot) = 0 Then Exit Function
    AddinAssets_GetLocalActivityLogDevRoot = CombinePath(CombinePath(assetsRoot, ACTIVITY_LOG_SYSTEM_FOLDER), ACTIVITY_LOG_ROOT_FOLDER)
End Function

Public Function AddinAssets_GetActivityLogRoot() As String
    Dim networkRoot As String
    Dim devRoot As String

    networkRoot = AddinAssets_GetNetworkAssetsRoot()
    If Len(networkRoot) > 0 Then
        AddinAssets_GetActivityLogRoot = CombinePath(CombinePath(networkRoot, ACTIVITY_LOG_SYSTEM_FOLDER), ACTIVITY_LOG_ROOT_FOLDER)
        Exit Function
    End If

    If StrComp(AddinAssets_GetDeploymentMode(), DEPLOYMENT_MODE_DEV, vbTextCompare) = 0 Then
        devRoot = AddinAssets_GetLocalActivityLogDevRoot()
        If Len(devRoot) > 0 Then AddinAssets_GetActivityLogRoot = devRoot
    End If
End Function

Public Function AddinAssets_GetActivityLogComputerStoreRoot() As String
    Dim rootPath As String
    rootPath = AddinAssets_GetActivityLogRoot()
    If Len(rootPath) = 0 Then Exit Function
    AddinAssets_GetActivityLogComputerStoreRoot = CombinePath(rootPath, ACTIVITY_LOG_COMPUTERS_FOLDER)
End Function

Public Function AddinAssets_GetActivityLogViewerWorkbookPath() As String
    Dim rootPath As String
    rootPath = AddinAssets_GetActivityLogRoot()
    If Len(rootPath) = 0 Then Exit Function
    AddinAssets_GetActivityLogViewerWorkbookPath = CombinePath(rootPath, ACTIVITY_LOG_VIEWER_WORKBOOK_NAME)
End Function
```

- [ ] **Step 3: Update smoke helper**

Update `AddinAssets_SmokeTest` so it still passes when old `workbook-state.xlsx` path is empty, and also verifies new activity log paths when assets root exists.

- [ ] **Step 4: Run source tests**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature assets-activity
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logs
```

Expected: `assets-activity` passes or fails only on checks that will be updated in later tasks; `activity-logs` still fails because runtime/store/viewer do not exist yet.

- [ ] **Step 5: Commit**

```powershell
git add addin/modAddinAssets.bas addin/tests/features/assets-activity.ps1
git commit -m "feat: add activity log asset paths"
```

---

### Task 3: Append-Only TSV Store

**Files:**
- Create: `_publish/audit-macros-private/addin/modActivityLogStore.bas`
- Modify: `_publish/audit-macros-private/addin/build/build-audit-addin.ps1`
- Test: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`

- [ ] **Step 1: Create module header**

Create `_publish/audit-macros-private/addin/modActivityLogStore.bas`:

```vb
Attribute VB_Name = "modActivityLogStore"
Option Explicit

Private Const MODULE_NAME As String = "modActivityLogStore"
Private Const LOG_FOLDER_NAME As String = "AuditAddin"
Private Const ACTIVITY_LOG_FOLDER_NAME As String = "activity-log"
Private Const SPOOL_FOLDER_NAME As String = "spool"
Private Const STATE_FOLDER_NAME As String = "state"
Private Const TSV_EXTENSION As String = ".tsv"
Private Const FOR_APPENDING As Long = 8
Private Const FOR_READING As Long = 1
Private Const FOR_WRITING As Long = 2
Private Const TRISTATE_TRUE As Long = -1
```

- [ ] **Step 2: Implement path helpers**

Add functions:

```vb
Public Function ActivityLogStore_GetLocalSpoolFolder() As String
Public Function ActivityLogStore_GetLocalStateFolder() As String
Public Function ActivityLogStore_GetLocalSpoolPath(ByVal eventMonth As Date) As String
Public Function ActivityLogStore_GetNetworkComputerLogPath(ByVal eventMonth As Date) As String
```

Rules:

- local spool: `%APPDATA%\AuditAddin\activity-log\spool\<MachineName>__YYYY-MM.tsv`
- local state: `%APPDATA%\AuditAddin\activity-log\state\<MachineName>__YYYY-MM.state`
- network/dev TSV: `AddinAssets_GetActivityLogComputerStoreRoot()\<MachineName>__YYYY-MM.tsv`

- [ ] **Step 3: Implement folder creation without delete/move**

Use `FileSystemObject.CreateFolder` recursively:

```vb
Private Sub EnsureFolderExists(ByVal folderPath As String)
```

Allowed: create missing folders. Forbidden: remove, rename, move, truncate.

- [ ] **Step 4: Implement TSV sanitization**

```vb
Public Function ActivityLogStore_SanitizeTsvField(ByVal value As String) As String
    value = Replace$(value, vbTab, " ")
    value = Replace$(value, vbCrLf, " ")
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    ActivityLogStore_SanitizeTsvField = value
End Function
```

- [ ] **Step 5: Implement append with header**

Add:

```vb
Public Sub ActivityLogStore_AppendEvent(ByVal eventLine As String)
```

Behavior:

- create local spool folder;
- create file with header if missing;
- append exactly one sanitized row;
- on any error: `Debug.Print`, no MsgBox.

- [ ] **Step 6: Implement sync state**

Store last flushed row count per month:

```text
LastFlushedLine=<number>
```

Functions:

```vb
Private Function ReadLastFlushedLine(ByVal statePath As String) As Long
Private Sub WriteLastFlushedLine(ByVal statePath As String, ByVal lineNumber As Long)
```

Writing the state file is allowed as local config/log writing. Do not delete old state.

- [ ] **Step 7: Implement network flush**

Add:

```vb
Public Function ActivityLogStore_FlushPendingRows() As Boolean
```

Behavior:

- read local spool for current month;
- skip header and rows already flushed;
- resolve network/dev monthly target with `ActivityLogStore_GetNetworkComputerLogPath(Date)`;
- if target path empty or folder unavailable, return `False`;
- ensure target folder;
- create target with same header if missing;
- append pending rows;
- update state after successful append;
- if target is read-only/locked, skip quickly and return `False`;
- no delete/move/rename/truncate.

- [ ] **Step 8: Wire module into build**

Modify `_publish/audit-macros-private/addin/build/build-audit-addin.ps1`:

```powershell
$activityLogStoreModulePath = Join-Path $repoRoot "addin\modActivityLogStore.bas"
...
if (-not (Test-Path $activityLogStoreModulePath)) { throw "Activity log store module not found." }
...
$vbProject.VBComponents.Import((New-VbeImportCopy -Path $activityLogStoreModulePath)) | Out-Null
```

Import it before `modActivityLogger.bas`.

- [ ] **Step 9: Run tests**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logs
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature build
```

Expected: activity-log source checks pass for store/build portions; viewer/runtime checks may still fail until later tasks.

- [ ] **Step 10: Commit**

```powershell
git add addin/modActivityLogStore.bas addin/build/build-audit-addin.ps1 addin/tests/features/activity-logs.ps1
git commit -m "feat: add activity log TSV store"
```

---

### Task 4: Runtime Activity Logger

**Files:**
- Modify: `_publish/audit-macros-private/addin/modActivityLogger.bas`
- Modify: `_publish/audit-macros-private/addin/CActivityLogAppEvents.cls`
- Modify: `_publish/audit-macros-private/addin/build/ThisWorkbook.addin.txt`
- Modify: `_publish/audit-macros-private/addin/tests/features/assets-activity.ps1`
- Test: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`

- [ ] **Step 1: Replace central workbook runtime path**

In `modActivityLogger.bas`, retire old central workbook writing from live event handlers. Keep compatibility procedures if tests/admin scripts expect them, but runtime event flow must call `ActivityLogStore_AppendEvent` instead of opening/saving a central workbook.

- [ ] **Step 2: Add runtime constants and state**

Use:

```vb
Private Const ACTIVITY_LOGGER_ENABLED As Boolean = True
Private Const HEARTBEAT_INTERVAL_SECONDS As Long = 300
Private mExcelSessionId As String
Private mEventSequence As Long
Private mNextHeartbeatAt As Date
Private mHeartbeatScheduled As Boolean
Private mTrackedWorkbookSessions As Object
```

If risk requires a safety valve, add hidden admin setting:

```vb
Public Sub ActivityLogger_Enable()
Public Sub ActivityLogger_Disable()
Public Function ActivityLogger_IsEnabled() As Boolean
```

Store it via `GetSetting`/`SaveSetting`, not Ribbon UI.

- [ ] **Step 3: Generate IDs**

Add:

```vb
Private Function BuildExcelSessionId() As String
Private Function NextEventId() As String
Private Function BuildWorkbookSessionId(ByVal targetWorkbook As Workbook) As String
```

`EventId = MachineName & "-" & ExcelSessionId & "-" & Format$(mEventSequence, "000000")`.

- [ ] **Step 4: Build event rows**

Add helper:

```vb
Private Function BuildActivityEventLine(ByVal eventType As String, ByVal targetWorkbook As Workbook, Optional ByVal notes As String = "") As String
```

It must output fields in the exact header order from the spec. Use `ActivityLogStore_SanitizeTsvField` for text.

- [ ] **Step 5: Implement startup**

`ActivityLogger_WorkbookOpen`:

- checks enabled;
- calls `StartActivityLogger`;
- does not synchronously scan all workbooks for heavy network writes.

`StartActivityLogger`:

- creates `CActivityLogAppEvents`;
- assigns `Application`;
- creates `ExcelSessionId`;
- writes `SESSION_START` to local spool;
- records already open saved user workbooks as lightweight `OPEN` events if needed;
- schedules heartbeat only if there are trackable workbooks.

- [ ] **Step 6: Implement event handlers**

`ActivityLogger_RecordWorkbookOpen`:

- ignore add-ins, unsaved workbooks, log files, central viewer workbook;
- append `OPEN`;
- track workbook session;
- schedule heartbeat.

`ActivityLogger_RecordWorkbookSave`:

- ensure open tracked;
- append `SAVE`;
- update snapshot.

`ActivityLogger_RecordWorkbookBeforeClose`:

- append local `MISSING` or `SESSION_END_ATTEMPT` only;
- do not flush network;
- forget tracking;
- if no tracked workbooks remain, cancel heartbeat.

- [ ] **Step 7: Implement heartbeat**

Add:

```vb
Public Sub ActivityLogger_RunHeartbeat()
```

Behavior:

- cancel current heartbeat marker at start;
- enumerate open user workbooks;
- append `SEEN` for tracked visible workbooks;
- append `MISSING` for tracked workbooks no longer visible;
- call `ActivityLogStore_FlushPendingRows`;
- reschedule only if `gApp` exists and at least one trackable workbook remains.

- [ ] **Step 8: Update `ThisWorkbook.addin.txt`**

Add:

```vb
Private mNextActivityLoggerInitAt As Date
```

In `Workbook_Open`:

```vb
mNextActivityLoggerInitAt = Now + TimeSerial(0, 0, 1)
Application.OnTime mNextActivityLoggerInitAt, RuntimeMacroTarget("ActivityLogger_WorkbookOpen")
```

In `Workbook_BeforeClose`, cancel it like hotkeys/zoom and keep:

```vb
StopActivityLogger
```

- [ ] **Step 9: Update event class**

`CActivityLogAppEvents.cls` should call only:

- `ActivityLogger_RecordWorkbookOpen`
- `ActivityLogger_RecordWorkbookSave` from `WorkbookAfterSave` when `Success=True`
- `ActivityLogger_RecordWorkbookBeforeClose`

If `WorkbookBeforeSave` stays, it must not do heavy work.

- [ ] **Step 10: Update old source tests**

Update `assets-activity.ps1`:

- remove requirement `Private Const ACTIVITY_LOGGER_ENABLED As Boolean = False`;
- remove assertions that forbid `RuntimeMacroTarget("ActivityLogger_WorkbookOpen")`;
- assert close path does not call central workbook open/save;
- keep safety checks for `ThisWorkbook.FullName` timer target.

- [ ] **Step 11: Run tests**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logger
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature setup
```

Expected: both pass.

- [ ] **Step 12: Commit**

```powershell
git add addin/modActivityLogger.bas addin/CActivityLogAppEvents.cls addin/build/ThisWorkbook.addin.txt addin/tests/features/assets-activity.ps1
git commit -m "feat: log workbook activity to per-computer TSV"
```

---

### Task 5: Admin Viewer Workbook Source

**Files:**
- Create: `_publish/audit-macros-private/addin/admin_activity_log/modActivityLogViewer.bas`
- Create: `_publish/audit-macros-private/addin/admin_activity_log/ThisWorkbook.activity-log-viewer.txt`
- Test: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`

- [ ] **Step 1: Create viewer module**

Create `modActivityLogViewer.bas` with:

```vb
Attribute VB_Name = "modActivityLogViewer"
Option Explicit

Private Const LOG_COMPUTERS_FOLDER_NAME As String = "computers"
Private Const SHEET_DASHBOARD As String = "Dashboard"
Private Const SHEET_RAW_EVENTS As String = "RawEvents"
Private Const SHEET_SESSIONS As String = "Sessions"
Private Const SHEET_FILES As String = "Files"
Private Const SHEET_USERS As String = "Users"
Private Const SHEET_COMPUTERS As String = "Computers"
Private Const SHEET_DIAGNOSTICS As String = "Diagnostics"
```

- [ ] **Step 2: Add public open hook**

```vb
Public Sub ActivityLogViewer_WorkbookOpen()
    On Error GoTo EH
    ActivityLogViewer_RefreshAll
    Exit Sub
EH:
    Debug.Print Now, "ActivityLogViewer open refresh failed: " & Err.Number & " - " & Err.Description
End Sub
```

- [ ] **Step 3: Add refresh importer**

```vb
Public Sub ActivityLogViewer_RefreshAll()
```

Behavior:

- resolve log root from workbook location: `ThisWorkbook.Path\computers`;
- if missing, create/clear `Diagnostics` with readable message;
- read all `*.tsv`;
- validate header;
- skip duplicate `EventId`;
- skip malformed rows;
- write `RawEvents`.

- [ ] **Step 4: Add derived sheets**

Implement simple first version:

- `Sessions`: group by `WorkbookSessionId`, compute `OpenedAt`, `LastSeenAt`, `MissingAt`, `ApproxClosedAt`, `OpenDurationMinutes`, `SaveCount`.
- `Files`: group by `WorkbookFullName`, count opens/saves, last activity.
- `Users`: group by `WindowsUser`, count events, unique files.
- `Computers`: group by `MachineName`, event count, last event.
- `Dashboard`: summary numbers and simple filter-friendly tables.
- `Diagnostics`: import status per source file.

- [ ] **Step 5: Create ThisWorkbook template**

Create:

```vb
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ActivityLogViewer_WorkbookOpen
    On Error GoTo 0
End Sub
```

- [ ] **Step 6: Run source test**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logs
```

Expected: viewer source checks pass except build script until next task.

- [ ] **Step 7: Commit**

```powershell
git add addin/admin_activity_log/modActivityLogViewer.bas addin/admin_activity_log/ThisWorkbook.activity-log-viewer.txt
git commit -m "feat: add activity log admin viewer source"
```

---

### Task 6: Admin Viewer Build Script

**Files:**
- Create: `_publish/audit-macros-private/addin/build/build-activity-log-viewer.ps1`
- Test: `_publish/audit-macros-private/addin/tests/features/activity-logs.ps1`

- [ ] **Step 1: Write build script parameters**

Script parameters:

```powershell
param(
    [string] $OutputPath = "",
    [string] $LogRoot = "",
    [Parameter(Mandatory = $true)]
    [string] $OpenPassword
)
```

If `OutputPath` is empty, use:

```powershell
Join-Path $env:APPDATA "AuditAddin\assets\_system\activity-logs\!Журнал активности Excel.xlsm"
```

If `LogRoot` is empty, use output folder. Ensure `computers` subfolder exists.

- [ ] **Step 2: Create workbook through Excel COM**

Build script should:

- create output parent folder and `computers` folder;
- start Excel COM hidden;
- create workbook;
- import `modActivityLogViewer.bas` through ANSI temp copy;
- replace ThisWorkbook code with `ThisWorkbook.activity-log-viewer.txt`;
- create required sheets;
- save as `.xlsm` with password-to-open:

```powershell
$workbook.SaveAs($resolvedOutputPath, 52, $OpenPassword)
```

- close workbook and quit Excel in `finally`.

- [ ] **Step 3: Add sample seed option only if needed**

Do not add sample data by default. If manual testing needs fixtures, keep them in test scripts, not in the admin workbook.

- [ ] **Step 4: Run source test**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logs
```

Expected: source tests pass.

- [ ] **Step 5: Build dev admin workbook**

Need real password from user before this step.

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\build\build-activity-log-viewer.ps1 -OpenPassword "<ASK_USER_PASSWORD>"
```

Expected output file:

```text
%APPDATA%\AuditAddin\assets\_system\activity-logs\!Журнал активности Excel.xlsm
```

- [ ] **Step 6: Commit**

```powershell
git add addin/build/build-activity-log-viewer.ps1 addin/tests/features/activity-logs.ps1
git commit -m "build: add activity log admin workbook builder"
```

Do not commit generated `%APPDATA%` admin workbook. If later a repo release copy is desired, agree location first.

---

### Task 7: Targeted Runtime Smoke Test

**Files:**
- Create: `_publish/audit-macros-private/addin/tests/activity-log-runtime-smoke-test.ps1`

- [ ] **Step 1: Write COM smoke test**

The test should:

- create temp `%APPDATA%`-like root using `AUDIT_ADDIN_APP_ROOT`;
- write `addin-updater.ini` with:

```text
DeploymentMode=Dev
AssetsRoot=<temp>\assets
```

- create temp workbook `ActivitySource.xlsx`;
- build or use `addin\release\AuditAddin.xlam`;
- open add-in in isolated Excel COM session;
- open source workbook;
- run/allow `ActivityLogger_WorkbookOpen`;
- save source workbook;
- run `ActivityLogger_RunHeartbeat`;
- close source workbook;
- quit Excel;
- assert local spool has `OPEN`, `SAVE`, `SEEN`;
- assert dev computer TSV exists under `<temp>\assets\_system\activity-logs\computers`.

- [ ] **Step 2: Avoid modal hangs**

Use:

```powershell
$excel.DisplayAlerts = $false
$excel.Visible = $false
$excel.AutomationSecurity = 1
```

Explicitly close all books and `Quit()` in `finally`.

- [ ] **Step 3: Run smoke test**

Run after publishing dev add-in:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\activity-log-runtime-smoke-test.ps1
```

Expected: outputs clear marker:

```text
ACTIVITY_LOG_RUNTIME_SMOKE_OK
```

- [ ] **Step 4: Commit**

```powershell
git add addin/tests/activity-log-runtime-smoke-test.ps1
git commit -m "test: add activity log runtime smoke"
```

---

### Task 8: Documentation Updates

**Files:**
- Modify: `_publish/audit-macros-private/docs/system-overview.md`
- Modify: `docs/superpowers/specs/2026-06-29-excel-activity-logs-design.md`

- [ ] **Step 1: Update system overview**

Add a short section:

- per-computer monthly TSV logs;
- runtime writes local spool then flushes to `NetworkAssetsRoot\_system\activity-logs\computers`;
- admin workbook `!Журнал активности Excel.xlsm`;
- no delete/move cleanup in runtime;
- heartbeat 5 minutes.

- [ ] **Step 2: Update spec status**

Change status to something like:

```text
Статус: реализуется по плану docs/superpowers/plans/2026-06-29-excel-activity-logs.md.
```

Do not claim implementation is complete until tests and dev install pass.

- [ ] **Step 3: Run documentation-relevant tests**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logger
```

- [ ] **Step 4: Commit**

```powershell
git add ..\docs\superpowers\specs\2026-06-29-excel-activity-logs-design.md docs/system-overview.md
git commit -m "docs: document activity log architecture"
```

Note: adjust `git add` paths depending on whether the command is run from workspace root or `_publish/audit-macros-private`.

---

### Task 9: Final Verification, Dev Publish, Install

**Files:**
- Generated: `_publish/audit-macros-private/addin/release/AuditAddin.xlam`
- Generated: `_publish/audit-macros-private/addin/release/AuditAddin.version.txt`
- Generated: `_publish/audit-macros-private/addin/release/AuditAddin.revision.txt`
- Generated: `_publish/audit-macros-private/addin/release/AuditAddinLoader.xlam`

- [ ] **Step 1: Run fast tests**

Run from `_publish/audit-macros-private`:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature activity-logger
powershell -ExecutionPolicy Bypass -File .\addin\tests\Invoke-AddinTests.ps1 -Feature build
```

Expected:

```text
ADDIN_TESTS_OK
```

- [ ] **Step 2: Publish dev release**

Close Excel first if needed.

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\build\publish-addin-release.ps1
```

Expected:

```text
RELEASE_ROOT=...
```

- [ ] **Step 3: Install dev add-in locally**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\build\Install-AuditAddin.ps1 -DeploymentMode Dev -ReleaseRoot .\addin\release
```

Expected:

```text
DEPLOYMENT_MODE=Dev
LOCAL_LOADER_PATH=...
LOCAL_ADDIN_PATH=...
```

- [ ] **Step 4: Run targeted runtime smoke**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\tests\activity-log-runtime-smoke-test.ps1
```

Expected:

```text
ACTIVITY_LOG_RUNTIME_SMOKE_OK
```

- [ ] **Step 5: Build admin workbook**

Requires user-provided open password:

```powershell
powershell -ExecutionPolicy Bypass -File .\addin\build\build-activity-log-viewer.ps1 -OpenPassword "<ASK_USER_PASSWORD>"
```

Expected:

```text
%APPDATA%\AuditAddin\assets\_system\activity-logs\!Журнал активности Excel.xlsm
```

- [ ] **Step 6: Stage only relevant files**

From `_publish/audit-macros-private`:

```powershell
git add addin/modAddinAssets.bas
git add addin/modActivityLogStore.bas
git add addin/modActivityLogger.bas
git add addin/CActivityLogAppEvents.cls
git add addin/build/ThisWorkbook.addin.txt
git add addin/build/build-audit-addin.ps1
git add addin/build/build-activity-log-viewer.ps1
git add addin/admin_activity_log/modActivityLogViewer.bas
git add addin/admin_activity_log/ThisWorkbook.activity-log-viewer.txt
git add addin/tests/Invoke-AddinTests.ps1
git add addin/tests/features/activity-logs.ps1
git add addin/tests/features/assets-activity.ps1
git add addin/tests/activity-log-runtime-smoke-test.ps1
git add docs/system-overview.md
git add -f addin/release/AuditAddin.xlam
git add addin/release/AuditAddin.version.txt
git add addin/release/AuditAddin.revision.txt
git add -f addin/release/AuditAddinLoader.xlam
```

From workspace root, stage updated spec/plan if not already committed:

```powershell
git add docs/superpowers/specs/2026-06-29-excel-activity-logs-design.md
git add docs/superpowers/plans/2026-06-29-excel-activity-logs.md
```

- [ ] **Step 7: Final commit and push**

Commit in `_publish/audit-macros-private`:

```powershell
git commit -m "feat: add safe Excel activity logging"
git push origin main
```

Commit workspace spec/plan if needed:

```powershell
git commit -m "docs: add Excel activity logs implementation plan"
```

Do not include unrelated dirty files.

---

## Final Manual Checks

- Open Excel normally with dev add-in installed.
- Open a saved workbook.
- Save it.
- Wait or manually run heartbeat if a hidden admin procedure exists.
- Confirm local spool exists under `%APPDATA%\AuditAddin\activity-log\spool`.
- Confirm dev computer TSV exists under `%APPDATA%\AuditAddin\assets\_system\activity-logs\computers`.
- Open `%APPDATA%\AuditAddin\assets\_system\activity-logs\!Журнал активности Excel.xlsm` with password.
- Confirm sheets fill automatically without pressing an update button.
- Close Excel and confirm no invisible hanging `EXCEL.EXE` remains.

## Known Blocker Before Full Execution

Need the exact password for opening `!Журнал активности Excel.xlsm`.

Implementation can start before this, but Task 6 Step 5 and Task 9 Step 5 cannot be completed without that password.
