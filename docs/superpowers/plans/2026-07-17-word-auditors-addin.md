# Word Add-in `Аудиторам` Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Создать отдельную Word VBA-надстройку `.dotm` с четырьмя Ribbon-командами, скрытым Activity Logger, dual-slot Setup/Loader для трех каналов и единой admin-книгой активности Excel+Word.

**Architecture:** Word-линия живет только в `word_addin/` и `word_addin_updater/`, загружается по цепочке `Setup -> Startup Loader -> local runtime` и работает с локальным кэшем шаблонов. Excel runtime и его TSV остаются неизменными; Word пишет отдельные per-session TSV в `word-computers`, а новая `!Журнал активности Office.xlsm` нормализует обе схемы.

**Tech Stack:** Word VBA, Excel VBA для admin viewer, RibbonX `customUI14`, PowerShell, Word/Excel COM, Open XML package editing, UTF-8 source + временный CP1251 import, UTF-16LE TSV/control.

**Required skills during execution:** `@test-driven-development`, `@windows-excel-vba-ru-paths`, `@verification-before-completion`, затем `@requesting-code-review` и `@finishing-a-development-branch`.

---

## Рабочая Среда И Зафиксированный Baseline

- Реализация идет в worktree `C:\Users\pavel\.config\superpowers\worktrees\audit-macros-private\word-auditors-addin`.
- Ветка: `codex/word-auditors-addin`.
- Спецификация: `C:\Users\pavel\OneDrive\Рабочий стол\Codex_venv\Макросы для аудита\docs\superpowers\specs\2026-07-17-word-auditors-addin-design.md`.
- До изменений `addin/tests/Invoke-AddinTests.ps1 -Feature source-policy` прошел.
- Существующий `all-fast` baseline падает в `sampling` с `derived sampling source readiness If is missing End If`; пользователь явно разрешил продолжать Word-разработку без исправления этого несвязанного Excel-теста.
- На протяжении работы запрещено менять `addin/`, `addin_updater/`, `shared/`, `profiles/` и legacy `addin/admin_activity_log/`.
- Company release не публикуется без отдельной команды и паролей. В этой реализации создаются и тестируются пайплайны; основной deliverable — dev runtime и локальная dev-установка.

## Карта Новых Файлов

```text
word_addin/
  README.md
  modWordAddinVersion.bas
  modWordConfig.bas
  modWordFileSystem.bas
  modWordRuntime.bas
  modWordRibbon.bas
  modWordCommands.bas
  modWordScheduler.bas
  CWordAppEvents.cls
  modWordTemplateCatalog.bas
  modWordTemplateOpener.bas
  frmWordTemplateOpener.code.txt
  modWordViewSettings.bas
  frmWordViewSettings.code.txt
  modWordHotkeys.bas
  CWordHotkeyTextBoxEvents.cls
  frmWordHotkeySettings.code.txt
  modWordActivityLogStore.bas
  modWordActivityLogger.bas
  CWordPendingSave.cls
  modWordDiagnostics.bas
  frmWordDiagnostics.code.txt
  assets/
    rd-template-open.png
    view-settings.png
    hotkeys-keycap.png
    diagnostics-status.png
  admin_activity_log/
    modOfficeActivityLogImport.bas
    modOfficeActivityLogViewer.bas
    ThisWorkbook.office-activity-log-viewer.txt
  build/
    AuditWordAddin.customUI14.xml
    ThisDocument.runtime.txt
    ThisDocument.loader.txt
    ThisDocument.setup.txt
    build-audit-word-addin.ps1
    build-audit-word-loader.ps1
    build-audit-word-setup.ps1
    build-office-activity-log-viewer.ps1
    Install-AuditWordAddin.ps1
    Uninstall-AuditWordAddin.ps1
    publish-word-addin-release.ps1
    publish-company-word-addin-release.ps1
    publish-office-activity-company-release.ps1
    protect-vba-project.py
  tests/
    Invoke-WordAddinTests.ps1
    lib/WordAddinTestUtils.ps1
    fixtures/New-WordAddinFixtures.ps1
    features/source-policy.ps1
    features/structure.ps1
    features/ribbon.ps1
    features/core-config.ps1
    features/scheduler.ps1
    features/rd-template-opener.ps1
    features/view-settings.ps1
    features/hotkeys.ps1
    features/activity-logger.ps1
    features/diagnostics.ps1
    features/setup-loader.ps1
    features/build.ps1
    features/office-activity-viewer.ps1
    features/release-structure.ps1
    setup-loader-integration-test.ps1
    smoke-built-word-addin.ps1
    office-activity-viewer-integration-test.ps1
  release/
    AuditWordAddin.dotm
    AuditWordAddin.version.txt

word_addin_updater/
  modWordUpdaterVersion.bas
  modWordUpdaterConfig.bas
  modWordUpdaterFileSystem.bas
  modWordAddinUpdater.bas
  modWordSetupBootstrap.bas

docs/
  word-addin-agent-instructions.md
  testers-word-addin-install.md
  system-overview.md
```

Файлы форм разделяются как в Excel-линии: build-скрипт создает UserForm/controls через VBIDE, затем импортирует код из `*.code.txt`. Большие обязанности не смешивать: parser каталога не показывает UI; logger store не подписывается на Word events; updater не входит в runtime `.dotm`.

---

### Task 1: Создать автономный Word test harness

**Files:**
- Create: `word_addin/README.md`
- Create: `word_addin/tests/lib/WordAddinTestUtils.ps1`
- Create: `word_addin/tests/Invoke-WordAddinTests.ps1`
- Create: `word_addin/tests/features/source-policy.ps1`
- Create: `word_addin/tests/features/structure.ps1`

- [ ] **Step 1: Написать harness-тест структуры**

`structure.ps1` должен требовать `word_addin/README.md`, папки `tests/features`, `build`, `assets`, `admin_activity_log`, а также проверять через Git, что ветка не меняет Excel-линии:

```powershell
$forbiddenChanges = @(git -C $repoRoot diff --name-only main...HEAD -- addin addin_updater shared profiles)
if ($forbiddenChanges.Count -gt 0) {
    throw "Word line must not modify Excel/runtime macro paths: $($forbiddenChanges -join ', ')"
}
```

- [ ] **Step 2: Запустить тест напрямую и увидеть RED**

Run: `& .\word_addin\tests\features\structure.ps1`

Expected: FAIL с отсутствующими Word-папками/README.

- [ ] **Step 3: Реализовать test utilities и runner**

`WordAddinTestUtils.ps1` должен дать `Get-WordAddinRepoRoot`, `Get-WordAddinText`, `Assert-TextContains`, `Assert-TextNotContains`, `Assert-NoWordRuntimeForbiddenMarkers`, `Get-WordChangedPaths`. Forbidden list повторяет company-ограничения: Git/HTTP/shell/Outlook/delete/move.

`Invoke-WordAddinTests.ps1` поддерживает `-Feature`, `-ChangedFiles`, `-List`; стабильный output:

```text
WORD_ADDIN_TEST_PLAN=<features>
WORD_ADDIN_TEST_START=<feature>
WORD_ADDIN_TEST_OK=<feature>; seconds=<n>
WORD_ADDIN_TESTS_OK=<count>
```

- [ ] **Step 4: Создать минимальную структуру и проверить GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature structure,source-policy`

Expected: `WORD_ADDIN_TESTS_OK=2`.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/README.md word_addin/tests
git commit -m "test: scaffold Word add-in test harness"
```

---

### Task 2: Зафиксировать Ribbon и runtime lifecycle

**Files:**
- Create: `word_addin/tests/features/ribbon.ps1`
- Create: `word_addin/build/AuditWordAddin.customUI14.xml`
- Create: `word_addin/build/ThisDocument.runtime.txt`
- Create: `word_addin/modWordAddinVersion.bas`
- Create: `word_addin/modWordRuntime.bas`
- Create: `word_addin/modWordRibbon.bas`
- Create: `word_addin/modWordCommands.bas`
- Copy: `addin/assets/{rd-template-open.png,view-settings.png,hotkeys-keycap.png,diagnostics-status.png}` to `word_addin/assets/`

- [ ] **Step 1: Написать RED Ribbon contract**

Тест парсит XML и требует ровно четыре элемента с `onAction`:

```text
wordOpenRdTemplate -> WordRibbon_OpenRdTemplate
wordConfigureView  -> WordRibbon_ConfigureView
wordHotkeys        -> WordRibbon_Hotkeys
wordDiagnostics    -> WordRibbon_Diagnostics
```

Он запрещает menu children, Activity Logger control и любые дополнительные `onAction`.

- [ ] **Step 2: Подтвердить RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature ribbon`

Expected: FAIL — XML/modules отсутствуют.

- [ ] **Step 3: Реализовать Ribbon/runtime shell**

`customUI14` создает tab `Аудиторам` после `TabReviewWord`, четыре группы и четыре large button. `modWordRibbon` хранит Ribbon object и четыре callbacks. `modWordCommands` дает ровно четыре публичных entry point:

```vb
Public Sub WordCommand_OpenRdTemplate()
Public Sub WordCommand_ConfigureView()
Public Sub WordCommand_Hotkeys()
Public Sub WordCommand_Diagnostics()
```

`ThisDocument.runtime.txt` вызывает `WordRuntime_AutoExec`/`WordRuntime_AutoExit`; lifecycle пока только инициализирует version и безопасно логирует ошибки через `Debug.Print`.

- [ ] **Step 4: Проверить Ribbon/source policy**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature ribbon,source-policy`

Expected: PASS; PNG-файлы существуют и Excel assets не изменены.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin
git commit -m "feat: add Word add-in runtime and Ribbon shell"
```

---

### Task 3: Реализовать config, user settings и безопасные path helpers

**Files:**
- Create: `word_addin/tests/features/core-config.ps1`
- Create: `word_addin/modWordConfig.bas`
- Create: `word_addin/modWordFileSystem.bas`
- Modify: `word_addin/modWordAddinVersion.bas`

- [ ] **Step 1: Написать RED pure-contract tests**

Тест требует публичные smoke helpers:

```vb
WordConfig_ReadValueForSmokeTest(configText, keyName)
WordConfig_UpsertValueForSmokeTest(configText, keyName, value)
WordPath_CombineForSmokeTest(leftPath, rightPath)
WordPath_IsWithinRootForSmokeTest(rootPath, candidatePath)
WordVersion_IsValidForSmokeTest(value)
WordVersion_CompareForSmokeTest(leftValue, rightValue)
```

Cases: Cyrillic paths, empty required path, duplicate key last-wins rejection, `..`, absolute RelativePath, prefix collision `templates-old`, versions `YYYY.MM.DD.HHMMSS`, invalid/downgrade.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature core-config`

Expected: FAIL — helpers отсутствуют.

- [ ] **Step 3: Реализовать минимальные модули**

`modWordConfig` читает UTF-16LE updater/user INI, сохраняет неизвестные ключи и не обращается к сети. `modWordFileSystem` содержит только разрешенные helpers (`FileExists`, `FolderExists`, `CreateFolder`, `CopyFile`, text streams). Root-boundary проверка канонизирует оба пути и сравнивает `LCase$(root & "\")` с началом candidate.

`modWordAddinVersion` возвращает build version и валидирует exact pattern.

- [ ] **Step 4: Run GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature core-config,source-policy`

Expected: PASS.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/modWordConfig.bas word_addin/modWordFileSystem.bas word_addin/modWordAddinVersion.bas word_addin/tests/features/core-config.ps1
git commit -m "feat: add Word add-in config and path contracts"
```

---

### Task 4: Добавить единый Word scheduler и Application events

**Files:**
- Create: `word_addin/tests/features/scheduler.ps1`
- Create: `word_addin/modWordScheduler.bas`
- Create: `word_addin/CWordAppEvents.cls`
- Modify: `word_addin/modWordRuntime.bas`
- Modify: `word_addin/build/ThisDocument.runtime.txt`

- [ ] **Step 1: Написать RED scheduler test**

Тест должен доказать, что строка `Application.OnTime` встречается только в `modWordScheduler.bas`, а update/assets/pending-save/heartbeat/flush ставятся как queue task types. Требовать `ScheduleOrCoalesce`, nearest `DueAt`, `WordScheduler_RunDueTasks`, event-driven self-heal и fully qualified macro path через runtime template.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature scheduler`

Expected: FAIL.

- [ ] **Step 3: Реализовать scheduler/event class**

Один in-memory dictionary хранит `TaskType -> DueAt/payload/sequence`. Новый task того же типа заменяет старый. Callback обрабатывает due tasks и планирует ближайший остаток. `CWordAppEvents` подписывается на `DocumentOpen`, `NewDocument`, `DocumentBeforeSave`, `DocumentBeforeClose`, `WindowActivate`, `Quit`; на этом шаге вызывает lifecycle hooks, которые безопасно существуют как no-op до следующих задач.

- [ ] **Step 4: Run GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature scheduler,source-policy`

Expected: PASS и ровно одна source-точка `Application.OnTime`.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/modWordScheduler.bas word_addin/CWordAppEvents.cls word_addin/modWordRuntime.bas word_addin/build/ThisDocument.runtime.txt word_addin/tests/features/scheduler.ps1
git commit -m "feat: add centralized Word scheduler and events"
```

---

### Task 5: Реализовать metadata-driven `Открыть шаблон РД`

**Files:**
- Create: `word_addin/tests/features/rd-template-opener.ps1`
- Create: `word_addin/tests/fixtures/New-WordAddinFixtures.ps1`
- Create: `word_addin/modWordTemplateCatalog.bas`
- Create: `word_addin/modWordTemplateOpener.bas`
- Create: `word_addin/frmWordTemplateOpener.code.txt`
- Modify: `word_addin/modWordCommands.bas`

- [ ] **Step 1: Написать RED catalog tests**

Fixture generator пишет UTF-8 BOM `word-rd-template-map.tsv` с колонками `TemplateId, AuditType, StageMode, StageVariant, RdCode, DisplayName, RelativePath, Enabled`. Cases: zero rows, enabled/disabled, duplicate id, duplicate selection key, missing file, unknown extra column, absolute path, `..`, `.docx`, valid `.dotx/.dotm`.

Тест требует parser API и `Documents.Add Template:=resolvedPath, NewTemplate:=False, Visible:=True`.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature rd-template-opener`

Expected: FAIL.

- [ ] **Step 3: Реализовать parser/opener/form**

Parser работает только с локальными `AssetsRoot\sync` и `RdTemplateRoot`, возвращает diagnostics collection и активные records. Форма каскадно заполняет вид аудита, этапность/вариант, код/имя. On-demand handler не обращается к `Z:`; пустой/битый каталог показывает одно понятное сообщение. Новый документ остается несохраненным, source template не меняется.

- [ ] **Step 4: Run GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature rd-template-opener,source-policy`

Expected: PASS.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/modWordTemplateCatalog.bas word_addin/modWordTemplateOpener.bas word_addin/frmWordTemplateOpener.code.txt word_addin/modWordCommands.bas word_addin/tests
git commit -m "feat: add Word RD template catalog and opener"
```

---

### Task 6: Реализовать `Настроить вид`

**Files:**
- Create: `word_addin/tests/features/view-settings.ps1`
- Create: `word_addin/modWordViewSettings.bas`
- Create: `word_addin/frmWordViewSettings.code.txt`
- Modify: `word_addin/modWordCommands.bas`
- Modify: `word_addin/CWordAppEvents.cls`

- [ ] **Step 1: Написать RED view tests**

Тест требует safe defaults (`Enabled=No`, все `NoChange`), допустимый zoom, allowlist режимов Word, navigation pane и formatting marks. Event hooks должны применять только активное окно. Source contract сохраняет исходный `Document.Saved` и не вызывает `Document.Save`.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature view-settings`

Expected: FAIL.

- [ ] **Step 3: Реализовать форму и применение**

Одна форма содержит tabs/frames `При открытии` и `При закрытии`, поля zoom/view/navigation/formatting, checkboxes automation и две внутренние кнопки `Применить сейчас`. Применять `Window.View.Type`, `View.Zoom.Percentage`, `Window.DocumentMap`, `View.ShowAll` best-effort по одному свойству; ошибка одного свойства не отменяет остальные.

`CWordAppEvents` вызывает open profile на `DocumentOpen/NewDocument`, close profile на `DocumentBeforeClose` без файловых/сетевых операций.

- [ ] **Step 4: Run GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature view-settings`

Expected: PASS.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/modWordViewSettings.bas word_addin/frmWordViewSettings.code.txt word_addin/modWordCommands.bas word_addin/CWordAppEvents.cls word_addin/tests/features/view-settings.ps1
git commit -m "feat: add Word document view settings"
```

---

### Task 7: Реализовать `Горячие клавиши`

**Files:**
- Create: `word_addin/tests/features/hotkeys.ps1`
- Create: `word_addin/modWordHotkeys.bas`
- Create: `word_addin/CWordHotkeyTextBoxEvents.cls`
- Create: `word_addin/frmWordHotkeySettings.code.txt`
- Modify: `word_addin/modWordCommands.bas`
- Modify: `word_addin/modWordRuntime.bas`

- [ ] **Step 1: Написать RED hotkey contract**

Allowlist ровно из четырех команд. Тест требует сохранение/restore `CustomizationContext`, runtime template context, conflict lookup, удаление только own bindings, внешнее хранение в `user-settings.ini`, восстановление `NormalTemplate.Saved`/runtime `Saved` без `Save`.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature hotkeys`

Expected: FAIL.

- [ ] **Step 3: Реализовать manager/form**

Использовать `KeyBindings.Add KeyCategory:=wdKeyCategoryMacro` с полностью квалифицированными Word command macros. На старте сессии переустановить bindings из settings; при конфликте форма показывает существующую command string и просит подтверждение. Cleanup всегда возвращает исходный customization context/saved flags.

- [ ] **Step 4: Run GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature hotkeys,source-policy`

Expected: PASS.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/modWordHotkeys.bas word_addin/CWordHotkeyTextBoxEvents.cls word_addin/frmWordHotkeySettings.code.txt word_addin/modWordCommands.bas word_addin/modWordRuntime.bas word_addin/tests/features/hotkeys.ps1
git commit -m "feat: add Word add-in hotkeys"
```

---

### Task 8: Реализовать Word Activity Logger без retention

**Files:**
- Create: `word_addin/tests/features/activity-logger.ps1`
- Create: `word_addin/modWordActivityLogStore.bas`
- Create: `word_addin/modWordActivityLogger.bas`
- Create: `word_addin/CWordPendingSave.cls`
- Modify: `word_addin/CWordAppEvents.cls`
- Modify: `word_addin/modWordScheduler.bas`
- Modify: `word_addin/modWordRuntime.bas`

- [ ] **Step 1: Написать RED schema/store tests**

Тест фиксирует 20-колоночный WordV1 header, UTF-16LE BOM/CRLF, filename `<Machine>__<WordSessionId>__YYYY-MM.tsv`, event types `SESSION_START/OPEN/SAVE/SEEN/MISSING/SESSION_END_ATTEMPT`, 300-second heartbeat и отсутствие любых retention/delete/truncate helpers.

Отдельно требовать, что `DocumentBeforeSave`/`DocumentBeforeClose` не содержат `FileDateTime`, `FileLen`, FSO, network root или flush.

- [ ] **Step 2: Написать RED control/flush tests**

Pure helpers проверяют `UploadDefault`, enabled/disabled machines и fail-closed contract. Устаревший local `Yes` + missing network control должен вернуть `False`. Sync-state ведется отдельно на session file.

- [ ] **Step 3: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature activity-logger`

Expected: FAIL.

- [ ] **Step 4: Реализовать logger/store**

`modWordActivityLogger` ведет in-memory document snapshots и pending-save records. Pending verification coalesce-ится sequence, максимум пять попыток/15 секунд через central scheduler; неоднозначный save пропускается. `modWordActivityLogStore` сначала append-ит local spool, затем scheduled activity task читает свежий network `logger-control.ini` и только при allow дописывает новые строки в `NetworkActivityLogRoot\word-computers`.

Ни один runtime path не удаляет/перемещает/обрезает старые файлы.

- [ ] **Step 5: Run GREEN**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature activity-logger,scheduler,source-policy`

Expected: PASS.

- [ ] **Step 6: Коммит**

```powershell
git add word_addin/modWordActivityLogStore.bas word_addin/modWordActivityLogger.bas word_addin/CWordPendingSave.cls word_addin/CWordAppEvents.cls word_addin/modWordScheduler.bas word_addin/modWordRuntime.bas word_addin/tests/features/activity-logger.ps1
git commit -m "feat: add Word activity logger"
```

---

### Task 9: Реализовать `Диагностика`

**Files:**
- Create: `word_addin/tests/features/diagnostics.ps1`
- Create: `word_addin/modWordDiagnostics.bas`
- Create: `word_addin/frmWordDiagnostics.code.txt`
- Modify: `word_addin/modWordCommands.bas`

- [ ] **Step 1: Написать RED diagnostics test**

Требовать version, loaded runtime path, active/pending slots, `Application.StartupPath`, loader path, mode/channel, package/version/assets/activity paths, catalog counts, logger spool/pending/last flush, three retry deadlines, scheduler status и last loader log line. Запретить update/fix/logger buttons.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature diagnostics`

Expected: FAIL.

- [ ] **Step 3: Реализовать read-only form**

Все getters best-effort и не инициируют сеть. Network paths показываются как строки из config. Catalog counts читаются из локального registry/cache. Ошибка одного getter добавляет diagnostic line, но форма открывается.

- [ ] **Step 4: Run GREEN и коммит**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature diagnostics`

```powershell
git add word_addin/modWordDiagnostics.bas word_addin/frmWordDiagnostics.code.txt word_addin/modWordCommands.bas word_addin/tests/features/diagnostics.ps1
git commit -m "feat: add Word add-in diagnostics"
```

---

### Task 10: Реализовать dual-slot Setup/Loader/updater

**Files:**
- Create: `word_addin/tests/features/setup-loader.ps1`
- Create: `word_addin/tests/setup-loader-integration-test.ps1`
- Create: `word_addin_updater/modWordUpdaterVersion.bas`
- Create: `word_addin_updater/modWordUpdaterConfig.bas`
- Create: `word_addin_updater/modWordUpdaterFileSystem.bas`
- Create: `word_addin_updater/modWordAddinUpdater.bas`
- Create: `word_addin_updater/modWordSetupBootstrap.bas`
- Create: `word_addin/build/ThisDocument.loader.txt`
- Create: `word_addin/build/ThisDocument.setup.txt`

- [ ] **Step 1: Написать RED source contract**

Проверить literal company paths, channel-relative package/version, `Application.StartupPath`, slots A/B, `ActiveRuntimeSlot`, `PendingRuntimeSlot`, `LoaderProtocolVersion=1`, three cooldown keys, Loader local-only AutoExec, `Installed=False` before fallback, health-check match, Setup entry points `AutoOpen` и `AutoNew` с единым idempotent session flag и три явные Setup-ветки:

```text
First install: validate/copy slot/config -> copy Loader last as commit point
Reinstall with loaded protocol-v1 Loader: do not overwrite Loader; update inactive slot/config only
Protocol mismatch: stop before any slot/config/Startup mutation and show migration instruction
```

- [ ] **Step 2: Написать RED isolated integration test**

В проверенном temp root создать три fake channel-shaped folders `word-addin-dev`, `word-addin-testers`, `word-addin-stable` с package/version/assets/control. Не удалять ничего вне temp. Cases: `AutoOpen` и `AutoNew` вызывают один общий bootstrap, а idempotent flag не допускает второй mutation/message в той же Word-сессии; first install каждого канала; repeated Setup с уже загруженным protocol-v1 Loader и Word-open model без перезаписи Loader; protocol mismatch без частичной мутации; failure runtime/version/config до Loader commit point без Startup artifact; newer staging to inactive slot; next-start activation; corrupt pending fallback with exactly one installed runtime model; invalid/downgrade version; missing config; read-only Startup; separate cooldown; assets allowlist/catalog-last.

- [ ] **Step 3: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature setup-loader`

Run: `& .\word_addin\tests\setup-loader-integration-test.ps1`

Expected: оба FAIL.

- [ ] **Step 4: Реализовать updater/setup modules**

Loader AutoExec: config -> unload known slot add-ins -> pending health check -> active fallback -> call runtime maintenance hook. Maintenance копирует более новый runtime/version только в inactive slot, ставит pending, синхронизирует `templates/**`, затем registry, и отдельно refresh-ит activity control.

Setup предоставляет оба Word lifecycle entry point: `AutoOpen` для прямого открытия `.dotm` и `AutoNew` для создания документа на его основе. Оба вызывают один bootstrap, который в начале атомарно устанавливает единый session-level idempotent flag и поэтому выполняется не более одного раза в конкретном Word process.

Setup реализует ветвление до записи: при first install валидирует package/version, подготавливает slot/config и копирует stable Loader последним; при protocol-v1 reinstall никогда не перезаписывает loaded Loader и обновляет только inactive/pending slot/config; при protocol mismatch или любой ошибке до commit point прекращает работу без нового Startup Loader.

Никаких delete/move/shell/network clients. Network attempt — одна серия FSO operations после `IsReady`, затем 4h cooldown сервиса.

- [ ] **Step 5: Run GREEN**

Expected outputs:

```text
WORD_SETUP_LOADER_SOURCE_OK
WORD_SETUP_LOADER_INTEGRATION_OK
```

- [ ] **Step 6: Коммит**

```powershell
git add word_addin_updater word_addin/build/ThisDocument.loader.txt word_addin/build/ThisDocument.setup.txt word_addin/tests/features/setup-loader.ps1 word_addin/tests/setup-loader-integration-test.ps1
git commit -m "feat: add Word add-in setup and dual-slot loader"
```

---

### Task 11: Собрать runtime/Loader/Setup и выполнить Word COM smoke

**Files:**
- Create: `word_addin/tests/features/build.ps1`
- Create: `word_addin/tests/smoke-built-word-addin.ps1`
- Create: `word_addin/build/build-audit-word-addin.ps1`
- Create: `word_addin/build/build-audit-word-loader.ps1`
- Create: `word_addin/build/build-audit-word-setup.ps1`
- Create: `word_addin/build/Install-AuditWordAddin.ps1`
- Create: `word_addin/build/Uninstall-AuditWordAddin.ps1`
- Modify: `word_addin/release/AuditWordAddin.dotm`
- Modify: `word_addin/release/AuditWordAddin.version.txt`

- [ ] **Step 1: Написать RED build tests**

Требовать полный import list, временный CP1251 root, programmatic UserForms, customUI/images injection, output/version markers, cleanup temp и отсутствие Excel source changes.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature build`

Expected: FAIL.

- [ ] **Step 3: Реализовать build scripts**

Использовать отдельный невидимый Word COM instance, `AutomationSecurity`, `DisplayAlerts=0`, VBIDE import из CP1251, сохранение `wdFormatXMLTemplateMacroEnabled`. Ribbon внедрять после закрытия Word через `System.IO.Compression`, добавляя `customUI/customUI14.xml`, images, relationships и content types.

Installer разворачивает dev runtime/config/Loader без VBA-пароля. Он не закрывает пользовательский Word: при уже загруженном файле использует inactive/pending slot либо сообщает о restart.

- [ ] **Step 4: Собрать artifacts**

Run: `& .\word_addin\build\build-audit-word-addin.ps1 -OutputPath .\word_addin\release\AuditWordAddin.dotm`

Run: `& .\word_addin\build\build-audit-word-loader.ps1 -OutputPath $env:TEMP\AuditWordAddinLoader.dotm`

Run: `& .\word_addin\build\build-audit-word-setup.ps1 -OutputPath $env:TEMP\AuditWordAddinSetup.dotm`

Expected: `WORD_*_BUILD_OK` и валидные ZIP packages.

- [ ] **Step 5: Выполнить последовательный COM smoke**

Smoke фиксирует все pre-existing `WINWORD.EXE` PID и не закрывает их. Он использует именно уже собранные runtime/Loader/Setup `.dotm`, создает изолированные temp Startup/AppData/channel roots, а затем через реальный Word COM последовательно проверяет:

- прямое открытие собранного Setup (`AutoOpen`) и создание документа на его основе (`AutoNew`) при одном idempotent bootstrap на Word process;
- фактическую first install: runtime/config сначала, собранный Loader в temp Startup последним;
- закрытие только созданного тестом Word process, следующий реальный запуск Word через temp Startup, активацию Loader и ровно одного installed runtime;
- повторный запуск собранного Setup при загруженном protocol-v1 Loader без его перезаписи, затем следующий Word restart с pending slot;
- загрузку runtime как единственной global template и четыре команды;
- `Documents.Add`, view без dirty, hotkeys без Normal save и diagnostics;
- обычный Save и успешный Save As;
- отмененный Save As без ложного `SAVE`;
- повторные pending saves с sequence/coalescing;
- close-before-callback без ложного подтверждения;
- намеренный конфликт/пропуск Word `OnTime` и self-heal на следующем Word event;
- две параллельные созданные тестом Word-сессии с разными TSV/state paths;
- `OPEN/SAVE/SEEN/MISSING` и flush только при свежем разрешающем network control;
- offline/missing/broken network control и stale local `Yes` как fail-closed без flush;
- закрытие только созданных тестом Word processes и сохранность всех исходных PID.

Run: `& .\word_addin\tests\smoke-built-word-addin.ps1`

Expected: `WORD_ADDIN_BUILT_SMOKE_OK`.

- [ ] **Step 6: Коммит**

```powershell
git add word_addin/build word_addin/tests word_addin/release
git commit -m "build: add Word add-in packages and COM smoke"
```

---

### Task 12: Реализовать единую `!Журнал активности Office.xlsm`

**Files:**
- Create: `word_addin/tests/features/office-activity-viewer.ps1`
- Create: `word_addin/tests/office-activity-viewer-integration-test.ps1`
- Modify: `word_addin/tests/fixtures/New-WordAddinFixtures.ps1`
- Create: `word_addin/admin_activity_log/modOfficeActivityLogImport.bas`
- Create: `word_addin/admin_activity_log/modOfficeActivityLogViewer.bas`
- Create: `word_addin/admin_activity_log/ThisWorkbook.office-activity-log-viewer.txt`
- Create: `word_addin/build/build-office-activity-log-viewer.ps1`

- [ ] **Step 1: Написать RED mixed-schema tests**

Fixture generator пишет фактические UTF-16LE BOM/CRLF ExcelV1 и WordV1 TSV. Cases: valid rows, partial tail, broken header, UTF-8 BOM forward fixture, no-BOM reject, duplicate within app/machine, same EventId across app/machine, same session ids across app/machine.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature office-activity-viewer`

Expected: FAIL.

- [ ] **Step 3: Реализовать importer/mapper**

Importer читает `computers/*.tsv` как ExcelV1 и `word-computers/*.tsv` как WordV1, BOM-detects, maps headers by name и создает normalized row с `OfficeApp`, `SourceSchema`, common document/session fields. Ключи:

```text
EventKey    = OfficeApp|MachineName|EventId
OfficeKey   = OfficeApp|MachineName|OfficeSessionId
DocumentKey = OfficeApp|MachineName|DocumentSessionId
```

- [ ] **Step 4: Реализовать viewer/admin controls**

Листы: `Dashboard`, `RawEvents`, `OfficeSessions`, `DocumentSessions`, `Files`, `Users`, `Computers`, `Diagnostics`. `DocumentSessions` использует составной document-session key, а `Files` является отдельным file-level отчетом по нормализованному пути; оба содержат и фильтруют `OfficeApp`. Control editor пишет общий UTF-16LE `logger-control.ini`; обычный refresh его не меняет.

- [ ] **Step 5: Реализовать Excel COM builder и integration test**

Builder принимает `-ActivityLogRoot -OutputPath -OpenPassword`, создает `.xlsm` из новой Word-линии и не читает/не меняет legacy VBA. Integration test строит workbook с test password, открывает его, вызывает `OfficeActivityLogViewer_RefreshAll`, проверяет mixed counts/keys/sheets и закрывает только свой Excel process.

Run: `& .\word_addin\tests\office-activity-viewer-integration-test.ps1`

Expected: `OFFICE_ACTIVITY_VIEWER_INTEGRATION_OK`.

- [ ] **Step 6: Коммит**

```powershell
git add word_addin/admin_activity_log word_addin/build/build-office-activity-log-viewer.ps1 word_addin/tests
git commit -m "feat: add unified Excel and Word activity viewer"
```

---

### Task 13: Добавить dev/company release-пайплайны и policy tests

**Files:**
- Create: `word_addin/tests/features/release-structure.ps1`
- Create: `word_addin/tests/company-package-policy-smoke-test.ps1`
- Create: `word_addin/tests/vba-protection-smoke-test.ps1`
- Create: `word_addin/build/publish-word-addin-release.ps1`
- Create: `word_addin/build/publish-company-word-addin-release.ps1`
- Create: `word_addin/build/publish-office-activity-company-release.ps1`
- Create: `word_addin/build/protect-vba-project.py`

- [ ] **Step 1: Написать RED release tests**

Проверить exact layouts `company/На рабочий комп/Не трогать`, `Результаты проверок`, archives, three channel-agnostic artifacts, admin activity resource layout и desktop mirror paths. Package policy сканирует готовые `.dotm/.xlsm`, protection test проверяет VBA project lock.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature release-structure`

Expected: FAIL.

- [ ] **Step 3: Реализовать publishers**

Dev publisher обновляет только `word_addin/release/AuditWordAddin.dotm` и version.

Word company publisher требует явное `-AclConfirmed` и выполняет строго упорядоченный pipeline:

1. source/features PASS;
2. собрать staging runtime/Loader/Setup из dev source;
3. поставить VBA protection на все три package;
4. один раз прогнать package policy по защищенному staging;
5. один раз прогнать VBA protection smoke;
6. потребовать ручное подтверждение ACL для channel packages, `word-assets`, общего control и raw-log каталогов и записать его в timestamped human-readable report;
7. сформировать остальные timestamped human-readable reports;
8. только после всех PASS и ACL confirmation record архивировать предыдущий current release;
9. заменить `company/На рабочий комп` и `Результаты проверок`;
10. только после replacement обновить desktop mirror и zip;
11. объявить success.

Office activity publisher требует `-OpenPassword`, `-VbaPassword` и явное `-AclConfirmed`. Он выполняет строго:

1. CP1251 import + Excel COM build в staging;
2. open-password verification;
3. VBA protection и protection verification;
4. mixed UTF-16LE ExcelV1/WordV1 import/dedupe/session/files checks;
5. package/source policy и `git diff` guard для `addin/`, `addin_updater/`, `shared/`, `profiles/`, legacy viewer;
6. timestamped reports и ACL confirmation record;
7. только после PASS архивировать предыдущий Office activity current release;
8. заменить `office-activity-company/Ресурсы для общего склада` и reports;
9. только затем обновить desktop mirror;
10. объявить success.

Любой FAIL до шага archive оставляет current release/mirror неизменными. Office publisher кладет только `Ресурсы для общего склада/assets/_system/activity-logs/!Журнал активности Office.xlsm` плюс reports.

Ни один publisher не пишет на `Z:` автоматически.

- [ ] **Step 4: Проверить source/pipeline без фактического company release**

Release tests дополнительно проверяют обязательный `-AclConfirmed` и порядок pipeline markers: protect -> policy -> protection smoke -> ACL -> reports -> archive -> replace -> mirror -> success, а для Office viewer — build -> open-password -> protect -> mixed-import -> source guard -> ACL -> archive -> replace -> mirror -> success.

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature release-structure,source-policy`

Expected: PASS. Не запускать company publishers без пользовательских паролей/команды.

- [ ] **Step 5: Коммит**

```powershell
git add word_addin/build word_addin/tests
git commit -m "build: add Word company release pipelines"
```

---

### Task 14: Обновить runbooks и системную архитектуру

**Files:**
- Create: `docs/word-addin-agent-instructions.md`
- Create: `docs/testers-word-addin-install.md`
- Modify: `docs/system-overview.md`
- Modify: `word_addin/README.md`

- [ ] **Step 1: Написать docs assertions в `structure.ps1`**

Требовать literal каналы, word-assets, common activity root, local paths, dual slots, no-retention, exactly four Ribbon commands, build/test/install commands и запрет менять Excel line.

- [ ] **Step 2: Run RED**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature structure`

Expected: FAIL на недостающих docs.

- [ ] **Step 3: Написать документацию**

Runbook должен быть самостоятельным: карта путей, dev/company, setup/loader lifecycle, template catalog schema, Office logger/viewer, known OnTime/Windows timeout limitations, release rules, troubleshooting и тестовые entry points.

- [ ] **Step 4: Run GREEN и коммит**

```powershell
& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature structure
git add docs word_addin/README.md word_addin/tests/features/structure.ps1
git commit -m "docs: add Word add-in runbooks"
```

---

### Task 15: Полная проверка, локальная dev-установка и handoff

**Files:**
- Update: `word_addin/release/AuditWordAddin.dotm`
- Update: `word_addin/release/AuditWordAddin.version.txt`
- No changes allowed: `addin/`, `addin_updater/`, `shared/`, `profiles/`

- [ ] **Step 1: Запустить Word fast suite**

Run: `& .\word_addin\tests\Invoke-WordAddinTests.ps1 -Feature all-fast`

Expected: все Word features PASS.

- [ ] **Step 2: Запустить integration/COM suites последовательно**

```powershell
& .\word_addin\tests\setup-loader-integration-test.ps1
& .\word_addin\tests\smoke-built-word-addin.ps1
& .\word_addin\tests\office-activity-viewer-integration-test.ps1
```

Expected: три `*_OK`, без новых зависших Word/Excel processes.

- [ ] **Step 3: Проверить неизменность Excel-линии**

Run: `git diff --exit-code main...HEAD -- addin addin_updater shared profiles`

Run: `& .\addin\tests\Invoke-AddinTests.ps1 -Feature source-policy`

Expected: empty diff и `ADDIN_SOURCE_POLICY_OK`. Не использовать падающий baseline `all-fast` как Word acceptance; отдельно записать известный исходный sampling failure.

- [ ] **Step 4: Пересобрать финальный dev package и установить локально**

```powershell
& .\word_addin\build\publish-word-addin-release.ps1
& .\word_addin\build\Install-AuditWordAddin.ps1 -DeploymentMode Dev -NetworkEnabled No
```

Если пользовательский Word открыт, не закрывать его автоматически; installer должен безопасно поставить pending slot и сообщить о restart.

- [ ] **Step 5: Запустить review/verification skills**

Использовать `@requesting-code-review`, исправить реальные замечания, затем `@verification-before-completion`. Повторить измененные тесты после каждой правки.

- [ ] **Step 6: Финальный коммит feature branch**

```powershell
git add word_addin word_addin_updater docs
git commit -m "feat: add Word auditors add-in"
git status --short -uall
```

Expected: worktree clean.

- [ ] **Step 7: Завершить ветку**

Использовать `@finishing-a-development-branch` для выбора merge/PR. После интеграции выполнить корневой `tools/Assert-CleanWorktree.ps1` в основном workspace и убедиться, что внешний и вложенный репозитории чисты.
