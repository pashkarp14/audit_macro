# Дизайн Word-надстройки `Аудиторам`

Дата: 2026-07-17

Статус: дизайн согласован в чате; ожидает документального ревью и плана реализации.

## Контекст

Нужна первая отдельная надстройка `Аудиторам` для настольного Microsoft Word на Windows. Надстройка создается на VBA в формате `.dotm`, использует собственный Ribbon и не встраивается в существующую Excel-надстройку.

Word-линия должна повторять проверенную схему доставки Excel-линии: локальный runtime, легкий loader в Word Startup, отдельный setup и три сетевых канала `dev`, `testers`, `stable`. При этом код, пакеты, настройки и ресурсы Word не должны менять или переиспользовать runtime-файлы Excel-надстройки.

Первая версия содержит четыре видимые команды и один скрытый фоновый сервис Activity Logger. Количество и виды шаблонов рабочих документов пока неизвестны, поэтому каталог шаблонов должен быть управляемым данными, а не зашитым в `.dotm`.

## Цели

- Создать самостоятельную Word-надстройку `AuditWordAddin.dotm`.
- Добавить на вкладку `Аудиторам` ровно четыре согласованные Ribbon-команды.
- Создать установку и автообновление по схеме `Setup -> Word Startup Loader -> local runtime`.
- Поддержать сетевые каналы `word-addin-dev`, `word-addin-testers`, `word-addin-stable` одним channel-agnostic пакетом.
- Поддержать произвольный каталог Word-шаблонов через TSV-реестр.
- Добавить скрытый Activity Logger с теми же типами действий, которые фактически отслеживает Excel-логгер.
- Сохранить fail-open запуск: локальный runtime загружается без сетевых вызовов; сетевые сервисы выполняют одну best-effort попытку и затем отдельный cooldown. Первый системный доступ к недоступному mapped drive может ждать Windows timeout — это явно раскрываемое и требующее подтверждения ограничение чистого VBA.
- Сразу предусмотреть dev-сборку без VBA-пароля и отдельный company-пайплайн с VBA-защитой.

## Не Цели Первой Версии

- Не менять `addin/`, `addin_updater/`, `AuditAddin.xlam`, Excel Ribbon, Excel Activity Logger или Excel admin-книгу.
- Не добавлять очистку текста, оформление таблиц, экспорт в PDF, проверку документа, инструкцию, конструктор РД или другие пользовательские команды.
- Не добавлять видимую кнопку Activity Logger или дополнительные Ribbon-пункты.
- Не записывать текст документов, нажатия клавиш, изменения текста, выделения, перемещения курсора, печать или содержимое буфера обмена.
- Не вводить ограничение хранения на шесть месяцев, автоматическую ротацию, архивирование или очистку Word-логов.
- Не фиксировать заранее перечень видов аудита, этапов, кодов РД или количество шаблонов.
- Не использовать интернет, Git, shell, PowerShell, `cmd.exe`, `WScript.Shell`, HTTP-клиенты или внешние updater-процессы в runtime.
- Не изменять `Normal.dotm`, глобальные настройки безопасности Word или документы пользователя ради хранения настроек надстройки.
- Не публиковать company-пакет в сетевые каналы без отдельной команды на company-release.

## Целевая Платформа

- Настольный Microsoft Word для Windows с VBA и RibbonX `customUI14`.
- Поддержка 32-bit и 64-bit Office: runtime не использует Win32 API declarations.
- Word for Mac и Word Online не входят в scope.
- Dev-сборщик требует локальный Microsoft Word и включенный на машине разработчика `Trust access to the VBA project object model`; конечным пользователям этот доступ не нужен.
- Setup/runtime требуют разрешенного корпоративной политикой запуска подписанных/доверенных macro-enabled templates; код не ослабляет политику Office.

## Принятое Архитектурное Решение

Используется отдельная Word-линия со своими исходниками, пакетами, локальным корнем и сетевыми ресурсами.

```text
Word Setup
    -> копирует Word Loader в Application.StartupPath
    -> копирует Word runtime в %APPDATA%\AuditWordAddin\installed
    -> записывает локальный конфиг

Word startup
    -> автоматически загружает Loader из Application.StartupPath
    -> Loader загружает активный или pending runtime из одного из двух локальных слотов
    -> runtime best-effort ставит обновление в неактивный слот для следующего старта
    -> runtime best-effort синхронизирует разрешенные Word assets

Word runtime
    -> показывает четыре команды Ribbon
    -> работает только с локальным кэшем ресурсов
    -> поднимает события вида документа, хоткеев и Activity Logger
```

Загрузка `.dotm` из Word Startup и управление global templates опираются на штатную модель Word: [`Application.StartupPath`](https://learn.microsoft.com/en-us/office/vba/api/word.application.startuppath) и объект [`AddIn`](https://learn.microsoft.com/en-us/office/vba/api/word.addin).

## Структура Репозитория

Новая линия располагается во вложенном репозитории `_publish/audit-macros-private`:

```text
word_addin/
  build/
  assets/
  admin_activity_log/
  tests/
    features/
    fixtures/
  release/
  *.bas
  *.cls
  *.frm
  *.frx
  *.code.txt

word_addin_updater/
  tests/
  *.bas
  *.code.txt
```

`word_addin/` не является профилем обычного macro updater-а. Модули Word не помещаются в `shared/` или `profiles/`.

Исходники VBA хранятся в UTF-8. Для импорта через VBE сборочные скрипты создают временные ANSI/CP1251-копии, чтобы русские строки не портились. Временные копии не коммитятся.

## Сетевые Каналы

Общий company-root остается рядом с Excel-каналами:

```text
Z:\Project\DAVK\Оборудование\Addins\
```

Word-каналы:

```text
Z:\Project\DAVK\Оборудование\Addins\word-addin-dev\
Z:\Project\DAVK\Оборудование\Addins\word-addin-testers\
Z:\Project\DAVK\Оборудование\Addins\word-addin-stable\
```

Каждый канал имеет одинаковую структуру:

```text
AuditWordAddinSetup.dotm
Не трогать\
  AuditWordAddin.dotm
  AuditWordAddin.version.txt
  AuditWordAddinLoader.dotm
```

Пакет channel-agnostic. Setup формирует пути обновления из собственной папки запуска:

```text
<SetupFolder>\Не трогать\AuditWordAddin.dotm
<SetupFolder>\Не трогать\AuditWordAddin.version.txt
```

Поэтому конечные пользователи запускают Setup из фактической папки выбранного сетевого канала. Имя канала нужно только для диагностики; обновляющий код не содержит трех отдельных hardcoded-пакетов.

## Сетевые Word-Ресурсы

Общий склад Word-ресурсов отделен от каналов runtime:

```text
Z:\Project\DAVK\Оборудование\Addins\word-assets\
  templates\
    <произвольная структура шаблонов>
  sync\
    word-rd-template-map.tsv
  _system\
    activity-logs\
      computers\
        <MachineName>__<WordSessionId>__YYYY-MM.tsv
      logger-control.ini
      !Журнал активности Word.xlsm
```

Папка `templates\` может быть пустой в первой сборке. Ни количество шаблонов, ни их иерархия не являются частью контракта runtime. Сами измерения выбора `вид аудита -> этапность/вариант -> код РД` зафиксированы исходным пользовательским сценарием; динамическими остаются их значения и число шаблонов.

Синхронизация работает по allowlist и в таком порядке:

1. рекурсивно копирует `templates\`;
2. копирует `_system\activity-logs\logger-control.ini`;
3. последним копирует `sync\word-rd-template-map.tsv`, чтобы новый каталог не ссылался на еще не скопированные шаблоны.

Никакие другие файлы `word-assets` в пользовательский кэш не попадают. В частности, не копируются `_system\activity-logs\computers\`, `!Журнал активности Word.xlsm` и чужие сетевые логи.

## Локальные Пути

Авторитетный Word Startup определяется во время выполнения через `Application.StartupPath`, а не через захардкоженную строку `%APPDATA%\Microsoft\Word\STARTUP`.

```text
<Application.StartupPath>\AuditWordAddinLoaderLocal.dotm

%APPDATA%\AuditWordAddin\installed\slots\A\AuditWordAddinLocal.dotm
%APPDATA%\AuditWordAddin\installed\slots\A\AuditWordAddin.version.txt
%APPDATA%\AuditWordAddin\installed\slots\B\AuditWordAddinLocal.dotm
%APPDATA%\AuditWordAddin\installed\slots\B\AuditWordAddin.version.txt
%APPDATA%\AuditWordAddin\word-addin-updater.ini
%APPDATA%\AuditWordAddin\word-addin-updater.log
%APPDATA%\AuditWordAddin\user-settings.ini
%APPDATA%\AuditWordAddin\assets\
%APPDATA%\AuditWordAddin\activity-log\spool\
%APPDATA%\AuditWordAddin\activity-log\state\
```

Настройки вида и горячих клавиш лежат в `user-settings.ini`; они не записываются в открытые документы и не сохраняются в `Normal.dotm`.

## Dev И Company Упаковка

Одна dev-кодовая база используется для обоих вариантов:

- dev runtime: `word_addin/release/AuditWordAddin.dotm`, без VBA-пароля;
- company runtime: тот же проверенный runtime под VBA-паролем;
- company package: защищенный runtime плюс Loader и Setup.

Личный сетевой канал `word-addin-dev` является company-style предрелизом: он имеет Setup/Loader и защищенный runtime, как `testers` и `stable`. Локальная разработческая установка на машине Павла использует незапароленный dev runtime.

Локальная структура company-release:

```text
word_addin/release/company/
  На рабочий комп/
    AuditWordAddinSetup.dotm
    Не трогать/
      AuditWordAddin.dotm
      AuditWordAddin.version.txt
      AuditWordAddinLoader.dotm
  Результаты проверок/

word_addin/release/company-archive/<version>/
```

Зеркало актуального company-комплекта на этой машине:

```text
C:\Users\pavel\OneDrive\Рабочий стол\Релиз надстройки Аудиторам Word\
```

Оно создается только при явном company-release и содержит актуальный комплект, результаты проверок и `На рабочий комп.zip`. Старые версии уходят в repo archive, а не накапливаются в зеркале рабочего стола.

Word-assets публикуются отдельным asset/template-release и не смешиваются с обычным runtime-пакетом. Реальные шаблоны не создаются и не придумываются до их передачи Павлом.

## Установка

`AuditWordAddinSetup.dotm` является одноразовым bootstrap-пакетом. Он должен корректно отработать как при открытии шаблона, так и при создании Word нового документа на его основе; повторные lifecycle-вызовы защищены idempotent-флагом.

Setup:

1. определяет собственную папку;
2. проверяет наличие подпапки `Не трогать` и трех обязательных файлов;
3. получает фактический `Application.StartupPath`;
4. создает контролируемые локальные папки, если их нет;
5. валидирует source version и package до изменения Word Startup;
6. при первой установке копирует runtime/version в слот `A`, проверяет копии и готовит его как active;
7. при повторной установке копирует runtime/version в неактивный слот, проверяет копии и готовит его как pending;
8. полностью записывает и перечитывает config с путями выбранного канала и Word assets;
9. только после готового slot/config при первой установке копирует стабильный Loader protocol v1 как `AuditWordAddinLoaderLocal.dotm`; Loader является commit point установки;
10. сообщает об успешной установке и необходимости перезапустить Word.

Setup не изменяет Trust Center, `Normal.dotm`, реестр безопасности Office и параметры корпоративной политики. Если политика Word блокирует VBA из сетевой папки, это внешнее условие установки; код не пытается его обходить.

Loader v1 намеренно минимален и стабилен: он только выбирает локальный слот, загружает `.dotm` и вызывает health check. Повторный Setup не перезаписывает уже загруженный Loader с тем же protocol version. Если будущая версия потребует смены Loader protocol, это отдельная миграция вне текущего scope; Setup v1 обязан остановиться до частичной установки и дать понятную инструкцию, а не пытаться перезаписать заблокированный Startup-файл.

Повторный Setup безопасно обновляет runtime/config через неактивный слот. Обычные runtime-обновления после установки также используют slot-механику и не требуют перезаписи Loader-а.

## Конфиг Loader-а

`word-addin-updater.ini` содержит фактические абсолютные значения, записанные Setup:

```text
DeploymentMode=Dev|Work
Channel=Dev|Testers|Stable|Custom
NetworkEnabled=Yes|No
UpdateMode=Auto
NetworkPackagePath=<channel>\Не трогать\AuditWordAddin.dotm
NetworkVersionPath=<channel>\Не трогать\AuditWordAddin.version.txt
NetworkAssetsRoot=Z:\Project\DAVK\Оборудование\Addins\word-assets
LocalRuntimeRoot=%APPDATA%\AuditWordAddin\installed\slots
ActiveRuntimeSlot=A|B
PendingRuntimeSlot=A|B|
LoaderProtocolVersion=1
AssetsRoot=%APPDATA%\AuditWordAddin\assets
RdTemplateRoot=%APPDATA%\AuditWordAddin\assets\templates
RuntimeRetryAfter=
AssetsRetryAfter=
ActivityLogRetryAfter=
LastNetworkCheckAt=
LastSuccessfulUpdateAt=
LastAssetsSyncAt=
```

В реальном файле переменные окружения раскрыты до абсолютных путей. Пустой или некорректный обязательный путь считается недоступным и не отправляется в FSO/Word API.

## Обновление При Старте Word

Версия runtime имеет формат `YYYY.MM.DD.HHMMSS`, например `2026.07.17.183240`. Version-файл содержит одну такую строку. Пустая, пониженная, синтаксически битая или не соответствующая health check версия не активируется. Поле `revision` в диагностике равно этой version string; отдельного revision-файла нет.

Word автоматически загружает `AuditWordAddinLoaderLocal.dotm` из Startup. При отсутствующем/битом config Loader пишет локальную диагностику, не пытается угадать пути и тихо завершает работу. При валидном config Loader выполняет легкий `AutoExec` без сетевых обращений:

1. читает локальный config;
2. если указан pending slot, сначала пытается загрузить его через Word AddIns;
3. вызывает обязательный `AuditWordAddin_HealthCheck` и сверяет ответ с version-файлом слота;
4. при успехе делает pending slot активным;
5. при ошибке явно устанавливает failed add-in `Installed=False`, проверяет, что он выгружен, и очищает pending state;
6. только после подтвержденной выгрузки failed pending загружает предыдущий active slot;
7. если active slot битый, тем же unload-before-fallback протоколом пробует второй валидный слот;
8. если не удалось гарантировать выгрузку неуспешного runtime, прекращает загрузку и не допускает две одноименные global templates/Ribbon;
9. после успешной локальной загрузки передает runtime задачи проверки обновления и assets sync.

Перед выбором слота Loader перечисляет оба известных slot paths и добивается состояния, в котором установлен не более чем один Word runtime. `AddIn.Delete` и удаление файлов не используются; управление выполняется только через `AddIn.Installed`.

Runtime update task выполняется уже после появления локальной надстройки в Word:

1. соблюдает отдельный `RuntimeRetryAfter`;
2. делает не более одного набора сетевых package/version проверок;
3. сравнивает валидные version strings без автоматического downgrade;
4. копирует более новый package/version только в неактивный слот;
5. проверяет наличие, ненулевой разумный размер, ZIP-сигнатуру `.dotm` и совпадение version-файла;
6. только после полного копирования записывает `PendingRuntimeSlot`;
7. активирует обновление на следующем старте Word через health check.

Таким образом, обновление Word runtime применяется с задержкой в один перезапуск. Незавершенное копирование затрагивает только неактивный слот и не уничтожает последнюю пригодную версию.

Assets task имеет независимый `AssetsRetryAfter`, работает по allowlist и `DateLastModified`. Logger flush имеет независимый `ActivityLogRetryAfter`. Ошибка одного сетевого сервиса не отключает два других.

Отказ сети не является отказом Word-надстройки. Setup гарантирует наличие первоначального локального runtime. Если обновление не удалось, используется предыдущая локальная копия.

Loader/runtime не удаляют и не перемещают файлы. Для обновления разрешено только контролируемое копирование в неактивный слот. Assets обновляются без удаления локальных файлов: новые и более свежие сетевые файлы копируются, отсутствующие в сети локальные файлы остаются.

## Единый Планировщик Word

Word может поддерживать только один background timer `Application.OnTime`; постановка нового таймера отменяет предыдущий. Поэтому надстройка не создает независимые таймеры для logger-а, pending save и синхронизации. Это ограничение зафиксировано в [документации Microsoft](https://learn.microsoft.com/en-us/office/vba/api/word.application.ontime).

В runtime существует один централизованный scheduler:

- все внутренние задачи лежат в in-memory очереди с `DueAt` и типом задачи;
- один полностью квалифицированный callback планируется на ближайший `DueAt`;
- callback обрабатывает просроченные задачи, сохраняет незавершенные и ставит следующий единственный timer;
- update check, assets sync, pending-save verification, heartbeat и logger flush не вызывают `Application.OnTime` напрямую;
- повторные задачи одного типа coalesce-ятся, чтобы очередь не росла;
- на `DocumentOpen`, `NewDocument`, `DocumentBeforeSave`, `DocumentBeforeClose`, `WindowActivate` и каждом Ribbon callback scheduler проверяет пропущенные `DueAt` и self-heal-ит очередь.

Если другая Word-надстройка после нас поставит собственный `OnTime`, наш callback может быть отменен. Если наш scheduler поставит callback последним, он может отменить чужой timer. Word не предоставляет API обнаружения или совместного владения этим единственным timer-ом. Поэтому scheduler используется только когда есть pending-save, открытые отслеживаемые документы или отложенная startup-задача; отсутствие callback считается best-effort пропуском, а не причиной блокировки Word. Пропущенные задачи выполняются при следующем событии/команде, а `SEEN` остается приблизительным.

Win32 `SetTimer`, бесконечные `DoEvents`-циклы и внешние timer-процессы не используются из-за риска reentrancy/crash и company-ограничений.

## Ribbon

Используется `customUI14`. На ленте создается отдельная вкладка `Аудиторам`, размещенная после вкладки рецензирования Word.

Видимы ровно четыре Ribbon-контрола с `onAction`:

```text
Аудиторам
  Рабочие документы
    Открыть шаблон РД

  Вид документа
    Настроить вид

  Клавиши
    Горячие клавиши

  Сервис
    Диагностика
```

Для четырех основных команд используются отдельные осмысленные PNG-иконки. Разрешено скопировать подходящие исходные изображения из `addin/assets` в `word_addin/assets`, не меняя Excel Ribbon и его assets.

`Настроить вид` является одной кнопкой и открывает единую форму. Настройки `При открытии`, `При закрытии` и две операции `Применить сейчас` находятся внутри формы, а не являются отдельными Ribbon-командами.

Activity Logger не получает кнопку, menu item или пользовательскую панель.

## Команда `Открыть шаблон РД`

Команда читает только локальный кэш:

```text
%APPDATA%\AuditWordAddin\assets\sync\word-rd-template-map.tsv
%APPDATA%\AuditWordAddin\assets\templates\
```

Сетевых обращений в on-demand сценарии нет.

### Формат Реестра

`word-rd-template-map.tsv` хранится в UTF-8 с BOM и имеет обязательные колонки:

```text
TemplateId
AuditType
StageMode
StageVariant
RdCode
DisplayName
RelativePath
Enabled
```

Правила:

- `TemplateId` уникален и стабилен;
- комбинация `AuditType + StageMode + StageVariant + RdCode` не дублируется среди включенных строк;
- имена обязательных колонок сравниваются без учета регистра, неизвестные дополнительные колонки игнорируются;
- значения обрезаются по краям; `Enabled=Yes` включает строку, `No` скрывает без учета регистра;
- `RelativePath` задается относительно `templates\`;
- runtime строит абсолютный candidate через FSO, канонизирует root/candidate и проверяет root-boundary с завершающим `\`; абсолютные пути, `..` и любой выход за корень шаблонов запрещены;
- разрешены `.dotx` и `.dotm`;
- отсутствующий файл не ломает весь каталог, но строка попадает в диагностику;
- битый заголовок или невозможность прочитать реестр дают быстрое понятное сообщение.

Форма выбора строит значения динамически в последовательности:

1. вид аудита;
2. этапность/вариант;
3. код и отображаемое имя РД.

Никакие варианты не зашиваются в код. Если активных шаблонов нет, команда показывает понятное сообщение и завершается.

После выбора создается новый несохраненный документ через `Documents.Add` на основе локального `.dotx`/`.dotm`. Исходный шаблон не открывается для редактирования и не изменяется. Это штатная модель Word [`Documents.Add`](https://learn.microsoft.com/en-us/office/vba/api/word.documents.add).

## Команда `Настроить вид`

Одна Ribbon-кнопка открывает единую панель с двумя наборами настроек: `При открытии` и `При закрытии`.

Для каждого набора доступны:

- включить/выключить автоматическое применение;
- масштаб: число в допустимом диапазоне или `Не менять`;
- режим: `Не менять`, `Разметка печати`, `Черновик`, `Веб-документ`, `Структура`, `Режим чтения`;
- область навигации: `Не менять`, `Показать`, `Скрыть`;
- знаки форматирования: `Не менять`, `Показать`, `Скрыть`.

Безопасные дефолты:

- автоматизация при открытии выключена;
- автоматизация при закрытии выключена;
- все поля равны `Не менять`.

Кнопки `Применить открытие сейчас` и `Применить закрытие сейчас` применяют соответствующий профиль к активному окну без изменения сохраненных дефолтов.

Автоматическое применение использует Word Application events `DocumentOpen`, `NewDocument` и `DocumentBeforeClose`. Настройки применяются к окну документа, а не глобально ко всем документам Word. API вида основан на Word [`View`](https://learn.microsoft.com/en-us/office/vba/api/word.view), [`View.Type`](https://learn.microsoft.com/en-us/office/vba/api/word.view.type), [`View.Zoom`](https://learn.microsoft.com/en-us/office/vba/api/word.view.zoom) и области навигации [`Window.DocumentMap`](https://learn.microsoft.com/en-us/office/vba/api/word.window.documentmap).

Критический контракт: применение вида не изменяет содержимое документа, не сбрасывает `Document.Saved` и не создает запрос на сохранение при закрытии.

## Команда `Горячие клавиши`

Единая панель разрешает назначать сочетания только четырем поддерживаемым Ribbon-командам:

- `Открыть шаблон РД`;
- `Настроить вид`;
- `Горячие клавиши`;
- `Диагностика`.

Используется Word [`KeyBindings`](https://learn.microsoft.com/en-us/office/vba/api/word.keybindings), а не Excel `Application.OnKey`.

Перед добавлением binding код запоминает текущий [`CustomizationContext`](https://learn.microsoft.com/en-us/office/vba/api/word.global.customizationcontext), временно устанавливает его в локально загруженный runtime template и гарантированно восстанавливает прежнее значение в cleanup-блоке. `Normal.dotm` не используется как контекст и не сохраняется.

Сочетания хранятся в `%APPDATA%\AuditWordAddin\user-settings.ini` и повторно применяются в памяти при каждой сессии Word. Runtime `.dotm` не сохраняется после настройки клавиш.

Панель:

- показывает текущее сочетание;
- обнаруживает конфликт с существующей Word-командой;
- требует явного подтверждения перед перекрытием конфликта;
- умеет снять только binding, созданный Word-надстройкой;
- не удаляет и не переписывает сочетания других templates/надстроек;
- восстанавливает исходные `Saved`-флаги runtime template и `NormalTemplate`, не выполняя их `Save`.

Закрытие Word после назначения клавиш не должно показывать запрос на сохранение `Normal.dotm` или runtime `.dotm`.

## Команда `Диагностика`

Команда показывает read-only панель с базовой информацией:

- версия Word-надстройки (она же revision);
- полный путь загруженного `AuditWordAddinLocal.dotm`, active/pending slot и версия каждого слота;
- фактический `Application.StartupPath`;
- ожидаемый и найденный путь Loader-а;
- `DeploymentMode` и `Channel`;
- сетевые package/version paths;
- локальный runtime path;
- `NetworkAssetsRoot`, `AssetsRoot`, `RdTemplateRoot`;
- путь и статус `word-rd-template-map.tsv`;
- количество активных, найденных и отсутствующих шаблонов;
- состояние Activity Logger;
- локальный spool и сетевой computer-log root;
- последняя успешная выгрузка и наличие невыгруженных строк;
- три независимых retry-after и состояние единого scheduler-а;
- последняя содержательная строка loader log.

Диагностика ничего не исправляет и не обновляет. В ней нет кнопок updater-а, Activity Logger, конструктора или других функций.

## Activity Logger

Activity Logger является скрытым best-effort сервисом. Он повторяет фактический уровень событий Excel-логгера, но использует Word-объекты и отдельное хранилище.

Типы событий:

- `SESSION_START`: Word-логгер поднят в текущей сессии;
- `OPEN`: сохраненный документ замечен открытым;
- `SAVE`: best-effort подтверждено сохранение документа по состоянию Word и файла;
- `SEEN`: heartbeat подтверждает, что документ остается открытым;
- `MISSING`: ранее отслеживаемый документ закрыт или больше не виден;
- `SESSION_END_ATTEMPT`: Word или runtime начали завершение.

Heartbeat имеет целевой интервал 300 секунд только при наличии отслеживаемых документов. Он идет через единый scheduler, перечисляет `Application.Documents`, пишет легкие локальные события и пытается выгрузить накопленные строки. Из-за единственного Word `OnTime` фактический интервал приблизительный; пропущенный timer восстанавливается при следующем Word event или Ribbon callback.

Несохраненные документы без пути не логируются. При первом успешном сохранении такого документа logger сначала backfill-ит `OPEN`, затем пишет `SAVE`, как Excel-реализация.

Word не имеет `DocumentAfterSave`, поэтому абсолютной гарантии фиксации каждого успешного сохранения нет. Logger использует следующую best-effort state machine:

1. `DocumentBeforeSave` создает или обновляет один pending-save record для объекта документа;
2. record удерживает только in-memory данные: ссылку на `Document`, внутренний document token, `SaveAsUI`, исходные `FullName`, `Saved`, ранее кэшированный logger-ом file snapshot и время события; `DocumentBeforeSave` не вызывает `FileDateTime`, `FileLen`, FSO или сеть;
3. единый scheduler ставит проверку примерно через одну секунду и уже вне save-event читает новый file snapshot; если Word еще выполняет background save, допускается до пяти проверок в пределах 15 секунд;
4. сохранение считается подтвержденным, если появился валидный файл и обнаружен хотя бы один надежный переход: новый/изменившийся `FullName`, изменение времени/размера файла или переход `Saved=False -> True`;
5. для первого сохранения несохраненного документа пишутся backfill `OPEN`, затем `SAVE`;
6. если документ закрыт, Save As отменен, признаки противоречивы или истекло окно проверки, `SAVE` не пишется, а причина остается только в локальной диагностике;
7. повторные `DocumentBeforeSave` одного документа coalesce-ятся с увеличением sequence, чтобы старый callback не подтвердил более новую попытку.

Успешные обычный Save и Save As являются обязательными smoke-сценариями. Для AutoSave, облачных/синхронизируемых путей и сохранения сторонним VBA точность остается best-effort; ложный `SAVE` хуже пропущенного, поэтому неоднозначное состояние пропускается. События Word Application подключаются через отдельный class module с `WithEvents`; это соответствует штатной модели [Application events](https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word) и [`DocumentBeforeSave`](https://learn.microsoft.com/en-us/office/vba/api/word.application.documentbeforesave).

Logger игнорирует:

- runtime, Loader и Setup Word-надстройки;
- `Normal.dotm` и загруженные global templates;
- admin-книгу журнала;
- файлы логов;
- несохраненные документы без пути;
- служебные документы внутри контролируемых папок `%APPDATA%\AuditWordAddin`.

### Поля TSV

```text
EventId
EventAt
EventType
WordSessionId
DocumentSessionId
DocumentName
DocumentFullName
DocumentFolder
LastWriteTime
FileSize
WindowsUser
WordUser
MachineName
AddinVersion
HeartbeatIntervalSeconds
SourcePath
Notes
OpenMode
WriteReserved
ReadOnlyRecommended
```

Табы и переносы строк в полях заменяются пробелами. `EventId` уникален в пределах машины и сессии и используется admin-книгой для дедупликации.

### Локальный Spool И Сеть

Каждое событие сначала append-only записывается в:

```text
%APPDATA%\AuditWordAddin\activity-log\spool\<MachineName>__<WordSessionId>__YYYY-MM.tsv
```

`WordSessionId` содержит время старта и случайный discriminator. Поэтому два параллельных процесса Word не пишут в один локальный файл. Для каждого spool-файла sync-state отдельно хранит номер последней выгруженной строки. При доступности сети logger дописывает только новые строки в одноименный уникальный файл:

```text
Z:\Project\DAVK\Оборудование\Addins\word-assets\_system\activity-logs\computers\<MachineName>__<WordSessionId>__YYYY-MM.tsv
```

Каждая Word-сессия пишет только в свой месячный файл, поэтому между параллельными `WINWORD.EXE` нет shared append/state. Центральная admin-книга не открывается и не изменяется пользовательским runtime.

`logger-control.ini` повторяет staged rollout Excel-логгера:

```text
UploadDefault=No
UploadEnabledComputers=
UploadDisabledComputers=
UpdatedAt=
UpdatedBy=
```

Локальное логирование активно по умолчанию. Выгрузка в сеть включается централизованно для всех или отдельных машин. Недоступная или read-only сеть оставляет строки в локальном spool до следующей попытки.

### Хранение

В первой версии нет ограничения по сроку хранения.

- Runtime не удаляет, не перемещает, не архивирует и не обрезает старые TSV/state-файлы.
- Admin-книга импортирует все доступные месяцы.
- Локальные и сетевые файлы остаются до отдельного будущего решения о retention.

Месячное разбиение позволяет позже добавить календарную retention-политику без смены схемы событий, но такая логика не входит в текущую реализацию.

## Admin-Книга Word-Логов

Отдельная admin-only книга:

```text
!Журнал активности Word.xlsm
```

Она собирается из Word-линии исходников и не меняет `!Журнал активности Excel.xlsm`.

При открытии книга:

- читает `activity-logs\computers\*.tsv`;
- проверяет заголовки;
- игнорирует и диагностирует битые/частично записанные строки;
- убирает дубли по `EventId`;
- строит raw events, document sessions, files, users, computers и diagnostics;
- дает admin-контролы включения/выключения сетевой выгрузки.

Company-сборка admin-книги принимает пароль открытия параметром сборочного скрипта. Пароль не хранится в Git. До явного Word company-release достаточно реализовать исходники, builder и тесты; фактический защищенный viewer собирается только при передаче пароля или отдельной release-команде.

## Модель Доступа К Сети

Пути `word-addin-dev`, `word-addin-testers`, `word-addin-stable`, `word-assets` и desktop mirror придуманы рядом с Excel-линией и явно разрешены Павлом в этом обсуждении. Код надстройки не создает и не меняет NTFS/share ACL; права являются обязательной внешней предпосылкой company-публикации.

Рекомендуемая граница доверия:

- обычные пользователи имеют только чтение channel package, `word-assets\templates` и `word-assets\sync`;
- запись package, version, templates, registry и `logger-control.ini` разрешена только администраторам/издателям;
- `!Журнал активности Word.xlsm` и чтение всех raw TSV доступны только администраторам/аудиторам;
- на `activity-logs\computers` пользователям разрешены создание файла и запись данных без просмотра чужих файлов;
- `CREATOR OWNER` получает чтение/дозапись собственного session-файла, администраторы получают чтение всех файлов;
- локальный `%APPDATA%\AuditWordAddin` доступен текущему Windows-пользователю.

Пароль admin-книги не считается защитой raw TSV. Пока цифровая подпись отложена, read-only ACL сетевых `.dotm`/version/assets и admin-only ACL логов являются обязательной границей поставки. Company release checklist требует ручного подтверждения этих прав; при неподтвержденной ACL-модели пакет не объявляется готовым к массовому каналу.

## Обработка Ошибок

Общий принцип: Word и пользовательский документ важнее любой функции надстройки.

On-demand команды:

- валидируют обязательные пути до файловых вызовов;
- быстро завершаются с одним понятным сообщением;
- не делают повторных сетевых попыток внутри пользовательского клика;
- не оставляют частично созданные пользовательские документы.

Фоновые сервисы:

- не показывают `MsgBox`;
- пишут краткую причину в updater/logger log или `Debug.Print`;
- никогда не обращаются к сети из `DocumentBeforeSave`/`DocumentBeforeClose`;
- загружают локальный runtime до постановки любых сетевых задач;
- для каждого сервиса делают максимум один preflight/operation set за попытку без рекурсивного обхода сетевого корня;
- для mapped drive сначала проверяют существование диска и `IsReady`; отрицательный preflight завершает задачу до обращения к package/assets/log path;
- при сетевой ошибке ставят независимый четырехчасовой cooldown для runtime update, assets или logger flush;
- продолжают использовать локальный runtime и локальные assets.

Чистый VBA не может прервать уже начавшийся системный вызов к зависшему mapped drive, поэтому абсолютный hard timeout не обещается. Измеримый контракт такой: Loader `AutoExec` не делает сетевых вызовов и на исправном локальном диске целится завершиться менее чем за одну секунду; первая неудачная сетевая операция может ждать Windows timeout, но после нее тот же сервис не трогает сеть четыре часа. Закрытие и сохранение документов выполняют только локальные операции.

Это осознанное платформенное ограничение требует подтверждения вместе с данной спецификацией. Альтернатива с настоящим отменяемым timeout потребовала бы отдельного внешнего updater-процесса или shell/network helper, что прямо исключено company-политикой текущего scope.

Битый каталог шаблонов не мешает работе вида, хоткеев, диагностики и logger-а. Ошибка Activity Logger не мешает сохранению или закрытию документа.

## Ограничения Безопасности

Для dev и company runtime одинаково запрещены:

- интернет и HTTP-клиенты;
- Git как источник обновления;
- `Shell`, `WScript.Shell`, PowerShell, `cmd.exe`, `curl`;
- обход Office Trust Center;
- отправка через Outlook;
- `Kill`, `RmDir`, `DeleteFile`, `DeleteFolder`, `MoveFile`, `MoveFolder`, `Name old As new`;
- запись в `Normal.dotm`;
- изменение пользовательских документов ради служебных настроек или логов.

Разрешены только контролируемые `FileExists`, `FolderExists`, `CreateFolder`, `CopyFile`, чтение/запись текстовых config/log/TSV, `ADODB.Stream` и штатные Word APIs.

## Сборка И Release

Dev build:

1. запускает source-level feature tests;
2. создает временные CP1251-копии VBA;
3. через отдельный Word COM instance собирает `.dotm`;
4. встраивает `customUI14` и четыре PNG-иконки;
5. проверяет package structure и обязательные VBA-компоненты;
6. обновляет `word_addin/release/AuditWordAddin.dotm` и version file;
7. записывает список существующих `WINWORD.EXE`, использует отдельный automation instance и завершает только созданный тестом процесс;
8. устанавливает свежий dev runtime на машину Павла для ручной проверки, не закрывая чужой Word-процесс без отдельного разрешения.

Company build запускается только по явной команде:

1. собирает staging из проверенной dev-кодовой базы;
2. защищает VBA runtime, Loader и Setup;
3. проверяет готовые `.dotm` на запрещенные package-маркеры;
4. проверяет VBA-защиту;
5. только после этого обновляет `company/На рабочий комп` и `Результаты проверок`;
6. архивирует предыдущий company-release;
7. обновляет desktop mirror и zip.

## Стратегия Тестирования

### Source-Level Feature Tests

- `ribbon`: ровно четыре контрола с `onAction`, без дочерних menu-команд;
- `rd-template-opener`: TSV-схема, динамические значения, disabled/missing/duplicate/traversal cases;
- `view-settings`: дефолты, parsing, диапазоны масштаба, события открытия/закрытия;
- `hotkeys`: только allowlist-команд, runtime customization context, отсутствие `Normal.dotm`;
- `diagnostics`: обязательные поля и безопасная обработка пустых путей;
- `scheduler`: единственная точка `Application.OnTime`, queue/coalescing/self-heal и отсутствие независимых timer-ов;
- `activity`: Word events, pending-save state machine, запрет file/network calls в `DocumentBeforeSave`, per-session monthly spool, sync-state, отсутствие retention/delete;
- `setup-loader`: channel-relative paths, StartupPath, Loader-last commit point, dual-slot unload-before-fallback, version protocol, независимые cooldown и asset allowlist;
- `source-policy`: запрещенные shell/network/delete/move-маркеры;
- `build`: обязательные исходники, Ribbon XML, icons и release layout.

### Word COM / Runtime Smoke

Тесты Word выполняются последовательно в отдельном automation instance:

- загрузить собранный runtime как global template;
- проверить вкладку и четыре команды;
- открыть fixture-каталог с `.dotx` и `.dotm` и создать новые несохраненные документы;
- проверить понятное поведение при пустом/битом каталоге;
- применить каждый профиль вида и убедиться, что `Document.Saved` не изменился;
- проверить события `DocumentOpen`, `NewDocument`, `DocumentBeforeClose`;
- назначить/снять хоткеи и убедиться, что `Normal.dotm` не изменен и нет save prompt;
- проверить диагностику в dev/offline режимах;
- получить `OPEN`, `SAVE`, `SEEN`, `MISSING` в локальном TSV;
- проверить обычный Save, успешный Save As, отмененный Save As, повторные pending-save и закрытие до callback;
- проверить конфликт/пропуск `OnTime` и self-heal при следующем Word event;
- запустить две Word-сессии и убедиться, что они используют разные spool/network TSV;
- проверить network flush в локальную имитацию и offline fallback;
- закрыть automation Word без зависшего процесса или модального окна.

Тесты фиксируют исходные PID, не закрывают пользовательские `WINWORD.EXE` и не продолжают destructive cleanup, если не могут доказать, какой процесс создали сами. Если локальная ручная установка требует перезапуска открытого Word, пользователь получает короткую просьбу закрыть его.

### Setup/Loader Integration

Во временных каталогах проверяются:

- установка из каждой channel-shaped папки;
- запись фактических package/version paths;
- установка Loader-а в полученный Word Startup;
- первоначальное локальное копирование runtime в слот `A`;
- staging более новой версии в неактивный слот и активация только на следующем старте;
- health-check failure pending-слота и загрузка предыдущего active-слота;
- после любого fallback установлен ровно один runtime и виден ровно один Ribbon;
- пустая, битая и пониженная version string;
- нулевой/оборванный/не-ZIP runtime package;
- загрузка старой локальной версии при недоступной сети;
- медленный/отключенный `Z:` и три независимых четырехчасовых cooldown;
- повторный Setup при уже загруженном стабильном Loader-е;
- ошибка runtime/version/config до commit point не оставляет новый Loader в Startup;
- read-only Startup и macro-blocked Setup с понятной ошибкой без частичной установки;
- asset sync по allowlist/`DateLastModified` без удаления и с каталогом последним;
- исключение центральных Word-логов и admin-книги из пользовательского кэша.

### Company Release Checks

- package scan по итоговым `.dotm`;
- отсутствие dev-хвостов и запрещенных маркеров;
- подтверждение VBA protection;
- корректная структура `На рабочий комп\Не трогать`;
- ручное подтверждение ACL для channel packages, assets, control и raw logs;
- человекочитаемые отчеты с временем проверки;
- desktop mirror/zip только после успешных проверок.

## Критерии Приемки Первой Версии

- Excel-надстройка и ее файлы не изменены.
- Собирается отдельный незапароленный dev `AuditWordAddin.dotm`.
- Word показывает вкладку `Аудиторам` ровно с четырьмя согласованными Ribbon-командами.
- Пустой каталог шаблонов поддерживается без ошибки; fixture-шаблоны открываются как новые несохраненные документы.
- Настройки вида работают при открытии/закрытии и по команде, не помечая документ измененным.
- Хоткеи назначаются только allowlist-командам и не меняют `Normal.dotm`.
- Диагностика показывает runtime, Startup, channel, paths, catalog и logger status.
- Activity Logger пишет отдельные Word TSV и не записывает содержимое документов.
- В Activity Logger отсутствует логика шестимесячной retention/очистки.
- Setup/Loader поддерживают `dev`, `testers`, `stable`, dual-slot fallback и независимые cooldown.
- Runtime остается работоспособным без сети.
- В runtime есть только одна точка постановки Word `Application.OnTime`; конфликт имеет документированную best-effort семантику.
- Два параллельных Word-процесса не пишут в один TSV/state-файл.
- Source tests и релевантные COM smoke проходят последовательно.
- Оба Git-репозитория остаются чистыми после зафиксированных коммитов.

## Отложенные Решения

- Реальные строки `word-rd-template-map.tsv` и фактические Word-шаблоны появятся после их передачи и согласования.
- Календарная retention-политика Activity Logger откладывается; первая версия хранит все месяцы.
- Пароль открытия company admin-книги передается только при фактическом company-release.
- Цифровая подпись Office-пакетов не вводится без отдельного решения.
