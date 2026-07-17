# Дизайн Word-надстройки `Аудиторам`

Дата: 2026-07-17

Статус: дизайн согласован в чате; ожидает документального ревью и плана реализации.

## Контекст

Нужна первая отдельная надстройка `Аудиторам` для настольного Microsoft Word на Windows. Надстройка создается на VBA в формате `.dotm`, использует собственный Ribbon и не встраивается в существующую Excel-надстройку.

Word-линия должна повторять проверенную схему доставки Excel-линии: локальный runtime, легкий loader в Word Startup, отдельный setup и три сетевых канала `dev`, `testers`, `stable`. При этом код, пакеты, настройки и ресурсы Word не должны менять или переиспользовать runtime-файлы Excel-надстройки.

Первая версия содержит четыре видимые команды и один скрытый фоновый сервис Activity Logger. Количество и виды шаблонов рабочих документов пока неизвестны, поэтому каталог шаблонов должен быть управляемым данными, а не зашитым в `.dotm`.

## Цели

- Создать самостоятельную Word-надстройку `AuditWordAddin.dotm`.
- Добавить на вкладку `Аудиторам` только четыре согласованные команды.
- Создать установку и автообновление по схеме `Setup -> Word Startup Loader -> local runtime`.
- Поддержать сетевые каналы `word-addin-dev`, `word-addin-testers`, `word-addin-stable` одним channel-agnostic пакетом.
- Поддержать произвольный каталог Word-шаблонов через TSV-реестр.
- Добавить скрытый Activity Logger с теми же типами действий, которые фактически отслеживает Excel-логгер.
- Сохранить отказоустойчивость: недоступная сеть, битый путь или необязательный ресурс не должны подвешивать Word.
- Сразу предусмотреть dev-сборку без VBA-пароля и отдельный company-пайплайн с VBA-защитой.

## Не Цели Первой Версии

- Не менять `addin/`, `addin_updater/`, `AuditAddin.xlam`, Excel Ribbon, Excel Activity Logger или Excel admin-книгу.
- Не добавлять очистку текста, оформление таблиц, экспорт в PDF, проверку документа, инструкцию, конструктор РД или другие пользовательские команды.
- Не добавлять видимую кнопку Activity Logger на Ribbon.
- Не записывать текст документов, нажатия клавиш, изменения текста, выделения, перемещения курсора, печать или содержимое буфера обмена.
- Не вводить ограничение хранения на шесть месяцев, автоматическую ротацию, архивирование или очистку Word-логов.
- Не фиксировать заранее перечень видов аудита, этапов, кодов РД или количество шаблонов.
- Не использовать интернет, Git, shell, PowerShell, `cmd.exe`, `WScript.Shell`, HTTP-клиенты или внешние updater-процессы в runtime.
- Не изменять `Normal.dotm`, глобальные настройки безопасности Word или документы пользователя ради хранения настроек надстройки.
- Не публиковать company-пакет в сетевые каналы без отдельной команды на company-release.

## Принятое Архитектурное Решение

Используется отдельная Word-линия со своими исходниками, пакетами, локальным корнем и сетевыми ресурсами.

```text
Word Setup
    -> копирует Word Loader в Application.StartupPath
    -> копирует Word runtime в %APPDATA%\AuditWordAddin\installed
    -> записывает локальный конфиг

Word startup
    -> автоматически загружает Loader из Application.StartupPath
    -> Loader при возможности обновляет локальный runtime
    -> Loader загружает локальный runtime как global template/add-in
    -> Loader best-effort синхронизирует Word assets

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
        <MachineName>__YYYY-MM.tsv
      logger-control.ini
      !Журнал активности Word.xlsm
```

Папка `templates\` может быть пустой в первой сборке. Ни количество шаблонов, ни их иерархия не являются частью контракта runtime.

Loader синхронизирует в локальный кэш только пользовательские assets и безопасный control-файл логгера. Он не копирует на компьютеры пользователей:

- `_system\activity-logs\computers\`;
- `!Журнал активности Word.xlsm`;
- чужие сетевые логи.

## Локальные Пути

Авторитетный Word Startup определяется во время выполнения через `Application.StartupPath`, а не через захардкоженную строку `%APPDATA%\Microsoft\Word\STARTUP`.

```text
<Application.StartupPath>\AuditWordAddinLoaderLocal.dotm

%APPDATA%\AuditWordAddin\installed\AuditWordAddinLocal.dotm
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
5. копирует Loader как `AuditWordAddinLoaderLocal.dotm`;
6. копирует runtime как `AuditWordAddinLocal.dotm`;
7. копирует/записывает локальную версию;
8. записывает config с путями выбранного канала и Word assets;
9. сообщает об успешной установке и необходимости перезапустить Word.

Setup не изменяет Trust Center, `Normal.dotm`, реестр безопасности Office и параметры корпоративной политики. Если политика Word блокирует VBA из сетевой папки, это внешнее условие установки; код не пытается его обходить.

Повторный запуск Setup безопасно обновляет Loader, runtime и config. Обычно он нужен только при первой установке или изменении самого Loader/Setup; обычные runtime-обновления выполняет Loader.

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
LocalInstallPath=%APPDATA%\AuditWordAddin\installed\AuditWordAddinLocal.dotm
AssetsRoot=%APPDATA%\AuditWordAddin\assets
RdTemplateRoot=%APPDATA%\AuditWordAddin\assets\templates
NetworkRetryAfter=
LastNetworkCheckAt=
LastSuccessfulUpdateAt=
LastAssetsSyncAt=
```

В реальном файле переменные окружения раскрыты до абсолютных путей. Пустой или некорректный обязательный путь считается недоступным и не отправляется в FSO/Word API.

## Обновление При Старте Word

Word автоматически загружает `AuditWordAddinLoaderLocal.dotm` из Startup. Loader выполняет легкий `AutoExec`:

1. читает локальный config;
2. если действует `NetworkRetryAfter`, пропускает сеть;
3. если проверка разрешена, один раз проверяет package/version выбранного канала;
4. при более новой версии копирует runtime в локальный controlled path до загрузки runtime;
5. при любой сетевой ошибке записывает лог и ставит четырехчасовой cooldown;
6. независимо от результата сети загружает последнюю пригодную локальную копию через Word AddIns;
7. после загрузки runtime запускает отложенную best-effort синхронизацию Word assets по `DateLastModified`.

Отказ сети не является отказом Word-надстройки. Setup гарантирует наличие первоначального локального runtime. Если обновление не удалось, используется предыдущая локальная копия.

Loader не удаляет и не перемещает файлы. Для обновления разрешено только контролируемое копирование. Assets обновляются без удаления локальных файлов: новые и более свежие сетевые файлы копируются, отсутствующие в сети локальные файлы остаются.

## Ribbon

Используется `customUI14`. На ленте создается отдельная вкладка `Аудиторам`, размещенная после вкладки рецензирования Word.

Видимы только четыре согласованные команды:

```text
Аудиторам
  Рабочие документы
    Открыть шаблон РД

  Вид документа
    Настроить вид
      При открытии...
      При закрытии...
      Применить открытие сейчас
      Применить закрытие сейчас

  Клавиши
    Горячие клавиши

  Сервис
    Диагностика
```

Для четырех основных команд используются отдельные осмысленные PNG-иконки. Разрешено скопировать подходящие исходные изображения из `addin/assets` в `word_addin/assets`, не меняя Excel Ribbon и его assets.

Activity Logger не получает кнопку, menu item или пользовательскую панель.

## Команда `Открыть шаблон РД`

Команда читает только локальный кэш:

```text
%APPDATA%\AuditWordAddin\assets\sync\word-rd-template-map.tsv
%APPDATA%\AuditWordAddin\assets\templates\
```

Сетевых обращений в on-demand сценарии нет.

### Формат Реестра

`word-rd-template-map.tsv` хранится в UTF-8 и имеет колонки:

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
- `Enabled=Yes` включает строку, `No` скрывает;
- `RelativePath` задается относительно `templates\`;
- абсолютные пути, `..` и выход за корень шаблонов запрещены;
- разрешены `.dotx` и `.dotm`;
- отсутствующий файл не ломает весь каталог, но строка попадает в диагностику;
- битый заголовок или невозможность прочитать реестр дают быстрое понятное сообщение.

Форма выбора строит значения динамически в последовательности:

1. вид аудита;
2. этапность/вариант;
3. код и отображаемое имя РД.

Никакие варианты не зашиваются в код. Если активных шаблонов нет, команда показывает понятное сообщение и завершается.

После выбора создается новый несохраненный документ через `Documents.Add` на основе локального `.dotx`/`.dotm`. Исходный шаблон не открывается для редактирования и не изменяется. Это штатная модель Word [`Documents.Add`](https://learn.microsoft.com/nl-nl/office/vba/api/word.documents.add).

## Команда `Настроить вид`

Есть одна панель с двумя наборами настроек: `При открытии` и `При закрытии`.

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

Единая панель разрешает назначать сочетания только поддерживаемым Word-командам:

- `Открыть шаблон РД`;
- `Настройки при открытии`;
- `Настройки при закрытии`;
- `Применить открытие сейчас`;
- `Применить закрытие сейчас`;
- `Горячие клавиши`;
- `Диагностика`.

Используется Word [`KeyBindings`](https://learn.microsoft.com/en-us/office/vba/api/word.keybindings), а не Excel `Application.OnKey`.

Перед добавлением binding код устанавливает [`CustomizationContext`](https://learn.microsoft.com/en-us/office/vba/api/word.global.customizationcontext) в локально загруженный runtime template. `Normal.dotm` не используется как контекст и не сохраняется.

Сочетания хранятся в `%APPDATA%\AuditWordAddin\user-settings.ini` и повторно применяются в памяти при каждой сессии Word. Runtime `.dotm` не сохраняется после настройки клавиш.

Панель:

- показывает текущее сочетание;
- обнаруживает конфликт с существующей Word-командой;
- требует явного подтверждения перед перекрытием конфликта;
- умеет снять только binding, созданный Word-надстройкой;
- не удаляет и не переписывает сочетания других templates/надстроек.

Закрытие Word после назначения клавиш не должно показывать запрос на сохранение `Normal.dotm` или runtime `.dotm`.

## Команда `Диагностика`

Команда показывает read-only панель с базовой информацией:

- версия и revision Word-надстройки;
- полный путь загруженного `AuditWordAddinLocal.dotm`;
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
- последняя содержательная строка loader log.

Диагностика ничего не исправляет и не обновляет. В ней нет кнопок updater-а, Activity Logger, конструктора или других функций.

## Activity Logger

Activity Logger является скрытым best-effort сервисом. Он повторяет фактический уровень событий Excel-логгера, но использует Word-объекты и отдельное хранилище.

Типы событий:

- `SESSION_START`: Word-логгер поднят в текущей сессии;
- `OPEN`: сохраненный документ замечен открытым;
- `SAVE`: подтверждено успешное сохранение документа;
- `SEEN`: heartbeat подтверждает, что документ остается открытым;
- `MISSING`: ранее отслеживаемый документ закрыт или больше не виден;
- `SESSION_END_ATTEMPT`: Word или runtime начали завершение.

Heartbeat выполняется раз в 300 секунд только при наличии отслеживаемых документов. Он перечисляет `Application.Documents`, пишет легкие локальные события и пытается выгрузить накопленные строки.

Несохраненные документы без пути не логируются. При первом успешном сохранении такого документа logger сначала backfill-ит `OPEN`, затем пишет `SAVE`, как Excel-реализация.

Word предоставляет `DocumentBeforeSave`, поэтому logger сохраняет легкий pre-save snapshot и выполняет ограниченную отложенную проверку результата. `SAVE` пишется только когда у документа появился валидный путь и подтверждено сохранение; отмененный Save/Save As не считается успешным событием. События Word Application подключаются через отдельный class module с `WithEvents`; это соответствует штатной модели [Application events](https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-application-object-word) и [`DocumentBeforeSave`](https://learn.microsoft.com/en-us/office/vba/api/word.application.documentbeforesave).

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
%APPDATA%\AuditWordAddin\activity-log\spool\<MachineName>__YYYY-MM.tsv
```

Sync-state хранит номер последней выгруженной строки. При доступности сети logger дописывает только новые строки в:

```text
Z:\Project\DAVK\Оборудование\Addins\word-assets\_system\activity-logs\computers\<MachineName>__YYYY-MM.tsv
```

Каждый компьютер пишет только в свой месячный файл. Центральная admin-книга не открывается и не изменяется пользовательским runtime.

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
- не блокируют открытие/закрытие Word;
- при сетевой ошибке переходят на четырехчасовой cooldown;
- продолжают использовать локальный runtime и локальные assets.

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
7. устанавливает свежий dev runtime на машину Павла для ручной проверки, не закрывая чужой Word-процесс без отдельного разрешения.

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

- `ribbon`: ровно четыре пользовательские команды и нужные callback-и;
- `rd-template-opener`: TSV-схема, динамические значения, disabled/missing/duplicate/traversal cases;
- `view-settings`: дефолты, parsing, диапазоны масштаба, события открытия/закрытия;
- `hotkeys`: только allowlist-команд, runtime customization context, отсутствие `Normal.dotm`;
- `diagnostics`: обязательные поля и безопасная обработка пустых путей;
- `activity`: Word events, типы событий, monthly spool, sync-state, отсутствие retention/delete;
- `setup-loader`: channel-relative paths, StartupPath, local fallback, cooldown, asset exclusions;
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
- проверить, что отмененный Save As не создает `SAVE`;
- проверить network flush в локальную имитацию и offline fallback;
- закрыть automation Word без зависшего процесса или модального окна.

Тесты не закрывают пользовательские `WINWORD.EXE` автоматически. Если локальная ручная установка требует перезапуска открытого Word, пользователь получает короткую просьбу закрыть его.

### Setup/Loader Integration

Во временных каталогах проверяются:

- установка из каждой channel-shaped папки;
- запись фактических package/version paths;
- установка Loader-а в полученный Word Startup;
- первоначальное локальное копирование runtime;
- обновление более новой версии до загрузки runtime;
- загрузка старой локальной версии при недоступной сети;
- четырехчасовой cooldown;
- asset sync по `DateLastModified` без удаления;
- исключение центральных Word-логов и admin-книги из пользовательского кэша.

### Company Release Checks

- package scan по итоговым `.dotm`;
- отсутствие dev-хвостов и запрещенных маркеров;
- подтверждение VBA protection;
- корректная структура `На рабочий комп\Не трогать`;
- человекочитаемые отчеты с временем проверки;
- desktop mirror/zip только после успешных проверок.

## Критерии Приемки Первой Версии

- Excel-надстройка и ее файлы не изменены.
- Собирается отдельный незапароленный dev `AuditWordAddin.dotm`.
- Word показывает вкладку `Аудиторам` ровно с четырьмя согласованными командами.
- Пустой каталог шаблонов поддерживается без ошибки; fixture-шаблоны открываются как новые несохраненные документы.
- Настройки вида работают при открытии/закрытии и по команде, не помечая документ измененным.
- Хоткеи назначаются только allowlist-командам и не меняют `Normal.dotm`.
- Диагностика показывает runtime, Startup, channel, paths, catalog и logger status.
- Activity Logger пишет отдельные Word TSV и не записывает содержимое документов.
- В Activity Logger отсутствует логика шестимесячной retention/очистки.
- Setup/Loader поддерживают `dev`, `testers`, `stable`, локальный fallback и cooldown.
- Runtime остается работоспособным без сети.
- Source tests и релевантные COM smoke проходят последовательно.
- Оба Git-репозитория остаются чистыми после зафиксированных коммитов.

## Отложенные Решения

- Реальные строки `word-rd-template-map.tsv` и фактические Word-шаблоны появятся после их передачи и согласования.
- Календарная retention-политика Activity Logger откладывается; первая версия хранит все месяцы.
- Пароль открытия company admin-книги передается только при фактическом company-release.
- Цифровая подпись Office-пакетов не вводится без отдельного решения.
