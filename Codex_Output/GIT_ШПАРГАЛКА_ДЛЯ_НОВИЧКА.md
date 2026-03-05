# Git шпаргалка для новичка (2 компьютера, очень просто)

## 0) Что это вообще
- `git` хранит историю твоих файлов.
- `GitHub` это сайт, куда ты отправляешь эту историю.
- Ты работаешь на 2 ПК:
- рабочий ПК (где тестируешь Excel)
- домашний ПК (где правишь код)

## 1) Три главных слова
- `commit` = сохранил изменения на этом компьютере.
- `push` = отправил эти сохранения на GitHub.
- `pull` = забрал с GitHub свежие изменения.

Запомни как мантру:
- сделал правки -> `commit` -> `push`
- сел на другой ПК -> `pull`

## 2) Один раз настроить (на каждом ПК)
Открой PowerShell и выполни:

```powershell
git --version
git config --global user.name "ТВОЕ_ИМЯ"
git config --global user.email "ТВОЙ_EMAIL"
```

Если `git --version` пишет ошибку, значит git не установлен.

## 3) Где работать
Всегда сначала перейти в папку проекта:

```powershell
cd "C:\Users\pavel\OneDrive\Рабочий стол\Codex_venv\Макросы для аудита"
```

Проверка:

```powershell
git status
```

Если видишь текст типа `On branch main` - всё ок.

## 4) Твой обычный цикл (рабочий ПК)
### Шаг A. Выгрузи VBA из Excel
- В Excel нажми кнопку `ExportAllModules_UI`.
- В проекте обновится папка `_vba_sync`.

### Шаг B. Отправь изменения на GitHub
В PowerShell:

```powershell
cd "C:\Users\pavel\OneDrive\Рабочий стол\Codex_venv\Макросы для аудита"
git status
git add _vba_sync
git commit -m "Экспорт VBA с рабочего ПК"
git push
```

Если после `git commit` пишет `nothing to commit`, значит изменений нет.

## 5) Твой обычный цикл (домашний ПК)
### Шаг A. Забери свежие изменения

```powershell
cd "C:\Users\pavel\OneDrive\Рабочий стол\Codex_venv\Макросы для аудита"
git pull
```

### Шаг B. Правки
- Правишь файлы в `_vba_sync` (сам или со мной).

### Шаг C. Отправь обратно

```powershell
git add _vba_sync
git commit -m "Правки VBA"
git push
```

## 6) Снова на рабочем ПК после правок
Сначала забери изменения:

```powershell
git pull
```

Потом в Excel:
- нажми `ImportAllModules_UI`
- проверь работу макросов

## 7) Самые частые ошибки и что делать
### Ошибка: `fatal: not a git repository`
Ты не в той папке. Сделай:

```powershell
cd "C:\Users\pavel\OneDrive\Рабочий стол\Codex_venv\Макросы для аудита"
```

### Ошибка при `push`: `rejected`
Кто-то уже отправил изменения раньше тебя.
Сделай:

```powershell
git pull
git push
```

### Ошибка при `pull` из-за локальных изменений
Сначала сохрани свои изменения:

```powershell
git add _vba_sync
git commit -m "Промежуточное сохранение"
git pull
```

## 8) Команды “подсмотреть”
```powershell
git status
git log --oneline -n 10
git diff
```

## 9) Самый короткий сценарий (если вообще ничего не помнишь)
На ПК, где есть правки:

```powershell
git add _vba_sync
git commit -m "Обновил VBA"
git push
```

На другом ПК:

```powershell
git pull
```
