# MCP Collabora Holix

MCP-сервер, который читает и правит офисные файлы в одной папке: Word, Excel, PowerPoint и LibreOffice.

Агент общается с сервером по stdio и вызывает инструменты. Инструменты меняют настоящие `.docx`, `.xlsx`, `.pptx`, `.odt`, `.ods` и `.odp` на диске. Файл никуда не отправляется на конвертацию.

[English](README.md) · [Справочник](docs/ru/reference.md) · [Tool reference](docs/en/reference.md)

## Что он делает

Укажите серверу каталог. Попросите агента найти документ, прочитать его, заменить формулировку, оформить фразу, вставить картинку, записать формулу или добавить слайд. Результат — это файл на диске.

Holix Studio подключает этот сервер сам, когда включено редактирование офиса, и открытая вкладка Collabora перечитывает файл после удачной правки. Claude, OpenClaw, Hermes, Codex и Holix используют тот же сервер. Фрагменты конфигурации ниже.

## Что можно попросить

| | |
|---|---|
| Документы | DOCX, XLSX, PPTX, ODT, ODS, ODP в рабочей папке |
| Текст | Найти файлы, прочитать текст, заменить фразу, дописать абзац в DOCX или ODT |
| Оформление | Шрифт, размер, полужирный, курсив, подчёркивание, цвет и заливка абзаца. Именованные стили вроде Heading 1 для DOCX и ODT |
| Медиа | PNG, JPEG, GIF, WebP, BMP и TIFF во всех шести типах. MP4, WebM и MOV воспроизводятся на слайдах PPTX и ODP и в кадре ODT. DOCX хранит видео и ссылку |
| Презентации | Список, новый слайд, заголовок, заметки и удаление слайдов PPTX и ODP |
| Листы | Новые листы, ячейки, формулы, гиперссылки и именованные диапазоны Excel. Формулы ODS принимают запись Excel, например `=Other!A1` |
| Диаграммы | Столбчатая, линейная и круговая на XLSX и PPTX |
| Таблицы | Чтение, добавление и правка ячейки в DOCX, PPTX, ODT и ODP. Сетка на XLSX или ODS |

Пути задаются относительно рабочей папки, например `reports/plan.docx` и `media/cover.png`. Путь, который выходит из этой папки, сервер отклоняет.

## Честные границы

- Заливка абзаца применяется только когда весь текст абзаца совпадает с указанной фразой.
- Именованные стили есть у DOCX и ODT. Диаграммы — у XLSX и PPTX. Именованные диапазоны — у XLSX.
- DOCX не воспроизводит видео. Файл сохраняется внутри документа, рядом ставится ссылка.
- Нет анимации слайдов, переходов и мастер-слайдов.
- `office_read_tool` возвращает текст. Числа и формулы читает `office_sheet_tool` с действием `read`.
- Дописать абзац можно в DOCX и ODT.
- В Holix Studio перезагрузка заменяет несохранённые нажатия в этой вкладке Collabora сохранённым файлом.

Полные контракты — в [справочнике](docs/ru/reference.md).

## Установка

Нужен Python 3.12 или новее.

```bash
pip install "git+https://github.com/javded-itres/mcp-collabora-holix.git"
python -m holix_office
```

`python -m holix_office` ждёт клиента MCP на stdin. Так и должно быть. Рабочую папку задайте до запуска клиента:

```bash
export HOLIX_OFFICE_WORKSPACE="$HOME/Documents"
```

Путь должен быть абсолютным. После добавления или смены сервера откройте новый чат: уже запущенный процесс MCP сам инструменты не перечитывает.

## Подключение

### Holix Studio

Отдельный фрагмент не нужен. Когда редактирование офиса включено, Studio записывает сервер `holix_office` в профиль Holix, назначает его основному агенту и ставит навык `office_documents`. Рабочая папка — это рабочая папка профиля, а не случайный каталог чата.

После включения офиса откройте новый чат.

### Holix

В `config.yaml` профиля:

```yaml
mcp_servers:
  holix_office:
    transport: stdio
    command: mcp-collabora-holix
    args: []
    env:
      HOLIX_OFFICE_WORKSPACE: /absolute/path/to/documents
```

Запустите новую сессию. В Studio эту запись лучше не править руками: Studio сама создаёт `holix_office` и обновляет его.

### Claude

Claude Desktop, macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`. Windows: `%APPDATA%\Claude\claude_desktop_config.json`.

```json
{
  "mcpServers": {
    "mcp-collabora-holix": {
      "command": "mcp-collabora-holix",
      "args": [],
      "env": {
        "HOLIX_OFFICE_WORKSPACE": "/absolute/path/to/documents"
      }
    }
  }
}
```

Полностью закройте Claude Desktop и откройте снова.

Claude Code из каталога проекта:

```bash
claude mcp add --transport stdio --env HOLIX_OFFICE_WORKSPACE=/absolute/path/to/documents mcp-collabora-holix -- mcp-collabora-holix
```

Если `mcp-collabora-holix` нет в `PATH`, подставьте путь из `command -v mcp-collabora-holix` или команду `python` с аргументами `["-m", "holix_office"]`.

### OpenClaw

```bash
openclaw mcp add mcp-collabora-holix \
  --command mcp-collabora-holix \
  --env HOLIX_OFFICE_WORKSPACE=/absolute/path/to/documents
openclaw mcp doctor mcp-collabora-holix --probe
```

Та же запись в конфигурации OpenClaw:

```text
mcp:
  servers:
    mcp-collabora-holix:
      command: mcp-collabora-holix
      transport: stdio
      enabled: true
      env:
        HOLIX_OFFICE_WORKSPACE: /absolute/path/to/documents
```

### Hermes

`~/.hermes/config.yaml`:

```yaml
mcp_servers:
  mcp-collabora-holix:
    command: mcp-collabora-holix
    args: []
    env:
      HOLIX_OFFICE_WORKSPACE: /absolute/path/to/documents
```

Дальше `hermes chat`. Инструменты видны как `mcp__mcp-collabora-holix__office_list_tool` и остальные `office_*`.

### Codex

`~/.codex/config.toml` или `.codex/config.toml` доверенного проекта:

```toml
[mcp_servers.mcp-collabora-holix]
command = "mcp-collabora-holix"
args = []

[mcp_servers.mcp-collabora-holix.env]
HOLIX_OFFICE_WORKSPACE = "/absolute/path/to/documents"
```

Или:

```bash
codex mcp add mcp-collabora-holix --env HOLIX_OFFICE_WORKSPACE=/absolute/path/to/documents -- mcp-collabora-holix
```

В TUI Codex команда `/mcp` показывает сервер.

## Рабочая папка

| Переменная | Смысл |
|---|---|
| `HOLIX_OFFICE_WORKSPACE` | Абсолютный путь папки, которую инструменты могут читать и менять |
| `HOLIX_STUDIO_WORKSPACE_ROOT` | Ставит Holix Studio, если публичная переменная не задана |
| `HOLIX_PROFILE` | Имя профиля Studio, нужно только чтобы обновить открытый редактор |

Если ни одна переменная пути не задана, рабочей папкой становится текущий каталог процесса.

## Лицензия

MIT. © 2026 Pavel Lukyanov и участники ITRES.
