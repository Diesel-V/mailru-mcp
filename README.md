# mailru-mcp

Локальный MCP-сервер для почты (IMAP/SMTP) и календаря (CalDAV) Mail.ru, транспорт stdio.

## Инструменты

| Инструмент | Что делает |
|---|---|
| `list_folders` | список папок |
| `search_messages` | поиск: текст, отправитель, тема, даты, только непрочитанные |
| `get_message` | письмо целиком + список вложений (не помечает прочитанным) |
| `set_flags` | прочитано / флажок |
| `move_message` | перенос в папку |
| `trash_message` | в Корзину (без безвозвратного удаления) |
| `create_draft` | черновик в «Черновики» |
| `send_message` | отправка — только при `MAILRU_ALLOW_SEND=1` |
| `calendar_list` | календари, доступные по CalDAV |
| `calendar_events` | встречи за период (повторы раскрыты, участники, ссылка на звонок) |
| `calendar_free_slots` | встречи и свободные слоты по дням в рабочие часы; `extra_busy` — занятость из других календарей |
| `calendar_event` | одна встреча целиком: все участники, описание, повтор |
| `calendar_create_event` | создать встречу (участники, повтор, весь день) |
| `calendar_update_event` | изменить встречу или одно вхождение серии |
| `calendar_delete_event` | удалить встречу или одно вхождение серии |
| `calendar_respond` | принять / отклонить / «возможно» |

## Установка

1. Mail.ru → Настройки → Безопасность → **Пароли для внешних приложений** → создать
   пароль с доступом IMAP и SMTP. Обычный пароль от ящика по IMAP не работает.
2. Зависимости (Python 3.10+):

   ```
   python -m venv .venv
   .venv/bin/pip install -r requirements.txt        # Windows: .venv\Scripts\pip
   ```

3. В конфиг Claude Desktop (`claude_desktop_config.json`, Настройки → Разработчик →
   Изменить конфигурацию) добавить:

   ```json
   {
     "mcpServers": {
       "mailru": {
         "command": "/ABS/PATH/mailru-mcp/.venv/bin/python",
         "args": ["/ABS/PATH/mailru-mcp/server.py"],
         "env": {
           "MAILRU_USER": "you@mail.ru",
           "MAILRU_APP_PASSWORD": "пароль-приложения",
           "MAILRU_ALLOW_SEND": "0"
         }
       }
     }
   }
   ```

4. Перезапустить приложение.

## Переменные окружения

- `MAILRU_USER`, `MAILRU_APP_PASSWORD` — обязательные
- `MAILRU_ALLOW_SEND=1` — включить `send_message`
- `MAILRU_FROM_NAME` — отображаемое имя отправителя
- `MAILRU_CALDAV_URL` — по умолчанию `https://calendar.mail.ru`
- `MAILRU_CALDAV_PASSWORD` — отдельный пароль приложения для календаря, если почтовый не подходит
- `MAILRU_TZ` — часовой пояс календаря, по умолчанию `Europe/Moscow`
- `MAILRU_CALENDAR_DEFAULT` — календарь для новых встреч, по умолчанию «Основной»
- `MAILRU_CALENDAR_READONLY=1` — запретить создание, правку, удаление и ответы
- `MAILRU_IMAP_HOST/PORT`, `MAILRU_SMTP_HOST/PORT` — по умолчанию `imap.mail.ru:993`, `smtp.mail.ru:465`

## Ограничения

- Пароль приложения лежит в конфиге открытым текстом — ограничьте права на файл.
- Вложения не скачиваются, только перечисляются.
- В одном поиске на сервер уходит одно условие с кириллицей; остальные фильтруются локально
  по последним письмам.

## Лицензия

MIT
