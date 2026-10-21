#!/bin/bash
# Добавляет пароль приложения для календаря (CalDAV) в конфиг Claude Desktop.
# Почтовый пароль не трогает. Запуск: bash configure_calendar.sh
set -e
cd "$(dirname "$0")"
DIR="$(pwd)"
CFG="$HOME/Library/Application Support/Claude/claude_desktop_config.json"
[ -x "$DIR/.venv/bin/python" ] || { echo "Нет .venv — сначала: bash $DIR/setup.sh" >&2; exit 1; }
[ -s "$CFG" ] || { echo "Нет конфига $CFG — сначала: bash $DIR/configure.sh" >&2; exit 1; }
read -r -s -p "Пароль приложения для календаря (ввод скрыт): " MAILRU_CALDAV_PASSWORD; echo
[ -n "$MAILRU_CALDAV_PASSWORD" ] || { echo "Пустой пароль" >&2; exit 1; }
export MAILRU_CALDAV_PASSWORD CFG
"$DIR/.venv/bin/python" - <<'PY'
import json, os, shutil, sys, time
cfg = os.environ["CFG"]
with open(cfg, encoding="utf-8") as f:
    data = json.load(f)
srv = data.get("mcpServers", {}).get("mailru")
if not srv:
    sys.exit("В конфиге нет сервера mailru — сначала: bash configure.sh")
bak = f"{cfg}.bak-{time.strftime('%Y%m%d-%H%M%S')}"
shutil.copy2(cfg, bak)
print("Резервная копия:", bak)
srv.setdefault("env", {})["MAILRU_CALDAV_PASSWORD"] = os.environ["MAILRU_CALDAV_PASSWORD"]
with open(cfg, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
os.chmod(cfg, 0o600)
print("Готово:", cfg)
PY
echo "Проверяю календарь..."
MAILRU_USER="$("$DIR/.venv/bin/python" -c "
import json, os
print(json.load(open(os.environ['CFG'], encoding='utf-8'))['mcpServers']['mailru']['env']['MAILRU_USER'])")"
export MAILRU_USER
"$DIR/.venv/bin/python" -c "
import mailru_calendar as cal
cs = cal.calendars(refresh=True)
print('CalDAV: OK, календарей:', len(cs))
for c in cs: print(' -', c['name'])
" || echo "Календарь не ответил — текст ошибки выше." >&2
echo "Перезапустите Claude Desktop."
