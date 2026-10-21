#!/bin/bash
# Создаёт .venv с Python 3.10+ и ставит зависимости. Запуск: bash setup.sh
set -e
cd "$(dirname "$0")"
DIR="$(pwd)"
PY=""
for c in python3.13 python3.12 python3.11 python3.10 python3 \
         /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if command -v "$c" >/dev/null 2>&1 && \
     "$c" -c 'import sys; sys.exit(sys.version_info < (3, 10))' 2>/dev/null; then
    PY="$(command -v "$c")"; break
  fi
done
rm -rf .venv
if [ -n "$PY" ]; then
  echo "Python: $PY"
  "$PY" -m venv .venv
  .venv/bin/pip install -q --upgrade pip
  .venv/bin/pip install -q -r requirements.txt
elif command -v uv >/dev/null 2>&1; then
  echo "Python 3.10+ не найден, ставлю через uv"
  uv venv --python 3.12 .venv
  uv pip install -q --python .venv/bin/python -r requirements.txt
else
  echo "Нужен Python 3.10+: brew install python@3.12 — и запустить снова" >&2
  exit 1
fi
.venv/bin/python -c "import server; print('Сервер импортируется: OK')"
cat <<JSON

Добавьте в ~/Library/Application Support/Claude/claude_desktop_config.json:

{
  "mcpServers": {
    "mailru": {
      "command": "$DIR/.venv/bin/python",
      "args": ["$DIR/server.py"],
      "env": {
        "MAILRU_USER": "you@mail.ru",
        "MAILRU_APP_PASSWORD": "пароль-приложения",
        "MAILRU_ALLOW_SEND": "0"
      }
    }
  }
}
JSON
