#!/bin/bash
# Прописывает сервер mailru в конфиг Claude Desktop. Запуск: bash configure.sh
set -e
cd "$(dirname "$0")"
DIR="$(pwd)"
CFG="$HOME/Library/Application Support/Claude/claude_desktop_config.json"
[ -x "$DIR/.venv/bin/python" ] || { echo "Нет .venv — сначала: bash $DIR/setup.sh" >&2; exit 1; }
read -r -p "Адрес ящика Mail.ru: " MAILRU_USER
read -r -s -p "Пароль приложения (ввод скрыт): " MAILRU_APP_PASSWORD; echo
read -r -p "Разрешить отправку писем? [y/N]: " SEND
case "$SEND" in y|Y|д|Д) MAILRU_ALLOW_SEND=1 ;; *) MAILRU_ALLOW_SEND=0 ;; esac
[ -n "$MAILRU_USER" ] && [ -n "$MAILRU_APP_PASSWORD" ] || { echo "Пустой логин или пароль" >&2; exit 1; }
export MAILRU_USER MAILRU_APP_PASSWORD MAILRU_ALLOW_SEND DIR CFG
"$DIR/.venv/bin/python" - <<'PY'
import json, os, shutil, time
cfg, d = os.environ["CFG"], os.environ["DIR"]
data = {}
if os.path.exists(cfg) and os.path.getsize(cfg):
    with open(cfg, encoding="utf-8") as f:
        data = json.load(f)
    bak = f"{cfg}.bak-{time.strftime('%Y%m%d-%H%M%S')}"
    shutil.copy2(cfg, bak)
    print("Резервная копия:", bak)
os.makedirs(os.path.dirname(cfg), exist_ok=True)
data.setdefault("mcpServers", {})["mailru"] = {
    "command": f"{d}/.venv/bin/python",
    "args": [f"{d}/server.py"],
    "env": {k: os.environ[k] for k in
            ("MAILRU_USER", "MAILRU_APP_PASSWORD", "MAILRU_ALLOW_SEND")},
}
with open(cfg, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
os.chmod(cfg, 0o600)
print("Готово:", cfg)
PY
echo "Проверяю вход в ящик..."
"$DIR/.venv/bin/python" -c "
import server
with server._imap() as c:
    print('IMAP: OK, папок:', len(server._folders(c)))
" || echo "Вход не удался — проверьте адрес и пароль приложения (нужен доступ IMAP)." >&2
echo "Перезапустите Claude Desktop."
