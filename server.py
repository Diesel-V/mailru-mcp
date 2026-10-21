#!/usr/bin/env python3
"""MCP-сервер для почты (IMAP + SMTP) и календаря (CalDAV) Mail.ru, транспорт stdio.

Переменные окружения:
  MAILRU_USER           адрес ящика (обязательно)
  MAILRU_APP_PASSWORD   пароль для внешнего приложения (обязательно)
  MAILRU_ALLOW_SEND     "1" — зарегистрировать инструмент send_message
  MAILRU_FROM_NAME      отображаемое имя отправителя (необязательно)
  MAILRU_IMAP_HOST/PORT по умолчанию imap.mail.ru:993
  MAILRU_SMTP_HOST/PORT по умолчанию smtp.mail.ru:465
  MAILRU_CALDAV_URL     по умолчанию https://calendar.mail.ru
  MAILRU_CALDAV_PASSWORD отдельный пароль приложения для календаря (необязательно)
  MAILRU_TZ             часовой пояс календаря, по умолчанию Europe/Moscow
"""
from __future__ import annotations

import base64
import email
import email.policy
import imaplib
import os
import re
import smtplib
import ssl
import time
from contextlib import contextmanager
from datetime import date, timedelta
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from html.parser import HTMLParser
from typing import Optional

try:  # mcp 2.x
    from mcp.server.mcpserver import MCPServer as FastMCP
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP

import mailru_calendar as cal

USER = os.environ.get("MAILRU_USER", "")
PASSWORD = os.environ.get("MAILRU_APP_PASSWORD", "")
ALLOW_SEND = os.environ.get("MAILRU_ALLOW_SEND", "") == "1"
FROM_NAME = os.environ.get("MAILRU_FROM_NAME", "")
IMAP_HOST = os.environ.get("MAILRU_IMAP_HOST", "imap.mail.ru")
IMAP_PORT = int(os.environ.get("MAILRU_IMAP_PORT", "993"))
SMTP_HOST = os.environ.get("MAILRU_SMTP_HOST", "smtp.mail.ru")
SMTP_PORT = int(os.environ.get("MAILRU_SMTP_PORT", "465"))

POLICY = email.policy.default
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

mcp = FastMCP("mailru")


try:  # mcp 2.x
    from mcp.server.mcpserver.exceptions import ToolError
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp.exceptions import ToolError


def tool():
    """@mcp.tool(), но текст любой ошибки доходит до клиента.

    mcp 2.x скрывает сообщение у всего, кроме ToolError: клиент видит только
    «Error executing tool <name>». Здесь исключение переводится в ToolError,
    а traceback пишется в stderr (лог сервера).
    """
    import functools
    import sys
    import traceback

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            try:
                return fn(*args, **kwargs)
            except ToolError:
                raise
            except Exception as e:
                traceback.print_exc(file=sys.stderr)
                raise ToolError(f"{type(e).__name__}: {e}") from e
        return mcp.tool()(wrapper)
    return deco


# ---------- IMAP modified UTF-7 (имена папок) ----------

def utf7_decode(s: str) -> str:
    def repl(m: re.Match) -> str:
        chunk = m.group(1)
        if not chunk:
            return "&"
        b = chunk.replace(",", "/")
        b += "=" * (-len(b) % 4)
        return base64.b64decode(b).decode("utf-16-be")
    return re.sub(r"&([^-]*)-", repl, s)


def utf7_encode(s: str) -> str:
    out: list[str] = []
    buf: list[str] = []

    def flush() -> None:
        if buf:
            b = base64.b64encode("".join(buf).encode("utf-16-be")).decode()
            out.append("&" + b.rstrip("=").replace("/", ",") + "-")
            buf.clear()

    for ch in s:
        if 0x20 <= ord(ch) <= 0x7E:
            flush()
            out.append("&-" if ch == "&" else ch)
        else:
            buf.append(ch)
    flush()
    return "".join(out)


def _q(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _mbox(name: str) -> str:
    return _q(utf7_encode(name))


# ---------- соединение ----------

@contextmanager
def _imap():
    if not USER or not PASSWORD:
        raise RuntimeError("Не заданы MAILRU_USER / MAILRU_APP_PASSWORD")
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT,
                             ssl_context=ssl.create_default_context(), timeout=30)
    try:
        conn.login(USER, PASSWORD)
        yield conn
    finally:
        try:
            conn.logout()
        except Exception:
            pass


def _ok(typ: str, data, what: str) -> None:
    if typ != "OK":
        raise RuntimeError(f"{what}: {data!r}")


def _caps(conn) -> set[str]:
    typ, data = conn.capability()
    _ok(typ, data, "CAPABILITY")
    return set(data[0].decode("ascii", "replace").upper().split())


def _uid(uid: str) -> str:
    uid = str(uid).strip()
    if not uid.isdigit():
        raise ValueError("uid должен быть числом")
    return uid


_LIST_RE = re.compile(rb'\((?P<flags>[^)]*)\)\s+(?:"[^"]*"|NIL)\s+(?P<name>.*)')


def _folders(conn) -> list[dict]:
    typ, data = conn.list()
    _ok(typ, data, "LIST")
    res = []
    for item in data:
        line, lit = (item[0], item[1]) if isinstance(item, tuple) else (item, None)
        if not line:
            continue
        m = _LIST_RE.match(line)
        if not m:
            continue
        raw = lit if lit is not None else m.group("name").strip()
        name = raw.decode("ascii", "replace")
        if len(name) >= 2 and name[0] == '"' and name[-1] == '"':
            name = name[1:-1].replace('\\"', '"').replace("\\\\", "\\")
        res.append({"name": utf7_decode(name),
                    "flags": m.group("flags").decode("ascii", "replace").split()})
    return res


def _special(conn, flag: str, fallbacks: list[str]) -> str:
    folders = _folders(conn)
    for f in folders:
        if flag in f["flags"]:
            return f["name"]
    names = {f["name"] for f in folders}
    for fb in fallbacks:
        if fb in names:
            return fb
    raise RuntimeError(f"Не найдена папка {flag}")


def _select(conn, folder: str, readonly: bool = True) -> None:
    typ, data = conn.select(_mbox(folder), readonly=readonly)
    _ok(typ, data, f"SELECT {folder}")


# ---------- разбор писем ----------

def _h(msg, name: str) -> str:
    try:
        return str(msg.get(name, "") or "")
    except Exception:
        raw = msg.get_all(name, failobj=[""])
        return str(raw[0]) if raw else ""


def _parse_fetch(data) -> list[dict]:
    """Разбор ответа UID FETCH (FLAGS RFC822.SIZE BODY[HEADER.FIELDS ...])."""
    out = []
    for i, item in enumerate(data):
        if not isinstance(item, tuple):
            continue
        meta = item[0]
        if i + 1 < len(data) and isinstance(data[i + 1], bytes):
            meta += b" " + data[i + 1]
        m_uid = re.search(rb"UID (\d+)", meta)
        if not m_uid:
            continue
        m_flags = re.search(rb"FLAGS \(([^)]*)\)", meta)
        m_size = re.search(rb"RFC822\.SIZE (\d+)", meta)
        flags = m_flags.group(1).decode("ascii", "replace").split() if m_flags else []
        h = email.message_from_bytes(item[1], policy=POLICY)
        out.append({
            "uid": m_uid.group(1).decode(),
            "date": _h(h, "Date"),
            "from": _h(h, "From"),
            "to": _h(h, "To"),
            "subject": _h(h, "Subject"),
            "unread": "\\Seen" not in flags,
            "flagged": "\\Flagged" in flags,
            "size": int(m_size.group(1)) if m_size else None,
        })
    out.sort(key=lambda r: int(r["uid"]), reverse=True)
    return out


class _HTMLText(HTMLParser):
    BLOCK = {"br", "p", "div", "tr", "li", "h1", "h2", "h3", "h4", "table", "blockquote"}

    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "head"):
            self.skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("script", "style", "head") and self.skip:
            self.skip -= 1

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data)


def _html_to_text(html: str) -> str:
    p = _HTMLText()
    p.feed(html)
    text = "".join(p.parts).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" ?\n ?", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _body_text(msg) -> str:
    part = msg.get_body(preferencelist=("plain", "html"))
    if part is None:
        return ""
    try:
        text = part.get_content()
    except Exception:
        text = (part.get_payload(decode=True) or b"").decode("utf-8", "replace")
    if part.get_content_type() == "text/html":
        text = _html_to_text(text)
    return text


def _imap_date(iso: str) -> str:
    d = date.fromisoformat(iso)
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year}"


# ---------- инструменты: чтение ----------

@tool()
def list_folders() -> list[dict]:
    """Список папок ящика Mail.ru с IMAP-флагами (\\Sent, \\Drafts, \\Trash, \\Junk)."""
    with _imap() as conn:
        return _folders(conn)


@tool()
def search_messages(
    folder: str = "INBOX",
    query: Optional[str] = None,
    sender: Optional[str] = None,
    subject: Optional[str] = None,
    since: Optional[str] = None,
    before: Optional[str] = None,
    unread_only: bool = False,
    limit: int = 20,
) -> list[dict]:
    """Поиск писем в папке, новые сверху. Возвращает uid, дату, отправителя, тему, флаги.

    query — подстрока в заголовках и теле; sender — в поле From; subject — в теме.
    since / before — даты в формате YYYY-MM-DD (since включительно, before не включая).
    Без условий возвращает последние письма папки.
    """
    limit = max(1, min(int(limit), 100))
    crit: list[str] = []
    literal: Optional[tuple[str, str]] = None
    client: dict[str, str] = {}

    if unread_only:
        crit.append("UNSEEN")
    if since:
        crit += ["SINCE", _imap_date(since)]
    if before:
        crit += ["BEFORE", _imap_date(before)]
    # В одной команде imaplib передаёт только один не-ASCII литерал:
    # приоритет у query, остальные не-ASCII условия фильтруются на клиенте.
    for key, val in (("TEXT", query), ("FROM", sender), ("SUBJECT", subject)):
        if not val:
            continue
        if val.isascii():
            crit += [key, _q(val)]
        elif literal is None:
            literal = (key, val)
        else:
            client[key] = val.casefold()

    with _imap() as conn:
        _select(conn, folder)
        if literal:
            conn.literal = literal[1].encode("utf-8")
            typ, data = conn.uid("SEARCH", "CHARSET", "UTF-8", *crit, literal[0])
        else:
            typ, data = conn.uid("SEARCH", *(crit or ["ALL"]))
        _ok(typ, data, "SEARCH")
        uids = [u.decode() for u in (data[0] or b"").split()]
        if not uids:
            return []
        take = min(limit * 5, 200) if client else limit
        uids = uids[-take:]
        typ, data = conn.uid(
            "FETCH", ",".join(uids),
            "(FLAGS RFC822.SIZE BODY.PEEK[HEADER.FIELDS (FROM TO SUBJECT DATE)])")
        _ok(typ, data, "FETCH")
        rows = _parse_fetch(data)

    if "FROM" in client:
        rows = [r for r in rows if client["FROM"] in r["from"].casefold()]
    if "SUBJECT" in client:
        rows = [r for r in rows if client["SUBJECT"] in r["subject"].casefold()]
    return rows[:limit]


@tool()
def get_message(uid: str, folder: str = "INBOX", max_chars: int = 20000) -> dict:
    """Письмо целиком: заголовки, текст (HTML приводится к тексту), список вложений.

    Не помечает письмо прочитанным. Текст обрезается до max_chars.
    """
    uid = _uid(uid)
    with _imap() as conn:
        _select(conn, folder)
        typ, data = conn.uid("FETCH", uid, "(BODY.PEEK[])")
        _ok(typ, data, "FETCH")
    raw = next((it[1] for it in data if isinstance(it, tuple)), None)
    if raw is None:
        raise ValueError(f"Письмо uid={uid} не найдено в папке {folder}")
    msg = email.message_from_bytes(raw, policy=POLICY)
    text = _body_text(msg)
    attachments = []
    for p in msg.iter_attachments():
        payload = p.get_payload(decode=True)
        attachments.append({"filename": p.get_filename(),
                            "type": p.get_content_type(),
                            "size": len(payload) if payload else None})
    return {
        "uid": uid,
        "folder": folder,
        "date": _h(msg, "Date"),
        "from": _h(msg, "From"),
        "to": _h(msg, "To"),
        "cc": _h(msg, "Cc"),
        "subject": _h(msg, "Subject"),
        "message_id": _h(msg, "Message-ID"),
        "body": text[:max_chars],
        "truncated": len(text) > max_chars,
        "attachments": attachments,
    }


ATTACH_DIR = os.environ.get(
    "MAILRU_ATTACH_DIR",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachments"))


def _safe_name(name: str) -> str:
    name = os.path.basename(name or "").strip().replace("\x00", "")
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name or "attachment"


@tool()
def save_attachments(uid: str, folder: str = "INBOX",
                     filenames: Optional[list[str]] = None,
                     dest_dir: Optional[str] = None) -> dict:
    """Сохранить вложения письма на диск.

    filenames — какие вложения сохранить (по имени); по умолчанию все.
    dest_dir — папка назначения; по умолчанию MAILRU_ATTACH_DIR
    (attachments/ рядом с сервером). Существующие файлы не перезаписываются:
    к имени добавляется суффикс. Возвращает пути сохранённых файлов.
    """
    uid = _uid(uid)
    with _imap() as conn:
        _select(conn, folder)
        typ, data = conn.uid("FETCH", uid, "(BODY.PEEK[])")
        _ok(typ, data, "FETCH")
    raw = next((it[1] for it in data if isinstance(it, tuple)), None)
    if raw is None:
        raise ValueError(f"Письмо uid={uid} не найдено в папке {folder}")
    msg = email.message_from_bytes(raw, policy=POLICY)
    out_dir = os.path.expanduser(dest_dir or ATTACH_DIR)
    os.makedirs(out_dir, exist_ok=True)
    saved, skipped = [], []
    for p in msg.iter_attachments():
        name = p.get_filename() or "attachment"
        if filenames and name not in filenames:
            skipped.append(name)
            continue
        payload = p.get_payload(decode=True) or b""
        base, ext = os.path.splitext(_safe_name(name))
        path = os.path.join(out_dir, base + ext)
        i = 1
        while os.path.exists(path):
            path = os.path.join(out_dir, f"{base} ({i}){ext}")
            i += 1
        with open(path, "wb") as f:
            f.write(payload)
        saved.append({"filename": name, "path": path, "size": len(payload)})
    if filenames:
        missing = [n for n in filenames if n not in {x["filename"] for x in saved}]
        if missing:
            raise ValueError(f"Вложения не найдены: {missing}; есть: {skipped}")
    return {"saved": saved, "dir": out_dir}


# ---------- инструменты: изменение ----------

@tool()
def set_flags(uid: str, folder: str = "INBOX",
              read: Optional[bool] = None, flagged: Optional[bool] = None) -> str:
    """Пометить письмо прочитанным/непрочитанным (read) и/или флажком (flagged)."""
    uid = _uid(uid)
    with _imap() as conn:
        _select(conn, folder, readonly=False)
        for value, flag in ((read, "\\Seen"), (flagged, "\\Flagged")):
            if value is None:
                continue
            typ, data = conn.uid("STORE", uid, "+FLAGS" if value else "-FLAGS", f"({flag})")
            _ok(typ, data, "STORE")
    return "ok"


def _move(conn, uid: str, destination: str) -> None:
    caps = _caps(conn)
    if "MOVE" in caps:
        typ, data = conn.uid("MOVE", uid, _mbox(destination))
        _ok(typ, data, "MOVE")
        return
    typ, data = conn.uid("COPY", uid, _mbox(destination))
    _ok(typ, data, "COPY")
    typ, data = conn.uid("STORE", uid, "+FLAGS", "(\\Deleted)")
    _ok(typ, data, "STORE")
    if "UIDPLUS" in caps:  # обычный EXPUNGE задел бы чужие \Deleted — не используем
        conn.uid("EXPUNGE", uid)


@tool()
def move_message(uid: str, destination: str, folder: str = "INBOX") -> str:
    """Переместить письмо из folder в папку destination (имя как в list_folders)."""
    uid = _uid(uid)
    with _imap() as conn:
        _select(conn, folder, readonly=False)
        _move(conn, uid, destination)
    return f"moved to {destination}"


def _uids(uids: list[str]) -> str:
    if not uids:
        raise ValueError("пустой список uid")
    return ",".join(_uid(u) for u in uids)


@tool()
def create_folder(name: str) -> str:
    """Создать папку в ящике. Вложенность — через «/» (например «Архив/2026»)."""
    with _imap() as conn:
        if name in {f["name"] for f in _folders(conn)}:
            return f"already exists: {name}"
        typ, data = conn.create(_mbox(name))
        _ok(typ, data, "CREATE")
        conn.subscribe(_mbox(name))
    return f"created: {name}"


@tool()
def move_messages(uids: list[str], destination: str, folder: str = "INBOX",
                  create: bool = True) -> str:
    """Переместить несколько писем из folder в destination одним вызовом.

    create=True — создать папку destination, если её нет.
    """
    uidset = _uids(uids)
    with _imap() as conn:
        if create and destination not in {f["name"] for f in _folders(conn)}:
            typ, data = conn.create(_mbox(destination))
            _ok(typ, data, "CREATE")
            conn.subscribe(_mbox(destination))
        _select(conn, folder, readonly=False)
        _move(conn, uidset, destination)
    return f"moved {len(uids)} to {destination}"


@tool()
def mark_read(uids: list[str], folder: str = "INBOX", read: bool = True) -> str:
    """Пометить несколько писем прочитанными (read=False — непрочитанными)."""
    uidset = _uids(uids)
    with _imap() as conn:
        _select(conn, folder, readonly=False)
        typ, data = conn.uid("STORE", uidset, "+FLAGS" if read else "-FLAGS", "(\\Seen)")
        _ok(typ, data, "STORE")
    return f"ok: {len(uids)}"


@tool()
def trash_message(uid: str, folder: str = "INBOX") -> str:
    """Переместить письмо в Корзину (не удаляет безвозвратно)."""
    uid = _uid(uid)
    with _imap() as conn:
        trash = _special(conn, "\\Trash", ["Корзина", "Trash"])
        _select(conn, folder, readonly=False)
        _move(conn, uid, trash)
    return f"moved to {trash}"


def _quote(orig) -> str:
    """Цитата исходного письма в стиле Mail.ru: шапка + текст с префиксом '> '."""
    head = [f"От кого: {_h(orig, 'From')}", f"Кому: {_h(orig, 'To')}"]
    if _h(orig, "Cc"):
        head.append(f"Копия: {_h(orig, 'Cc')}")
    head += [f"Дата: {_h(orig, 'Date')}", f"Тема: {_h(orig, 'Subject')}"]
    text = _body_text(orig).rstrip()
    quoted = "\n".join(("> " + ln) if ln else ">" for ln in (head + [""] + text.splitlines()))
    return "\n\n" + quoted + "\n"


def _build(to: list[str], subject: str, body: str,
           cc: Optional[list[str]], bcc: Optional[list[str]],
           reply_to_uid: Optional[str], reply_folder: str,
           quote: bool = True) -> EmailMessage:
    msg = EmailMessage(policy=email.policy.SMTP)
    msg["From"] = formataddr((FROM_NAME, USER)) if FROM_NAME else USER
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    if bcc:
        msg["Bcc"] = ", ".join(bcc)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=USER.split("@")[-1] or None)
    if reply_to_uid:
        with _imap() as conn:
            _select(conn, reply_folder)
            typ, data = conn.uid("FETCH", _uid(reply_to_uid), "(BODY.PEEK[])")
            _ok(typ, data, "FETCH")
        raw = next((it[1] for it in data if isinstance(it, tuple)), None)
        if raw is None:
            raise ValueError(f"Исходное письмо uid={reply_to_uid} не найдено")
        orig = email.message_from_bytes(raw, policy=POLICY)
        mid = _h(orig, "Message-ID").strip()
        if mid:
            msg["In-Reply-To"] = mid
            msg["References"] = (_h(orig, "References").strip() + " " + mid).strip()
        topic = _h(orig, "Thread-Topic").strip()
        if topic:
            msg["Thread-Topic"] = topic
        if quote:
            body = body + _quote(orig)
    msg.set_content(body)
    return msg


def _attach_files(msg: EmailMessage, paths: Optional[list[str]]) -> None:
    import mimetypes
    for path in paths or []:
        path = os.path.expanduser(path)
        if not os.path.isfile(path):
            raise ValueError(f"Файл не найден: {path}")
        ctype, _ = mimetypes.guess_type(path)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype,
                               filename=os.path.basename(path))


@tool()
def create_draft(to: list[str], subject: str, body: str,
                 cc: Optional[list[str]] = None, bcc: Optional[list[str]] = None,
                 reply_to_uid: Optional[str] = None, reply_folder: str = "INBOX",
                 attachments: Optional[list[str]] = None, quote: bool = True) -> str:
    """Сохранить черновик в папку «Черновики». Ничего не отправляет.

    reply_to_uid — uid письма, на которое это ответ (проставит In-Reply-To/References
    и по умолчанию процитирует исходное письмо со всей перепиской ниже текста).
    quote=False — не цитировать исходное письмо.
    attachments — пути к локальным файлам, которые приложить.
    """
    msg = _build(to, subject, body, cc, bcc, reply_to_uid, reply_folder, quote)
    _attach_files(msg, attachments)
    with _imap() as conn:
        drafts = _special(conn, "\\Drafts", ["Черновики", "Drafts"])
        typ, data = conn.append(_mbox(drafts), "(\\Draft)",
                                imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        _ok(typ, data, "APPEND")
    return f"draft saved to {drafts}"


def _save_sent(msg) -> str:
    """Положить копию отправленного письма в «Отправленные» (SMTP Mail.ru сам не сохраняет)."""
    try:
        with _imap() as conn:
            sent = _special(conn, "\\Sent", ["Отправленные", "Sent"])
            typ, data = conn.append(_mbox(sent), "(\\Seen)",
                                    imaplib.Time2Internaldate(time.time()), msg.as_bytes())
            _ok(typ, data, "APPEND")
        return f"copy saved to {sent}"
    except Exception as e:  # письмо уже ушло — не роняем вызов
        return f"copy NOT saved to Sent: {e}"


@tool()
def forward_message(uid: str, to: list[str], body: str = "", folder: str = "INBOX",
                    subject: Optional[str] = None, cc: Optional[list[str]] = None,
                    send: bool = False) -> str:
    """Переслать письмо вместе со всеми вложениями, добавив сверху свой текст body.

    send=False — сохранить в «Черновики»; send=True — отправить (нужен MAILRU_ALLOW_SEND=1).
    Получатели исходного письма не копируются: уходит только на to и cc.
    """
    uid = _uid(uid)
    if send and not ALLOW_SEND:
        raise RuntimeError("Отправка выключена (MAILRU_ALLOW_SEND != 1)")
    with _imap() as conn:
        _select(conn, folder)
        typ, data = conn.uid("FETCH", uid, "(BODY.PEEK[])")
        _ok(typ, data, "FETCH")
    raw = next((it[1] for it in data if isinstance(it, tuple)), None)
    if raw is None:
        raise ValueError(f"Письмо uid={uid} не найдено в папке {folder}")
    orig = email.message_from_bytes(raw, policy=POLICY)
    osubj = _h(orig, "Subject")
    if not subject:
        subject = osubj if osubj.lower().startswith(("fwd:", "fw:")) else "Fwd: " + osubj
    quoted = ("-------- Пересылаемое сообщение --------\n"
              f"От кого: {_h(orig, 'From')}\n"
              f"Кому: {_h(orig, 'To')}\n"
              f"Дата: {_h(orig, 'Date')}\n"
              f"Тема: {osubj}\n\n" + _body_text(orig))
    msg = EmailMessage(policy=email.policy.SMTP)
    msg["From"] = formataddr((FROM_NAME, USER)) if FROM_NAME else USER
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=USER.split("@")[-1] or None)
    msg.set_content((body.rstrip() + "\n\n" if body.strip() else "") + quoted)
    names = []
    for part in orig.iter_attachments():
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        maintype, _, subtype = part.get_content_type().partition("/")
        name = part.get_filename() or "attachment"
        msg.add_attachment(payload, maintype=maintype, subtype=subtype or "octet-stream",
                           filename=name)
        names.append(name)
    if send:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT,
                              context=ssl.create_default_context(), timeout=60) as smtp:
            smtp.login(USER, PASSWORD)
            refused = smtp.send_message(msg)
        return (f"sent to {msg['To']}, attachments: {names}, "
                f"refused: {refused or 'none'}; {_save_sent(msg)}")
    with _imap() as conn:
        drafts = _special(conn, "\\Drafts", ["Черновики", "Drafts"])
        typ, data = conn.append(_mbox(drafts), "(\\Draft)",
                                imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        _ok(typ, data, "APPEND")
    return f"draft saved to {drafts}, attachments: {names}"


# ---------- инструменты: календарь (CalDAV) ----------

@tool()
def calendar_list() -> list[dict]:
    """Календари Mail.ru, доступные по CalDAV: имя и адрес."""
    return cal.calendars(refresh=True)


@tool()
def calendar_events(
    date_from: str,
    date_to: Optional[str] = None,
    calendars: Optional[list[str]] = None,
    query: Optional[str] = None,
    include_declined: bool = False,
    details: bool = False,
) -> dict:
    """Встречи из календаря Mail.ru за период, по возрастанию времени.

    date_from / date_to — даты YYYY-MM-DD, обе включительно (без date_to — один день).
    calendars — подстроки имён календарей (по умолчанию все).
    query — подстрока в названии, месте или описании.
    У каждой встречи есть uid (для calendar_event и правок), у повторяющихся — occurrence.
    my_status: organizer — встреча ваша; accepted / declined / tentative / needs-action —
    вы приглашены. Настоящего организатора чужой встречи Mail.ru не отдаёт.
    Повторяющиеся встречи раскрыты в отдельные вхождения. Отменённые не выводятся;
    отклонённые — только при include_declined. details добавляет описание и статусы участников.
    Время — в поясе MAILRU_TZ (по умолчанию Europe/Moscow).
    """
    events, errors = cal.fetch(date_from, date_to, calendars)
    declined = sum(1 for e in events if e["my_status"] == "declined" and not e["cancelled"])
    # отклонённое вхождение серии Mail.ru отдаёт как отменённое — показываем его
    # вместе с отклонёнными
    rows = [e for e in events
            if include_declined or (not e["cancelled"] and e["my_status"] != "declined")]
    if query:
        q = query.casefold()
        rows = [e for e in rows if q in " ".join(
            (e["summary"], e["location"], e["description"])).casefold()]
    out = {"timezone": cal.TZ_NAME, "count": len(rows),
           "events": [cal.public(e, details) for e in rows]}
    if declined and not include_declined:
        out["declined_hidden"] = declined
    if errors:
        out["errors"] = errors
    return out


@tool()
def calendar_free_slots(
    date_from: str,
    date_to: Optional[str] = None,
    day_start: str = "10:00",
    day_end: str = "19:00",
    min_minutes: int = 30,
    weekdays_only: bool = True,
    calendars: Optional[list[str]] = None,
    extra_busy: Optional[list[dict]] = None,
    tentative_is_busy: bool = True,
    all_day_blocks: bool = False,
) -> dict:
    """Свободные слоты и встречи по дням в рабочие часы по календарю Mail.ru.

    date_from / date_to — даты YYYY-MM-DD включительно. day_start / day_end — рабочие часы HH:MM.
    min_minutes — минимальная длина слота. Прошедшее время сегодня не предлагается.
    extra_busy — дополнительные занятые интервалы из других календарей (например, Google):
    [{"start": ISO, "end": ISO, "summary": "..."}]; время без смещения считается местным.
    tentative_is_busy — считать ли занятыми встречи «под вопросом».
    all_day_blocks — блокируют ли день события на весь день (по умолчанию только упоминаются).
    Отклонённые, отменённые и помеченные «свободен» встречи время не занимают.
    """
    d1, d2, ws, we = cal._window(date_from, date_to)
    events, errors = cal.fetch(date_from, date_to, calendars)
    busy, blocked, per_day = [], set(), {}

    def note(day, row):
        per_day.setdefault(day.isoformat(), []).append(row)

    for e in events:
        if not e["busy"] or (e["tentative"] and not tentative_is_busy):
            continue
        if e["all_day"]:
            d = e["start_dt"].date()
            while d < e["end_dt"].date():
                note(d, {"all_day": True, "summary": e["summary"]})
                if all_day_blocks:
                    blocked.add(d)
                d += timedelta(days=1)
            continue
        busy.append((e["start_dt"], e["end_dt"]))
        row = {"start": e["start_dt"].strftime("%H:%M"), "end": e["end_dt"].strftime("%H:%M"),
               "summary": e["summary"]}
        if e["tentative"]:
            row["tentative"] = True
        if e["link"]:
            row["link"] = e["link"]
        note(e["start_dt"].date(), row)
    for item in extra_busy or []:
        s, en = cal._iso(item["start"]), cal._iso(item["end"])
        if en <= ws or s >= we or en <= s:
            continue
        s, en = s.astimezone(ws.tzinfo), en.astimezone(ws.tzinfo)
        busy.append((s, en))
        note(s.date(), {"start": s.strftime("%H:%M"), "end": en.strftime("%H:%M"),
                        "summary": item.get("summary") or "(занято)", "source": "extra"})

    days = cal.free_slots(busy, d1, d2, day_start, day_end, min_minutes,
                          weekdays_only, blocked_days=blocked)
    for day in days:
        rows = per_day.get(day["date"], [])
        day["meetings"] = sorted(rows, key=lambda r: r.get("start", ""))
    out = {"timezone": cal.TZ_NAME, "work_hours": f"{day_start}-{day_end}",
           "min_minutes": min_minutes, "days": days}
    if errors:
        out["errors"] = errors
    return out


if ALLOW_SEND:
    @mcp.tool()
    def send_message(to: list[str], subject: str, body: str,
                     cc: Optional[list[str]] = None, bcc: Optional[list[str]] = None,
                     reply_to_uid: Optional[str] = None, reply_folder: str = "INBOX",
                     attachments: Optional[list[str]] = None, quote: bool = True) -> str:
        """Отправить письмо с ящика Mail.ru через SMTP. Действие необратимо.

        reply_to_uid — uid письма, на которое это ответ (по умолчанию цитирует
        исходное письмо со всей перепиской ниже текста; quote=False — без цитаты).
        attachments — пути к локальным файлам, которые приложить.
        """
        msg = _build(to, subject, body, cc, bcc, reply_to_uid, reply_folder, quote)
        _attach_files(msg, attachments)
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT,
                              context=ssl.create_default_context(), timeout=30) as smtp:
            smtp.login(USER, PASSWORD)
            refused = smtp.send_message(msg)
        saved = _save_sent(msg)
        names = [os.path.basename(os.path.expanduser(a)) for a in attachments or []]
        att = f", attachments: {names}" if names else ""
        if refused:
            return f"sent, refused recipients: {refused}{att}; {saved}"
        return f"sent, message-id {msg['Message-ID']}{att}; {saved}"


# ---------- инструменты: календарь, одна встреча и запись ----------

@tool()
def calendar_event(uid: str, occurrence: Optional[str] = None) -> dict:
    """Полные данные одной встречи по uid (uid есть в выдаче calendar_events).

    Все участники со статусами, полное описание, правило повтора, изменённые вхождения.
    occurrence — дата YYYY-MM-DD или время начала конкретного вхождения серии; с ним
    возвращаются данные именно этого вхождения и встреча ищется быстрее.
    Организатор: Mail.ru не отдаёт его по CalDAV. my_status = organizer означает, что
    встреча ваша; иначе вы приглашены, а кто организатор — неизвестно.
    """
    return cal.event(uid, occurrence)


@tool()
def calendar_create_event(
    summary: str,
    start: str,
    end: Optional[str] = None,
    attendees: Optional[list[str]] = None,
    description: str = "",
    location: str = "",
    rrule: str = "",
    calendar: Optional[str] = None,
    busy: bool = True,
) -> dict:
    """Создать встречу в календаре Mail.ru. Вы — организатор.

    start / end — время YYYY-MM-DDTHH:MM в поясе MAILRU_TZ (без end — 30 минут) либо даты
    YYYY-MM-DD для события на весь день (обе включительно).
    attendees — адреса участников; им уходит приглашение от вашего имени.
    rrule — правило повтора iCalendar, например FREQ=WEEKLY;BYDAY=MO,WE или
    FREQ=DAILY;COUNT=5. calendar — подстрока имени календаря (по умолчанию «Основной»).
    busy=False — встреча не занимает время. Ссылку на звонок кладите в location или description.
    """
    return cal.create_event(summary, start, end, calendar, description, location,
                            attendees, rrule, busy)


@tool()
def calendar_update_event(
    uid: str,
    occurrence: Optional[str] = None,
    summary: Optional[str] = None,
    start: Optional[str] = None,
    end: Optional[str] = None,
    description: Optional[str] = None,
    location: Optional[str] = None,
    add_attendees: Optional[list[str]] = None,
    remove_attendees: Optional[list[str]] = None,
    rrule: Optional[str] = None,
    busy: Optional[bool] = None,
) -> dict:
    """Изменить встречу: название, время, описание, место, участников, повтор.

    Меняются только переданные поля. Пустая строка в description / location / rrule очищает поле.
    occurrence — дата YYYY-MM-DD или время начала одного вхождения серии: тогда меняется
    только оно (перенос одной встречи из серии). Без occurrence меняется вся серия; start
    тогда задаёт дату и время первого вхождения.
    start без end сохраняет длительность. Формат времени как в calendar_create_event.
    Менять имеет смысл свои встречи; правка чужой останется только в вашем календаре.
    """
    return cal.update_event(uid, occurrence, summary, start, end, description, location,
                            add_attendees, remove_attendees, rrule, busy)


@tool()
def calendar_delete_event(uid: str, occurrence: Optional[str] = None) -> dict:
    """Удалить встречу или одно вхождение серии. Необратимо.

    occurrence — дата YYYY-MM-DD или время начала вхождения: удаляется только оно.
    Без occurrence удаляется встреча целиком, у повторяющейся — вся серия.
    Своя встреча отменяется для участников; чужая убирается из вашего календаря
    (чтобы отказаться от участия, используйте calendar_respond).
    """
    return cal.delete_event(uid, occurrence)


@tool()
def calendar_respond(uid: str, status: str, occurrence: Optional[str] = None,
                     comment: Optional[str] = None) -> dict:
    """Ответить на приглашение: status = accepted, declined или tentative.

    occurrence — дата YYYY-MM-DD или время начала вхождения: ответ только на него.
    Без occurrence ответ относится ко всей встрече или серии.
    comment — причина или пояснение к ответу (до 300 знаков). Записывается в ответ;
    проверьте по calendar_event (поле my_comment), сохранил ли его сервер.
    Отклонённое вхождение серии Mail.ru показывает в вашем календаре как отменённое;
    в calendar_events оно видно только с include_declined.
    """
    return cal.respond(uid, status, occurrence, comment)


if __name__ == "__main__":
    mcp.run()
