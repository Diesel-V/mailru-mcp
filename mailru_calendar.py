"""Календарь Mail.ru по CalDAV: календари, события, свободные слоты.

Переменные окружения:
  MAILRU_USER / MAILRU_APP_PASSWORD   те же, что для почты
  MAILRU_CALDAV_PASSWORD   отдельный пароль приложения для календаря (необязательно)
  MAILRU_CALDAV_URL        по умолчанию https://calendar.mail.ru
  MAILRU_TZ                часовой пояс вывода, по умолчанию Europe/Moscow
  MAILRU_CALENDAR_DEFAULT  календарь для новых встреч (по умолчанию «Основной»)
  MAILRU_CALENDAR_READONLY 1 — запретить создание, правку, удаление и ответы
"""
from __future__ import annotations

import base64
import http.client
import os
import re
import ssl
import urllib.error
import urllib.request
import uuid
import xml.etree.ElementTree as ET
from datetime import date, datetime, time as dtime, timedelta, timezone
from typing import Optional
from urllib.parse import urljoin, urlsplit
from xml.sax.saxutils import escape as _xml
from zoneinfo import ZoneInfo

from dateutil.rrule import rrulestr

USER = os.environ.get("MAILRU_CALDAV_USER") or os.environ.get("MAILRU_USER", "")
PASSWORD = (os.environ.get("MAILRU_CALDAV_PASSWORD")
            or os.environ.get("MAILRU_APP_PASSWORD", ""))
BASE_URL = os.environ.get("MAILRU_CALDAV_URL", "https://calendar.mail.ru").rstrip("/")
TZ_NAME = os.environ.get("MAILRU_TZ", "Europe/Moscow")

D = "{DAV:}"
C = "{urn:ietf:params:xml:ns:caldav}"
WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
MAX_DAYS = 62

# Имена поясов Windows (приглашения из Outlook) -> IANA
WIN_TZ = {
    "Russian Standard Time": "Europe/Moscow",
    "Kaliningrad Standard Time": "Europe/Kaliningrad",
    "Russia Time Zone 3": "Europe/Samara",
    "Ekaterinburg Standard Time": "Asia/Yekaterinburg",
    "N. Central Asia Standard Time": "Asia/Novosibirsk",
    "North Asia Standard Time": "Asia/Krasnoyarsk",
    "North Asia East Standard Time": "Asia/Irkutsk",
    "Yakutsk Standard Time": "Asia/Yakutsk",
    "Vladivostok Standard Time": "Asia/Vladivostok",
    "Belarus Standard Time": "Europe/Minsk",
    "Central Asia Standard Time": "Asia/Almaty",
    "Arabian Standard Time": "Asia/Dubai",
    "Turkey Standard Time": "Europe/Istanbul",
    "FLE Standard Time": "Europe/Kiev",
    "GMT Standard Time": "Europe/London",
    "W. Europe Standard Time": "Europe/Berlin",
    "Central European Standard Time": "Europe/Warsaw",
    "Romance Standard Time": "Europe/Paris",
    "Eastern Standard Time": "America/New_York",
    "Pacific Standard Time": "America/Los_Angeles",
    "UTC": "UTC",
}

CONF_RE = re.compile(
    r"https?://[^\s<>\"'\\]*(?:telemost|zoom\.|meet\.google|teams\.microsoft|teams\.live|"
    r"jazz\.sber|salutejazz|sberjazz|ktalk|dion\.vc|trueconf|webinar\.ru|mts-link|"
    r"calls\.mail\.ru|vk\.com/call|talk\.kontur|meet\.)[^\s<>\"'\\]*", re.I)


def local_tz() -> ZoneInfo:
    return ZoneInfo(TZ_NAME)


# ---------- HTTP / WebDAV ----------

class _KeepMethod(urllib.request.HTTPRedirectHandler):
    """Редиректы с сохранением метода и тела (PROPFIND/REPORT)."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if code not in (301, 302, 303, 307, 308):
            return None
        old, new = urlsplit(req.full_url), urlsplit(newurl)
        if new.scheme != "https":
            return None
        hdrs = dict(req.header_items())
        if new.hostname != old.hostname:  # пароль чужому хосту не отдаём
            hdrs = {k: v for k, v in hdrs.items() if k.lower() != "authorization"}
        return urllib.request.Request(newurl, data=req.data, headers=hdrs,
                                      method=req.get_method())


_opener = urllib.request.build_opener(
    urllib.request.HTTPSHandler(context=ssl.create_default_context()), _KeepMethod())


def _dav(method: str, url: str, body: str, depth: str) -> tuple[str, ET.Element]:
    if not USER or not PASSWORD:
        raise RuntimeError("Не заданы MAILRU_USER / MAILRU_APP_PASSWORD")
    token = base64.b64encode(f"{USER}:{PASSWORD}".encode("utf-8")).decode()
    req = urllib.request.Request(url, data=body.encode("utf-8"), method=method, headers={
        "Authorization": "Basic " + token,
        "Depth": depth,
        "Content-Type": "application/xml; charset=utf-8",
        "User-Agent": "mailru-mcp/1.0",
    })
    try:
        with _opener.open(req, timeout=30) as resp:
            final, raw = resp.geturl(), resp.read()
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise RuntimeError(
                "CalDAV 401: пароль приложения не принят для календаря. Нужен пароль "
                "приложения Mail.ru с доступом к календарю (CalDAV).") from None
        if e.code == 403:
            raise RuntimeError(
                "CalDAV 403: доступ к календарю по CalDAV запрещён "
                "(возможно, закрыт администратором домена).") from None
        raise RuntimeError(f"CalDAV {method} {url}: HTTP {e.code} {e.reason}") from None
    except urllib.error.URLError as e:
        raise RuntimeError(f"CalDAV {method} {url}: {e.reason}") from None
    except (http.client.HTTPException, OSError) as e:
        raise RuntimeError(f"CalDAV {method} {url}: {e!r}") from None
    try:
        return final, ET.fromstring(raw)
    except ET.ParseError:
        raise RuntimeError(f"CalDAV {method} {url}: ответ не XML: {raw[:200]!r}") from None


def _propfind(url: str, props: str, depth: str = "0") -> tuple[str, ET.Element]:
    body = ('<?xml version="1.0" encoding="utf-8"?>'
            '<D:propfind xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">'
            f"<D:prop>{props}</D:prop></D:propfind>")
    return _dav("PROPFIND", url, body, depth)


def _href_in(root: ET.Element, prop: str) -> Optional[str]:
    el = root.find(f".//{prop}/{D}href")
    return el.text.strip() if el is not None and el.text else None


_calendars_cache: Optional[list[dict]] = None


def calendars(refresh: bool = False) -> list[dict]:
    """Календари пользователя: [{name, url}]."""
    global _calendars_cache
    if _calendars_cache is not None and not refresh:
        return _calendars_cache
    principal = None
    errors = []
    for start in (BASE_URL + "/.well-known/caldav", BASE_URL + "/"):
        try:
            final, root = _propfind(start, "<D:current-user-principal/>")
        except RuntimeError as e:
            if "401" in str(e) or "403" in str(e):
                raise
            errors.append(str(e))
            continue
        href = _href_in(root, f"{D}current-user-principal")
        if href:
            principal = urljoin(final, href)
            break
    if not principal:
        raise RuntimeError("CalDAV: не найден principal. " + "; ".join(errors))
    final, root = _propfind(principal, "<C:calendar-home-set/>")
    home = _href_in(root, f"{C}calendar-home-set")
    if not home:
        raise RuntimeError("CalDAV: не найден calendar-home-set")
    home = urljoin(final, home)
    final, root = _propfind(
        home, "<D:resourcetype/><D:displayname/><C:supported-calendar-component-set/>", "1")
    out = []
    for resp in root.findall(f"{D}response"):
        if resp.find(f".//{D}resourcetype/{C}calendar") is None:
            continue
        comps = [c.get("name") for c in resp.findall(f".//{C}comp")]
        if comps and "VEVENT" not in comps:
            continue
        href = resp.findtext(f"{D}href") or ""
        name = (resp.findtext(f".//{D}displayname") or "").strip()
        url = urljoin(final, href.strip())
        out.append({"name": name or url.rstrip("/").rsplit("/", 1)[-1], "url": url})
    _calendars_cache = out
    return out


def _query(cal_url: str, flt: str) -> list[dict]:
    """calendar-query с фильтром по VEVENT -> [{url, etag, ics}]."""
    body = ('<?xml version="1.0" encoding="utf-8"?>'
            '<C:calendar-query xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">'
            "<D:prop><D:getetag/><C:calendar-data/></D:prop>"
            '<C:filter><C:comp-filter name="VCALENDAR"><C:comp-filter name="VEVENT">'
            f"{flt}"
            "</C:comp-filter></C:comp-filter></C:filter></C:calendar-query>")
    final, root = _dav("REPORT", cal_url, body, "1")
    out = []
    for resp in root.findall(f"{D}response"):
        data = resp.find(f".//{C}calendar-data")
        if data is None or not data.text:
            continue
        href = (resp.findtext(f"{D}href") or "").strip()
        out.append({"url": urljoin(final, href),
                    "etag": (resp.findtext(f".//{D}getetag") or "").strip(),
                    "ics": data.text})
    return out


def _time_range(ws: datetime, we: datetime) -> str:
    fmt = "%Y%m%dT%H%M%SZ"
    return (f'<C:time-range start="{ws.astimezone(timezone.utc).strftime(fmt)}" '
            f'end="{we.astimezone(timezone.utc).strftime(fmt)}"/>')


def _report(cal_url: str, ws: datetime, we: datetime) -> list[str]:
    return [i["ics"] for i in _query(cal_url, _time_range(ws, we))]


# ---------- iCalendar ----------

def _split_line(line: str) -> tuple[str, dict, str]:
    """NAME;PARAM=val;PARAM="a:b":value -> (NAME, {PARAM: val}, value)."""
    in_q, i = False, 0
    for i, ch in enumerate(line):
        if ch == '"':
            in_q = not in_q
        elif ch == ":" and not in_q:
            break
    else:
        return line.upper(), {}, ""
    head, value = line[:i], line[i + 1:]
    parts, cur, in_q = [], "", False
    for ch in head:
        if ch == '"':
            in_q = not in_q
        elif ch == ";" and not in_q:
            parts.append(cur)
            cur = ""
            continue
        cur += ch
    parts.append(cur)
    params = {}
    for p in parts[1:]:
        k, _, v = p.partition("=")
        params[k.upper()] = v.strip('"')
    return parts[0].upper(), params, value


def parse_ics(text: str) -> list[dict]:
    """Дерево компонентов: {name, props: [(name, params, value)], children}."""
    lines = re.sub(r"\r?\n[ \t]", "", text).splitlines()
    roots: list[dict] = []
    stack: list[dict] = []
    for line in lines:
        if not line.strip():
            continue
        name, params, value = _split_line(line)
        if name == "BEGIN":
            comp = {"name": value.strip().upper(), "props": [], "children": []}
            (stack[-1]["children"] if stack else roots).append(comp)
            stack.append(comp)
        elif name == "END":
            if stack:
                stack.pop()
        elif stack:
            stack[-1]["props"].append((name, params, value))
    return roots


def _get(comp: dict, name: str):
    for n, p, v in comp["props"]:
        if n == name:
            return p, v
    return None


def _all(comp: dict, name: str):
    return [(p, v) for n, p, v in comp["props"] if n == name]


def _text(v: str) -> str:
    return (v.replace("\\n", "\n").replace("\\N", "\n").replace("\\,", ",")
            .replace("\\;", ";").replace("\\\\", "\\")).strip()


def _offset(v: str) -> timedelta:
    sign = -1 if v.startswith("-") else 1
    v = v.lstrip("+-")
    return sign * timedelta(hours=int(v[:2]), minutes=int(v[2:4] or 0))


def _vtimezones(vcal: dict) -> dict:
    out = {}
    for comp in vcal["children"]:
        if comp["name"] != "VTIMEZONE":
            continue
        tzid = _get(comp, "TZID")
        if not tzid:
            continue
        subs = ([c for c in comp["children"] if c["name"] == "STANDARD"]
                or comp["children"])
        for sub in subs:
            off = _get(sub, "TZOFFSETTO")
            if off:
                try:
                    out[tzid[1]] = timezone(_offset(off[1].strip()))
                except ValueError:
                    pass
                break
    return out


def _tz(tzid: str, vtz: dict):
    tzid = tzid.strip()
    for name in (tzid, WIN_TZ.get(tzid, "")):
        if name:
            try:
                return ZoneInfo(name)
            except Exception:
                pass
    return vtz.get(tzid) or local_tz()


def _parse_dt(params: dict, value: str, vtz: dict):
    """-> (date, True) для событий на весь день или (aware datetime, False)."""
    value = value.strip()
    if params.get("VALUE", "").upper() == "DATE" or re.fullmatch(r"\d{8}", value):
        return datetime.strptime(value[:8], "%Y%m%d").date(), True
    dt = datetime.strptime(value.rstrip("Z")[:15], "%Y%m%dT%H%M%S")
    if value.endswith("Z"):
        return dt.replace(tzinfo=timezone.utc), False
    tz = _tz(params["TZID"], vtz) if params.get("TZID") else local_tz()
    return dt.replace(tzinfo=tz), False


def _duration(v: str) -> timedelta:
    m = re.fullmatch(r"([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?",
                     v.strip().upper())
    if not m:
        return timedelta(0)
    w, d, h, mi, s = (int(x or 0) for x in m.groups()[1:])
    td = timedelta(weeks=w, days=d, hours=h, minutes=mi, seconds=s)
    return -td if m.group(1) == "-" else td


def _wall(dt_or_date, is_date: bool, tz) -> datetime:
    """Наивное «настенное» время в поясе события."""
    if is_date:
        return datetime.combine(dt_or_date, dtime())
    return dt_or_date.astimezone(tz).replace(tzinfo=None)


def _fix_until(rule: str, tz) -> str:
    """dateutil требует UNTIL того же вида, что и DTSTART; разворачиваем в настенном времени."""
    def repl(m: re.Match) -> str:
        v = m.group(1)
        if len(v) == 8:
            u = datetime.strptime(v, "%Y%m%d").replace(hour=23, minute=59, second=59)
        else:
            u = datetime.strptime(v.rstrip("Z")[:15], "%Y%m%dT%H%M%S")
            if v.endswith("Z"):
                u = u.replace(tzinfo=timezone.utc).astimezone(tz).replace(tzinfo=None)
        return "UNTIL=" + u.strftime("%Y%m%dT%H%M%S")
    return re.sub(r"UNTIL=([0-9TZ]+)", repl, rule, flags=re.I)


def _mail(v: str) -> str:
    return re.sub(r"^mailto:", "", v.strip(), flags=re.I).lower()


def _event_fields(ev: dict, me: str) -> dict:
    def txt(name: str) -> str:
        got = _get(ev, name)
        return _text(got[1]) if got else ""

    attendees, my_status, my_comment = [], "", ""
    for p, v in _all(ev, "ATTENDEE"):
        addr = _mail(v)
        status = p.get("PARTSTAT", "NEEDS-ACTION").lower()
        row = {"name": p.get("CN") or addr, "email": addr, "status": status}
        if p.get("X-RESPONSE-COMMENT"):
            row["comment"] = p["X-RESPONSE-COMMENT"]
        attendees.append(row)
        if addr == me:
            my_status = status
            my_comment = p.get("X-RESPONSE-COMMENT", "")
    # Mail.ru по CalDAV всегда отдаёт ORGANIZER = владелец календаря, настоящего
    # организатора в данных нет. Достоверно известно только одно: если меня нет среди
    # участников, встреча моя. Чужой адрес в ORGANIZER (другие серверы) показываем как есть.
    org = _get(ev, "ORGANIZER")
    org_mail = _mail(org[1]) if org else ""
    organizer = ""
    if org_mail and org_mail != me:
        organizer = org[0].get("CN") or org_mail
    elif org_mail == me and not my_status:
        my_status = "organizer"
    link = ""
    for _, _, v in ev["props"]:
        m = CONF_RE.search(v)
        if m:
            link = m.group(0).rstrip(".,;)")
            break
    status = txt("STATUS").upper()
    transparent = txt("TRANSP").upper() == "TRANSPARENT"
    return {
        "summary": txt("SUMMARY") or "(без названия)",
        "location": txt("LOCATION"),
        "description": txt("DESCRIPTION"),
        "organizer": organizer,
        "attendees": attendees,
        "my_status": my_status,
        "my_comment": my_comment or txt("COMMENT"),
        "link": link,
        "cancelled": status == "CANCELLED",
        "tentative": status == "TENTATIVE" or my_status == "tentative",
        "busy": not transparent and status != "CANCELLED" and my_status != "declined",
        "uid": txt("UID"),
    }


def expand(ics_text: str, ws: datetime, we: datetime, me: str = "",
           calendar: str = "") -> list[dict]:
    """События из одного ICS-ресурса, попадающие в окно [ws, we), с раскрытием повторов."""
    loc = local_tz()
    me = me.lower()
    out: list[dict] = []
    for vcal in parse_ics(ics_text):
        if vcal["name"] != "VCALENDAR":
            continue
        vtz = _vtimezones(vcal)
        by_uid: dict[str, list[dict]] = {}
        for comp in vcal["children"]:
            if comp["name"] == "VEVENT":
                uid = _get(comp, "UID")
                by_uid.setdefault(uid[1] if uid else str(id(comp)), []).append(comp)
        for group in by_uid.values():
            masters = [e for e in group if not _get(e, "RECURRENCE-ID")]
            overrides = [e for e in group if _get(e, "RECURRENCE-ID")]
            for ev in masters + overrides:
                ds = _get(ev, "DTSTART")
                if not ds:
                    continue
                try:
                    start, is_date = _parse_dt(ds[0], ds[1], vtz)
                except (ValueError, KeyError):
                    continue
                tz = loc if is_date else start.tzinfo
                s0 = _wall(start, is_date, tz)
                de, du = _get(ev, "DTEND"), _get(ev, "DURATION")
                if de:
                    try:
                        end, end_is_date = _parse_dt(de[0], de[1], vtz)
                        dur = _wall(end, end_is_date, tz) - s0
                    except (ValueError, KeyError):
                        dur = timedelta(0)
                elif du:
                    dur = _duration(du[1])
                else:
                    dur = timedelta(days=1) if is_date else timedelta(0)
                if dur < timedelta(0):
                    dur = timedelta(0)

                lo = ws.astimezone(tz).replace(tzinfo=None) - dur
                hi = we.astimezone(tz).replace(tzinfo=None)
                rules = [v for _, v in _all(ev, "RRULE")]
                rdates = _all(ev, "RDATE")
                is_master = ev in masters
                recurring = bool(rules or rdates) and is_master
                if recurring:
                    starts = set()
                    if lo <= s0 <= hi:
                        starts.add(s0)
                    for rule in rules:
                        try:
                            rr = rrulestr("RRULE:" + _fix_until(rule, tz), dtstart=s0)
                            starts.update(rr.between(lo, hi, inc=True))
                        except Exception:
                            continue
                    skip = set()
                    for name, coll in (("RDATE", starts), ("EXDATE", skip)):
                        for p, v in _all(ev, name):
                            for item in v.split(","):
                                try:
                                    d, d_is_date = _parse_dt(p, item.split("/")[0], vtz)
                                    w = _wall(d, d_is_date, tz)
                                except (ValueError, KeyError):
                                    continue
                                if name == "EXDATE" or lo <= w <= hi:
                                    coll.add(w)
                    for o in overrides:
                        p, v = _get(o, "RECURRENCE-ID")
                        try:
                            d, d_is_date = _parse_dt(p, v, vtz)
                            skip.add(_wall(d, d_is_date, tz))
                        except (ValueError, KeyError):
                            pass
                    if is_date:  # EXDATE для событий на весь день сравниваем по дате
                        skip_days = {s.date() for s in skip}
                        starts = {s for s in starts if s.date() not in skip_days}
                    else:
                        starts -= skip
                else:
                    starts = {s0}

                rid_wall = None
                if not is_master:
                    rp = _get(ev, "RECURRENCE-ID")
                    try:
                        rd, rd_is_date = _parse_dt(rp[0], rp[1], vtz)
                        rid_wall = _wall(rd, rd_is_date, tz)
                    except (ValueError, KeyError):
                        pass
                fields = None
                for s in sorted(starts):
                    st = s.replace(tzinfo=tz).astimezone(loc)
                    en = (s + dur).replace(tzinfo=tz).astimezone(loc)
                    hit = (st < we and en > ws) if en > st else (ws <= st < we)
                    if not hit:
                        continue
                    if fields is None:
                        fields = _event_fields(ev, me)
                    out.append({**fields, "start_dt": st, "end_dt": en, "all_day": is_date,
                                "recurring": recurring or not is_master,
                                "override": not is_master,
                                "occ_dt": ((rid_wall or s).replace(tzinfo=tz).astimezone(loc)
                                           if (recurring or not is_master) else None),
                                "calendar": calendar})
    return out


# ---------- сбор событий ----------

def _day(s: str) -> date:
    try:
        return date.fromisoformat(str(s).strip()[:10])
    except ValueError:
        raise ValueError(f"Дата должна быть в формате YYYY-MM-DD: {s!r}") from None


def _window(date_from: str, date_to: Optional[str]) -> tuple[date, date, datetime, datetime]:
    loc = local_tz()
    d1 = _day(date_from)
    d2 = _day(date_to) if date_to else d1
    if d2 < d1:
        raise ValueError("date_to раньше date_from")
    if (d2 - d1).days >= MAX_DAYS:
        raise ValueError(f"Период не больше {MAX_DAYS} дней")
    ws = datetime.combine(d1, dtime(), tzinfo=loc)
    we = datetime.combine(d2 + timedelta(days=1), dtime(), tzinfo=loc)
    return d1, d2, ws, we


def fetch(date_from: str, date_to: Optional[str] = None,
          only: Optional[list[str]] = None) -> tuple[list[dict], list[str]]:
    """Все события за период из выбранных календарей. -> (события, ошибки)."""
    _, _, ws, we = _window(date_from, date_to)
    cals = calendars()
    if only:
        keys = [k.casefold() for k in only]
        cals = [c for c in cals if any(k in c["name"].casefold() for k in keys)]
        if not cals:
            raise ValueError(f"Нет календарей с именами {only}; "
                             f"есть: {[c['name'] for c in calendars()]}")
    events, errors, seen = [], [], set()
    for cal in cals:
        try:
            for ics in _report(cal["url"], ws, we):
                for e in expand(ics, ws, we, USER, cal["name"]):
                    key = (e["uid"], e["start_dt"])
                    if key not in seen:
                        seen.add(key)
                        events.append(e)
        except RuntimeError as err:
            errors.append(f"{cal['name']}: {err}")
    events.sort(key=lambda e: (e["start_dt"], e["end_dt"]))
    return events, errors


def public(e: dict, details: bool = False) -> dict:
    """Событие в виде для выдачи."""
    if e["all_day"]:
        start = e["start_dt"].date().isoformat()
        end = (e["end_dt"] - timedelta(days=1)).date().isoformat()
    else:
        start = e["start_dt"].isoformat(timespec="minutes")
        end = e["end_dt"].isoformat(timespec="minutes")
    row = {"summary": e["summary"], "start": start, "end": end}
    if e["all_day"]:
        row["all_day"] = True
    for key in ("location", "link", "organizer", "my_status"):
        if e[key]:
            row[key] = e[key]
    if e["tentative"]:
        row["tentative"] = True
    if e["cancelled"]:
        row["cancelled"] = True
    if not e["busy"]:
        row["busy"] = False
    if e["recurring"]:
        row["recurring"] = True
        if e.get("occ_dt"):
            row["occurrence"] = (e["occ_dt"].date().isoformat() if e["all_day"]
                                 else e["occ_dt"].isoformat(timespec="minutes"))
    row["uid"] = e["uid"]
    row["attendees_count"] = len(e["attendees"])
    row["calendar"] = e["calendar"]
    if details:
        row["attendees"] = [f"{a['name']} ({a['status']})" for a in e["attendees"]]
        row["description"] = e["description"][:2000]
    elif e["attendees"]:
        row["attendees"] = [a["name"] for a in e["attendees"][:8]]
    return row


# ---------- свободные слоты ----------

def _hm(s: str) -> dtime:
    m = re.fullmatch(r"(\d{1,2}):(\d{2})", str(s).strip())
    if not m:
        raise ValueError(f"Время должно быть в формате HH:MM: {s!r}")
    return dtime(int(m.group(1)), int(m.group(2)))


def _iso(s: str) -> datetime:
    dt = datetime.fromisoformat(str(s).strip().replace("Z", "+00:00"))
    return dt if dt.tzinfo else dt.replace(tzinfo=local_tz())


def free_slots(busy: list[tuple[datetime, datetime]], d1: date, d2: date,
               day_start: str = "10:00", day_end: str = "19:00", min_minutes: int = 30,
               weekdays_only: bool = True, now: Optional[datetime] = None,
               blocked_days: Optional[set] = None) -> list[dict]:
    """Свободные окна по дням в рабочие часы. busy — занятые интервалы (aware datetime)."""
    loc = local_tz()
    t1, t2 = _hm(day_start), _hm(day_end)
    if t2 <= t1:
        raise ValueError("day_end должен быть позже day_start")
    need = timedelta(minutes=max(5, int(min_minutes)))
    now = (now or datetime.now(loc)).astimezone(loc)
    if now.minute % 5 or now.second or now.microsecond:  # округляем вверх до 5 минут
        now = now.replace(second=0, microsecond=0) + timedelta(minutes=5 - now.minute % 5)
    merged: list[list[datetime]] = []
    for s, e in sorted((s.astimezone(loc), e.astimezone(loc)) for s, e in busy if e > s):
        if merged and s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    days = []
    d = d1
    while d <= d2:
        if weekdays_only and d.weekday() >= 5:
            d += timedelta(days=1)
            continue
        lo = max(datetime.combine(d, t1, tzinfo=loc), now)
        hi = datetime.combine(d, t2, tzinfo=loc)
        slots = []
        if not (blocked_days and d in blocked_days):
            cur = lo
            for s, e in merged:
                if e <= cur or s >= hi:
                    continue
                if s - cur >= need:
                    slots.append((cur, s))
                cur = max(cur, e)
            if hi - cur >= need:
                slots.append((cur, hi))
        days.append({
            "date": d.isoformat(),
            "weekday": WEEKDAYS[d.weekday()],
            "free": [{"start": s.strftime("%H:%M"), "end": e.strftime("%H:%M"),
                      "minutes": int((e - s).total_seconds() // 60)} for s, e in slots],
            "free_minutes": int(sum((e - s).total_seconds() for s, e in slots) // 60),
        })
        d += timedelta(days=1)
    return days


# ======================================================================
# Запись: одна встреча, создание, правка, удаление, ответ на приглашение
# ======================================================================

READONLY = os.environ.get("MAILRU_CALENDAR_READONLY", "") == "1"
DEFAULT_CALENDAR = os.environ.get("MAILRU_CALENDAR_DEFAULT", "")
_RECUR = ("RRULE", "EXDATE", "RDATE", "EXRULE")
_MAIL_RE = re.compile(r"[^@\s<>,;:\"]+@[^@\s<>,;:\"]+\.[^@\s<>,;:\"]+")


def _need_write() -> None:
    if READONLY:
        raise RuntimeError("Запись в календарь выключена (MAILRU_CALENDAR_READONLY=1)")


def _http(method: str, url: str, body: Optional[str] = None,
          headers: Optional[dict] = None) -> int:
    if not USER or not PASSWORD:
        raise RuntimeError("Не заданы MAILRU_USER / MAILRU_APP_PASSWORD")
    token = base64.b64encode(f"{USER}:{PASSWORD}".encode("utf-8")).decode()
    hdrs = {"Authorization": "Basic " + token, "User-Agent": "mailru-mcp/1.0"}
    hdrs.update(headers or {})
    req = urllib.request.Request(
        url, data=body.encode("utf-8") if body is not None else None,
        method=method, headers=hdrs)
    try:
        with _opener.open(req, timeout=30) as resp:
            resp.read()
            return resp.status
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise RuntimeError("CalDAV 401: пароль приложения не принят для календаря.") from None
        if e.code == 412:
            raise RuntimeError("CalDAV 412: встреча изменилась на сервере (или уже существует) "
                               "— запросите её заново и повторите.") from None
        try:
            detail = e.read()[:300].decode("utf-8", "replace").strip()
        except Exception:
            detail = ""
        raise RuntimeError(f"CalDAV {method}: HTTP {e.code} {e.reason}. {detail}".strip()) from None
    except urllib.error.URLError as e:
        raise RuntimeError(f"CalDAV {method} {url}: {e.reason}") from None
    except (http.client.HTTPException, OSError) as e:
        raise RuntimeError(f"CalDAV {method} {url}: {e!r}") from None


# ---------- текст iCalendar построчно (неизвестные свойства сохраняются как есть) ----------

def _unfold(text: str) -> list[str]:
    return [l for l in re.sub(r"\r?\n[ \t]", "", text).splitlines() if l.strip()]


def _fold(line: str) -> str:
    if len(line.encode("utf-8")) <= 75:
        return line
    out, cur = [], b""
    for ch in line:
        b = ch.encode("utf-8")
        if len(cur) + len(b) > 75:
            out.append(cur.decode("utf-8"))
            cur = b" " + b
        else:
            cur += b
    out.append(cur.decode("utf-8"))
    return "\r\n".join(out)


def _esc(s: str) -> str:
    return (str(s).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
            .replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n"))


def _pname(line: str) -> str:
    return re.split(r"[;:]", line, maxsplit=1)[0].strip().upper()


def _segments(text: str) -> list:
    """Строки ресурса; каждый VEVENT — отдельный список строк."""
    segs: list = []
    cur = None
    for l in _unfold(text):
        u = l.strip().upper()
        if cur is None:
            if u == "BEGIN:VEVENT":
                cur = [l]
            else:
                segs.append(l)
            continue
        cur.append(l)
        if u == "END:VEVENT":
            segs.append(cur)
            cur = None
    if cur is not None:
        segs.append(cur)
    return segs


def _render(segs: list) -> str:
    lines: list[str] = []
    for s in segs:
        lines.extend(s if isinstance(s, list) else [s])
    return "\r\n".join(_fold(l) for l in lines) + "\r\n"


def _comp(block: list[str]) -> dict:
    return parse_ics("\r\n".join(block))[0]


def _top(block: list[str]) -> list[int]:
    """Индексы свойств самого VEVENT (без вложенных VALARM)."""
    idx, depth = [], 0
    for i in range(1, len(block) - 1):
        n = _pname(block[i])
        if n == "BEGIN":
            depth += 1
        elif n == "END":
            depth -= 1
        elif depth == 0:
            idx.append(i)
    return idx


def _bdel(block: list[str], name: str) -> None:
    for i in reversed([i for i in _top(block) if _pname(block[i]) == name]):
        del block[i]


def _bset(block: list[str], name: str, line: str) -> None:
    idx = [i for i in _top(block) if _pname(block[i]) == name]
    if idx:
        block[idx[0]] = line
        for i in reversed(idx[1:]):
            del block[i]
    else:
        block.insert(len(block) - 1, line)


def _stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _touch(block: list[str], bump: bool = True) -> None:
    now = _stamp()
    _bset(block, "DTSTAMP", "DTSTAMP:" + now)
    _bset(block, "LAST-MODIFIED", "LAST-MODIFIED:" + now)
    if bump:
        seq = _get(_comp(block), "SEQUENCE")
        try:
            n = int(seq[1]) + 1 if seq else 1
        except ValueError:
            n = 1
        _bset(block, "SEQUENCE", f"SEQUENCE:{n}")


def _when(s: str):
    """-> (date, True) или (aware datetime, False). Время без смещения — в поясе MAILRU_TZ."""
    s = str(s).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return date.fromisoformat(s), True
    try:
        dt = datetime.fromisoformat(s.replace(" ", "T").replace("Z", "+00:00"))
    except ValueError:
        raise ValueError(f"Нужно время YYYY-MM-DDTHH:MM или дата YYYY-MM-DD: {s!r}") from None
    return (dt if dt.tzinfo else dt.replace(tzinfo=local_tz())), False


def _fmt(name: str, val, is_date: bool, like: dict, vtz: dict, utc: bool = False) -> str:
    """Строка даты того же вида, что у образца like (параметры DTSTART события)."""
    if is_date:
        d = val.date() if isinstance(val, datetime) else val
        return f"{name};VALUE=DATE:{d:%Y%m%d}"
    tzid = like.get("TZID")
    if tzid and not utc:
        return f"{name};TZID={tzid}:{val.astimezone(_tz(tzid, vtz)):%Y%m%dT%H%M%S}"
    return f"{name}:{val.astimezone(timezone.utc):%Y%m%dT%H%M%SZ}"


def _shown(val, is_date: bool, end: bool = False) -> str:
    if is_date:
        d = val.date() if isinstance(val, datetime) else val
        return (d - timedelta(days=1) if end else d).isoformat()
    return val.astimezone(local_tz()).isoformat(timespec="minutes")


def _emails(items) -> list[str]:
    out = []
    for a in items or []:
        a = _mail(str(a))
        if not _MAIL_RE.fullmatch(a):
            raise ValueError(f"Не похоже на адрес почты: {a!r}")
        if a not in out:
            out.append(a)
    return out


def _attendee_line(addr: str) -> str:
    return f"ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:{addr}"


def _line_mail(line: str) -> str:
    return _mail(_split_line(line)[2])


# ---------- поиск ресурса и вхождения ----------

def _has_uid(ics: str, uid: str) -> bool:
    for vcal in parse_ics(ics):
        for comp in vcal["children"]:
            if comp["name"] == "VEVENT":
                u = _get(comp, "UID")
                if u and u[1].strip() == uid:
                    return True
    return False


def _locate(uid: str, when: Optional[str] = None) -> dict:
    """Ресурс встречи по UID -> {url, etag, ics, calendar}."""
    uid = str(uid).strip()
    if not uid:
        raise ValueError("Пустой uid")
    loc = local_tz()
    if when:
        v, is_date = _when(when)
        d = v if is_date else v.astimezone(loc).date()
        ws = datetime.combine(d - timedelta(days=1), dtime(), tzinfo=loc)
        we = ws + timedelta(days=3)
    else:
        ws = datetime.combine(datetime.now(loc).date() - timedelta(days=14), dtime(), tzinfo=loc)
        we = ws + timedelta(days=120)
    by_uid = ('<C:prop-filter name="UID"><C:text-match collation="i;octet">'
              f"{_xml(uid)}</C:text-match></C:prop-filter>")
    errors = []
    for flt in (by_uid, _time_range(ws, we)):
        for cal in calendars():
            try:
                items = _query(cal["url"], flt)
            except RuntimeError as e:
                errors.append(str(e))
                continue
            for it in items:
                if _has_uid(it["ics"], uid):
                    return {**it, "calendar": cal["name"]}
    hint = "" if when else " Укажите occurrence (дату встречи), чтобы искать по дате."
    raise ValueError(f"Встреча uid={uid} не найдена.{hint}"
                     + (" Ошибки: " + "; ".join(errors[:2]) if errors else ""))


def _vtz_of(ics: str) -> dict:
    for vcal in parse_ics(ics):
        if vcal["name"] == "VCALENDAR":
            return _vtimezones(vcal)
    return {}


def _pick(segs: list, uid: str):
    master, overrides = None, []
    for s in segs:
        if not isinstance(s, list):
            continue
        c = _comp(s)
        u = _get(c, "UID")
        if not u or u[1].strip() != uid:
            continue
        if _get(c, "RECURRENCE-ID"):
            overrides.append(s)
        else:
            master = s
    if master is None and not overrides:
        raise ValueError(f"В ресурсе нет встречи uid={uid}")
    return master, overrides


def _occurrence(ics: str, uid: str, occurrence: str) -> dict:
    """Одно вхождение встречи по дате или времени начала."""
    v, is_date = _when(occurrence)
    loc = local_tz()
    if is_date:
        ws = datetime.combine(v, dtime(), tzinfo=loc)
        we = ws + timedelta(days=1)
    else:
        ws, we = v, v + timedelta(minutes=1)
    evs = [e for e in expand(ics, ws, we, USER) if e["uid"] == uid]
    if is_date:
        evs = [e for e in evs if e["start_dt"].date() == v
               or (e["occ_dt"] and e["occ_dt"].date() == v)]
    else:
        evs = [e for e in evs if e["start_dt"] == v or e["occ_dt"] == v]
    if not evs:
        raise ValueError(f"У встречи нет вхождения на {occurrence}")
    if len(evs) > 1:
        raise ValueError(f"На {occurrence} несколько вхождений — укажите время начала: "
                         + ", ".join(e["start_dt"].isoformat(timespec="minutes") for e in evs))
    return evs[0]


def _find_override(overrides: list, occ: datetime, vtz: dict):
    for o in overrides:
        p, v = _get(_comp(o), "RECURRENCE-ID")
        try:
            d, d_is_date = _parse_dt(p, v, vtz)
        except (ValueError, KeyError):
            continue
        if (d_is_date and d == occ.date()) or (not d_is_date and d == occ):
            return o
    return None


def _new_override(master: list[str], ev: dict, vtz: dict) -> list[str]:
    like = _get(_comp(master), "DTSTART")[0]
    top = set(_top(master))
    blk = [l for i, l in enumerate(master)
           if not (i in top and _pname(l) in _RECUR + ("DURATION",))]
    occ, is_date = ev["occ_dt"], ev["all_day"]
    end = occ + (ev["end_dt"] - ev["start_dt"])
    _bset(blk, "DTSTART", _fmt("DTSTART", occ, is_date, like, vtz))
    _bset(blk, "DTEND", _fmt("DTEND", end, is_date, like, vtz))
    _bset(blk, "RECURRENCE-ID", _fmt("RECURRENCE-ID", occ, is_date, like, vtz, utc=True))
    return blk


def _targets(item: dict, segs: list, uid: str, occurrence: Optional[str], vtz: dict):
    """Блоки VEVENT, которые надо менять -> (блоки, вхождение или None)."""
    master, overrides = _pick(segs, uid)
    if not occurrence:
        return ([master] if master else []) + overrides, None
    ev = _occurrence(item["ics"], uid, occurrence)
    if not ev["recurring"]:
        return ([master] if master else overrides), None
    hit = _find_override(overrides, ev["occ_dt"], vtz)
    if hit is None:
        if master is None:
            raise ValueError("Вхождение не найдено в ресурсе")
        hit = _new_override(master, ev, vtz)
        pos = next(i for i, s in enumerate(segs) if s is master)
        segs.insert(pos + 1, hit)
    return [hit], ev


def _put(item: dict, segs: list) -> None:
    hdrs = {"Content-Type": "text/calendar; charset=utf-8"}
    if item.get("etag"):
        hdrs["If-Match"] = item["etag"]
    _http("PUT", item["url"], _render(segs), hdrs)


def _choose_calendar(name: Optional[str]) -> dict:
    cals = calendars()
    if not cals:
        raise RuntimeError("Нет доступных календарей")
    want = (name or DEFAULT_CALENDAR or "").casefold()
    if want:
        hit = [c for c in cals if want in c["name"].casefold()]
        if len(hit) != 1:
            raise ValueError(f"Календарь {name or DEFAULT_CALENDAR!r}: найдено {len(hit)}; "
                             f"есть: {[c['name'] for c in cals]}")
        return hit[0]
    for c in cals:
        if c["name"].casefold() == "основной":
            return c
    return cals[0]


def _block_times(block: list[str], vtz: dict):
    """-> (start, end, is_date, параметры DTSTART); end для дат — исключающий."""
    c = _comp(block)
    p, v = _get(c, "DTSTART")
    start, is_date = _parse_dt(p, v, vtz)
    de, du = _get(c, "DTEND"), _get(c, "DURATION")
    if de:
        end, _ = _parse_dt(de[0], de[1], vtz)
    elif du:
        end = start + _duration(du[1])
    else:
        end = start + (timedelta(days=1) if is_date else timedelta(0))
    return start, end, is_date, p


# ---------- операции ----------

def event(uid: str, occurrence: Optional[str] = None) -> dict:
    """Полные данные одной встречи."""
    item = _locate(uid, occurrence)
    uid = str(uid).strip()
    vtz = _vtz_of(item["ics"])
    master, overrides = _pick(_segments(item["ics"]), uid)
    me = USER.lower()
    if occurrence:
        ev = _occurrence(item["ics"], uid, occurrence)
        f = ev
        start = _shown(ev["start_dt"], ev["all_day"])
        end = _shown(ev["end_dt"], ev["all_day"], end=True)
        all_day = ev["all_day"]
    else:
        blk = master or overrides[0]
        f = _event_fields(_comp(blk), me)
        s, e, all_day, _ = _block_times(blk, vtz)
        start, end = _shown(s, all_day), _shown(e, all_day, end=True)
    out = {"uid": uid, "calendar": item["calendar"], "summary": f["summary"],
           "start": start, "end": end}
    if all_day:
        out["all_day"] = True
    for key in ("location", "link", "organizer", "my_status", "my_comment"):
        if f[key]:
            out[key] = f[key]
    if f["cancelled"]:
        out["cancelled"] = True
    if master is not None:
        c = _comp(master)
        rules = [v for _, v in _all(c, "RRULE")]
        if rules:
            out["rrule"] = rules[0] if len(rules) == 1 else rules
            out["series_start"] = _shown(_block_times(master, vtz)[0], all_day)
            out["excluded_count"] = sum(len(v.split(",")) for _, v in _all(c, "EXDATE"))
    if occurrence and f.get("occ_dt"):
        out["occurrence"] = _shown(f["occ_dt"], all_day)
    out["attendees"] = [{k: a[k] for k in ("email", "status", "comment") if k in a}
                        for a in f["attendees"]]
    out["description"] = f["description"][:8000]
    if overrides and not occurrence:
        rows = []
        for o in overrides:
            oc = _comp(o)
            p, v = _get(oc, "RECURRENCE-ID")
            try:
                rid, rid_is_date = _parse_dt(p, v, vtz)
                s, _, o_is_date, _ = _block_times(o, vtz)
            except (ValueError, KeyError, TypeError):
                continue
            row = {"occurrence": _shown(rid, rid_is_date), "start": _shown(s, o_is_date)}
            if (_get(oc, "STATUS") or ({}, ""))[1].strip().upper() == "CANCELLED":
                row["cancelled"] = True
            rows.append(row)
        out["changed_occurrences"] = sorted(rows, key=lambda r: r["occurrence"])[-20:]
    return out


def create_event(summary: str, start: str, end: Optional[str] = None,
                 calendar: Optional[str] = None, description: str = "", location: str = "",
                 attendees: Optional[list[str]] = None, rrule: str = "",
                 busy: bool = True) -> dict:
    _need_write()
    if not str(summary).strip():
        raise ValueError("Пустое название встречи")
    s, s_is_date = _when(start)
    if end:
        e, e_is_date = _when(end)
        if e_is_date != s_is_date:
            raise ValueError("start и end должны быть оба датами или оба временем")
    else:
        e, e_is_date = (s, True) if s_is_date else (s + timedelta(minutes=30), False)
    if s_is_date:
        if e < s:
            raise ValueError("end раньше start")
        e = e + timedelta(days=1)
    elif e <= s:
        raise ValueError("end должен быть позже start")
    rrule = re.sub(r"^RRULE:", "", str(rrule or "").strip(), flags=re.I)
    if rrule:
        try:
            rrulestr("RRULE:" + _fix_until(rrule, local_tz()), dtstart=datetime(2026, 1, 1))
        except Exception as err:
            raise ValueError(f"Некорректное правило повтора RRULE: {err}") from None
    me = USER.lower()
    people = [a for a in _emails(attendees) if a != me]
    cal = _choose_calendar(calendar)

    like, vt = {}, []
    if not s_is_date:
        z, y = local_tz(), datetime.now().year
        o1 = datetime(y, 1, 1, tzinfo=z).utcoffset()
        o2 = datetime(y, 7, 1, tzinfo=z).utcoffset()
        if o1 == o2:  # пояс без перехода на летнее время — пишем местное время с TZID
            mins = int(o1.total_seconds() // 60)
            off = f"{'+' if mins >= 0 else '-'}{abs(mins) // 60:02d}{abs(mins) % 60:02d}"
            like = {"TZID": TZ_NAME}
            vt = ["BEGIN:VTIMEZONE", f"TZID:{TZ_NAME}", "BEGIN:STANDARD",
                  "DTSTART:19700101T000000", f"TZOFFSETFROM:{off}", f"TZOFFSETTO:{off}",
                  "END:STANDARD", "END:VTIMEZONE"]
    uid, now = str(uuid.uuid4()), _stamp()
    ev = ["BEGIN:VEVENT", f"UID:{uid}", f"DTSTAMP:{now}", f"CREATED:{now}",
          f"LAST-MODIFIED:{now}", "SEQUENCE:0", "STATUS:CONFIRMED",
          "TRANSP:" + ("OPAQUE" if busy else "TRANSPARENT"),
          "SUMMARY:" + _esc(str(summary).strip()),
          _fmt("DTSTART", s, s_is_date, like, {}), _fmt("DTEND", e, s_is_date, like, {})]
    if description:
        ev.append("DESCRIPTION:" + _esc(description))
    if location:
        ev.append("LOCATION:" + _esc(location))
    if rrule:
        ev.append("RRULE:" + rrule)
    ev.append(f"ORGANIZER:mailto:{me}")
    ev += [_attendee_line(a) for a in people]
    ev.append("END:VEVENT")
    segs = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//mailru-mcp//RU",
            "CALSCALE:GREGORIAN", *vt, ev, "END:VCALENDAR"]
    url = cal["url"].rstrip("/") + "/" + uid + ".ics"
    _http("PUT", url, _render(segs),
          {"Content-Type": "text/calendar; charset=utf-8", "If-None-Match": "*"})
    out = {"created": True, "uid": uid, "calendar": cal["name"],
           "summary": str(summary).strip(), "start": _shown(s, s_is_date),
           "end": _shown(e, s_is_date, end=True)}
    if people:
        out["attendees"] = people
    if rrule:
        out["rrule"] = rrule
    return out


def update_event(uid: str, occurrence: Optional[str] = None, summary: Optional[str] = None,
                 start: Optional[str] = None, end: Optional[str] = None,
                 description: Optional[str] = None, location: Optional[str] = None,
                 add_attendees: Optional[list[str]] = None,
                 remove_attendees: Optional[list[str]] = None,
                 rrule: Optional[str] = None, busy: Optional[bool] = None) -> dict:
    _need_write()
    if all(x is None for x in (summary, start, end, description, location, rrule, busy)) \
            and not add_attendees and not remove_attendees:
        raise ValueError("Не указано, что менять")
    if summary is not None and not str(summary).strip():
        raise ValueError("Пустое название встречи")
    item = _locate(uid, occurrence)
    uid = str(uid).strip()
    vtz = _vtz_of(item["ics"])
    segs = _segments(item["ics"])
    blocks, occ = _targets(item, segs, uid, occurrence, vtz)
    if rrule is not None and occ is not None:
        raise ValueError("Правило повтора меняется только у всей серии (без occurrence)")
    me = USER.lower()
    add = [a for a in _emails(add_attendees) if a != me]
    drop = set(_emails(remove_attendees))
    if rrule is not None:
        rrule = re.sub(r"^RRULE:", "", str(rrule).strip(), flags=re.I)
        if rrule:
            try:
                rrulestr("RRULE:" + _fix_until(rrule, local_tz()), dtstart=datetime(2026, 1, 1))
            except Exception as err:
                raise ValueError(f"Некорректное правило повтора RRULE: {err}") from None
    changed, shown = [], {}
    for blk in blocks:
        is_master = not _get(_comp(blk), "RECURRENCE-ID")
        if summary is not None:
            _bset(blk, "SUMMARY", "SUMMARY:" + _esc(str(summary).strip()))
        for name, val in (("DESCRIPTION", description), ("LOCATION", location)):
            if val is not None:
                if val:
                    _bset(blk, name, f"{name}:" + _esc(val))
                else:
                    _bdel(blk, name)
        if busy is not None:
            _bset(blk, "TRANSP", "TRANSP:" + ("OPAQUE" if busy else "TRANSPARENT"))
        if drop:
            for i in reversed([i for i in _top(blk) if _pname(blk[i]) == "ATTENDEE"
                               and _line_mail(blk[i]) in drop]):
                del blk[i]
        if add:
            have = {_line_mail(blk[i]) for i in _top(blk) if _pname(blk[i]) == "ATTENDEE"}
            for a in add:
                if a not in have:
                    blk.insert(len(blk) - 1, _attendee_line(a))
        # время меняем у самой встречи: у мастера серии либо у выбранного вхождения
        if (start is not None or end is not None) and (is_master or occ is not None
                                                       or len(blocks) == 1):
            cur_s, cur_e, is_date, like = _block_times(blk, vtz)
            new_s, new_e = cur_s, cur_e
            if start is not None:
                v, v_is_date = _when(start)
                if v_is_date != is_date:
                    raise ValueError("Нельзя менять вид встречи (весь день / по времени) — "
                                     "создайте новую")
                new_s = v
                new_e = v + (cur_e - cur_s)
            if end is not None:
                v, v_is_date = _when(end)
                if v_is_date != is_date:
                    raise ValueError("start и end должны быть одного вида")
                new_e = v + timedelta(days=1) if is_date else v
            if new_e <= new_s:
                raise ValueError("Конец встречи должен быть позже начала")
            _bset(blk, "DTSTART", _fmt("DTSTART", new_s, is_date, like, vtz))
            _bdel(blk, "DURATION")
            _bset(blk, "DTEND", _fmt("DTEND", new_e, is_date, like, vtz))
            shown = {"start": _shown(new_s, is_date), "end": _shown(new_e, is_date, end=True)}
        if rrule is not None and is_master:
            if rrule:
                _bset(blk, "RRULE", "RRULE:" + rrule)
            else:
                for n in _RECUR:
                    _bdel(blk, n)
        _touch(blk)
        changed.append(blk)
    _put(item, segs)
    f = _event_fields(_comp(changed[0]), me)
    out = {"updated": True, "uid": uid, "calendar": item["calendar"], "summary": f["summary"],
           "scope": "occurrence" if occ is not None else "event", **shown}
    if occ is not None:
        out["occurrence"] = _shown(occ["occ_dt"], occ["all_day"])
    out["attendees_count"] = len(f["attendees"])
    return out


def delete_event(uid: str, occurrence: Optional[str] = None) -> dict:
    _need_write()
    item = _locate(uid, occurrence)
    uid = str(uid).strip()
    vtz = _vtz_of(item["ics"])
    segs = _segments(item["ics"])
    master, overrides = _pick(segs, uid)
    title = _event_fields(_comp(master or overrides[0]), USER.lower())["summary"]
    ev = _occurrence(item["ics"], uid, occurrence) if occurrence else None
    if ev is None or not ev["recurring"] or (master is None and len(overrides) <= 1):
        hdrs = {"If-Match": item["etag"]} if item.get("etag") else {}
        _http("DELETE", item["url"], None, hdrs)
        return {"deleted": True, "uid": uid, "summary": title, "scope": "event",
                "calendar": item["calendar"]}
    hit = _find_override(overrides, ev["occ_dt"], vtz)
    if hit is not None:
        segs[:] = [s for s in segs if s is not hit]
    if master is not None:
        if ev["all_day"]:
            line = f"EXDATE;VALUE=DATE:{ev['occ_dt'].date():%Y%m%d}"
        else:
            line = f"EXDATE:{ev['occ_dt'].astimezone(timezone.utc):%Y%m%dT%H%M%SZ}"
        master.insert(len(master) - 1, line)
        _touch(master)
    _put(item, segs)
    return {"deleted": True, "uid": uid, "summary": title, "scope": "occurrence",
            "occurrence": _shown(ev["occ_dt"], ev["all_day"]), "calendar": item["calendar"]}


_PARTSTAT = {"accepted": "ACCEPTED", "accept": "ACCEPTED", "yes": "ACCEPTED",
             "declined": "DECLINED", "decline": "DECLINED", "no": "DECLINED",
             "tentative": "TENTATIVE", "maybe": "TENTATIVE"}


def respond(uid: str, status: str, occurrence: Optional[str] = None,
            comment: Optional[str] = None) -> dict:
    _need_write()
    note = re.sub(r"\s+", " ", str(comment or "")).replace('"', "'").strip()[:300]
    part = _PARTSTAT.get(str(status).strip().lower())
    if not part:
        raise ValueError("status: accepted, declined или tentative")
    item = _locate(uid, occurrence)
    uid = str(uid).strip()
    vtz = _vtz_of(item["ics"])
    segs = _segments(item["ics"])
    blocks, occ = _targets(item, segs, uid, occurrence, vtz)
    me, n = USER.lower(), 0
    for blk in blocks:
        for i in _top(blk):
            if _pname(blk[i]) != "ATTENDEE" or _line_mail(blk[i]) != me:
                continue
            m = re.search(r":mailto:", blk[i], flags=re.I)
            head, rest = blk[i][:m.start()], blk[i][m.start():]
            if re.search(r"PARTSTAT=", head, flags=re.I):
                head = re.sub(r"PARTSTAT=[^;:]*", "PARTSTAT=" + part, head, flags=re.I)
            else:
                head += ";PARTSTAT=" + part
            head = re.sub(r';X-RESPONSE-COMMENT=(?:"[^"]*"|[^;:]*)', "", head, flags=re.I)
            if note:
                head += f';X-RESPONSE-COMMENT="{note}"'
            blk[i] = head + rest
            n += 1
        if note:
            _bset(blk, "COMMENT", "COMMENT:" + _esc(note))
        else:
            _bdel(blk, "COMMENT")
        _touch(blk, bump=False)
    if not n:
        raise ValueError("Вас нет среди участников этой встречи (вероятно, вы организатор) — "
                         "отвечать не на что")
    _put(item, segs)
    title = _event_fields(_comp(blocks[0]), me)["summary"]
    out = {"responded": True, "uid": uid, "summary": title, "status": part.lower(),
           "scope": "occurrence" if occ is not None else "event",
           "calendar": item["calendar"]}
    if occ is not None:
        out["occurrence"] = _shown(occ["occ_dt"], occ["all_day"])
    if note:
        out["comment"] = note
    return out
