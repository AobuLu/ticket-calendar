import csv
import io
import re
from datetime import datetime, timedelta
from pathlib import Path

import requests
from icalendar import Calendar, Event, Alarm
from pytz import timezone

SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRM6THpS8w6h0b91xbauEz-sp1dl3uhReR5AJpocWh_CuYQcKVm6DNedGCaXJxJTAda5IEXYdfzAhyd/pub?output=csv"

SHOWS = [
    "幽灵",
    "怪物",
    "奥尔菲斯",
    "末日迷途",
    "嗜血博士",
    "阴天",
    "亡灵之旅冥犬与少年",
    "时光代理人",
    "辛吉路的画材店",
    "秘密花园",
    "新贵",
    "星期三",
]

ACCOUNT_SHOWS = {
    "延幕文化": ["幽灵", "怪物", "奥尔菲斯"],
    "涵金文化": ["末日迷途", "嗜血博士", "阴天"],
    "五十五文化": ["亡灵之旅冥犬与少年"],
    "接着奏乐接着舞Musicals": ["时光代理人"],
    "辛吉路的画材店成都站": ["辛吉路的画材店"],
    "Mioji米欧吉": ["秘密花园"],
    "小鹿追马DeerDrama": ["新贵"],
    "一台好戏Focustage": ["星期三"],
}

MIN_ALLOWED = datetime(2026, 3, 1, 0, 0)
TZ = timezone("Asia/Shanghai")

# 同行日期+时间
YMDHM = re.compile(
    r'(?P<year>20\d{2})年\s*(?P<month>\d{1,2})月\s*(?P<day>\d{1,2})[日号]\s*'
    r'(?:[（(][^)）]+[)）])?\s*(?P<hour>\d{1,2})[:：](?P<minute>\d{2})'
)
MDHM = re.compile(
    r'(?P<month>\d{1,2})月\s*(?P<day>\d{1,2})[日号]\s*'
    r'(?:[（(][^)）]+[)）])?\s*(?P<hour>\d{1,2})[:：](?P<minute>\d{2})'
)

# 单独日期 / 单独时间
DATE_ONLY = re.compile(
    r'^(?:(?P<year>20\d{2})年\s*)?(?P<month>\d{1,2})月\s*(?P<day>\d{1,2})[日号]\s*(?:[（(][^)）]+[)）])?$'
)
TIME_ONLY = re.compile(
    r'^(?P<hour>\d{1,2})[:：](?P<minute>\d{2})\s*(?P<label>.+)$'
)

# 兑换时间
EXCHANGE = re.compile(
    r'(?:兑换时间|资格兑换时间)\s*[:：]?\s*'
    r'(?:(?P<year>20\d{2})年\s*)?'
    r'(?P<month>\d{1,2})月\s*(?P<day>\d{1,2})[日号]?\s*'
    r'(?P<hour>\d{1,2})[:：](?P<minute>\d{2})'
)


def load_rows():
    r = requests.get(SHEET_URL, timeout=30)
    r.raise_for_status()
    text = r.content.decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(text)))


def get_any(row, *names):
    for name in names:
        val = row.get(name)
        if val is not None and str(val).strip():
            return str(val).strip()
    return ""


def infer_show(text, account_name):
    hits = [s for s in SHOWS if s in text]
    if len(hits) == 1:
        return hits[0]

    mapped = ACCOUNT_SHOWS.get(account_name, [])
    valid = [s for s in mapped if s in SHOWS]
    if len(valid) == 1:
        return valid[0]

    return None


def dt(y, m, d, h, mi):
    return datetime(int(y), int(m), int(d), int(h), int(mi))


def clean_label(label):
    label = re.sub(r'^[\s\-—–·•]+', '', label.strip())
    label = re.sub(r'^[📍📌⏰🌟✨🎈🐕🔔]+\s*', '', label)
    return label.strip("()（） ")


def is_valid_label(label):
    label = label.strip()

    # 太短
    if len(label) <= 1:
        return False

    # 纯时间
    if re.fullmatch(r'\d{1,2}[:：]\d{2}', label):
        return False

    # 像“3月6日23:59”“4月11日23:59”
    if re.fullmatch(r'\d{1,2}月\d{1,2}日\d{1,2}[:：]\d{2}', label):
        return False

    # 明显噪音
    if label in {"起", "开票时间", "开票", "更多", "详情"}:
        return False

    return True


def build_uid(show_name, sale_type_raw, sale_time):
    safe = re.sub(r'\s+', '', sale_type_raw)
    return f"{show_name}_{sale_time.strftime('%Y%m%dT%H%M')}_{safe}"


def extract_events(text, published_at, show_name):
    year_hint = 2026
    m = re.search(r'(20\d{2})', published_at)
    if m:
        year_hint = int(m.group(1))

    out = []
    pending_date = None

    # 先抓兑换时间
    for m in EXCHANGE.finditer(text):
        y = m.group("year") or year_hint
        when = dt(
            y,
            m.group("month"),
            m.group("day"),
            m.group("hour"),
            m.group("minute"),
        )
        out.append((when, "优先购兑换"))

    lines = [x.strip() for x in text.splitlines() if x.strip()]

    for line in lines:
        # 1. 同一行：完整年月日时分
        m = YMDHM.search(line)
        if m:
            when = dt(
                m.group("year"),
                m.group("month"),
                m.group("day"),
                m.group("hour"),
                m.group("minute"),
            )
            label = clean_label(line[m.end():])
            if label and is_valid_label(label):
                out.append((when, label))
            pending_date = None
            continue

        # 2. 同一行：月日时分
        m = MDHM.search(line)
        if m:
            when = dt(
                year_hint,
                m.group("month"),
                m.group("day"),
                m.group("hour"),
                m.group("minute"),
            )
            label = clean_label(line[m.end():])
            if label and is_valid_label(label):
                out.append((when, label))
            pending_date = None
            continue

        # 3. 单独日期行
        m = DATE_ONLY.search(line)
        if m:
            y = m.group("year") or year_hint
            pending_date = (int(y), int(m.group("month")), int(m.group("day")))
            continue

        # 4. 单独时间行，接上一行日期
        m = TIME_ONLY.search(line)
        if m and pending_date:
            y, month, day = pending_date
            when = dt(y, month, day, m.group("hour"), m.group("minute"))
            label = clean_label(m.group("label"))
            if label and is_valid_label(label):
                out.append((when, label))
            pending_date = None
            continue

    # 去重 + 时间过滤
    seen = set()
    keep = []
    for when, label in out:
        if when < MIN_ALLOWED:
            continue
        key = (show_name, when.isoformat(timespec="minutes"), label)
        if key not in seen:
            seen.add(key)
            keep.append((when, label))

    return keep


def main():
    rows = load_rows()

    cal = Calendar()
    cal.add("prodid", "-//Ticket Calendar//CN//")
    cal.add("version", "2.0")

    seen_uid = set()

    for row in rows:
        account_name = get_any(row, "账号")
        weibo_url = get_any(row, "微博链接")
        body = get_any(row, "自动抓取正文", "正文备用", "正文备用（可选）")
        published_at = get_any(row, "发布时间", "发布时间（可选）", "时间戳", "Timestamp")

        if not account_name or not weibo_url or not body:
            continue

        show_name = infer_show(body, account_name)
        if not show_name:
            continue

        events = extract_events(body, published_at, show_name)

        for when, sale_type_raw in events:
            uid = build_uid(show_name, sale_type_raw, when)
            if uid in seen_uid:
                continue
            seen_uid.add(uid)

            start = TZ.localize(when)
            end = TZ.localize(when + timedelta(minutes=5))

            event = Event()
            event.add("uid", uid)
            event.add("summary", f"{show_name}｜{sale_type_raw}")
            event.add("dtstart", start)
            event.add("dtend", end)
            event.add("description", f"来源账号：{account_name}\n微博：{weibo_url}")

            alarm = Alarm()
            alarm.add("action", "DISPLAY")
            alarm.add("description", f"提醒：{show_name}｜{sale_type_raw}")
            alarm.add("trigger", timedelta(minutes=-10))
            event.add_component(alarm)

            cal.add_component(event)

    Path("docs").mkdir(exist_ok=True)
    with open("docs/ticket_calendar.ics", "wb") as f:
        f.write(cal.to_ical())


if __name__ == "__main__":
    main()
