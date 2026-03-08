import csv
import requests
from datetime import datetime
from icalendar import Calendar, Event
from pathlib import Path

SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRM6THpS8w6h0b91xbauEz-sp1dl3uhReR5AJpocWh_CuYQcKVm6DNedGCaXJxJTAda5IEXYdfzAhyd/pub?output=csv"

def load_rows():
    r = requests.get(SHEET_URL)
    text = r.content.decode("utf-8-sig")
    return list(csv.DictReader(text.splitlines()))

def parse_time(text):
    try:
        return datetime.strptime(text, "%Y-%m-%d %H:%M")
    except:
        return None

def build_calendar(rows):

    cal = Calendar()
    cal.add("prodid", "-//Ticket Calendar//")
    cal.add("version", "2.0")

    for r in rows:

        show = r.get("剧名","")
        time = r.get("开票时间","")
        platform = r.get("平台","")

        if not show or not time:
            continue

        dt = parse_time(time)
        if not dt:
            continue

        event = Event()
        event.add("summary", f"{show}｜{platform}开票")
        event.add("dtstart", dt)
        event.add("dtend", dt)

        cal.add_component(event)

    Path("docs").mkdir(exist_ok=True)

    with open("docs/ticket_calendar.ics","wb") as f:
        f.write(cal.to_ical())

rows = load_rows()
build_calendar(rows)
