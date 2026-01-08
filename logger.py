import time
import requests
import sqlite3
import os
from datetime import datetime, timezone
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# ================= CONFIG =================
API_KEY = os.environ["STEAM_API_KEY"]
STEAM_ID = os.environ["STEAM_ID"]
POLL_INTERVAL = 60  # seconds

DB = "sessions.db"
EXCEL_FILE = "steam_sessions.xlsx"

STATUS_ONLINE = {1, 2, 3, 4, 5, 6}
# =========================================

# ---------- Time ----------
def utc_now():
    return datetime.now(timezone.utc)

def utc_iso():
    return utc_now().isoformat()

# ---------- Excel Init ----------
def init_excel():
    wb = Workbook()

    ws = wb.active
    ws.title = "Sessions"
    ws.append([
        "Name", "SteamID",
        "Session Start (UTC)", "Session End (UTC)",
        "Duration (minutes)"
    ])

    wb.create_sheet("Summary")
    wb.create_sheet("DailyStats")
    wb.create_sheet("HourHeatmap")
    wb.create_sheet("Charts")

    wb.save(EXCEL_FILE)

def ensure_excel():
    if not os.path.exists(EXCEL_FILE):
        init_excel()

def append_session(row):
    wb = load_workbook(EXCEL_FILE)
    wb["Sessions"].append(row)
    wb.save(EXCEL_FILE)

# ---------- Rebuild Analytics ----------
def rebuild_excel():
    wb = load_workbook(EXCEL_FILE)

    ws_s = wb["Sessions"]
    ws_sum = wb["Summary"]
    ws_daily = wb["DailyStats"]
    ws_heat = wb["HourHeatmap"]
    ws_charts = wb["Charts"]

    ws_sum.delete_rows(1, ws_sum.max_row)
    ws_daily.delete_rows(1, ws_daily.max_row)
    ws_heat.delete_rows(1, ws_heat.max_row)
    ws_charts._charts.clear()

    # ---------------- Summary ----------------
    ws_sum.append(["Name", "Total Sessions", "Total Minutes", "Avg Session"])

    sessions_by_name = defaultdict(list)

    for r in ws_s.iter_rows(min_row=2, values_only=True):
        name, _, start, end, dur = r
        if dur:
            sessions_by_name[name].append((start, end, dur))

    for name, rows in sessions_by_name.items():
        durs = [r[2] for r in rows]
        ws_sum.append([
            name,
            len(durs),
            round(sum(durs), 2),
            round(sum(durs) / len(durs), 2)
        ])

    # ---------------- Pie Chart ----------------
    pie = PieChart()
    pie.title = "Total Playtime Share"

    pie_data = Reference(ws_sum, min_col=3, min_row=1, max_row=ws_sum.max_row)
    pie_labels = Reference(ws_sum, min_col=1, min_row=2, max_row=ws_sum.max_row)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_labels)
    ws_charts.add_chart(pie, "A1")

    # ---------------- Daily + Rolling 7-day ----------------
    ws_daily.append(["Date", "Total Minutes", "7-Day Rolling Avg"])

    daily_totals = defaultdict(float)

    for rows in sessions_by_name.values():
        for start, _, dur in rows:
            d = datetime.fromisoformat(start).date()
            daily_totals[d] += dur

    dates = sorted(daily_totals.keys())
    values = [daily_totals[d] for d in dates]

    for i, d in enumerate(dates):
        window = values[max(0, i-6):i+1]
        ws_daily.append([
            d.isoformat(),
            round(values[i], 2),
            round(sum(window) / len(window), 2)
        ])

    line = LineChart()
    line.title = "7-Day Rolling Average (Minutes)"
    line.y_axis.title = "Minutes"
    line.x_axis.title = "Date"

    data_ref = Reference(ws_daily, min_col=3, min_row=1, max_row=ws_daily.max_row)
    cats = Reference(ws_daily, min_col=1, min_row=2, max_row=ws_daily.max_row)
    line.add_data(data_ref, titles_from_data=True)
    line.set_categories(cats)
    ws_charts.add_chart(line, "A20")

    # ---------------- Hour-of-Day Heatmap ----------------
    ws_heat.append(["Hour"] + list(range(24)))

    hour_counts = defaultdict(lambda: [0]*24)

    for name, rows in sessions_by_name.items():
        for start, end, _ in rows:
            s = datetime.fromisoformat(start)
            e = datetime.fromisoformat(end)
            h = s.hour
            hour_counts[name][h] += 1

    for name, hours in hour_counts.items():
        ws_heat.append([name] + hours)

    # ---------------- Weekday vs Weekend ----------------
    weekday_minutes = 0
    weekend_minutes = 0

    for rows in sessions_by_name.values():
        for start, _, dur in rows:
            d = datetime.fromisoformat(start)
            if d.weekday() < 5:
                weekday_minutes += dur
            else:
                weekend_minutes += dur

    ws_tmp = wb.create_sheet("TempWeek")
    ws_tmp.append(["Type", "Minutes"])
    ws_tmp.append(["Weekday", weekday_minutes])
    ws_tmp.append(["Weekend", weekend_minutes])

    bar = BarChart()
    bar.title = "Weekday vs Weekend Playtime"
    bar.y_axis.title = "Minutes"

    bar_data = Reference(ws_tmp, min_col=2, min_row=1, max_row=3)
    bar_cats = Reference(ws_tmp, min_col=1, min_row=2, max_row=3)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(bar_cats)
    ws_charts.add_chart(bar, "A40")

    wb.remove(ws_tmp)
    wb.save(EXCEL_FILE)

# ---------- Database ----------
conn = sqlite3.connect(DB)
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS sessions (
    steamid TEXT, name TEXT,
    start_time TEXT, end_time TEXT,
    duration_minutes REAL,
    exported INTEGER DEFAULT 0
)
""")

cur.execute("""
CREATE TABLE IF NOT EXISTS last_state (
    steamid TEXT PRIMARY KEY,
    state INTEGER,
    last_seen TEXT
)
""")

conn.commit()

# ---------- Steam ----------
def get_profiles():
    ids = ",".join(
        f["steamid"]
        for f in requests.get(
            f"https://api.steampowered.com/ISteamUser/GetFriendList/v1/"
            f"?key={API_KEY}&steamid={STEAM_ID}&relationship=friend"
        ).json()["friendslist"]["friends"]
    )
    return requests.get(
        f"https://api.steampowered.com/ISteamUser/GetPlayerSummaries/v2/"
        f"?key={API_KEY}&steamids={ids}"
    ).json()["response"]["players"]

# ---------- Run ----------
ensure_excel()
print("Steam daemon running → Advanced Excel Analytics")

while True:
    try:
        for p in get_profiles():
            sid, name, state = p["steamid"], p["personaname"], p["personastate"]

            cur.execute("SELECT state FROM last_state WHERE steamid=?", (sid,))
            row = cur.fetchone()

            if row is None:
                cur.execute("INSERT INTO last_state VALUES (?, ?, ?)", (sid, state, utc_iso()))
                continue

            prev = row[0]

            if prev == 0 and state in STATUS_ONLINE:
                cur.execute(
                    "INSERT INTO sessions (steamid, name, start_time) VALUES (?, ?, ?)",
                    (sid, name, utc_iso())
                )

            if prev in STATUS_ONLINE and state == 0:
                cur.execute("""
                    SELECT rowid, start_time FROM sessions
                    WHERE steamid=? AND end_time IS NULL
                    ORDER BY start_time DESC LIMIT 1
                """, (sid,))
                s = cur.fetchone()
                if s:
                    rid, start = s
                    end = utc_now()
                    dur = (end - datetime.fromisoformat(start)).total_seconds() / 60
                    cur.execute("""
                        UPDATE sessions SET end_time=?, duration_minutes=? WHERE rowid=?
                    """, (end.isoformat(), dur, rid))

            cur.execute(
                "UPDATE last_state SET state=?, last_seen=? WHERE steamid=?",
                (state, utc_iso(), sid)
            )

        conn.commit()

        cur.execute("""
            SELECT rowid, name, steamid, start_time, end_time, duration_minutes
            FROM sessions WHERE end_time IS NOT NULL AND exported=0
        """)

        rows = cur.fetchall()
        for r in rows:
            append_session([r[1], r[2], r[3], r[4], round(r[5], 2)])
            cur.execute("UPDATE sessions SET exported=1 WHERE rowid=?", (r[0],))

        if rows:
            rebuild_excel()

        conn.commit()

    except Exception as e:
        print("Error:", e)

    time.sleep(POLL_INTERVAL)
