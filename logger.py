import time
import requests
import sqlite3
import os
from pathlib import Path
from datetime import datetime, timezone
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

# ================= CONFIG =================
API_KEY = os.environ["STEAM_API_KEY"]
STEAM_ID = os.environ["STEAM_ID"]
POLL_INTERVAL = 60  # 5 minutes

# ================= PATHS (Railway-safe) =================
DATA_DIR = Path("/app/data")
DATA_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = DATA_DIR / "sessions.db"
EXCEL_FILE = DATA_DIR / "steam_sessions.xlsx"

# ================= DATABASE =================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sessions (
            steamid TEXT,
            name TEXT,
            start_ts TEXT,
            end_ts TEXT,
            duration REAL,
            exported INTEGER DEFAULT 0
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS active (
            steamid TEXT PRIMARY KEY,
            name TEXT,
            start_ts TEXT
        )
    """)
    conn.commit()
    conn.close()

# ================= STEAM API =================
def get_friends():
    url = "https://api.steampowered.com/ISteamUser/GetFriendList/v1/"
    r = requests.get(url, params={"key": API_KEY, "steamid": STEAM_ID})
    r.raise_for_status()
    return [f["steamid"] for f in r.json()["friendslist"]["friends"]]

def get_player_summaries(ids):
    url = "https://api.steampowered.com/ISteamUser/GetPlayerSummaries/v2/"
    r = requests.get(url, params={
        "key": API_KEY,
        "steamids": ",".join(ids)
    })
    r.raise_for_status()
    return r.json()["response"]["players"]

# ================= EXCEL =================
def init_excel():
    wb = Workbook()

    ws = wb.active
    ws.title = "Sessions"
    ws.append([
        "Name",
        "SteamID",
        "Session Start (UTC)",
        "Session End (UTC)",
        "Duration (minutes)"
    ])

    wb.create_sheet("Summary")
    wb.create_sheet("DailyStats")
    wb.create_sheet("HourHeatmap")
    wb.create_sheet("Charts")

    wb.save(EXCEL_FILE)

def ensure_excel():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not EXCEL_FILE.exists():
        init_excel()

def append_session(row):
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Sessions"]
    ws.append(row)
    wb.save(EXCEL_FILE)

# ================= ANALYTICS =================
def rebuild_excel():
    wb = load_workbook(EXCEL_FILE)
    sessions = wb["Sessions"]

    summary = wb["Summary"]
    daily = wb["DailyStats"]
    heat = wb["HourHeatmap"]
    charts = wb["Charts"]

    summary.delete_rows(1, summary.max_row)
    daily.delete_rows(1, daily.max_row)
    heat.delete_rows(1, heat.max_row)
    charts.delete_rows(1, charts.max_row)

    summary.append(["Name", "Total Minutes", "Sessions", "Avg Session (min)"])
    daily.append(["Date", "Total Minutes"])
    heat.append(["Name"] + list(range(24)))

    totals = defaultdict(float)
    counts = defaultdict(int)
    daily_totals = defaultdict(float)
    hour_map = defaultdict(lambda: [0] * 24)

    for row in sessions.iter_rows(min_row=2, values_only=True):
        name, sid, start, end, dur = row
        totals[name] += dur
        counts[name] += 1

        start_dt = datetime.fromisoformat(start)
        daily_totals[start_dt.date()] += dur
        hour_map[name][start_dt.hour] += 1

    for name in totals:
        summary.append([
            name,
            round(totals[name], 1),
            counts[name],
            round(totals[name] / counts[name], 1)
        ])

    for d in sorted(daily_totals):
        daily.append([str(d), round(daily_totals[d], 1)])

    for name, hours in hour_map.items():
        heat.append([name] + hours)

    # ----- Charts -----
    charts.append(["Name", "Total Minutes"])
    for r in summary.iter_rows(min_row=2, values_only=True):
        charts.append([r[0], r[1]])

    pie = PieChart()
    labels = Reference(charts, min_col=1, min_row=2, max_row=charts.max_row)
    data = Reference(charts, min_col=2, min_row=1, max_row=charts.max_row)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Total Playtime Share"
    charts.add_chart(pie, "E2")

    line = LineChart()
    data = Reference(daily, min_col=2, min_row=1, max_row=daily.max_row)
    cats = Reference(daily, min_col=1, min_row=2, max_row=daily.max_row)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.title = "Daily Total Minutes"
    charts.add_chart(line, "E20")

    wb.save(EXCEL_FILE)

# ================= MAIN LOOP =================
def main():
    ensure_excel()
    init_db()

    while True:
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            friend_ids = get_friends()
            players = get_player_summaries(friend_ids)

            now = datetime.now(timezone.utc).isoformat()

            online_now = {}

            for p in players:
                if p.get("personastate", 0) > 0:
                    online_now[p["steamid"]] = p["personaname"]

            # Handle logins
            for sid, name in online_now.items():
                cur.execute("SELECT 1 FROM active WHERE steamid=?", (sid,))
                if not cur.fetchone():
                    cur.execute(
                        "INSERT INTO active VALUES (?,?,?)",
                        (sid, name, now)
                    )

            # Handle logouts
            cur.execute("SELECT steamid, name, start_ts FROM active")
            for sid, name, start in cur.fetchall():
                if sid not in online_now:
                    start_dt = datetime.fromisoformat(start)
                    end_dt = datetime.fromisoformat(now)
                    duration = (end_dt - start_dt).total_seconds() / 60

                    cur.execute("""
                        INSERT INTO sessions
                        VALUES (?,?,?,?,?,0)
                    """, (sid, name, start, now, duration))

                    cur.execute("DELETE FROM active WHERE steamid=?", (sid,))

            # Export sessions to Excel
            cur.execute("""
                SELECT rowid, steamid, name, start_ts, end_ts, duration
                FROM sessions WHERE exported=0
            """)

            rows = cur.fetchall()
            for r in rows:
                append_session([
                    r[2], r[1], r[3], r[4], round(r[5], 2)
                ])
                cur.execute(
                    "UPDATE sessions SET exported=1 WHERE rowid=?",
                    (r[0],)
                )

            if rows:
                rebuild_excel()

            conn.commit()
            conn.close()

        except Exception as e:
            print("Error:", e)

        time.sleep(POLL_INTERVAL)

# ================= ENTRY =================
if __name__ == "__main__":
    print("Steam daemon running → Railway + Excel analytics")
    main()
