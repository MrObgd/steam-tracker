import time
import requests
import sqlite3
import os
from pathlib import Path
from datetime import datetime, timezone
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, PieChart, Reference

# ================= CONFIG =================
API_KEY = os.environ["STEAM_API_KEY"]
STEAM_ID = os.environ["STEAM_ID"]
DISCORD_WEBHOOK = os.environ.get("DISCORD_WEBHOOK_URL")

POLL_INTERVAL = 60  # seconds

ACTIVE_STATES = {1, 2, 5, 6}
IDLE_STATES = {3, 4}

# ================= PATHS =================
DATA_DIR = Path("/app/data")
DATA_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = DATA_DIR / "sessions.db"
EXCEL_FILE = DATA_DIR / "steam_sessions.xlsx"
LAST_UPLOAD_FILE = DATA_DIR / "last_discord_upload.txt"

# ================= DATABASE =================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS segments (
            steamid TEXT,
            name TEXT,
            start_ts TEXT,
            end_ts TEXT,
            duration REAL,
            segment_type TEXT,
            exported INTEGER DEFAULT 0
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS active (
            steamid TEXT PRIMARY KEY,
            name TEXT,
            start_ts TEXT,
            segment_type TEXT
        )
    """)

    conn.commit()
    conn.close()

# ================= STEAM API =================
def get_tracked_ids():
    r = requests.get(
        "https://api.steampowered.com/ISteamUser/GetFriendList/v1/",
        params={"key": API_KEY, "steamid": STEAM_ID},
        timeout=15
    )
    r.raise_for_status()

    ids = {f["steamid"] for f in r.json()["friendslist"]["friends"]}
    ids.add(STEAM_ID)
    return list(ids)

def get_player_summaries(ids):
    r = requests.get(
        "https://api.steampowered.com/ISteamUser/GetPlayerSummaries/v2/",
        params={"key": API_KEY, "steamids": ",".join(ids)},
        timeout=15
    )
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
        "Active Duration (minutes)"
    ])

    wb.create_sheet("Summary")
    wb.create_sheet("DailyStats")
    wb.create_sheet("HourHeatmap")
    wb.create_sheet("Charts")

    wb.save(EXCEL_FILE)

def ensure_excel():
    if not EXCEL_FILE.exists():
        init_excel()

def append_session(row):
    wb = load_workbook(EXCEL_FILE)
    wb["Sessions"].append(row)
    wb.save(EXCEL_FILE)

# ================= ANALYTICS =================
def rebuild_excel():
    wb = load_workbook(EXCEL_FILE)
    sessions = wb["Sessions"]

    summary = wb["Summary"]
    daily = wb["DailyStats"]
    heat = wb["HourHeatmap"]
    charts = wb["Charts"]

    for ws in (summary, daily, heat, charts):
        ws.delete_rows(1, ws.max_row)

    summary.append(["Name", "Active Minutes", "Sessions", "Avg Session"])
    daily.append(["Date", "Active Minutes"])
    heat.append(["Name"] + list(range(24)))

    totals = defaultdict(float)
    counts = defaultdict(int)
    daily_totals = defaultdict(float)
    hour_map = defaultdict(lambda: [0] * 24)

    for name, sid, start, end, dur in sessions.iter_rows(min_row=2, values_only=True):
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

    charts.append(["Name", "Active Minutes"])
    for r in summary.iter_rows(min_row=2, values_only=True):
        charts.append([r[0], r[1]])

    pie = PieChart()
    pie.add_data(
        Reference(charts, min_col=2, min_row=1, max_row=charts.max_row),
        titles_from_data=True
    )
    pie.set_categories(
        Reference(charts, min_col=1, min_row=2, max_row=charts.max_row)
    )
    pie.title = "Active Playtime Share"
    charts.add_chart(pie, "E2")

    line = LineChart()
    line.add_data(
        Reference(daily, min_col=2, min_row=1, max_row=daily.max_row),
        titles_from_data=True
    )
    line.set_categories(
        Reference(daily, min_col=1, min_row=2, max_row=daily.max_row)
    )
    line.title = "Daily Active Minutes"
    charts.add_chart(line, "E20")

    wb.save(EXCEL_FILE)

# ================= DISCORD =================
def should_upload_today():
    today = datetime.now(timezone.utc).date()
    if LAST_UPLOAD_FILE.exists():
        if datetime.fromisoformat(LAST_UPLOAD_FILE.read_text()).date() == today:
            return False
    LAST_UPLOAD_FILE.write_text(datetime.now(timezone.utc).isoformat())
    return True

def upload_to_discord():
    if not DISCORD_WEBHOOK or not EXCEL_FILE.exists():
        return

    with open(EXCEL_FILE, "rb") as f:
        requests.post(
            DISCORD_WEBHOOK,
            data={"content": "📊 Daily Steam activity (idle excluded)"},
            files={"file": ("steam_sessions.xlsx", f)},
            timeout=30
        )

# ================= MAIN LOOP =================
def main():
    ensure_excel()
    init_db()

    while True:
        try:
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()

            now = datetime.now(timezone.utc).isoformat()
            players = get_player_summaries(get_tracked_ids())

            current = {}
            for p in players:
                state = p.get("personastate", 0)
                if state in ACTIVE_STATES:
                    current[p["steamid"]] = (p["personaname"], "active")
                elif state in IDLE_STATES:
                    current[p["steamid"]] = (p["personaname"], "idle")

            # Start segments
            for sid, (name, seg_type) in current.items():
                cur.execute("SELECT segment_type FROM active WHERE steamid=?", (sid,))
                row = cur.fetchone()
                if not row or row[0] != seg_type:
                    if row:
                        cur.execute("SELECT start_ts, segment_type FROM active WHERE steamid=?", (sid,))
                        start, old_type = cur.fetchone()
                        dur = (datetime.fromisoformat(now) - datetime.fromisoformat(start)).total_seconds() / 60
                        cur.execute(
                            "INSERT INTO segments VALUES (?,?,?,?,?, ?,0)",
                            (sid, name, start, now, dur, old_type)
                        )
                        cur.execute("DELETE FROM active WHERE steamid=?", (sid,))
                    cur.execute(
                        "INSERT OR REPLACE INTO active VALUES (?,?,?,?)",
                        (sid, name, now, seg_type)
                    )

            # End segments
            cur.execute("SELECT steamid, name, start_ts, segment_type FROM active")
            for sid, name, start, seg_type in cur.fetchall():
                if sid not in current:
                    dur = (datetime.fromisoformat(now) - datetime.fromisoformat(start)).total_seconds() / 60
                    cur.execute(
                        "INSERT INTO segments VALUES (?,?,?,?,?, ?,0)",
                        (sid, name, start, now, dur, seg_type)
                    )
                    cur.execute("DELETE FROM active WHERE steamid=?", (sid,))

            # Export ACTIVE segments only
            cur.execute("""
                SELECT rowid, steamid, name, start_ts, end_ts, duration
                FROM segments
                WHERE exported=0 AND segment_type='active'
            """)
            rows = cur.fetchall()

            for r in rows:
                append_session([r[2], r[1], r[3], r[4], round(r[5], 2)])
                cur.execute("UPDATE segments SET exported=1 WHERE rowid=?", (r[0],))

            if rows:
                rebuild_excel()
                if should_upload_today():
                    upload_to_discord()

            conn.commit()
            conn.close()

        except Exception as e:
            print("Error:", e)

        time.sleep(POLL_INTERVAL)

# ================= ENTRY =================
if __name__ == "__main__":
    print("Steam daemon running → idle excluded, preserved internally")
    main()
