import time
import requests
import sqlite3
import os
import logging
from pathlib import Path
from datetime import datetime, timezone
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, PieChart, Reference

# ================= CONFIG & OBSERVABILITY LOGGING =================
# Standardized logging replaces basic prints for better monitoring of script health.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler()]
)

API_KEY = os.environ.get("STEAM_API_KEY")
STEAM_ID = os.environ.get("STEAM_ID")
DISCORD_WEBHOOK = os.environ.get("DISCORD_WEBHOOK_URL")

POLL_INTERVAL = 60  # Check every minute
DATA_DIR = Path("/app/data")
DB_PATH = DATA_DIR / "sessions.db"
EXCEL_FILE = DATA_DIR / "steam_sessions.xlsx"
LAST_UPLOAD_FILE = DATA_DIR / "last_discord_upload.txt"

ACTIVE_STATES = {1, 2, 5, 6} # Online, Busy, Looking to Play/Trade
IDLE_STATES = {3, 4}        # Away, Snooze

# ================= DATABASE LAYER =================
def init_db():
    """Initializes tables and ensures database integrity on startup."""
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS segments (
                steamid TEXT, name TEXT, start_ts TEXT, end_ts TEXT,
                duration REAL, segment_type TEXT, exported INTEGER DEFAULT 0
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS active (
                steamid TEXT PRIMARY KEY, name TEXT, start_ts TEXT, segment_type TEXT
            )
        """)
        # Cleanup logic: If the script crashed, move 'active' sessions to 'segments'.
        cleanup_stale_sessions(conn)

def cleanup_stale_sessions(conn):
    """Closes any sessions left hanging if the script was previously interrupted."""
    now = datetime.now(timezone.utc).isoformat()
    cur = conn.cursor()
    cur.execute("SELECT steamid, name, start_ts, segment_type FROM active")
    stale = cur.fetchall()
    if stale:
        logging.info(f"Observability: Closing {len(stale)} stale sessions from last run.")
        for sid, name, start, seg in stale:
            dur = (datetime.fromisoformat(now) - datetime.fromisoformat(start)).total_seconds() / 60
            cur.execute("INSERT INTO segments VALUES (?,?,?,?,?,?,0)", (sid, name, start, now, dur, seg))
        cur.execute("DELETE FROM active")
        conn.commit()

# ================= EXCEL ENGINE (VISUALS) =================
def ensure_excel_structure():
    """Creates a styled Excel file with multiple analytic sheets."""
    if not EXCEL_FILE.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw_Sessions"
        
        headers = ["Name", "SteamID", "Start (UTC)", "End (UTC)", "Duration (Min)"]
        ws.append(headers)
        
        # Professional header styling
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            
        wb.create_sheet("Activity_Summary")
        wb.create_sheet("Daily_Trends")
        wb.create_sheet("Dashboard")
        wb.save(EXCEL_FILE)

def rebuild_visual_analytics():
    """Aggregates data and generates visual charts for high observability."""
    try:
        wb = load_workbook(EXCEL_FILE)
        ws_raw = wb["Raw_Sessions"]
        ws_sum = wb["Activity_Summary"]
        ws_trends = wb["Daily_Trends"]
        ws_dash = wb["Dashboard"]

        # Clear existing summary/trends before recalculating
        ws_sum.delete_rows(1, ws_sum.max_row)
        ws_trends.delete_rows(1, ws_trends.max_row)
        ws_sum.append(["User", "Total Hours"])
        ws_trends.append(["Date", "Minutes Played"])

        user_map = defaultdict(float)
        trend_map = defaultdict(float)

        # Aggregate sessions from the raw data
        for row in ws_raw.iter_rows(min_row=2, values_only=True):
            if not row[0]: continue
            user_map[row[0]] += row[4]
            date_key = row[2][:10] # Extract YYYY-MM-DD
            trend_map[date_key] += row[4]

        for name, mins in user_map.items():
            ws_sum.append([name, round(mins/60, 2)])
        
        for d in sorted(trend_map.keys()):
            ws_trends.append([d, round(trend_map[d], 1)])

        # Clear and recreate Dashboard Charts
        while ws_dash.charts:
            ws_dash.remove_chart(ws_dash.charts[0])
        
        # Pie Chart: Total Playtime Share
        pie = PieChart()
        pie.add_data(Reference(ws_sum, min_col=2, min_row=1, max_row=len(user_map)+1), titles_from_data=True)
        pie.set_categories(Reference(ws_sum, min_col=1, min_row=2, max_row=len(user_map)+1))
        pie.title = "Total Gaming Share (Hours)"
        ws_dash.add_chart(pie, "A1")

        # Line Chart: Activity Trend Over Time
        line = LineChart()
        line.add_data(Reference(ws_trends, min_col=2, min_row=1, max_row=len(trend_map)+1), titles_from_data=True)
        line.set_categories(Reference(ws_trends, min_col=1, min_row=2, max_row=len(trend_map)+1))
        line.title = "Daily Activity Trend"
        line.y_axis.title = "Minutes"
        ws_dash.add_chart(line, "I1")

        # Column Auto-Width adjustment for readability
        for sheet in [ws_raw, ws_sum, ws_trends]:
            for col in sheet.columns:
                max_length = max((len(str(cell.value or "")) for cell in col), default=10)
                sheet.column_dimensions[col[0].column_letter].width = max_length + 2

        wb.save(EXCEL_FILE)
        logging.info("Observability: Excel dashboard updated.")
    except Exception as e:
        logging.error(f"Excel Update Failed: {e}")

# ================= API HELPERS =================
def steam_api_get(endpoint, params):
    """Wrapper with timeout and basic error handling for Steam API calls."""
    url = f"https://api.steampowered.com/{endpoint}"
    params['key'] = API_KEY
    try:
        r = requests.get(url, params=params, timeout=15)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        logging.warning(f"Steam API Call failed ({endpoint}): {e}")
        return None

def upload_to_discord():
    """Uploads the finalized Excel report to Discord once daily."""
    if not DISCORD_WEBHOOK or not EXCEL_FILE.exists():
        return
    try:
        with open(EXCEL_FILE, "rb") as f:
            requests.post(
                DISCORD_WEBHOOK,
                data={"content": "📊 **Daily Steam Activity Report Updated**"},
                files={"file": ("steam_sessions.xlsx", f)},
                timeout=30
            )
    except Exception as e:
        logging.error(f"Discord Upload Failed: {e}")

# ================= MAIN TRACKING LOOP =================
def main():
    if not API_KEY or not STEAM_ID:
        logging.error("Configuration Error: API Key or Steam ID missing.")
        return

    init_db()
    ensure_excel_structure()

    logging.info("Steam Tracking Daemon Online. Monitoring activity...")

    while True:
        try:
            now_dt = datetime.now(timezone.utc)
            now = now_dt.isoformat()

            # 1. Fetch friend list and summaries
            friends_data = steam_api_get("ISteamUser/GetFriendList/v1/", {"steamid": STEAM_ID})
            if not friends_data: 
                time.sleep(POLL_INTERVAL)
                continue
            
            ids = [f["steamid"] for f in friends_data["friendslist"]["friends"]] + [STEAM_ID]
            summary_data = steam_api_get("ISteamUser/GetPlayerSummaries/v2/", {"steamids": ",".join(ids)})
            if not summary_data:
                time.sleep(POLL_INTERVAL)
                continue

            # 2. Process current states
            current_snapshot = {}
            for p in summary_data["response"]["players"]:
                sid, name = p["steamid"], p["personaname"]
                state = p.get("personastate", 0)
                seg = "playing" if p.get("gameid") else ("online" if state in ACTIVE_STATES else "idle")
                
                if state in ACTIVE_STATES or state in IDLE_STATES:
                    current_snapshot[sid] = (name, seg)

            # 3. Synchronize states with Database
            with sqlite3.connect(DB_PATH) as conn:
                cur = conn.cursor()
                cur.execute("SELECT steamid, segment_type, start_ts FROM active")
                active_map = {row[0]: (row[1], row[2]) for row in cur.fetchall()}

                for sid, (name, seg) in current_snapshot.items():
                    if sid not in active_map:
                        cur.execute("INSERT INTO active VALUES (?,?,?,?)", (sid, name, now, seg))
                    elif active_map[sid][0] != seg:
                        # State changed: Close old segment and start new one
                        old_seg, start = active_map[sid]
                        dur = (now_dt - datetime.fromisoformat(start)).total_seconds() / 60
                        cur.execute("INSERT INTO segments VALUES (?,?,?,?,?,?,0)", (sid, name, start, now, dur, old_seg))
                        cur.execute("UPDATE active SET start_ts=?, segment_type=? WHERE steamid=?", (now, seg, sid))

                # Handle users who went offline
                for sid, (old_seg, start) in active_map.items():
                    if sid not in current_snapshot:
                        dur = (now_dt - datetime.fromisoformat(start)).total_seconds() / 60
                        cur.execute("INSERT INTO segments VALUES (?,?,?,?,?,?,0)", (sid, "User Offline", start, now, dur, old_seg))
                        cur.execute("DELETE FROM active WHERE steamid=?", (sid,))

                # 4. Sync new "Playing" sessions to the Excel Raw Log
                cur.execute("SELECT rowid, name, steamid, start_ts, end_ts, duration FROM segments WHERE exported=0 AND segment_type='playing'")
                rows = cur.fetchall()
                if rows:
                    wb = load_workbook(EXCEL_FILE)
                    ws = wb["Raw_Sessions"]
                    for r in rows:
                        ws.append([r[1], r[2], r[3], r[4], round(r[5], 2)])
                        cur.execute("UPDATE segments SET exported=1 WHERE rowid=?", (r[0],))
                    wb.save(EXCEL_FILE)
                    rebuild_visual_analytics()
                    
                    # Daily Discord Upload Logic
                    today = now_dt.date()
                    last_date = None
                    if LAST_UPLOAD_FILE.exists():
                        last_date = datetime.fromisoformat(LAST_UPLOAD_FILE.read_text()).date()
                    
                    if last_date != today:
                        upload_to_discord()
                        LAST_UPLOAD_FILE.write_text(now)

        except Exception as e:
            logging.error(f"Tracking Loop Error: {e}")

        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    main()