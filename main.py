import tkinter as tk
from tkinter import messagebox, ttk
import time
import threading
from datetime import datetime
import requests
import random
import csv
import io
import os
import json

# -------------------------
# Paths & defaults
# -------------------------
documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
app_folder = os.path.join(documents_folder, "RocketLaunchCountdown")
os.makedirs(app_folder, exist_ok=True)

COUNTDOWN_HTML = os.path.join(app_folder, "countdown.html")
GONOGO_HTML = os.path.join(app_folder, "gonogo.html")
GONOGO_JS = os.path.join(app_folder, "gonogo_data.js")
SETTINGS_FILE = os.path.join(app_folder, "settings.json")

# Default CSV link you provided (kept for CSV fetch fallback)
DEFAULT_CSV_LINK = "https://docs.google.com/spreadsheets/d/1UPJTW8vH2mgEzispjg_Y_zSqYTFaLoxuoZnqleVlSZ0/export?format=csv&gid=855477916"

session = requests.Session()
appVersion = "0.3.0"

# -------------------------
# Settings helpers
# -------------------------
def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        # default settings
        s = {"mode": "sheet", "sheet_url": ""}
        save_settings(s)
        return s
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"mode": "sheet", "sheet_url": ""}

def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)

# -------------------------
# Sheet fetching / manager
# -------------------------
def col_letters_to_index(col_letters: str) -> int:
    """Convert column letters like 'A' or 'AA' to 0-based index."""
    col_letters = col_letters.upper()
    idx = 0
    for ch in col_letters:
        if 'A' <= ch <= 'Z':
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def fetch_cell(sheet_url, cell_ref, timeout=2):
    """Fetch a specific cell value from a public Google Sheet CSV link.

    cell_ref like 'L2' or 'AA10'. Returns string value or 'Error: ...'.
    """
    try:
        # add cache-buster to avoid stale cached CSV from Google
        if sheet_url:
            cb = int(time.time() * 1000)
            fetch_url = sheet_url + ("&" if "?" in sheet_url else "?") + f"cb={cb}"
        else:
            fetch_url = sheet_url
        resp = session.get(fetch_url, timeout=timeout)
        resp.raise_for_status()
        reader = csv.reader(io.StringIO(resp.text))
        data = list(reader)
        # parse cell_ref
        # split letters then digits
        letters = ''.join([c for c in cell_ref if c.isalpha()])
        digits = ''.join([c for c in cell_ref if c.isdigit()])
        # invalid cell ref -> treat as failure
        if not letters or not digits:
            return None
        col = col_letters_to_index(letters)
        row = int(digits) - 1
        if row < 0 or col < 0:
            return None
        if row >= len(data) or col >= len(data[row]):
            return None
        return data[row][col].strip()
    except Exception as e:
        # network or parsing error -> return None so caller can treat as failure
        print(f"[WARN] fetch_cell error for {cell_ref} @ {sheet_url}: {e}")
        return None


# Note: single-sheet behavior only. Multi-sheet manager removed; use top-level
# settings keys `sheet_url` and `sheet_cells` in the Settings dialog.

# -------------------------
# Fetch (CSV) Go/No-Go (fallback/prefill)
# -------------------------
def fetch_gonogo_csv(csv_link=DEFAULT_CSV_LINK, timeout=3):
    """
    Fetch Go/No-Go parameters from CSV export (rows 2-4, col 12).
    Returns [Range, Weather, Vehicle] (strings).
    """
    try:
        # add cache-buster to try to get fresh CSV content from Google
        cb = int(time.time() * 1000)
        fetch_url = csv_link + ("&" if "?" in csv_link else "?") + f"cb={cb}"
        resp = session.get(fetch_url, timeout=timeout)
        resp.raise_for_status()
        reader = csv.reader(io.StringIO(resp.text))
        data = list(reader)
        gonogo = []
        # rows index 1,2,3 correspond to spreadsheet rows 2,3,4
        for i in [1, 2, 3]:
            value = data[i][11] if len(data) > i and len(data[i]) > 11 else "N/A"
            gonogo.append(value.strip().upper())
        return gonogo
    except Exception as e:
        print(f"[WARN] Failed to fetch Go/No-Go CSV: {e}")
        return None


def parse_csv_and_get_cell(csv_text, cell_ref):
    """Parse CSV text and return the value at cell_ref (A1-style) or None if missing."""
    try:
        reader = csv.reader(io.StringIO(csv_text))
        data = list(reader)
        letters = ''.join([c for c in cell_ref if c.isalpha()])
        digits = ''.join([c for c in cell_ref if c.isdigit()])
        # invalid cell ref -> signal failure
        if not letters or not digits:
            return None
        col = col_letters_to_index(letters)
        row = int(digits) - 1
        if row < 0 or col < 0:
            return None
        if row >= len(data) or col >= len(data[row]):
            return None
        return data[row][col].strip()
    except Exception:
        return None


def try_fetch_csv_text(url, timeout=3):
    """Fetch CSV text with cache-buster; return text or None."""
    try:
        if not url:
            return None
        cb = int(time.time() * 1000)
        fetch_url = url + ("&" if "?" in url else "?") + f"cb={cb}"
        resp = session.get(fetch_url, timeout=timeout)
        resp.raise_for_status()
        return resp.text
    except Exception:
        return None


def build_gviz_csv_url(sheet_url):
    """Try to construct a gviz CSV URL from common sheet URLs.
    Example: https://docs.google.com/spreadsheets/d/<id>/gviz/tq?tqx=out:csv&gid=<gid>
    Returns None if it can't parse an id.
    """
    try:
        if not sheet_url:
            return None
        parts = sheet_url.split('/d/')
        if len(parts) < 2:
            return None
        spid = parts[1].split('/')[0]
        gid = None
        # try to extract gid param if present
        if 'gid=' in sheet_url:
            for seg in sheet_url.replace('?', '&').split('&'):
                if seg.startswith('gid='):
                    gid = seg.split('=', 1)[1]
                    break
        if not gid:
            gid = '0'
        return f"https://docs.google.com/spreadsheets/d/{spid}/gviz/tq?tqx=out:csv&gid={gid}"
    except Exception:
        return None


def fetch_gonogo_values(sheet_url=None, csv_fallback=None, cells=None, timeout=1):
    """Attempt to fetch the three gonogo values (Range, Weather, Vehicle).
    Returns a list [R, W, V] or None on failure.
    Tries gviz CSV first (usually fresher), then the export CSV, then fetch_gonogo_csv.
    """
    try:
        url = sheet_url or csv_fallback or DEFAULT_CSV_LINK
        # try gviz CSV first
        gviz = build_gviz_csv_url(url)
        csv_text = None
        if gviz:
            csv_text = try_fetch_csv_text(gviz, timeout=min(2, timeout))
        if not csv_text:
            csv_text = try_fetch_csv_text(url, timeout=timeout)
        if csv_text:
            # If specific cell mappings provided, use them and treat missing mapping as failure
            if cells:
                try:
                    r = parse_csv_and_get_cell(csv_text, cells.get('Range', ''))
                    w = parse_csv_and_get_cell(csv_text, cells.get('Weather', ''))
                    v = parse_csv_and_get_cell(csv_text, cells.get('Vehicle', ''))
                    if r is not None and w is not None and v is not None:
                        return [r.strip().upper(), w.strip().upper(), v.strip().upper()]
                    return None
                except Exception:
                    return None
            # No mappings provided: fall back to default L2/L3/L4 positions (column 12 -> index 11)
            reader = csv.reader(io.StringIO(csv_text))
            data = list(reader)
            vals = []
            for i in [1,2,3]:
                v = data[i][11] if len(data) > i and len(data[i]) > 11 else "N/A"
                vals.append(v.strip().upper())
            return vals
        # last-resort: use existing CSV fetch helper which does similar work
        return fetch_gonogo_csv(url)
    except Exception:
        return None

# -------------------------
# Utility
# -------------------------
def get_status_color(status):
    try:
        return "green" if str(status).strip().upper() == "GO" else "red"
    except Exception:
        return "white"

def ensure_iframe_url(url):
    """
    Convert many common Google sheet URLs to an embeddable pubhtml URL.
    If user pasted a 'publish to web' link already, return as-is.
    If they pasted an /edit? URL, attempt to convert to preview/pub versions.
    """
    if not url:
        return ""
    if "pubhtml" in url:
        return url
    # If it's the spreadsheet "edit" URL, try to convert to /preview (works in many cases)
    if "/edit" in url:
        return url.split("/edit")[0] + "/preview"
    # If it's the direct docs/d/<id>/ URL without pubhtml, attempt the embed pattern
    # This won't always preserve sheet/tab selection; best is to instruct users to Publish to web.
    return url

# -------------------------
# HTML writers
# -------------------------
def write_countdown_html(mission_name, timer_text):
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    margin: 0;
    background-color: black;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    color: white;
    font-family: Consolas, monospace;
}}
#mission {{ font-size: 4vw; margin-bottom: 0; }}
#timer {{ font-size: 8vw; margin-bottom: 40px; }}
</style>
<script>setTimeout(()=>location.reload(),1000);</script>
</head>
<body>
<div id="mission">{mission_name}</div>
<div id="timer">{timer_text}</div>
</body>
</html>"""
    with open(COUNTDOWN_HTML, "w", encoding="utf-8") as f:
        f.write(html)

def write_gonogo_html_from_values(values):
    """
    values: [Range, Weather, Vehicle] strings like "GO" or "NO-GO"
    """
    # ensure three items:
    vals = (values + ["N/A","N/A","N/A"])[:3]
    ts = int(time.time() * 1000)
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
body {{
    margin: 0;
    background-color: black;
    color: white;
    font-family: Consolas, monospace;
    display: flex;
    justify-content: center;
    align-items: center;
    height: 100vh;
}}
#gonogo {{
    display: flex;
    gap: 40px;
}}
.status-box {{
    border: 2px solid white;
    padding: 20px 40px;
    font-size: 2.5vw;
    text-align: center;
    background-color: #111;
}}
.go {{ color: #0f0; }}
.nogo {{ color: #f00; }}
</style>
</script>
<script>
// Reload frequently but keep it small to be responsive in OBS browser-source
setTimeout(()=>location.reload(),1000);
</script>
</head>
<body>
<!-- updated: {ts} -->
<div id="gonogo">
    <div class="status-box {'go' if vals[0].strip().upper()=='GO' else 'nogo'}">Range: {vals[0]}</div>
    <div class="status-box {'go' if vals[2].strip().upper()=='GO' else 'nogo'}">Vehicle: {vals[2]}</div>
    <div class="status-box {'go' if vals[1].strip().upper()=='GO' else 'nogo'}">Weather: {vals[1]}</div>
</div>
</body>
</html>"""
    # atomic write: write to temp file then rename
    tmp = GONOGO_HTML + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html)
    try:
        os.replace(tmp, GONOGO_HTML)
    except Exception:
        # best-effort fallback
        with open(GONOGO_HTML, "w", encoding="utf-8") as f:
            f.write(html)

def write_gonogo_html_iframe(sheet_embed_url):
    """
    Writes a gonogo.html that embeds the published Google Sheet via iframe.
    Prefer users to publish-to-web and paste the pubhtml URL.
    """
    # safe empty handling
    ts = int(time.time() * 1000)
    if not sheet_embed_url:
        content = "<div style='color:orange'>No sheet URL set in settings</div>"
    else:
        # add cache-busting query param so iframe reloads when we update the file
        glue = '&' if '?' in sheet_embed_url else '?'
        content = f'<iframe src="{sheet_embed_url}{glue}cb={ts}" width="100%" height="600" style="border:none;"></iframe>'
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<style>
body {{
    margin: 0;
    background-color: black;
    color: white;
    font-family: Consolas, monospace;
}}
.container {{
    width: 95%;
    margin: 10px auto;
}}
</style>
<script>
setTimeout(()=>location.reload(),1000);
</script>
</head>
<body>
<!-- updated: {ts} -->
<div class="container">
{content}
</div>
</body>
</html>"""
    tmp = GONOGO_HTML + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(html)
    try:
        os.replace(tmp, GONOGO_HTML)
    except Exception:
        with open(GONOGO_HTML, "w", encoding="utf-8") as f:
            f.write(html)

# -------------------------
# Countdown App
# -------------------------
class CountdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"RocketLaunchCountdown {appVersion}")
        self.root.config(bg="black")
        self.root.attributes("-topmost", True)
        self.root.geometry("900x700")

        # state
        self.running = False
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.target_time = None
        self.hold_start_time = None
        self.remaining_time = 0
        self.mission_name = "Placeholder Mission"

        # manual gonogo statuses (default N/A)
        self.manual_range = "N/A"
        self.manual_weather = "N/A"
        self.manual_vehicle = "N/A"

        # load settings
        self.settings = load_settings()
        # if sheet mode and sheet_url empty, try to fallback to CSV link to prefill GUI
        self.csv_link = DEFAULT_CSV_LINK

        # initial gonogo_html: if settings.mode == sheet -> iframe embed, else fill from CSV/fallback or manual
        if self.settings.get("mode", "sheet") == "sheet":
            embed_url = ensure_iframe_url(self.settings.get("sheet_url", ""))
            write_gonogo_html_iframe(embed_url)
        else:
            # manual: populate from csv fetch as initial values or keep N/A
            vals = fetch_gonogo_csv(self.csv_link)
            # csv returns [Range, Weather, Vehicle] index 0,1,2 -> but our layout expects [Range, Weather, Vehicle]
            write_gonogo_html_from_values(vals)
            self.manual_range, self.manual_weather, self.manual_vehicle = vals[0], vals[1], vals[2]

        write_countdown_html("Placeholder Mission", "T-00:00:00")

        # GUI layout
        self.titletext = tk.Label(root, text="RocketLaunchCountdown", font=("Consolas", 24), fg="white", bg="black")
        self.titletext.pack(pady=(10, 0))

        self.text = tk.Label(root, text="T-00:00:00", font=("Consolas", 80, "bold"), fg="white", bg="black")
        self.text.pack(pady=(0, 5))

        # mission input
        frame_top = tk.Frame(root, bg="black")
        frame_top.pack(pady=5)
        tk.Label(frame_top, text="Mission Name:", fg="white", bg="black").pack(side="left")
        self.mission_entry = tk.Entry(frame_top, width=30, font=("Arial", 16))
        self.mission_entry.insert(0, self.mission_name)
        self.mission_entry.pack(side="left", padx=5)

        # Mode toggle (duration/clock)
        frame_mode = tk.Frame(root, bg="black")
        frame_mode.pack(pady=5)
        self.mode_var = tk.StringVar(value="duration")
        self.radio_duration = tk.Radiobutton(frame_mode, text="Duration", variable=self.mode_var, value="duration",
                                             fg="white", bg="black", selectcolor="black", command=self.update_inputs)
        self.radio_duration.pack(side="left", padx=5)
        self.radio_clock = tk.Radiobutton(frame_mode, text="Clock Time", variable=self.mode_var, value="clock",
                                          fg="white", bg="black", selectcolor="black", command=self.update_inputs)
        self.radio_clock.pack(side="left", padx=5)

        # Duration inputs
        frame_duration = tk.Frame(root, bg="black")
        frame_duration.pack(pady=5)
        tk.Label(frame_duration, text="H", fg="white", bg="black").pack(side="left")
        self.hours_entry = tk.Entry(frame_duration, width=3, font=("Arial", 16))
        self.hours_entry.insert(0, "0")
        self.hours_entry.pack(side="left", padx=2)
        tk.Label(frame_duration, text="M", fg="white", bg="black").pack(side="left")
        self.minutes_entry = tk.Entry(frame_duration, width=3, font=("Arial", 16))
        self.minutes_entry.insert(0, "5")
        self.minutes_entry.pack(side="left", padx=2)
        tk.Label(frame_duration, text="S", fg="white", bg="black").pack(side="left")
        self.seconds_entry = tk.Entry(frame_duration, width=3, font=("Arial", 16))
        self.seconds_entry.insert(0, "0")
        self.seconds_entry.pack(side="left", padx=2)

        # Clock input
        frame_clock = tk.Frame(root, bg="black")
        frame_clock.pack(pady=5)
        tk.Label(frame_clock, text="HH:MM", fg="white", bg="black").pack(side="left")
        self.clock_entry = tk.Entry(frame_clock, width=7, font=("Arial", 16))
        self.clock_entry.insert(0, "14:00")
        self.clock_entry.pack(side="left", padx=5)

        # control buttons
        frame_buttons = tk.Frame(root, bg="black")
        frame_buttons.pack(pady=10)
        self.start_btn = tk.Button(frame_buttons, text="▶ Start", command=self.start, font=("Arial", 14))
        self.start_btn.grid(row=0, column=0, padx=5)
        self.hold_btn = tk.Button(frame_buttons, text="⏸ Hold", command=self.hold, font=("Arial", 14))
        self.hold_btn.grid(row=0, column=1, padx=5)
        self.resume_btn = tk.Button(frame_buttons, text="⏵ Resume", command=self.resume, font=("Arial", 14))
        self.resume_btn.grid(row=0, column=1, padx=5)
        self.resume_btn.grid_remove()
        self.scrub_btn = tk.Button(frame_buttons, text="🚫 Scrub", command=self.scrub, font=("Arial", 14), fg="red")
        self.scrub_btn.grid(row=0, column=2, padx=5)
        self.reset_btn = tk.Button(frame_buttons, text="⟳ Reset", command=self.reset, font=("Arial", 14))
        self.reset_btn.grid(row=0, column=3, padx=5)

        # Settings + manual GO/NO-GO controls
        ctrl_frame = tk.Frame(root, bg="black")
        ctrl_frame.pack(pady=10)

        self.settings_btn = tk.Button(ctrl_frame, text="⚙ Settings", command=self.open_settings, font=("Arial", 12))
        self.settings_btn.pack(side="left", padx=6)

        # Quick view for sheet data (single-sheet mode)
        self.sheet_show_btn = tk.Button(ctrl_frame, text="Show Sheet Data", command=self.show_sheet_data, font=("Arial", 12))
        self.sheet_show_btn.pack(side="left", padx=6)
        self.refresh_btn = tk.Button(ctrl_frame, text="Refresh Now", command=self.refresh_gonogo_now, font=("Arial", 12))
        self.refresh_btn.pack(side="left", padx=6)
        # Rapid poll button: temporarily speed up polling for quick spreadsheet edits
        self.rapid_btn = tk.Button(ctrl_frame, text="Rapid Poll", command=lambda: self.start_rapid_poll(15), font=("Arial", 12))
        self.rapid_btn.pack(side="left", padx=6)

        # Manual GO/NO-GO toggles (visible only in manual mode)
        self.manual_frame = tk.Frame(ctrl_frame, bg="black")
        self.manual_frame.pack(side="left", padx=10)
        tk.Label(self.manual_frame, text="Manual:", fg="white", bg="black").pack(side="left", padx=(0,6))
        self.range_btn = tk.Button(self.manual_frame, text="Range: N/A", command=self.toggle_range, font=("Arial", 12))
        self.range_btn.pack(side="left", padx=4)
        self.weather_btn = tk.Button(self.manual_frame, text="Weather: N/A", command=self.toggle_weather, font=("Arial", 12))
        self.weather_btn.pack(side="left", padx=4)
        self.vehicle_btn = tk.Button(self.manual_frame, text="Vehicle: N/A", command=self.toggle_vehicle, font=("Arial", 12))
        self.vehicle_btn.pack(side="left", padx=4)

        # Go/No-Go display labels in main GUI (mirror)
        frame_gn = tk.Frame(root, bg="black")
        frame_gn.pack(pady=10)
        self.range_label = tk.Label(frame_gn, text="RANGE: N/A", font=("Consolas", 18), fg="white", bg="black")
        self.range_label.pack()
        self.weather_label = tk.Label(frame_gn, text="WEATHER: N/A", font=("Consolas", 18), fg="white", bg="black")
        self.weather_label.pack()
        self.vehicle_label = tk.Label(frame_gn, text="VEHICLE: N/A", font=("Consolas", 18), fg="white", bg="black")
        self.vehicle_label.pack()

        # Initialize gonogo cache and apply initial values (non-blocking preferred)
        self.gonogo_values = ["N/A", "N/A", "N/A"]
        # rapid polling window timestamp (seconds since epoch)
        self._rapid_until = 0
        # last rapid-update timestamp (seconds since epoch)
        self.last_gonogo_update = 0
        # inflight flag so we only have one background fetch at a time
        self._gonogo_fetch_inflight = False

        # Try a quick synchronous prefill if sheet_cells are configured (small/fast requests only)
        try:
            if self.settings.get("mode", "sheet") == "sheet":
                cells = self.settings.get("sheet_cells", {})
                url = self.settings.get("sheet_url") or self.csv_link
                if cells and url:
                    # fetch CSV once and parse mapped cells locally to avoid multiple HTTP requests
                    try:
                        cb = int(time.time() * 1000)
                        fetch_url = url + ("&" if "?" in url else "?") + f"cb={cb}"
                        resp = session.get(fetch_url, timeout=2)
                        resp.raise_for_status()
                        csv_text = resp.text
                        r = parse_csv_and_get_cell(csv_text, cells.get("Range", "")) or "N/A"
                        w = parse_csv_and_get_cell(csv_text, cells.get("Weather", "")) or "N/A"
                        v = parse_csv_and_get_cell(csv_text, cells.get("Vehicle", "")) or "N/A"
                        self.gonogo_values = [r.upper(), w.upper(), v.upper()]
                    except Exception:
                        # fallback to CSV fetch helper
                        self.gonogo_values = fetch_gonogo_csv(self.csv_link) or ["N/A", "N/A", "N/A"]
                else:
                    # quick CSV fallback
                    self.gonogo_values = fetch_gonogo_csv(self.csv_link) or ["N/A", "N/A", "N/A"]
            else:
                self.gonogo_values = [self.manual_range, self.manual_weather, self.manual_vehicle]
        except Exception:
            # keep defaults on failure
            pass

        # apply initial gonogo UI and HTML
        write_gonogo_html_from_values(self.gonogo_values)
        self.range_label.config(text=f"RANGE: {self.gonogo_values[0]}", fg=get_status_color(self.gonogo_values[0]))
        self.weather_label.config(text=f"WEATHER: {self.gonogo_values[1]}", fg=get_status_color(self.gonogo_values[1]))
        self.vehicle_label.config(text=f"VEHICLE: {self.gonogo_values[2]}", fg=get_status_color(self.gonogo_values[2]))

        # start background poller for Go/No-Go updates
        self.gonogo_poll_interval = int(self.settings.get("gonogo_interval", 10))
        # failure/backoff state
        self._gonogo_failures = 0
        self._gonogo_backoff_until = 0
        self._gonogo_max_failures = int(self.settings.get("gonogo_max_failures", 5))
        self.start_gonogo_poller()

        # footer
        footer_frame = tk.Frame(root, bg="black")
        footer_frame.pack(side="bottom", pady=0, fill="x")
        self.footer_label = tk.Label(footer_frame, text=f"Made by HamsterSpaceNerd3000 — v{appVersion}",
                                        font=("Consolas", 12), fg="black", bg="white")
        self.footer_label.pack(fill="x")

        # initialize visibility + values
        self.update_inputs()
        self.apply_settings_to_ui()
        self.update_clock()

    # ----------------------------
    # Settings UI and behavior
    # ----------------------------
    def apply_settings_to_ui(self):
        # show/hide manual controls based on settings.mode
        mode = self.settings.get("mode", "sheet")
        if mode == "manual":
            self.manual_frame.pack(side="left", padx=10)
        else:
            # sheet mode -> hide manual controls
            self.manual_frame.pack_forget()

    def show_sheet_data(self):
        # Single-sheet mode: use top-level sheet_url and sheet_cells
        cells = self.settings.get("sheet_cells", {})
        url = self.settings.get("sheet_url", "")
        if not cells or not url:
            messagebox.showinfo("Info", "No sheet URL or cell mappings configured in Settings.")
            return
        output = ["--- Sheet Data ---"]
        for name, cell in cells.items():
            val = fetch_cell(url, cell)
            output.append(f"{name}: {val}")
        messagebox.showinfo("Sheet Data", "\n".join(output))

    def refresh_gonogo_now(self):
        # Reset backoff and attempt an aggressive, short polling loop (up to 1s)
        def worker():
            prev = list(self.gonogo_values)
            deadline = time.time() + 1.0  # 1 second budget to try to get fresh data
            # clear temporary backoff to allow immediate requests
            self._gonogo_failures = 0
            self._gonogo_backoff_until = 0
            while time.time() < deadline:
                try:
                    self.poll_gonogo_once()
                except Exception:
                    pass
                # if values changed, stop early
                if self.gonogo_values != prev:
                    break
                time.sleep(0.18)
        threading.Thread(target=worker, daemon=True).start()

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.config(bg="black")
        win.geometry("700x180")
        tk.Label(win, text="Data Source Mode:", bg="black", fg="white").pack(anchor="w", padx=10, pady=(10,0))
        mode_var = tk.StringVar(value=self.settings.get("mode", "sheet"))

        rb1 = tk.Radiobutton(win, text="Google Sheet (embed published sheet)", variable=mode_var, value="sheet",
                             bg="black", fg="white", selectcolor="black")
        rb1.pack(anchor="w", padx=20)
        rb2 = tk.Radiobutton(win, text="Manual GO/NOGO (use GUI buttons)", variable=mode_var, value="manual",
                             bg="black", fg="white", selectcolor="black")
        rb2.pack(anchor="w", padx=20)

        tk.Label(win, text="Google Sheet embed URL (Publish to web → copy link):", bg="black", fg="white").pack(anchor="w", padx=10, pady=(10,0))
        sheet_entry = tk.Entry(win, width=100)
        sheet_entry.insert(0, self.settings.get("sheet_url", ""))
        sheet_entry.pack(padx=10, pady=(0,10))

        # Cell mappings: allow user to specify exact cells for Range/Weather/Vehicle
        tk.Label(win, text="Cell mappings (A1-style). Use selected sheet or these values:", bg="black", fg="white").pack(anchor="w", padx=10)
        cells = self.settings.get("sheet_cells", {})
        mapping_frame = tk.Frame(win, bg="black")
        mapping_frame.pack(anchor="w", padx=10, pady=(4,10))
        tk.Label(mapping_frame, text="Range:", bg="black", fg="white").grid(row=0, column=0, sticky="w")
        range_entry = tk.Entry(mapping_frame, width=8)
        range_entry.insert(0, cells.get("Range", "L2"))
        range_entry.grid(row=0, column=1, padx=6)
        tk.Label(mapping_frame, text="Weather:", bg="black", fg="white").grid(row=0, column=2, sticky="w")
        weather_entry = tk.Entry(mapping_frame, width=8)
        weather_entry.insert(0, cells.get("Weather", "L3"))
        weather_entry.grid(row=0, column=3, padx=6)
        tk.Label(mapping_frame, text="Vehicle:", bg="black", fg="white").grid(row=0, column=4, sticky="w")
        vehicle_entry = tk.Entry(mapping_frame, width=8)
        vehicle_entry.insert(0, cells.get("Vehicle", "L4"))
        vehicle_entry.grid(row=0, column=5, padx=6)

        def save_settings_cmd():
            new_mode = mode_var.get()
            new_url = sheet_entry.get().strip()
            self.settings["mode"] = new_mode
            self.settings["sheet_url"] = new_url
            # save cell mappings to top-level settings (used when no selected_sheet is set)
            new_cells = {
                "Range": range_entry.get().strip(),
                "Weather": weather_entry.get().strip(),
                "Vehicle": vehicle_entry.get().strip()
            }
            # store mappings (only keys with non-empty values)
            self.settings["sheet_cells"] = {k: v for k, v in new_cells.items() if v}
            # single-sheet mode: mappings are stored in top-level sheet_cells only
            save_settings(self.settings)
            # regenerate gonogo HTML according to mode
            if new_mode == "sheet":
                embed = ensure_iframe_url(new_url)
                write_gonogo_html_iframe(embed)
                # prefill labels from top-level sheet_url & sheet_cells if provided
                cells = self.settings.get("sheet_cells", {})
                url = self.settings.get("sheet_url")
                if cells and url:
                    r = fetch_cell(url, cells.get("Range", "")) or "N/A"
                    w = fetch_cell(url, cells.get("Weather", "")) or "N/A"
                    v = fetch_cell(url, cells.get("Vehicle", "")) or "N/A"
                    vals = [r, w, v]
                else:
                    vals = fetch_gonogo_csv(self.csv_link)
                self.range_label.config(text=f"RANGE: {vals[0]}", fg=get_status_color(vals[0]))
                self.weather_label.config(text=f"WEATHER: {vals[1]}", fg=get_status_color(vals[1]))
                self.vehicle_label.config(text=f"VEHICLE: {vals[2]}", fg=get_status_color(vals[2]))
            else:
                # manual: write current manual values
                write_gonogo_html_from_values([self.manual_range, self.manual_weather, self.manual_vehicle])
            self.apply_settings_to_ui()
            win.destroy()

        btn_frame = tk.Frame(win, bg="black")
        btn_frame.pack(fill="x", pady=(0,10))
        tk.Button(btn_frame, text="Save", command=save_settings_cmd, width=12).pack(side="right", padx=10)

    # ----------------------------
    # Inputs
    # ----------------------------
    def update_inputs(self):
        if self.mode_var.get() == "duration":
            self.hours_entry.config(state="normal")
            self.minutes_entry.config(state="normal")
            self.seconds_entry.config(state="normal")
            self.clock_entry.config(state="disabled")
        else:
            self.hours_entry.config(state="disabled")
            self.minutes_entry.config(state="disabled")
            self.seconds_entry.config(state="disabled")
            self.clock_entry.config(state="normal")

    # ----------------------------
    # Manual toggle callbacks
    # ----------------------------
    def toggle_range(self):
        if self.settings.get("mode", "sheet") != "manual":
            return
        self.manual_range = "GO" if self.manual_range.strip().upper() != "GO" else "NOGO"
        self.range_btn.config(text=f"Range: {self.manual_range}")
        self.range_label.config(text=f"RANGE: {self.manual_range}", fg=get_status_color(self.manual_range))
        write_gonogo_html_from_values([self.manual_range, self.manual_weather, self.manual_vehicle])

    def toggle_weather(self):
        if self.settings.get("mode", "sheet") != "manual":
            return
        self.manual_weather = "GO" if self.manual_weather.strip().upper() != "GO" else "NOGO"
        self.weather_btn.config(text=f"Weather: {self.manual_weather}")
        self.weather_label.config(text=f"WEATHER: {self.manual_weather}", fg=get_status_color(self.manual_weather))
        write_gonogo_html_from_values([self.manual_range, self.manual_weather, self.manual_vehicle])

    def toggle_vehicle(self):
        if self.settings.get("mode", "sheet") != "manual":
            return
        self.manual_vehicle = "GO" if self.manual_vehicle.strip().upper() != "GO" else "NOGO"
        self.vehicle_btn.config(text=f"Vehicle: {self.manual_vehicle}")
        self.vehicle_label.config(text=f"VEHICLE: {self.manual_vehicle}", fg=get_status_color(self.manual_vehicle))
        write_gonogo_html_from_values([self.manual_range, self.manual_weather, self.manual_vehicle])

    # ----------------------------
    # Control logic
    # ----------------------------
    def start(self):
        self.mission_name = self.mission_entry.get().strip() or "Placeholder Mission"
        self.running = True
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.show_hold_button()
        try:
            if self.mode_var.get() == "duration":
                h = int(self.hours_entry.get())
                m = int(self.minutes_entry.get())
                s = int(self.seconds_entry.get())
                total_seconds = h * 3600 + m * 60 + s
            else:
                now = datetime.now()
                parts = [int(p) for p in self.clock_entry.get().split(":")]
                h, m = parts[0], parts[1]
                s = parts[2] if len(parts) == 3 else 0
                target_today = now.replace(hour=h, minute=m, second=s, microsecond=0)
                if target_today < now:
                    target_today = target_today.replace(day=now.day + 1)
                total_seconds = (target_today - now).total_seconds()
        except Exception:
            self.text.config(text="Invalid time")
            write_countdown_html(self.mission_name, "Invalid time")
            return
        self.target_time = time.time() + total_seconds
        self.remaining_time = total_seconds

    def hold(self):
        if self.running and not self.on_hold and not self.scrubbed:
            self.on_hold = True
            self.hold_start_time = time.time()
            self.remaining_time = max(0, self.target_time - self.hold_start_time)
            self.show_resume_button()

    def resume(self):
        if self.running and self.on_hold and not self.scrubbed:
            self.on_hold = False
            self.target_time = time.time() + self.remaining_time
            self.show_hold_button()

    def show_hold_button(self):
        self.resume_btn.grid_remove()
        self.hold_btn.grid()

    def show_resume_button(self):
        self.hold_btn.grid_remove()
        self.resume_btn.grid()

    def scrub(self):
        self.scrubbed = True
        self.running = False
        write_countdown_html(self.mission_name, "SCRUB")
        self.text.config(text="SCRUB")

    def reset(self):
        self.running = False
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.text.config(text="T-00:00:00")
        write_countdown_html(self.mission_name, "T-00:00:00")
        self.show_hold_button()

    # ----------------------------
    # Format/clock loop
    # ----------------------------
    def format_time(self, seconds, prefix="T-"):
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{prefix}{h:02}:{m:02}:{s:02}"

    def update_clock(self):
        now_time = time.time()

        # Timer logic
        if self.running and not self.scrubbed:
            if getattr(self, 'paused', False):
                try:
                    secs = int(self.remaining_time)
                except Exception:
                    secs = 0
                timer_text = self.format_time(secs, "T-")
            else:
                if self.on_hold:
                    elapsed = int(now_time - self.hold_start_time)
                    timer_text = self.format_time(elapsed, "H+")
                elif self.target_time:
                    diff = int(self.target_time - now_time)
                    if diff <= 0 and not self.counting_up:
                        self.counting_up = True
                        self.target_time = now_time
                        diff = 0
                    if self.counting_up:
                        elapsed = int(now_time - self.target_time)
                        timer_text = self.format_time(elapsed, "T+")
                    else:
                        timer_text = self.format_time(diff, "T-")
                else:
                    timer_text = "T-00:00:00"
        else:
            timer_text = self.text.cget("text")

        self.text.config(text=timer_text)
        write_countdown_html(self.mission_name, timer_text)

        # Gonogo updates are handled by background poller which updates `self.gonogo_values`
        # Here we just ensure the display mirrors the latest cache (non-blocking)
        vals = self.gonogo_values
        # Only update UI here; file writes are done by the poller when values change
        self.range_label.config(text=f"RANGE: {vals[0]}", fg=get_status_color(vals[0]))
        self.weather_label.config(text=f"WEATHER: {vals[1]}", fg=get_status_color(vals[1]))
        self.vehicle_label.config(text=f"VEHICLE: {vals[2]}", fg=get_status_color(vals[2]))

        # Rapid near-instant fetch: if it's been >0.1s since last rapid update, start a quick background fetch
        try:
            if now_time - getattr(self, 'last_gonogo_update', 0) > 0.1 and not getattr(self, '_gonogo_fetch_inflight', False):
                # mark inflight and perform quick fetch in background
                self._gonogo_fetch_inflight = True
                def rapid_fetch():
                    try:
                        # Try fetching mapped cells if configured
                        url = self.settings.get('sheet_url') or self.csv_link
                        cells = self.settings.get('sheet_cells', {})
                        new_vals = None
                        if cells and url:
                            # fetch CSV once via helper (gviz preferred inside helper)
                            # try to use gviz first, then export
                            gviz = build_gviz_csv_url(url)
                            csv_text = None
                            if gviz:
                                csv_text = try_fetch_csv_text(gviz, timeout=1)
                            if not csv_text:
                                csv_text = try_fetch_csv_text(url, timeout=1)
                            if csv_text:
                                r = parse_csv_and_get_cell(csv_text, cells.get('Range', ''))
                                w = parse_csv_and_get_cell(csv_text, cells.get('Weather', ''))
                                v = parse_csv_and_get_cell(csv_text, cells.get('Vehicle', ''))
                                if r is not None and w is not None and v is not None:
                                    new_vals = [str(r).upper(), str(w).upper(), str(v).upper()]
                        if new_vals is None:
                            # fallback to generic fetch helper with short timeout, pass mappings so it won't mix sources
                            new_vals = fetch_gonogo_values(url, self.csv_link, cells=cells, timeout=1)
                        if new_vals and new_vals != self.gonogo_values:
                            self.gonogo_values = new_vals
                            # schedule UI update on main thread
                            try:
                                self.root.after(0, lambda: (
                                    write_gonogo_html_from_values(self.gonogo_values),
                                    self.range_label.config(text=f"RANGE: {self.gonogo_values[0]}", fg=get_status_color(self.gonogo_values[0])),
                                    self.weather_label.config(text=f"WEATHER: {self.gonogo_values[1]}", fg=get_status_color(self.gonogo_values[1])),
                                    self.vehicle_label.config(text=f"VEHICLE: {self.gonogo_values[2]}", fg=get_status_color(self.gonogo_values[2]))
                                ))
                            except Exception:
                                pass
                        self.last_gonogo_update = time.time()
                    finally:
                        self._gonogo_fetch_inflight = False
                threading.Thread(target=rapid_fetch, daemon=True).start()
        except Exception:
            pass

        # schedule next tick
        self.root.after(500, self.update_clock)  # 2× per second

    # ----------------------------
    # Go/No-Go poller (background)
    # ----------------------------
    def start_gonogo_poller(self):
        # run the poller in a background thread to avoid blocking the UI
        def poll_loop():
            while True:
                try:
                    self.poll_gonogo_once()
                except Exception:
                    pass
                # support rapid poll mode: if _rapid_until is set and not expired, use 1s
                now = time.time()
                if getattr(self, '_rapid_until', 0) > now:
                    poll_interval = 1.0
                else:
                    # use configured interval (ensure float and minimum 0.5s)
                    try:
                        poll_interval = max(0.5, float(self.gonogo_poll_interval))
                    except Exception:
                        poll_interval = 1.0
                time.sleep(poll_interval)
        t = threading.Thread(target=poll_loop, daemon=True)
        t.start()

    def start_rapid_poll(self, seconds=15):
        """Enable rapid polling for a short duration (seconds)."""
        try:
            self._rapid_until = time.time() + float(seconds)
            # kick off a UI refresher to show rapid status on button
            try:
                self.root.after(200, self._update_rapid_button_ui)
            except Exception:
                pass
        except Exception:
            pass

    def _update_rapid_button_ui(self):
        # update the rapid button label to indicate remaining time
        now = time.time()
        if getattr(self, '_rapid_until', 0) > now:
            remaining = int(getattr(self, '_rapid_until', 0) - now)
            try:
                self.rapid_btn.config(text=f"Rapid ({remaining}s)")
            except Exception:
                pass
            # continue updating until expired
            try:
                self.root.after(500, self._update_rapid_button_ui)
            except Exception:
                pass
        else:
            try:
                self.rapid_btn.config(text="Rapid Poll")
            except Exception:
                pass

    def poll_gonogo_once(self):
        # Determine what to fetch: top-level mapping or fallback
        mode = self.settings.get("mode", "sheet")
        if mode != "sheet":
            # manual mode: nothing to poll
            return
        # check backoff window
        now = time.time()
        if getattr(self, '_gonogo_backoff_until', 0) > now:
            return

        cells = self.settings.get("sheet_cells", {})
        url = self.settings.get("sheet_url") or self.csv_link

        new_vals = None
        # Try mapped fetch first but use a single CSV HTTP request to reduce latency
        if cells and url:
            try:
                cb = int(time.time() * 1000)
                fetch_url = url + ("&" if "?" in url else "?") + f"cb={cb}"
                resp = session.get(fetch_url, timeout=3)
                resp.raise_for_status()
                csv_text = resp.text
                r = parse_csv_and_get_cell(csv_text, cells.get("Range", ""))
                w = parse_csv_and_get_cell(csv_text, cells.get("Weather", ""))
                v = parse_csv_and_get_cell(csv_text, cells.get("Vehicle", ""))
                if r is not None and w is not None and v is not None:
                    new_vals = [str(r).upper(), str(w).upper(), str(v).upper()]
            except Exception as e:
                print(f"[WARN] mapped CSV fetch failed: {e}")

        # If mapped failed or not provided, try CSV fallback (also uses cache-busted URL)
        if new_vals is None:
            csv_vals = fetch_gonogo_csv(url)
            if csv_vals is not None:
                new_vals = csv_vals

        if new_vals is None:
            # treat as failure: increment failure count and compute backoff
            self._gonogo_failures = getattr(self, '_gonogo_failures', 0) + 1
            # exponential backoff with jitter (seconds)
            backoff = min(60, (2 ** min(self._gonogo_failures, 6)))
            jitter = random.uniform(0, backoff * 0.3)
            wait = backoff + jitter
            self._gonogo_backoff_until = now + wait
            print(f"[WARN] gonogo fetch failed #{self._gonogo_failures}; backing off for {wait:.1f}s")
            # If too many failures, switch to iframe-only mode (less aggressive polling)
            if self._gonogo_failures >= getattr(self, '_gonogo_max_failures', 5):
                print("[WARN] Switching to iframe fallback due to repeated failures")
                embed = ensure_iframe_url(self.settings.get('sheet_url', ''))
                write_gonogo_html_iframe(embed)
                # set longer backoff
                self._gonogo_backoff_until = now + 300
            return

        # Success: reset failures and apply new values
        self._gonogo_failures = 0

        # if changed, update cache and write files via main thread
        if new_vals != self.gonogo_values:
            self.gonogo_values = new_vals
            def apply_update():
                write_gonogo_html_from_values(self.gonogo_values)
                self.range_label.config(text=f"RANGE: {self.gonogo_values[0]}", fg=get_status_color(self.gonogo_values[0]))
                self.weather_label.config(text=f"WEATHER: {self.gonogo_values[1]}", fg=get_status_color(self.gonogo_values[1]))
                self.vehicle_label.config(text=f"VEHICLE: {self.gonogo_values[2]}", fg=get_status_color(self.gonogo_values[2]))
            try:
                self.root.after(0, apply_update)
            except Exception:
                # if root is gone, just ignore
                pass

# -------------------------
# Main: splash then launch
# -------------------------
if __name__ == "__main__":
    def show_splash_and_start():
        splash = tk.Tk()
        splash.title(f"RocketLaunchCountdown — Initialization {appVersion}")
        splash.config(bg="black")
        splash.geometry("500x200")
        splash.attributes("-topmost", True)

        tk.Label(splash, text="RocketLaunchCountdown", fg="white", bg="black", font=("Arial", 20, "bold")).pack(pady=(12,4))
        info = tk.Label(splash, text="Loading resources...", fg="#ccc", bg="black", font=("Arial", 12))
        info.pack(pady=6)
        cont_btn = tk.Button(splash, text="Continue", state="disabled", width=16)
        cont_btn.pack(pady=8)

        init_state = {'done': False, 'error': None}

        def init_worker():
            try:
                # pre-create files
                s = load_settings()
                # create initial countdown & gonogo files
                write_countdown_html("Placeholder Mission", "T-00:00:00")
                if s.get("mode", "sheet") == "sheet":
                    write_gonogo_html_iframe(ensure_iframe_url(s.get("sheet_url", "")))
                else:
                    vals = fetch_gonogo_csv(DEFAULT_CSV_LINK)
                    write_gonogo_html_from_values(vals)
                init_state['done'] = True
            except Exception as e:
                init_state['error'] = str(e)
                init_state['done'] = True

        threading.Thread(target=init_worker, daemon=True).start()

        def check_init():
            if init_state['done']:
                if init_state['error']:
                    info.config(text=f"Initialization error: {init_state['error']}")
                    cont_btn.config(state="normal")
                else:
                    info.config(text="Ready. You may open browser sources now, then click Continue.")
                    cont_btn.config(state="normal")
                    splash.after(5000, on_continue)  # auto-continue after 5s
                return
            splash.after(200, check_init)

        def on_continue():
            splash.destroy()
            root = tk.Tk()
            app = CountdownApp(root)
            root.mainloop()

        cont_btn.config(command=on_continue)
        splash.after(200, check_init)
        splash.mainloop()

    show_splash_and_start()
