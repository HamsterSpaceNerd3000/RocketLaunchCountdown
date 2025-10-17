import tkinter as tk
import time
import threading
from datetime import datetime
import requests
import csv
import io
import os
import json


# Get the user's Documents folder (cross-platform)
documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

# Create your app folder inside Documents
app_folder = os.path.join(documents_folder, "RocketLaunchCountdown")
os.makedirs(app_folder, exist_ok=True)

# Define file paths
COUNTDOWN_HTML = os.path.join(app_folder, "countdown.html")
GONOGO_HTML = os.path.join(app_folder, "gonogo.html")
SHEET_LINK = "https://docs.google.com/spreadsheets/d/1UPJTW8vH2mgEzispjg_Y_zSqYTFaLoxuoZnqleVlSZ0/export?format=csv&gid=855477916"
session = requests.Session()
appVersion = "0.3.0"
SETTINGS_FILE = os.path.join(app_folder, "settings.json")

# Default settings
DEFAULT_SETTINGS = {
    "mode": "spreadsheet",  # or 'buttons'
    "sheet_link": SHEET_LINK,
    # rows are 1-based as shown in spreadsheet; default 2,3,4 -> indices 1,2,3
    "range_row": 2,
    "weather_row": 3,
    "vehicle_row": 4,
    "column": 12  # 1-based column (default column 12 -> index 11)
}


def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as fh:
                return json.load(fh)
    except Exception:
        pass
    # ensure default saved
    save_settings(DEFAULT_SETTINGS)
    return DEFAULT_SETTINGS.copy()


def save_settings(s):
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as fh:
            json.dump(s, fh, indent=2)
    except Exception:
        pass


# -------------------------
# Fetch Go/No-Go Data
# -------------------------
def fetch_gonogo():
    """Fetch Go/No-Go parameters either from configured spreadsheet or return manual button values."""
    settings = load_settings()
    mode = settings.get('mode', 'spreadsheet')
    # If manual mode, read values from a runtime stash (set by the GUI buttons)
    if mode == 'buttons':
        # stored values will be on the app class; fallback to N/A
        try:
            return [getattr(fetch_gonogo, 'manual_range', 'N/A'),
                    getattr(fetch_gonogo, 'manual_weather', 'N/A'),
                    getattr(fetch_gonogo, 'manual_vehicle', 'N/A')]
        except Exception:
            return ['N/A', 'N/A', 'N/A']

    # spreadsheet mode
    link = settings.get('sheet_link', SHEET_LINK)
    col = max(1, int(settings.get('column', 12))) - 1
    rows = [int(settings.get('range_row', 2)) - 1,
            int(settings.get('weather_row', 3)) - 1,
            int(settings.get('vehicle_row', 4)) - 1]
    try:
        resp = session.get(link, timeout=3)
        resp.raise_for_status()
        reader = csv.reader(io.StringIO(resp.text))
        data = list(reader)
        gonogo = []
        for r in rows:
            val = 'N/A'
            if 0 <= r < len(data) and len(data[r]) > col:
                val = data[r][col]
            gonogo.append(val.strip().upper())
        return gonogo
    except Exception as e:
        print(f"[ERROR] Failed to fetch Go/No-Go from sheet: {e}")
        return ["ERROR", "ERROR", "ERROR"]


# -------------------------
# Helper for color
# -------------------------
def get_status_color(status):
    """Return color name for a Go/No-Go status string."""
    try:
        return "green" if str(status).strip().upper() == "GO" else "red"
    except Exception:
        return "white"

# -------------------------
# Write Countdown HTML
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
<script>
setTimeout(() => location.reload(), 1000);
</script>
</head>
<body>
<div id="mission">{mission_name}</div>
<div id="timer">{timer_text}</div>
</body>
</html>"""
    with open(COUNTDOWN_HTML, "w", encoding="utf-8") as f:
        f.write(html)

# -------------------------
# Write Go/No-Go HTML
# -------------------------
def write_gonogo_html(gonogo_values=None):
    if gonogo_values is None:
        gonogo_values = ["N/A", "N/A", "N/A"]
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
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
<script>
setTimeout(() => location.reload(), 5000);
</script>
</head>
<body>
<div id="gonogo">
    <div class="status-box {'go' if gonogo_values[0].lower()=='go' else 'nogo'}">Range: {gonogo_values[0]}</div>
    <div class="status-box {'go' if gonogo_values[2].lower()=='go' else 'nogo'}">Vehicle: {gonogo_values[2]}</div>
    <div class="status-box {'go' if gonogo_values[1].lower()=='go' else 'nogo'}">Weather: {gonogo_values[1]}</div>
</div>
</body>
</html>"""
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
        self.root.geometry("800x615")

        # State
        self.running = False
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.target_time = None
        self.hold_start_time = None
        self.remaining_time = 0
        self.mission_name = "Placeholder Mission"
        # fetch_gonogo() returns [Range, Weather, Vehicle] to match gonogo.html writer
        self.gonogo_values = fetch_gonogo()
        self.last_gonogo_update = time.time()

        # Title
        self.titletext = tk.Label(root, text="RocketLaunchCountdown", font=("Consolas", 24), fg="white", bg="black")
        self.titletext.pack(pady=(10, 0))

        # Display
        self.text = tk.Label(root, text="T-00:00:00", font=("Consolas", 80, "bold"), fg="white", bg="black")
        self.text.pack(pady=(0, 5))

        # Mission name input
        frame_top = tk.Frame(root, bg="black")
        frame_top.pack(pady=5)
        tk.Label(frame_top, text="Mission Name:", fg="white", bg="black").pack(side="left")
        self.mission_entry = tk.Entry(frame_top, width=20, font=("Arial", 18))
        self.mission_entry.insert(0, self.mission_name)
        self.mission_entry.pack(side="left")

        # Mode toggle
        frame_mode = tk.Frame(root, bg="black")
        frame_mode.pack(pady=5)

        self.mode_var = tk.StringVar(value="duration")

        self.radio_duration = tk.Radiobutton(
            frame_mode,
            text="Duration",
            variable=self.mode_var,
            value="duration",
            fg="white",
            bg="black",
            selectcolor="black",  # makes the dot visible
            command=self.update_inputs
        )
        self.radio_duration.pack(side="left", padx=5)

        self.radio_clock = tk.Radiobutton(
            frame_mode,
            text="Clock Time",
            variable=self.mode_var,
            value="clock",
            fg="white",
            bg="black",
            selectcolor="black",  # makes the dot visible
            command=self.update_inputs
        )
        self.radio_clock.pack(side="left", padx=5)

        # Duration inputs
        frame_duration = tk.Frame(root, bg="black")
        frame_duration.pack(pady=5)
        tk.Label(frame_duration, text="H", fg="white", bg="black").pack(side="left")
        self.hours_entry = tk.Entry(frame_duration, width=3, font=("Arial", 18))
        self.hours_entry.insert(0, "0")
        self.hours_entry.pack(side="left", padx=2)
        tk.Label(frame_duration, text="M", fg="white", bg="black").pack(side="left")
        self.minutes_entry = tk.Entry(frame_duration, width=3, font=("Arial", 18))
        self.minutes_entry.insert(0, "5")
        self.minutes_entry.pack(side="left", padx=2)
        tk.Label(frame_duration, text="S", fg="white", bg="black").pack(side="left")
        self.seconds_entry = tk.Entry(frame_duration, width=3, font=("Arial", 18))
        self.seconds_entry.insert(0, "0")
        self.seconds_entry.pack(side="left", padx=2)

        # Clock time input
        frame_clock = tk.Frame(root, bg="black")
        frame_clock.pack(pady=5)
        tk.Label(frame_clock, text="HH:MM", fg="white", bg="black").pack(side="left")
        self.clock_entry = tk.Entry(frame_clock, width=7, font=("Arial", 18))
        self.clock_entry.insert(0, "14:00")
        self.clock_entry.pack(side="left", padx=5)

        # Control buttons
        frame_buttons = tk.Frame(root, bg="black")
        frame_buttons.pack(pady=10)

        self.start_btn = tk.Button(frame_buttons, text="▶ Start", command=self.start, font=("Arial", 14))
        self.start_btn.grid(row=0, column=0, padx=5)

        # Hold and resume share the same position
        self.hold_btn = tk.Button(frame_buttons, text="⏸ Hold", command=self.hold, font=("Arial", 14))
        self.hold_btn.grid(row=0, column=1, padx=5)

        self.resume_btn = tk.Button(frame_buttons, text="⏵ Resume", command=self.resume, font=("Arial", 14))
        self.resume_btn.grid(row=0, column=1, padx=5)
        self.resume_btn.grid_remove()  # hidden at start

        self.scrub_btn = tk.Button(frame_buttons, text="🚫 Scrub", command=self.scrub, font=("Arial", 14), fg="red")
        self.scrub_btn.grid(row=0, column=2, padx=5)

        self.reset_btn = tk.Button(frame_buttons, text="⟳ Reset", command=self.reset, font=("Arial", 14))
        self.reset_btn.grid(row=0, column=3, padx=5)
        # Settings button moved next to control buttons (match size/style)
        settings_btn = tk.Button(frame_buttons, text="Settings", command=self.show_settings_window, font=("Arial", 14), width=10)
        settings_btn.grid(row=0, column=4, padx=6)

        # Note: gonogo mode switching remains in Settings; manual buttons appear when mode == 'buttons'

        # Manual Go/No-Go buttons will go next to control buttons
        self.manual_frame = tk.Frame(root, bg="black")
        self.manual_frame.pack(pady=6)

        # Buttons now toggle current state between GO and NOGO
        self.range_toggle_btn = tk.Button(self.manual_frame, text="Range: Toggle", width=12,
                                          command=lambda: self._toggle_manual('range'))
        self.weather_toggle_btn = tk.Button(self.manual_frame, text="Weather: Toggle", width=12,
                                           command=lambda: self._toggle_manual('weather'))
        self.vehicle_toggle_btn = tk.Button(self.manual_frame, text="Vehicle: Toggle", width=12,
                                           command=lambda: self._toggle_manual('vehicle'))

        # Placeholders; visibility will be controlled by settings
        self.range_toggle_btn.grid(row=0, column=0, padx=4, pady=2)
        self.weather_toggle_btn.grid(row=0, column=1, padx=4, pady=2)
        self.vehicle_toggle_btn.grid(row=0, column=2, padx=4, pady=2)

        frame_gn = tk.Frame(root, bg="black")
        frame_gn.pack(pady=10)
        # Labels displayed: Range, Weather, Vehicle — match write_gonogo_html ordering
        self.range_label = tk.Label(frame_gn, text="RANGE: N/A", font=("Consolas", 20), fg="white", bg="black")
        self.range_label.pack()
        self.weather_label = tk.Label(frame_gn, text="WEATHER: N/A", font=("Consolas", 20), fg="white", bg="black")
        self.weather_label.pack()
        self.vehicle_label = tk.Label(frame_gn, text="VEHICLE: N/A", font=("Consolas", 20), fg="white", bg="black")
        self.vehicle_label.pack()

        # Footer
        footer_frame = tk.Frame(root, bg="black")
        footer_frame.pack(side="bottom", pady=0, fill="x")

        self.footer_label = tk.Label(
            footer_frame,
            text="Made by HamsterSpaceNerd3000",  # or whatever you want
            font=("Consolas", 12),
            fg="black",
            bg="white"
        )
        self.footer_label.pack(fill="x")
        self.update_inputs()
        # set initial manual button visibility from settings
        self.update_manual_visibility()
        self.update_clock()

    # ----------------------------
    # Settings window
    # ----------------------------
    def show_settings_window(self):
        settings = load_settings()

        win = tk.Toplevel(self.root)
        win.config(bg="black")
        win.title("Settings")
        win.geometry("560x200")
        win.transient(self.root)

        # Mode selection
        frame_mode = tk.Frame(win)
        frame_mode.config(bg="black")
        frame_mode.pack(fill='x', pady=8, padx=8)
        tk.Label(frame_mode, text="Mode:", fg="white", bg="black").pack(side='left')
        mode_var = tk.StringVar(value=settings.get('mode', 'spreadsheet'))
        tk.Radiobutton(frame_mode, text='Spreadsheet', variable=mode_var, value='spreadsheet', fg="white", bg="black", selectcolor="black").pack(side='left', padx=8)
        tk.Radiobutton(frame_mode, text='Buttons (manual)', variable=mode_var, value='buttons', fg="white", bg="black", selectcolor="black").pack(side='left', padx=8)

        # Spreadsheet config
        frame_sheet = tk.LabelFrame(win, text='Spreadsheet configuration', fg='white', bg='black')
        frame_sheet.config(bg="black")
        frame_sheet.pack(fill='x', padx=8, pady=6)
        tk.Label(frame_sheet, text='Sheet link (CSV export):', fg='white', bg='black').pack(anchor='w')
        sheet_entry = tk.Entry(frame_sheet, width=80, fg='white', bg='#222', insertbackground='white')
        sheet_entry.pack(fill='x', padx=6, pady=4)
        sheet_entry.insert(0, settings.get('sheet_link', SHEET_LINK))

        # Accept cells in 'L3' format for each parameter
        cell_frame = tk.Frame(frame_sheet)
        cell_frame.config(bg="black")
        cell_frame.pack(fill='x', padx=6, pady=2)
        tk.Label(cell_frame, text='Range cell (e.g. L3):', fg='white', bg='black').grid(row=0, column=0)
        range_cell = tk.Entry(cell_frame, width=8, fg='white', bg='#222', insertbackground='white')
        range_cell.grid(row=0, column=1, padx=4)
        # show as L3 if present, otherwise build from numeric settings
        try:
            if 'range_cell' in settings:
                range_cell.insert(0, settings.get('range_cell'))
            else:
                # convert numeric row/column to cell like L3
                col = settings.get('column', DEFAULT_SETTINGS['column'])
                row = settings.get('range_row', DEFAULT_SETTINGS['range_row'])
                # column number to letters
                def col_to_letters(n):
                    s = ''
                    while n > 0:
                        n, r = divmod(n - 1, 26)
                        s = chr(ord('A') + r) + s
                    return s
                range_cell.insert(0, f"{col_to_letters(col)}{row}")
        except Exception:
            range_cell.insert(0, f"L3")

        tk.Label(cell_frame, text='Weather cell (e.g. L4):', fg='white', bg='black').grid(row=0, column=2)
        weather_cell = tk.Entry(cell_frame, width=8, fg='white', bg='#222', insertbackground='white')
        weather_cell.grid(row=0, column=3, padx=4)
        try:
            if 'weather_cell' in settings:
                weather_cell.insert(0, settings.get('weather_cell'))
            else:
                col = settings.get('column', DEFAULT_SETTINGS['column'])
                row = settings.get('weather_row', DEFAULT_SETTINGS['weather_row'])
                def col_to_letters(n):
                    s = ''
                    while n > 0:
                        n, r = divmod(n - 1, 26)
                        s = chr(ord('A') + r) + s
                    return s
                weather_cell.insert(0, f"{col_to_letters(col)}{row}")
        except Exception:
            weather_cell.insert(0, f"L4")

        tk.Label(cell_frame, text='Vehicle cell (e.g. L5):', fg='white', bg='black').grid(row=0, column=4)
        vehicle_cell = tk.Entry(cell_frame, width=8, fg='white', bg='#222', insertbackground='white')
        vehicle_cell.grid(row=0, column=5, padx=4)
        try:
            if 'vehicle_cell' in settings:
                vehicle_cell.insert(0, settings.get('vehicle_cell'))
            else:
                col = settings.get('column', DEFAULT_SETTINGS['column'])
                row = settings.get('vehicle_row', DEFAULT_SETTINGS['vehicle_row'])
                def col_to_letters(n):
                    s = ''
                    while n > 0:
                        n, r = divmod(n - 1, 26)
                        s = chr(ord('A') + r) + s
                    return s
                vehicle_cell.insert(0, f"{col_to_letters(col)}{row}")
        except Exception:
            vehicle_cell.insert(0, f"L5")

        # Manual buttons config
        frame_buttons_cfg = tk.LabelFrame(win, text='Manual Go/No-Go (Buttons mode)', fg='white', bg='black')
        frame_buttons_cfg.config(bg='black')
        frame_buttons_cfg.pack(fill='x', padx=8, pady=6)

        def set_manual(val_type, val):
            # store on fetch_gonogo func for now
            if val_type == 'range':
                fetch_gonogo.manual_range = val
            elif val_type == 'weather':
                fetch_gonogo.manual_weather = val
            elif val_type == 'vehicle':
                fetch_gonogo.manual_vehicle = val

        # helper to set manual and update UI from main app
        def set_manual_and_update(val_type, val):
            set_manual(val_type, val)
            # update labels and write html
            self.gonogo_values = fetch_gonogo()
            # update GUI labels immediately
            self.range_label.config(text=f"RANGE: {self.gonogo_values[0]}", fg=get_status_color(self.gonogo_values[0]))
            self.weather_label.config(text=f"WEATHER: {self.gonogo_values[1]}", fg=get_status_color(self.gonogo_values[1]))
            self.vehicle_label.config(text=f"VEHICLE: {self.gonogo_values[2]}", fg=get_status_color(self.gonogo_values[2]))
            write_gonogo_html(self.gonogo_values)

        # Save/Cancel
        def cell_to_rc(cell_str):
            s = (cell_str or '').strip().upper()
            if not s:
                return None, None
            # split letters and digits
            letters = ''
            digits = ''
            for ch in s:
                if ch.isalpha():
                    letters += ch
                elif ch.isdigit():
                    digits += ch
            if not digits:
                return None, None
            # convert letters to number
            col = 0
            for ch in letters:
                col = col * 26 + (ord(ch) - ord('A') + 1)
            row = int(digits)
            return row, col

        def on_save():
            # parse cells
            r_row, r_col = cell_to_rc(range_cell.get())
            w_row, w_col = cell_to_rc(weather_cell.get())
            v_row, v_col = cell_to_rc(vehicle_cell.get())
            # fallbacks
            if r_row is None:
                r_row = DEFAULT_SETTINGS['range_row']
            if w_row is None:
                w_row = DEFAULT_SETTINGS['weather_row']
            if v_row is None:
                v_row = DEFAULT_SETTINGS['vehicle_row']
            # determine column to use (prefer range column, else weather, else vehicle, else default)
            col_val = r_col or w_col or v_col or DEFAULT_SETTINGS['column']
            new_settings = {
                'mode': mode_var.get(),
                'sheet_link': sheet_entry.get().strip() or SHEET_LINK,
                'range_row': int(r_row),
                'weather_row': int(w_row),
                'vehicle_row': int(v_row),
                'column': int(col_val),
                # persist the textual cells for convenience
                'range_cell': range_cell.get().strip().upper(),
                'weather_cell': weather_cell.get().strip().upper(),
                'vehicle_cell': vehicle_cell.get().strip().upper(),
                # persist manual values if present
                'manual_range': getattr(fetch_gonogo, 'manual_range', None),
                'manual_weather': getattr(fetch_gonogo, 'manual_weather', None),
                'manual_vehicle': getattr(fetch_gonogo, 'manual_vehicle', None)
            }
            save_settings(new_settings)
            # update immediately
            self.gonogo_values = fetch_gonogo()
            write_gonogo_html(self.gonogo_values)
            # update manual visibility in main UI
            self.update_manual_visibility()
            win.destroy()

        def on_cancel():
            win.destroy()

        btn_frame = tk.Frame(win)
        btn_frame = tk.Frame(win, bg='black')
        btn_frame.pack(fill='x', pady=8)
        tk.Button(btn_frame, text='Save', command=on_save, fg='white', bg='#333', activebackground='#444').pack(side='right', padx=8)
        tk.Button(btn_frame, text='Cancel', command=on_cancel, fg='white', bg='#333', activebackground='#444').pack(side='right')


    # ----------------------------
    # Update input visibility based on mode
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
    # Manual controls & helpers
    # ----------------------------
    def set_manual(self, which, val):
        # normalize
        v = (val or '').strip().upper()
        if which == 'range':
            fetch_gonogo.manual_range = v
        elif which == 'weather':
            fetch_gonogo.manual_weather = v
        elif which == 'vehicle':
            fetch_gonogo.manual_vehicle = v
        # update GUI and HTML
        self.gonogo_values = fetch_gonogo()
        try:
            self.range_label.config(text=f"RANGE: {self.gonogo_values[0]}", fg=get_status_color(self.gonogo_values[0]))
            self.weather_label.config(text=f"WEATHER: {self.gonogo_values[1]}", fg=get_status_color(self.gonogo_values[1]))
            self.vehicle_label.config(text=f"VEHICLE: {self.gonogo_values[2]}", fg=get_status_color(self.gonogo_values[2]))
        except Exception:
            pass
        write_gonogo_html(self.gonogo_values)
        # persist manual values immediately so they survive restarts
        try:
            s = load_settings()
            s['manual_range'] = getattr(fetch_gonogo, 'manual_range', s.get('manual_range'))
            s['manual_weather'] = getattr(fetch_gonogo, 'manual_weather', s.get('manual_weather'))
            s['manual_vehicle'] = getattr(fetch_gonogo, 'manual_vehicle', s.get('manual_vehicle'))
            save_settings(s)
        except Exception:
            pass

    def update_manual_visibility(self):
        s = load_settings()
        mode = s.get('mode', 'spreadsheet')
        visible = (mode == 'buttons')
        # show or hide manual frame
        if visible:
            self.manual_frame.pack(pady=6)
        else:
            self.manual_frame.pack_forget()

    def _toggle_manual(self, which):
        # get current values (Range, Weather, Vehicle)
        cur = fetch_gonogo()
        # map which to index
        idx_map = {'range': 0, 'weather': 1, 'vehicle': 2}
        idx = idx_map.get(which, 0)
        try:
            cur_val = (cur[idx] or '').strip().upper()
        except Exception:
            cur_val = 'N/A'
        # toggle: if GO -> NOGO, else -> GO
        new_val = 'NOGO' if cur_val == 'GO' else 'GO'
        self.set_manual(which, new_val)

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
    # Clock updating
    # ----------------------------
    def format_time(self, seconds, prefix="T-"):
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{prefix}{h:02}:{m:02}:{s:02}"

    def update_clock(self):
        now_time = time.time()

        # Update timer
        if self.running and not self.scrubbed:
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

        # Update Go/No-Go every 10 seconds
        if now_time - self.last_gonogo_update > 0.1:
            # fetch_gonogo returns [Range, Weather, Vehicle]
            self.range_status, self.weather, self.vehicle = fetch_gonogo()
            self.range_label.config(text=f"RANGE: {self.range_status}", fg=get_status_color(self.range_status))
            self.weather_label.config(text=f"WEATHER: {self.weather}", fg=get_status_color(self.weather))
            self.vehicle_label.config(text=f"VEHICLE: {self.vehicle}", fg=get_status_color(self.vehicle))
            self.gonogo_values = [self.range_status, self.weather, self.vehicle]
            write_gonogo_html(self.gonogo_values)
            self.last_gonogo_update = now_time

        self.root.after(200, self.update_clock)


if __name__ == "__main__":
    # Show a small splash/loading GUI while we fetch initial data and write HTML files.
    def show_splash_and_start():
        splash = tk.Tk()
        splash.title("RocketLaunchCountdown — Initialaization")
        splash.config(bg="black")
        splash.geometry("400x175")
        splash.attributes("-topmost", True)

        title = tk.Label(splash, text="RocketLaunchCountdown", fg="white", bg="black", font=("Arial", 20, "bold"))
        title.pack(pady=(10,0))

        lbl = tk.Label(splash, text="Loading resources...", fg="white", bg="black", font=("Arial", 14))
        lbl.pack(pady=(0,5))

        info = tk.Label(splash, text="Fetching Go/No-Go and preparing HTML files.", fg="#ccc", bg="black", font=("Arial", 10))
        info.pack()

        cont_btn = tk.Button(splash, text="Continue", state="disabled", width=12)
        cont_btn.pack(pady=8)

        # Footer
        footer_frame = tk.Frame(splash, bg="black")
        footer_frame.pack(side="bottom", pady=0, fill="x")

        footer_label = tk.Label(
            footer_frame,
            text="Made by HamsterSpaceNerd3000",  # or whatever you want
            font=("Consolas", 12),
            fg="black",
            bg="white"
        )
        footer_label.pack(fill="x")

        # Shared flag to indicate initialization complete
        init_state = { 'done': False, 'error': None }

        def init_worker():
            try:
                # perform the same initial writes you had before
                gonogo = fetch_gonogo()
                write_countdown_html("Placeholder Mission", "T-00:00:00")
                write_gonogo_html(gonogo)
                init_state['done'] = True
            except Exception as e:
                init_state['error'] = str(e)
                init_state['done'] = True

        # Start background initialization
        threading.Thread(target=init_worker, daemon=True).start()

        def check_init():
            if init_state['done']:
                if init_state['error']:
                    info.config(text=f"Initialization error: {init_state['error']}")
                else:
                    info.config(text="Ready. You may open browser sources now, then click Continue.")
                    cont_btn.config(state="normal")
                    splash.after(5000, on_continue)
                return
            splash.after(200, check_init)

        def on_continue():
            splash.destroy()
            # now create the real main window
            root = tk.Tk()
            app = CountdownApp(root)
            root.mainloop()

        cont_btn.config(command=on_continue)
        # begin polling
        splash.after(100, check_init)
        splash.mainloop()

    show_splash_and_start()
