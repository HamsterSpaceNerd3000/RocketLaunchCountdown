import tkinter as tk
from tkinter import colorchooser
import time
import threading
from datetime import datetime, timedelta
import re
import requests
import csv
import io
import os
import json
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

# Get the user's Documents folder (cross-platform)
documents_folder = os.path.join(os.path.expanduser("~"), "Documents")

# Create your app folder inside Documents
app_folder = os.path.join(documents_folder, "RocketLaunchCountdown")
os.makedirs(app_folder, exist_ok=True)

# Define file paths
COUNTDOWN_HTML = os.path.join(app_folder, "countdown.html")
GONOGO_HTML = os.path.join(app_folder, "gonogo.html")
SHEET_LINK = ""
session = requests.Session()
appVersion = "0.6.0"
SETTINGS_FILE = os.path.join(app_folder, "settings.json")

# Default settings
DEFAULT_SETTINGS = {
    "mode": "spreadsheet",
    "sheet_link": SHEET_LINK,
    "range_row": 2,
    "weather_row": 3,
    "vehicle_row": 4,
    "column": 12,
    "hide_mission_name": False,
}
# default timezone: 'local' uses system local tz, otherwise an IANA name or 'UTC'
DEFAULT_SETTINGS.setdefault('timezone', 'local')

# Appearance defaults
DEFAULT_SETTINGS.setdefault('bg_color', '#000000')
DEFAULT_SETTINGS.setdefault('text_color', '#FFFFFF')
DEFAULT_SETTINGS.setdefault('font_family', 'Consolas')
DEFAULT_SETTINGS.setdefault('mission_font_px', 24)
DEFAULT_SETTINGS.setdefault('timer_font_px', 80)
DEFAULT_SETTINGS.setdefault('gn_bg_color', '#111111')
DEFAULT_SETTINGS.setdefault('gn_border_color', '#FFFFFF')
DEFAULT_SETTINGS.setdefault('gn_go_color', '#00FF00')
DEFAULT_SETTINGS.setdefault('gn_nogo_color', '#FF0000')
DEFAULT_SETTINGS.setdefault('gn_font_px', 20)
DEFAULT_SETTINGS.setdefault('appearance_mode', 'dark')

# HTML-only appearance defaults (these should not affect the Python GUI)
DEFAULT_SETTINGS.setdefault('html_bg_color', DEFAULT_SETTINGS.get('bg_color', '#000000'))
DEFAULT_SETTINGS.setdefault('html_text_color', DEFAULT_SETTINGS.get('text_color', '#FFFFFF'))
DEFAULT_SETTINGS.setdefault('html_font_family', DEFAULT_SETTINGS.get('font_family', 'Consolas'))
DEFAULT_SETTINGS.setdefault('html_mission_font_px', DEFAULT_SETTINGS.get('mission_font_px', 24))
DEFAULT_SETTINGS.setdefault('html_timer_font_px', DEFAULT_SETTINGS.get('timer_font_px', 80))
DEFAULT_SETTINGS.setdefault('html_gn_bg_color', DEFAULT_SETTINGS.get('gn_bg_color', '#111111'))
DEFAULT_SETTINGS.setdefault('html_gn_border_color', DEFAULT_SETTINGS.get('gn_border_color', '#FFFFFF'))
DEFAULT_SETTINGS.setdefault('html_gn_go_color', DEFAULT_SETTINGS.get('gn_go_color', '#00FF00'))
DEFAULT_SETTINGS.setdefault('html_gn_nogo_color', DEFAULT_SETTINGS.get('gn_nogo_color', '#FF0000'))
DEFAULT_SETTINGS.setdefault('html_gn_font_px', DEFAULT_SETTINGS.get('gn_font_px', 20))

# Auto-hold times: list of seconds before T at which timer should automatically enter hold
DEFAULT_SETTINGS.setdefault('auto_hold_times', [])

# A small list of common timezone choices.
TIMEZONE_CHOICES = [
    'local', 'UTC', 'US/Eastern', 'US/Central', 'US/Mountain', 'US/Pacific',
    'Europe/London', 'Europe/Paris', 'Asia/Tokyo', 'Australia/Sydney'
]

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
        s = str(status or '').strip().upper()
        # normalize to letters only so variants like 'NO GO', 'NO-GO', 'NOGO' match
        norm = re.sub(r'[^A-Z]', '', s)
        if norm == 'GO':
            return 'green'
        if norm == 'NOGO':
            return 'red'
        # fallback: treat unknown/empty as white
        return 'white'
    except Exception:
        return "white"


def format_status_display(status):
    try:
        s = str(status or '').strip().upper()
        norm = re.sub(r'[^A-Z]', '', s)
        if norm == 'GO':
            return 'GO'
        if norm == 'NOGO':
            return 'NO-GO'
        return s
    except Exception:
        return str(status or '')

# -------------------------
# Write Countdown HTML
# -------------------------
def write_countdown_html(mission_name, timer_text):
    s = load_settings()
    # Prefer HTML-specific settings; fall back to GUI appearance settings for backwards compatibility
    bg = s.get('html_bg_color', s.get('bg_color', '#000000'))
    text = s.get('html_text_color', s.get('text_color', '#FFFFFF'))
    font = s.get('html_font_family', s.get('font_family', 'Consolas, monospace'))
    mission_px = int(s.get('html_mission_font_px', s.get('mission_font_px', 48)))
    timer_px = int(s.get('html_timer_font_px', s.get('timer_font_px', 120)))
    
    # Mission name hidden setting
    hide_mission_name = s.get("hide_mission_name", False)
    mission_div_hidden = f'<div id="mission">{mission_name}</div>' if not hide_mission_name else ''

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    margin: 0;
    background-color: {bg};
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    color: {text};
    font-family: {font};
}}
#mission {{ font-size: {mission_px}px; margin-bottom: 0; }}
#timer {{ font-size: {timer_px}px; margin-bottom: 40px; }}
</style>
<script>
setTimeout(() => location.reload(), 1000);
</script>
</head>
<body>
{mission_div_hidden}
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
    s = load_settings()
    # Prefer HTML-specific settings; fall back to GUI appearance settings for backwards compatibility
    bg = s.get('html_bg_color', s.get('bg_color', '#000000'))
    text = s.get('html_text_color', s.get('text_color', '#FFFFFF'))
    font = s.get('html_font_family', s.get('font_family', 'Consolas, monospace'))
    gn_bg = s.get('html_gn_bg_color', s.get('gn_bg_color', '#111111'))
    gn_border = s.get('html_gn_border_color', s.get('gn_border_color', '#FFFFFF'))
    gn_go = s.get('html_gn_go_color', s.get('gn_go_color', '#00FF00'))
    gn_nogo = s.get('html_gn_nogo_color', s.get('gn_nogo_color', '#FF0000'))
    gn_px = int(s.get('html_gn_font_px', s.get('gn_font_px', 28)))
    # normalize and format display values so variants like 'NO GO' become 'NO-GO'
    disp0 = format_status_display(gonogo_values[0])
    disp1 = format_status_display(gonogo_values[1])
    disp2 = format_status_display(gonogo_values[2])
    n0 = re.sub(r'[^A-Z]', '', (str(gonogo_values[0] or '')).strip().upper())
    n1 = re.sub(r'[^A-Z]', '', (str(gonogo_values[1] or '')).strip().upper())
    n2 = re.sub(r'[^A-Z]', '', (str(gonogo_values[2] or '')).strip().upper())

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    margin: 0;
    background-color: {bg};
    color: {text};
    font-family: {font};
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
    border: 2px solid {gn_border};
    padding: 20px 40px;
    font-size: {gn_px}px;
    text-align: center;
    background-color: {gn_bg};
}}
.go {{ color: {gn_go}; }}
.nogo {{ color: {gn_nogo}; }}
</style>
<script>
setTimeout(() => location.reload(), 5000);
</script>
</head>
<body>
    <div id="gonogo">
    <div class="status-box {'go' if n0=='GO' else 'nogo'}">Range: {disp0}</div>
    <div class="status-box {'go' if n2=='GO' else 'nogo'}">Vehicle: {disp2}</div>
    <div class="status-box {'go' if n1=='GO' else 'nogo'}">Weather: {disp1}</div>
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
        # track which auto-holds we've already triggered for the current run
        self._auto_hold_triggered = set()

        # Title
        self.titletext = tk.Label(root, text=f"RocketLaunchCountdown {appVersion}", font=("Consolas", 24), fg="white", bg="black")
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

        # Auto-hold quick button (opens H/M/S dialog)
        def open_autohold_dialog():
            dlg = tk.Toplevel(self.root)
            dlg.transient(self.root)
            dlg.title('Set Auto-hold')
            dlg.geometry('320x110')
            # theme according to appearance mode
            ssettings = load_settings()
            mode_local = ssettings.get('appearance_mode', DEFAULT_SETTINGS.get('appearance_mode', 'dark'))
            if mode_local == 'dark':
                dlg_bg = '#000000'; dlg_fg = '#FFFFFF'; entry_bg = '#222222'; btn_bg = '#FFFFFF'; btn_fg = '#000000'
            else:
                dlg_bg = '#FFFFFF'; dlg_fg = '#000000'; entry_bg = '#b4b4b4'; btn_bg = '#000000'; btn_fg = '#FFFFFF'
            dlg.config(bg=dlg_bg)
            tk.Label(dlg, text='Auto-hold time (H M S):', fg=dlg_fg, bg=dlg_bg).pack(pady=(6,0))
            box = tk.Frame(dlg, bg=dlg_bg)
            box.pack(pady=6)
            h_entry = tk.Entry(box, width=3, font=('Arial', 12), bg=entry_bg, fg=dlg_fg, insertbackground=dlg_fg)
            m_entry = tk.Entry(box, width=3, font=('Arial', 12), bg=entry_bg, fg=dlg_fg, insertbackground=dlg_fg)
            s_entry = tk.Entry(box, width=3, font=('Arial', 12), bg=entry_bg, fg=dlg_fg, insertbackground=dlg_fg)
            h_entry.pack(side='left', padx=4)
            tk.Label(box, text='H', fg=dlg_fg, bg=dlg_bg).pack(side='left')
            m_entry.pack(side='left', padx=4)
            tk.Label(box, text='M', fg=dlg_fg, bg=dlg_bg).pack(side='left')
            s_entry.pack(side='left', padx=4)
            tk.Label(box, text='S', fg=dlg_fg, bg=dlg_bg).pack(side='left')

            # populate with first configured value if present
            try:
                ssettings = load_settings()
                a = (ssettings.get('auto_hold_times') or [])
                if a:
                    secs = int(a[0])
                    hh = secs // 3600; mm = (secs % 3600) // 60; ss = secs % 60
                    h_entry.insert(0, str(hh)); m_entry.insert(0, str(mm)); s_entry.insert(0, str(ss))
            except Exception:
                pass

            def do_save():
                try:
                    hh = int(h_entry.get() or 0)
                    mm = int(m_entry.get() or 0)
                    ss = int(s_entry.get() or 0)
                    total = max(0, hh*3600 + mm*60 + ss)
                except Exception:
                    total = 0
                try:
                    ssettings = load_settings()
                    # replace with single auto-hold time (list with one element)
                    ssettings['auto_hold_times'] = [int(total)] if total > 0 else []
                    save_settings(ssettings)
                    # update runtime set so this run will consider the new value
                    self._auto_hold_triggered = set()
                except Exception:
                    pass
                dlg.destroy()

            btnf = tk.Frame(dlg, bg=dlg_bg)
            btnf.pack(fill='x', pady=6)
            tk.Button(btnf, text='Save', command=do_save, width=10, bg=btn_bg, fg=btn_fg, activebackground='#444').pack(side='right', padx=6)
            tk.Button(btnf, text='Cancel', command=dlg.destroy, width=10, bg=btn_bg, fg=btn_fg, activebackground='#444').pack(side='right')

            # recursively theme dialog to ensure consistency
            try:
                self._theme_recursive(dlg, dlg_bg, dlg_fg, btn_bg, btn_fg)
            except Exception:
                pass

        tk.Button(frame_duration, text='Auto-hold...', command=open_autohold_dialog, fg='white', bg='#333', width=12).pack(side='left', padx=8)

        # Clock time input (separate HH, MM, SS boxes)
        frame_clock = tk.Frame(root, bg="black")
        frame_clock.pack(pady=5)
        tk.Label(frame_clock, text="Clock (HH:MM:SS)", fg="white", bg="black").pack(side="left")
        self.clock_hours_entry = tk.Entry(frame_clock, width=3, font=("Arial", 18), fg='white', bg='#111', insertbackground='white')
        self.clock_hours_entry.insert(0, "14")
        self.clock_hours_entry.pack(side="left", padx=2)
        tk.Label(frame_clock, text=":", fg="white", bg="black").pack(side="left")
        self.clock_minutes_entry = tk.Entry(frame_clock, width=3, font=("Arial", 18), fg='white', bg='#111', insertbackground='white')
        self.clock_minutes_entry.insert(0, "00")
        self.clock_minutes_entry.pack(side="left", padx=2)
        tk.Label(frame_clock, text=":", fg="white", bg="black").pack(side="left")
        self.clock_seconds_entry = tk.Entry(frame_clock, width=3, font=("Arial", 18), fg='white', bg='#111', insertbackground='white')
        self.clock_seconds_entry.insert(0, "00")
        self.clock_seconds_entry.pack(side="left", padx=2)

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
        self.settings_btn = tk.Button(frame_buttons, text="Settings", command=self.show_settings_window, font=("Arial", 14), width=10)
        self.settings_btn.grid(row=0, column=4, padx=6)

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
        # Apply appearance settings at startup so the mission entry and other widgets reflect the saved mode
        try:
            self.apply_appearance_settings()
        except Exception:
            pass
        self.update_clock()

    # ----------------------------
    # Settings window
    # ----------------------------
    def show_settings_window(self):
        settings = load_settings()
        win = tk.Toplevel(self.root)
        win.transient(self.root)
        win.title("Settings")
        win.geometry("560x275")
        # apply current appearance mode so the settings window matches the main UI
        s_local = load_settings()
        mode_local = s_local.get('appearance_mode', 'dark')
        if mode_local == 'dark':
            win_bg = '#000000'; win_text = '#FFFFFF'; btn_bg = '#FFFFFF'; btn_fg = '#000000'
        else:
            win_bg = '#FFFFFF'; win_text = '#000000'; btn_bg = '#000000'; btn_fg = '#FFFFFF'
        win.config(bg=win_bg)
        # set per-window widget defaults so nested widgets inherit the chosen theme
        try:
            win.option_add('*Foreground', win_text)
            win.option_add('*Background', win_bg)
            # entry specific defaults
            win.option_add('*Entry.Background', '#222' if mode_local == 'dark' else '#b4b4b4')
            win.option_add('*Entry.Foreground', win_text if mode_local == 'dark' else '#000000')
        except Exception:
            pass
        # keep track of this Toplevel so other dialogs can close it if needed
        try:
            self.settings_win = win
            def _clear_settings_ref(evt=None):
                try:
                    self.settings_win = None
                except Exception:
                    pass
            win.bind('<Destroy>', _clear_settings_ref)
        except Exception:
            pass

        # Mode selection
        frame_mode = tk.Frame(win)
        frame_mode.config(bg=win_bg)
        frame_mode.pack(fill='x', pady=8, padx=8)
        tk.Label(frame_mode, text="Mode:", fg=win_text, bg=win_bg).pack(side='left')
        mode_var = tk.StringVar(value=settings.get('mode', 'spreadsheet'))
        tk.Radiobutton(frame_mode, text='Spreadsheet', variable=mode_var, value='spreadsheet', fg=win_text, bg=win_bg, selectcolor=win_bg).pack(side='left', padx=8)
        tk.Radiobutton(frame_mode, text='Buttons (manual)', variable=mode_var, value='buttons', fg=win_text, bg=win_bg, selectcolor=win_bg).pack(side='left', padx=8)

        # Spreadsheet config
        frame_sheet = tk.LabelFrame(win, text='Spreadsheet configuration', fg=win_text, bg=win_bg)
        frame_sheet.config(bg=win_bg)
        frame_sheet.pack(fill='x', padx=8, pady=6)
        tk.Label(frame_sheet, text='Sheet link (CSV export):', fg=win_text, bg=win_bg).pack(anchor='w')
        # entry background chosen to contrast with window background
        sheet_entry_bg = '#222' if mode_local == 'dark' else '#b4b4b4'
        sheet_entry_fg = win_text if mode_local == 'dark' else '#000000'
        sheet_entry = tk.Entry(frame_sheet, width=80, fg=sheet_entry_fg, bg=sheet_entry_bg, insertbackground=sheet_entry_fg)
        sheet_entry.pack(fill='x', padx=6, pady=4)
        sheet_entry.insert(0, settings.get('sheet_link', SHEET_LINK))

        # Accept cells in 'L3' format for each parameter
        cell_frame = tk.Frame(frame_sheet)
        cell_frame.config(bg=win_bg)
        cell_frame.pack(fill='x', padx=6, pady=2)
        tk.Label(cell_frame, text='Range cell (e.g. L3):', fg=win_text, bg=win_bg).grid(row=0, column=0)
        range_cell_bg = '#222' if mode_local == 'dark' else '#b4b4b4'
        range_cell_fg = win_text if mode_local == 'dark' else '#000000'
        range_cell = tk.Entry(cell_frame, width=8, fg=range_cell_fg, bg=range_cell_bg, insertbackground=range_cell_fg)
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

        tk.Label(cell_frame, text='Weather cell (e.g. L4):', fg=win_text, bg=win_bg).grid(row=0, column=2)
        weather_cell = tk.Entry(cell_frame, width=8, fg=range_cell_fg, bg=range_cell_bg, insertbackground=range_cell_fg)
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

        tk.Label(cell_frame, text='Vehicle cell (e.g. L5):', fg=win_text, bg=win_bg).grid(row=0, column=4)
        vehicle_cell = tk.Entry(cell_frame, width=8, fg=range_cell_fg, bg=range_cell_bg, insertbackground=range_cell_fg)
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
        frame_buttons_cfg = tk.LabelFrame(win, text='Manual Go/No-Go (Buttons mode)', fg=win_text, bg=win_bg)
        frame_buttons_cfg.config(bg=win_bg)
        frame_buttons_cfg.pack(fill='x', padx=8, pady=6)

        # (Auto-hold configuration removed from Settings — managed from main UI)

        # Appearance settings are in a separate window
        frame_appearance_btn = tk.Frame(win, bg=win_bg)
        frame_appearance_btn.pack(fill='x', padx=8, pady=6)
        tk.Button(frame_appearance_btn, text='Appearance...', command=lambda: self.show_appearance_window(), fg=btn_fg, bg=btn_bg, activebackground='#444').pack(side='left')

        # Timezone selector
        tz_frame = tk.Frame(frame_sheet, bg=win_bg)
        tz_frame.pack(fill='x', padx=6, pady=4)
        tk.Label(tz_frame, text='Timezone:', fg=win_text, bg=win_bg).pack(side='left')
        tz_var = tk.StringVar(value=settings.get('timezone', DEFAULT_SETTINGS.get('timezone', 'local')))
        # OptionMenu with a few choices, but user may edit the text to any IANA name
        tz_menu = tk.OptionMenu(tz_frame, tz_var, *TIMEZONE_CHOICES)
        tz_menu.config(fg=win_text, bg=range_cell_bg, activebackground='#333')
        tz_menu.pack(side='left', padx=6)

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
                'manual_vehicle': getattr(fetch_gonogo, 'manual_vehicle', None),
                'timezone': tz_var.get(),
                # preserve appearance settings (edited in Appearance window)
                'bg_color': settings.get('bg_color', '#000000'),
                'text_color': settings.get('text_color', '#FFFFFF'),
                'gn_bg_color': settings.get('gn_bg_color', '#111111'),
                'gn_border_color': settings.get('gn_border_color', '#FFFFFF'),
                'gn_go_color': settings.get('gn_go_color', '#00FF00'),
                'gn_nogo_color': settings.get('gn_nogo_color', '#FF0000'),
                'font_family': settings.get('font_family', 'Consolas'),
                'mission_font_px': int(settings.get('mission_font_px', 48)),
                'timer_font_px': int(settings.get('timer_font_px', 120)),
                'gn_font_px': int(settings.get('gn_font_px', 28))
            }
            # Auto-hold editing removed from Settings window; keep existing settings value
            new_settings['auto_hold_times'] = settings.get('auto_hold_times', [])
            # preserve the appearance_mode so saving Settings doesn't accidentally remove it
            try:
                new_settings['appearance_mode'] = settings.get('appearance_mode', DEFAULT_SETTINGS.get('appearance_mode', 'dark'))
            except Exception:
                new_settings['appearance_mode'] = DEFAULT_SETTINGS.get('appearance_mode', 'dark')
            save_settings(new_settings)
            # update immediately
            self.gonogo_values = fetch_gonogo()
            write_gonogo_html(self.gonogo_values)
            # update manual visibility in main UI
            self.update_manual_visibility()
            # appearance changes are applied only from the Appearance window
            win.destroy()

        def on_cancel():
            win.destroy()

        btn_frame = tk.Frame(win, bg=win_bg)
        btn_frame.pack(fill='x', pady=8)
        tk.Button(btn_frame, text='Save', command=on_save, fg=btn_fg, bg=btn_bg, activebackground='#444').pack(side='right', padx=8)
        tk.Button(btn_frame, text='Cancel', command=on_cancel, fg=btn_fg, bg=btn_bg, activebackground='#444').pack(side='right')
        # ensure the new toplevel gets recursively themed like the main window
        try:
            self._theme_recursive(win, win_bg, win_text, btn_bg, btn_fg)
        except Exception:
            pass


    # ----------------------------
    # Update input visibility based on mode
    # ----------------------------
    def update_inputs(self):
        if self.mode_var.get() == "duration":
            self.hours_entry.config(state="normal")
            self.minutes_entry.config(state="normal")
            self.seconds_entry.config(state="normal")
            self.clock_hours_entry.config(state="disabled")
            self.clock_minutes_entry.config(state="disabled")
            self.clock_seconds_entry.config(state="disabled")
        else:
            self.hours_entry.config(state="disabled")
            self.minutes_entry.config(state="disabled")
            self.seconds_entry.config(state="disabled")
            self.clock_hours_entry.config(state="normal")
            self.clock_minutes_entry.config(state="normal")
            self.clock_seconds_entry.config(state="normal")

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

    def apply_appearance_settings(self):
        """Apply appearance-related settings to the running Tk UI."""
        s = load_settings()
        # If an appearance_mode preset is selected, override specific settings with the preset
        mode = s.get('appearance_mode', None)
        if mode == 'dark':
            s.update({
                'bg_color': '#000000', 'text_color': '#FFFFFF', 'gn_bg_color': '#111111',
                'gn_border_color': '#FFFFFF', 'gn_go_color': '#00FF00', 'gn_nogo_color': '#FF0000',
                'font_family': 'Consolas', 'mission_font_px': 44, 'timer_font_px': 80, 'gn_font_px': 24
            })
        elif mode == 'light':
            s.update({
                'bg_color': '#FFFFFF', 'text_color': '#000000', 'gn_bg_color': '#EEEEEE',
                'gn_border_color': '#333333', 'gn_go_color': '#008800', 'gn_nogo_color': '#AA0000',
                'font_family': 'Consolas', 'mission_font_px': 44, 'timer_font_px': 80, 'gn_font_px': 24
            })
        bg = s.get('bg_color', '#000000')
        text = s.get('text_color', '#FFFFFF')
        font_family = s.get('font_family', 'Consolas')
        timer_px = int(s.get('timer_font_px', 100))
        mission_px = int(s.get('mission_font_px', 48))
        gn_px = int(s.get('gn_font_px', 24))
        gn_bg = s.get('gn_bg_color', '#111111')
        gn_border = s.get('gn_border_color', '#FFFFFF')
        gn_go = s.get('gn_go_color', '#00FF00')
        gn_nogo = s.get('gn_nogo_color', '#FF0000')
        # apply to main window elements
        try:
            self.root.config(bg=bg)
            self.titletext.config(fg=text, bg=bg, font=(font_family, 20))
            # timer label
            self.text.config(fg=text, bg=bg, font=(font_family, timer_px, 'bold'))
            # GN labels: set bg and font, and color depending on GO/NOGO
            def style_gn_label(lbl, value):
                try:
                    lbl.config(bg=bg, font=(font_family, gn_px))
                    v = (value or '').strip().upper()
                    if v == 'GO':
                        lbl.config(fg=gn_go)
                    elif v in ('NOGO', 'NO-GO'):
                        lbl.config(fg=gn_nogo)
                    else:
                        lbl.config(fg=text)
                except Exception:
                    pass

            style_gn_label(self.range_label, getattr(self, 'range_status', None))
            style_gn_label(self.weather_label, getattr(self, 'weather', None))
            style_gn_label(self.vehicle_label, getattr(self, 'vehicle', None))

            # Buttons: invert colors depending on mode
            # dark mode -> buttons white bg, black text
            # light mode -> buttons black bg, white text
            if mode == 'dark':
                btn_bg = '#FFFFFF'
                btn_fg = '#000000'
                active_bg = '#DDDDDD'
            else:
                btn_bg = '#000000'
                btn_fg = '#FFFFFF'
                active_bg = '#222222'

            for btn in (self.start_btn, self.hold_btn, self.resume_btn, self.scrub_btn, self.reset_btn, self.settings_btn):
                try:
                    # preserve scrub button's custom color (red) if set
                    try:
                        cur_fg = btn.cget('fg')
                    except Exception:
                        cur_fg = None
                    if btn is getattr(self, 'scrub_btn', None) and cur_fg:
                        # keep existing foreground (usually red)
                        btn.config(bg=btn_bg, activebackground=active_bg)
                    else:
                        btn.config(bg=btn_bg, fg=btn_fg, activebackground=active_bg)
                except Exception:
                    pass

            # Manual toggle buttons
            for btn in (self.range_toggle_btn, self.weather_toggle_btn, self.vehicle_toggle_btn):
                try:
                    btn.config(bg=btn_bg, fg=btn_fg)
                except Exception:
                    pass

            # manual frame and footer
            try:
                self.manual_frame.config(bg=bg)
                # Footer should invert colors depending on mode:
                # - dark mode -> white background, black text
                # - light mode -> black background, white text
                mode = s.get('appearance_mode', 'dark')
                if mode == 'dark':
                    footer_bg = '#FFFFFF'
                    footer_fg = '#000000'
                else:
                    footer_bg = '#000000'
                    footer_fg = '#FFFFFF'
                try:
                    self.footer_label.config(bg=footer_bg, fg=footer_fg)
                except Exception:
                    # fall back to generic theme
                    self.footer_label.config(bg=bg, fg=text)
            except Exception:
                pass
        except Exception:
            pass
        # Recursively theme frames and common widgets so no frame is left with old colors
        try:
            self._theme_recursive(self.root, bg, text, btn_bg, btn_fg)
        except Exception:
            pass

    def update_gn_labels(self, range_val, weather_val, vehicle_val):
        """Update GN label texts and apply theme-aware styling."""
        s = load_settings()
        gn_px = int(s.get('gn_font_px', 28))
        font_family = s.get('font_family', 'Consolas')
        bg = s.get('bg_color', '#000000')
        text = s.get('text_color', '#FFFFFF')
        gn_go = s.get('gn_go_color', '#00FF00')
        gn_nogo = s.get('gn_nogo_color', '#FF0000')
        # Range
        try:
            display_range = format_status_display(range_val)
            self.range_label.config(text=f"RANGE: {display_range}", bg=bg, font=(font_family, gn_px))
            rv = (range_val or '').strip().upper()
            rnorm = re.sub(r'[^A-Z]', '', rv)
            if rnorm == 'GO':
                self.range_label.config(fg=gn_go)
            elif rnorm == 'NOGO':
                self.range_label.config(fg=gn_nogo)
            else:
                self.range_label.config(fg=text)
        except Exception:
            pass

        # Weather
        try:
            display_weather = format_status_display(weather_val)
            self.weather_label.config(text=f"WEATHER: {display_weather}", bg=bg, font=(font_family, gn_px))
            wv = (weather_val or '').strip().upper()
            wnorm = re.sub(r'[^A-Z]', '', wv)
            if wnorm == 'GO':
                self.weather_label.config(fg=gn_go)
            elif wnorm == 'NOGO':
                self.weather_label.config(fg=gn_nogo)
            else:
                self.weather_label.config(fg=text)
        except Exception:
            pass

        # Vehicle
        try:
            display_vehicle = format_status_display(vehicle_val)
            self.vehicle_label.config(text=f"VEHICLE: {display_vehicle}", bg=bg, font=(font_family, gn_px))
            vv = (vehicle_val or '').strip().upper()
            vnorm = re.sub(r'[^A-Z]', '', vv)
            if vnorm == 'GO':
                self.vehicle_label.config(fg=gn_go)
            elif vnorm == 'NOGO':
                self.vehicle_label.config(fg=gn_nogo)
            else:
                self.vehicle_label.config(fg=text)
        except Exception:
            pass

    def _theme_recursive(self, widget, bg, text, btn_bg, btn_fg):
        # load settings so we can theme GN label backgrounds if configured
        s = load_settings()
        for child in widget.winfo_children():
            # Frame and LabelFrame
            try:
                if isinstance(child, (tk.Frame, tk.LabelFrame)):
                    try:
                        child.config(bg=bg)
                    except Exception:
                        pass
                # Labels: set bg, but don't override GN label fg
                if isinstance(child, tk.Label):
                    try:
                        # preserve GN label fg colors and don't override the footer label (it has a special inverted style)
                        if child in (getattr(self, 'range_label', None), getattr(self, 'weather_label', None), getattr(self, 'vehicle_label', None)):
                            # GN labels keep fg but should have themed bg
                            child.config(bg=s.get('gn_bg_color', bg))
                        elif child is getattr(self, 'footer_label', None):
                            # footer_label was already styled by apply_appearance_settings; don't override it here
                            pass
                        else:
                            child.config(bg=bg, fg=text)
                    except Exception:
                        pass
                # Entries: avoid overriding entries that were explicitly styled (like mission_entry)
                if isinstance(child, tk.Entry):
                    try:
                        # Set entry bg/fg depending on appearance mode
                        mode_local = s.get('appearance_mode', 'dark')
                        if mode_local == 'dark':
                            child.config(bg='#222222', fg=text, insertbackground=text)
                        else:
                            # light mode entries should contrast with the white background
                            child.config(bg='#b4b4b4', fg='#000000', insertbackground='#000000')
                    except Exception:
                        pass
                # OptionMenu/Menubutton
                if isinstance(child, tk.Menubutton):
                    try:
                        child.config(bg=btn_bg, fg=btn_fg, activebackground='#555')
                    except Exception:
                        pass
                # Radiobutton / Checkbutton
                if isinstance(child, (tk.Radiobutton, tk.Checkbutton)):
                    try:
                        # selectcolor is the indicator background; set it to match the overall bg for neatness
                        child.config(bg=bg, fg=text, selectcolor=bg, activebackground=bg)
                    except Exception:
                        pass
                # Buttons: ensure themed background and correct fg
                if isinstance(child, tk.Button):
                    try:
                        # don't override scrub button's fg if it has a special color
                        if child is getattr(self, 'scrub_btn', None):
                            child.config(bg=btn_bg, activebackground='#555')
                        else:
                            child.config(bg=btn_bg, fg=btn_fg, activebackground='#555')
                    except Exception:
                        pass
            except Exception:
                pass
            # Recurse
            try:
                if hasattr(child, 'winfo_children'):
                    self._theme_recursive(child, bg, text, btn_bg, btn_fg)
            except Exception:
                pass

    def show_appearance_window(self):
        # Re-implemented appearance window with proper theming and layout
        settings = load_settings()
        win = tk.Toplevel(self.root)
        win.transient(self.root)
        win.title('Appearance')
        win.geometry('520x475')

        # derive colors from appearance_mode so the dialog matches the main UI
        mode_local = settings.get('appearance_mode', 'dark')
        if mode_local == 'dark':
            win_bg = '#000000'; win_text = '#FFFFFF'; btn_bg = '#FFFFFF'; btn_fg = '#000000'; entry_bg = '#222'; entry_fg = '#FFFFFF'
        else:
            win_bg = '#FFFFFF'; win_text = '#000000'; btn_bg = '#000000'; btn_fg = '#FFFFFF'; entry_bg = '#b4b4b4'; entry_fg = '#000000'
        win.config(bg=win_bg)

        tk.Label(win, text='Choose UI mode:', fg=win_text, bg=win_bg).pack(anchor='w', padx=12, pady=(10,0))
        mode_var = tk.StringVar(value=settings.get('appearance_mode', 'dark'))
        modes = ['dark', 'light']
        mode_menu = tk.OptionMenu(win, mode_var, *modes)
        mode_menu.config(fg=win_text, bg=entry_bg, activebackground='#333')
        mode_menu.pack(anchor='w', padx=12, pady=6)

        def on_save_mode():
            choice = mode_var.get()
            presets = {
                'dark': {
                    'bg_color': '#000000', 'text_color': '#FFFFFF', 'gn_bg_color': '#111111',
                    'gn_border_color': '#FFFFFF', 'gn_go_color': '#00FF00', 'gn_nogo_color': '#FF0000',
                    'font_family': 'Consolas', 'mission_font_px': 24, 'timer_font_px': 80, 'gn_font_px': 20
                },
                'light': {
                    'bg_color': '#FFFFFF', 'text_color': '#000000', 'gn_bg_color': '#EEEEEE',
                    'gn_border_color': '#333333', 'gn_go_color': '#008800', 'gn_nogo_color': '#AA0000',
                    'font_family': 'Consolas', 'mission_font_px': 24, 'timer_font_px': 80, 'gn_font_px': 20
                }
            }
            p = presets.get(choice, {})
            s = load_settings()
            s['appearance_mode'] = choice
            s.update(p)
            save_settings(s)
            try:
                self.apply_appearance_settings()
                write_countdown_html(self.mission_name, self.text.cget('text'))
                write_gonogo_html(self.gonogo_values)
            except Exception:
                pass
            # close appearance window
            win.destroy()
            # also close the settings window if it is open
            try:
                if getattr(self, 'settings_win', None):
                    try:
                        self.settings_win.destroy()
                    except Exception:
                        pass
            except Exception:
                pass

        def choose_color(entry_widget):
            try:
                col = colorchooser.askcolor()
                if col and col[1]:
                    entry_widget.delete(0, tk.END)
                    entry_widget.insert(0, col[1])
            except Exception:
                pass

        s = load_settings()
        html_frame = tk.LabelFrame(win, text='HTML appearance (streaming)', fg=win_text, bg=win_bg)
        html_frame.config(bg=win_bg)
        html_frame.pack(fill='x', padx=8, pady=6)

        # layout HTML appearance fields in a grid
        tk.Label(html_frame, text='Background:', fg=win_text, bg=win_bg).grid(row=0, column=0, sticky='w', padx=6, pady=4)
        bg_entry = tk.Entry(html_frame, width=12, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        bg_entry.grid(row=0, column=1, padx=6, pady=4)
        bg_entry.insert(0, s.get('html_bg_color', s.get('bg_color', '#000000')))
        tk.Button(html_frame, text='Choose', command=lambda: choose_color(bg_entry), fg=btn_fg, bg=btn_bg).grid(row=0, column=2, padx=6)

        tk.Label(html_frame, text='Text:', fg=win_text, bg=win_bg).grid(row=1, column=0, sticky='w', padx=6, pady=4)
        text_entry = tk.Entry(html_frame, width=12, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        text_entry.grid(row=1, column=1, padx=6, pady=4)
        text_entry.insert(0, s.get('html_text_color', s.get('text_color', '#FFFFFF')))
        tk.Button(html_frame, text='Choose', command=lambda: choose_color(text_entry), fg=btn_fg, bg=btn_bg).grid(row=1, column=2, padx=6)

        tk.Label(html_frame, text='GN GO:', fg=win_text, bg=win_bg).grid(row=2, column=0, sticky='w', padx=6, pady=4)
        gn_go_entry = tk.Entry(html_frame, width=12, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        gn_go_entry.grid(row=2, column=1, padx=6, pady=4)
        gn_go_entry.insert(0, s.get('html_gn_go_color', s.get('gn_go_color', '#00FF00')))
        tk.Button(html_frame, text='Choose', command=lambda: choose_color(gn_go_entry), fg=btn_fg, bg=btn_bg).grid(row=2, column=2, padx=6)

        tk.Label(html_frame, text='GN NO-GO:', fg=win_text, bg=win_bg).grid(row=3, column=0, sticky='w', padx=6, pady=4)
        gn_nogo_entry = tk.Entry(html_frame, width=12, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        gn_nogo_entry.grid(row=3, column=1, padx=6, pady=4)
        gn_nogo_entry.insert(0, s.get('html_gn_nogo_color', s.get('gn_nogo_color', '#FF0000')))
        tk.Button(html_frame, text='Choose', command=lambda: choose_color(gn_nogo_entry), fg=btn_fg, bg=btn_bg).grid(row=3, column=2, padx=6)

        tk.Label(html_frame, text='GN box bg:', fg=win_text, bg=win_bg).grid(row=4, column=0, sticky='w', padx=6, pady=4)
        gn_box_bg_entry = tk.Entry(html_frame, width=12, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        gn_box_bg_entry.grid(row=4, column=1, padx=6, pady=4)
        gn_box_bg_entry.insert(0, s.get('html_gn_bg_color', s.get('gn_bg_color', '#111111')))
        tk.Button(html_frame, text='Choose', command=lambda: choose_color(gn_box_bg_entry), fg=btn_fg, bg=btn_bg).grid(row=4, column=2, padx=6)

        tk.Label(html_frame, text='GN border:', fg=win_text, bg=win_bg).grid(row=5, column=0, sticky='w', padx=6, pady=4)
        gn_border_entry = tk.Entry(html_frame, width=12, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        gn_border_entry.grid(row=5, column=1, padx=6, pady=4)
        gn_border_entry.insert(0, s.get('html_gn_border_color', s.get('gn_border_color', '#FFFFFF')))
        tk.Button(html_frame, text='Choose', command=lambda: choose_color(gn_border_entry), fg=btn_fg, bg=btn_bg).grid(row=5, column=2, padx=6)

        tk.Label(html_frame, text='Font family:', fg=win_text, bg=win_bg).grid(row=6, column=0, sticky='w', padx=6, pady=4)
        font_entry = tk.Entry(html_frame, width=20, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        font_entry.grid(row=6, column=1, padx=6, pady=4, columnspan=2, sticky='w')
        font_entry.insert(0, s.get('html_font_family', s.get('font_family', 'Consolas')))

        tk.Label(html_frame, text='Mission px:', fg=win_text, bg=win_bg).grid(row=7, column=0, sticky='w', padx=6, pady=4)
        mission_px_entry = tk.Entry(html_frame, width=6, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        mission_px_entry.grid(row=7, column=1, padx=6, pady=4, sticky='w')
        mission_px_entry.insert(0, str(s.get('html_mission_font_px', s.get('mission_font_px', 24))))

        tk.Label(html_frame, text='Timer px:', fg=win_text, bg=win_bg).grid(row=8, column=0, sticky='w', padx=6, pady=4)
        timer_px_entry = tk.Entry(html_frame, width=6, fg=entry_fg, bg=entry_bg, insertbackground=entry_fg)
        timer_px_entry.grid(row=8, column=1, padx=6, pady=4, sticky='w')
        timer_px_entry.insert(0, str(s.get('html_timer_font_px', s.get('timer_font_px', 80))))

        # Add a checkbox to hide mission name in HTML output
        self.hide_mission_name_var = tk.BooleanVar(value=s.get("hide_mission_name", False))
        hide_mission_name_cb = tk.Checkbutton(html_frame, text="Hide mission name in HTML output", variable=self.hide_mission_name_var, fg=win_text, bg=win_bg, selectcolor=win_bg, activebackground=win_bg)

        hide_mission_name_cb.grid(row=9, column=0, columnspan=3, sticky='w', padx=6, pady=4)
        
        def save_html_prefs():
            try:
                s_local = load_settings()
                s_local['html_bg_color'] = bg_entry.get().strip() or s_local.get('html_bg_color')
                s_local['html_text_color'] = text_entry.get().strip() or s_local.get('html_text_color')
                s_local['html_gn_go_color'] = gn_go_entry.get().strip() or s_local.get('html_gn_go_color')
                s_local['html_gn_nogo_color'] = gn_nogo_entry.get().strip() or s_local.get('html_gn_nogo_color')
                s_local['html_gn_bg_color'] = gn_box_bg_entry.get().strip() or s_local.get('html_gn_bg_color')
                s_local['html_gn_border_color'] = gn_border_entry.get().strip() or s_local.get('html_gn_border_color')
                s_local['html_font_family'] = font_entry.get().strip() or s_local.get('html_font_family')
                s_local["hide_mission_name"] = self.hide_mission_name_var.get()
                try:
                    s_local['html_mission_font_px'] = int(mission_px_entry.get())
                except Exception:
                    pass
                try:
                    s_local['html_timer_font_px'] = int(timer_px_entry.get())
                except Exception:
                    pass
                save_settings(s_local)
                write_countdown_html(self.mission_name, self.text.cget('text'))
                write_gonogo_html(self.gonogo_values)
            except Exception:
                pass

        def reset_html_defaults():
            try:
                s_local = load_settings()
                s_local['html_bg_color'] = DEFAULT_SETTINGS.get('html_bg_color')
                s_local['html_text_color'] = DEFAULT_SETTINGS.get('html_text_color')
                s_local['html_font_family'] = DEFAULT_SETTINGS.get('html_font_family')
                s_local['html_mission_font_px'] = DEFAULT_SETTINGS.get('html_mission_font_px')
                s_local['html_timer_font_px'] = DEFAULT_SETTINGS.get('html_timer_font_px')
                s_local['html_gn_bg_color'] = DEFAULT_SETTINGS.get('html_gn_bg_color')
                s_local['html_gn_border_color'] = DEFAULT_SETTINGS.get('html_gn_border_color')
                s_local['html_gn_go_color'] = DEFAULT_SETTINGS.get('html_gn_go_color')
                s_local['html_gn_nogo_color'] = DEFAULT_SETTINGS.get('html_gn_nogo_color')
                s_local['html_gn_font_px'] = DEFAULT_SETTINGS.get('html_gn_font_px')
                save_settings(s_local)
                # update UI fields
                bg_entry.delete(0, tk.END); bg_entry.insert(0, s_local['html_bg_color'])
                text_entry.delete(0, tk.END); text_entry.insert(0, s_local['html_text_color'])
                gn_go_entry.delete(0, tk.END); gn_go_entry.insert(0, s_local['html_gn_go_color'])
                gn_nogo_entry.delete(0, tk.END); gn_nogo_entry.insert(0, s_local['html_gn_nogo_color'])
                gn_box_bg_entry.delete(0, tk.END); gn_box_bg_entry.insert(0, s_local['html_gn_bg_color'])
                gn_border_entry.delete(0, tk.END); gn_border_entry.insert(0, s_local['html_gn_border_color'])
                font_entry.delete(0, tk.END); font_entry.insert(0, s_local['html_font_family'])
                mission_px_entry.delete(0, tk.END); mission_px_entry.insert(0, str(s_local['html_mission_font_px']))
                timer_px_entry.delete(0, tk.END); timer_px_entry.insert(0, str(s_local['html_timer_font_px']))
                write_countdown_html(self.mission_name, self.text.cget('text'))
                write_gonogo_html(self.gonogo_values)
            except Exception:
                pass

        html_btns = tk.Frame(html_frame, bg=win_bg)
        html_btns.grid(row=10, column=0, columnspan=3, pady=6)
        tk.Button(html_btns, text='Save (HTML only)', command=save_html_prefs, fg=btn_fg, bg=btn_bg).pack(side='right', padx=6)
        tk.Button(html_btns, text='Reset HTML defaults', command=reset_html_defaults, fg=btn_fg, bg=btn_bg).pack(side='right')

        btn_frame = tk.Frame(win, bg=win_bg)
        btn_frame.pack(fill='x', pady=8, padx=8)
        tk.Button(btn_frame, text='Save', command=on_save_mode, fg=btn_fg, bg=btn_bg).pack(side='right', padx=6)
        tk.Button(btn_frame, text='Cancel', command=win.destroy, fg=btn_fg, bg=btn_bg).pack(side='right')

        try:
            self._theme_recursive(win, win_bg, win_text, btn_bg, btn_fg)
        except Exception:
            pass

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
        new_val = 'NO-GO' if cur_val == 'GO' else 'GO'
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
                # read separate HH, MM, SS boxes
                h = int(self.clock_hours_entry.get() or 0)
                m = int(self.clock_minutes_entry.get() or 0)
                s = int(self.clock_seconds_entry.get() or 0)
                # determine timezone from settings
                ssettings = load_settings()
                tzname = ssettings.get('timezone', DEFAULT_SETTINGS.get('timezone', 'local'))
                if ZoneInfo is None or tzname in (None, '', 'local'):
                    # naive local time handling (existing behavior) — use timedelta to roll day
                    target_today = now.replace(hour=h, minute=m, second=s, microsecond=0)
                    if target_today <= now:
                        target_today = target_today + timedelta(days=1)
                    total_seconds = (target_today - now).total_seconds()
                else:
                    try:
                        tz = ZoneInfo(tzname)
                        # construct aware "now" in that timezone and create the target time
                        now_tz = datetime.now(tz)
                        target = now_tz.replace(hour=h, minute=m, second=s, microsecond=0)
                        # if target already passed in that tz, roll to next day
                        if target <= now_tz:
                            target = target + timedelta(days=1)
                        # compute total seconds using aware-datetime subtraction to avoid epoch mixing
                        total_seconds = (target - now_tz).total_seconds()
                    except Exception:
                        # fallback to naive local behavior
                        target_today = now.replace(hour=h, minute=m, second=s, microsecond=0)
                        if target_today <= now:
                            target_today = target_today + timedelta(days=1)
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
                # auto-hold detection: if configured times include this remaining value, enter hold
                try:
                    s = load_settings()
                    ah = set(int(x) for x in s.get('auto_hold_times', []) or [])
                except Exception:
                    ah = set()
                if diff in ah and diff not in getattr(self, '_auto_hold_triggered', set()):
                    # trigger hold
                    self._auto_hold_triggered.add(diff)
                    self.hold()
                    # show_hold_button/other UI changes handled by hold()
                    # After entering hold, update countdown display via next tick
                
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
            # update texts and styles using theme
            try:
                self.update_gn_labels(self.range_status, self.weather, self.vehicle)
            except Exception:
                # fallback to simple config
                self.range_label.config(text=f"RANGE: {self.range_status}")
                self.weather_label.config(text=f"WEATHER: {self.weather}")
                self.vehicle_label.config(text=f"VEHICLE: {self.vehicle}")
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

        # Footer uses inverted colors: white bg/black text in dark mode, black bg/white text in light mode
        s = load_settings()
        splash_mode = s.get('appearance_mode', 'dark')
        if splash_mode == 'dark':
            splash_footer_bg = '#FFFFFF'
            splash_footer_fg = '#000000'
        else:
            splash_footer_bg = '#000000'
            splash_footer_fg = '#FFFFFF'

        footer_label = tk.Label(
            footer_frame,
            text="Made by HamsterSpaceNerd3000",
            font=("Consolas", 12),
            fg=splash_footer_fg,
            bg=splash_footer_bg
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
                    # show a visible countdown before auto-start; allow Continue to skip
                    AUTO_START_SECONDS = 5
                    remaining = AUTO_START_SECONDS
                    cont_btn.config(state='normal')

                    def tick():
                        nonlocal remaining
                        if remaining <= 0:
                            on_continue()
                            return
                        info.config(text=f"Ready — auto-starting in {remaining}...")
                        cont_btn.config(text=f"Continue ({remaining})")
                        remaining -= 1
                        splash.after(1000, tick)

                    # clicking Continue will immediately proceed
                    cont_btn.config(command=on_continue)
                    tick()
                return
            splash.after(200, check_init)

        def on_continue():
            splash.destroy()
            # now create the real main window
            root = tk.Tk()
            app = CountdownApp(root)
            root.mainloop()

        # begin polling
        splash.after(100, check_init)
        splash.mainloop()

    show_splash_and_start()