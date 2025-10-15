import tkinter as tk
import time
import threading
from datetime import datetime
import requests
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
# Fetch (CSV) Go/No-Go (fallback/prefill)
# -------------------------
def fetch_gonogo_csv(csv_link=DEFAULT_CSV_LINK, timeout=3):
    """
    Fetch Go/No-Go parameters from CSV export (rows 2-4, col 12).
    Returns [Range, Weather, Vehicle] (strings).
    """
    try:
        resp = session.get(csv_link, timeout=timeout)
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
        print(f"[ERROR] Failed to fetch Go/No-Go CSV: {e}")
        return ["N/A", "N/A", "N/A"]

# -------------------------
# Utility
# -------------------------
def get_status_color(status):
    status = (status or "").strip().upper()
    if status == "GO": return "green"
    if status == "NOGO" or status == "NOGO" or status == "NOGO": return "red"
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
setTimeout(()=>location.reload(),5000);
</script>
</head>
<body>
<div id="gonogo">
    <div class="status-box {'go' if vals[0].strip().upper()=='GO' else 'nogo'}">Range: {vals[0]}</div>
    <div class="status-box {'go' if vals[2].strip().upper()=='GO' else 'nogo'}">Vehicle: {vals[2]}</div>
    <div class="status-box {'go' if vals[1].strip().upper()=='GO' else 'nogo'}">Weather: {vals[1]}</div>
</div>
</body>
</html>"""
    with open(GONOGO_HTML, "w", encoding="utf-8") as f:
        f.write(html)

def write_gonogo_html_iframe(sheet_embed_url):
    """
    Writes a gonogo.html that embeds the published Google Sheet via iframe.
    Prefer users to publish-to-web and paste the pubhtml URL.
    """
    # safe empty handling
    if not sheet_embed_url:
        content = "<div style='color:orange'>No sheet URL set in settings</div>"
    else:
        content = f'<iframe src="{sheet_embed_url}" width="100%" height="600" style="border:none;"></iframe>'
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
}}
.container {{
    width: 95%;
    margin: 10px auto;
}}
</style>
<script>
setTimeout(()=>location.reload(),3000);
</script>
</head>
<body>
<div class="container">
{content}
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

        def save_settings_cmd():
            new_mode = mode_var.get()
            new_url = sheet_entry.get().strip()
            self.settings["mode"] = new_mode
            self.settings["sheet_url"] = new_url
            save_settings(self.settings)
            # regenerate gonogo HTML according to mode
            if new_mode == "sheet":
                embed = ensure_iframe_url(new_url)
                write_gonogo_html_iframe(embed)
                # prefill labels from CSV fetch if possible
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

        # Update gonogo block depending on settings.mode
        mode = self.settings.get("mode", "sheet")
        if mode == "sheet":
            # ensure embed is current
            embed = ensure_iframe_url(self.settings.get("sheet_url", ""))
            write_gonogo_html_iframe(embed)
            # optionally update mirrored labels by fetching CSV (best-effort)
            vals = fetch_gonogo_csv(self.csv_link)
            self.range_label.config(text=f"RANGE: {vals[0]}", fg=get_status_color(vals[0]))
            self.weather_label.config(text=f"WEATHER: {vals[1]}", fg=get_status_color(vals[1]))
            self.vehicle_label.config(text=f"VEHICLE: {vals[2]}", fg=get_status_color(vals[2]))
        else:
            # manual mode: write gonogo based on manual toggles and update labels
            vals = [self.manual_range, self.manual_weather, self.manual_vehicle]
            write_gonogo_html_from_values(vals)
            self.range_label.config(text=f"RANGE: {vals[0]}", fg=get_status_color(vals[0]))
            self.weather_label.config(text=f"WEATHER: {vals[1]}", fg=get_status_color(vals[1]))
            self.vehicle_label.config(text=f"VEHICLE: {vals[2]}", fg=get_status_color(vals[2]))

        # schedule next tick
        self.root.after(500, self.update_clock)  # 2× per second

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
