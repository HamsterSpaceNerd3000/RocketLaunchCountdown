# RLCs main operating file. It pulls the countdown info from that script and pulls setting etc. It handles all of the GUI, but pulls information from other scripts.

# Imports
import dearpygui.dearpygui as dpg
import os
import settings
from datetime import datetime, timedelta
import countdown_handler as ch
import path_handler as ph
import subprocess
import time
import json
import spreadsheet_handler as sh
import threading
import re
import sys

# DPG init
dpg.create_context()

version = "0.8.0"

error_theme = None

# Countdown class
class CountdownMain:
    def __init__(self):
        self.running = False
        self.target_time = None
        self.hold = False
        self.scrubbed = False
        self.scrub_label = ""
        self.hold_start_time = None

        self.statuses = {"WEATHER": 0, "RANGE": 0, "VEHICLE": 0}
        self.status_values = {"WEATHER": "N/A", "RANGE": "N/A", "VEHICLE": "N/A"}
        self.auto_hold = False

        self.settings = settings.defaults

        self.load_settings()

        self.target_set = dpg.get_value("target_in")

        self.last_sync_time = 0
        self.sync_interval = 3  # Seconds between sheet updates
    
    # Function to load settings and update approapriate values
    def load_settings(self):
        settings.load()
        self.update_auto_hold_thres()
        self.update_min_hold_reset()

    # Function to update the auto hold time
    def update_auto_hold_thres(self):
        try:
            h, m, s = map(int, self.settings["auto_hold_time"].split(":"))
            self.auto_hold_thres = h * 3600 + m * 60 + s
        except:
            self.auto_hold_thres = 0

    # Function to update the minimum time the clock is set to after a hold
    def update_min_hold_reset(self):
        try:
            h, m, s = map(int, self.settings["min_hold_reset"].split(":"))
            self.min_hold_reset_thres = h * 3600 + m * 60 + s
        except:
            self.min_hold_reset_thres = 0
    
    # Function to toggle any window
    def toggle_window(self, tag):
        if dpg.is_item_shown(tag):
            dpg.hide_item(tag)
        else:
            dpg.show_item(tag)

    
    # Function to toggle the countdown prefix (T- / L-)
    def toggle_countdown_prefix(self):
        self.settings["prefix"] = "L-" if self.settings["prefix"] == "T-" else "T-"
        if not self.running or (self.target_time and (self.target_time - datetime.now()).total_seconds() > 0):
            dpg.set_value("prefix_text", self.settings["prefix"])
    
    # Function to reset everything to load
    def reset(self):
        self.running = False
        self.target_time = None
        self.hold = False
        self.scrubbed = False
        self.hold_start_time = None

        if dpg.does_item_exist("prefix_text"):
            dpg.set_value("prefix_text", self.settings["prefix"])
        if dpg.does_item_exist("countdown_text"):
            dpg.set_value("countdown_text", "00:00:00")
            dpg.configure_item("countdown_text", color=(255, 255, 255))
        if dpg.does_item_exist("target_in"):
            dpg.set_value("target_in", "")
        
        log_to_console("System Reset.")
    # Function to toggle a hold
    def toggle_hold(self):
        if not self.running or self.scrubbed:
            return
        
        diff_sec = (self.target_time - datetime.now()).total_seconds()

        if not self.hold:
            # Start hold
            if diff_sec <= 0:
                return

            # Time recycle check
            if diff_sec < self.min_hold_reset_thres:
                self.target_time = datetime.now() + timedelta(seconds=self.min_hold_reset_thres)
            
            self.hold = True
            self.hold_start_time = datetime.now()
        else:
            hold_duration = datetime.now() - self.hold_start_time
            self.target_time += hold_duration
            self.hold = False

    # Function to trigger a scrub
    def trigger_scrub(self):
        if self.target_time:
            diff_sec = (self.target_time - datetime.now()).total_seconds()
            self.scrub_label = "RUD" if diff_sec <= 0 else "SCRUB"
        else:
            self.scrub_label = "SCRUB"
        
        self.scrubbed, self.running, self.hold = True, False, False

    # Function to cycle status
    def cycle_status(self, key):
        self.statuses[key] = (self.statuses[key] + 1) % 3
        self.update_status_gui(key, self.statuses[key]) 
    
    def update_status_gui(self, key, idx):
        tag = f"{key.lower()}_status_text"
        if not dpg.does_item_exist(tag):
            return
        
        current_val = str(dpg.get_value(tag))
        if "%" in current_val:
            return 

        labels = ["NO-GO", "GO", "N/A"]
        colors = [(255, 0, 0), (0, 255, 0), (128, 128, 128)]
        
        dpg.set_value(tag, labels[idx])
        dpg.configure_item(tag, color=colors[idx])
    
    def start_countdown(self):
        self.target_set = dpg.get_value("target_in")
        ch.start_countdown(self)
    
    # Logic update function
    def update(self):
        if dpg.does_item_exist("countdown_text"):
            window_width = dpg.get_item_width("countdown")
        else:
            window_width = 800
        controls_width = dpg.get_item_width("Controls") or 400
        is_touch = self.settings["touch_screen"]

        box_w, main_h = (window_width - 60) / 3, (100 if is_touch else 60)
        box_h = box_w * 0.50

        # Status boxes have fixed width of 210 in the status_display
        status_box_w = 210
        status_box_h = 225 * 0.5
        
        for item in ["weather", "range", "vehicle"]:
            tag = f"{item}_box"
            if dpg.does_item_exist(tag):
                dpg.configure_item(tag, width=int(status_box_w), height=int(status_box_h))
        
        if dpg.does_item_exist("start_button"):
            for btn in ["start_button", "hold_button", "scrub_button", "reset_button"]:
                dpg.configure_item(btn, height=main_h)
            
            btn_w = (controls_width - 42) / 3
            for btn in ["btn_w_manual", "btn_r_manual", "btn_v_manual"]:
                dpg.configure_item(btn, width=int(btn_w), height=(80 if is_touch else 40))
        
        if self.settings["show_mission"] and dpg.does_item_exist("mission_title"):
            m_size = dpg.get_text_size(dpg.get_value("mission_title"), font=fonts.get("large", 0))
            if m_size:
                dpg.set_item_indent("mission_title", max(0, int((window_width / 2) - (m_size[0] / 2) + self.settings["centering_offset"])))
        
        if dpg.does_item_exist("count_group"):
            count_text = f"{dpg.get_value('prefix_text')} {dpg.get_value('countdown_text')}"
            count_size = dpg.get_text_size(count_text, font=fonts.get("huge", 0))
            if count_size:
                dpg.set_item_indent("count_group", max(0, int((window_width/2) - (count_size[0]/2) + self.settings["centering_offset"])))
        
        # Center status text items
        fixed_box_width = 210
        for item in ["weather", "range", "vehicle"]:
            name_tag = f"{item}_name_text"
            status_tag = f"{item}_status_text"
            if dpg.does_item_exist(name_tag):
                name_size = dpg.get_text_size(dpg.get_value(name_tag), font=fonts.get("status", 0))
                if name_size:
                    indent = max(0, int(((fixed_box_width / 2) - (name_size[0] / 2)) * 0.99))
                    dpg.set_item_indent(name_tag, indent)
            if dpg.does_item_exist(status_tag):
                status_size = dpg.get_text_size(dpg.get_value(status_tag), font=fonts.get("status", 0))
                if status_size:
                    indent = max(0, int(((fixed_box_width / 2) - (status_size[0] / 2)) * 0.99))
                    dpg.set_item_indent(status_tag, indent)
        
        if self.scrubbed:
            dpg.set_value("prefix_text", ""); dpg.set_value("countdown_text", self.scrub_label)
            dpg.configure_item("countdown_text", color=(255, 0, 0))
        
        elif self.hold:
            dpg.set_value("prefix_text", "H+")
            dpg.set_value("countdown_text", self.format_time(int((datetime.now() - self.hold_start_time).total_seconds())))
        
        elif self.running:
            prefix, display_val = ch.run_countdown(self)
            dpg.set_value("prefix_text", prefix)
            dpg.set_value("countdown_text", self.format_time(display_val))

        # Touchscreen width check
        is_touch = self.settings["touch_screen"]
        side_padding = 50 if is_touch else 35
        btn_w = (controls_width - side_padding) / 3 

        # Main controls width checking
        main_btn_w = (controls_width - 32) / 2 # For the 2-column layout (START/HOLD)

        if dpg.does_item_exist("start_button"):
            dpg.configure_item("start_button", width=int(main_btn_w))
            dpg.configure_item("hold_button", width=int(main_btn_w))
            dpg.configure_item("scrub_button", width=int(main_btn_w))
            dpg.configure_item("reset_button", width=int(main_btn_w))

    # Function to format time
    def format_time(self, seconds):
        ts = int(seconds)
        return f"{ts//3600:02}:{(ts%3600)//60:02}:{ts%60:02}"
    
    # Function to export the control state to popout instances
    def export_state(self):
        if self.scrubbed:
            color = (255, 0, 0)
        else:
            color = (255, 255, 255)

        status_labels = ["NO-GO", "GO", "N/A"]
        status_colors = [(255, 0, 0), (0, 255, 0), (128, 128, 128)]

        status_data = {}

        for key, idx in self.statuses.items():
            status_data[key] = {
                "label": status_labels[idx],
                "color": status_colors[idx]
            }

        data = {
            "mission_name": self.settings["mission_name"],
            "show_mission": self.settings["show_mission"],
            "prefix": dpg.get_value("prefix_text") if dpg.does_item_exist("prefix_text") else "",
            "countdown_text": dpg.get_value("countdown_text") if dpg.does_item_exist("countdown_text") else "00:00:00",
            "centering_offset": self.settings["centering_offset"],
            "scrubbed": self.scrubbed,
            "hold": self.hold,
            "countdown_color": color,
            "statuses": status_data,
            "status_values": self.status_values,
            "manual_concerns": self.settings["manual_concerns"],

            "box_bg_color": self.settings["box_bg_color"],
            "box_outline": self.settings["box_outline"],
            "box_border_width": self.settings["box_border_width"]

        }

        # JSON dump logic
        db_folder = ph.PATH / "database"
        final_path = db_folder / "countdown_state.json"

        with open(final_path, "w") as f:
            json.dump(data, f)

    
    # Refresh data from spreadsheet
    def spreadsheet_refresh(self):
        def background_task():
            data = sh.load_sheet_data()
            
            if "error" in data:
                log_to_console(data["error"])
                return

            # Robust mapping dictionary
            mapping = {"NO-GO": 0, "GO": 1, "N/A": 2}
            
            for key in ["WEATHER", "RANGE", "VEHICLE"]:
                raw_val = str(data.get(key.lower(), "N/A")).strip().upper()
                self.status_values[key] = raw_val
                
                if key == "WEATHER" and "%" in raw_val:
                    try:
                        # Convert "85%" -> 85
                        percent = int(raw_val.replace("%", ""))
                        
                        # Determine Color and Status based on thresholds
                        if percent >= 60:
                            status_color = (0, 255, 0)   # Green
                            self.statuses[key] = 1       # GO
                        elif 45 <= percent < 60:
                            status_color = (255, 165, 0) # Orange
                            self.statuses[key] = 1       # Still GO (logic-wise), but orange visual
                        else:
                            status_color = (255, 0, 0)   # Red
                            self.statuses[key] = 0       # NO-GO

                        # Update GUI
                        if dpg.does_item_exist("weather_status_text"):
                            dpg.set_value("weather_status_text", raw_val)
                            dpg.configure_item("weather_status_text", color=status_color)
                    except ValueError:
                        # Fallback if the percentage isn't a valid number
                        self.update_status_gui(key, 2) 
                else:
                    # Standard GO/NO-GO mapping for Range and Vehicle
                    idx = mapping.get(raw_val, 2)
                    self.statuses[key] = idx
                    self.update_status_gui(key, idx)

            # Update Concerns
            raw_concerns = data.get("concerns", "No Data")

            # If there is actual data, split by comma and add bullets
            if raw_concerns and raw_concerns != "No Data":
                # Split by comma, strip whitespace, and filter out empty strings
                items = [item.strip() for item in raw_concerns.split(",") if item.strip()]
                # Join with newlines and a bullet point
                concerns = "\n".join([f"- {item}" for item in items])
            else:
                concerns = "No current concerns."

            self.settings["manual_concerns"] = concerns
            if dpg.does_item_exist("concerns_text"):
                dpg.set_value("concerns_text", concerns)
            
            log_to_console(f"Sync: W:{data.get('weather')} R:{data.get('range')} V:{data.get('vehicle')}")

        threading.Thread(target=background_task, daemon=True).start()

    def handle_link_paste(self, app_data):
        # 'app_data' is what the user actually typed/pasted into the box
        new_url = app_data
        
        # Correct way to update the dictionary
        self.settings["spreadsheet_link"] = new_url

        # Look for the GID
        try:
            match = re.search(r"gid=(\d+)", new_url)
            if match:
                extracted_gid = match.group(1)
                self.settings["sheet_gid"] = extracted_gid
                if dpg.does_item_exist("gid_input"):
                    dpg.set_value("gid_input", extracted_gid)
        except Exception as e:
            log_to_console(f"Regex Error: {e}")

    def reset_settings(self):
        self.settings = settings.defaults.copy()

        settings.save(self.settings)

        if dpg.does_item_exist("mission_title"):
            dpg.set_value("mission_title", self.settings["mission_name"])
        
        dpg.set_value("target_in", "")
        if dpg.does_item_exist("sheet_name_input"):
            dpg.set_value("sheet_name_input", "")
        if dpg.does_item_exist("gid_input"):
            dpg.set_value("gid_input", "")
        
        apply_theme()

        log_to_console("Settings reset to defaults")
state = CountdownMain()

# Font handling
fonts = {}

# Grab script root
root = ph.PATH

with dpg.font_registry():
    try:
        font_p = root / "database" / "ShareTechMono-Regular.ttf"
        # Create your sizes
        fonts["huge"] = dpg.add_font(font_p, 120)
        fonts["large"] = dpg.add_font(font_p, 60)
        fonts["status"] = dpg.add_font(font_p, 40)
        fonts["default"] = dpg.add_font(font_p, 12) # Add a standard size
        
        # This one line makes EVERYTHING use this font by default
        dpg.bind_font(fonts["default"]) 
    except Exception as e:
        print(f"EXACT PATH: {os.path.abspath(font_p)}")
        # This prints the REAL reason it's failing
        print(f"Font Load Failed with error: {e}")
        pass

box_theme = None

# Function to apply theme before creating windows
def apply_theme():
    global box_theme, error_theme

    bg = [int(max(0, min(255, c * 255))) for c in state.settings["box_bg_color"]]
    border = [int(max(0, min(255, c * 255))) for c in state.settings["box_outline"]]
    txt_color = [int(max(0, min(255, c * 255))) for c in state.settings["txt_color"]]
    border_width = state.settings["box_border_width"]
    
    # Delete old theme if it exists
    if box_theme and dpg.does_item_exist(box_theme):
        dpg.delete_item(box_theme)

    with dpg.theme() as box_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_ChildBg, bg)
            dpg.add_theme_color(dpg.mvThemeCol_Border, border)
            dpg.add_theme_style(dpg.mvStyleVar_ChildBorderSize, border_width)
            dpg.add_theme_color(dpg.mvThemeCol_Text, txt_color)


    with dpg.theme() as error_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_color(dpg.mvThemeCol_Border, (255, 0, 0), category=dpg.mvThemeCat_Core)
            dpg.add_theme_style(dpg.mvStyleVar_WindowBorderSize, 2, category=dpg.mvThemeCat_Core)

    dpg.bind_theme(box_theme)


def select_display(display_type):
    dw  = DisplayWindow()
    if display_type == "countdown":
        if dpg.does_item_exist("countdown"):
            pass
        else:
            dw.countdown_display()
    elif display_type == "status":
        if dpg.does_item_exist("status_display"):
            pass
        else:
            dw.status_display()
    elif display_type == "concerns":
        if dpg.does_item_exist("concerns_display"):
            pass
        else:
            dw.concerns_display()
    
    dpg.configure_item("window_select", show=False)

# Function to open the countdown popout
def countdown_popout():
    if getattr(sys, 'frozen', False):
        popout_bin = ph.PATH / "countdown" 
        subprocess.Popen([str(popout_bin)], cwd=str(ph.PATH), env=os.environ)
    else:
        subprocess.Popen([sys.executable, "popouts/countdown.py"])
    select_display("countdown")
    dpg.configure_item("window_select", show=False)

# Function to open the status popout
def status_popout():
    if getattr(sys, 'frozen', False):
        popout_bin = ph.PATH / "status" 
        subprocess.Popen([str(popout_bin)], cwd=str(ph.PATH), env=os.environ)
    else:
        subprocess.Popen([sys.executable, "popouts/status.py"])
    select_display("status")
    dpg.configure_item("window_select", show=False)

# Function to open the concerns popout
def concerns_popout():
    if getattr(sys, 'frozen', False):
        popout_bin = ph.PATH / "concerns"
        subprocess.Popen([str(popout_bin)], cwd=str(ph.PATH), env=os.environ)
    else:
        subprocess.Popen([sys.executable, "popouts/concerns.py"])
    select_display("concerns")
    dpg.configure_item("window_select", show=False)

# Display window class
class DisplayWindow:
    def __init__(self):
        self.window_width = 702
        self.window_x = 415
    
    # Countdown display window
    def countdown_display(self):
        with dpg.window(label="Countdown Clock", tag="countdown", width=self.window_width, height=232, pos=[self.window_x, 0]):
            dpg.add_text(state.settings["mission_name"], tag="mission_title", show=state.settings["show_mission"])
            if "large" in fonts:
                dpg.bind_item_font("mission_title", fonts["large"])
            
            # Countdown group
            count_group_tag = "count_group"
            with dpg.group(horizontal=True, tag=count_group_tag):
                dpg.add_text(state.settings["prefix"], tag="prefix_text")
                dpg.add_text("00:00:00", tag="countdown_text")
                if "huge" in fonts:
                    dpg.bind_item_font("prefix_text", fonts["huge"])
                    dpg.bind_item_font("countdown_text", fonts["huge"])
        apply_theme()
    
    # Status display window
    def status_display(self):
        with dpg.window(label="Status Display", tag="status_display", width=self.window_width, height=148, pos=[self.window_x, 232]):
            with dpg.group(horizontal=True):
                for item in ["WEATHER", "RANGE", "VEHICLE"]:

                    box_tag = f"{item.lower()}_box"
                    name_tag = f"{item.lower()}_name_text"
                    status_tag = f"{item.lower()}_status_text"
                    
                    with dpg.child_window(tag=box_tag, width=210, height=-1):
                        dpg.add_text(f"{item}", tag=name_tag)
                        dpg.add_text("NO-GO", tag=status_tag)
                        if "status" in fonts:
                            dpg.bind_item_font(name_tag, fonts["status"])
                            dpg.bind_item_font(status_tag, fonts["status"])
                    
                    if item != "VEHICLE":
                        dpg.add_spacer(width=12)
                
                dpg.add_spacer(width=15, tag="status_spacer_right")
        apply_theme()
        for key in ["WEATHER", "RANGE", "VEHICLE"]:
            current_idx = state.statuses.get(key, 2) # Default to 2 (N/A) if not found
            state.update_status_gui(key, current_idx)

    # Major concerns display window
    def concerns_display(self):
        # Increased window height from 160 to 250
        with dpg.window(label="Major Concerns Display", tag="concerns_display", width=self.window_width, height=250, pos=[self.window_x, 380]):
            dpg.add_text("MAJOR CONCERNS: ")
            # Increased child box height from 100 to 180
            with dpg.child_window(tag="concerns_box", width=-1, height=180): 
                c_txt = dpg.add_text(state.settings["manual_concerns"], tag="concerns_text")
                if "status" in fonts:
                    dpg.bind_item_font(c_txt, fonts["status"])
        apply_theme()

# Control window
with dpg.window(label="Controls", tag="Controls", width=415, height=525, pos=[0, 0]):
    with dpg.group(horizontal=True):
        dpg.add_button(label="ADD DISPLAY", tag="add_display_button", width=195, height=40, callback=lambda: dpg.configure_item("window_select", show=True))
        dpg.add_button(label="SETTINGS", width=195, height=40, callback=lambda: state.toggle_window("SettingsWin"))
    dpg.add_separator()
    with dpg.group(horizontal=True):
        dpg.add_input_text(label="MISSION NAME", width=200, default_value=state.settings["mission_name"], callback=lambda s, a: (state.settings.update({"mission_name": a}), dpg.set_value("mission_title", a)))
        dpg.add_button(label="T-/L-", width=-1, callback=state.toggle_countdown_prefix)
    dpg.add_checkbox(label="SHOW MISSION NAME", default_value=state.settings["show_mission"], callback=lambda s, a: (state.settings.update({"show_mission": a}), dpg.configure_item("mission_title", show=a)))
    dpg.add_separator()
    dpg.add_input_text(tag="target_in", hint="Enter Countdown (HH:MM:SS)", width=-1)
    with dpg.group(horizontal=True):
        dpg.add_button(label="START", tag="start_button", width=195, callback=state.start_countdown)
        dpg.add_button(label="HOLD/RESUME", tag="hold_button", width=195, callback=state.toggle_hold)
    with dpg.group(horizontal=True):
        dpg.add_button(label="SCRUB", tag="scrub_button", width=195, callback=state.trigger_scrub)
        dpg.add_button(label="RESET", tag="reset_button", width=195, callback=state.reset)
    dpg.add_spacer(height=5)
    dpg.add_separator()
    dpg.add_text("MANUAL STATUS")
    with dpg.group(horizontal=True):
        dpg.add_button(label="WEATHER", tag="btn_w_manual", width=125, callback=lambda: state.cycle_status("WEATHER"))
        dpg.add_button(label="RANGE", tag="btn_r_manual", width=125, callback=lambda: state.cycle_status("RANGE"))
        dpg.add_button(label="VEHICLE", tag="btn_v_manual", width=125, callback=lambda: state.cycle_status("VEHICLE"))
    with dpg.group(horizontal=True):
        dpg.add_text("MAJOR CONCERNS")
        dpg.add_input_text(tag="concerns_in", width=-1, default_value=state.settings["manual_concerns"], callback=lambda s, a: (state.settings.update({"manual_concerns": a}), dpg.set_value("concerns_text", a)))
    dpg.add_spacer(height=5)
    dpg.add_separator()
    dpg.add_text("AUTO-HOLD SETTINGS")
    dpg.add_checkbox(label="ENABLE AUTO-HOLD", default_value=state.auto_hold, callback=lambda s, a: setattr(state, "auto_hold", a))
    with dpg.group(horizontal=True):
        dpg.add_text("Hold At (HH:MM:SS)")
        dpg.add_input_text(default_value=state.settings["auto_hold_time"], width=-1, callback=lambda s, a: (state.settings.update({"auto_hold_time": a}), state.update_auto_hold_thres()))
    with dpg.group(horizontal=True):
        dpg.add_text("Min Hold Reset (HH:MM:SS)")
        dpg.add_input_text(default_value=state.settings["min_hold_reset"], width=-1, callback=lambda s, a: (state.settings.update({"min_hold_reset": a}), state.update_min_hold_reset()))

# Window select popup
with dpg.popup(parent="add_display_button", mousebutton=dpg.mvMouseButton_Left, tag="window_select"):
    dpg.add_text("Select Display Type:")
    dpg.add_separator()
    with dpg.group(horizontal=True):
        dpg.add_button(label="COUNTDOWN CLOCK", callback=lambda: select_display("countdown"))
        dpg.add_button(arrow=True, direction=dpg.mvDir_Right, tag="countdown_pop_btn", callback=lambda: countdown_popout())
        with dpg.tooltip("countdown_pop_btn"):
            dpg.add_text("Countdown Popout")
    with dpg.group(horizontal=True):
        dpg.add_button(label="STATUS DISPLAY", callback=lambda: select_display("status"))
        dpg.add_button(arrow=True, direction=dpg.mvDir_Right, tag="status_pop_btn", callback=lambda: status_popout())
        with dpg.tooltip("status_pop_btn"):
            dpg.add_text("Status Popout")
    with dpg.group(horizontal=True):
        dpg.add_button(label="CONCERNS DISPLAY", callback=lambda: select_display("concerns"))
        dpg.add_button(arrow=True, direction=dpg.mvDir_Right, tag="concerns_pop_btn", callback=lambda: concerns_popout())
        with dpg.tooltip("concerns_pop_btn"):
            dpg.add_text("Concerns Popout")


# Spreadsheet window
with dpg.window(label="Spreadsheet Manager", tag="spreadsheet", width=420, height=300, pos=[415, 380], show=False):
    dpg.add_text("Spreadsheet Link")
    dpg.add_input_text(default_value=state.settings["spreadsheet_link"], width=-1, callback=lambda s, a: state.handle_link_paste(a))
    dpg.add_text("Spreadsheet GID (At the end of the link)")
    dpg.add_input_text(tag="gid_input", default_value=state.settings["sheet_gid"], width=-1, callback=lambda s, a: state.settings.update({"sheet_gid": a}))
    dpg.add_text("Weather Cell")
    dpg.add_input_text(default_value=state.settings["weather_sheet_cell"], width=-1, callback=lambda s, a: (state.settings.update({"weather_sheet_cell": a})))
    dpg.add_text("Range Cell")
    dpg.add_input_text(default_value=state.settings["range_sheet_cell"], width=-1, callback=lambda s, a: (state.settings.update({"range_sheet_cell": a})))
    dpg.add_text("Vehicle Cell")
    dpg.add_input_text(default_value=state.settings["vehicle_sheet_cell"], width=-1, callback=lambda s, a: (state.settings.update({"vehicle_sheet_cell": a})))
    dpg.add_text("Concerns Cell")
    dpg.add_input_text(default_value=state.settings["concerns_sheet_cell"], width=-1, callback=lambda s, a: (state.settings.update({"concerns_sheet_cell": a})))
    dpg.add_button(label="SAVE", width=-1, height=30, callback=lambda: settings.save(state.settings))


# Settings window
with dpg.window(label="Settings", tag="SettingsWin", width=415, height=300, pos=[415, 80], show=False):
   dpg.add_checkbox(label="TOUCH SCREEN", default_value=state.settings["touch_screen"], callback=lambda s, a: state.settings.update({"touch_screen": a}))
   dpg.add_color_edit(label="Box Background Color", default_value=state.settings["box_bg_color"], no_alpha=True, alpha_bar=False, callback=lambda s, a: (state.settings.update({"box_bg_color": a[:3]}), apply_theme()))
   dpg.add_color_edit(label="Box Outline Color", default_value=state.settings["box_outline"], no_alpha=True, alpha_bar=False, callback=lambda s, a: (state.settings.update({"box_outline": a[:3]}), apply_theme()))
   dpg.add_color_edit(label="Text Color", default_value=state.settings["txt_color"], no_alpha=True, alpha_bar=False, callback=lambda s, a: (state.settings.update({"txt_color": a[:3]}), apply_theme()))
   dpg.add_button(label="SPREADSHEET", width=-1, height=30, callback=lambda: state.toggle_window("spreadsheet"))
   dpg.add_button(label="SAVE", width=-1, height=30, callback=lambda: settings.save(state.settings))
   dpg.add_button(label="RESET TO DEFAULTS", width=-1, height=30, callback=state.reset_settings)

console_logs = []

# Console logging function
def log_to_console(error):
    console_logs.append(error)
    if len(console_logs) > 50:
        console_logs.pop(0)

    if dpg.does_item_exist("console_widget"):
        current_log_string = "\n".join(console_logs)
        dpg.set_value("console_widget", current_log_string)

        dpg.set_y_scroll("console_scroll", dpg.get_y_scroll_max("console_scroll"))


# Error console
with dpg.window(label="System Console", tag="console_window", width=390, height=200, pos=[830, 0], show=False):
    with dpg.child_window(tag="console_scroll", autosize_x=True, border=True):
        dpg.add_text("", tag="console_widget")

# Menubar
with dpg.viewport_menu_bar():
    with dpg.menu(label="Displays"):
        dpg.add_menu_item(label="Countdown Clock", callback=lambda: select_display("countdown"))
        dpg.add_menu_item(label="Status Display", callback=lambda: select_display("status"))
        dpg.add_menu_item(label="Major Concerns Display", callback=lambda: select_display("concerns"))
        dpg.add_menu_item(label="Controls", callback=lambda: state.toggle_window("Controls"))
    with dpg.menu(label="Help"):
        dpg.add_menu_item(label="Guide")
        dpg.add_menu_item(label="Console", callback=lambda: state.toggle_window("console_window"))
    with dpg.menu(label="Popouts"):
        dpg.add_menu_item(label="Countdown", callback=lambda: countdown_popout())
        dpg.add_menu_item(label="Status Popout", callback=lambda: status_popout())
        dpg.add_menu_item(label="Concerns Popout", callback=lambda: concerns_popout())
    with dpg.menu(label="Settings"):
        dpg.add_menu_item(label="Main Settings", callback=lambda: state.toggle_window("SettingsWin"))
        dpg.add_menu_item(label="Spreadsheet Settings", callback=lambda: state.toggle_window("spreadsheet"))

# DPG wrap up
dpg.create_viewport(title=f"RocketLaunchCountdown v{version}", width=1231, height=720, small_icon="RLCLogo.ico", large_icon="RLCLogo.ico")
dpg.setup_dearpygui()
dpg.set_viewport_always_top(True)
dpg.show_viewport()
apply_theme()
state.update()
for key in state.statuses:
    state.update_status_gui(key, state.statuses[key])
while dpg.is_dearpygui_running():
    # Sync handler
    current_time = time.time()
    if current_time - state.last_sync_time > state.sync_interval:
        # Link check
        if state.settings.get("spreadsheet_link"):
            state.spreadsheet_refresh() 
            state.last_sync_time = current_time

    # Frame handeling
    state.update()
    state.export_state()
    dpg.render_dearpygui_frame()
dpg.destroy_context()