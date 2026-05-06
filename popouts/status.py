# RLCs status popout. Ran as a subprocess by main

# Imports
import dearpygui.dearpygui as dpg
import json
import os

import popout_init
popout_init.bootstrap()

import path_handler as ph

# ---------- File Path ----------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

dpg.create_context()

fonts = {}
with dpg.font_registry():
    root = ph.PATH
    try:
        font_p = root / "database" / "ShareTechMono-Regular.ttf"
        fonts["status"] = dpg.add_font(font_p, 60)
    except:
        pass

# ---------- UI ----------
with dpg.window(label="GO / NO-GO", tag="gonogo_window", no_collapse=True):

    with dpg.group(horizontal=True):
        dpg.add_spacer(width=15, tag="status_spacer_left")
        for item in ["WEATHER", "RANGE", "VEHICLE"]:
            name_tag = f"{item}_name"
            status_tag = f"{item}_status"
            box_tag = f"{item}_box"

            with dpg.child_window(tag=box_tag, width=250, height=168):
                dpg.add_text(item, tag=name_tag)
                dpg.add_text("NO-GO", tag=status_tag)

                if "status" in fonts:
                    dpg.bind_item_font(name_tag, fonts["status"])
                    dpg.bind_item_font(status_tag, fonts["status"])

            if item != "VEHICLE":
                dpg.add_spacer(width=12)
        
        dpg.add_spacer(width=15, tag="status_spacer_right")

# ---------- Sync ----------
def apply_popout_theme(data):
    """Applies theme using colors directly from the state data."""
    try:
        # Pull colors from the state dictionary
        # Assuming main.py saves them as 0-255 or 0.0-1.0
        # If they are already 0-255 in the state file, remove the * 255
        bg = data.get("box_bg_color", [15, 15, 15])
        border = data.get("box_outline", [255, 255, 255])
        txt_color = data.get("txt_color", [255, 255, 255])
        border_width = data.get("box_border_width", 1)

        with dpg.theme() as global_theme:
            with dpg.theme_component(dpg.mvChildWindow):
                dpg.add_theme_color(dpg.mvThemeCol_ChildBg, bg)
                dpg.add_theme_color(dpg.mvThemeCol_Border, border)
                dpg.add_theme_style(dpg.mvStyleVar_ChildBorderSize, border_width)
                dpg.add_theme_color(dpg.mvThemeCol_Text, txt_color)
            
            with dpg.theme_component(dpg.mvAll):
                dpg.add_theme_color(dpg.mvThemeCol_Text, txt_color)

        dpg.bind_theme(global_theme)
    except Exception as e:
        pass

state_path = ph.PATH / "database" / "countdown_state.json"

def update_from_file():
    if not os.path.exists(state_path):
        return

    try:
        with open(state_path, "r") as f:
            data = json.load(f)

        # 1. Update the Theme first
        apply_popout_theme(data)

        # 2. Update the Status Labels and Colors
        if "statuses" in data:
            for key, info in data["statuses"].items():
                status_tag = f"{key}_status"
                if dpg.does_item_exist(status_tag):
                    if key == "WEATHER" and "status_values" in data:
                        weather_value = data["status_values"].get("WEATHER", info["label"])
                        dpg.set_value(status_tag, weather_value)

                        try:
                            if "%" in weather_value:
                                percent = int(weather_value.replace("%", ""))
                                if percent >= 60:
                                    color = (0, 255, 0)
                                elif 45 <= percent < 60:
                                    color = (255, 165, 0)
                                else:
                                    color = (255, 0, 0)
                                dpg.configure_item(status_tag, color=color)
                            else:
                                dpg.configure_item(status_tag, color=tuple(info["color"]))
                        except ValueError:
                            dpg.configure_item(status_tag, color=tuple(info["color"]))
                    else:
                        dpg.set_value(status_tag, info["label"])
                        dpg.configure_item(status_tag, color=tuple(info["color"]))
    except:
        pass

# Center text within status boxes (matches main.py behavior)
def center_status_text():
    fixed_box_width = 240
    for item in ["WEATHER", "RANGE", "VEHICLE"]:
        name_tag = f"{item}_name"
        status_tag = f"{item}_status"
        
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


# ---------- Viewport ----------
dpg.create_viewport(title="GO / NO-GO", width=870, height=220)
dpg.set_viewport_always_top(True)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("gonogo_window", True)

while dpg.is_dearpygui_running():
    update_from_file()
    center_status_text()
    dpg.render_dearpygui_frame()

dpg.destroy_context()
