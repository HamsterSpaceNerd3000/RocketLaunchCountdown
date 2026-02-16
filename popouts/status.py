# RLCs status popout. Ran as a subprocess by main

# Imports
import dearpygui.dearpygui as dpg
import json
import os

# ---------- File Path ----------
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE_FILE = os.path.join(BASE_DIR, "countdown_state.json")

dpg.create_context()

fonts = {}
with dpg.font_registry():
    try:
        font_p = "C:/Windows/Fonts/consola.ttf"
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
def update_from_file():
    if not os.path.exists(STATE_FILE):
        return

    try:
        with open(STATE_FILE, "r") as f:
            data = json.load(f)

        if "statuses" not in data:
            return

        for key, info in data["statuses"].items():
            status_tag = f"{key}_status"

            if dpg.does_item_exist(status_tag):
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
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("gonogo_window", True)

while dpg.is_dearpygui_running():
    update_from_file()
    center_status_text()
    dpg.render_dearpygui_frame()

dpg.destroy_context()
