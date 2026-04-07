# RLCs concerns popout.

# Imports
import dearpygui.dearpygui as dpg
import json
import os

dpg.create_context()

fonts = {}
with dpg.font_registry():
    try:
        font_p = "C:/Windows/Fonts/consola.ttf"
        fonts["status"] = dpg.add_font(font_p, 40)
    except:
        pass

with dpg.window(tag="main_window", no_title_bar=True, no_resize=True, no_move=True):
    dpg.add_text("MAJOR CONCERNS:")
    with dpg.child_window(tag="concerns_box", width=-1, height=-1):
        dpg.add_text("", tag="concerns_text")

    if "status" in fonts:
        dpg.bind_item_font("concerns_text", fonts["status"])

dpg.create_viewport(title="Major Concerns", width=800, height=250)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("main_window", True)
dpg.set_viewport_always_top(True)

def update():
    try:
        path = os.path.join(os.path.dirname(__file__), "..", "countdown_state.json")
        with open(path, "r") as f:
            data = json.load(f)

        dpg.set_value("concerns_text", data.get("manual_concerns", ""))

    except:
        pass

while dpg.is_dearpygui_running():
    update()
    dpg.render_dearpygui_frame()

dpg.destroy_context()
