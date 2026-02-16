# RLCs countdown clock popout window. This script gets run a subprocess by main.

# Imports
import dearpygui.dearpygui as dpg
import json
import os

dpg.create_context()

fonts = {}
with dpg.font_registry():
    try:
        font_p = "C:/Windows/Fonts/consola.ttf"
        fonts["huge"] = dpg.add_font(font_p, 120)
        fonts["large"] = dpg.add_font(font_p, 60)
    except:
        pass


# UI layout
with dpg.window(
    label="Countdown Clock",
    tag="countdown"
):
    dpg.add_text("", tag="mission_title")
    if "large" in fonts:
        dpg.bind_item_font("mission_title", fonts["large"])

    with dpg.group(horizontal=True, tag="count_group"):
        dpg.add_text("", tag="prefix_text")
        dpg.add_text("00:00:00", tag="countdown_text")

        if "huge" in fonts:
            dpg.bind_item_font("prefix_text", fonts["huge"])
            dpg.bind_item_font("countdown_text", fonts["huge"])


# Sync from main
def update_from_file():
    if not os.path.exists("countdown_state.json"):
        return

    try:
        with open("countdown_state.json", "r") as f:
            data = json.load(f)

        dpg.set_value("mission_title", data["mission_name"])
        dpg.configure_item("mission_title", show=data["show_mission"])

        dpg.set_value("prefix_text", data["prefix"])
        dpg.set_value("countdown_text", data["countdown_text"])

        # apply countdown text color if provided
        if "countdown_color" in data:
            dpg.configure_item("countdown_text", color=data["countdown_color"])

        # Center mission name
        if data["show_mission"]:
            m_size = dpg.get_text_size(
                data["mission_name"],
                font=fonts.get("large", 0)
            )
            if m_size:
                w = dpg.get_viewport_width()
                indent = max(0, int((w / 2) - (m_size[0] / 2)
                                    + data["centering_offset"]))
                dpg.set_item_indent("mission_title", indent)

        # Center countdown
        count_text = f"{data['prefix']} {data['countdown_text']}"
        count_size = dpg.get_text_size(
            count_text,
            font=fonts.get("huge", 0)
        )
        if count_size:
            w = dpg.get_viewport_width()
            indent = max(0, int((w / 2) - (count_size[0] / 2)
                                + data["centering_offset"]))
            dpg.set_item_indent("count_group", indent)

    except:
        pass


# Viewport
dpg.create_viewport(title="Countdown Clock", width=702, height=232)
dpg.setup_dearpygui()
dpg.set_primary_window("countdown", True)
dpg.show_viewport()

while dpg.is_dearpygui_running():
    update_from_file()
    dpg.render_dearpygui_frame()

dpg.destroy_context()