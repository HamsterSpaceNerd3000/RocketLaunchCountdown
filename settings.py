# RLCs settings handler. Handles loading, saving, editings, and defaults

# Imports
import os
import json
from path_handler import get_config_folder

config_file = get_config_folder() / "config.json"

# Default settings
defaults = {
    "mission_name": "PLACEHOLDER",
    "show_mission": True,
    "prefix": "T-",
    "box_bg_color": [0.0, 0.0, 0.0, 255],
    "box_outline": [0.5, 0.5, 0.5, 255],
    "txt_color": [255, 255, 255, 255],
    "box_border_width": 2,
    "auto_hold_time": "00:02:00",
    "min_hold_reset": "00:00:40",
    "manual_concerns": "NO CONCERNS",
    "touch_screen": False,
    "centering_offset": 0,
    "countdown_text_color": [1.0, 1.0, 1.0, 255],
    "status_text_color": [1.0, 1.0, 1.0, 255],
    "mission_text_color": [1.0, 1.0, 1.0, 255],
    "spreadsheet_link": "",
    "weather_sheet_cell": "",
    "range_sheet_cell": "",
    "vehicle_sheet_cell": "",
    "concerns_sheet_cell": "",
    "sheet_gid": ""
}

def load():
    """
    Function to load settings from config file, if the config file is renamed, this will break.
    """
    if os.path.exists(config_file):
        try:
            with open(config_file, "r") as f:
                defaults.update(json.load(f))
        except: pass

def save(data_to_save):
    """
    Function to save changed settings to config file.

    data_to_save: The changed settings you want saved, this can be a single setting or the whole batch.
    """
    with open(config_file, "w") as f:
        json.dump(data_to_save, f, indent=4)