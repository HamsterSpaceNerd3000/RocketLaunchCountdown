# RLCs settings handler. Handles loading, saving, editings, and defaults

# Imports
import os
import json

config_file = "config.json"

# Default settings
defaults = {
    "mission_name": "PLACEHOLDER",
    "show_mission": True,
    "prefix": "T-",
    "box_bg_color": [0.0, 0.0, 0.0, 255],
    "box_outline": [0.5, 0.5, 0.5, 255],
    "box_border_width": 2,
    "auto_hold_time": "00:02:00",
    "min_hold_reset": "00:00:40",
    "manual_concerns": "NO CONCERNS",
    "touch_screen": False,
    "centering_offset": 0,
    "countdown_text_color": [1.0, 1.0, 1.0, 255],
    "status_text_color": [1.0, 1.0, 1.0, 255],
    "mission_text_color": [1.0, 1.0, 1.0, 255]
}

def load():
    if os.path.exists(config_file):
        try:
            with open(config_file, "r") as f:
                defaults.update(json.load(f))
        except: pass

def save():
    with open(config_file, "w") as f:
        json.dump(defaults, f)