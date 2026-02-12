# RLCs countdown handler. It handles all the major countdown logic.

# Imports
from datetime import datetime, timedelta
from typing import TYPE_CHECKING
import math

# Type checking for main
if TYPE_CHECKING:
    from main import CountdownMain

# Function to start the countdown
def start_countdown(countdown: "CountdownMain"):
    if countdown.scrubbed:
        return
    
    t_str = countdown.target_set

    try:
        h, m, s = map(int, t_str.split(":"))
        countdown.target_time = datetime.now() + timedelta(hours=h, minutes=m, seconds=s)
        countdown.running = True
        countdown.hold = False
    except:
        countdown.target_time = datetime.now() + timedelta(minutes=10)
        countdown.running = True
    

# Function to run the countdown
def run_countdown(countdown: "CountdownMain"):
    diff_sec = (countdown.target_time - datetime.now()).total_seconds()
    if countdown.auto_hold and 0 < diff_sec <= countdown.auto_hold_thres:
        countdown.auto_hold = False; countdown.toggle_hold()
    
    if diff_sec <= 0:
        prefix, display_val = "T+", abs(int(diff_sec))
    else: 
        prefix, display_val = countdown.settings["prefix"], math.ceil(diff_sec)
    
    return prefix, display_val