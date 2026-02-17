# RLCs error handler. This script has a series of function called by a system when something goes wrong, and the user should known.

# Imports
from datetime import datetime

# Error function
def get_error(code):
    time = datetime.now().strftime("%H:%M:%S")
    errors = {
        100: "Sheetlink not found",
        101: "Network Timeout",
        102: "Invalid Cell Reference",
        404: "Invalid Sheetlink (Check URL)"
    }
    msg = errors.get(code, "Unknown System Error")
    return f"[{time}] ERROR {code}: {msg}"