# Script to handle paths for RLC

# Imports
from pathlib import Path
import os
import sys

def get_root():
    # Bundle check
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).resolve().parent
    
# Establish path
PATH = get_root()

def get_config_folder():
    """Get the RocketLaunchCountdown folder in Documents"""
    documents = Path.home() / "Documents" / "RocketLaunchCountdown"
    documents.mkdir(parents=True, exist_ok=True)
    return documents