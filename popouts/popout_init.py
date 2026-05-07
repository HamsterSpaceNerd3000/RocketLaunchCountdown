# Script to initiliaze popout requirements

import sys
from pathlib import Path

def bootstrap():
    """Locates the project root and adds it to the system path."""
    # Fetch project root
    path_root = Path(__file__).resolve().parents[1]
    
    if str(path_root) not in sys.path:
        sys.path.append(str(path_root))
    
    return path_root