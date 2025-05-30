'''Package / frozen entry-point.'''

import sys
import os
from pathlib import Path

# Get the application directory
if getattr(sys, 'frozen', False):
    # Running as PyInstaller executable
    application_path = sys._MEIPASS
else:
    # Running as Python script
    application_path = Path(__file__).resolve().parent

# Add the application path to sys.path
sys.path.insert(0, str(application_path))

# Import the main module
import main

if __name__ == "__main__":
    main.main() 