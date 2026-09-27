"""Launcher script for PyInstaller — allows relative imports to work when bundled."""
import sys
import os

# Ensure src/ is on sys.path for PyInstaller bundled mode
_bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
if _bundle_dir not in sys.path:
    sys.path.insert(0, _bundle_dir)

from src.main import main

sys.exit(main())
