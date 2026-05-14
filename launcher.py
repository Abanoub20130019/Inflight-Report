"""
Standalone launcher for Weekly Inflight Report Generator.
Bundles Streamlit so the app runs without requiring a separate Python installation.
"""
import os
import sys

# PyInstaller sets sys._MEIPASS when running from a bundled executable.
# In normal Python, we fall back to the script's directory.
if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

app_path = os.path.join(base_dir, 'app.py')

# Launch Streamlit in non-development mode for a cleaner UI
sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false"]

if __name__ == '__main__':
    import streamlit.web.cli as stcli
    sys.exit(stcli.main())
