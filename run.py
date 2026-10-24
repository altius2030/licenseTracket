#!/usr/bin/env python
"""
Run script for License Tracker.
Installs dependencies if needed and launches the app.
"""
import sys
import subprocess
import os

def install_requirements():
    """Install required packages."""
    req_file = os.path.join(os.path.dirname(__file__), "requirements.txt")
    if os.path.exists(req_file):
        print("Installing dependencies...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file])
        print("Dependencies installed.")

def main():
    # Check if customtkinter is available
    try:
        import customtkinter
    except ImportError:
        install_requirements()
    
    # Run the app
    from main import main as app_main
    app_main()

if __name__ == "__main__":
    main()