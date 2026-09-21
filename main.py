"""
College Report Generator - Main Entry Point
Launches the Desktop GUI application.
"""
import sys
import os

# Ensure project root is on sys.path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from app.gui.app import ReportGeneratorApp


def main():
    app = ReportGeneratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
