"""
Simple launcher for the Document Intelligence GUI
"""

import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from gui_app import main
    main()
except ImportError as e:
    print(f"Missing dependencies: {e}")
    print("Please install requirements: pip install -r requirements.txt")
except Exception as e:
    print(f"Error starting application: {e}")
    input("Press Enter to exit...")