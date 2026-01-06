#!/usr/bin/env python3
"""
Invoice Management System - Main Entry Point

A local Windows desktop application for managing Excel-based invoice data
from multiple testing labs (BV, ITS, TUV, SGS).

Usage:
    python main.py
"""

import sys
import os

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui import main

if __name__ == "__main__":
    main()
