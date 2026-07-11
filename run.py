#!/usr/bin/env python3
"""Entry point for ReportSoft - Consolidados."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from main_10 import main

if __name__ == "__main__":
    main()
