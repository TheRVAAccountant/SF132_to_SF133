#!/usr/bin/env python3
"""
Entry point script for SF132 to SF133 Reconciliation Tool.

This script provides a simple launcher for the refactored package.
"""

import sys
from src.sf132_sf133_recon.main import main

if __name__ == "__main__":
    sys.exit(main())