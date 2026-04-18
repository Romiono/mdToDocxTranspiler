#!/usr/bin/env python3
"""Legacy entry point — delegates to src/transpiler.py."""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from transpiler import main  # noqa: E402

if __name__ == '__main__':
    main()
