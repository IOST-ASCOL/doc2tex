#!/usr/bin/env python3
# Shortcut so `python cli.py ...` still works without installing the package.
# After `pip install .` you can just run `doc2tex ...` instead.

import sys

from doc2tex.cli import main

if __name__ == '__main__':
    sys.exit(main())
