#!/usr/bin/env python3
# Shortcut so `python web.py` still works without installing the package.
# After `pip install .` you can just run `doc2tex-web` instead.

from doc2tex.web import app, start_server  # noqa: F401  (app is handy for `flask --app web run`)

if __name__ == '__main__':
    start_server()
