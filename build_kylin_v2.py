#!/usr/bin/env python3
"""Build script for Kylin binary - patches PyInstaller isolated subprocess (WSL2 workaround)."""
import os
os.environ["PYTHONMALLOC"] = "malloc"

# Monkey-patch PyInstaller isolation BEFORE importing PyInstaller
# Must patch BOTH isolated.call AND _parent.call because isolated/__init__.py
# does "from ._parent import call" which creates a separate reference.
import PyInstaller.isolated._parent as isolated_parent
import PyInstaller.isolated as isolated

def inline_call(function, *args, **kwargs):
    """Run target function inline instead of spawning a subprocess (WSL2 workaround)."""
    return function(*args, **kwargs)

isolated_parent.call = inline_call
isolated.call = inline_call
print("Patched: isolated.call AND _parent.call now run inline")

import PyInstaller.__main__
PyInstaller.__main__.run([
    '--onefile',
    '--name', 'cx_new_v6_kylin',
    '--hidden-import', 'tkinter',
    '--hidden-import', '_tkinter',
    '--hidden-import', 'pandas',
    '--hidden-import', 'openpyxl',
    '--hidden-import', 'numpy',
    '--hidden-import', 'openpyxl.styles',
    '--hidden-import', 'openpyxl.utils',
    '--hidden-import', 'openpyxl.cell',
    '--clean',
    '--noconfirm',
    'cx_new_v6_kylin.py'
])
