#!/usr/bin/env python3
"""Build script for Kylin binary - patches WSL2 subprocess issue."""
import os
os.environ["PYTHONMALLOC"] = "malloc"

# Monkey-patch PyInstaller isolation BEFORE importing PyInstaller
import PyInstaller.isolated._parent as isolated_parent
_original_call = isolated_parent.call

def inline_call(function, *args, **kwargs):
    """Run target function inline instead of spawning a subprocess (WSL2 workaround)."""
    return function(*args, **kwargs)

isolated_parent.call = inline_call
print("Patched: isolated_python.call now runs inline")

import PyInstaller.__main__
PyInstaller.__main__.run([
    '--onefile', '--name', 'cx_new_v6_kylin_final',
    '--hidden-import', 'tkinter', '--hidden-import', '_tkinter',
    '--hidden-import', 'pandas', '--hidden-import', 'openpyxl',
    '--hidden-import', 'numpy',
    '--clean', '--noconfirm',
    'cx_new_v6_kylin.py'
])
