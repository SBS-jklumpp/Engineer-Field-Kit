"""
Local override for PyInstaller's tkinter pre-find hook.

The upstream hook drops tkinter when Tcl/Tk auto-detection reports a broken
installation. In this project we bundle Tcl/Tk assets manually from app.spec,
so we intentionally do not clear hook_api.search_dirs here.
"""


def pre_find_module_path(hook_api):
    return
