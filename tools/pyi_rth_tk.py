import os
import sys


def _set_tk_env():
    base = getattr(sys, "_MEIPASS", None)
    if not base:
        return
    tcl_dir = os.path.join(base, "tcl", "tcl8.6")
    tk_dir = os.path.join(base, "tcl", "tk8.6")
    if os.path.isdir(tcl_dir):
        os.environ["TCL_LIBRARY"] = tcl_dir
    if os.path.isdir(tk_dir):
        os.environ["TK_LIBRARY"] = tk_dir


_set_tk_env()
