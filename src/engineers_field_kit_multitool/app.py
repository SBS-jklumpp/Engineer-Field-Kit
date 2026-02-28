import csv
import datetime as dt
import html
import json
import os
from pathlib import Path
import queue
import re
import sys
import tempfile
import threading
import time
import tkinter as tk
import webbrowser
from urllib.parse import urljoin
from tkinter import filedialog, messagebox, scrolledtext, ttk

import mistune
import numpy as np
import serial
from serial.tools import list_ports
try:
    from .styles import (
        DARK_ACCENT,
        DARK_BG,
        DARK_BORDER,
        DARK_MUTED,
        DARK_OK,
        DARK_PANEL,
        DARK_PANEL_2,
        DARK_TEXT,
        LIGHT_BG,
        LIGHT_MUTED,
        LIGHT_OK,
        LIGHT_PANEL_2,
        LIGHT_TEXT,
        apply_theme,
    )
except ImportError:
    from styles import (
        DARK_ACCENT,
        DARK_BG,
        DARK_BORDER,
        DARK_MUTED,
        DARK_OK,
        DARK_PANEL,
        DARK_PANEL_2,
        DARK_TEXT,
        LIGHT_BG,
        LIGHT_MUTED,
        LIGHT_OK,
        LIGHT_PANEL_2,
        LIGHT_TEXT,
        apply_theme,
    )

SENSOR_TEST_DIR = r"I:\common\products\SensorTests\SBE83"
WARN_NS = 10.0
FAIL_NS = 20.0
MAX_UNITS_PER_SESSION = 10
MAX_PORTS = 10
PRECAL_TEST_SUBDIR = "PreCalTest"
DEBUG_RESULTS_SUBDIR = "SBE83_Debug"
DEFAULT_OPERATOR = "Justin"
DEFAULT_SALINITY_PSU = "0.0"
COMM_RETRY_TIMEOUT_S = 12.0
COMM_RETRY_INTERVAL_S = 1.0
SAMPLE_RETRY_TIMEOUT_S = 8.0
SAMPLE_RETRY_INTERVAL_S = 0.6

TSR_FIELDS = [
    "red_phase",
    "blue_phase",
    "red_blue_phase",
    "red_voltage",
    "blue_voltage",
    "raw_temp_voltage",
    "red_pll_voltage",
    "blue_pll_voltage",
    "red_blue_pll_voltage",
    "electronics_temp_voltage",
]

DEFAULT_FIELD_DESCRIPTIONS = {
    "red_phase": "Red phase",
    "blue_phase": "Blue phase",
    "red_blue_phase": "Red-Blue phase",
    "red_voltage": "Red voltage",
    "blue_voltage": "Blue voltage",
    "raw_temp_voltage": "Raw temperature voltage",
    "red_pll_voltage": "Red PLL voltage",
    "blue_pll_voltage": "Blue PLL voltage",
    "red_blue_pll_voltage": "Red-Blue PLL voltage",
    "electronics_temp_voltage": "Electronics temperature voltage",
}

LIVE_PLOT_FIELDS = {
    "Red Phase": "red_phase",
    "Blue Phase": "blue_phase",
    "Red Voltage": "red_voltage",
    "Blue Voltage": "blue_voltage",
    "Raw Temp Voltage": "raw_temp_voltage",
    "Electronics Temp Voltage": "electronics_temp_voltage",
    "Red PLL Voltage": "red_pll_voltage",
    "Blue PLL Voltage": "blue_pll_voltage",
}

SESSION_PLOT_FIELDS = {
    "Red Noise (ns)": "red_noise_ns",
    "Blue Noise (ns)": "blue_noise_ns",
    "Red Voltage Std": "red_voltage_std",
    "Red Voltage Avg": "red_voltage_avg",
    "Blue Voltage Std": "blue_voltage_std",
    "Blue Voltage Avg": "blue_voltage_avg",
    "Red PLL Voltage Std": "red_pll_voltage_std",
    "Red PLL Voltage Avg": "red_pll_voltage_avg",
    "Blue PLL Voltage Std": "blue_pll_voltage_std",
    "Blue PLL Voltage Avg": "blue_pll_voltage_avg",
    "Raw Temp Voltage Std": "raw_temp_voltage_std",
    "Raw Temp Voltage Avg": "raw_temp_voltage_avg",
    "Electronics Temp Voltage Std": "electronics_temp_voltage_std",
    "Electronics Temp Voltage Avg": "electronics_temp_voltage_avg",
}

LIVE_PORT_COLORS = [
    "#2563eb",
    "#dc2626",
    "#059669",
    "#d97706",
    "#7c3aed",
    "#0891b2",
    "#be123c",
    "#4f46e5",
    "#65a30d",
    "#0f766e",
]

BAUD_OPTIONS = [1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400, 460800, 921600]
UNIT_SCALE_FACTORS = {
    "raw": 1.0,
    "milli": 1000.0,
    "micro": 1000000.0,
    "kilo": 0.001,
}

DEFAULT_SAMPLE_SETUP = {
    "tsr_fields": list(TSR_FIELDS),
    "default_field_descriptions": dict(DEFAULT_FIELD_DESCRIPTIONS),
    "live_plot_fields": dict(LIVE_PLOT_FIELDS),
    "session_plot_fields": dict(SESSION_PLOT_FIELDS),
    "unit_scale_factors": dict(UNIT_SCALE_FACTORS),
}

APP_NAME = "Seabird Sensor Digital Workbench"
APP_SUBTITLE = "Digital Sensor Workbench"
APP_TAGLINE = "Serial | Plot | Analyze | Debug"
APP_VERSION = "v1.1.2"
APP_AUTHOR = "Justin Klumpp"
APP_COMPANY = "Seabird Scientific"
APP_CONTACT_EMAIL = "jklumpp@seabird.com"
APP_ABOUT_SUMMARY = (
    "Seabird Sensor Digital Workbench is a desktop tool for connecting to digital sensors, "
    "running repeatable bench tests, and visualizing both live and session data. "
    "It combines serial communications, parser setup, plotting, and run logging in one workflow "
    "to speed engineering diagnostics while keeping outputs traceable."
)


def _resolve_app_config_file():
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        return os.path.join(exe_dir, "sbs_dsw_config.json")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "engineers_field_kit_multitool_config.json")


APP_CONFIG_FILE = _resolve_app_config_file()


class SBE83GuiApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.app_config = self._load_app_config()
        self._last_normal_geometry = None
        layout_state = self.app_config.get("layout_state", {})
        if isinstance(layout_state, dict):
            saved_normal = layout_state.get("normal_window_geometry")
            if isinstance(saved_normal, str) and saved_normal.strip():
                self._last_normal_geometry = saved_normal.strip()
        self._set_initial_window_size()
        self._apply_saved_window_state()
        self.sample_setup_defaults = self._load_sample_setup_defaults(self.app_config.get("sample_setup_defaults"))
        self.tsr_fields = list(self.sample_setup_defaults["tsr_fields"])
        self.default_field_descriptions = dict(self.sample_setup_defaults["default_field_descriptions"])
        self.default_live_plot_fields = dict(self.sample_setup_defaults["live_plot_fields"])
        self.default_session_plot_fields = dict(self.sample_setup_defaults["session_plot_fields"])
        self.unit_scale_factors = dict(self.sample_setup_defaults["unit_scale_factors"])
        self.dark_mode_var = tk.BooleanVar(value=bool(self.app_config.get("dark_mode", True)))
        self.port_station_collapsed_var = tk.BooleanVar(value=bool(self.app_config.get("port_station_collapsed", False)))
        self.config_mode_var = tk.BooleanVar(value=bool(self.app_config.get("config_mode", False)))
        self.debug_mode_var = tk.BooleanVar(value=bool(self.app_config.get("debug_mode", False)))
        self.mode_choice_var = tk.StringVar(value="debug" if bool(self.debug_mode_var.get()) else "production")
        self.mode_status_var = tk.StringVar(value="")
        stored_non_debug_root = str(self.app_config.get("non_debug_results_root", "")).strip()
        self.non_debug_results_root = stored_non_debug_root or SENSOR_TEST_DIR
        self.test_setup_visible = not bool(self.app_config.get("test_setup_collapsed", False))
        self.setup_manual_override = False
        apply_theme(self.root, dark_mode=bool(self.dark_mode_var.get()))

        self.serial_pool = {}  # port -> serial.Serial
        self.available_ports = []
        self.port_slots = {}
        self.debug_tabs = {}
        self.debug_max_lines = 1500

        self.session_start = dt.datetime.now()
        self.session_id = self.session_start.strftime("%Y%m%d_%H%M%S")
        self.session_rows = []
        self.session_serials = set()
        self.session_dir = os.path.join(self.non_debug_results_root, "sessions", PRECAL_TEST_SUBDIR)
        os.makedirs(self.session_dir, exist_ok=True)
        self.session_csv = os.path.join(self.session_dir, f"sbe83_session_{self.session_id}.csv")
        self.live_run_series_by_port = {}  # port -> field -> list[float]
        self.live_run_total_samples_by_port = {}  # port -> int
        self.live_run_serial_by_port = {}  # port -> serial label
        self.live_port_colors = {}  # port -> hex color
        self.run_in_progress = False
        self.run_thread = None
        self.run_threads = {}
        self.active_run_ports = set()
        self.run_state_lock = threading.Lock()
        self.tree_sort_state = {}
        self.delimiter_var = tk.StringVar(value=",")
        self.example_sample_var = tk.StringVar()
        self.measureand_rows = []
        self.live_plot_fields = dict(self.default_live_plot_fields)
        self.base_session_plot_fields = dict(self.default_session_plot_fields)
        self.session_plot_fields = dict(self.default_session_plot_fields)
        self.sample_field_defs = self._load_persisted_sample_field_defs(self.app_config.get("sample_field_defs"))
        if not self.sample_field_defs:
            self.sample_field_defs = self._default_sample_field_defs()
        self.field_meta_by_key = {}
        self.derived_fields = []
        self.live_visible_ports = set()
        self.stream_threads = {}
        self.stream_stop_events = {}
        self.stream_enabled = {}
        self.manual_capture_rows = []
        self.reference_session_rows = []
        self.reference_session_path = ""
        self.ui_event_queue = queue.Queue()
        self.shutdown_event = threading.Event()
        self.profile_dir = os.path.join(self.session_dir, "profiles")
        os.makedirs(self.profile_dir, exist_ok=True)
        self.parser_trim_prefix_var = tk.StringVar(value="")
        self.parser_token_start_var = tk.IntVar(value=0)
        self.parser_regex_var = tk.StringVar(value="")
        self.sample_command_var = tk.StringVar(value="tsr")
        self.baudrate_var = tk.IntVar(value=9600)
        self.live_autoscale_var = tk.BooleanVar(value=True)
        self.live_ymin_var = tk.StringVar(value="")
        self.live_ymax_var = tk.StringVar(value="")
        self.live_show_points_var = tk.BooleanVar(value=True)
        self.live_visible_only_var = tk.BooleanVar(value=False)
        self.live_x_start_var = tk.IntVar(value=1)
        self.live_x_end_var = tk.IntVar(value=0)
        self.console_detached = False
        self.console_send_cr_var = tk.BooleanVar(value=True)
        self.console_send_lf_var = tk.BooleanVar(value=True)
        self.console_display_mode_var = tk.StringVar(value="ascii")
        self.batch_runs_remaining_by_port = {}
        self.runs_left_var = tk.StringVar(value="Runs left: n/a")
        self.sample_format_expanded = False
        self._dock_console_callback = self.root.register(self._dock_console_tab)
        self.about_window = None
        self._layout_save_after_id = None

        self._build_ui()
        self._apply_results_root(self.non_debug_results_root, log_change=False)
        self._set_test_setup_visibility(self.test_setup_visible, persist=False)
        if self.debug_mode_var.get():
            self._set_debug_mode(True, persist=False, announce=False)
        else:
            self._update_results_root_controls()
            self._update_mode_status_text()
        self.refresh_ports()
        self._apply_theme(persist=False)
        self.root.after(60, self._process_ui_events)
        self.log(f"Session started: {self.session_id}")
        self.log(f"Session summary file: {self.session_csv}")

    def _load_app_config(self):
        try:
            with open(APP_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _load_sample_setup_defaults(self, payload):
        defaults = dict(DEFAULT_SAMPLE_SETUP)
        result = {
            "tsr_fields": list(defaults["tsr_fields"]),
            "default_field_descriptions": dict(defaults["default_field_descriptions"]),
            "live_plot_fields": dict(defaults["live_plot_fields"]),
            "session_plot_fields": dict(defaults["session_plot_fields"]),
            "unit_scale_factors": dict(defaults["unit_scale_factors"]),
        }
        if not isinstance(payload, dict):
            return result

        tsr_fields = payload.get("tsr_fields")
        if isinstance(tsr_fields, list):
            clean = [str(x).strip() for x in tsr_fields if str(x).strip()]
            if clean:
                result["tsr_fields"] = clean

        field_desc = payload.get("default_field_descriptions")
        if isinstance(field_desc, dict):
            clean = {}
            for k, v in field_desc.items():
                key = str(k).strip()
                desc = str(v).strip()
                if key and desc:
                    clean[key] = desc
            if clean:
                result["default_field_descriptions"] = clean

        live_plot = payload.get("live_plot_fields")
        if isinstance(live_plot, dict):
            clean = {}
            for label, key in live_plot.items():
                label_text = str(label).strip()
                key_text = str(key).strip()
                if label_text and key_text:
                    clean[label_text] = key_text
            if clean:
                result["live_plot_fields"] = clean

        session_plot = payload.get("session_plot_fields")
        if isinstance(session_plot, dict):
            clean = {}
            for label, key in session_plot.items():
                label_text = str(label).strip()
                key_text = str(key).strip()
                if label_text and key_text:
                    clean[label_text] = key_text
            if clean:
                result["session_plot_fields"] = clean

        scale_factors = payload.get("unit_scale_factors")
        if isinstance(scale_factors, dict):
            clean = {}
            for name, factor in scale_factors.items():
                scale_name = str(name).strip()
                try:
                    scale_val = float(factor)
                except Exception:
                    continue
                if scale_name:
                    clean[scale_name] = scale_val
            if clean:
                result["unit_scale_factors"] = clean
        return result

    def _load_persisted_sample_field_defs(self, payload):
        if not isinstance(payload, list) or not payload:
            return []
        defs = []
        for i, item in enumerate(payload):
            if not isinstance(item, dict):
                continue
            idx = item.get("index", i)
            try:
                idx = int(idx)
            except Exception:
                idx = i
            key = self._sanitize_measureand_key(item.get("key", f"field_{i + 1}"), idx)
            description = str(item.get("description", key.replace("_", " ").title())).strip() or key.replace("_", " ").title()
            scale_name = str(item.get("scale", "raw")).strip() or "raw"
            if scale_name not in self.unit_scale_factors:
                scale_name = "raw"
            defs.append(
                {
                    "index": idx,
                    "key": key,
                    "description": description,
                    "unit": str(item.get("unit", "")).strip(),
                    "scale": scale_name,
                    "min_val": str(item.get("min_val", "")).strip(),
                    "max_val": str(item.get("max_val", "")).strip(),
                    "stuck_n": str(item.get("stuck_n", "")).strip(),
                    "expr": str(item.get("expr", "")).strip(),
                    "plot_live": bool(item.get("plot_live", False)),
                    "plot_session": bool(item.get("plot_session", False)),
                    "live_default": bool(item.get("live_default", i == 0)),
                }
            )
        return sorted(defs, key=lambda d: d.get("index", 0))

    def _snapshot_sample_setup_defaults(self):
        return {
            "tsr_fields": [d["key"] for d in self.sample_field_defs],
            "default_field_descriptions": {d["key"]: d["description"] for d in self.sample_field_defs},
            "live_plot_fields": {d["description"]: d["key"] for d in self.sample_field_defs if d.get("plot_live")},
            "session_plot_fields": dict(self.base_session_plot_fields),
            "unit_scale_factors": dict(self.unit_scale_factors),
        }

    def _persist_sample_setup_defaults(self):
        self.app_config["sample_field_defs"] = list(self.sample_field_defs)
        self.app_config["sample_setup_defaults"] = self._snapshot_sample_setup_defaults()
        self._save_app_config()

    def _save_app_config(self):
        data = dict(self.app_config) if isinstance(self.app_config, dict) else {}
        data["dark_mode"] = bool(self.dark_mode_var.get())
        data["port_station_collapsed"] = bool(self.port_station_collapsed_var.get())
        if hasattr(self, "config_mode_var"):
            data["config_mode"] = bool(self.config_mode_var.get())
        if hasattr(self, "debug_mode_var"):
            data["debug_mode"] = bool(self.debug_mode_var.get())
        data["non_debug_results_root"] = str(getattr(self, "non_debug_results_root", SENSOR_TEST_DIR)).strip() or SENSOR_TEST_DIR
        data["test_setup_collapsed"] = not bool(getattr(self, "test_setup_visible", True))
        data["layout_state"] = self._capture_layout_state()
        if "sample_setup_defaults" not in data:
            data["sample_setup_defaults"] = self._snapshot_sample_setup_defaults()
        if "sample_field_defs" not in data:
            data["sample_field_defs"] = list(self.sample_field_defs)
        try:
            with open(APP_CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.app_config = data
        except Exception as exc:
            self.log(f"Config save failed: {exc}")

    def _console_font(self):
        return ("Consolas", 10)

    def _set_initial_window_size(self):
        try:
            screen_w = int(self.root.winfo_screenwidth())
            screen_h = int(self.root.winfo_screenheight())
        except Exception:
            self.root.geometry("960x600")
            return
        target_w = max(860, min(960, int(screen_w * 0.5)))
        target_h = max(540, min(600, int(screen_h * 0.5)))
        target_w = min(target_w, max(800, screen_w - 60))
        target_h = min(target_h, max(520, screen_h - 80))
        pos_x = max(0, (screen_w - target_w) // 2)
        pos_y = max(0, (screen_h - target_h) // 3)
        self.root.geometry(f"{target_w}x{target_h}+{pos_x}+{pos_y}")

    def _apply_saved_window_state(self):
        layout = self.app_config.get("layout_state", {})
        if not isinstance(layout, dict):
            return
        state = str(layout.get("window_state", "")).lower()
        normal_geometry = layout.get("normal_window_geometry")
        if isinstance(normal_geometry, str) and normal_geometry.strip():
            self._last_normal_geometry = normal_geometry.strip()

        geometry = layout.get("window_geometry")
        if state == "zoomed" and isinstance(self._last_normal_geometry, str) and self._last_normal_geometry.strip():
            geometry = self._last_normal_geometry.strip()
        if isinstance(geometry, str) and geometry.strip():
            try:
                self.root.geometry(geometry.strip())
            except Exception:
                pass
        if state == "zoomed":
            try:
                self.root.state("zoomed")
            except Exception:
                try:
                    self.root.attributes("-zoomed", True)
                except Exception:
                    pass

    def _capture_layout_state(self):
        out = {}
        try:
            state = str(self.root.state()).lower()
        except Exception:
            state = "normal"
        if state == "iconic":
            state = "normal"
        out["window_state"] = state
        try:
            current_geometry = self.root.winfo_geometry()
            out["window_geometry"] = current_geometry
            if state == "normal" and isinstance(current_geometry, str) and current_geometry.strip():
                self._last_normal_geometry = current_geometry.strip()
            if isinstance(self._last_normal_geometry, str) and self._last_normal_geometry.strip():
                out["normal_window_geometry"] = self._last_normal_geometry.strip()
        except Exception:
            pass
        if hasattr(self, "main_split"):
            try:
                out["main_split_sash"] = int(self.main_split.sashpos(0))
            except Exception:
                pass
        if hasattr(self, "plot_split"):
            try:
                out["plot_split_sash"] = int(self.plot_split.sashpos(0))
            except Exception:
                pass
        out["sample_format_expanded"] = bool(getattr(self, "sample_format_expanded", False))
        if hasattr(self, "main_notebook"):
            try:
                out["active_tab_text"] = str(self.main_notebook.tab(self.main_notebook.select(), "text"))
            except Exception:
                pass
        return out

    def _schedule_layout_save(self, delay_ms=700):
        if not hasattr(self, "root"):
            return
        if self._layout_save_after_id is not None:
            try:
                self.root.after_cancel(self._layout_save_after_id)
            except Exception:
                pass
        try:
            self._layout_save_after_id = self.root.after(delay_ms, self._flush_layout_save)
        except Exception:
            self._layout_save_after_id = None

    def _flush_layout_save(self):
        self._layout_save_after_id = None
        try:
            self._save_app_config()
        except Exception:
            pass

    def _set_split_defaults(self):
        profile = self._layout_profile()
        if hasattr(self, "main_split"):
            self.root.update_idletasks()
            try:
                total_h = int(self.main_split.winfo_height())
            except Exception:
                total_h = 0
            if total_h >= 220:
                min_notebook_h = max(int(total_h / 3), 320 if profile == "wide" else 280)
                default_top_h = max(220, int(total_h * (0.30 if profile == "wide" else 0.42)))
                default_top_h = min(default_top_h, total_h - min_notebook_h)
                try:
                    self.main_split.sashpos(0, default_top_h)
                except Exception:
                    pass
        self._enforce_main_notebook_min_height()
        self._adjust_plot_split_height(force=True)

    def _main_split_limits(self):
        if not hasattr(self, "main_split"):
            return None
        try:
            total_h = int(self.main_split.winfo_height())
        except Exception:
            return None
        if total_h < 200:
            return None
        min_notebook_h = max(int(total_h / 3), 320 if self._layout_profile() == "wide" else 280)
        min_top_h = 130
        max_top_h = max(min_top_h, total_h - min_notebook_h)
        return (min_top_h, max_top_h)

    def _plot_split_limits(self):
        if not hasattr(self, "plot_split"):
            return None
        try:
            total_h = int(self.plot_split.winfo_height())
        except Exception:
            return None
        if total_h < 220:
            return None
        min_top_h = 110
        max_top_h = max(min_top_h, total_h - 170)
        return (min_top_h, max_top_h)

    def _apply_startup_layout(self):
        self._set_split_defaults()
        self._restore_layout_from_config()
        self._update_config_mode_button()
        if self.config_mode_var.get():
            self._set_config_mode(True, persist=False)

    def _restore_layout_from_config(self):
        layout = self.app_config.get("layout_state", {})
        if not isinstance(layout, dict):
            return
        self.root.update_idletasks()
        main_sash = layout.get("main_split_sash")
        if isinstance(main_sash, int) and hasattr(self, "main_split"):
            try:
                limits = self._main_split_limits()
                if limits is not None:
                    lo, hi = limits
                    self.main_split.sashpos(0, max(lo, min(hi, main_sash)))
            except Exception:
                pass
        tab_text = str(layout.get("active_tab_text", "")).strip().lower()
        if tab_text and hasattr(self, "main_notebook"):
            for tab_id in self.main_notebook.tabs():
                try:
                    if str(self.main_notebook.tab(tab_id, "text")).strip().lower() == tab_text:
                        self.main_notebook.select(tab_id)
                        break
                except Exception:
                    continue

    def _on_root_configure(self, _event=None):
        try:
            if str(self.root.state()).lower() == "normal":
                geom = self.root.winfo_geometry()
                if isinstance(geom, str) and geom.strip():
                    self._last_normal_geometry = geom.strip()
        except Exception:
            pass
        self._enforce_main_notebook_min_height()
        self._adjust_plot_split_height(force=False)
        self._schedule_layout_save(delay_ms=800)

    def _enforce_main_notebook_min_height(self):
        if not hasattr(self, "main_split"):
            return
        try:
            total_h = int(self.main_split.winfo_height())
            pos = int(self.main_split.sashpos(0))
        except Exception:
            return
        if total_h < 200:
            return
        min_top_h = 130
        min_notebook_h = max(int(total_h / 3), 320 if self._layout_profile() == "wide" else 280)
        max_top_h = max(min_top_h, total_h - min_notebook_h)
        if pos < min_top_h:
            try:
                self.main_split.sashpos(0, min_top_h)
            except Exception:
                pass
            return
        if pos > max_top_h:
            try:
                self.main_split.sashpos(0, max_top_h)
            except Exception:
                pass

    def _adjust_plot_split_height(self, force=False):
        if not hasattr(self, "plot_split"):
            return
        try:
            total_h = int(self.plot_split.winfo_height())
            pos = int(self.plot_split.sashpos(0))
        except Exception:
            return
        if total_h < 220:
            return
        profile = self._layout_profile()

        if self.sample_format_expanded:
            min_top_h = max(int(total_h / 3), 240 if profile == "wide" else 220)
            target_top_h = max(min_top_h, int(total_h * (0.36 if profile == "wide" else 0.42)))
            target_top_h = min(target_top_h, total_h - 140)
            if force:
                try:
                    self.plot_split.sashpos(0, target_top_h)
                except Exception:
                    pass
            elif pos < min_top_h:
                try:
                    self.plot_split.sashpos(0, min_top_h)
                except Exception:
                    pass
        elif force:
            target_top_h = min(max(130, int(total_h * (0.18 if profile == "wide" else 0.22))), total_h - 160)
            try:
                self.plot_split.sashpos(0, target_top_h)
            except Exception:
                pass

    def _layout_profile(self):
        try:
            state = str(self.root.state()).lower()
        except Exception:
            state = ""
        if state == "zoomed":
            return "wide"
        try:
            w = int(self.root.winfo_width())
        except Exception:
            w = 0
        return "wide" if w >= 1500 else "compact"

    def _theme_colors(self):
        if self.dark_mode_var.get():
            return {
                "bg": DARK_BG,
                "panel": DARK_PANEL,
                "panel_2": DARK_PANEL_2,
                "border": DARK_BORDER,
                "fg": DARK_TEXT,
                "muted": DARK_MUTED,
                "accent": DARK_ACCENT,
                "ok": DARK_OK,
                "hero_bg": "#1f2a3a",
                "hero_border": "#5d7ea0",
                "hero_title": "#f8fbff",
                "canvas": "#0b1220",
                "entry": DARK_PANEL,
            }
        return {
            "bg": LIGHT_BG,
            "panel": "#ffffff",
            "panel_2": LIGHT_PANEL_2,
            "border": "#9bb0c7",
            "fg": LIGHT_TEXT,
            "muted": LIGHT_MUTED,
            "accent": "#0f62fe",
            "ok": LIGHT_OK,
            "hero_bg": "#dce8f6",
            "hero_border": "#7aa0c7",
            "hero_title": "#12253a",
            "canvas": "#f3f8ff",
            "entry": "#ffffff",
        }

    def _status_colors(self):
        if self.dark_mode_var.get():
            return {
                "DISCONNECTED": ("#5f6b7a", "#ffffff"),
                "CONNECTED": ("#0f62fe", "#ffffff"),
                "RUNNING": ("#f59e0b", "#111827"),
                "COMPLETE": ("#14b8a6", "#042f2e"),
                "PASS": ("#22c55e", "#052e16"),
                "WARN": ("#fb923c", "#111827"),
                "FAIL": ("#ef4444", "#ffffff"),
                "ERROR": ("#a855f7", "#ffffff"),
            }
        return {
            "DISCONNECTED": ("#7c8796", "#ffffff"),
            "CONNECTED": ("#1d4ed8", "#ffffff"),
            "RUNNING": ("#f59e0b", "#111827"),
            "COMPLETE": ("#0d9488", "#ffffff"),
            "PASS": ("#16a34a", "#ffffff"),
            "WARN": ("#ea580c", "#ffffff"),
            "FAIL": ("#dc2626", "#ffffff"),
            "ERROR": ("#7e22ce", "#ffffff"),
        }

    def _apply_theme_recursive(self, widget, colors):
        try:
            widget.configure(bg=colors["bg"], fg=colors["fg"])
        except Exception:
            pass
        try:
            widget.configure(background=colors["bg"], foreground=colors["fg"])
        except Exception:
            pass
        for child in widget.winfo_children():
            self._apply_theme_recursive(child, colors)

    def _apply_direct_widget_theme(self, colors):
        if hasattr(self, "log_box"):
            self.log_box.configure(bg=colors["entry"], fg=colors["fg"], insertbackground=colors["fg"])
        if hasattr(self, "live_text"):
            self.live_text.configure(bg=colors["entry"], fg=colors["fg"], insertbackground=colors["fg"])
        if hasattr(self, "live_canvas"):
            self.live_canvas.configure(bg=colors["canvas"])
        if hasattr(self, "measureand_canvas"):
            self.measureand_canvas.configure(bg=colors["canvas"])
        for info in getattr(self, "debug_tabs", {}).values():
            try:
                info["text"].configure(bg=colors["entry"], fg=colors["fg"], insertbackground=colors["fg"])
            except Exception:
                pass
        if hasattr(self, "hero_frame"):
            self.hero_frame.configure(bg=colors["hero_bg"], highlightbackground=colors["hero_border"])
        if hasattr(self, "hero_title_label"):
            self.hero_title_label.configure(bg=colors["hero_bg"], fg=colors["hero_title"])
        if hasattr(self, "hero_right"):
            self.hero_right.configure(bg=colors["hero_bg"])
        if hasattr(self, "hero_right_top"):
            self.hero_right_top.configure(bg=colors["hero_bg"])
        for attr in (
            "hero_help_link",
            "hero_about_link",
            "hero_results_link",
            "hero_live_link",
            "hero_config_link",
            "hero_reset_link",
        ):
            if hasattr(self, attr):
                getattr(self, attr).configure(bg=colors["hero_bg"], fg=colors["muted"])
        if hasattr(self, "hero_version_label"):
            self.hero_version_label.configure(bg=colors["hero_bg"], fg=colors["accent"])
        if hasattr(self, "hero_tagline_label"):
            self.hero_tagline_label.configure(bg=colors["hero_bg"], fg=colors["muted"])
        if hasattr(self, "port_grid_empty_label"):
            self.port_grid_empty_label.configure(bg=colors["bg"], fg=colors["muted"])
        for slot in getattr(self, "port_slots", {}).values():
            slot["card"].configure(bg=colors["panel"], highlightbackground=colors["border"])
            slot["slot_label"].configure(bg=colors["panel"], fg=colors["fg"])
            slot["port_label"].configure(bg=colors["panel"], fg=colors["fg"])
            slot["serial_label"].configure(bg=colors["panel"], fg=colors["muted"])
            state_bg, state_fg = self.status_color(slot["state_var"].get())
            slot["state_label"].configure(bg=state_bg, fg=state_fg)

    def _apply_theme(self, persist=True):
        apply_theme(self.root, dark_mode=bool(self.dark_mode_var.get()))
        font = self._console_font()
        if hasattr(self, "live_text"):
            self.live_text.configure(font=font)
        if hasattr(self, "log_box"):
            self.log_box.configure(font=font)
        for info in getattr(self, "debug_tabs", {}).values():
            try:
                info["text"].configure(font=font)
            except Exception:
                pass
        colors = self._theme_colors()
        self._apply_theme_recursive(self.root, colors)
        self._apply_direct_widget_theme(colors)
        if persist:
            self._save_app_config()

    def _on_toggle_dark_mode(self):
        self._apply_theme(persist=True)

    def _on_header_link_enter(self, widget):
        try:
            widget.configure(fg=self._theme_colors()["accent"])
        except Exception:
            pass

    def _on_header_link_leave(self, widget):
        try:
            if hasattr(self, "hero_config_link") and widget is self.hero_config_link and bool(self.config_mode_var.get()):
                widget.configure(fg=self._theme_colors()["accent"])
                return
            widget.configure(fg=self._theme_colors()["muted"])
        except Exception:
            pass

    def _update_config_mode_button(self):
        if not hasattr(self, "hero_config_link"):
            return
        enabled = bool(self.config_mode_var.get())
        self.hero_config_link.configure(fg=self._theme_colors()["accent"] if enabled else self._theme_colors()["muted"])

    def _set_config_mode(self, enabled, persist=True):
        self.config_mode_var.set(bool(enabled))
        self._update_config_mode_button()
        if enabled:
            try:
                self.main_notebook.select(self.sample_setup_tab)
            except Exception:
                pass
        else:
            self._focus_live_layout(update_mode=False)
        if persist:
            self._save_app_config()

    def _toggle_config_mode(self):
        self._set_config_mode(not bool(self.config_mode_var.get()), persist=True)

    def _focus_results_layout(self):
        if bool(self.config_mode_var.get()):
            self.config_mode_var.set(False)
            self._update_config_mode_button()
        if hasattr(self, "main_notebook") and hasattr(self, "test_tab"):
            try:
                self.main_notebook.select(self.test_tab)
            except Exception:
                pass
        if not hasattr(self, "main_split"):
            return
        self.root.update_idletasks()
        try:
            total_h = int(self.main_split.winfo_height())
        except Exception:
            return
        if total_h < 220:
            return
        notebook_target_h = max(int(total_h * (0.78 if self._layout_profile() == "wide" else 0.72)), int(total_h / 3), 320)
        notebook_target_h = min(notebook_target_h, total_h - 140)
        top_h = total_h - notebook_target_h
        limits = self._main_split_limits()
        if limits is not None:
            lo, hi = limits
            top_h = max(lo, min(hi, top_h))
        try:
            self.main_split.sashpos(0, top_h)
        except Exception:
            pass
        self._save_app_config()

    def _focus_live_layout(self, update_mode=True):
        if update_mode and bool(self.config_mode_var.get()):
            self.config_mode_var.set(False)
            self._update_config_mode_button()
        if hasattr(self, "main_notebook") and hasattr(self, "plot_tab"):
            try:
                self.main_notebook.select(self.plot_tab)
            except Exception:
                pass
        if hasattr(self, "main_split"):
            self.root.update_idletasks()
            try:
                total_h = int(self.main_split.winfo_height())
            except Exception:
                total_h = 0
            if total_h >= 220:
                notebook_target_h = max(int(total_h * (0.82 if self._layout_profile() == "wide" else 0.76)), int(total_h / 3), 340)
                notebook_target_h = min(notebook_target_h, total_h - 130)
                top_h = total_h - notebook_target_h
                limits = self._main_split_limits()
                if limits is not None:
                    lo, hi = limits
                    top_h = max(lo, min(hi, top_h))
                try:
                    self.main_split.sashpos(0, top_h)
                except Exception:
                    pass
        self._save_app_config()

    def _reset_layout(self):
        if getattr(self, "console_detached", False):
            try:
                self._dock_console_tab()
            except Exception:
                pass
        if bool(self.config_mode_var.get()):
            self.config_mode_var.set(False)
            self._update_config_mode_button()
        try:
            state = str(self.root.state()).lower()
        except Exception:
            state = ""
        if state != "zoomed":
            self._set_initial_window_size()
        self.root.after(40, self._set_split_defaults)
        self.root.after(80, self._save_app_config)

    def _apply_port_station_visibility(self, persist=True):
        if not hasattr(self, "port_station_frame"):
            return
        if self.port_station_collapsed_var.get():
            self.port_station_frame.pack_forget()
            if hasattr(self, "port_station_toggle_btn"):
                self.port_station_toggle_btn.configure(text="Show Port Station")
        else:
            target = self.setup_frame
            if not bool(getattr(self, "test_setup_visible", True)) and hasattr(self, "actions_frame"):
                target = self.actions_frame
            self.port_station_frame.pack(fill=tk.X, pady=2, before=target)
            if hasattr(self, "port_station_toggle_btn"):
                self.port_station_toggle_btn.configure(text="Hide Port Station")
        if persist:
            self._save_app_config()

    def _toggle_port_station_visibility(self):
        self.port_station_collapsed_var.set(not self.port_station_collapsed_var.get())
        self._apply_port_station_visibility(persist=True)

    def _fit_main_split_to_top_content(self):
        if not hasattr(self, "main_split") or not hasattr(self, "top_frame"):
            return
        self.root.update_idletasks()
        try:
            total_h = int(self.main_split.winfo_height())
            req_top_h = int(self.top_frame.winfo_reqheight()) + 12
        except Exception:
            return
        if total_h < 220:
            return
        limits = self._main_split_limits()
        if limits is None:
            return
        lo, hi = limits
        target = max(lo, min(hi, req_top_h))
        try:
            self.main_split.sashpos(0, target)
        except Exception:
            pass

    def _set_test_setup_visibility(self, visible, auto=False, persist=True):
        visible = bool(visible)
        self.test_setup_visible = visible
        if visible:
            if not self.setup_frame.winfo_manager():
                self.setup_frame.pack(fill=tk.X, pady=4, before=self.actions_frame)
        else:
            if self.setup_frame.winfo_manager():
                self.setup_frame.pack_forget()
        if hasattr(self, "setup_toggle_btn"):
            self.setup_toggle_btn.configure(text="Hide Test Setup" if visible else "Show Test Setup")
        self._apply_port_station_visibility(persist=False)
        self.root.after(40, self._fit_main_split_to_top_content)
        if auto:
            self.log("Test Setup auto-hidden after all fields were entered.")
        if persist:
            self._save_app_config()

    def _toggle_test_setup_visibility(self):
        next_visible = not bool(self.test_setup_visible)
        if next_visible:
            self.setup_manual_override = True
        else:
            self.setup_manual_override = False
        self._set_test_setup_visibility(next_visible, auto=False, persist=True)

    def _build_ui(self):
        self.root.minsize(860, 540)
        self.hero_frame = tk.Frame(self.root, bg=DARK_PANEL_2, highlightthickness=1, highlightbackground=DARK_BORDER)
        self.hero_frame.pack(fill=tk.X, padx=8, pady=(8, 0))
        self.hero_title_label = tk.Label(
            self.hero_frame,
            text=APP_NAME,
            bg=DARK_PANEL_2,
            fg=DARK_TEXT,
            font=("Segoe UI Semibold", 15),
            padx=12,
            pady=8,
        )
        self.hero_title_label.pack(side=tk.LEFT)
        self.hero_right = tk.Frame(self.hero_frame, bg=DARK_PANEL_2)
        self.hero_right.pack(side=tk.RIGHT, padx=12)
        self.hero_right_top = tk.Frame(self.hero_right, bg=DARK_PANEL_2)
        self.hero_right_top.pack(anchor="e")
        self.hero_help_link = tk.Label(
            self.hero_right_top,
            text="Help",
            bg=DARK_PANEL_2,
            fg=DARK_MUTED,
            font=("Segoe UI Semibold", 11),
            cursor="hand2",
            padx=0,
            pady=0,
        )
        self.hero_help_link.pack(side=tk.LEFT, padx=(0, 12))
        self.hero_help_link.bind("<Button-1>", lambda _e: self.open_readme_help())
        self.hero_help_link.bind("<Enter>", lambda _e, w=self.hero_help_link: self._on_header_link_enter(w))
        self.hero_help_link.bind("<Leave>", lambda _e, w=self.hero_help_link: self._on_header_link_leave(w))
        self.hero_about_link = tk.Label(
            self.hero_right_top,
            text="About",
            bg=DARK_PANEL_2,
            fg=DARK_MUTED,
            font=("Segoe UI Semibold", 11),
            cursor="hand2",
            padx=0,
            pady=0,
        )
        self.hero_about_link.pack(side=tk.LEFT, padx=(0, 12))
        self.hero_about_link.bind("<Button-1>", lambda _e: self.open_about_dialog())
        self.hero_about_link.bind("<Enter>", lambda _e, w=self.hero_about_link: self._on_header_link_enter(w))
        self.hero_about_link.bind("<Leave>", lambda _e, w=self.hero_about_link: self._on_header_link_leave(w))
        self.hero_results_link = tk.Label(
            self.hero_right_top,
            text="Results",
            bg=DARK_PANEL_2,
            fg=DARK_MUTED,
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=0,
            pady=0,
        )
        self.hero_results_link.pack(side=tk.LEFT, padx=(0, 10))
        self.hero_results_link.bind("<Button-1>", lambda _e: self._focus_results_layout())
        self.hero_results_link.bind("<Enter>", lambda _e, w=self.hero_results_link: self._on_header_link_enter(w))
        self.hero_results_link.bind("<Leave>", lambda _e, w=self.hero_results_link: self._on_header_link_leave(w))
        self.hero_live_link = tk.Label(
            self.hero_right_top,
            text="Live",
            bg=DARK_PANEL_2,
            fg=DARK_MUTED,
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=0,
            pady=0,
        )
        self.hero_live_link.pack(side=tk.LEFT, padx=(0, 10))
        self.hero_live_link.bind("<Button-1>", lambda _e: self._focus_live_layout())
        self.hero_live_link.bind("<Enter>", lambda _e, w=self.hero_live_link: self._on_header_link_enter(w))
        self.hero_live_link.bind("<Leave>", lambda _e, w=self.hero_live_link: self._on_header_link_leave(w))
        self.hero_config_link = tk.Label(
            self.hero_right_top,
            text="Config",
            bg=DARK_PANEL_2,
            fg=DARK_MUTED,
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=0,
            pady=0,
        )
        self.hero_config_link.pack(side=tk.LEFT, padx=(0, 10))
        self.hero_config_link.bind("<Button-1>", lambda _e: self._toggle_config_mode())
        self.hero_config_link.bind("<Enter>", lambda _e, w=self.hero_config_link: self._on_header_link_enter(w))
        self.hero_config_link.bind("<Leave>", lambda _e, w=self.hero_config_link: self._on_header_link_leave(w))
        self.hero_reset_link = tk.Label(
            self.hero_right_top,
            text="Reset",
            bg=DARK_PANEL_2,
            fg=DARK_MUTED,
            font=("Segoe UI", 10),
            cursor="hand2",
            padx=0,
            pady=0,
        )
        self.hero_reset_link.pack(side=tk.LEFT, padx=(0, 12))
        self.hero_reset_link.bind("<Button-1>", lambda _e: self._reset_layout())
        self.hero_reset_link.bind("<Enter>", lambda _e, w=self.hero_reset_link: self._on_header_link_enter(w))
        self.hero_reset_link.bind("<Leave>", lambda _e, w=self.hero_reset_link: self._on_header_link_leave(w))
        self.hero_version_label = tk.Label(
            self.hero_right_top, text=APP_VERSION, bg=DARK_PANEL_2, fg=DARK_ACCENT, font=("Segoe UI Semibold", 11)
        )
        self.hero_version_label.pack(side=tk.LEFT)
        self.hero_tagline_label = tk.Label(self.hero_right, text=APP_TAGLINE, bg=DARK_PANEL_2, fg=DARK_MUTED, font=("Segoe UI", 10))
        self.hero_tagline_label.pack(anchor="e")

        self.main_split = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.main_split.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.top_frame = ttk.Frame(self.main_split, padding=8)

        conn = ttk.LabelFrame(self.top_frame, text="Connection (up to 10 COM ports)", padding=8)
        conn.pack(fill=tk.X, pady=4)
        conn.grid_columnconfigure(0, weight=1)

        conn_inputs = ttk.Frame(conn)
        conn_inputs.grid(row=0, column=0, sticky="ew")
        conn_inputs.grid_columnconfigure(1, weight=1)
        conn_inputs.grid_columnconfigure(3, weight=1)
        conn_inputs.grid_columnconfigure(4, weight=0)
        ttk.Label(conn_inputs, text="Selected COM Port").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=(0, 6))
        self.com_var = tk.StringVar(value="COM5")
        self.com_combo = ttk.Combobox(conn_inputs, textvariable=self.com_var, width=14, state="readonly")
        self.com_combo.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        ttk.Label(conn_inputs, text="Baud").grid(row=0, column=2, sticky="e", padx=(14, 6), pady=(0, 6))
        self.baud_combo = ttk.Combobox(conn_inputs, textvariable=self.baudrate_var, width=10, state="readonly", values=BAUD_OPTIONS)
        self.baud_combo.grid(row=0, column=3, sticky="ew", pady=(0, 6))
        ttk.Button(conn_inputs, text="Refresh Ports", command=self.refresh_ports).grid(row=0, column=4, sticky="ew", padx=(14, 0), pady=(0, 6))

        conn_actions = ttk.Frame(conn)
        conn_actions.grid(row=1, column=0, sticky="ew")
        for col in range(6):
            conn_actions.grid_columnconfigure(col, weight=1)
        ttk.Button(conn_actions, text="Connect Selected", command=self.connect_selected_port).grid(
            row=0, column=0, padx=2, sticky="ew"
        )
        ttk.Button(conn_actions, text="Reconnect @ Baud", command=self.reconnect_selected_port).grid(
            row=0, column=1, padx=2, sticky="ew"
        )
        ttk.Button(conn_actions, text="Disconnect Selected", command=self.disconnect_selected_port).grid(
            row=0, column=2, padx=2, sticky="ew"
        )
        ttk.Button(conn_actions, text="Connect All", command=self.connect_all_ports).grid(
            row=0, column=3, padx=2, sticky="ew"
        )
        ttk.Button(conn_actions, text="Disconnect All", command=self.disconnect_all_ports).grid(
            row=0, column=4, padx=2, sticky="ew"
        )
        self.port_station_toggle_btn = ttk.Button(
            conn_actions, text="Hide Port Station", command=self._toggle_port_station_visibility
        )
        self.port_station_toggle_btn.grid(row=0, column=5, padx=2, sticky="ew")

        conn_status_row = ttk.Frame(conn)
        conn_status_row.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        conn_status_row.grid_columnconfigure(1, weight=1)
        self.conn_status = tk.StringVar(value="Connected ports: 0")
        self.conn_status_label = ttk.Label(conn_status_row, textvariable=self.conn_status, style="Muted.TLabel")
        self.conn_status_label.grid(row=0, column=0, sticky="w")
        self.connected_ports_var = tk.StringVar(value="None")
        ttk.Label(conn_status_row, text="Connected list:", style="Muted.TLabel").grid(row=0, column=1, sticky="e", padx=(12, 6))
        self.connected_ports_label = ttk.Label(conn_status_row, textvariable=self.connected_ports_var, style="OK.TLabel")
        self.connected_ports_label.grid(row=0, column=2, sticky="w")

        self.port_station_frame = ttk.LabelFrame(self.top_frame, text="Port Station View (10 slots)", padding=4)
        self.port_station_frame.pack(fill=tk.X, pady=2)
        self._build_port_grid(self.port_station_frame)

        self.setup_frame = ttk.LabelFrame(self.top_frame, text="Test Setup", padding=8)
        self.setup_frame.pack(fill=tk.X, pady=4)

        self.operator_var = tk.StringVar(value=DEFAULT_OPERATOR)
        self.notes_var = tk.StringVar()
        self.bath_temp_c_var = tk.StringVar()
        self.salinity_psu_var = tk.StringVar(value=DEFAULT_SALINITY_PSU)
        self.bath_id_var = tk.StringVar()
        self.results_root_var = tk.StringVar(value=SENSOR_TEST_DIR)
        self.sample_count_var = tk.IntVar(value=50)
        self.batch_run_count_var = tk.IntVar(value=1)
        self.batch_delay_s_var = tk.DoubleVar(value=5.0)

        for col in (1, 3, 5):
            self.setup_frame.grid_columnconfigure(col, weight=1)
        ttk.Label(self.setup_frame, text="Operator").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=(0, 6))
        ttk.Entry(self.setup_frame, textvariable=self.operator_var, width=20, state="readonly").grid(row=0, column=1, padx=(0, 12), pady=(0, 6), sticky="ew")
        ttk.Label(self.setup_frame, text="Bath ID").grid(row=0, column=2, sticky="e", padx=(0, 6), pady=(0, 6))
        ttk.Entry(self.setup_frame, textvariable=self.bath_id_var, width=18).grid(row=0, column=3, padx=(0, 12), pady=(0, 6), sticky="ew")
        ttk.Label(self.setup_frame, text="Notes").grid(row=0, column=4, sticky="e", padx=(0, 6), pady=(0, 6))
        ttk.Entry(self.setup_frame, textvariable=self.notes_var, width=28).grid(row=0, column=5, pady=(0, 6), sticky="ew")

        ttk.Label(self.setup_frame, text="Bath Temp (C)").grid(row=1, column=0, sticky="e", padx=(0, 6), pady=(0, 6))
        ttk.Entry(self.setup_frame, textvariable=self.bath_temp_c_var, width=20).grid(row=1, column=1, padx=(0, 12), pady=(0, 6), sticky="ew")
        ttk.Label(self.setup_frame, text="Salinity (PSU)").grid(row=1, column=2, sticky="e", padx=(0, 6), pady=(0, 6))
        ttk.Entry(self.setup_frame, textvariable=self.salinity_psu_var, width=20, state="readonly").grid(row=1, column=3, padx=(0, 12), pady=(0, 6), sticky="ew")
        ttk.Label(self.setup_frame, text="Samples").grid(row=1, column=4, sticky="e", padx=(0, 6), pady=(0, 6))
        ttk.Spinbox(self.setup_frame, from_=20, to=500, textvariable=self.sample_count_var, width=8).grid(row=1, column=5, pady=(0, 6), sticky="ew")
        ttk.Label(self.setup_frame, text="Results Root").grid(row=2, column=0, sticky="e", padx=(0, 6), pady=(2, 0))
        self.results_root_entry = ttk.Entry(self.setup_frame, textvariable=self.results_root_var, width=78, state="readonly")
        self.results_root_entry.grid(
            row=2, column=1, columnspan=4, padx=(0, 12), sticky="ew", pady=(2, 0)
        )
        self.results_root_browse_btn = ttk.Button(self.setup_frame, text="Browse", command=self.browse_results_root)
        self.results_root_browse_btn.grid(
            row=2, column=5, sticky="ew", pady=(2, 0)
        )
        for var in (
            self.operator_var,
            self.notes_var,
            self.bath_id_var,
            self.bath_temp_c_var,
            self.salinity_psu_var,
            self.results_root_var,
            self.sample_count_var,
        ):
            var.trace_add("write", self._on_setup_field_changed)
        self._apply_port_station_visibility(persist=False)

        self.actions_frame = ttk.LabelFrame(self.top_frame, text="Actions", padding=8)
        self.actions_frame.pack(fill=tk.X, pady=4)
        self.actions_frame.grid_columnconfigure(0, weight=1)
        action_buttons = ttk.Frame(self.actions_frame)
        action_buttons.grid(row=0, column=0, sticky="ew")
        for col in range(10):
            action_buttons.grid_columnconfigure(col, weight=1)

        self.run_btn = ttk.Button(
            action_buttons, text="Run Test", command=self.run_unit_test, state=tk.DISABLED, style="Primary.TButton"
        )
        self.run_btn.grid(row=0, column=0, padx=2, sticky="ew")
        ttk.Button(action_buttons, text="Save JSON", command=self.save_session_json).grid(
            row=0, column=1, padx=2, sticky="ew"
        )
        ttk.Button(action_buttons, text="Reset", command=self.reset_session).grid(
            row=0, column=2, padx=2, sticky="ew"
        )
        ttk.Button(action_buttons, text="Live Plot", command=self.show_plot_tab).grid(
            row=0, column=3, padx=2, sticky="ew"
        )
        ttk.Button(action_buttons, text="Console", command=self.show_console_tab).grid(
            row=0, column=4, padx=2, sticky="ew"
        )
        self.detach_console_btn = ttk.Button(action_buttons, text="Detach", command=self.detach_console_window)
        self.detach_console_btn.grid(row=0, column=5, padx=2, sticky="ew")
        ttk.Button(action_buttons, text="Session Plot", command=self.plot_current_session).grid(
            row=0, column=6, padx=2, sticky="ew"
        )
        ttk.Button(action_buttons, text="Load Session", command=self.load_session_plot).grid(
            row=0, column=7, padx=2, sticky="ew"
        )
        ttk.Button(action_buttons, text="Reload JSON", command=self.reload_current_session_plot).grid(
            row=0, column=8, padx=2, sticky="ew"
        )
        ttk.Button(action_buttons, text="CSV Column", command=self.toggle_csv_column).grid(
            row=0, column=9, padx=2, sticky="ew"
        )

        action_status = ttk.Frame(self.actions_frame)
        action_status.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        action_status.grid_columnconfigure(10, weight=1)
        ttk.Label(action_status, text="Runs").grid(row=0, column=0, sticky="e")
        ttk.Spinbox(action_status, from_=1, to=50, textvariable=self.batch_run_count_var, width=6).grid(row=0, column=1, padx=(4, 16), sticky="w")
        ttk.Label(action_status, text="Delay (s)").grid(row=0, column=2, sticky="e")
        ttk.Spinbox(action_status, from_=0, to=300, increment=1, textvariable=self.batch_delay_s_var, width=6).grid(row=0, column=3, padx=(4, 16), sticky="w")
        ttk.Checkbutton(action_status, text="Dark Mode", variable=self.dark_mode_var, command=self._on_toggle_dark_mode).grid(row=0, column=4, padx=(0, 16), sticky="w")
        ttk.Label(action_status, text="Mode").grid(row=0, column=5, sticky="e")
        ttk.Radiobutton(
            action_status,
            text="Production",
            variable=self.mode_choice_var,
            value="production",
            command=self._on_mode_choice_changed,
        ).grid(row=0, column=6, padx=(4, 8), sticky="w")
        ttk.Radiobutton(
            action_status,
            text="Debug",
            variable=self.mode_choice_var,
            value="debug",
            command=self._on_mode_choice_changed,
        ).grid(row=0, column=7, padx=(0, 16), sticky="w")

        self.limit_var = tk.StringVar(value=f"Units tested: 0 / {MAX_UNITS_PER_SESSION}")
        ttk.Label(action_status, textvariable=self.limit_var, style="Accent.TLabel").grid(row=0, column=8, sticky="w")
        ttk.Label(action_status, textvariable=self.runs_left_var, style="Muted.TLabel").grid(row=0, column=9, padx=(16, 0), sticky="w")
        self.setup_toggle_btn = ttk.Button(action_status, text="Hide Test Setup", command=self._toggle_test_setup_visibility)
        self.setup_toggle_btn.grid(row=0, column=10, padx=(16, 0), sticky="e")
        self.mode_status_entry = ttk.Entry(action_status, textvariable=self.mode_status_var, state="readonly")
        self.mode_status_entry.grid(row=1, column=0, columnspan=11, pady=(8, 0), sticky="ew")

        self.main_notebook = ttk.Notebook(self.main_split)
        self.main_split.add(self.top_frame, weight=0)
        self.main_split.add(self.main_notebook, weight=1)

        self.test_tab = ttk.Frame(self.main_notebook, padding=8)
        self.plot_tab = ttk.Frame(self.main_notebook, padding=8)
        self.sample_setup_tab = ttk.Frame(self.main_notebook, padding=8)
        self.log_tab = ttk.Frame(self.main_notebook, padding=8)
        self.debug_tab = tk.Frame(self.main_notebook, padx=8, pady=8)

        self.main_notebook.add(self.test_tab, text="Test Results")
        self.main_notebook.add(self.plot_tab, text="Live Plot")
        self.main_notebook.add(self.debug_tab, text="Serial Consoles")
        self.main_notebook.add(self.log_tab, text="Run Log")
        self.main_notebook.add(self.sample_setup_tab, text="Sample Setup")

        mid = ttk.Frame(self.test_tab)
        mid.pack(fill=tk.BOTH, expand=True)

        cols = (
            "timestamp",
            "port",
            "serial",
            "red_ns",
            "red_v_std",
            "red_v_avg",
            "blue_ns",
            "blue_v_std",
            "blue_v_avg",
            "red_pll_v_std",
            "red_pll_v_avg",
            "blue_pll_v_std",
            "blue_pll_v_avg",
            "raw_temp_v_std",
            "raw_temp_v_avg",
            "elec_temp_v_std",
            "elec_temp_v_avg",
            "flags",
            "sample_csv",
        )
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=12)
        tree_scroll_y = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        for c, w in [
            ("timestamp", 145),
            ("port", 75),
            ("serial", 75),
            ("red_ns", 75),
            ("red_v_std", 90),
            ("red_v_avg", 90),
            ("blue_ns", 75),
            ("blue_v_std", 90),
            ("blue_v_avg", 90),
            ("red_pll_v_std", 105),
            ("red_pll_v_avg", 105),
            ("blue_pll_v_std", 105),
            ("blue_pll_v_avg", 105),
            ("raw_temp_v_std", 115),
            ("raw_temp_v_avg", 115),
            ("elec_temp_v_std", 118),
            ("elec_temp_v_avg", 118),
            ("flags", 220),
            ("sample_csv", 390),
        ]:
            self.tree.heading(c, text=c, command=lambda col=c: self.sort_tree_by_column(col))
            self.tree.column(c, width=w, minwidth=60, anchor="w", stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree.configure(displaycolumns=tuple(c for c in cols if c != "sample_csv"))
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        setup_panel = ttk.LabelFrame(self.sample_setup_tab, text="Generic Sample Format", padding=8)
        setup_panel.pack(fill=tk.BOTH, expand=True)
        self._build_sample_format_controls(setup_panel)

        live = ttk.LabelFrame(self.plot_tab, text="Current Run Live View", padding=8)
        live.pack(fill=tk.BOTH, expand=True)

        live_controls = ttk.Frame(live)
        live_controls.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(live_controls, text="Plot field").pack(side=tk.LEFT)
        self.live_field_var = tk.StringVar(value="Red Phase")
        self.live_field_combo = ttk.Combobox(
            live_controls,
            textvariable=self.live_field_var,
            state="readonly",
            width=18,
            values=list(self.live_plot_fields.keys()),
        )
        self.live_field_combo.pack(side=tk.LEFT, padx=(6, 16))
        self.live_field_combo.bind("<<ComboboxSelected>>", self._on_live_field_changed)
        ttk.Button(live_controls, text="Refresh Plot", command=self.refresh_live_plot).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(live_controls, text="Auto Y", variable=self.live_autoscale_var, command=self.refresh_live_plot).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Label(live_controls, text="Ymin").pack(side=tk.LEFT)
        ttk.Entry(live_controls, textvariable=self.live_ymin_var, width=8).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(live_controls, text="Ymax").pack(side=tk.LEFT)
        ttk.Entry(live_controls, textvariable=self.live_ymax_var, width=8).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Checkbutton(live_controls, text="Points", variable=self.live_show_points_var, command=self.refresh_live_plot).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Checkbutton(
            live_controls, text="Filter Ports", variable=self.live_visible_only_var, command=self.refresh_live_plot
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(live_controls, text="Visible Ports", command=self.select_visible_ports).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(live_controls, text="X Start").pack(side=tk.LEFT)
        ttk.Entry(live_controls, textvariable=self.live_x_start_var, width=6).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(live_controls, text="X End").pack(side=tk.LEFT)
        ttk.Entry(live_controls, textvariable=self.live_x_end_var, width=6).pack(side=tk.LEFT, padx=(4, 12))

        self.live_std_var = tk.StringVar(value="Std Dev: n/a")
        ttk.Label(live_controls, textvariable=self.live_std_var).pack(side=tk.LEFT, padx=(0, 16))
        self.live_samples_var = tk.StringVar(value="Samples: 0 / 0")
        ttk.Label(live_controls, textvariable=self.live_samples_var).pack(side=tk.LEFT)

        live_grid = ttk.Frame(live)
        live_grid.pack(fill=tk.BOTH, expand=True)
        live_grid.columnconfigure(0, weight=3)
        live_grid.columnconfigure(1, weight=2)
        live_grid.rowconfigure(0, weight=1)

        self.live_canvas = tk.Canvas(live_grid, bg="#0b1220", height=230, highlightthickness=1, highlightbackground=DARK_BORDER)
        self.live_canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.live_canvas.bind("<Configure>", lambda _e: self.refresh_live_plot())

        self.live_text = scrolledtext.ScrolledText(
            live_grid,
            height=12,
            state=tk.DISABLED,
            font=self._console_font(),
            bg=DARK_PANEL,
            fg=DARK_TEXT,
            insertbackground=DARK_TEXT,
            relief=tk.FLAT,
        )
        self.live_text.grid(row=0, column=1, sticky="nsew")

        bottom = ttk.LabelFrame(self.log_tab, text="Run Log", padding=8)
        bottom.pack(fill=tk.BOTH, expand=True)
        self.log_box = scrolledtext.ScrolledText(
            bottom,
            height=14,
            state=tk.DISABLED,
            font=self._console_font(),
            bg=DARK_PANEL,
            fg=DARK_TEXT,
            insertbackground=DARK_TEXT,
            relief=tk.FLAT,
        )
        self.log_box.pack(fill=tk.BOTH, expand=True)

        debug = ttk.LabelFrame(self.debug_tab, text="Serial Debug Consoles (Per COM Port)", padding=8)
        debug.pack(fill=tk.BOTH, expand=True)
        debug_actions = ttk.Frame(debug)
        debug_actions.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(debug_actions, text="Clear Selected Debug Tab", command=self.clear_selected_debug_tab).pack(side=tk.LEFT)
        ttk.Button(debug_actions, text="Export Console CSV", command=self.export_manual_capture).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Checkbutton(debug_actions, text="CR", variable=self.console_send_cr_var).pack(side=tk.LEFT, padx=(14, 0))
        ttk.Checkbutton(debug_actions, text="LF", variable=self.console_send_lf_var).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(debug_actions, text="Display").pack(side=tk.LEFT)
        ttk.Radiobutton(debug_actions, text="ASCII", variable=self.console_display_mode_var, value="ascii").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Radiobutton(debug_actions, text="HEX", variable=self.console_display_mode_var, value="hex").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Radiobutton(debug_actions, text="DEC", variable=self.console_display_mode_var, value="dec").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Radiobutton(debug_actions, text="BIN", variable=self.console_display_mode_var, value="bin").pack(side=tk.LEFT, padx=(4, 0))
        self.debug_notebook = ttk.Notebook(debug)
        self.debug_notebook.pack(fill=tk.BOTH, expand=True)
        self._rebuild_measureand_editor_rows(self.sample_field_defs)
        self._apply_measureand_config(show_message=False)
        self.root.after(40, self._apply_startup_layout)
        self.root.bind("<Configure>", self._on_root_configure, add="+")
        self.root.bind("<ButtonRelease-1>", lambda _e: self._schedule_layout_save(delay_ms=250), add="+")
        self.main_split.bind("<ButtonRelease-1>", lambda _e: self._schedule_layout_save(delay_ms=250), add="+")
        self.main_notebook.bind("<<NotebookTabChanged>>", lambda _e: self._schedule_layout_save(delay_ms=250), add="+")

    def show_plot_tab(self):
        self.main_notebook.select(self.plot_tab)

    def show_console_tab(self):
        if self.console_detached:
            try:
                self.root.tk.call("wm", "deiconify", str(self.debug_tab))
                self.root.tk.call("raise", str(self.debug_tab))
            except Exception:
                pass
            return
        self.main_notebook.select(self.debug_tab)

    def toggle_sample_format_panel(self, expanded=None):
        self.sample_format_expanded = bool(True if expanded is None else expanded)
        if self.sample_format_expanded and hasattr(self, "sample_setup_tab"):
            try:
                self.main_notebook.select(self.sample_setup_tab)
            except Exception:
                pass
        elif hasattr(self, "plot_tab"):
            try:
                self.main_notebook.select(self.plot_tab)
            except Exception:
                pass
        self._schedule_layout_save(delay_ms=300)

    def _markdown_to_basic_html(self, markdown_text):
        lines = markdown_text.splitlines()
        out = []
        in_code = False
        in_list = False
        link_re = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")

        def convert_inline(text):
            escaped = html.escape(text)
            return link_re.sub(lambda m: f'<a href="{html.escape(m.group(2), quote=True)}">{html.escape(m.group(1))}</a>', escaped)

        for raw in lines:
            line = raw.rstrip("\n")
            if line.strip().startswith("```"):
                if in_code:
                    out.append("</code></pre>")
                    in_code = False
                else:
                    if in_list:
                        out.append("</ul>")
                        in_list = False
                    out.append("<pre><code>")
                    in_code = True
                continue

            if in_code:
                out.append(html.escape(line))
                continue

            if not line.strip():
                if in_list:
                    out.append("</ul>")
                    in_list = False
                continue

            heading = re.match(r"^(#{1,6})\s+(.*)$", line)
            if heading:
                if in_list:
                    out.append("</ul>")
                    in_list = False
                level = len(heading.group(1))
                out.append(f"<h{level}>{convert_inline(heading.group(2).strip())}</h{level}>")
                continue

            bullet = re.match(r"^\s*[-*]\s+(.*)$", line)
            if bullet:
                if not in_list:
                    out.append("<ul>")
                    in_list = True
                out.append(f"<li>{convert_inline(bullet.group(1).strip())}</li>")
                continue

            if in_list:
                out.append("</ul>")
                in_list = False
            out.append(f"<p>{convert_inline(line.strip())}</p>")

        if in_list:
            out.append("</ul>")
        if in_code:
            out.append("</code></pre>")
        return "\n".join(out)

    @staticmethod
    def _render_markdown_html(markdown_text):
        try:
            md = mistune.create_markdown(plugins=["table"])
            return md(markdown_text)
        except Exception:
            return None

    @staticmethod
    def _rewrite_help_links(html_text, page_uri, base_uri):
        """Normalize links in rendered help HTML for temp-file viewing."""
        if not html_text:
            return html_text
        abs_scheme = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*:")
        href_re = re.compile(r'href=(["\'])(.*?)\1', re.IGNORECASE)

        def repl(match):
            quote = match.group(1)
            href = (match.group(2) or "").strip()
            if not href:
                return match.group(0)
            if href.startswith("#"):
                new_href = f"{page_uri}{href}"
            elif href.startswith("//") or abs_scheme.match(href):
                new_href = href
            else:
                new_href = urljoin(base_uri, href)
            return f'href={quote}{html.escape(new_href, quote=True)}{quote}'

        return href_re.sub(repl, html_text)

    def _open_contact_email(self):
        try:
            webbrowser.open(f"mailto:{APP_CONTACT_EMAIL}")
        except Exception:
            messagebox.showerror("Contact Error", f"Could not open email client for:\n{APP_CONTACT_EMAIL}")

    def open_about_dialog(self):
        if self.about_window is not None:
            try:
                if self.about_window.winfo_exists():
                    self.about_window.deiconify()
                    self.about_window.lift()
                    self.about_window.focus_force()
                    return
            except Exception:
                self.about_window = None

        colors = self._theme_colors()
        bg = colors["panel_2"]
        fg = colors["fg"]
        muted = colors["muted"]
        accent = colors["accent"]
        border = colors["border"]

        win = tk.Toplevel(self.root)
        self.about_window = win
        win.title(f"About {APP_NAME}")
        win.resizable(False, False)
        win.transient(self.root)
        win.configure(bg=bg, padx=16, pady=14)

        # Standard About order: product, version, publisher, author, contact, summary.
        tk.Label(win, text=APP_NAME, bg=bg, fg=fg, font=("Segoe UI Semibold", 14)).grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        tk.Label(win, text=f"Version {APP_VERSION}", bg=bg, fg=accent, font=("Segoe UI Semibold", 11)).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(2, 10)
        )

        sep = tk.Frame(win, height=1, bg=border)
        sep.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        tk.Label(win, text="Company", bg=bg, fg=muted, font=("Segoe UI Semibold", 10)).grid(row=3, column=0, sticky="e", padx=(0, 10))
        tk.Label(win, text=APP_COMPANY, bg=bg, fg=fg, font=("Segoe UI", 10)).grid(row=3, column=1, sticky="w")
        tk.Label(win, text="Author", bg=bg, fg=muted, font=("Segoe UI Semibold", 10)).grid(row=4, column=0, sticky="e", padx=(0, 10), pady=(4, 0))
        tk.Label(win, text=APP_AUTHOR, bg=bg, fg=fg, font=("Segoe UI", 10)).grid(row=4, column=1, sticky="w", pady=(4, 0))
        tk.Label(win, text="Contact", bg=bg, fg=muted, font=("Segoe UI Semibold", 10)).grid(row=5, column=0, sticky="e", padx=(0, 10), pady=(4, 0))
        email_link = tk.Label(
            win,
            text=APP_CONTACT_EMAIL,
            bg=bg,
            fg=accent,
            font=("Segoe UI", 10, "underline"),
            cursor="hand2",
        )
        email_link.grid(row=5, column=1, sticky="w", pady=(4, 0))
        email_link.bind("<Button-1>", lambda _e: self._open_contact_email())

        tk.Label(win, text="Overview", bg=bg, fg=muted, font=("Segoe UI Semibold", 10)).grid(
            row=6, column=0, sticky="ne", padx=(0, 10), pady=(10, 0)
        )
        tk.Label(
            win,
            text=APP_ABOUT_SUMMARY,
            bg=bg,
            fg=fg,
            font=("Segoe UI", 10),
            justify="left",
            wraplength=520,
        ).grid(row=6, column=1, sticky="w", pady=(10, 0))

        close_btn = ttk.Button(win, text="Close", width=10, command=lambda: on_close())
        close_btn.grid(row=7, column=1, sticky="e", pady=(14, 0))

        win.grid_columnconfigure(1, weight=1)
        win.update_idletasks()
        px = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - win.winfo_width()) // 2)
        py = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - win.winfo_height()) // 3)
        win.geometry(f"+{px}+{py}")

        def on_close():
            self.about_window = None
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)
        win.grab_set()
        win.focus_force()

    def open_readme_help(self):
        candidates = []
        if getattr(sys, "frozen", False):
            meipass = getattr(sys, "_MEIPASS", "")
            if meipass:
                candidates.append(Path(meipass) / "README.md")
            exe_dir = Path(sys.executable).resolve().parent
            candidates.append(exe_dir / "README.md")
            candidates.append(exe_dir.parent / "README.md")
        app_file = Path(__file__).resolve()
        candidates.append(app_file.parents[2] / "README.md")
        candidates.append(Path.cwd() / "README.md")

        readme_path = None
        for c in candidates:
            if c.exists():
                readme_path = c
                break
        if readme_path is None:
            attempted = "\n".join(str(c) for c in candidates)
            messagebox.showerror("Help Not Found", f"README.md not found. Checked:\n{attempted}")
            return
        try:
            markdown_text = readme_path.read_text(encoding="utf-8", errors="replace")
            body_html = self._render_markdown_html(markdown_text)
            if not body_html:
                body_html = self._markdown_to_basic_html(markdown_text)
            base_href = readme_path.parent.resolve().as_uri()
            if not base_href.endswith("/"):
                base_href += "/"
            with tempfile.NamedTemporaryFile("w", delete=False, suffix="_engineers_field_kit_help.html", encoding="utf-8") as f:
                temp_path = Path(f.name)
            page_uri = temp_path.resolve().as_uri()
            body_html = self._rewrite_help_links(body_html, page_uri, base_href)
            page = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_NAME} Help</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; max-width: 980px; margin: 24px auto; padding: 0 16px; line-height: 1.55; background:#f8fafc; color:#0f172a; }}
    h1,h2,h3,h4,h5,h6 {{ color:#0f172a; margin-top: 1.2em; }}
    pre {{ background:#0f172a; color:#e2e8f0; padding: 12px; overflow:auto; border-radius: 8px; }}
    code {{ font-family: Consolas, monospace; }}
    a {{ color:#0f62fe; text-decoration:none; }}
    a:hover {{ text-decoration:underline; }}
    ul {{ padding-left: 1.2rem; }}
    ol {{ padding-left: 1.2rem; }}
    p {{ margin: 0.55em 0; }}
    table {{ border-collapse: collapse; width: 100%; margin: 12px 0; }}
    th, td {{ border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; }}
    th {{ background: #e2e8f0; }}
    blockquote {{ margin: 12px 0; padding: 8px 12px; border-left: 3px solid #94a3b8; background: #f1f5f9; }}
    img {{ max-width: 100%; height: auto; border: 1px solid #cbd5e1; border-radius: 8px; }}
    .hdr {{ margin-bottom: 12px; padding-bottom: 8px; border-bottom: 1px solid #cbd5e1; }}
  </style>
</head>
<body>
  <div class="hdr"><h1>{APP_NAME} Help</h1><div>Source: {readme_path}</div></div>
  {body_html}
</body>
</html>"""
            with open(temp_path, "w", encoding="utf-8", newline="\n") as f:
                f.write(page)
            webbrowser.open(temp_path.as_uri())
        except Exception as exc:
            messagebox.showerror("Help Error", f"Could not open help page:\n{exc}")

    def _install_output_root(self):
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))

    def _debug_results_root(self):
        return os.path.join(self._install_output_root(), DEBUG_RESULTS_SUBDIR)

    def _update_mode_status_text(self):
        mode_text = "DEBUG" if self.debug_mode_var.get() else "PRODUCTION"
        root = self._current_results_root()
        hint = "Install-local safe output" if self.debug_mode_var.get() else "Production output root"
        self.mode_status_var.set(f"Mode: {mode_text} | Save Root: {root} | {hint}")

    def _on_mode_choice_changed(self):
        choice = str(self.mode_choice_var.get()).strip().lower()
        enabled = choice == "debug"
        if enabled == bool(self.debug_mode_var.get()):
            self._update_mode_status_text()
            return
        self._set_debug_mode(enabled, persist=True, announce=True)

    def _set_debug_mode(self, enabled, persist=True, announce=True):
        enabled = bool(enabled)
        self.debug_mode_var.set(enabled)
        if hasattr(self, "mode_choice_var"):
            self.mode_choice_var.set("debug" if enabled else "production")
        if enabled:
            debug_root = self._debug_results_root()
            current_root = self._current_results_root()
            if current_root and os.path.abspath(current_root) != os.path.abspath(debug_root):
                self.non_debug_results_root = current_root
            self._apply_results_root(debug_root, log_change=False)
            if announce:
                self.log(f"Debug mode enabled. Results root set to install directory: {debug_root}")
                self.log(f"Session summary file: {self.session_csv}")
        else:
            restore_root = str(self.non_debug_results_root).strip() or SENSOR_TEST_DIR
            self._apply_results_root(restore_root, log_change=False)
            if announce:
                self.log(f"Debug mode disabled. Results root restored: {restore_root}")
                self.log(f"Session summary file: {self.session_csv}")
        self._update_results_root_controls()
        self._update_mode_status_text()
        self._on_setup_field_changed()
        if persist:
            self._save_app_config()

    def _update_results_root_controls(self):
        if hasattr(self, "results_root_entry"):
            entry_state = tk.DISABLED if self.debug_mode_var.get() else "readonly"
            self.results_root_entry.configure(state=entry_state)
        if hasattr(self, "results_root_browse_btn"):
            state = tk.DISABLED if self.debug_mode_var.get() else tk.NORMAL
            self.results_root_browse_btn.configure(state=state)

    def _current_results_root(self):
        if hasattr(self, "results_root_var"):
            value = self.results_root_var.get().strip()
            if value:
                return value
        return SENSOR_TEST_DIR

    def _apply_results_root(self, root_path, log_change=True):
        root = os.path.abspath(str(root_path).strip())
        os.makedirs(root, exist_ok=True)
        if not self.debug_mode_var.get():
            self.non_debug_results_root = root
        if hasattr(self, "results_root_var"):
            self.results_root_var.set(root)
        self.session_dir = os.path.join(root, "sessions", PRECAL_TEST_SUBDIR)
        os.makedirs(self.session_dir, exist_ok=True)
        self.profile_dir = os.path.join(self.session_dir, "profiles")
        os.makedirs(self.profile_dir, exist_ok=True)
        self.session_csv = os.path.join(self.session_dir, f"sbe83_session_{self.session_id}.csv")
        self._update_mode_status_text()
        if log_change:
            self.log(f"Results root updated: {root}")
            self.log(f"Session summary file: {self.session_csv}")

    def browse_results_root(self):
        if self.debug_mode_var.get():
            messagebox.showinfo("Debug Mode Active", "Disable Debug Mode to change Results Root.")
            return
        initial = self._current_results_root()
        start_dir = initial if os.path.isdir(initial) else os.getcwd()
        path = filedialog.askdirectory(title="Select Results Root Folder", initialdir=start_dir)
        if not path:
            return
        try:
            self._apply_results_root(path, log_change=True)
            self._save_app_config()
        except Exception as exc:
            messagebox.showerror("Results Root Error", str(exc))

    def detach_console_window(self):
        try:
            if not self.console_detached:
                if str(self.debug_tab) in self.main_notebook.tabs():
                    self.main_notebook.forget(self.debug_tab)
                self.root.update_idletasks()
                self.root.tk.call("wm", "manage", str(self.debug_tab))
                self.root.tk.call("wm", "title", str(self.debug_tab), "Serial Consoles")
                self.root.tk.call("wm", "geometry", str(self.debug_tab), "980x640+120+120")
                self.root.tk.call("wm", "protocol", str(self.debug_tab), "WM_DELETE_WINDOW", self._dock_console_callback)
                self.detach_console_btn.configure(text="Dock Console")
                self.console_detached = True
            else:
                self._dock_console_tab()
        except Exception as exc:
            # Keep the console reachable if detach/dock fails on this platform.
            self._dock_console_tab()
            messagebox.showerror("Detach Not Available", f"Could not detach/dock console on this system:\n{exc}")

    def _dock_console_tab(self):
        try:
            self.root.tk.call("wm", "protocol", str(self.debug_tab), "WM_DELETE_WINDOW", "")
        except Exception:
            pass
        try:
            self.root.tk.call("wm", "forget", str(self.debug_tab))
        except Exception:
            pass
        self.root.update_idletasks()
        if str(self.debug_tab) not in self.main_notebook.tabs():
            self.main_notebook.add(self.debug_tab, text="Serial Consoles")
        self.main_notebook.select(self.debug_tab)
        self.detach_console_btn.configure(text="Detach Console")
        self.console_detached = False

    @staticmethod
    def _sanitize_measureand_key(text, idx):
        key = re.sub(r"[^a-zA-Z0-9]+", "_", str(text).strip().lower()).strip("_")
        return key or f"field_{idx + 1}"

    def _default_sample_field_defs(self):
        defs = []
        for idx, key in enumerate(self.tsr_fields):
            unit = "V" if "voltage" in key else ("ns" if "phase" in key else "")
            scale = "milli" if unit == "V" else "raw"
            defs.append(
                {
                    "index": idx,
                    "key": key,
                    "description": self.default_field_descriptions.get(key, key.replace("_", " ").title()),
                    "unit": unit,
                    "scale": scale,
                    "min_val": "",
                    "max_val": "",
                    "stuck_n": "",
                    "expr": "",
                    "plot_live": key in self.default_live_plot_fields.values(),
                    "plot_session": key in {"red_phase", "blue_phase", "red_voltage", "blue_voltage"},
                    "live_default": key == "red_phase",
                }
            )
        return defs

    def _build_sample_format_controls(self, parent):
        self.sample_format_body = ttk.Frame(parent)
        self.sample_format_body.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(self.sample_format_body)
        top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(top, text="Example sample").pack(side=tk.LEFT)
        ttk.Entry(top, textvariable=self.example_sample_var, width=72).pack(side=tk.LEFT, padx=(6, 8), fill=tk.X, expand=True)
        ttk.Label(top, text="Delimiter").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(top, textvariable=self.delimiter_var, width=3).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top, text="Quick Setup + Plot", command=self.quick_setup_from_example).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Load From Example", command=self.load_measureands_from_example).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Apply Measureands", command=self.apply_measureands_from_editor).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Save Profile", command=self.save_parser_profile).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Load Profile", command=self.load_parser_profile).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Reset Default", command=self.reset_measureands_default).pack(side=tk.LEFT)

        parser_row = ttk.Frame(self.sample_format_body)
        parser_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(parser_row, text="Sample Cmd").pack(side=tk.LEFT)
        ttk.Entry(parser_row, textvariable=self.sample_command_var, width=10).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(parser_row, text="Trim Prefix").pack(side=tk.LEFT)
        ttk.Entry(parser_row, textvariable=self.parser_trim_prefix_var, width=12).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(parser_row, text="Start Token").pack(side=tk.LEFT)
        ttk.Spinbox(parser_row, from_=0, to=20, textvariable=self.parser_token_start_var, width=5).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(parser_row, text="Regex").pack(side=tk.LEFT)
        ttk.Entry(parser_row, textvariable=self.parser_regex_var, width=22).pack(side=tk.LEFT, padx=(6, 0), fill=tk.X, expand=True)

        ttk.Label(
            self.sample_format_body,
            text="Paste one sensor output line, then use Quick Setup + Plot or edit names/descriptions and pick live/session fields.",
            foreground=DARK_MUTED,
        ).pack(anchor="w", pady=(0, 6))

        editor_wrap = ttk.Frame(self.sample_format_body)
        editor_wrap.pack(fill=tk.BOTH, expand=True)
        self.measureand_canvas = tk.Canvas(editor_wrap, height=190, bg=DARK_BG, highlightthickness=1, highlightbackground=DARK_BORDER)
        self.measureand_xscroll = ttk.Scrollbar(editor_wrap, orient=tk.HORIZONTAL, command=self.measureand_canvas.xview)
        self.measureand_yscroll = ttk.Scrollbar(editor_wrap, orient=tk.VERTICAL, command=self.measureand_canvas.yview)
        self.measureand_canvas.configure(xscrollcommand=self.measureand_xscroll.set, yscrollcommand=self.measureand_yscroll.set)
        self.measureand_canvas.grid(row=0, column=0, sticky="nsew")
        self.measureand_yscroll.grid(row=0, column=1, sticky="ns")
        self.measureand_xscroll.grid(row=1, column=0, sticky="ew")
        editor_wrap.rowconfigure(0, weight=1)
        editor_wrap.columnconfigure(0, weight=1)

        self.measureand_editor = ttk.Frame(self.measureand_canvas)
        self.measureand_canvas_window = self.measureand_canvas.create_window((0, 0), window=self.measureand_editor, anchor="nw")
        self.measureand_editor.bind(
            "<Configure>", lambda _e: self.measureand_canvas.configure(scrollregion=self.measureand_canvas.bbox("all"))
        )
        self.measureand_canvas.bind(
            "<Configure>",
            lambda e: self.measureand_canvas.itemconfigure(self.measureand_canvas_window, width=max(e.width, 980)),
        )
        self.measureand_default_live_idx = tk.IntVar(value=0)
        self.sample_format_expanded = True

    def _rebuild_measureand_editor_rows(self, defs):
        for child in self.measureand_editor.winfo_children():
            child.destroy()
        self.measureand_rows = []
        headers = ("Idx", "Field Key", "Description", "Unit", "Scale", "Min", "Max", "StuckN", "Derived Expr", "Live", "Session", "Default")

        # Use a single shared grid for headers + rows to keep columns aligned.
        column_mins = {
            0: 34,
            1: 160,
            2: 420,
            3: 58,
            4: 62,
            5: 52,
            6: 52,
            7: 62,
            8: 220,
            9: 42,
            10: 56,
            11: 56,
        }
        for col, minsize in column_mins.items():
            weight = 1 if col in (2, 8) else 0
            self.measureand_editor.columnconfigure(col, minsize=minsize, weight=weight)

        for col, txt in enumerate(headers):
            ttk.Label(self.measureand_editor, text=txt).grid(row=0, column=col, sticky="w", padx=(0, 8), pady=(0, 4))
        if not defs:
            return

        default_idx = 0
        for i, d in enumerate(defs):
            if d.get("live_default"):
                default_idx = i
                break
        self.measureand_default_live_idx.set(default_idx)

        for i, d in enumerate(defs):
            row_idx = i + 1
            ttk.Label(self.measureand_editor, text=str(i + 1), width=4).grid(row=row_idx, column=0, sticky="w")
            key_var = tk.StringVar(value=d.get("key", f"field_{i + 1}"))
            desc_var = tk.StringVar(value=d.get("description", ""))
            unit_var = tk.StringVar(value=d.get("unit", ""))
            scale_var = tk.StringVar(value=d.get("scale", "raw"))
            min_var = tk.StringVar(value=str(d.get("min_val", "")))
            max_var = tk.StringVar(value=str(d.get("max_val", "")))
            stuck_var = tk.StringVar(value=str(d.get("stuck_n", "")))
            expr_var = tk.StringVar(value=d.get("expr", ""))
            live_var = tk.BooleanVar(value=bool(d.get("plot_live", False)))
            session_var = tk.BooleanVar(value=bool(d.get("plot_session", False)))
            ttk.Entry(self.measureand_editor, textvariable=key_var, width=24).grid(row=row_idx, column=1, sticky="ew", padx=(0, 8))
            ttk.Entry(self.measureand_editor, textvariable=desc_var).grid(row=row_idx, column=2, sticky="ew", padx=(0, 8))
            ttk.Entry(self.measureand_editor, textvariable=unit_var, width=8).grid(row=row_idx, column=3, sticky="w", padx=(0, 8))
            ttk.Combobox(
                self.measureand_editor, textvariable=scale_var, width=7, state="readonly", values=list(self.unit_scale_factors.keys())
            ).grid(
                row=row_idx, column=4, sticky="w", padx=(0, 8)
            )
            ttk.Entry(self.measureand_editor, textvariable=min_var, width=8).grid(row=row_idx, column=5, sticky="w", padx=(0, 8))
            ttk.Entry(self.measureand_editor, textvariable=max_var, width=8).grid(row=row_idx, column=6, sticky="w", padx=(0, 8))
            ttk.Entry(self.measureand_editor, textvariable=stuck_var, width=6).grid(row=row_idx, column=7, sticky="w", padx=(0, 8))
            ttk.Entry(self.measureand_editor, textvariable=expr_var, width=18).grid(row=row_idx, column=8, sticky="ew", padx=(0, 8))
            ttk.Checkbutton(self.measureand_editor, variable=live_var).grid(row=row_idx, column=9, sticky="w")
            ttk.Checkbutton(self.measureand_editor, variable=session_var).grid(row=row_idx, column=10, sticky="w")
            ttk.Radiobutton(self.measureand_editor, variable=self.measureand_default_live_idx, value=i).grid(
                row=row_idx, column=11, sticky="w"
            )
            self.measureand_rows.append(
                {
                    "index": i,
                    "key_var": key_var,
                    "description_var": desc_var,
                    "unit_var": unit_var,
                    "scale_var": scale_var,
                    "min_var": min_var,
                    "max_var": max_var,
                    "stuck_var": stuck_var,
                    "expr_var": expr_var,
                    "plot_live_var": live_var,
                    "plot_session_var": session_var,
                }
            )

    def _current_measureand_defs_from_editor(self):
        defs = []
        default_idx = self.measureand_default_live_idx.get() if self.measureand_rows else 0
        for row in self.measureand_rows:
            idx = row["index"]
            key = self._sanitize_measureand_key(row["key_var"].get(), idx)
            description = row["description_var"].get().strip() or key.replace("_", " ").title()
            defs.append(
                {
                    "index": idx,
                    "key": key,
                    "description": description,
                    "unit": row["unit_var"].get().strip(),
                    "scale": row["scale_var"].get().strip() or "raw",
                    "min_val": row["min_var"].get().strip(),
                    "max_val": row["max_var"].get().strip(),
                    "stuck_n": row["stuck_var"].get().strip(),
                    "expr": row["expr_var"].get().strip(),
                    "plot_live": bool(row["plot_live_var"].get()),
                    "plot_session": bool(row["plot_session_var"].get()),
                    "live_default": idx == default_idx,
                }
            )
        return defs

    def _example_tokens(self):
        raw = self.example_sample_var.get().strip()
        delim = self.delimiter_var.get() or ","
        if not raw:
            messagebox.showwarning("No Example", "Paste one serial output line first.")
            return None
        tokens = [t.strip() for t in raw.split(delim)]
        if len(tokens) < 2:
            messagebox.showerror("Invalid Example", "Example line did not split into at least two fields.")
            return None
        return tokens

    def quick_setup_from_example(self):
        tokens = self._example_tokens()
        if not tokens:
            return
        defs = []
        session_count = min(6, len(tokens))
        for i, _token in enumerate(tokens):
            key = f"column_{i + 1}"
            desc = f"Column {i + 1}"
            defs.append(
                {
                    "index": i,
                    "key": key,
                    "description": desc,
                    "unit": "",
                    "scale": "raw",
                    "min_val": "",
                    "max_val": "",
                    "stuck_n": "",
                    "expr": "",
                    "plot_live": True,
                    "plot_session": i < session_count,
                    "live_default": i == 0,
                }
            )
        self.sample_field_defs = defs
        self._rebuild_measureand_editor_rows(defs)
        self._apply_measureand_config(show_message=False)
        self._persist_sample_setup_defaults()
        self.show_plot_tab()
        self.log(f"Quick sample setup applied from example line: {len(tokens)} fields.")
        messagebox.showinfo(
            "Quick Setup Applied",
            f"Configured {len(tokens)} columns for live plotting.\n"
            "Use Apply Measureands after any manual edits to fine-tune fields.",
        )

    def load_measureands_from_example(self):
        tokens = self._example_tokens()
        if not tokens:
            return

        prev = {d.get("index"): d for d in self.sample_field_defs}
        defs = []
        for i, _token in enumerate(tokens):
            d = prev.get(i, {})
            key = d.get("key") or f"field_{i + 1}"
            desc = d.get("description") or f"Field {i + 1}"
            defs.append(
                {
                    "index": i,
                    "key": self._sanitize_measureand_key(key, i),
                    "description": desc,
                    "unit": d.get("unit", ""),
                    "scale": d.get("scale", "raw"),
                    "min_val": d.get("min_val", ""),
                    "max_val": d.get("max_val", ""),
                    "stuck_n": d.get("stuck_n", ""),
                    "expr": d.get("expr", ""),
                    "plot_live": bool(d.get("plot_live", i < 4)),
                    "plot_session": bool(d.get("plot_session", i < 4)),
                    "live_default": bool(d.get("live_default", i == 0)),
                }
            )
        self._rebuild_measureand_editor_rows(defs)

    def apply_measureands_from_editor(self):
        defs = self._current_measureand_defs_from_editor()
        self.sample_field_defs = defs
        self._apply_measureand_config(show_message=True)
        self._persist_sample_setup_defaults()

    def reset_measureands_default(self):
        self.sample_field_defs = self._default_sample_field_defs()
        self._rebuild_measureand_editor_rows(self.sample_field_defs)
        self._apply_measureand_config(show_message=True)
        self._persist_sample_setup_defaults()

    def _apply_measureand_config(self, show_message=False):
        self.sample_field_defs = sorted(self.sample_field_defs, key=lambda d: d.get("index", 0))
        live_fields = {}
        session_fields = dict(self.base_session_plot_fields)
        self.field_meta_by_key = {}
        self.derived_fields = []

        for d in self.sample_field_defs:
            key = d["key"]
            desc = d["description"]
            min_val = self._to_float_or_none(d.get("min_val", ""))
            max_val = self._to_float_or_none(d.get("max_val", ""))
            stuck_n = self._to_int_or_none(d.get("stuck_n", ""))
            expr = (d.get("expr", "") or "").strip()
            self.field_meta_by_key[key] = {
                "description": desc,
                "unit": d.get("unit", ""),
                "scale": d.get("scale", "raw") if d.get("scale", "raw") in self.unit_scale_factors else "raw",
                "min_val": min_val,
                "max_val": max_val,
                "stuck_n": stuck_n,
                "expr": expr,
            }
            if expr:
                self.derived_fields.append({"key": key, "expr": expr})
            if d.get("plot_live"):
                live_fields[desc] = key
            if d.get("plot_session"):
                session_fields[f"{desc} Std"] = f"{key}_std"
                session_fields[f"{desc} Avg"] = f"{key}_avg"

        if not live_fields:
            first = self.sample_field_defs[0] if self.sample_field_defs else {"description": "Field 1", "key": "field_1"}
            live_fields[first["description"]] = first["key"]

        self.live_plot_fields = live_fields
        self.session_plot_fields = session_fields
        current = self.live_field_var.get() if hasattr(self, "live_field_var") else ""
        live_labels = list(self.live_plot_fields.keys())
        if hasattr(self, "live_field_combo"):
            self.live_field_combo.configure(values=live_labels)
        if current not in self.live_plot_fields:
            preferred = None
            for d in self.sample_field_defs:
                if d.get("live_default") and d.get("plot_live"):
                    preferred = d["description"]
                    break
            self.live_field_var.set(preferred or live_labels[0])
        self._reset_live_series_for_current_fields()
        self.update_live_std_label()
        self._update_live_samples_label()
        self.refresh_live_plot()
        if show_message:
            messagebox.showinfo("Measureands Updated", f"Configured {len(self.sample_field_defs)} fields for parser/plots.")

    def _reset_live_series_for_current_fields(self):
        keys = list(self.live_plot_fields.values())
        for port in list(self.live_run_series_by_port.keys()):
            self.live_run_series_by_port[port] = {k: [] for k in keys}

    @staticmethod
    def _to_float_or_none(v):
        text = str(v).strip()
        if not text:
            return None
        try:
            return float(text)
        except Exception:
            return None

    @staticmethod
    def _to_int_or_none(v):
        text = str(v).strip()
        if not text:
            return None
        try:
            return int(float(text))
        except Exception:
            return None

    def _is_safe_expr(self, expr):
        return bool(re.fullmatch(r"[a-zA-Z0-9_+\-*/().\s]+", expr or ""))

    def _evaluate_derived_fields(self, parsed):
        for item in self.derived_fields:
            expr = item["expr"]
            if not self._is_safe_expr(expr):
                parsed[item["key"]] = np.nan
                continue
            env = {k: float(v) for k, v in parsed.items() if np.isfinite(v)}
            try:
                parsed[item["key"]] = float(eval(expr, {"__builtins__": {}}, env))
            except Exception:
                parsed[item["key"]] = np.nan

    def _field_scale_factor(self, key):
        meta = self.field_meta_by_key.get(key, {})
        return self.unit_scale_factors.get(meta.get("scale", "raw"), 1.0)

    def _field_label_with_unit(self, key, default_label):
        meta = self.field_meta_by_key.get(key, {})
        unit = (meta.get("unit") or "").strip()
        scale = meta.get("scale", "raw")
        if not unit and "voltage" in key:
            unit = "V"
            if scale == "raw":
                scale = "milli"
        suffix = ""
        if unit:
            if scale == "milli":
                suffix = f"m{unit}"
            elif scale == "micro":
                suffix = f"u{unit}"
            elif scale == "kilo":
                suffix = f"k{unit}"
            else:
                suffix = unit
        return f"{default_label} ({suffix})" if suffix else default_label

    def save_parser_profile(self):
        path = filedialog.asksaveasfilename(
            title="Save Parser Profile",
            initialdir=self.profile_dir,
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        payload = {
            "sample_command": self.sample_command_var.get().strip(),
            "delimiter": self.delimiter_var.get(),
            "trim_prefix": self.parser_trim_prefix_var.get(),
            "token_start": int(self.parser_token_start_var.get()),
            "regex": self.parser_regex_var.get(),
            "fields": self._current_measureand_defs_from_editor(),
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        self.log(f"Saved parser profile: {path}")

    def load_parser_profile(self):
        path = filedialog.askopenfilename(
            title="Load Parser Profile",
            initialdir=self.profile_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as exc:
            messagebox.showerror("Load Failed", f"Could not load profile:\n{exc}")
            return
        self.sample_command_var.set(str(payload.get("sample_command", "tsr")))
        self.delimiter_var.set(str(payload.get("delimiter", ",")))
        self.parser_trim_prefix_var.set(str(payload.get("trim_prefix", "")))
        self.parser_token_start_var.set(int(payload.get("token_start", 0)))
        self.parser_regex_var.set(str(payload.get("regex", "")))
        fields = payload.get("fields", [])
        if not isinstance(fields, list) or not fields:
            messagebox.showwarning("Invalid Profile", "Profile did not contain valid field definitions.")
            return
        self.sample_field_defs = fields
        self._rebuild_measureand_editor_rows(fields)
        self._apply_measureand_config(show_message=True)
        self._persist_sample_setup_defaults()
        self.log(f"Loaded parser profile: {path}")

    def plot_current_session(self):
        if not self.session_rows:
            messagebox.showinfo("No Data", "No unit results in this session yet.")
            return
        self._open_session_plot_window(self.session_rows, f"Session Plot - {self.session_id}")

    def load_session_plot(self):
        path = filedialog.askopenfilename(
            title="Load Session JSON To Plot",
            initialdir=self.session_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        self._load_and_plot_session_file(path)

    def reload_current_session_plot(self):
        current_json = os.path.join(self.session_dir, f"sbe83_session_{self.session_id}.json")
        if os.path.exists(current_json):
            self._load_and_plot_session_file(current_json)
            return
        messagebox.showinfo(
            "Session JSON Missing",
            f"No current session JSON found yet:\n{current_json}\n\nSave Session JSON first, or choose a file manually.",
        )
        self.load_session_plot()

    def _load_and_plot_session_file(self, path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            messagebox.showerror("Load Failed", f"Could not read JSON:\n{path}\n\n{exc}")
            return

        rows = self._normalize_session_rows(data)
        if not rows:
            messagebox.showerror(
                "No Plot Data",
                "Loaded file did not contain plottable numeric session fields.",
            )
            return

        self._open_session_plot_window(rows, f"Session Plot - {os.path.basename(path)}")

    @staticmethod
    def _normalize_session_rows(data):
        if isinstance(data, list):
            return [r for r in data if isinstance(r, dict)]
        if isinstance(data, dict):
            if isinstance(data.get("rows"), list):
                return [r for r in data["rows"] if isinstance(r, dict)]
            return [data]
        return []

    def _open_session_plot_window(self, rows, title):
        if not rows:
            messagebox.showinfo("No Data", "No rows available to plot.")
            return

        plot_fields = self._session_plot_fields_for_rows(rows)
        available = []
        for label, key in plot_fields.items():
            if any(np.isfinite(self._to_float(r.get(key))) for r in rows):
                available.append(label)
        if not available:
            messagebox.showerror("No Plot Data", "No supported numeric metrics found in session rows.")
            return

        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("1050x460")

        top = ttk.Frame(win, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text=f"Current runs: {len(rows)}  |  Compare two sessions by serial").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(top, text="Plot field").pack(side=tk.LEFT)
        metric_label_var = tk.StringVar(value=available[0])
        metric_combo = ttk.Combobox(top, textvariable=metric_label_var, state="readonly", width=32, values=available)
        metric_combo.pack(side=tk.LEFT, padx=(6, 10))
        plot_paused_var = tk.BooleanVar(value=False)
        ttk.Button(top, text="Load Reference Session", command=lambda: load_reference_and_render()).pack(side=tk.LEFT, padx=(0, 8))
        pause_btn = ttk.Button(top, text="Pause Plot")
        pause_btn.pack(side=tk.LEFT, padx=(0, 8))
        reference_name_var = tk.StringVar(value="Reference: not loaded")
        ttk.Label(top, textvariable=reference_name_var, foreground=DARK_MUTED).pack(side=tk.LEFT, padx=(6, 0))

        selectors = ttk.Frame(win, padding=(8, 0, 8, 6))
        selectors.pack(fill=tk.X)
        current_group = ttk.LabelFrame(selectors, text="Current Session Sensors", padding=6)
        current_group.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        ref_group = ttk.LabelFrame(selectors, text="Reference Session Sensors", padding=6)
        ref_group.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        current_list = tk.Listbox(current_group, selectmode=tk.EXTENDED, exportselection=False, height=5)
        current_list.pack(fill=tk.BOTH, expand=True)
        ref_list = tk.Listbox(ref_group, selectmode=tk.EXTENDED, exportselection=False, height=5)
        ref_list.pack(fill=tk.BOTH, expand=True)

        current_rows_var = list(rows)
        reference_rows_var = []

        def _serials_for_rows(in_rows):
            serials = sorted({(str(r.get("serial", "")).strip() or "UNKNOWN") for r in in_rows})
            return serials

        def _populate_sensor_list(listbox, serials):
            listbox.delete(0, tk.END)
            for serial in serials:
                listbox.insert(tk.END, serial)
            if serials:
                listbox.selection_set(0, tk.END)

        def _selected_serials(listbox):
            selected_idx = listbox.curselection()
            if not selected_idx:
                return set(listbox.get(0, tk.END))
            return {listbox.get(i) for i in selected_idx}

        _populate_sensor_list(current_list, _serials_for_rows(current_rows_var))
        _populate_sensor_list(ref_list, [])

        canvas = tk.Canvas(win, bg="#0b1220", highlightthickness=1, highlightbackground=DARK_BORDER)
        canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        def render():
            if plot_paused_var.get():
                return
            metric_key = plot_fields.get(metric_label_var.get(), "red_noise_ns")
            current_serials = _selected_serials(current_list)
            ref_serials = _selected_serials(ref_list)
            current_filtered = [r for r in current_rows_var if (str(r.get("serial", "")).strip() or "UNKNOWN") in current_serials]
            ref_filtered = [r for r in reference_rows_var if (str(r.get("serial", "")).strip() or "UNKNOWN") in ref_serials]
            self._draw_session_metric_plot(
                canvas,
                current_filtered,
                ref_filtered,
                metric_key,
                metric_label_var.get(),
                current_label="Current",
                reference_label="Reference",
            )

        def load_reference_and_render():
            loaded = self._load_reference_rows()
            if not loaded:
                return
            reference_rows_var.clear()
            reference_rows_var.extend(self.reference_session_rows)
            reference_name_var.set(f"Reference: {os.path.basename(self.reference_session_path)} ({len(reference_rows_var)} runs)")
            _populate_sensor_list(ref_list, _serials_for_rows(reference_rows_var))
            merged_fields = self._session_plot_fields_for_rows(current_rows_var + reference_rows_var)
            merged_available = []
            for label, key in merged_fields.items():
                if any(np.isfinite(self._to_float(r.get(key))) for r in (current_rows_var + reference_rows_var)):
                    merged_available.append(label)
            if merged_available:
                metric_combo.configure(values=merged_available)
                if metric_label_var.get() not in merged_available:
                    metric_label_var.set(merged_available[0])
                plot_fields.clear()
                plot_fields.update(merged_fields)
            render()

        def toggle_pause():
            paused = not plot_paused_var.get()
            plot_paused_var.set(paused)
            pause_btn.configure(text="Resume Plot" if paused else "Pause Plot")
            if not paused:
                render()

        pause_btn.configure(command=toggle_pause)

        metric_combo.bind("<<ComboboxSelected>>", lambda _e: render())
        current_list.bind("<<ListboxSelect>>", lambda _e: render())
        ref_list.bind("<<ListboxSelect>>", lambda _e: render())
        canvas.bind("<Configure>", lambda _e: render())
        render()

    def _load_reference_rows(self):
        path = filedialog.askopenfilename(
            title="Load Reference Session JSON",
            initialdir=self.session_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return False
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.reference_session_rows = self._normalize_session_rows(data)
            self.reference_session_path = path
            self.log(f"Loaded reference rows: {len(self.reference_session_rows)} from {path}")
            return bool(self.reference_session_rows)
        except Exception as exc:
            messagebox.showerror("Reference Load Failed", str(exc))
            return False

    def _session_plot_fields_for_rows(self, rows):
        fields = dict(self.session_plot_fields)
        keys = set()
        for row in rows:
            keys.update(row.keys())
        for key in sorted(keys):
            if key.endswith("_std"):
                base = key[: -len("_std")]
                label = base.replace("_", " ").title()
                fields.setdefault(f"{label} Std", key)
            elif key.endswith("_avg"):
                base = key[: -len("_avg")]
                label = base.replace("_", " ").title()
                fields.setdefault(f"{label} Avg", key)
        return fields

    @staticmethod
    def _to_float(val):
        try:
            return float(val)
        except (TypeError, ValueError):
            return np.nan

    @staticmethod
    def _session_metric_color(metric_key):
        if "red" in metric_key:
            return "#dc2626"
        if "blue" in metric_key:
            return "#2563eb"
        if "temp" in metric_key:
            return "#059669"
        if "pll" in metric_key:
            return "#7c3aed"
        return "#4b5563"

    def _draw_session_metric_plot(
        self,
        canvas,
        current_rows,
        reference_rows,
        metric_key,
        metric_label,
        current_label="Current",
        reference_label="Reference",
    ):
        canvas.delete("all")
        if not current_rows and not reference_rows:
            canvas.create_text(20, 20, text="No session data to plot.", anchor="nw", fill=DARK_MUTED)
            return

        width = max(int(canvas.winfo_width()), 320)
        height = max(int(canvas.winfo_height()), 220)
        left, top, right, bottom = 70, 30, width - 20, height - 55
        pw = max(right - left, 1)
        ph = max(bottom - top, 1)

        rows_by_name = {
            current_label: list(current_rows),
            reference_label: list(reference_rows),
        }
        finite_raw_y = []
        for rows in rows_by_name.values():
            for row in rows:
                val = self._to_float(row.get(metric_key))
                if np.isfinite(val):
                    finite_raw_y.append(val)
        if not finite_raw_y:
            canvas.create_text(20, 20, text=f"No finite values for {metric_label}.", anchor="nw", fill=DARK_MUTED)
            return

        base_key = metric_key
        if metric_key.endswith("_std"):
            base_key = metric_key[: -len("_std")]
        elif metric_key.endswith("_avg"):
            base_key = metric_key[: -len("_avg")]
        scale_factor = self._field_scale_factor(base_key)
        y_axis_label = self._field_label_with_unit(base_key, metric_label)

        y_vals = [v * scale_factor for v in finite_raw_y]
        ymin = min(y_vals)
        ymax = max(y_vals)
        if ymin == ymax:
            pad = 1.0 if ymin == 0 else abs(ymin) * 0.1
            ymin -= pad
            ymax += pad
        else:
            pad = (ymax - ymin) * 0.08
            ymin -= pad
            ymax += pad

        canvas.create_rectangle(left, top, right, bottom, outline=DARK_BORDER, width=1)
        y_ticks = 5
        for i in range(y_ticks + 1):
            frac = i / y_ticks
            y = bottom - frac * ph
            val = ymin + frac * (ymax - ymin)
            canvas.create_line(left, y, right, y, fill="#1f2937")
            canvas.create_text(left - 8, y, text=self.fmt(val), anchor="e", fill=DARK_MUTED)

        serials = []
        for rows in rows_by_name.values():
            for row in rows:
                serial = str(row.get("serial", "")).strip() or "UNKNOWN"
                if serial not in serials:
                    serials.append(serial)
        n_serials = len(serials)
        x_pad = min(max(width * 0.04, 20.0), 60.0)
        x_left = left + x_pad
        x_right = right - x_pad
        if n_serials <= 1:
            serial_base_x = {serials[0]: (x_left + x_right) / 2.0}
        else:
            step = (x_right - x_left) / (n_serials - 1)
            start = x_left
            serial_base_x = {
                serial: start + (i * step)
                for i, serial in enumerate(serials)
            }

        for i, serial in enumerate(serials):
            x = serial_base_x[serial]
            if n_serials == 1:
                band_left, band_right = x_left, x_right
            else:
                prev_x = serial_base_x[serials[i - 1]] if i > 0 else x_left
                next_x = serial_base_x[serials[i + 1]] if i < n_serials - 1 else x_right
                band_left = (prev_x + x) / 2.0 if i > 0 else x_left
                band_right = (x + next_x) / 2.0 if i < n_serials - 1 else x_right
            shade = "#0f172a" if i % 2 == 0 else "#111827"
            canvas.create_rectangle(band_left, top, band_right, bottom, fill=shade, outline="")

        label_step = max(1, n_serials // 12)
        for i, serial in enumerate(serials):
            x = serial_base_x[serial]
            canvas.create_line(x, top, x, bottom, fill="#1f2937", dash=(2, 3))
            if i % label_step == 0 or i == n_serials - 1:
                canvas.create_text(
                    x,
                    bottom + 14,
                    text=serial,
                    anchor="n",
                    fill=DARK_MUTED,
                    font=("Segoe UI", 7),
                )
        canvas.create_line(x_left, top, x_left, bottom, fill=DARK_BORDER)
        canvas.create_line(x_right, top, x_right, bottom, fill=DARK_BORDER)

        canvas.create_text((left + right) / 2, height - 18, text="Sensor Serial", anchor="center", fill=DARK_TEXT)
        canvas.create_text(18, (top + bottom) / 2, text=y_axis_label, angle=90, anchor="center", fill=DARK_TEXT)

        y_den = max(ymax - ymin, 1e-12)
        point_meta = []
        both_sessions = bool(current_rows) and bool(reference_rows)
        session_colors = {
            current_label: "#f59e0b",
            reference_label: "#22c55e",
        }
        session_offset = {
            current_label: -8.0 if both_sessions else 0.0,
            reference_label: 8.0 if both_sessions else 0.0,
        }
        for session_name, rows in rows_by_name.items():
            color = session_colors.get(session_name, self._session_metric_color(metric_key))
            serial_row_idxs = {serial: [] for serial in serials}
            for idx, row in enumerate(rows):
                serial = str(row.get("serial", "")).strip() or "UNKNOWN"
                serial_row_idxs[serial].append(idx)

            x_positions = [left + pw / 2.0] * len(rows)
            for serial in serials:
                idxs = serial_row_idxs[serial]
                count = len(idxs)
                if count == 0:
                    continue
                if n_serials > 1:
                    neighbor_gap = (x_right - x_left) / (n_serials - 1)
                else:
                    neighbor_gap = max(x_right - x_left, 1.0)
                cluster_span = min(20.0, neighbor_gap * 0.35)
                intra_step = 0.0 if count <= 1 else (cluster_span / (count - 1))
                start = -((count - 1) * intra_step) / 2.0
                for j, row_idx in enumerate(idxs):
                    x_positions[row_idx] = serial_base_x[serial] + session_offset[session_name] + start + (j * intra_step)

            for i, row in enumerate(rows):
                raw_val = self._to_float(row.get(metric_key))
                if not np.isfinite(raw_val):
                    continue
                scaled_val = raw_val * scale_factor
                x = x_positions[i]
                y = bottom - ((scaled_val - ymin) / y_den) * ph
                canvas.create_oval(x - 3, y - 3, x + 3, y + 3, fill=color, outline=color)
                run_idx = row.get("run_index", i + 1)
                serial = str(row.get("serial", "")).strip() or "UNKNOWN"
                point_meta.append(
                    {
                        "x": x,
                        "y": y,
                        "text": f"{session_name} | {serial} | r{run_idx} | y={self.fmt(scaled_val)}",
                    }
                )

        legend = []
        if current_rows:
            legend.append((current_label, session_colors[current_label], len(current_rows)))
        if reference_rows:
            legend.append((reference_label, session_colors[reference_label], len(reference_rows)))
        legend_x = left + 8
        legend_y = top + 8
        for i, (name, color, nrows) in enumerate(legend):
            y = legend_y + i * 15
            canvas.create_line(legend_x, y, legend_x + 16, y, fill=color, width=3)
            canvas.create_text(legend_x + 22, y, text=f"{name} ({nrows})", anchor="w", fill=DARK_TEXT)

        self._bind_session_plot_hover(canvas, point_meta)

    def _bind_session_plot_hover(self, canvas, point_meta):
        canvas._session_hover_points = point_meta

        def on_motion(event):
            points = getattr(canvas, "_session_hover_points", [])
            if not points:
                canvas.delete("hover_tip")
                return
            best = None
            best_d2 = 81.0
            for p in points:
                dx = p["x"] - event.x
                dy = p["y"] - event.y
                d2 = dx * dx + dy * dy
                if d2 <= best_d2:
                    best_d2 = d2
                    best = p
            canvas.delete("hover_tip")
            if not best:
                return
            tx = best["x"] + 10
            ty = best["y"] - 10
            text_id = canvas.create_text(tx, ty, text=best["text"], anchor="sw", fill=DARK_TEXT, tags="hover_tip")
            x0, y0, x1, y1 = canvas.bbox(text_id)
            canvas.create_rectangle(x0 - 4, y0 - 2, x1 + 4, y1 + 2, fill="#0f172a", outline=DARK_BORDER, tags="hover_tip")
            canvas.tag_raise(text_id)

        def on_leave(_event):
            canvas.delete("hover_tip")

        if not getattr(canvas, "_session_hover_bound", False):
            canvas.bind("<Motion>", on_motion)
            canvas.bind("<Leave>", on_leave)
            canvas._session_hover_bound = True

    @staticmethod
    def _sort_value(col, raw):
        if raw is None:
            return (1, "")
        text = str(raw).strip()
        if not text:
            return (1, "")
        if col == "timestamp":
            try:
                return (0, dt.datetime.fromisoformat(text))
            except ValueError:
                return (1, text.lower())
        numeric_cols = {
            "red_ns",
            "red_v_std",
            "red_v_avg",
            "blue_ns",
            "blue_v_std",
            "blue_v_avg",
            "red_pll_v_std",
            "red_pll_v_avg",
            "blue_pll_v_std",
            "blue_pll_v_avg",
            "raw_temp_v_std",
            "raw_temp_v_avg",
            "elec_temp_v_std",
            "elec_temp_v_avg",
        }
        if col in numeric_cols:
            try:
                return (0, float(text))
            except ValueError:
                return (1, text.lower())
        return (0, text.lower())

    def sort_tree_by_column(self, col):
        descending = self.tree_sort_state.get(col, False)
        self.tree_sort_state[col] = not descending

        row_ids = list(self.tree.get_children(""))
        row_ids.sort(
            key=lambda iid: self._sort_value(col, self.tree.set(iid, col)),
            reverse=descending,
        )
        for idx, iid in enumerate(row_ids):
            self.tree.move(iid, "", idx)

        for heading_col in self.tree["columns"]:
            suffix = ""
            if heading_col == col:
                suffix = " (desc)" if descending else " (asc)"
            self.tree.heading(heading_col, text=f"{heading_col}{suffix}", command=lambda c=heading_col: self.sort_tree_by_column(c))

    def toggle_csv_column(self):
        display_cols = list(self.tree["displaycolumns"])
        if "sample_csv" in display_cols:
            display_cols = [c for c in display_cols if c != "sample_csv"]
        else:
            display_cols = list(self.tree["columns"])
        self.tree.configure(displaycolumns=tuple(display_cols))

    def _build_port_grid(self, parent):
        self.port_grid_parent = parent
        self.port_grid_empty_label = tk.Label(
            parent,
            text="No connected ports",
            bg=DARK_BG,
            fg=DARK_MUTED,
            font=("Segoe UI", 10),
        )
        for idx in range(MAX_PORTS):
            r = idx // 5
            c = idx % 5
            card = tk.Frame(
                parent,
                bd=0,
                relief=tk.FLAT,
                padx=4,
                pady=2,
                bg=DARK_PANEL,
                highlightthickness=1,
                highlightbackground=DARK_BORDER,
            )
            card.grid(row=r, column=c, padx=3, pady=2, sticky="nsew")

            slot_var = tk.StringVar(value=f"Slot {idx + 1}")
            port_var = tk.StringVar(value="(empty)")
            serial_var = tk.StringVar(value="SN: -")
            state_var = tk.StringVar(value="DISCONNECTED")

            slot_label = tk.Label(card, textvariable=slot_var, font=("Segoe UI Semibold", 10), bg=DARK_PANEL, fg=DARK_TEXT)
            slot_label.pack(anchor="w")
            port_label = tk.Label(card, textvariable=port_var, fg=DARK_TEXT, bg=DARK_PANEL, font=("Segoe UI", 10))
            port_label.pack(anchor="w")
            serial_label = tk.Label(card, textvariable=serial_var, fg=DARK_MUTED, bg=DARK_PANEL, font=("Segoe UI", 9))
            serial_label.pack(anchor="w")
            state_label = tk.Label(
                card, textvariable=state_var, fg="white", bg="#6b7280", width=11, relief=tk.FLAT, padx=4, pady=2, font=("Segoe UI Semibold", 9)
            )
            state_label.pack(anchor="w", pady=(2, 0))

            self.port_slots[idx] = {
                "card": card,
                "slot_var": slot_var,
                "port_var": port_var,
                "serial_var": serial_var,
                "state_var": state_var,
                "slot_label": slot_label,
                "port_label": port_label,
                "serial_label": serial_label,
                "state_label": state_label,
            }

        for c in range(5):
            parent.grid_columnconfigure(c, weight=0)

    def _ui_post(self, event, *args, **kwargs):
        self.ui_event_queue.put((event, args, kwargs))

    def _log_main(self, msg: str):
        ts = dt.datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state=tk.NORMAL)
        self.log_box.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state=tk.DISABLED)
        
    def log(self, msg: str):
        if threading.current_thread() is threading.main_thread():
            self._log_main(msg)
            return
        self._ui_post("log", msg)

    def refresh_ports(self):
        ports = sorted({p.device for p in list_ports.comports()})
        self.available_ports = ports[:MAX_PORTS]
        if not self.available_ports:
            self.available_ports = ["COM5"]
        self.com_combo["values"] = self.available_ports
        if self.com_var.get() not in self.available_ports:
            self.com_var.set(self.available_ports[0])
        self.sync_debug_tabs()
        self.update_port_grid()
        self.log(f"Ports detected (max {MAX_PORTS} shown): {', '.join(self.available_ports)}")

    def update_connection_labels(self):
        connected = sorted(self.serial_pool.keys())
        self.conn_status.set(f"Connected ports: {len(connected)}")
        self.connected_ports_var.set(", ".join(connected) if connected else "None")
        if not self.live_visible_ports:
            self.live_visible_ports = set(connected)
        self.update_run_button_state()
        self.sync_debug_tabs()
        self.update_port_grid()

    def _on_setup_field_changed(self, *_):
        self.update_run_button_state()
        self._maybe_auto_hide_test_setup()

    def _setup_fields_ready_for_auto_hide(self):
        required_text = (
            self.operator_var.get().strip(),
            self.bath_id_var.get().strip(),
            self.notes_var.get().strip(),
            self.bath_temp_c_var.get().strip(),
            self.salinity_psu_var.get().strip(),
            self.results_root_var.get().strip(),
        )
        if not all(required_text):
            return False
        try:
            return int(self.sample_count_var.get()) >= 20
        except Exception:
            return False

    def _maybe_auto_hide_test_setup(self):
        if not self._setup_fields_ready_for_auto_hide():
            self.setup_manual_override = False
            return
        if not self.test_setup_visible:
            return
        if self.setup_manual_override:
            return
        self._set_test_setup_visibility(False, auto=True)

    def required_setup_values(self):
        return {
            "Operator": self.operator_var.get().strip(),
        }

    def missing_setup_fields(self):
        return [name for name, value in self.required_setup_values().items() if not value]

    def update_run_button_state(self):
        connected = bool(self.serial_pool)
        setup_complete = not self.missing_setup_fields()
        allow_run = connected and setup_complete and not self.run_in_progress
        self.run_btn.configure(state=tk.NORMAL if allow_run else tk.DISABLED)

    def _update_runs_left_label(self):
        if not self.batch_runs_remaining_by_port:
            self.runs_left_var.set("Runs left: n/a")
            return
        total_left = sum(max(0, int(v)) for v in self.batch_runs_remaining_by_port.values())
        parts = [f"{port}:{max(0, int(v))}" for port, v in sorted(self.batch_runs_remaining_by_port.items())]
        self.runs_left_var.set(f"Runs left: {total_left} total ({' | '.join(parts)})")

    def status_color(self, status):
        return self._status_colors().get(status, self._status_colors()["DISCONNECTED"])

    def find_slot_by_port(self, port):
        for idx in range(MAX_PORTS):
            if self.port_slots[idx]["port_var"].get() == port:
                return idx
        return None

    def set_port_status(self, port, status, serial=None):
        if threading.current_thread() is not threading.main_thread():
            self._ui_post("set_port_status", port, status, serial=serial)
            return
        idx = self.find_slot_by_port(port)
        if idx is None:
            return
        slot = self.port_slots[idx]
        slot["state_var"].set(status)
        state_bg, state_fg = self.status_color(status)
        slot["state_label"].configure(bg=state_bg, fg=state_fg)
        if serial:
            slot["serial_var"].set(f"SN: {serial}")
        elif status == "DISCONNECTED":
            slot["serial_var"].set("SN: -")

    def sync_debug_tabs(self):
        for port in sorted(set(self.available_ports) | set(self.serial_pool.keys())):
            self.ensure_debug_tab(port)

    def ensure_debug_tab(self, port):
        if port in self.debug_tabs:
            return self.debug_tabs[port]["text"]
        tab = ttk.Frame(self.debug_notebook)
        text = scrolledtext.ScrolledText(
            tab,
            height=9,
            wrap=tk.NONE,
            font=self._console_font(),
            bg=DARK_PANEL,
            fg=DARK_TEXT,
            insertbackground=DARK_TEXT,
            relief=tk.FLAT,
        )
        text.pack(fill=tk.BOTH, expand=True, pady=(0, 6))
        text.insert(tk.END, "Terminal mode: type command here and press Enter to send.\n")
        text.bind("<Return>", lambda e, p=port: self._on_debug_text_enter(e, p))
        controls = ttk.Frame(tab)
        controls.pack(fill=tk.X)
        cmd_var = tk.StringVar()
        entry = ttk.Entry(controls, textvariable=cmd_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        entry.bind("<Return>", lambda _e, p=port: self.send_debug_command(p))
        ttk.Button(controls, text="Send", command=lambda p=port: self.send_debug_command(p)).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(controls, text="Read", command=lambda p=port: self.read_debug_responses(p)).pack(side=tk.LEFT, padx=(6, 0))
        stream_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            controls, text="Stream", variable=stream_var, command=lambda p=port, v=stream_var: self.toggle_stream(p, v.get())
        ).pack(side=tk.LEFT, padx=(6, 0))
        self.debug_notebook.add(tab, text=port)
        self.debug_tabs[port] = {"tab": tab, "text": text, "lines": 1, "cmd_var": cmd_var, "stream_var": stream_var}
        self.stream_enabled[port] = False
        return text

    def _on_debug_text_enter(self, event, port):
        box = self.debug_tabs[port]["text"]
        line = box.get("insert linestart", "insert lineend").strip()
        if line:
            line = line.lstrip(">").strip()
            self.send_debug_command(port, cmd=line)
        self._ensure_console_trailing_newline(port)
        return "break"

    def _ensure_console_trailing_newline(self, port):
        info = self.debug_tabs.get(port)
        if not info:
            return
        box = info["text"]
        tail = box.get("end-2c", "end-1c")
        if tail != "\n":
            box.insert(tk.END, "\n")
        box.mark_set(tk.INSERT, tk.END)
        box.see(tk.END)

    def clear_debug_tab(self, port):
        info = self.debug_tabs.get(port)
        if not info:
            return
        box = info["text"]
        box.delete("1.0", tk.END)
        box.insert(tk.END, "Terminal mode: type command here and press Enter to send.\n")
        info["lines"] = 0

    def clear_selected_debug_tab(self):
        if not self.debug_notebook.tabs():
            return
        current = self.debug_notebook.select()
        for port, info in self.debug_tabs.items():
            if str(info["tab"]) != current:
                continue
            self.clear_debug_tab(port)
            self.log(f"Cleared serial debug tab: {port}")
            break

    def port_is_running(self, port):
        with self.run_state_lock:
            return port in self.active_run_ports

    def reconnect_selected_port(self):
        port = self.com_var.get().strip()
        if not port:
            messagebox.showwarning("No Port", "Select a COM port first.")
            return
        self.disconnect_selected_port()
        self.connect_selected_port()

    def send_debug_command(self, port, cmd=None):
        if self.port_is_running(port):
            messagebox.showwarning("Port Busy", f"{port} is running a test. Wait for completion before manual commands.")
            return
        ser = self.serial_pool.get(port)
        if not ser or not ser.is_open:
            messagebox.showwarning("Not Connected", f"{port} is not connected.")
            return
        from_entry = cmd is None
        if cmd is None:
            cmd = self.debug_tabs[port]["cmd_var"].get().strip()
        else:
            cmd = str(cmd).strip()
        if not cmd:
            messagebox.showwarning("No Command", "Type a command before pressing Send.")
            return
        try:
            payload_bytes = self._console_command_bytes(cmd)
            if from_entry:
                box = self.debug_tabs[port]["text"]
                box.insert(tk.END, f"> {cmd}\n")
                self._ensure_console_trailing_newline(port)
            self.serial_debug(port, "TX", payload_bytes)
            ser.write(payload_bytes)
            self.debug_tabs[port]["cmd_var"].set("")
            self._ensure_console_trailing_newline(port)
            threading.Thread(target=self._read_debug_responses_quick, args=(port, 0.35), daemon=True).start()
        except Exception as exc:
            self.log(f"[{port}] Manual command failed: {exc}")
            messagebox.showerror("Manual Command Error", f"{port}: {exc}")

    def _read_debug_responses_quick(self, port, window_s=0.35):
        ser = self.serial_pool.get(port)
        if not ser or not ser.is_open or self.port_is_running(port) or self.stream_enabled.get(port):
            return
        old_timeout = ser.timeout
        try:
            ser.timeout = 0.08
            deadline = time.time() + window_s
            while time.time() < deadline and not self.shutdown_event.is_set():
                self._read_debug_line(ser, port=port)
        except Exception as exc:
            self.log(f"[{port}] Quick read failed: {exc}")
        finally:
            try:
                ser.timeout = old_timeout
            except Exception:
                pass

    def read_debug_responses(self, port, window_s=1.2):
        if self.port_is_running(port):
            messagebox.showwarning("Port Busy", f"{port} is running a test. Wait for completion before manual reads.")
            return
        if self.stream_enabled.get(port):
            messagebox.showinfo("Stream Active", f"{port} stream mode is active. Stop stream for manual read window.")
            return
        ser = self.serial_pool.get(port)
        if not ser or not ser.is_open:
            messagebox.showwarning("Not Connected", f"{port} is not connected.")
            return
        old_timeout = ser.timeout
        rx_count = 0
        try:
            ser.timeout = 0.15
            deadline = time.time() + window_s
            while time.time() < deadline:
                raw = self._read_debug_line(ser, port=port)
                if not raw:
                    continue
                rx_count += 1
        finally:
            ser.timeout = old_timeout
        if rx_count == 0:
            self.log(f"[{port}] Manual read: no response")

    def serial_debug(self, port, direction, payload):
        if threading.current_thread() is not threading.main_thread():
            self._ui_post("debug_line", port, direction, payload)
            return
        self._append_debug_line(port, direction, payload)

    def _append_debug_line(self, port, direction, payload):
        if not port:
            return
        payload_text = self._format_console_payload(payload)
        box = self.ensure_debug_tab(port)
        info = self.debug_tabs[port]
        ts = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{ts}] {direction}: {payload_text}\n"
        box.insert(tk.END, line)
        info["lines"] += 1
        if info["lines"] > self.debug_max_lines:
            box.delete("1.0", "2.0")
            info["lines"] -= 1
        box.see(tk.END)
        self.manual_capture_rows.append(
            {
                "timestamp": dt.datetime.now().isoformat(timespec="milliseconds"),
                "port": port,
                "direction": direction,
                "payload": payload_text,
            }
        )

    def _process_ui_events(self):
        drained = 0
        while drained < 180:
            try:
                event, args, kwargs = self.ui_event_queue.get_nowait()
            except queue.Empty:
                break
            if event == "debug_line":
                self._append_debug_line(*args, **kwargs)
            elif event == "log":
                self._log_main(*args, **kwargs)
            elif event == "set_port_status":
                self.set_port_status(*args, **kwargs)
            elif event == "clear_live_run_view":
                self.clear_live_run_view(*args, **kwargs)
            elif event == "append_live_run_sample":
                self.append_live_run_sample(*args, **kwargs)
            elif event == "run_result":
                self._apply_run_result(*args, **kwargs)
            elif event == "finish_port_run":
                self._finish_port_run(*args, **kwargs)
            elif event == "show_error":
                messagebox.showerror(*args, **kwargs)
            elif event == "show_warning":
                messagebox.showwarning(*args, **kwargs)
            elif event == "show_info":
                messagebox.showinfo(*args, **kwargs)
            drained += 1
        if not self.shutdown_event.is_set():
            self.root.after(60, self._process_ui_events)

    def toggle_stream(self, port, enabled):
        self.stream_enabled[port] = bool(enabled)
        if enabled:
            self.start_stream_reader(port)
        else:
            self.stop_stream_reader(port)

    def start_stream_reader(self, port):
        ser = self.serial_pool.get(port)
        if not ser or not ser.is_open:
            return
        if port in self.stream_threads and self.stream_threads[port].is_alive():
            return
        stop_event = threading.Event()
        self.stream_stop_events[port] = stop_event

        def worker():
            old_timeout = ser.timeout
            try:
                ser.timeout = 0.08
                while not stop_event.is_set() and not self.shutdown_event.is_set():
                    if self.port_is_running(port):
                        time.sleep(0.1)
                        continue
                    raw = ser.readline()
                    if raw:
                        self.serial_debug(port, "RX", raw)
            except Exception as exc:
                self.log(f"[{port}] Stream reader stopped: {exc}")
            finally:
                try:
                    ser.timeout = old_timeout
                except Exception:
                    pass

        t = threading.Thread(target=worker, daemon=True)
        self.stream_threads[port] = t
        t.start()

    def stop_stream_reader(self, port):
        ev = self.stream_stop_events.get(port)
        if ev:
            ev.set()

    def export_manual_capture(self):
        if not self.manual_capture_rows:
            messagebox.showinfo("No Data", "No console TX/RX rows captured yet.")
            return
        path = filedialog.asksaveasfilename(
            title="Export Console Capture CSV",
            initialdir=self.session_dir,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["timestamp", "port", "direction", "payload"])
            writer.writeheader()
            writer.writerows(self.manual_capture_rows)
        self.log(f"Console capture exported: {path}")

    def update_port_grid(self):
        # Only show currently connected ports to keep the station view compact.
        connected = [p for p in sorted(self.serial_pool.keys()) if self.serial_pool[p].is_open][:MAX_PORTS]
        n = len(connected)

        if n <= 0:
            cols = 1
        elif n <= 2:
            cols = n
        elif n <= 4:
            cols = 2
        elif n <= 6:
            cols = 3
        elif n <= 8:
            cols = 4
        else:
            cols = 5

        for c in range(5):
            self.port_grid_parent.grid_columnconfigure(c, weight=1 if c < cols else 0)

        if n == 0:
            self.port_grid_empty_label.grid(row=0, column=0, columnspan=5, sticky="w", padx=6, pady=4)
        else:
            self.port_grid_empty_label.grid_remove()

        for idx in range(MAX_PORTS):
            slot = self.port_slots[idx]
            card = slot["card"]
            if idx < n:
                port = connected[idx]
                slot["port_var"].set(port)
                if slot["state_var"].get() not in {"RUNNING", "COMPLETE", "PASS", "WARN", "FAIL", "ERROR"}:
                    slot["state_var"].set("CONNECTED")
                if slot["serial_var"].get() == "SN: -":
                    slot["serial_var"].set("SN: (pending)")
                state_bg, state_fg = self.status_color(slot["state_var"].get())
                slot["state_label"].configure(bg=state_bg, fg=state_fg)
                r = idx // cols
                c = idx % cols
                card.grid(row=r, column=c, padx=3, pady=2, sticky="nsew")
            else:
                slot["port_var"].set("(empty)")
                slot["serial_var"].set("SN: -")
                slot["state_var"].set("DISCONNECTED")
                state_bg, state_fg = self.status_color("DISCONNECTED")
                slot["state_label"].configure(bg=state_bg, fg=state_fg)
                card.grid_remove()

    def connect_selected_port(self):
        port = self.com_var.get().strip()
        if not port:
            messagebox.showwarning("No Port", "Select a COM port first.")
            return
        if port in self.serial_pool and self.serial_pool[port].is_open:
            self.log(f"Port already connected: {port}")
            self.update_connection_labels()
            return
        try:
            baud = int(self.baudrate_var.get())
            ser = serial.Serial(port=port, baudrate=baud, bytesize=8, parity="N", stopbits=1, timeout=2)
            self.serial_pool[port] = ser
            self.log(f"Connected: {port} @ {baud}")
            self.set_port_status(port, "CONNECTED")
        except Exception as exc:
            messagebox.showerror("Connection Error", f"{port}: {exc}")
            self.set_port_status(port, "ERROR")
        self.update_connection_labels()

    def disconnect_selected_port(self):
        port = self.com_var.get().strip()
        self.stop_stream_reader(port)
        ser = self.serial_pool.pop(port, None)
        if not ser:
            self.log(f"Port not connected: {port}")
            self.update_connection_labels()
            return
        try:
            ser.close()
        except Exception:
            pass
        self.log(f"Disconnected: {port}")
        self.set_port_status(port, "DISCONNECTED")
        self.update_connection_labels()

    def connect_all_ports(self):
        self.refresh_ports()
        count = 0
        baud = int(self.baudrate_var.get())
        for port in self.available_ports:
            if port in self.serial_pool and self.serial_pool[port].is_open:
                continue
            try:
                self.serial_pool[port] = serial.Serial(
                    port=port, baudrate=baud, bytesize=8, parity="N", stopbits=1, timeout=2
                )
                count += 1
                self.set_port_status(port, "CONNECTED")
            except Exception as exc:
                self.log(f"Connect failed {port}: {exc}")
                self.set_port_status(port, "ERROR")
        self.log(f"Connect-all complete: {count} new connection(s) @ {baud}")
        self.update_connection_labels()

    def disconnect_all_ports(self):
        ports = list(self.serial_pool.keys())
        for port in ports:
            self.stop_stream_reader(port)
            try:
                self.serial_pool[port].close()
            except Exception:
                pass
            self.serial_pool.pop(port, None)
            self.set_port_status(port, "DISCONNECTED")
        self.log("All ports disconnected")
        self.update_connection_labels()

    def _console_command_bytes(self, cmd):
        suffix = b""
        if self.console_send_cr_var.get():
            suffix += b"\r"
        if self.console_send_lf_var.get():
            suffix += b"\n"
        return str(cmd).encode("utf-8", errors="replace") + suffix

    def _format_console_payload(self, payload):
        raw = payload if isinstance(payload, (bytes, bytearray)) else str(payload).encode("utf-8", errors="replace")
        mode = self.console_display_mode_var.get()
        if mode == "hex":
            return " ".join(f"{b:02X}" for b in raw)
        if mode == "dec":
            return " ".join(str(b) for b in raw)
        if mode == "bin":
            return " ".join(f"{b:08b}" for b in raw)

        out = []
        for b in raw:
            if 32 <= b <= 126:
                out.append(chr(b))
            elif b == 13:
                out.append("<CR>")
            elif b == 10:
                out.append("<LF>")
            elif b == 9:
                out.append("<TAB>")
            else:
                out.append(f"<0x{b:02X}>")
        return "".join(out)

    def _read_debug_line(self, ser, port=None):
        raw = ser.readline()
        if raw:
            self.serial_debug(port, "RX", raw)
        return raw

    def send_cmd(self, ser, cmd: str, port=None):
        self.serial_debug(port, "TX", cmd)
        ser.write((cmd + "\r\n").encode("utf-8"))

    def read_line(self, ser, port=None) -> str:
        line = ser.readline().decode("utf-8", errors="ignore").strip()
        if line:
            self.serial_debug(port, "RX", line)
        return line

    def query_key_value(self, ser, cmd: str, port=None, settle_s: float = 0.15, read_s: float = 2.2):
        deadline = time.time() + COMM_RETRY_TIMEOUT_S
        attempt = 0
        lines = []
        while time.time() < deadline and not self.shutdown_event.is_set():
            attempt += 1
            ser.reset_input_buffer()
            self.send_cmd(ser, cmd, port=port)
            time.sleep(settle_s)

            lines = []
            started = time.time()
            while time.time() - started <= read_s:
                if self.shutdown_event.is_set():
                    break
                line = self.read_line(ser, port=port)
                if not line:
                    continue
                lines.append(line)

            if lines:
                kv = {}
                for line in lines:
                    if "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    kv[key.strip()] = value.strip()
                return kv, lines

            remaining = max(0.0, deadline - time.time())
            if remaining <= 0:
                break
            port_text = f"[{port}] " if port else ""
            self.log(f"{port_text}No response to '{cmd}' (attempt {attempt}); retrying ({remaining:.0f}s left before timeout)...")
            time.sleep(min(COMM_RETRY_INTERVAL_S, remaining))

        raise TimeoutError(f"{cmd.upper()} communication timed out after {COMM_RETRY_TIMEOUT_S:.0f}s.")

    def clear_live_run_view(self, total_samples, port=None, serial_number=None):
        if threading.current_thread() is not threading.main_thread():
            self._ui_post("clear_live_run_view", total_samples, port=port, serial_number=serial_number)
            return
        if not port:
            return
        self.live_run_series_by_port[port] = {field: [] for field in self.live_plot_fields.values()}
        self.live_run_total_samples_by_port[port] = int(total_samples)
        self.live_run_serial_by_port[port] = serial_number or self.live_run_serial_by_port.get(port, "")
        self._ensure_live_port_color(port)
        self.update_live_std_label()
        self._update_live_samples_label()
        self.refresh_live_plot()

    def append_live_run_sample(self, sample_idx, sample, port=None):
        if threading.current_thread() is not threading.main_thread():
            self._ui_post("append_live_run_sample", sample_idx, sample, port=port)
            return
        if not port:
            return
        if port not in self.live_run_series_by_port:
            self.live_run_series_by_port[port] = {field: [] for field in self.live_plot_fields.values()}
        self._ensure_live_port_color(port)
        parsed = sample["parsed"]
        for field in self.live_run_series_by_port[port]:
            self.live_run_series_by_port[port][field].append(parsed.get(field, np.nan))
        self.update_live_std_label()

        serial_label = f"[{port}]" if port else "[NO-PORT]"
        preview_items = []
        for key in list(self.live_plot_fields.values())[:6]:
            preview_items.append(f"{key}={self.fmt(parsed.get(key, np.nan))}")
        preview = " ".join(preview_items) if preview_items else "(no live fields selected)"
        line = (
            f"{serial_label} sample {sample_idx}/{self.live_run_total_samples_by_port.get(port, 0)} "
            f"{preview} "
            f"raw={sample['raw']}\n"
        )
        self.live_text.configure(state=tk.NORMAL)
        self.live_text.insert(tk.END, line)
        self.live_text.see(tk.END)
        self.live_text.configure(state=tk.DISABLED)
        self._update_live_samples_label()
        self.refresh_live_plot()

    def _on_live_field_changed(self, _event=None):
        self.update_live_std_label()
        self.refresh_live_plot()

    def update_live_std_label(self):
        current_field = self.live_plot_fields.get(self.live_field_var.get(), next(iter(self.live_plot_fields.values())))
        parts = []
        for port in sorted(self.live_run_series_by_port.keys()):
            vals = np.array(
                [v for v in self.live_run_series_by_port[port].get(current_field, []) if np.isfinite(v)],
                dtype=float,
            )
            if len(vals) >= 2:
                std_val = float(np.std(vals, ddof=1))
            elif len(vals) == 1:
                std_val = 0.0
            else:
                continue
            parts.append(f"{port}={self.fmt(std_val)}")
        if parts:
            self.live_std_var.set(f"Std Dev ({self.live_field_var.get()}): " + " | ".join(parts))
        else:
            self.live_std_var.set(f"Std Dev ({self.live_field_var.get()}): n/a")

    def refresh_live_plot(self):
        if not hasattr(self, "live_canvas"):
            return
        c = self.live_canvas
        c.delete("all")
        width = max(int(c.winfo_width()), 240)
        height = max(int(c.winfo_height()), 160)

        left = 52
        right = width - 16
        top = 20
        bottom = height - 34
        c.create_rectangle(left, top, right, bottom, outline=DARK_BORDER)

        field_label = self.live_field_var.get()
        field = self.live_plot_fields.get(field_label, next(iter(self.live_plot_fields.values())))
        x_start = self._to_int_or_none(self.live_x_start_var.get())
        x_end_cfg = self._to_int_or_none(self.live_x_end_var.get())
        x_start = max(1, x_start if x_start is not None else 1)
        x_end_cfg = x_end_cfg if x_end_cfg is not None else 0
        series_by_port = {}
        for port, d in self.live_run_series_by_port.items():
            if self.live_visible_only_var.get() and self.live_visible_ports and port not in self.live_visible_ports:
                continue
            vals = np.array([v if np.isfinite(v) else np.nan for v in d.get(field, [])], dtype=float)
            if len(vals) >= x_start:
                vals = vals[x_start - 1 :]
            else:
                vals = np.array([], dtype=float)
            if x_end_cfg > 0:
                vals = vals[: max(0, x_end_cfg - x_start + 1)]
            finite_idx = np.where(np.isfinite(vals))[0]
            series_by_port[port] = (vals, finite_idx)
        scale_factor = self._field_scale_factor(field)
        c.create_text((left + right) // 2, 8, text=f"{field_label} vs Sample Count", fill=DARK_TEXT)
        c.create_text((left + right) // 2, height - 10, text="Sample Count", fill=DARK_TEXT)
        c.create_text(14, (top + bottom) // 2, text=self._field_label_with_unit(field, "Value"), angle=90, fill=DARK_TEXT)

        all_y = []
        max_n = 0
        for vals, finite_idx in series_by_port.values():
            if len(vals) > max_n:
                max_n = len(vals)
            if len(finite_idx) > 0:
                all_y.extend((vals[finite_idx] * scale_factor).tolist())
        if len(all_y) == 0:
            c.create_text((left + right) // 2, (top + bottom) // 2, text="Collecting samples...", fill=DARK_MUTED)
            return

        y_min = float(np.min(np.array(all_y, dtype=float)))
        y_max = float(np.max(np.array(all_y, dtype=float)))
        if self.live_autoscale_var.get():
            if y_min == y_max:
                pad = abs(y_min) * 0.01 if y_min != 0 else 0.01
                y_min -= pad
                y_max += pad
            else:
                pad = (y_max - y_min) * 0.08
                y_min -= pad
                y_max += pad
        else:
            y_min_cfg = self._to_float_or_none(self.live_ymin_var.get())
            y_max_cfg = self._to_float_or_none(self.live_ymax_var.get())
            if y_min_cfg is not None and y_max_cfg is not None and y_max_cfg > y_min_cfg:
                y_min, y_max = y_min_cfg, y_max_cfg

        n = max(max_n, 1)
        x_den = max(n - 1, 1)
        y_den = max(y_max - y_min, 1e-12)
        legend_items = []
        for port in sorted(series_by_port.keys()):
            vals, finite_idx = series_by_port[port]
            if len(finite_idx) == 0:
                continue
            color = self._ensure_live_port_color(port)
            points = []
            for idx in finite_idx:
                x = left + (idx / x_den) * (right - left)
                yv = vals[idx] * scale_factor
                y = bottom - ((yv - y_min) / y_den) * (bottom - top)
                points.extend((x, y))
            if len(points) >= 4:
                c.create_line(*points, fill=color, width=2.0, smooth=False)
            if self.live_show_points_var.get():
                for i in range(0, len(points), 2):
                    c.create_oval(points[i] - 2, points[i + 1] - 2, points[i] + 2, points[i + 1] + 2, fill=color, outline="")

            finite_vals = np.array([v for v in vals[finite_idx] if np.isfinite(v)], dtype=float)
            if len(finite_vals) >= 2:
                std_text = self.fmt(float(np.std(finite_vals, ddof=1)))
            elif len(finite_vals) == 1:
                std_text = self.fmt(0.0)
            else:
                std_text = "n/a"
            legend_items.append((port, color, std_text))

        c.create_text(left - 4, top, text=f"{y_max:.4f}", anchor="e", fill=DARK_MUTED)
        c.create_text(left - 4, bottom, text=f"{y_min:.4f}", anchor="e", fill=DARK_MUTED)
        x_end_label = x_start + n - 1
        c.create_text(left, bottom + 14, text=str(x_start), anchor="w", fill=DARK_MUTED)
        c.create_text(right, bottom + 14, text=str(x_end_label), anchor="e", fill=DARK_MUTED)

        if legend_items:
            legend_pad = 6
            row_h = 14
            legend_w = 190
            legend_h = legend_pad * 2 + row_h * len(legend_items)
            legend_x0 = right - legend_w - 4
            legend_y0 = top + 4
            c.create_rectangle(legend_x0, legend_y0, legend_x0 + legend_w, legend_y0 + legend_h, fill="#0f172a", outline=DARK_BORDER)
            for i, (port, color, std_text) in enumerate(legend_items):
                y = legend_y0 + legend_pad + i * row_h + 7
                c.create_line(legend_x0 + 8, y, legend_x0 + 22, y, fill=color, width=2)
                c.create_text(legend_x0 + 28, y, text=f"{port}  s={std_text}", anchor="w", fill=DARK_TEXT)

    def _ensure_live_port_color(self, port):
        if port not in self.live_port_colors:
            idx = len(self.live_port_colors) % len(LIVE_PORT_COLORS)
            self.live_port_colors[port] = LIVE_PORT_COLORS[idx]
        return self.live_port_colors[port]

    def select_visible_ports(self):
        ports = sorted(self.live_run_series_by_port.keys())
        if not ports:
            messagebox.showinfo("No Ports", "No live ports available yet.")
            return
        win = tk.Toplevel(self.root)
        win.title("Visible Live Ports")
        vars_by_port = {}
        default_selected = self.live_visible_ports or set(ports)
        for i, p in enumerate(ports):
            v = tk.BooleanVar(value=p in default_selected)
            ttk.Checkbutton(win, text=p, variable=v).grid(row=i, column=0, sticky="w", padx=8, pady=2)
            vars_by_port[p] = v

        def apply_and_close():
            self.live_visible_ports = {p for p, v in vars_by_port.items() if v.get()}
            self.refresh_live_plot()
            win.destroy()

        ttk.Button(win, text="Apply", command=apply_and_close).grid(row=len(ports), column=0, sticky="e", padx=8, pady=8)

    def _update_live_samples_label(self):
        if not self.live_run_total_samples_by_port:
            self.live_samples_var.set("Samples: 0 / 0")
            return
        parts = []
        for port in sorted(self.live_run_total_samples_by_port.keys()):
            total = int(self.live_run_total_samples_by_port.get(port, 0))
            port_series = self.live_run_series_by_port.get(port, {})
            have = max((len(v) for v in port_series.values()), default=0)
            parts.append(f"{port}:{have}/{total}")
        self.live_samples_var.set("Samples: " + " | ".join(parts))

    def _reset_live_view_for_ports(self, ports, total_samples):
        self.live_run_series_by_port = {}
        self.live_run_total_samples_by_port = {}
        self.live_run_serial_by_port = {}
        self.live_port_colors = {}
        for port in ports:
            self.live_run_series_by_port[port] = {field: [] for field in self.live_plot_fields.values()}
            self.live_run_total_samples_by_port[port] = int(total_samples)
            self._ensure_live_port_color(port)
        self.live_text.configure(state=tk.NORMAL)
        self.live_text.delete("1.0", tk.END)
        self.live_text.configure(state=tk.DISABLED)
        self.update_live_std_label()
        self._update_live_samples_label()
        self.refresh_live_plot()

    def take_sample(self, ser, port=None):
        deadline = time.time() + SAMPLE_RETRY_TIMEOUT_S
        attempt = 0
        while time.time() < deadline and not self.shutdown_event.is_set():
            attempt += 1
            ser.reset_input_buffer()
            sample_cmd = self.sample_command_var.get().strip() or "tsr"
            self.send_cmd(ser, sample_cmd, port=port)
            line1 = self.read_line(ser, port=port)
            line2 = self.read_line(ser, port=port)

            candidates = [line2, line1]
            raw = ""
            for c in candidates:
                if c and (self.delimiter_var.get() or ",") in c:
                    raw = c
                    break
            if raw:
                fields, parsed = self._parse_sample_payload(raw)
                return {
                    "raw": raw,
                    "fields": fields,
                    "parsed": parsed,
                }

            remaining = max(0.0, deadline - time.time())
            if remaining <= 0:
                break
            port_text = f"[{port}] " if port else ""
            self.log(f"{port_text}No TSR sample response (attempt {attempt}); retrying ({remaining:.0f}s left before timeout)...")
            time.sleep(min(SAMPLE_RETRY_INTERVAL_S, remaining))

        raise TimeoutError(f"{(self.sample_command_var.get().strip() or 'sample').upper()} sample communication timed out after {SAMPLE_RETRY_TIMEOUT_S:.0f}s.")

    def _parse_sample_payload(self, raw):
        content = str(raw).strip()
        pattern = self.parser_regex_var.get().strip()
        if pattern:
            try:
                m = re.search(pattern, content)
            except re.error:
                m = None
            if m:
                content = m.group(1) if m.groups() else m.group(0)
        trim_prefix = self.parser_trim_prefix_var.get()
        if trim_prefix and content.startswith(trim_prefix):
            content = content[len(trim_prefix) :].strip()
        delim = self.delimiter_var.get() or ","
        fields = [x.strip() for x in content.split(delim)]
        start_idx = max(0, int(self.parser_token_start_var.get()))
        if start_idx > 0:
            fields = fields[start_idx:]

        parsed = {}
        for d in self.sample_field_defs:
            idx = d["index"]
            name = d["key"]
            parsed[name] = self.to_float(fields[idx]) if idx < len(fields) else np.nan
        self._evaluate_derived_fields(parsed)
        return fields, parsed

    @staticmethod
    def to_float(v):
        try:
            return float(v)
        except Exception:
            return np.nan

    @staticmethod
    def chunked_std(vals, chunks=10):
        arr = np.array([x for x in vals if np.isfinite(x)], dtype=float)
        if len(arr) < chunks:
            return np.nan
        parts = np.array_split(arr, chunks)
        return float(np.mean([np.std(p) for p in parts if len(p) > 0]))

    @staticmethod
    def _has_stuck_run(values, n_repeat):
        if not n_repeat or n_repeat < 2:
            return False
        run = 1
        prev = None
        for v in values:
            if not np.isfinite(v):
                run = 1
                prev = None
                continue
            if prev is not None and abs(v - prev) < 1e-12:
                run += 1
            else:
                run = 1
            prev = v
            if run >= n_repeat:
                return True
        return False

    def classify(self, red_ns, blue_ns):
        if not np.isfinite(red_ns) or not np.isfinite(blue_ns):
            return "UNKNOWN"
        if red_ns > FAIL_NS or blue_ns > FAIL_NS:
            return "FAIL"
        if red_ns > WARN_NS or blue_ns > WARN_NS:
            return "WARN"
        return "PASS"

    def parse_caldate(self, caldate_days):
        if caldate_days is None:
            return None
        try:
            days = int(str(caldate_days).strip())
            return dt.datetime(2000, 1, 1) + dt.timedelta(days=days)
        except Exception:
            return None

    @staticmethod
    def normalize_serial_key(key: str) -> str:
        return re.sub(r"[^a-z0-9#]", "", str(key).strip().lower())

    @staticmethod
    def parse_serial_value(value: str):
        m = re.search(r"(?<![0-9])([0-9]{3,6})(?![0-9])", str(value))
        return m.group(1) if m else None

    def extract_serial_number(self, ds, ds_lines):
        preferred_keys = {"serial#", "serial", "serialnumber"}
        blocked_keys = {"sensorfilmserial#"}

        for line in ds_lines:
            if "=" not in line:
                continue
            key, raw = line.split("=", 1)
            norm_key = self.normalize_serial_key(key)
            if norm_key in blocked_keys or norm_key not in preferred_keys:
                continue
            serial = self.parse_serial_value(raw)
            if serial:
                return serial

        for key, raw in ds.items():
            norm_key = self.normalize_serial_key(key)
            if norm_key in blocked_keys or norm_key not in preferred_keys:
                continue
            serial = self.parse_serial_value(raw)
            if serial:
                return serial
        return None

    def build_unit_folder(self, serial_number: str):
        folder = os.path.join(self._current_results_root(), serial_number, PRECAL_TEST_SUBDIR)
        os.makedirs(folder, exist_ok=True)
        return folder

    def unique_path(self, path):
        if not os.path.exists(path):
            return path
        root, ext = os.path.splitext(path)
        i = 1
        while True:
            candidate = f"{root}_{i}{ext}"
            if not os.path.exists(candidate):
                return candidate
            i += 1

    def append_session_row(self, row):
        is_new = not os.path.exists(self.session_csv)
        with open(self.session_csv, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(row.keys()))
            if is_new:
                writer.writeheader()
            writer.writerow(row)

    def collect_samples(self, ser, n_samples, port=None):
        samples = []
        for i in range(1, n_samples + 1):
            if self.shutdown_event.is_set():
                raise RuntimeError("Shutdown requested.")
            s = self.take_sample(ser, port=port)
            s["idx"] = i
            s["captured_at"] = dt.datetime.now().isoformat(timespec="milliseconds")
            samples.append(s)
            self.append_live_run_sample(i, s, port=port)
            if i % 10 == 0 or i == n_samples:
                self.log(f"Collected sample {i}/{n_samples}")
        return samples

    def compute_metrics(self, samples):
        value_map = {}
        for d in self.sample_field_defs:
            key = d["key"]
            value_map[key] = [s["parsed"].get(key, np.nan) for s in samples]

        red_phase = value_map.get("red_phase", [])
        blue_phase = value_map.get("blue_phase", [])
        red_blue_phase = value_map.get("red_blue_phase", [])
        red_voltage = value_map.get("red_voltage", [])
        blue_voltage = value_map.get("blue_voltage", [])
        red_pll_voltage = value_map.get("red_pll_voltage", [])
        blue_pll_voltage = value_map.get("blue_pll_voltage", [])
        raw_temp_voltage = value_map.get("raw_temp_voltage", [])
        elec_temp_voltage = value_map.get("electronics_temp_voltage", [])

        red_noise_ns = self.chunked_std(red_phase) * 1e3 if red_phase else np.nan
        blue_noise_ns = self.chunked_std(blue_phase) * 1e3 if blue_phase else np.nan
        red_blue_noise_ns = self.chunked_std(red_blue_phase) * 1e3 if red_blue_phase else np.nan

        red_v_std = float(np.nanstd(np.array(red_voltage, dtype=float))) if red_voltage else np.nan
        red_v_avg = float(np.nanmean(np.array(red_voltage, dtype=float))) if red_voltage else np.nan
        blue_v_std = float(np.nanstd(np.array(blue_voltage, dtype=float))) if blue_voltage else np.nan
        blue_v_avg = float(np.nanmean(np.array(blue_voltage, dtype=float))) if blue_voltage else np.nan
        red_pll_v_std = float(np.nanstd(np.array(red_pll_voltage, dtype=float))) if red_pll_voltage else np.nan
        red_pll_v_avg = float(np.nanmean(np.array(red_pll_voltage, dtype=float))) if red_pll_voltage else np.nan
        blue_pll_v_std = float(np.nanstd(np.array(blue_pll_voltage, dtype=float))) if blue_pll_voltage else np.nan
        blue_pll_v_avg = float(np.nanmean(np.array(blue_pll_voltage, dtype=float))) if blue_pll_voltage else np.nan
        raw_temp_v_std = float(np.nanstd(np.array(raw_temp_voltage, dtype=float))) if raw_temp_voltage else np.nan
        raw_temp_v_avg = float(np.nanmean(np.array(raw_temp_voltage, dtype=float))) if raw_temp_voltage else np.nan
        elec_temp_v_std = float(np.nanstd(np.array(elec_temp_voltage, dtype=float))) if elec_temp_voltage else np.nan
        elec_temp_v_avg = float(np.nanmean(np.array(elec_temp_voltage, dtype=float))) if elec_temp_voltage else np.nan

        flags = []
        if np.isfinite(red_noise_ns) and red_noise_ns == 0.0 and np.isfinite(red_v_std) and red_v_std > 0.0:
            flags.append("red_phase_flat_with_voltage_activity")
        if np.isfinite(blue_noise_ns) and blue_noise_ns == 0.0 and np.isfinite(blue_v_std) and blue_v_std > 0.0:
            flags.append("blue_phase_flat_with_voltage_activity")
        if np.isfinite(red_noise_ns) and np.isfinite(blue_noise_ns) and red_noise_ns == 0.0 and blue_noise_ns == 0.0:
            flags.append("both_phase_noise_zero")
        if np.isfinite(red_blue_noise_ns) and red_blue_noise_ns > 0 and "both_phase_noise_zero" in flags:
            flags.append("red_blue_phase_shows_activity")
        for d in self.sample_field_defs:
            key = d["key"]
            vals = np.array(value_map.get(key, []), dtype=float)
            meta = self.field_meta_by_key.get(key, {})
            min_v = meta.get("min_val")
            max_v = meta.get("max_val")
            if min_v is not None and np.any(vals[np.isfinite(vals)] < min_v):
                flags.append(f"{key}_below_min")
            if max_v is not None and np.any(vals[np.isfinite(vals)] > max_v):
                flags.append(f"{key}_above_max")
            stuck_n = meta.get("stuck_n")
            if stuck_n and self._has_stuck_run(vals, stuck_n):
                flags.append(f"{key}_stuck_{stuck_n}")

        metrics = {
            "red_noise_ns": red_noise_ns,
            "blue_noise_ns": blue_noise_ns,
            "red_blue_noise_ns": red_blue_noise_ns,
            "red_voltage_std": red_v_std,
            "red_voltage_avg": red_v_avg,
            "blue_voltage_std": blue_v_std,
            "blue_voltage_avg": blue_v_avg,
            "red_pll_voltage_std": red_pll_v_std,
            "red_pll_voltage_avg": red_pll_v_avg,
            "blue_pll_voltage_std": blue_pll_v_std,
            "blue_pll_voltage_avg": blue_pll_v_avg,
            "raw_temp_voltage_std": raw_temp_v_std,
            "raw_temp_voltage_avg": raw_temp_v_avg,
            "electronics_temp_voltage_std": elec_temp_v_std,
            "electronics_temp_voltage_avg": elec_temp_v_avg,
            "severity": self.classify(red_noise_ns, blue_noise_ns),
            "flags": ";".join(flags) if flags else "",
        }
        for key, values in value_map.items():
            arr = np.array(values, dtype=float)
            metrics[f"{key}_std"] = float(np.nanstd(arr)) if len(arr) else np.nan
            metrics[f"{key}_avg"] = float(np.nanmean(arr)) if len(arr) else np.nan
        return metrics

    def write_sample_csv(self, path, samples, serial_number):
        sample_fields = [d["key"] for d in self.sample_field_defs]
        fieldnames = ["sample_idx", "captured_at", "serial", "raw_sample"] + sample_fields
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for s in samples:
                row = {
                    "sample_idx": s["idx"],
                    "captured_at": s["captured_at"],
                    "serial": serial_number,
                    "raw_sample": s["raw"],
                }
                for name in sample_fields:
                    row[name] = s["parsed"].get(name, np.nan)
                writer.writerow(row)

    def run_unit_test(self):
        if self.run_in_progress:
            messagebox.showwarning("Run In Progress", "A unit test is already running. Wait for completion.")
            return
        connected_ports = sorted([p for p, s in self.serial_pool.items() if s and s.is_open])
        if not connected_ports:
            messagebox.showwarning("Not Connected", "Connect at least one COM port first.")
            return

        missing = self.missing_setup_fields()
        if missing:
            messagebox.showerror("Missing Setup Fields", "Fill required setup fields before starting:\n- " + "\n- ".join(missing))
            return

        if len(self.session_serials) >= MAX_UNITS_PER_SESSION:
            messagebox.showwarning("Session Limit", f"This session already has {MAX_UNITS_PER_SESSION} unique units.")
            return

        try:
            n_samples = int(self.sample_count_var.get())
        except Exception:
            messagebox.showerror("Input Error", "Samples must be an integer.")
            return

        if n_samples < 20:
            messagebox.showerror("Input Error", "Use at least 20 samples for stable noise stats.")
            return
        try:
            run_count = int(self.batch_run_count_var.get())
        except Exception:
            messagebox.showerror("Input Error", "Runs must be an integer.")
            return
        if run_count < 1:
            messagebox.showerror("Input Error", "Runs must be at least 1.")
            return
        try:
            delay_s = float(self.batch_delay_s_var.get())
        except Exception:
            messagebox.showerror("Input Error", "Delay (s) must be numeric.")
            return
        if delay_s < 0:
            messagebox.showerror("Input Error", "Delay (s) cannot be negative.")
            return

        setup = {
            "operator": self.operator_var.get().strip(),
            "notes": self.notes_var.get().strip(),
            "bath_id": self.bath_id_var.get().strip(),
            "bath_temp_c": self.bath_temp_c_var.get().strip(),
            "salinity_psu": self.salinity_psu_var.get().strip(),
        }

        self.run_in_progress = True
        self.update_run_button_state()
        with self.run_state_lock:
            self.active_run_ports = set(connected_ports)
            self.run_threads = {}
        self.batch_runs_remaining_by_port = {port: run_count for port in connected_ports}
        self._update_runs_left_label()
        self._reset_live_view_for_ports(connected_ports, n_samples)
        self.log(
            f"Starting parallel test across {len(connected_ports)} port(s): {', '.join(connected_ports)} "
            f"| runs={run_count}, samples={n_samples}, delay={delay_s:.1f}s"
        )
        for port in connected_ports:
            t = threading.Thread(
                target=self._run_unit_test_worker,
                args=(port, n_samples, run_count, delay_s, setup),
                daemon=True,
            )
            with self.run_state_lock:
                self.run_threads[port] = t
            t.start()

    def _run_unit_test_worker(self, selected_port, n_samples, run_count, delay_s, setup):
        ser = self.serial_pool.get(selected_port)
        if not ser or not ser.is_open:
            self.log(f"[{selected_port}] Run failed: selected COM port is not connected.")
            self.set_port_status(selected_port, "ERROR")
            self._ui_post("show_error", "Run Failed", f"{selected_port} is not connected.")
            self._ui_post("finish_port_run", selected_port)
            return

        try:
            self.log(f"[{selected_port}] Starting batch: {run_count} run(s), {n_samples} samples each, delay={delay_s:.1f}s")
            for run_idx in range(1, run_count + 1):
                if self.shutdown_event.is_set():
                    return
                ser = self.serial_pool.get(selected_port)
                if not ser or not ser.is_open:
                    raise RuntimeError(f"{selected_port} disconnected during batch run.")

                self.set_port_status(selected_port, "RUNNING")
                if run_count > 1:
                    self.log(f"[{selected_port}] Batch run {run_idx}/{run_count} starting...")
                self.log(f"[{selected_port}] Reading DS...")
                ds, ds_lines = self.query_key_value(ser, "ds", port=selected_port)
                self.log(f"[{selected_port}] Reading DC...")
                dc, dc_lines = self.query_key_value(ser, "dc", port=selected_port)

                serial_number = self.extract_serial_number(ds, ds_lines)
                if not serial_number:
                    serial_number = "UNKNOWN"
                self.set_port_status(selected_port, "RUNNING", serial=serial_number)

                if serial_number not in self.session_serials and len(self.session_serials) >= MAX_UNITS_PER_SESSION:
                    self._ui_post("show_warning", "Session Limit", f"This session already has {MAX_UNITS_PER_SESSION} unique units.")
                    self.log(f"[{selected_port}] Stopping batch due to session unique-unit limit.")
                    return

                self.log(f"[{selected_port}] Unit serial: {serial_number}")

                caldate = self.parse_caldate(dc.get("Caldate"))
                cal_age_days = (dt.datetime.now() - caldate).days if caldate else None

                self.log(f"[{selected_port}] Collecting {n_samples} samples from TSR stream...")
                self.clear_live_run_view(n_samples, port=selected_port, serial_number=serial_number)
                samples = self.collect_samples(ser, n_samples, port=selected_port)
                metrics = self.compute_metrics(samples)

                run_ts = dt.datetime.now().replace(microsecond=0)
                run_stamp = run_ts.isoformat().replace(":", "_")
                unit_dir = self.build_unit_folder(serial_number)

                sample_csv = self.unique_path(os.path.join(unit_dir, f"SBS83_SN{serial_number}_{run_stamp}_samples.csv"))
                unit_log = self.unique_path(os.path.join(unit_dir, f"SBS83_SN{serial_number}_{run_stamp}.log"))
                unit_json = self.unique_path(os.path.join(unit_dir, f"SBS83_SN{serial_number}_{run_stamp}_summary.json"))

                self.write_sample_csv(sample_csv, samples, serial_number)

                with open(unit_log, "w", encoding="utf-8") as f:
                    f.write(f"PORT: {selected_port}\n")
                    f.write("DS:\n")
                    for line in ds_lines:
                        f.write(line + "\n")
                    f.write("\nDC:\n")
                    for line in dc_lines:
                        f.write(line + "\n")
                    f.write("\nSetup:\n")
                    f.write(json.dumps(setup, indent=2) + "\n")
                    f.write("\nMetrics:\n")
                    f.write(f"Red noise (ns): {metrics['red_noise_ns']}\n")
                    f.write(f"Blue noise (ns): {metrics['blue_noise_ns']}\n")
                    f.write(f"Red-Blue noise (ns): {metrics['red_blue_noise_ns']}\n")
                    f.write(f"Red voltage std: {metrics['red_voltage_std']}\n")
                    f.write(f"Red voltage avg: {metrics['red_voltage_avg']}\n")
                    f.write(f"Blue voltage std: {metrics['blue_voltage_std']}\n")
                    f.write(f"Blue voltage avg: {metrics['blue_voltage_avg']}\n")
                    f.write(f"Red PLL voltage std: {metrics['red_pll_voltage_std']}\n")
                    f.write(f"Red PLL voltage avg: {metrics['red_pll_voltage_avg']}\n")
                    f.write(f"Blue PLL voltage std: {metrics['blue_pll_voltage_std']}\n")
                    f.write(f"Blue PLL voltage avg: {metrics['blue_pll_voltage_avg']}\n")
                    f.write(f"Raw temp voltage std: {metrics['raw_temp_voltage_std']}\n")
                    f.write(f"Raw temp voltage avg: {metrics['raw_temp_voltage_avg']}\n")
                    f.write(f"Electronics temp voltage std: {metrics['electronics_temp_voltage_std']}\n")
                    f.write(f"Electronics temp voltage avg: {metrics['electronics_temp_voltage_avg']}\n")
                    f.write(f"Flags: {metrics['flags']}\n")
                    f.write(f"Sample CSV: {sample_csv}\n")

                summary = {
                    "timestamp": run_ts.isoformat(timespec="seconds"),
                    "session_id": self.session_id,
                    "port": selected_port,
                    "serial": serial_number,
                    "operator": setup["operator"],
                    "notes": setup["notes"],
                    "bath_id": setup["bath_id"],
                    "bath_temp_c": setup["bath_temp_c"],
                    "salinity_psu": setup["salinity_psu"],
                    "sample_count": n_samples,
                    "run_index": run_idx,
                    "run_total": run_count,
                    "caldate": caldate.isoformat(timespec="seconds") if caldate else "",
                    "cal_age_days": cal_age_days if cal_age_days is not None else "",
                    "red_noise_ns": metrics["red_noise_ns"],
                    "blue_noise_ns": metrics["blue_noise_ns"],
                    "red_blue_noise_ns": metrics["red_blue_noise_ns"],
                    "red_voltage_std": metrics["red_voltage_std"],
                    "red_voltage_avg": metrics["red_voltage_avg"],
                    "blue_voltage_std": metrics["blue_voltage_std"],
                    "blue_voltage_avg": metrics["blue_voltage_avg"],
                    "red_pll_voltage_std": metrics["red_pll_voltage_std"],
                    "red_pll_voltage_avg": metrics["red_pll_voltage_avg"],
                    "blue_pll_voltage_std": metrics["blue_pll_voltage_std"],
                    "blue_pll_voltage_avg": metrics["blue_pll_voltage_avg"],
                    "raw_temp_voltage_std": metrics["raw_temp_voltage_std"],
                    "raw_temp_voltage_avg": metrics["raw_temp_voltage_avg"],
                    "electronics_temp_voltage_std": metrics["electronics_temp_voltage_std"],
                    "electronics_temp_voltage_avg": metrics["electronics_temp_voltage_avg"],
                    "severity": metrics["severity"],
                    "flags": metrics["flags"],
                    "sample_csv": sample_csv,
                    "unit_log": unit_log,
                    "unit_json": unit_json,
                }
                for d in self.sample_field_defs:
                    key = d["key"]
                    summary[f"{key}_std"] = metrics.get(f"{key}_std", np.nan)
                    summary[f"{key}_avg"] = metrics.get(f"{key}_avg", np.nan)

                with open(unit_json, "w", encoding="utf-8") as f:
                    json.dump(summary, f, indent=2)

                self._ui_post("run_result", summary, metrics, selected_port, serial_number, sample_csv)

                if run_idx < run_count and delay_s > 0:
                    self.log(f"[{selected_port}] Waiting {delay_s:.1f}s before run {run_idx + 1}/{run_count}...")
                    time.sleep(delay_s)

        except Exception as exc:
            self.log(f"[{selected_port}] Run failed: {exc}")
            self.set_port_status(selected_port, "ERROR")
            self._ui_post("show_error", "Run Failed", str(exc))
        finally:
            self._ui_post("finish_port_run", selected_port)

    def _apply_run_result(self, summary, metrics, selected_port, serial_number, sample_csv):
        self.append_session_row(summary)
        self.session_rows.append(summary)
        self.session_serials.add(serial_number)
        self.limit_var.set(f"Units tested: {len(self.session_serials)} / {MAX_UNITS_PER_SESSION}")
        run_total = int(summary.get("run_total", 1) or 1)
        run_index = int(summary.get("run_index", 1) or 1)
        self.batch_runs_remaining_by_port[selected_port] = max(0, run_total - run_index)
        self._update_runs_left_label()

        self.tree.insert(
            "",
            tk.END,
            values=(
                summary["timestamp"],
                selected_port,
                serial_number,
                self.fmt(metrics["red_noise_ns"]),
                self.fmt(metrics["red_voltage_std"]),
                self.fmt(metrics["red_voltage_avg"]),
                self.fmt(metrics["blue_noise_ns"]),
                self.fmt(metrics["blue_voltage_std"]),
                self.fmt(metrics["blue_voltage_avg"]),
                self.fmt(metrics["red_pll_voltage_std"]),
                self.fmt(metrics["red_pll_voltage_avg"]),
                self.fmt(metrics["blue_pll_voltage_std"]),
                self.fmt(metrics["blue_pll_voltage_avg"]),
                self.fmt(metrics["raw_temp_voltage_std"]),
                self.fmt(metrics["raw_temp_voltage_avg"]),
                self.fmt(metrics["electronics_temp_voltage_std"]),
                self.fmt(metrics["electronics_temp_voltage_avg"]),
                metrics["flags"],
                sample_csv,
            ),
        )

        self.log(
            f"[{selected_port}] Complete SN{serial_number}: red={self.fmt(metrics['red_noise_ns'])} ns, "
            f"blue={self.fmt(metrics['blue_noise_ns'])} ns"
        )
        if metrics["flags"]:
            self.log(f"[{selected_port}] Flags: {metrics['flags']}")
        self.log(f"[{selected_port}] Raw sample capture saved: {sample_csv}")
        self.set_port_status(selected_port, metrics["severity"], serial=serial_number)

    def _finish_port_run(self, port):
        idx = self.find_slot_by_port(port)
        if idx is not None:
            state = self.port_slots[idx]["state_var"].get()
            if state in {"PASS", "WARN", "FAIL"}:
                self.set_port_status(port, "COMPLETE")

        with self.run_state_lock:
            self.active_run_ports.discard(port)
            self.run_threads.pop(port, None)
            done = len(self.active_run_ports) == 0
        self.batch_runs_remaining_by_port[port] = 0
        if done:
            self.batch_runs_remaining_by_port = {}
        self._update_runs_left_label()
        if done:
            self.run_in_progress = False
            self.update_run_button_state()
            self.log("Parallel run complete on all connected ports.")

    @staticmethod
    def fmt(v):
        return "nan" if not np.isfinite(v) else f"{v:.3g}"

    def save_session_json(self):
        if not self.session_rows:
            messagebox.showinfo("No Data", "No unit results in this session yet.")
            return

        out_path = os.path.join(self.session_dir, f"sbe83_session_{self.session_id}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(self.session_rows, f, indent=2)
        self.log(f"Wrote session JSON: {out_path}")
        messagebox.showinfo("Saved", f"Session JSON saved:\n{out_path}")

    def reset_session(self):
        if self.run_in_progress:
            messagebox.showwarning("Run In Progress", "Wait for the active test run to finish before resetting session.")
            return
        if self.session_rows and not messagebox.askyesno(
            "Reset Session",
            "Start a new session and clear current table?",
        ):
            return

        for row in self.tree.get_children():
            self.tree.delete(row)

        self.session_start = dt.datetime.now()
        self.session_id = self.session_start.strftime("%Y%m%d_%H%M%S")
        self.session_rows = []
        self.session_serials = set()
        self.session_csv = os.path.join(self.session_dir, f"sbe83_session_{self.session_id}.csv")
        self.limit_var.set(f"Units tested: 0 / {MAX_UNITS_PER_SESSION}")
        self.batch_runs_remaining_by_port = {}
        self._update_runs_left_label()
        self.log(f"New session started: {self.session_id}")
        self.log(f"Session summary file: {self.session_csv}")
        self.update_port_grid()

    def shutdown(self):
        self.shutdown_event.set()
        if self._layout_save_after_id is not None:
            try:
                self.root.after_cancel(self._layout_save_after_id)
            except Exception:
                pass
            self._layout_save_after_id = None

        for ev in self.stream_stop_events.values():
            ev.set()

        run_threads = []
        with self.run_state_lock:
            run_threads = list(self.run_threads.values())

        stream_threads = list(self.stream_threads.values())
        for ser in list(self.serial_pool.values()):
            try:
                ser.close()
            except Exception:
                pass
        self.serial_pool.clear()

        for t in run_threads:
            if t and t.is_alive():
                t.join(timeout=0.6)
        for t in stream_threads:
            if t and t.is_alive():
                t.join(timeout=0.3)

        with self.run_state_lock:
            self.active_run_ports.clear()
            self.run_threads = {}
        self.run_in_progress = False
        try:
            self._save_app_config()
        except Exception:
            pass


def main():
    global SENSOR_TEST_DIR
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        SENSOR_TEST_DIR = os.path.join(exe_dir, "SBE83")
    try:
        os.makedirs(SENSOR_TEST_DIR, exist_ok=True)
    except Exception:
        fallback_root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sessions_local", "SBE83")
        os.makedirs(fallback_root, exist_ok=True)
        SENSOR_TEST_DIR = fallback_root
    root = tk.Tk()
    app = SBE83GuiApp(root)

    def on_close():
        app.shutdown()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
