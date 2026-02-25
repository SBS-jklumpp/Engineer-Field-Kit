import csv
import datetime as dt
import json
import os
import queue
import re
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import numpy as np
import serial
from serial.tools import list_ports
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

APP_NAME = "Engineer’s Field Kit – Multitool"
APP_SUBTITLE = "Engineering Multitool"
APP_TAGLINE = "Serial • Plot • Analyze • Debug"
APP_VERSION = "v1.0.1"
APP_AUTHOR = "Senior Electrical Engineer"
APP_CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "engineers_field_kit_multitool_config.json")


class SBE83GuiApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1280x800")
        self.app_config = self._load_app_config()
        self.dark_mode_var = tk.BooleanVar(value=bool(self.app_config.get("dark_mode", True)))
        self.port_station_collapsed_var = tk.BooleanVar(value=bool(self.app_config.get("port_station_collapsed", False)))
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
        self.session_dir = os.path.join(SENSOR_TEST_DIR, "sessions", PRECAL_TEST_SUBDIR)
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
        self.live_plot_fields = dict(LIVE_PLOT_FIELDS)
        self.base_session_plot_fields = dict(SESSION_PLOT_FIELDS)
        self.session_plot_fields = dict(SESSION_PLOT_FIELDS)
        self.sample_field_defs = self._default_sample_field_defs()
        self.field_meta_by_key = {}
        self.derived_fields = []
        self.live_visible_ports = set()
        self.stream_threads = {}
        self.stream_stop_events = {}
        self.stream_enabled = {}
        self.manual_capture_rows = []
        self.baseline_rows = []
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
        self.tabs_detached = False
        self.sample_format_expanded = False

        self._build_ui()
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

    def _save_app_config(self):
        data = {
            "dark_mode": bool(self.dark_mode_var.get()),
            "port_station_collapsed": bool(self.port_station_collapsed_var.get()),
        }
        try:
            with open(APP_CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as exc:
            self.log(f"Config save failed: {exc}")

    def _console_font(self):
        return ("Consolas", 10)

    def _theme_colors(self):
        if self.dark_mode_var.get():
            return {
                "bg": DARK_BG,
                "panel": DARK_PANEL,
                "panel_2": DARK_PANEL_2,
                "fg": DARK_TEXT,
                "muted": DARK_MUTED,
                "ok": DARK_OK,
                "canvas": "#0b1220",
                "entry": DARK_PANEL,
            }
        return {
            "bg": LIGHT_BG,
            "panel": "#ffffff",
            "panel_2": LIGHT_PANEL_2,
            "fg": LIGHT_TEXT,
            "muted": LIGHT_MUTED,
            "ok": LIGHT_OK,
            "canvas": "#f8fafc",
            "entry": "#ffffff",
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

    def _apply_port_station_visibility(self, persist=True):
        if not hasattr(self, "port_station_frame"):
            return
        if self.port_station_collapsed_var.get():
            self.port_station_frame.pack_forget()
            if hasattr(self, "port_station_toggle_btn"):
                self.port_station_toggle_btn.configure(text="Show Port Station View")
        else:
            self.port_station_frame.pack(fill=tk.X, pady=2, before=self.setup_frame)
            if hasattr(self, "port_station_toggle_btn"):
                self.port_station_toggle_btn.configure(text="Hide Port Station View")
        if persist:
            self._save_app_config()

    def _toggle_port_station_visibility(self):
        self.port_station_collapsed_var.set(not self.port_station_collapsed_var.get())
        self._apply_port_station_visibility(persist=True)

    def _build_ui(self):
        self.root.minsize(860, 560)
        hero = tk.Frame(self.root, bg=DARK_PANEL_2, highlightthickness=1, highlightbackground=DARK_BORDER)
        hero.pack(fill=tk.X, padx=8, pady=(8, 0))
        tk.Label(
            hero,
            text=APP_NAME,
            bg=DARK_PANEL_2,
            fg=DARK_TEXT,
            font=("Segoe UI Semibold", 13),
            padx=10,
            pady=6,
        ).pack(side=tk.LEFT)
        hero_right = tk.Frame(hero, bg=DARK_PANEL_2)
        hero_right.pack(side=tk.RIGHT, padx=10)
        tk.Label(hero_right, text=APP_VERSION, bg=DARK_PANEL_2, fg=DARK_ACCENT, font=("Segoe UI Semibold", 10)).pack(anchor="e")
        tk.Label(hero_right, text=APP_TAGLINE, bg=DARK_PANEL_2, fg=DARK_MUTED, font=("Segoe UI", 9)).pack(anchor="e")

        self.top_frame = ttk.Frame(self.root, padding=8)
        self.top_frame.pack(fill=tk.X)

        conn = ttk.LabelFrame(self.top_frame, text="Connection (up to 10 COM ports)", padding=8)
        conn.pack(fill=tk.X, pady=4)

        ttk.Label(conn, text="Selected COM Port").grid(row=0, column=0, sticky="w")
        self.com_var = tk.StringVar(value="COM5")
        self.com_combo = ttk.Combobox(conn, textvariable=self.com_var, width=14, state="readonly")
        self.com_combo.grid(row=0, column=1, padx=6, sticky="w")
        ttk.Label(conn, text="Baud").grid(row=0, column=2, sticky="e")
        self.baud_combo = ttk.Combobox(conn, textvariable=self.baudrate_var, width=10, state="readonly", values=BAUD_OPTIONS)
        self.baud_combo.grid(row=0, column=3, padx=6, sticky="w")

        ttk.Button(conn, text="Refresh Ports", command=self.refresh_ports).grid(row=0, column=4, padx=4)
        ttk.Button(conn, text="Connect Selected", command=self.connect_selected_port).grid(row=0, column=5, padx=4)
        ttk.Button(conn, text="Reconnect @ Baud", command=self.reconnect_selected_port).grid(row=0, column=6, padx=4)
        ttk.Button(conn, text="Disconnect Selected", command=self.disconnect_selected_port).grid(row=0, column=7, padx=4)
        ttk.Button(conn, text="Connect All", command=self.connect_all_ports).grid(row=0, column=8, padx=4)
        ttk.Button(conn, text="Disconnect All", command=self.disconnect_all_ports).grid(row=0, column=9, padx=4)
        self.port_station_toggle_btn = ttk.Button(conn, text="Hide Port Station View", command=self._toggle_port_station_visibility)
        self.port_station_toggle_btn.grid(row=0, column=10, padx=4)

        self.conn_status = tk.StringVar(value="Connected ports: 0")
        ttk.Label(conn, textvariable=self.conn_status, foreground=DARK_MUTED).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.connected_ports_var = tk.StringVar(value="None")
        ttk.Label(conn, text="Connected list:").grid(row=1, column=5, sticky="e", pady=(6, 0))
        ttk.Label(conn, textvariable=self.connected_ports_var, foreground=DARK_OK).grid(
            row=1, column=6, columnspan=4, sticky="w", pady=(6, 0)
        )

        self.port_station_frame = ttk.LabelFrame(self.top_frame, text="Port Station View (10 slots)", padding=4)
        self.port_station_frame.pack(fill=tk.X, pady=2)
        self._build_port_grid(self.port_station_frame)

        self.setup_frame = ttk.LabelFrame(self.top_frame, text="Test Setup (Salt Bath Controlled Environment)", padding=8)
        self.setup_frame.pack(fill=tk.X, pady=4)

        self.operator_var = tk.StringVar(value=DEFAULT_OPERATOR)
        self.station_var = tk.StringVar()
        self.bath_temp_c_var = tk.StringVar()
        self.salinity_psu_var = tk.StringVar(value=DEFAULT_SALINITY_PSU)
        self.bath_id_var = tk.StringVar()
        self.sample_count_var = tk.IntVar(value=50)
        self.batch_run_count_var = tk.IntVar(value=1)
        self.batch_delay_s_var = tk.DoubleVar(value=5.0)

        ttk.Label(self.setup_frame, text="Operator").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.setup_frame, textvariable=self.operator_var, width=20, state="readonly").grid(row=0, column=1, padx=5, sticky="w")
        ttk.Label(self.setup_frame, text="Station").grid(row=0, column=2, sticky="w")
        ttk.Entry(self.setup_frame, textvariable=self.station_var, width=20).grid(row=0, column=3, padx=5, sticky="w")
        ttk.Label(self.setup_frame, text="Bath ID").grid(row=0, column=4, sticky="w")
        ttk.Entry(self.setup_frame, textvariable=self.bath_id_var, width=18).grid(row=0, column=5, padx=5, sticky="w")

        ttk.Label(self.setup_frame, text="Bath Temp (C)").grid(row=1, column=0, sticky="w")
        ttk.Entry(self.setup_frame, textvariable=self.bath_temp_c_var, width=20).grid(row=1, column=1, padx=5, sticky="w")
        ttk.Label(self.setup_frame, text="Salinity (PSU)").grid(row=1, column=2, sticky="w")
        ttk.Entry(self.setup_frame, textvariable=self.salinity_psu_var, width=20, state="readonly").grid(row=1, column=3, padx=5, sticky="w")
        ttk.Label(self.setup_frame, text="Samples").grid(row=1, column=4, sticky="w")
        ttk.Spinbox(self.setup_frame, from_=20, to=500, textvariable=self.sample_count_var, width=8).grid(row=1, column=5, padx=5, sticky="w")
        for var in (
            self.operator_var,
            self.station_var,
            self.bath_id_var,
            self.bath_temp_c_var,
            self.salinity_psu_var,
        ):
            var.trace_add("write", self._on_setup_field_changed)
        self._apply_port_station_visibility(persist=False)

        actions = ttk.LabelFrame(self.top_frame, text="Actions", padding=8)
        actions.pack(fill=tk.X, pady=4)

        self.run_btn = ttk.Button(
            actions, text="Run Unit Test (all connected)", command=self.run_unit_test, state=tk.DISABLED, style="Primary.TButton"
        )
        self.run_btn.grid(row=0, column=0, padx=4)
        ttk.Button(actions, text="Save Session JSON", command=self.save_session_json).grid(row=0, column=1, padx=4)
        ttk.Button(actions, text="Reset Session", command=self.reset_session).grid(row=0, column=2, padx=4)
        ttk.Button(actions, text="Show Plot", command=self.show_plot_tab).grid(row=0, column=3, padx=4)
        ttk.Button(actions, text="Show Console", command=self.show_console_tab).grid(row=0, column=4, padx=4)
        self.detach_tabs_btn = ttk.Button(actions, text="Detach Tabs", command=self.detach_tabs_window)
        self.detach_tabs_btn.grid(row=0, column=5, padx=4)
        ttk.Button(actions, text="Plot Session", command=self.plot_current_session).grid(row=0, column=6, padx=4)
        ttk.Button(actions, text="Load Session Plot", command=self.load_session_plot).grid(row=0, column=7, padx=4)
        ttk.Button(actions, text="About", command=self.show_about_dialog).grid(row=0, column=8, padx=4)
        ttk.Button(actions, text="Reload Current Session JSON", command=self.reload_current_session_plot).grid(
            row=0, column=9, padx=4
        )
        ttk.Button(actions, text="Toggle CSV Column", command=self.toggle_csv_column).grid(row=0, column=10, padx=4)

        self.limit_var = tk.StringVar(value=f"Units tested: 0 / {MAX_UNITS_PER_SESSION}")
        ttk.Label(actions, textvariable=self.limit_var).grid(row=0, column=11, padx=12, sticky="w")
        ttk.Label(actions, text="Runs").grid(row=1, column=0, sticky="e", padx=(4, 2), pady=(6, 0))
        ttk.Spinbox(actions, from_=1, to=50, textvariable=self.batch_run_count_var, width=6).grid(
            row=1, column=1, sticky="w", padx=(0, 8), pady=(6, 0)
        )
        ttk.Label(actions, text="Delay (s)").grid(row=1, column=2, sticky="e", padx=(4, 2), pady=(6, 0))
        ttk.Spinbox(actions, from_=0, to=300, increment=1, textvariable=self.batch_delay_s_var, width=6).grid(
            row=1, column=3, sticky="w", padx=(0, 8), pady=(6, 0)
        )
        ttk.Checkbutton(actions, text="Dark Mode", variable=self.dark_mode_var, command=self._on_toggle_dark_mode).grid(
            row=1, column=4, sticky="w", padx=(8, 0), pady=(6, 0)
        )

        self.main_notebook = ttk.Notebook(self.root)
        self.main_notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.test_tab = ttk.Frame(self.main_notebook, padding=8)
        self.plot_tab = ttk.Frame(self.main_notebook, padding=8)
        self.log_tab = ttk.Frame(self.main_notebook, padding=8)
        self.debug_tab = ttk.Frame(self.main_notebook, padding=8)

        self.main_notebook.add(self.test_tab, text="Test Results")
        self.main_notebook.add(self.plot_tab, text="Live Plot")
        self.main_notebook.add(self.log_tab, text="Run Log")
        self.main_notebook.add(self.debug_tab, text="Serial Consoles")

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
            self.tree.column(c, width=w, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        self.tree.configure(displaycolumns=tuple(c for c in cols if c != "sample_csv"))
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        designer = ttk.LabelFrame(self.plot_tab, text="Generic Sample Format", padding=8)
        designer.pack(fill=tk.X, pady=(0, 6))
        self._build_sample_format_controls(designer)

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
        self.debug_notebook = ttk.Notebook(debug)
        self.debug_notebook.pack(fill=tk.BOTH, expand=True)
        self._rebuild_measureand_editor_rows(self.sample_field_defs)
        self._apply_measureand_config(show_message=False)

    def show_plot_tab(self):
        self.main_notebook.select(self.plot_tab)

    def show_console_tab(self):
        self.main_notebook.select(self.debug_tab)

    def toggle_sample_format_panel(self, expanded=None):
        if expanded is None:
            expanded = not self.sample_format_expanded
        self.sample_format_expanded = bool(expanded)
        if self.sample_format_expanded:
            self.sample_format_body.pack(fill=tk.BOTH, expand=True)
            self.sample_format_toggle_btn.configure(text="Hide Setup")
        else:
            self.sample_format_body.pack_forget()
            self.sample_format_toggle_btn.configure(text="Show Setup")

    def show_about_dialog(self):
        message = (
            f"{APP_NAME} {APP_VERSION}\n\n"
            f"{APP_SUBTITLE}\n"
            "Engineering workbench for serial sensor debug, validation, and recovery.\n"
            "Designed for troubleshooting failed production/test/calibration units,\n"
            "with live plotting, session analysis, and terminal-style serial tools.\n\n"
            f"Author: {APP_AUTHOR}"
        )
        messagebox.showinfo("About", message)

    def detach_tabs_window(self):
        try:
            if not self.tabs_detached:
                self.root.tk.call("wm", "manage", str(self.main_notebook))
                self.root.tk.call("wm", "title", str(self.main_notebook), "SBS Tabs")
                self.root.tk.call("wm", "geometry", str(self.main_notebook), "980x640+120+120")
                self.detach_tabs_btn.configure(text="Dock Tabs")
                self.tabs_detached = True
            else:
                self.root.tk.call("wm", "forget", str(self.main_notebook))
                self.main_notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
                self.detach_tabs_btn.configure(text="Detach Tabs")
                self.tabs_detached = False
        except Exception as exc:
            messagebox.showerror("Detach Not Available", f"Could not detach/dock tabs on this system:\n{exc}")

    @staticmethod
    def _sanitize_measureand_key(text, idx):
        key = re.sub(r"[^a-zA-Z0-9]+", "_", str(text).strip().lower()).strip("_")
        return key or f"field_{idx + 1}"

    def _default_sample_field_defs(self):
        defs = []
        for idx, key in enumerate(TSR_FIELDS):
            unit = "V" if "voltage" in key else ("ns" if "phase" in key else "")
            scale = "milli" if unit == "V" else "raw"
            defs.append(
                {
                    "index": idx,
                    "key": key,
                    "description": DEFAULT_FIELD_DESCRIPTIONS.get(key, key.replace("_", " ").title()),
                    "unit": unit,
                    "scale": scale,
                    "min_val": "",
                    "max_val": "",
                    "stuck_n": "",
                    "expr": "",
                    "plot_live": key in LIVE_PLOT_FIELDS.values(),
                    "plot_session": key in {"red_phase", "blue_phase", "red_voltage", "blue_voltage"},
                    "live_default": key == "red_phase",
                }
            )
        return defs

    def _build_sample_format_controls(self, parent):
        toggle_row = ttk.Frame(parent)
        toggle_row.pack(fill=tk.X, pady=(0, 6))
        self.sample_format_toggle_btn = ttk.Button(toggle_row, text="Show Setup", command=self.toggle_sample_format_panel)
        self.sample_format_toggle_btn.pack(side=tk.LEFT)
        ttk.Label(toggle_row, text="Setup parser/mapping only when needed.", foreground=DARK_MUTED).pack(side=tk.LEFT, padx=(10, 0))

        self.sample_format_body = ttk.Frame(parent)
        self.sample_format_body.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(self.sample_format_body)
        top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(top, text="Example sample").pack(side=tk.LEFT)
        ttk.Entry(top, textvariable=self.example_sample_var, width=72).pack(side=tk.LEFT, padx=(6, 8), fill=tk.X, expand=True)
        ttk.Label(top, text="Delimiter").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(top, textvariable=self.delimiter_var, width=3).pack(side=tk.LEFT, padx=(0, 8))
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
            text="Paste one sensor output line, then edit names/descriptions and pick live/session fields.",
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
        self.toggle_sample_format_panel(expanded=False)

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
                self.measureand_editor, textvariable=scale_var, width=7, state="readonly", values=list(UNIT_SCALE_FACTORS.keys())
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

    def load_measureands_from_example(self):
        raw = self.example_sample_var.get().strip()
        delim = self.delimiter_var.get() or ","
        if not raw:
            messagebox.showwarning("No Example", "Paste one serial output line first.")
            return
        tokens = [t.strip() for t in raw.split(delim)]
        if len(tokens) < 2:
            messagebox.showerror("Invalid Example", "Example line did not split into comma-separated fields.")
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

    def reset_measureands_default(self):
        self.sample_field_defs = self._default_sample_field_defs()
        self._rebuild_measureand_editor_rows(self.sample_field_defs)
        self._apply_measureand_config(show_message=True)

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
                "scale": d.get("scale", "raw") if d.get("scale", "raw") in UNIT_SCALE_FACTORS else "raw",
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
        return UNIT_SCALE_FACTORS.get(meta.get("scale", "raw"), 1.0)

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

        ttk.Label(top, text=f"Runs: {len(rows)}  |  Grouped by serial").pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(top, text="Plot field").pack(side=tk.LEFT)
        metric_label_var = tk.StringVar(value=available[0])
        metric_combo = ttk.Combobox(top, textvariable=metric_label_var, state="readonly", width=32, values=available)
        metric_combo.pack(side=tk.LEFT, padx=(6, 10))
        compare_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Compare to Baseline", variable=compare_var).pack(side=tk.LEFT, padx=(8, 8))

        canvas = tk.Canvas(win, bg="#0b1220", highlightthickness=1, highlightbackground=DARK_BORDER)
        canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        render = lambda: self._draw_session_metric_plot(
            canvas,
            rows,
            plot_fields.get(metric_label_var.get(), "red_noise_ns"),
            metric_label_var.get(),
            compare_baseline=bool(compare_var.get()),
        )
        metric_combo.bind("<<ComboboxSelected>>", lambda _e: render())
        compare_var.trace_add("write", lambda *_: render())
        canvas.bind("<Configure>", lambda _e: render())

        def load_baseline_and_render():
            if self._load_baseline_rows():
                compare_var.set(True)
                render()

        ttk.Button(top, text="Load Baseline", command=load_baseline_and_render).pack(side=tk.LEFT, padx=(0, 8))
        render()

    def _load_baseline_rows(self):
        path = filedialog.askopenfilename(
            title="Load Baseline Session JSON",
            initialdir=self.session_dir,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return False
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.baseline_rows = self._normalize_session_rows(data)
            self.log(f"Loaded baseline rows: {len(self.baseline_rows)} from {path}")
            return bool(self.baseline_rows)
        except Exception as exc:
            messagebox.showerror("Baseline Load Failed", str(exc))
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

    def _draw_session_metric_plot(self, canvas, rows, metric_key, metric_label, compare_baseline=False):
        canvas.delete("all")
        if not rows:
            canvas.create_text(20, 20, text="No session data to plot.", anchor="nw", fill=DARK_MUTED)
            return

        width = max(int(canvas.winfo_width()), 320)
        height = max(int(canvas.winfo_height()), 220)
        left, top, right, bottom = 70, 30, width - 20, height - 55
        pw = max(right - left, 1)
        ph = max(bottom - top, 1)

        raw_y_vals = [self._to_float(r.get(metric_key)) for r in rows]
        if compare_baseline and self.baseline_rows:
            base_by_serial = {}
            for row in self.baseline_rows:
                serial = str(row.get("serial", "")).strip() or "UNKNOWN"
                val = self._to_float(row.get(metric_key))
                if np.isfinite(val):
                    base_by_serial.setdefault(serial, []).append(val)
            base_mean = {k: float(np.mean(v)) for k, v in base_by_serial.items() if v}
            adjusted = []
            for i, row in enumerate(rows):
                serial = str(row.get("serial", "")).strip() or "UNKNOWN"
                cur = raw_y_vals[i]
                ref = base_mean.get(serial)
                adjusted.append(cur - ref if np.isfinite(cur) and ref is not None else np.nan)
            raw_y_vals = adjusted
        finite_raw_y = [v for v in raw_y_vals if np.isfinite(v)]
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
        if compare_baseline:
            y_axis_label = f"Delta {y_axis_label}"

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

        # Group rows by serial and distribute groups across the full axis width.
        serials = []
        for row in rows:
            serial = str(row.get("serial", "")).strip() or "UNKNOWN"
            if serial not in serials:
                serials.append(serial)
        n_serials = len(serials)
        # Dynamic side padding keeps endpoints visible without wasting large central area.
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

        # Alternating light bands make each serial set easier to identify.
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

        serial_row_idxs = {serial: [] for serial in serials}
        for idx, row in enumerate(rows):
            serial = str(row.get("serial", "")).strip() or "UNKNOWN"
            serial_row_idxs[serial].append(idx)

        x_positions = [left + pw / 2.0] * len(rows)
        run_labels = ["r1"] * len(rows)
        for serial in serials:
            idxs = serial_row_idxs[serial]
            count = len(idxs)
            if n_serials > 1:
                neighbor_gap = (x_right - x_left) / (n_serials - 1)
            else:
                neighbor_gap = max(x_right - x_left, 1.0)
            cluster_span = min(28.0, neighbor_gap * 0.55)
            intra_step = 0.0 if count <= 1 else (cluster_span / (count - 1))
            start = -((count - 1) * intra_step) / 2.0
            for j, row_idx in enumerate(idxs):
                x_positions[row_idx] = serial_base_x[serial] + start + (j * intra_step)
                run_idx_raw = rows[row_idx].get("run_index")
                try:
                    run_idx_num = int(run_idx_raw)
                except (TypeError, ValueError):
                    run_idx_num = j + 1
                run_labels[row_idx] = f"r{run_idx_num}"

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

        canvas.create_text((left + right) / 2, height - 18, text="Serial Number", anchor="center", fill=DARK_TEXT)
        canvas.create_text(18, (top + bottom) / 2, text=y_axis_label, angle=90, anchor="center", fill=DARK_TEXT)

        metric_color = self._session_metric_color(metric_key)
        for i, row in enumerate(rows):
            v = self._to_float(row.get(metric_key))
            if not np.isfinite(v):
                continue
            v *= scale_factor
            x = x_positions[i]
            y = bottom - ((v - ymin) / (ymax - ymin)) * ph
            canvas.create_oval(x - 2, y - 2, x + 2, y + 2, fill=metric_color, outline=metric_color)
            canvas.create_text(
                x + 5,
                y - 5,
                text=run_labels[i],
                anchor="sw",
                fill=DARK_MUTED,
                font=("Segoe UI", 6),
            )

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
            font=("Segoe UI", 9),
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

            tk.Label(card, textvariable=slot_var, font=("Segoe UI", 8, "bold"), bg=DARK_PANEL, fg=DARK_TEXT).pack(anchor="w")
            tk.Label(card, textvariable=port_var, fg=DARK_TEXT, bg=DARK_PANEL).pack(anchor="w")
            tk.Label(card, textvariable=serial_var, fg=DARK_MUTED, bg=DARK_PANEL).pack(anchor="w")
            state_label = tk.Label(card, textvariable=state_var, fg="white", bg="#6b7280", width=11, relief=tk.FLAT, padx=3, pady=1)
            state_label.pack(anchor="w", pady=(2, 0))

            self.port_slots[idx] = {
                "card": card,
                "slot_var": slot_var,
                "port_var": port_var,
                "serial_var": serial_var,
                "state_var": state_var,
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

    def required_setup_values(self):
        return {
            "Operator": self.operator_var.get().strip(),
            "Station": self.station_var.get().strip(),
            "Bath ID": self.bath_id_var.get().strip(),
            "Bath Temp (C)": self.bath_temp_c_var.get().strip(),
            "Salinity (PSU)": self.salinity_psu_var.get().strip(),
        }

    def missing_setup_fields(self):
        return [name for name, value in self.required_setup_values().items() if not value]

    def update_run_button_state(self):
        connected = bool(self.serial_pool)
        setup_complete = not self.missing_setup_fields()
        allow_run = connected and setup_complete and not self.run_in_progress
        self.run_btn.configure(state=tk.NORMAL if allow_run else tk.DISABLED)

    @staticmethod
    def status_color(status):
        colors = {
            "DISCONNECTED": "#6b7280",
            "CONNECTED": "#2563eb",
            "RUNNING": "#f59e0b",
            "COMPLETE": "#0ea5a3",
            "PASS": "#16a34a",
            "WARN": "#d97706",
            "FAIL": "#dc2626",
            "ERROR": "#7c3aed",
        }
        return colors.get(status, "#6b7280")

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
        slot["state_label"].configure(bg=self.status_color(status))
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
            if from_entry:
                box = self.debug_tabs[port]["text"]
                box.insert(tk.END, f"> {cmd}\n")
                self._ensure_console_trailing_newline(port)
            self.send_cmd(ser, cmd, port=port)
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
                self.read_line(ser, port=port)
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
                line = self.read_line(ser, port=port)
                if not line:
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
        payload = "" if payload is None else str(payload)
        box = self.ensure_debug_tab(port)
        info = self.debug_tabs[port]
        ts = dt.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        line = f"[{ts}] {direction}: {payload}\n"
        box.insert(tk.END, line)
        info["lines"] += 1
        if info["lines"] > self.debug_max_lines:
            box.delete("1.0", "2.0")
            info["lines"] -= 1
        box.see(tk.END)
        self.manual_capture_rows.append(
            {"timestamp": dt.datetime.now().isoformat(timespec="milliseconds"), "port": port, "direction": direction, "payload": payload}
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
                    line = ser.readline().decode("utf-8", errors="ignore").strip()
                    if line:
                        self.serial_debug(port, "RX", line)
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
                slot["state_label"].configure(bg=self.status_color(slot["state_var"].get()))
                r = idx // cols
                c = idx % cols
                card.grid(row=r, column=c, padx=3, pady=2, sticky="nsew")
            else:
                slot["port_var"].set("(empty)")
                slot["serial_var"].set("SN: -")
                slot["state_var"].set("DISCONNECTED")
                slot["state_label"].configure(bg=self.status_color("DISCONNECTED"))
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
        folder = os.path.join(SENSOR_TEST_DIR, serial_number, PRECAL_TEST_SUBDIR)
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
            "station": self.station_var.get().strip(),
            "bath_id": self.bath_id_var.get().strip(),
            "bath_temp_c": self.bath_temp_c_var.get().strip(),
            "salinity_psu": self.salinity_psu_var.get().strip(),
        }

        self.run_in_progress = True
        self.update_run_button_state()
        with self.run_state_lock:
            self.active_run_ports = set(connected_ports)
            self.run_threads = {}
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
                    "station": setup["station"],
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
        self.log(f"New session started: {self.session_id}")
        self.log(f"Session summary file: {self.session_csv}")
        self.update_port_grid()

    def shutdown(self):
        self.shutdown_event.set()

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


def main():
    global SENSOR_TEST_DIR
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





