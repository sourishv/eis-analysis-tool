import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog
import threading
import time
import os
import re
import webbrowser
import shutil
import subprocess
import numpy as np
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

try:
    import pypalmsens as ps
except Exception:
    ps = None

class EisAnalysisTool:
    def __init__(self, root):
        self.root = root
        self.root.title("EIS Analysis Tool")
        self.root.geometry("1080x760")
        self.root.minsize(980, 680)

        self.theme = {
            "bg": "#0f141b",
            "panel": "#171f29",
            "panel_alt": "#1c2633",
            "text": "#e9eff7",
            "muted": "#a8b8cb",
            "accent": "#3b82f6",
            "accent_active": "#5a97f7",
            "success": "#34a853",
            "warning": "#dca64a",
            "danger": "#e35d6a",
            "line": "#2f3f53",
            "line_soft": "#243140",
            "entry_bg": "#121a24",
            "tab_bg": "#1a2430",
            "tab_text": "#d8e3f0",
            "tab_hover": "#24364c",
            "btn_secondary": "#24364c",
            "btn_secondary_hover": "#2c4561",
            "btn_disabled": "#1e2a37",
            "diag_fail": "#b9434f",
            "diag_caution": "#be8a3a",
            "diag_pass": "#3f9353",
        }
        self.root.configure(bg=self.theme["bg"])

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("App.TFrame", background=self.theme["bg"])
        style.configure("Card.TFrame", background=self.theme["panel"], relief="solid", borderwidth=1, bordercolor=self.theme["line_soft"])
        style.configure("InnerCard.TFrame", background=self.theme["panel_alt"], relief="solid", borderwidth=1, bordercolor=self.theme["line"])
        style.configure("TLabel", background=self.theme["bg"], foreground=self.theme["text"], font=("Segoe UI", 11))
        style.configure("Card.TLabel", background=self.theme["panel"], foreground=self.theme["text"], font=("Segoe UI", 11))
        style.configure("Inner.Card.TLabel", background=self.theme["panel_alt"], foreground=self.theme["text"], font=("Segoe UI", 11))
        style.configure("Muted.Card.TLabel", background=self.theme["panel"], foreground=self.theme["muted"], font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=self.theme["bg"], foreground=self.theme["text"], font=("Segoe UI Semibold", 20))
        style.configure("SectionTitle.TLabel", background=self.theme["panel"], foreground=self.theme["text"], font=("Segoe UI Semibold", 12))
        style.configure("TEntry", fieldbackground=self.theme["entry_bg"], foreground=self.theme["text"], insertcolor=self.theme["text"], bordercolor=self.theme["line"], lightcolor=self.theme["line"], darkcolor=self.theme["line"], padding=(8, 6), font=("Segoe UI", 11))
        style.configure("Primary.TButton", font=("Segoe UI Semibold", 11), padding=(14, 9), background=self.theme["accent"], foreground=self.theme["bg"], bordercolor=self.theme["accent"])
        style.map("Primary.TButton", background=[("active", self.theme["accent_active"]), ("disabled", self.theme["btn_disabled"])], foreground=[("disabled", self.theme["muted"])])
        style.configure("Secondary.TButton", font=("Segoe UI Semibold", 11), padding=(14, 9), background=self.theme["btn_secondary"], foreground=self.theme["text"], bordercolor=self.theme["line"])
        style.map("Secondary.TButton", background=[("active", self.theme["btn_secondary_hover"]), ("disabled", self.theme["btn_disabled"])], foreground=[("disabled", self.theme["muted"])])
        style.configure("Run.TButton", font=("Segoe UI Semibold", 11), padding=(14, 9), background=self.theme["success"], foreground="white", bordercolor=self.theme["success"])
        style.map("Run.TButton", background=[("active", "#43ba62"), ("disabled", self.theme["btn_disabled"])], foreground=[("disabled", self.theme["muted"])])
        style.configure("Stop.TButton", font=("Segoe UI Semibold", 11), padding=(14, 9), background=self.theme["danger"], foreground="white", bordercolor=self.theme["danger"])
        style.map("Stop.TButton", background=[("active", "#f06b78"), ("disabled", self.theme["btn_disabled"])], foreground=[("disabled", self.theme["muted"])])
        style.configure("Status.TLabel", background=self.theme["panel"], foreground=self.theme["danger"], font=("Segoe UI", 11, "bold"))
        style.configure("TNotebook", background=self.theme["bg"], borderwidth=0, tabmargins=(2, 2, 2, 0))
        style.configure("TNotebook.Tab", font=("Segoe UI Semibold", 10), background=self.theme["tab_bg"], foreground=self.theme["tab_text"], padding=(16, 8))
        style.map("TNotebook.Tab", background=[("selected", self.theme["panel"]), ("active", self.theme["tab_hover"])], foreground=[("selected", self.theme["text"]), ("active", self.theme["text"])])
        style.configure("Hint.Card.TLabel", background=self.theme["panel_alt"], foreground=self.theme["muted"], font=("Segoe UI", 9))
        style.configure("Green.Horizontal.TProgressbar", troughcolor=self.theme["entry_bg"], background=self.theme["accent"], bordercolor=self.theme["line"], lightcolor=self.theme["accent"], darkcolor=self.theme["accent"])

        main_frame = ttk.Frame(root, style="App.TFrame")
        main_frame.pack(fill="both", expand=True, padx=14, pady=14)

        header_frame = ttk.Frame(main_frame, style="App.TFrame")
        header_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(header_frame, text="EIS Analysis Tool", style="Header.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text="Application for PalmSens EIS acquisition and diagnostics", style="TLabel").pack(anchor="w", pady=(2, 0))

        connect_frame = ttk.Frame(main_frame, style="Card.TFrame", padding=(14, 12))
        connect_frame.pack(fill="x", pady=(0, 12))
        connect_frame.columnconfigure(5, weight=1)

        ttk.Label(connect_frame, text="Connection", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w", padx=(2, 10))

        ttk.Label(connect_frame, text="Device", style="Card.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 6))
        self.device_var = tk.StringVar(value="Sensit BT")
        self.device_combo = ttk.Combobox(
            connect_frame,
            textvariable=self.device_var,
            values=["Sensit BT", "Simulated Mode", "Messy Data", "Calibration"],
            state="readonly",
            width=18,
        )
        self.device_combo.grid(row=0, column=2, sticky="w", padx=(0, 10))

        self.connect_btn = ttk.Button(connect_frame, text="Connect", command=self.start_connect_thread, style="Primary.TButton")
        self.connect_btn.grid(row=0, column=3, sticky="w", padx=(0, 6))

        self.disconnect_btn = ttk.Button(connect_frame, text="Disconnect", command=self.disconnect_device, style="Secondary.TButton", state="disabled")
        self.disconnect_btn.grid(row=0, column=4, sticky="w", padx=6)

        self.status_label = ttk.Label(
            connect_frame, 
            text="Status: Disconnected", 
            style="Status.TLabel"
        )
        self.status_label.grid(row=1, column=0, columnspan=6, sticky="w", padx=(2, 2), pady=(8, 0))
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)
        # Shared progress variable for progress bars (defined before any bar uses it)
        self.progress_var = tk.DoubleVar(value=0.0)

        # PyPalmSens runtime state
        self.ps_manager = None
        self.ps_instrument = None
        self.connection_mode = None  # "sensit_bt" or "simulated" or "messy" or "calibration"
        self.connection_in_progress = False
        self.cancel_connect_requested = False
        self.expected_points = 0
        self.measurement_in_progress = False
        self.target_mac = "00:16:A4:79:4E:03"
        self.measurement_start_time = None
        self.last_point_time = None
        self.last_point_count = 0
        self.stop_requested = False
        self.callback_debug_count = 0
        self.last_plot_update_time = 0  # Throttle plot updates (milliseconds)
        self.latest_plot_data = {
            'frequency': np.array([]),
            'z_real': np.array([]),
            'z_imag': np.array([]),
        }
        self.test_run_counter = 0
        self.sim_reference_profile = None
        self.shared_progress_frame = None
        self.shared_progress = None
        self.shared_progress_label = None
        self._measurement_drag_active = False
        self._osk_launch_cmd = self._detect_onscreen_keyboard_command()
        self._last_osk_launch_time = 0.0

        # --- Tab 1: Load Measurement ---
        self.eis_frame = ttk.Frame(self.notebook, style="Card.TFrame")
        self.notebook.add(self.eis_frame, text='Measurement Setup') # <-- Renamed Tab

        # Scrollable measurement setup for smaller touch displays.
        self.eis_canvas = tk.Canvas(
            self.eis_frame,
            bg=self.theme["panel"],
            highlightthickness=0,
            borderwidth=0,
        )
        self.eis_canvas.pack(side="left", fill="both", expand=True)
        self.eis_scrollbar = ttk.Scrollbar(self.eis_frame, orient="vertical", command=self.eis_canvas.yview)
        self.eis_scrollbar.pack(side="right", fill="y")
        self.eis_canvas.configure(yscrollcommand=self.eis_scrollbar.set)

        self.eis_scroll_content = ttk.Frame(self.eis_canvas, style="Card.TFrame")
        self.eis_canvas_window = self.eis_canvas.create_window((0, 0), window=self.eis_scroll_content, anchor="nw")
        self.eis_scroll_content.bind("<Configure>", self._on_measurement_content_configure)
        self.eis_canvas.bind("<Configure>", self._on_measurement_canvas_configure)
        self._bind_scroll_events(self.eis_frame)
        self._bind_scroll_events(self.eis_canvas)
        self._bind_scroll_events(self.eis_scroll_content)
        self.eis_canvas.bind("<ButtonPress-1>", self._on_measurement_drag_start, add="+")
        self.eis_canvas.bind("<B1-Motion>", self._on_measurement_drag, add="+")
        self.eis_scroll_content.bind("<ButtonPress-1>", self._on_measurement_drag_start, add="+")
        self.eis_scroll_content.bind("<B1-Motion>", self._on_measurement_drag, add="+")

        # --- NEW: Load Data Button ---
        # The parameter fields have been removed.
        load_frame = ttk.Frame(self.eis_scroll_content, style="Card.TFrame", padding=(18, 16))
        load_frame.pack(fill="both", expand=True)

        ttk.Label(load_frame, text="Measurement Configuration", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(load_frame, text="Tune scan parameters and run the real-time EIS test.", style="Muted.Card.TLabel").pack(anchor="w", pady=(2, 14))

        # --- Editable EIS parameter fields ---
        params = [
            ("Start Frequency (Hz)", "1e5", "Highest scan frequency (example: 100000)"),
            ("End Frequency (Hz)", "1e-1", "Lowest scan frequency (example: 0.1)"),
            ("Voltage Amplitude (mV)", "50", "AC excitation amplitude in millivolts"),
            ("Points per Decade", "5", "Resolution of the sweep (higher = denser scan)"),
        ]

        self.param_vars = {}
        params_frame = ttk.Frame(load_frame, style="InnerCard.TFrame", padding=(16, 14))
        params_frame.pack(fill="x", pady=(0, 18))
        params_frame.columnconfigure(0, weight=0)
        params_frame.columnconfigure(1, weight=1)

        for i, (label_text, default, helper_text) in enumerate(params):
            lbl = ttk.Label(params_frame, text=label_text, style="Inner.Card.TLabel")
            lbl.grid(row=i * 2, column=0, sticky="w", padx=(0, 12), pady=(4, 0))
            var = tk.StringVar(value=default)
            ent = ttk.Entry(params_frame, textvariable=var, width=24)
            ent.grid(row=i * 2, column=1, sticky="ew", pady=(4, 0))
            self._bind_entry_touch_focus(ent)
            hint = ttk.Label(params_frame, text=helper_text, style="Hint.Card.TLabel")
            hint.grid(row=i * 2 + 1, column=1, sticky="w", pady=(1, 8))
            self.param_vars[label_text] = var

        if self._osk_launch_cmd:
            ttk.Label(
                load_frame,
                text=f"On-screen keyboard available via: {self._osk_launch_cmd}",
                style="Hint.Card.TLabel",
            ).pack(anchor="w", pady=(0, 6))

            self.keyboard_btn = ttk.Button(
                load_frame,
                text="Open Keyboard",
                command=self.open_onscreen_keyboard,
                style="Secondary.TButton",
            )
            self.keyboard_btn.pack(anchor="w", pady=(0, 10))

        controls_frame = ttk.Frame(load_frame, style="Card.TFrame")
        controls_frame.pack(fill="x", pady=(2, 8))

        self.run_test_btn = ttk.Button(
            controls_frame,
            text="Run Test",
            command=self.start_run_test_thread,
            style="Run.TButton",
            state="disabled"
        )
        self.run_test_btn.pack(side="left")

        self.calibrate_btn = ttk.Button(
            controls_frame,
            text="Calibrate",
            command=self.start_calibration_thread,
            style="Secondary.TButton",
            state="disabled"
        )
        self.calibrate_btn.pack(side="right")

        self.stop_test_btn = ttk.Button(
            controls_frame,
            text="Stop Test",
            command=self.request_stop_measurement,
            style="Stop.TButton",
            state="disabled"
        )
        self.stop_test_btn.pack(side="left", padx=(10, 0))

        self.load_progress_lbl = ttk.Label(load_frame, text="No test running", style="Muted.Card.TLabel")
        self.load_progress_lbl.pack(anchor="w", pady=(8, 0))
        self._bind_drag_scroll_to_descendants(self.eis_scroll_content)
        # --- END OF CHANGES TO TAB 1 ---

        # --- Tab 2: Nyquist Plot ---
        nyquist_tab = ttk.Frame(self.notebook, style="Card.TFrame")
        self.notebook.add(nyquist_tab, text='Nyquist Plot')
        self.nyquist_fig = Figure(figsize=(6, 4), dpi=100, facecolor=self.theme["panel"])
        self.nyquist_ax = self.nyquist_fig.add_subplot(111)
        self.nyquist_canvas = FigureCanvasTkAgg(self.nyquist_fig, master=nyquist_tab)
        self.nyquist_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1, padx=10, pady=(10, 6))
        
        nyquist_export_frame = ttk.Frame(nyquist_tab, style="Card.TFrame")
        nyquist_export_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        self.export_nyquist_btn = ttk.Button(
            nyquist_export_frame,
            text="💾 Save to Device",
            style="Secondary.TButton",
            command=lambda: self.export_plot('nyquist'),
            state="disabled",
        )
        self.export_nyquist_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.export_nyquist_cloud_btn = ttk.Button(
            nyquist_export_frame,
            text="☁ Upload",
            style="Secondary.TButton",
            command=lambda: self.export_plot_to_cloud('nyquist'),
            state="disabled",
        )
        self.export_nyquist_cloud_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

        self.nyquist_ax.plot_data = ([], []) 
        self.nyquist_line = None
        self.nyquist_diag_text = None
        self.nyquist_quality_text = None
        self.nyquist_calibration_text = None
        
        # --- Tab 3: Bode Plot (Magnitude Only) ---
        bode_tab = ttk.Frame(self.notebook, style="Card.TFrame")
        self.notebook.add(bode_tab, text='Bode Plot')
        
        self.bode_fig = Figure(figsize=(6, 4), dpi=100, facecolor=self.theme["panel"])
        
        self.bode_ax_mag = self.bode_fig.add_axes([0.15, 0.15, 0.8, 0.75], zorder=1)
        self.bode_cbar_ax = self.bode_fig.add_axes([0.15, 0.15, 0.03, 0.75], zorder=2)
        
        self.bode_ax_mag.patch.set_alpha(0) 
        self.bode_ax_mag.plot_data = ([], []) 

        self.bode_line = None
        self.bode_diag_text = None
        self.bode_quality_text = None
        self.bode_calibration_text = None
        self.bode_eval_vline = None
        self.bode_eval_hline = None
        self.bode_eval_point = None
        self.bode_eval_text = None

        self.bode_canvas = FigureCanvasTkAgg(self.bode_fig, master=bode_tab)
        self.bode_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1, padx=10, pady=(10, 6))

        bode_export_frame = ttk.Frame(bode_tab, style="Card.TFrame")
        bode_export_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        self.export_bode_btn = ttk.Button(
            bode_export_frame,
            text="💾 Save to Device",
            style="Secondary.TButton",
            command=lambda: self.export_plot('bode'),
            state="disabled",
        )
        self.export_bode_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.export_bode_cloud_btn = ttk.Button(
            bode_export_frame,
            text="☁ Upload",
            style="Secondary.TButton",
            command=lambda: self.export_plot_to_cloud('bode'),
            state="disabled",
        )
        self.export_bode_cloud_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

        # --- Tab 4: Output Log ---
        log_tab = ttk.Frame(self.notebook, style="Card.TFrame", padding=(10, 10))
        self.notebook.add(log_tab, text='Output Log')
        # Progress bar shown while a test is running
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(log_tab, style='Green.Horizontal.TProgressbar', orient='horizontal', mode='determinate', variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', padx=4, pady=(4, 8))

        self.output_text = scrolledtext.ScrolledText(
            log_tab,
            height=10,
            state="disabled",
            font=("Consolas", 10),
            bg=self.theme["entry_bg"],
            fg=self.theme["text"],
            insertbackground=self.theme["text"],
            relief="solid",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=self.theme["line_soft"],
            highlightcolor=self.theme["line"],
        )
        self.output_text.pack(fill="both", expand=True, padx=4, pady=4)

        # --- Initialize Plots & Annotations ---
        self.init_nyquist_plot()
        self.init_bode_plot()

        self.nyquist_annot = self.nyquist_ax.annotate("", xy=(0,0), xytext=(15,15),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.35", fc=self.theme["panel_alt"], ec=self.theme["line"], alpha=0.95),
            color=self.theme["text"],
            arrowprops=dict(arrowstyle="->", color=self.theme["muted"]))
        self.nyquist_annot.set_visible(False)

        self.bode_annot = self.bode_ax_mag.annotate("", xy=(0,0), xytext=(15,15),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.35", fc=self.theme["panel_alt"], ec=self.theme["line"], alpha=0.95),
            color=self.theme["text"],
            arrowprops=dict(arrowstyle="->", color=self.theme["muted"]))
        self.bode_annot.set_visible(False)

        self.nyquist_canvas.mpl_connect("motion_notify_event", self.on_plot_hover)
        self.bode_canvas.mpl_connect("motion_notify_event", self.on_plot_hover)

    # --- REMOVED _create_param_entry ---

    # --- Connection Logic (Simulated) ---

    def start_connect_thread(self):
        """Connect to selected device mode."""
        selected_device = self.device_var.get().strip()
        self.connection_in_progress = True
        self.cancel_connect_requested = False
        if selected_device == "Sensit BT":
            self.connect_btn.config(text="Cancel", command=self.request_cancel_connect, state="normal")
        else:
            self.connect_btn.config(state="disabled")
        self.disconnect_btn.config(state="disabled")
        self.run_test_btn.config(state="disabled")
        self.status_label.config(text="Status: Connecting (USB/Bluetooth)...", foreground=self.theme["warning"])
        threading.Thread(target=self.connect_device, daemon=True).start()

    def request_cancel_connect(self):
        """Request cancellation of an in-progress real Bluetooth connection attempt."""
        if not self.connection_in_progress:
            return
        self.cancel_connect_requested = True
        self.connect_btn.config(state="disabled")
        self.status_label.config(text="Status: Cancelling connection...", foreground=self.theme["warning"])
        self.log_message("Cancel requested: connection attempt will stop as soon as possible.")

    def _finish_connection_cancelled(self):
        """Finalize UI state after a cancelled connection attempt."""
        self.connection_in_progress = False
        self.cancel_connect_requested = False
        self._set_disconnected_ui()
        self.log_message("Connection attempt cancelled.")

    def connect_simulated_device(self):
        """Connect to simulated runtime mode (no Bluetooth required)."""
        try:
            try:
                if self.ps_manager is not None:
                    self.ps_manager.disconnect()
            except Exception:
                pass

            self.ps_manager = None
            self.ps_instrument = None
            self.connection_mode = "simulated"
            self.log_message("Connected to Simulated Mode.")
            self.root.after(0, self._set_connected_ui, "Simulated Mode")
        except Exception as e:
            self.log_message(f"Simulated mode connection failed: {e}")
            self.connection_mode = None
            self.root.after(0, self._set_disconnected_ui)

    def connect_messy_device(self):
        """Connect to messy simulated runtime mode for setup/fit testing."""
        try:
            try:
                if self.ps_manager is not None:
                    self.ps_manager.disconnect()
            except Exception:
                pass

            self.ps_manager = None
            self.ps_instrument = None
            self.connection_mode = "messy"
            self.log_message("Connected to Messy Data mode.")
            self.root.after(0, self._set_connected_ui, "Messy Data")
        except Exception as e:
            self.log_message(f"Messy data mode connection failed: {e}")
            self.connection_mode = None
            self.root.after(0, self._set_disconnected_ui)

    def connect_calibration_device(self):
        """Connect to calibration runtime mode (simulated 3-pass sequence)."""
        try:
            try:
                if self.ps_manager is not None:
                    self.ps_manager.disconnect()
            except Exception:
                pass

            self.ps_manager = None
            self.ps_instrument = None
            self.connection_mode = "calibration"
            self.log_message("Connected to Calibration mode.")
            self.root.after(0, self._set_connected_ui, "Calibration")
        except Exception as e:
            self.log_message(f"Calibration mode connection failed: {e}")
            self.connection_mode = None
            self.root.after(0, self._set_disconnected_ui)

    def connect_device(self):
        """Discover and connect to a PalmSens instrument over USB or Bluetooth."""
        if ps is None:
            self.log_message("ERROR: PyPalmSens is not installed. Install with: pip install pypalmsens")
            self.connection_mode = None
            self.root.after(0, self._set_disconnected_ui)
            return

        try:
            # Best-effort cleanup of an existing manager before reconnecting
            try:
                if self.ps_manager is not None:
                    self.ps_manager.disconnect()
            except Exception:
                pass
            self.ps_manager = None
            self.ps_instrument = None

            # Discover in two passes: USB-first, then USB+Bluetooth fallback.
            self.log_message("Scanning for PalmSens instruments (USB first)...")
            discovery_profiles = [
                (
                    "USB",
                    dict(
                        ftdi=True,
                        usbcdc=True,
                        winusb=True,
                        bluetooth=False,
                        serial=True,
                        ignore_errors=True,
                    ),
                ),
                (
                    "USB/Bluetooth",
                    dict(
                        ftdi=True,
                        usbcdc=True,
                        winusb=True,
                        bluetooth=True,
                        serial=True,
                        ignore_errors=True,
                    ),
                ),
            ]

            instruments = []
            seen_ids = set()
            for profile_name, profile_args in discovery_profiles:
                if self.cancel_connect_requested:
                    self.root.after(0, self._finish_connection_cancelled)
                    return
                try:
                    discovered = ps.discover(**profile_args)
                except Exception as discover_error:
                    self.log_message(f"Discovery ({profile_name}) failed: {discover_error}")
                    continue

                self.log_message(f"Discovery ({profile_name}) found {len(discovered)} device(s).")
                for inst in discovered:
                    inst_key = self._instrument_unique_key(inst)
                    if inst_key in seen_ids:
                        continue
                    seen_ids.add(inst_key)
                    instruments.append(inst)

                # If USB pass already discovered a wired device, skip additional discovery passes.
                if profile_name == "USB":
                    usb_candidates = [inst for inst in instruments if not self._is_bluetooth_instrument(inst)]
                    if usb_candidates:
                        break

            if self.cancel_connect_requested:
                self.root.after(0, self._finish_connection_cancelled)
                return

            if not instruments:
                self.log_message("No PalmSens instruments found on USB/Bluetooth.")
                self.log_message("Check cable/power and verify the Pi can see USB serial devices (e.g., /dev/ttyACM* or /dev/ttyUSB*).")
                self.root.after(0, self._set_disconnected_ui)
                return

            self.log_message(f"Discovered {len(instruments)} device(s):")
            for idx, inst in enumerate(instruments, start=1):
                self.log_message(f"  {idx}. {self._describe_instrument(inst)}")

            # Fixed MAC targeting for Bluetooth instruments.
            # USB devices usually do not expose MAC-style identifiers.
            selected = None
            if self.target_mac:
                for inst in instruments:
                    if self._instrument_matches_mac(inst, self.target_mac):
                        selected = inst
                        break
                if selected is None:
                    usb_candidates = [inst for inst in instruments if not self._is_bluetooth_instrument(inst)]
                    if usb_candidates:
                        selected = usb_candidates[0]
                        self.log_message(
                            f"Configured MAC {self.target_mac} not found; falling back to USB device: {self._describe_instrument(selected)}"
                        )
                    else:
                        self.log_message(f"Configured MAC {self.target_mac} not found in discovered Bluetooth devices.")
                        self.root.after(0, self._set_disconnected_ui)
                        return
                else:
                    self.log_message(f"Using configured MAC match: {self._describe_instrument(selected)}")
            else:
                usb_candidates = [inst for inst in instruments if not self._is_bluetooth_instrument(inst)]
                selected = usb_candidates[0] if usb_candidates else instruments[0]

            if self.cancel_connect_requested:
                self.root.after(0, self._finish_connection_cancelled)
                return

            self.ps_instrument = selected

            # Retry once for transient socket/binding issues
            last_err = None
            for attempt in range(1, 3):
                if self.cancel_connect_requested:
                    self.root.after(0, self._finish_connection_cancelled)
                    return
                try:
                    self.ps_manager = ps.InstrumentManager(self.ps_instrument)
                    self.ps_manager.connect()
                    break
                except Exception as e:
                    last_err = e
                    self.log_message(f"Connect attempt {attempt}/2 failed: {e}")
                    try:
                        if self.ps_manager is not None:
                            self.ps_manager.disconnect()
                    except Exception:
                        pass
                    self.ps_manager = None
                    if attempt < 2:
                        time.sleep(1.0)
            if self.ps_manager is None:
                raise last_err if last_err is not None else RuntimeError("Unknown connection error")

            serial = self.ps_manager.get_instrument_serial()
            if self._is_bluetooth_instrument(selected):
                self.connection_mode = "sensit_bt"
                mode_label = "Sensit BT"
            else:
                self.connection_mode = "sensit_usb"
                mode_label = "Sensit USB"
            self.log_message(f"Connected to {self.ps_instrument.name} (Serial: {serial})")
            self.root.after(0, self._set_connected_ui, mode_label)
        except Exception as e:
            self.log_message(f"PalmSens connection failed: {e}")
            if "FT_DEVICE_NOT_OPENED" in str(e):
                self.log_message("Hint: FTDI device found but could not be opened. Close other app instances and unload kernel FTDI serial drivers (ftdi_sio/usbserial), then reconnect USB.")
            self.ps_manager = None
            self.ps_instrument = None
            self.connection_mode = None
            self.root.after(0, self._set_disconnected_ui)

    def _describe_instrument(self, instrument):
        """Return a readable description for logs with best-effort address fields."""
        name = getattr(instrument, 'name', 'unknown')
        interface = getattr(instrument, 'interface', 'unknown')
        # Try multiple common address-like fields
        address = None
        for attr in ('address', 'mac', 'mac_address', 'identifier', 'id'):
            try:
                value = getattr(instrument, attr)
                if value:
                    address = str(value)
                    break
            except Exception:
                continue
        if address:
            return f"name={name}, interface={interface}, addr={address}"
        return f"name={name}, interface={interface}"

    def _instrument_unique_key(self, instrument):
        """Build a stable identity key so cross-pass discovery does not duplicate the same device."""
        name = str(getattr(instrument, 'name', 'unknown'))
        interface = str(getattr(instrument, 'interface', 'unknown'))
        address = None
        for attr in ('address', 'mac', 'mac_address', 'identifier', 'id'):
            try:
                value = getattr(instrument, attr)
                if value:
                    address = str(value)
                    break
            except Exception:
                continue
        return f"{name}|{interface}|{address or ''}"

    def _instrument_matches_mac(self, instrument, target_mac):
        """Best-effort match against known address-like fields and BT name suffixes."""
        norm_target = str(target_mac).replace(':', '').replace('-', '').upper()
        target_tail6 = norm_target[-6:] if len(norm_target) >= 6 else norm_target
        target_tail4 = norm_target[-4:] if len(norm_target) >= 4 else norm_target

        for attr in ('address', 'mac', 'mac_address', 'identifier', 'id', 'name'):
            try:
                value = getattr(instrument, attr)
                if not value:
                    continue
                value_str = str(value)
                norm_value = value_str.replace(':', '').replace('-', '').upper()

                # Exact match when a full MAC-like value is available.
                if norm_value == norm_target:
                    return True

                # Fallback: some Bluetooth transports expose only device alias such as
                # "PS-4E03 (Bluetooth)". Match if any hex token is a suffix of target MAC.
                for token in re.findall(r'[0-9A-Fa-f]{4,12}', value_str):
                    token = token.upper()
                    if token == norm_target:
                        return True
                    if target_tail6 and token == target_tail6:
                        return True
                    if target_tail4 and token == target_tail4:
                        return True
                    if len(token) >= 4 and norm_target.endswith(token):
                        return True
            except Exception:
                continue
        return False

    def _is_bluetooth_instrument(self, instrument):
        """Return True when instrument interface metadata indicates Bluetooth transport."""
        interface = getattr(instrument, 'interface', '')
        return 'bluetooth' in str(interface).lower()

    def disconnect_device(self):
        """Disconnect active PalmSens instrument."""
        try:
            if self.ps_manager is not None:
                self.ps_manager.disconnect()
        except Exception as e:
            self.log_message(f"Disconnect warning: {e}")
        finally:
            self.ps_manager = None
            self.ps_instrument = None
            self.connection_mode = None
            self._set_disconnected_ui()
            self.log_message("Device disconnected.")

    def _set_connected_ui(self, connected_label="Connected"):
        self.connection_in_progress = False
        self.cancel_connect_requested = False
        self.status_label.config(text=f"Status: Connected ({connected_label})", foreground=self.theme["success"])
        self.connect_btn.config(text="Connect", command=self.start_connect_thread, state="disabled")
        self.disconnect_btn.config(state="normal")
        self.run_test_btn.config(state="normal")
        if self.connection_mode in ("sensit_bt", "sensit_usb"):
            self.calibrate_btn.config(state="normal")
        else:
            self.calibrate_btn.config(state="disabled")

    def _bind_entry_touch_focus(self, entry_widget):
        """Improve touch input by forcing focus and opening OSK when available."""
        entry_widget.bind("<ButtonRelease-1>", self._on_entry_touched, add="+")
        entry_widget.bind("<FocusIn>", self._on_entry_focused, add="+")

    def _on_entry_touched(self, event):
        try:
            event.widget.focus_set()
            event.widget.icursor(tk.END)
        except Exception:
            pass

    def _on_entry_focused(self, _event):
        self.open_onscreen_keyboard()

    def _detect_onscreen_keyboard_command(self):
        """Return the first available on-screen keyboard launcher command."""
        candidates = ["wvkbd-mobintl", "matchbox-keyboard", "onboard", "florence"]
        for cmd in candidates:
            if shutil.which(cmd):
                return cmd
        return None

    def open_onscreen_keyboard(self):
        """Launch an available on-screen keyboard without blocking the UI."""
        if not self._osk_launch_cmd:
            return

        now = time.time()
        if now - self._last_osk_launch_time < 1.5:
            return

        self._last_osk_launch_time = now
        try:
            subprocess.Popen([self._osk_launch_cmd], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            self.log_message(f"Unable to launch on-screen keyboard: {e}")

    def _bind_scroll_events(self, widget):
        widget.bind("<Enter>", self._on_measurement_enter_scrollable, add="+")
        widget.bind("<Leave>", self._on_measurement_leave_scrollable, add="+")

    def _bind_drag_scroll_to_descendants(self, parent_widget):
        """Enable drag-to-scroll from anywhere in the measurement setup content."""
        for child in parent_widget.winfo_children():
            child.bind("<ButtonPress-1>", self._on_measurement_drag_start, add="+")
            child.bind("<B1-Motion>", self._on_measurement_drag, add="+")
            if child.winfo_children():
                self._bind_drag_scroll_to_descendants(child)

    def _on_measurement_enter_scrollable(self, _event):
        self.root.bind_all("<MouseWheel>", self._on_measurement_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self._on_measurement_mousewheel_linux, add="+")
        self.root.bind_all("<Button-5>", self._on_measurement_mousewheel_linux, add="+")

    def _on_measurement_leave_scrollable(self, _event):
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")

    def _on_measurement_mousewheel(self, event):
        if event.delta == 0:
            return
        self.eis_canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    def _on_measurement_mousewheel_linux(self, event):
        if event.num == 4:
            self.eis_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.eis_canvas.yview_scroll(1, "units")

    def _on_measurement_drag_start(self, event):
        """Start drag-to-scroll gesture for touch displays."""
        self._measurement_drag_active = True
        self.eis_canvas.scan_mark(event.x, event.y)

    def _on_measurement_drag(self, event):
        """Continue drag-to-scroll gesture for touch displays."""
        if not self._measurement_drag_active:
            return
        self.eis_canvas.scan_dragto(event.x, event.y, gain=1)

    def _on_measurement_content_configure(self, _event):
        self.eis_canvas.configure(scrollregion=self.eis_canvas.bbox("all"))

    def _on_measurement_canvas_configure(self, event):
        self.eis_canvas.itemconfigure(self.eis_canvas_window, width=event.width)

    def _set_disconnected_ui(self):
        self.connection_in_progress = False
        self.cancel_connect_requested = False
        self.connect_btn.config(text="Connect", command=self.start_connect_thread, state="normal")
        self.disconnect_btn.config(state="disabled")
        self.run_test_btn.config(state="disabled")
        self.calibrate_btn.config(state="disabled")
        self.status_label.config(text="Status: Disconnected", foreground=self.theme["danger"])

    # --- Plotting Initialization ---

    def init_nyquist_plot(self):
        self.nyquist_ax.clear()
        self.nyquist_ax.set_facecolor(self.theme["panel"])
        self.nyquist_ax.set_xlabel('Z_real (Ohm)')
        self.nyquist_ax.set_ylabel('-Z_imaginary (Ohm)')
        self.nyquist_ax.set_title("Nyquist Plot")
        self.nyquist_ax.grid(True, color=self.theme["line"], linewidth=0.8, alpha=0.8)
        self.nyquist_ax.tick_params(colors=self.theme["muted"])
        self.nyquist_ax.xaxis.label.set_color(self.theme["text"])
        self.nyquist_ax.yaxis.label.set_color(self.theme["text"])
        self.nyquist_ax.title.set_color(self.theme["text"])
        for spine in self.nyquist_ax.spines.values():
            spine.set_color(self.theme["line"])
        
        # --- NEW: Use ticklabel_format to get the 10^5 offset ---
        # This tells matplotlib to use scientific notation (style='sci')
        # for both axes (axis='both') and to apply it if the numbers
        # are large (scilimits=(0,0)). useMathText makes it look nice.
        self.nyquist_ax.ticklabel_format(style='sci', axis='both', scilimits=(0,0), useMathText=True)
        # --- END NEW ---
        
        self.nyquist_fig.subplots_adjust(left=0.1, right=0.95, top=0.9, bottom=0.15)
        self.nyquist_canvas.draw()
        
    def init_bode_plot(self):
        """Initializes Bode plot with color bar and fixed axes."""
        
        self.bode_ax_mag.clear()
        self.bode_ax_mag.set_facecolor(self.theme["panel"])
        self.bode_ax_mag.set_ylim(1e2, 1e10)
        self.bode_ax_mag.set_xlim(1e-2, 1e5)
        self.bode_ax_mag.set_ylabel('|Z| (Ohm)')
        self.bode_ax_mag.set_xlabel('Frequency (Hz)')
        self.bode_ax_mag.set_title("Bode Plot")
        self.bode_ax_mag.grid(True, which='both', color=self.theme["line"], linewidth=0.8, alpha=0.8)
        self.bode_ax_mag.tick_params(colors=self.theme["muted"])
        self.bode_ax_mag.xaxis.label.set_color(self.theme["text"])
        self.bode_ax_mag.yaxis.label.set_color(self.theme["text"])
        self.bode_ax_mag.title.set_color(self.theme["text"])
        for spine in self.bode_ax_mag.spines.values():
            spine.set_color(self.theme["line"])
        self.bode_ax_mag.set_yscale('log')
        self.bode_ax_mag.set_xscale('log')
        
        self.bode_cbar_ax.clear()
        self.bode_cbar_ax.set_facecolor(self.theme["panel"])
        self.bode_cbar_ax.set_yscale('log')
        self.bode_cbar_ax.set_ylim(1e2, 1e10)
        
        self.bode_cbar_ax.axhspan(1e2, 1e5, facecolor=self.theme['diag_fail'], alpha=0.45) # Red
        self.bode_cbar_ax.axhspan(1e5, 1e7, facecolor=self.theme['diag_caution'], alpha=0.45) # Yellow
        self.bode_cbar_ax.axhspan(1e7, 1e10, facecolor=self.theme['diag_pass'], alpha=0.45) # Green

        self.bode_cbar_ax.set_xticks([])
        self.bode_cbar_ax.set_yticks([])
        self.bode_cbar_ax.set_yticklabels([])
        
        self.bode_cbar_ax.patch.set_alpha(0.6)
        
        self.bode_canvas.draw()

    # --- Plot Hover Logic (Unchanged) ---
    
    def update_annotation(self, annot, ax, x, y):
        """Updates and shows an annotation."""
        annot.xy = (x, y)
        if ax == self.bode_ax_mag:
            annot.set_text(f"Freq: {x:.2e} Hz\n|Z|: {y:.2e} Ohm")
        elif ax == self.nyquist_ax:
            annot.set_text(f"Z_real: {x:.2e}\n-Z_imag: {y:.2e}")
        annot.set_visible(True)

    def on_plot_hover(self, event):
        """Handles mouse hover events on plots."""
        if event.inaxes == self.bode_ax_mag:
            annot = self.bode_annot
            data_x, data_y = self.bode_ax_mag.plot_data
            is_log_x = True
        elif event.inaxes == self.nyquist_ax:
            annot = self.nyquist_annot
            data_x, data_y = self.nyquist_ax.plot_data
            is_log_x = False
        else:
            if self.bode_annot.get_visible():
                self.bode_annot.set_visible(False)
                self.bode_canvas.draw_idle()
            if self.nyquist_annot.get_visible():
                self.nyquist_annot.set_visible(False)
                self.nyquist_canvas.draw_idle()
            return

        if len(data_x) == 0: return
        x, y = event.xdata, event.ydata
        if x is None or y is None: return

        if is_log_x:
            log_dist_x = (np.log10(data_x) - np.log10(x))**2
            log_dist_y = (np.log10(data_y) - np.log10(y))**2
            dist = log_dist_x + log_dist_y
        else:
            x_range = np.max(data_x) - np.min(data_x)
            y_range = np.max(data_y) - np.min(data_y)
            if x_range == 0 or y_range == 0: return
            dist = ((data_x - x)/x_range)**2 + ((data_y - y)/y_range)**2

        if len(dist) > 0:
            idx = np.argmin(dist)
            point_x, point_y = data_x[idx], data_y[idx]
            
            self.update_annotation(annot, event.inaxes, point_x, point_y)
            event.canvas.draw_idle()
        else:
            if annot.get_visible():
                annot.set_visible(False)
                event.canvas.draw_idle()

    # --- Measurement Logic (Replaced with Load Logic) ---

    def log_message(self, msg):
        def _log():
            self.output_text.configure(state="normal")
            self.output_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {msg}\n")
            self.output_text.see(tk.END)
            self.output_text.configure(state="disabled")
        self.root.after(0, _log)

    def _safe_set_shared_progress_text(self, text):
        """Safely update shared progress label text if the widget still exists."""
        try:
            if self.shared_progress_label is not None and self.shared_progress_label.winfo_exists():
                self.shared_progress_label.config(text=text)
        except Exception:
            pass

    def _destroy_shared_progress_ui(self):
        """Safely destroy shared progress UI widgets and clear references."""
        try:
            if self.shared_progress_frame is not None and self.shared_progress_frame.winfo_exists():
                self.shared_progress_frame.destroy()
        except Exception:
            pass
        self.shared_progress_frame = None
        self.shared_progress = None
        self.shared_progress_label = None
    
    # --- REMOVED get_params_from_gui ---
    # --- REMOVED generate_fake_data ---

    def start_load_data_thread(self):
        """
        Runs the file dialog on the main thread, then starts the
        processing in a new thread.
        """
        # Deprecated: replaced by Run Test streaming. Keep for backward compatibility.
        self.log_message("Use 'Run Test' to stream data from the built-in CSV.")

    def process_data_file(self, filepath):
        """
        Reads and processes the CSV file in a thread.
        """
        try:
            self.log_message(f"Loading data from {filepath}...")
            
            # Read the CSV file using pandas
            df = pd.read_csv(filepath)
            
            self.log_message("File loaded. Processing data...")

            # Check for the columns from the user's screenshot
            required_cols = {'Frequency (Hz)', "Z' (Ω)", "-Z'' (Ω)", "Z (Ω)", "-Phase (°)", "Time (s)"}
            if not required_cols.issubset(df.columns):
                self.log_message("ERROR: CSV file is missing required columns.")
                self.log_message(f"  Missing: {required_cols - set(df.columns)}")
                self.log_message("  Required: " + ", ".join(required_cols))
                return

            # Extract data into numpy arrays
            freq = df['Frequency (Hz)'].to_numpy()
            z_real = df["Z' (Ω)"].to_numpy()
            z_imag_neg = df["-Z'' (Ω)"].to_numpy()
            z_imag = -z_imag_neg # Negate the -Z'' column to get Z''
            z_mag = df["Z (Ω)"].to_numpy()
            
            # Create the data dictionary for plotting
            data = {
                'frequency': freq,
                'z_real': z_real,
                'z_imag': z_imag # Pass the calculated Z_imag
            }
            
            self.log_message(f"Acquired {len(data['frequency'])} data points.")
            
            # --- Run Diagnosis ---
            # Pass the magnitude and frequency data directly
            diagnosis_result = self.diagnose_coating(z_mag, data['frequency'])
            self.log_message(f"Diagnosis: {diagnosis_result}")
            self.report_bode_data_quality(data['frequency'], z_mag)

            # --- Draw full Plots (on main thread) ---
            self.root.after(0, self.draw_plots, data)

        except Exception as e:
            self.log_message(f"Error processing file: {e}")
        finally:
            # Re-enable the run test button if this legacy path was used
            try:
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            except Exception:
                pass

    def diagnose_coating(self, z_mag_data, freq_data):
        """Analyzes impedance data to provide a coating diagnosis."""
        try:
            # Data from CSV is high-to-low freq, so lowest freq is at the end.
            low_freq_z_mag = z_mag_data[-1]
            
            self.log_message(f"Diagnosis based on low freq impedance: {low_freq_z_mag:.2e} Ohm")

            if low_freq_z_mag >= 1e7:
                return "Healthy Coating (Pass)"
            elif low_freq_z_mag >= 1e5:
                return "Coating needs monitoring (Caution)"
            else:
                return "Defective Coating, needs maintenance (Fail)"
        except Exception as e:
            self.log_message(f"Diagnosis error: {e}")
            return "Could not determine diagnosis."

    def _get_simulated_reference_profile(self):
        """Load and cache reference Bode profile from built-in simulated CSV."""
        if self.sim_reference_profile is not None:
            return self.sim_reference_profile

        ref_path = os.path.join(os.path.dirname(__file__), "11_12_25_test5.csv")
        if not os.path.exists(ref_path):
            self.log_message("Data quality check: reference simulated CSV not found; skipping reference match.")
            return None

        try:
            df = pd.read_csv(ref_path)
            required_cols = {'Frequency (Hz)', 'Z (Ω)'}
            if not required_cols.issubset(df.columns):
                self.log_message("Data quality check: reference CSV missing Frequency/Z columns.")
                return None

            freq = pd.to_numeric(df['Frequency (Hz)'], errors='coerce').to_numpy()
            z_mag = pd.to_numeric(df['Z (Ω)'], errors='coerce').to_numpy()

            mask = np.isfinite(freq) & np.isfinite(z_mag) & (freq > 0) & (z_mag > 0)
            if np.count_nonzero(mask) < 8:
                self.log_message("Data quality check: reference CSV has insufficient valid points.")
                return None

            freq_ref = np.asarray(freq[mask], dtype=float)
            z_ref = np.asarray(z_mag[mask], dtype=float)
            order = np.argsort(freq_ref)
            freq_ref = freq_ref[order]
            z_ref = z_ref[order]

            self.sim_reference_profile = {
                "freq": freq_ref,
                "zmag": z_ref,
                "logf": np.log10(freq_ref),
                "logz": np.log10(z_ref),
            }
            return self.sim_reference_profile
        except Exception as e:
            self.log_message(f"Data quality check: failed to load reference profile ({e}).")
            return None

    def assess_bode_data_quality(self, freq_data, z_mag_data):
        """Assess Bode data cleanliness via curve fit, smoothness, and reference matching."""
        result = {
            "ok": True,
            "summary": "Data quality: clean.",
            "warnings": [],
            "metrics": {},
        }

        try:
            freq = np.asarray(freq_data, dtype=float)
            z_mag = np.asarray(z_mag_data, dtype=float)

            mask = np.isfinite(freq) & np.isfinite(z_mag) & (freq > 0) & (z_mag > 0)
            if np.count_nonzero(mask) < 8:
                result["ok"] = False
                result["warnings"].append("Not enough valid Bode points for quality check.")
                result["summary"] = "Data quality warning: insufficient valid points."
                return result

            freq = freq[mask]
            z_mag = z_mag[mask]
            order = np.argsort(freq)
            freq = freq[order]
            z_mag = z_mag[order]

            logf = np.log10(freq)
            logz = np.log10(z_mag)

            fit_degree = 3 if len(logf) >= 12 else 2
            coeff = np.polyfit(logf, logz, deg=fit_degree)
            fit_vals = np.polyval(coeff, logf)
            residuals = logz - fit_vals
            rmse_fit = float(np.sqrt(np.mean(residuals ** 2)))
            ss_res = float(np.sum((logz - fit_vals) ** 2))
            ss_tot = float(np.sum((logz - np.mean(logz)) ** 2))
            r2_fit = 1.0 - (ss_res / ss_tot) if ss_tot > 0 else 1.0

            result["metrics"]["fit_rmse_log10"] = rmse_fit
            result["metrics"]["fit_r2"] = r2_fit

            if len(logf) >= 10:
                grad1 = np.gradient(logz, logf)
                grad2 = np.gradient(grad1, logf)
                roughness = float(np.median(np.abs(grad2)))
            else:
                roughness = 0.0
            result["metrics"]["roughness"] = roughness

            ref = self._get_simulated_reference_profile()
            if ref is not None:
                common_min = max(float(np.min(logf)), float(np.min(ref["logf"])))
                common_max = min(float(np.max(logf)), float(np.max(ref["logf"])))
                common_mask = (logf >= common_min) & (logf <= common_max)

                if np.count_nonzero(common_mask) >= 6:
                    test_logf = logf[common_mask]
                    test_logz = logz[common_mask]
                    ref_logz = np.interp(test_logf, ref["logf"], ref["logz"])
                    delta = test_logz - ref_logz
                    result["metrics"]["reference_mae_log10"] = float(np.mean(np.abs(delta)))
                    result["metrics"]["reference_max_log10"] = float(np.max(np.abs(delta)))
                else:
                    result["metrics"]["reference_mae_log10"] = None
                    result["metrics"]["reference_max_log10"] = None

            if rmse_fit > 0.12:
                result["warnings"].append(f"Curve-fit residual is high (RMSE={rmse_fit:.3f} decades).")
            if r2_fit < 0.94:
                result["warnings"].append(f"Bode trend fit is weak (R²={r2_fit:.3f}).")
            if roughness > 0.45:
                result["warnings"].append(f"Curve roughness is high ({roughness:.3f}).")

            mae_ref = result["metrics"].get("reference_mae_log10")
            max_ref = result["metrics"].get("reference_max_log10")
            if mae_ref is not None and mae_ref > 0.22:
                result["warnings"].append(f"Average deviation vs simulated reference is high ({mae_ref:.3f} decades).")
            if max_ref is not None and max_ref > 0.55:
                result["warnings"].append(f"Peak deviation vs simulated reference is high ({max_ref:.3f} decades).")

            if result["warnings"]:
                result["ok"] = False
                result["summary"] = "Data quality warning: Bode data looks noisy/mismatched. Check setup."
            return result
        except Exception as e:
            result["ok"] = False
            result["warnings"] = [f"Quality check failed: {e}"]
            result["summary"] = "Data quality warning: could not complete quality check."
            return result

    def report_bode_data_quality(self, freq_data, z_mag_data):
        """Run quality check and publish warning/details to output log and status label."""
        quality = self.assess_bode_data_quality(freq_data, z_mag_data)
        self.log_message(quality["summary"])

        if quality["ok"]:
            self.root.after(0, self.show_data_quality_on_plots, None)
            return

        for warning in quality.get("warnings", []):
            self.log_message(f"Quality detail: {warning}")

        self.root.after(0, self.load_progress_lbl.config, {"text": "Warning: data may be messy — check setup"})
        warning_text = "Faulty data detected: check wiring, connections, and setup"
        self.root.after(0, self.show_data_quality_on_plots, warning_text)
        self.root.after(0, self._show_quality_warning_popup, quality)

    def _show_quality_warning_popup(self, quality):
        """Show a popup dialog for faulty data quality results."""
        try:
            warning_lines = quality.get("warnings", [])[:4]
            details = "\n".join(f"- {line}" for line in warning_lines)
            if not details:
                details = "- Bode data appears noisy or mismatched."
            messagebox.showwarning(
                "Faulty Data Warning",
                "Bode data quality check failed.\n\n"
                "Please check electrical contacts, leads, and setup before trusting this run.\n\n"
                f"Details:\n{details}",
            )
        except Exception as e:
            self.log_message(f"Warning popup failed: {e}")

    def _show_calibration_result_popup(self, completed):
        """Show popup dialog when calibration is completed or stopped."""
        try:
            if completed:
                messagebox.showinfo(
                    "Calibration Complete",
                    "Calibration sequence finished successfully (3/3).",
                )
            else:
                messagebox.showwarning(
                    "Calibration Stopped",
                    "Calibration sequence was stopped before completion.",
                )
        except Exception as e:
            self.log_message(f"Calibration popup failed: {e}")

    def set_export_buttons_enabled(self, enabled):
        """Enable/disable plot export buttons together."""
        state = "normal" if enabled else "disabled"
        for attr in (
            "export_nyquist_btn",
            "export_nyquist_cloud_btn",
            "export_bode_btn",
            "export_bode_cloud_btn",
        ):
            btn = getattr(self, attr, None)
            if btn is not None:
                try:
                    btn.config(state=state)
                except Exception:
                    pass

    def show_data_quality_on_plots(self, warning_text):
        """Display or clear a data-quality warning banner on Nyquist and Bode plots."""
        try:
            if warning_text:
                badge_style = dict(boxstyle='round', facecolor=self.theme['diag_fail'], alpha=0.92)

                if self.nyquist_quality_text is None:
                    self.nyquist_quality_text = self.nyquist_ax.text(
                        0.5, 0.04, warning_text,
                        transform=self.nyquist_ax.transAxes,
                        fontsize=11, fontweight='bold', color='white',
                        horizontalalignment='center', verticalalignment='bottom',
                        bbox=badge_style,
                    )
                else:
                    self.nyquist_quality_text.set_text(warning_text)
                    self.nyquist_quality_text.set_bbox(badge_style)
                    self.nyquist_quality_text.set_visible(True)

                if self.bode_quality_text is None:
                    self.bode_quality_text = self.bode_ax_mag.text(
                        0.98, 0.96, warning_text,
                        transform=self.bode_ax_mag.transAxes,
                        fontsize=11, fontweight='bold', color='white',
                        horizontalalignment='right', verticalalignment='top',
                        bbox=badge_style,
                    )
                else:
                    self.bode_quality_text.set_text(warning_text)
                    self.bode_quality_text.set_bbox(badge_style)
                    self.bode_quality_text.set_visible(True)
            else:
                if self.nyquist_quality_text is not None:
                    self.nyquist_quality_text.set_visible(False)
                if self.bode_quality_text is not None:
                    self.bode_quality_text.set_visible(False)

            try:
                self.nyquist_canvas.draw_idle()
            except Exception:
                pass
            try:
                self.bode_canvas.draw_idle()
            except Exception:
                pass
        except Exception as e:
            self.log_message(f"Failed to update quality warning on plots: {e}")

    def show_calibration_status_on_plots(self, status_text):
        """Display or clear a prominent calibration-status banner on both plots."""
        try:
            if status_text:
                badge_style = dict(boxstyle='round', facecolor=self.theme['warning'], alpha=0.95)

                if self.nyquist_calibration_text is None:
                    self.nyquist_calibration_text = self.nyquist_ax.text(
                        0.5, 0.90, status_text,
                        transform=self.nyquist_ax.transAxes,
                        fontsize=14, fontweight='bold', color='black',
                        horizontalalignment='center', verticalalignment='top',
                        bbox=badge_style,
                    )
                else:
                    self.nyquist_calibration_text.set_text(status_text)
                    self.nyquist_calibration_text.set_bbox(badge_style)
                    self.nyquist_calibration_text.set_visible(True)

                if self.bode_calibration_text is None:
                    self.bode_calibration_text = self.bode_ax_mag.text(
                        0.5, 0.90, status_text,
                        transform=self.bode_ax_mag.transAxes,
                        fontsize=14, fontweight='bold', color='black',
                        horizontalalignment='center', verticalalignment='top',
                        bbox=badge_style,
                    )
                else:
                    self.bode_calibration_text.set_text(status_text)
                    self.bode_calibration_text.set_bbox(badge_style)
                    self.bode_calibration_text.set_visible(True)
            else:
                if self.nyquist_calibration_text is not None:
                    self.nyquist_calibration_text.set_visible(False)
                if self.bode_calibration_text is not None:
                    self.bode_calibration_text.set_visible(False)

            try:
                self.nyquist_canvas.draw_idle()
            except Exception:
                pass
            try:
                self.bode_canvas.draw_idle()
            except Exception:
                pass
        except Exception as e:
            self.log_message(f"Failed to update calibration status on plots: {e}")

    def clear_bode_threshold_indicator(self):
        """Remove Bode threshold-evaluation indicator artists if present."""
        for attr_name in ("bode_eval_vline", "bode_eval_hline", "bode_eval_point", "bode_eval_text"):
            artist = getattr(self, attr_name, None)
            if artist is None:
                continue
            try:
                artist.remove()
            except Exception:
                try:
                    artist.set_visible(False)
                except Exception:
                    pass
            setattr(self, attr_name, None)

    def show_bode_threshold_indicator(self, freq_data, z_mag_data):
        """Show which lowest-frequency Bode point is used for threshold evaluation."""
        try:
            freq = np.asarray(freq_data, dtype=float)
            z_mag = np.asarray(z_mag_data, dtype=float)
            valid = np.isfinite(freq) & np.isfinite(z_mag) & (freq > 0) & (z_mag > 0)
            if np.count_nonzero(valid) == 0:
                return

            freq = freq[valid]
            z_mag = z_mag[valid]
            idx = int(np.argmin(freq))
            eval_freq = float(freq[idx])
            eval_z = float(z_mag[idx])

            self.clear_bode_threshold_indicator()

            self.bode_eval_vline = self.bode_ax_mag.axvline(
                eval_freq,
                color=self.theme["warning"],
                linestyle='--',
                linewidth=1.6,
                alpha=0.95,
                zorder=20,
            )
            self.bode_eval_hline = self.bode_ax_mag.axhline(
                eval_z,
                color=self.theme["warning"],
                linestyle='--',
                linewidth=1.4,
                alpha=0.9,
                zorder=20,
            )
            (self.bode_eval_point,) = self.bode_ax_mag.plot(
                [eval_freq],
                [eval_z],
                marker='s',
                markersize=7,
                markerfacecolor=self.theme["warning"],
                markeredgecolor='black',
                linestyle='None',
                zorder=25,
            )
            self.bode_eval_text = self.bode_ax_mag.text(
                0.98,
                0.02,
                f"Threshold evaluation point\nlowest f = {eval_freq:.2e} Hz, |Z| = {eval_z:.2e} Ω",
                transform=self.bode_ax_mag.transAxes,
                horizontalalignment='right',
                verticalalignment='bottom',
                fontsize=9,
                color=self.theme["text"],
                bbox=dict(
                    boxstyle='round',
                    facecolor=self.theme["panel_alt"],
                    edgecolor=self.theme["warning"],
                    alpha=0.95,
                ),
                zorder=30,
            )
            self.bode_canvas.draw_idle()
        except Exception as e:
            self.log_message(f"Failed to draw Bode threshold indicator: {e}")

    def clear_plot_status_overlays(self):
        """Hide previous run's diagnosis and quality warnings from both plots."""
        try:
            if self.nyquist_diag_text is not None:
                try:
                    self.nyquist_diag_text.set_visible(False)
                except Exception:
                    pass
            if self.bode_diag_text is not None:
                try:
                    self.bode_diag_text.set_visible(False)
                except Exception:
                    pass

            if self.nyquist_quality_text is not None:
                try:
                    self.nyquist_quality_text.set_visible(False)
                except Exception:
                    pass
            if self.bode_quality_text is not None:
                try:
                    self.bode_quality_text.set_visible(False)
                except Exception:
                    pass

            if self.nyquist_calibration_text is not None:
                try:
                    self.nyquist_calibration_text.set_visible(False)
                except Exception:
                    pass
            if self.bode_calibration_text is not None:
                try:
                    self.bode_calibration_text.set_visible(False)
                except Exception:
                    pass

            self.clear_bode_threshold_indicator()

            try:
                self.nyquist_canvas.draw_idle()
            except Exception:
                pass
            try:
                self.bode_canvas.draw_idle()
            except Exception:
                pass
        except Exception as e:
            self.log_message(f"Failed to clear plot status overlays: {e}")

    def log_test_separator(self):
        """Write a clear separator in the output log at the start of each run."""
        self.test_run_counter += 1
        mode_label = {
            "sensit_bt": "Sensit BT",
            "sensit_usb": "Sensit USB",
            "simulated": "Simulated Mode",
            "messy": "Messy Data",
            "calibration": "Calibration",
        }.get(self.connection_mode, "Unknown Mode")

        separator = "=" * 66
        self.log_message(separator)
        self.log_message(f"TEST RUN {self.test_run_counter} STARTED • Mode: {mode_label}")
        self.log_message(separator)

    def log_calibration_stage_separator(self, test_index):
        """Write a clear separator for each calibration stage in the output log."""
        separator = "-" * 66
        if test_index < 3:
            stage_text = f"CALIBRATION STAGE {test_index}/3"
        else:
            stage_text = "FINAL MEASUREMENT STAGE 3/3"
        self.log_message(separator)
        self.log_message(stage_text)
        self.log_message(separator)
    
    def draw_plots(self, data):
        try:
            freq, z_real, z_imag = data['frequency'], data['z_real'], data['z_imag']
            self.latest_plot_data = {
                'frequency': np.asarray(freq),
                'z_real': np.asarray(z_real),
                'z_imag': np.asarray(z_imag),
            }
            z_mag = np.sqrt(z_real**2 + z_imag**2)
            z_imag_neg = -z_imag
            
            # --- 1. Nyquist Plot ---
            self.init_nyquist_plot() # Clears, sets formatters
            self.nyquist_ax.plot(z_real, z_imag_neg, 'o-', markersize=4, color=self.theme["accent"])
            self.nyquist_ax.plot_data = (z_real, z_imag_neg)
            self.nyquist_ax.relim()
            self.nyquist_ax.autoscale_view()
            
            # --- NEW: Add axis('equal') back ---
            # This makes the plot scales 1:1, just like MATLAB
            self.nyquist_ax.axis('equal') 
            
            self.nyquist_canvas.draw()

            # --- 2. Bode Plot ---
            self.init_bode_plot() 
            self.bode_ax_mag.loglog(freq, z_mag, 'o-', markersize=4, color=self.theme["accent"], zorder=10)
            self.bode_ax_mag.plot_data = (freq, z_mag)
            self.show_bode_threshold_indicator(freq, z_mag)
            self.bode_canvas.draw()
            
            self.log_message("Plots updated.")
            self.notebook.select(self.notebook.tabs()[2]) # Switch to Bode plot

        except Exception as e:
            self.log_message(f"Failed to draw plots: {e}")

    # --- New: Streaming load for Run Test ---
    def start_run_test_thread(self):
        """Starts a real EIS measurement using the connected PalmSens instrument."""
        if self.connection_mode is None:
            self.log_message("ERROR: No device connected. Click Connect first.")
            return
        if self.connection_mode in ("sensit_bt", "sensit_usb") and self.ps_manager is None:
            self.log_message("ERROR: No PalmSens instrument connected. Click Connect first.")
            return
        if self.measurement_in_progress:
            self.log_message("Measurement already in progress.")
            return

        self.measurement_in_progress = True
        self.stop_requested = False
        self.measurement_start_time = time.time()
        self.last_point_time = self.measurement_start_time
        self.last_point_count = 0
        self.set_export_buttons_enabled(False)
        self.log_test_separator()
        # Clear prior run annotations before starting a new test.
        self.clear_plot_status_overlays()
        # Disable button and switch to log
        self.run_test_btn.config(state="disabled")
        self.calibrate_btn.config(state="disabled")
        self.stop_test_btn.config(state="normal")
        # Show the Bode plot while streaming
        try:
            self.notebook.select(self.notebook.tabs()[2])
        except Exception:
            # fallback to last tab if indexing fails
            self.notebook.select(self.notebook.tabs()[-1])

        # Create shared progress bar in the bode tab above the save button
        try:
            bode_tab_widget = self.notebook.nametowidget(self.notebook.tabs()[2])
            # If a previous shared widget exists, remove it
            self._destroy_shared_progress_ui()

            self.shared_progress_frame = ttk.Frame(bode_tab_widget, style="Card.TFrame")
            # pack it before the save button so it appears above
            self.shared_progress = ttk.Progressbar(self.shared_progress_frame, style='Green.Horizontal.TProgressbar', orient='horizontal', mode='determinate', variable=self.progress_var, maximum=100)
            self.shared_progress.pack(side='left', fill='x', expand=True, padx=(0,8))
            self.shared_progress_label = ttk.Label(self.shared_progress_frame, text='0%', style="Card.TLabel")
            self.shared_progress_label.pack(side='right')
            # Insert above the save button
            try:
                self.shared_progress_frame.pack(side='bottom', fill='x', padx=10, pady=(6,2), before=self.export_bode_btn)
            except Exception:
                # fallback: pack then repack save button below
                self.shared_progress_frame.pack(side='bottom', fill='x', padx=10, pady=(6,2))
                try:
                    self.export_bode_btn.pack_forget()
                    self.export_bode_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0,0))
                except Exception:
                    pass
        except Exception:
            # If any failure, just continue without visual shared bar
            pass

        self.root.after(0, self.progress_var.set, 0.0)
        if self.connection_mode == "simulated":
            self.load_progress_lbl.config(text="Starting simulated test...")
        elif self.connection_mode == "messy":
            self.load_progress_lbl.config(text="Starting messy data simulation...")
        elif self.connection_mode == "calibration":
            self.load_progress_lbl.config(text="Starting calibration sequence (3 tests)...")
        else:
            self.load_progress_lbl.config(text="Starting EIS measurement...")
        self.root.after(5000, self.measurement_watchdog_tick)
        if self.connection_mode == "simulated":
            csv_path = os.path.join(os.path.dirname(__file__), "11_12_25_test5.csv")
            threading.Thread(target=self.stream_load_data, args=(csv_path,), daemon=True).start()
        elif self.connection_mode == "messy":
            csv_path = os.path.join(os.path.dirname(__file__), "11_12_25_test5.csv")
            threading.Thread(target=self.stream_load_data_messy, args=(csv_path,), daemon=True).start()
        elif self.connection_mode == "calibration":
            csv_path = os.path.join(os.path.dirname(__file__), "11_12_25_test5.csv")
            threading.Thread(target=self.run_calibration_sequence, args=(csv_path,), daemon=True).start()
        else:
            threading.Thread(target=self.run_real_eis_measurement, daemon=True).start()

    def request_stop_measurement(self):
        """Request cancellation of the active EIS measurement."""
        if not self.measurement_in_progress:
            self.log_message("No active measurement to stop.")
            return

        self.stop_requested = True
        self.stop_test_btn.config(state="disabled")
        self.load_progress_lbl.config(text="Stopping measurement...")
        if self.connection_mode in ("simulated", "messy", "calibration"):
            self.log_message("Stop requested for simulated test.")
        else:
            self.log_message("Stop requested. Sending stop signal to potentiostat...")
            threading.Thread(target=self._send_stop_signal_to_instrument, daemon=True).start()

    def start_calibration_thread(self):
        """Run 3 real calibration tests on connected Sensit BT instrument."""
        if self.connection_mode not in ("sensit_bt", "sensit_usb") or self.ps_manager is None:
            self.log_message("ERROR: Calibration requires an active PalmSens connection.")
            return
        if self.measurement_in_progress:
            self.log_message("Measurement already in progress.")
            return

        self.measurement_in_progress = True
        self.stop_requested = False
        self.measurement_start_time = time.time()
        self.last_point_time = self.measurement_start_time
        self.last_point_count = 0
        self.set_export_buttons_enabled(False)
        self.log_test_separator()
        self.clear_plot_status_overlays()

        self.run_test_btn.config(state="disabled")
        self.calibrate_btn.config(state="disabled")
        self.stop_test_btn.config(state="normal")

        try:
            self.notebook.select(self.notebook.tabs()[2])
        except Exception:
            self.notebook.select(self.notebook.tabs()[-1])

        try:
            bode_tab_widget = self.notebook.nametowidget(self.notebook.tabs()[2])
            self._destroy_shared_progress_ui()
            self.shared_progress_frame = ttk.Frame(bode_tab_widget, style="Card.TFrame")
            self.shared_progress = ttk.Progressbar(
                self.shared_progress_frame,
                style='Green.Horizontal.TProgressbar',
                orient='horizontal',
                mode='determinate',
                variable=self.progress_var,
                maximum=100,
            )
            self.shared_progress.pack(side='left', fill='x', expand=True, padx=(0, 8))
            self.shared_progress_label = ttk.Label(self.shared_progress_frame, text='0%', style="Card.TLabel")
            self.shared_progress_label.pack(side='right')
            try:
                self.shared_progress_frame.pack(side='bottom', fill='x', padx=10, pady=(6, 2), before=self.export_bode_btn)
            except Exception:
                self.shared_progress_frame.pack(side='bottom', fill='x', padx=10, pady=(6, 2))
        except Exception:
            pass

        self.root.after(0, self.progress_var.set, 0.0)
        self.load_progress_lbl.config(text="Starting real calibration sequence (3 tests)...")
        self.root.after(5000, self.measurement_watchdog_tick)
        threading.Thread(target=self.run_real_calibration_sequence, daemon=True).start()

    def run_real_calibration_sequence(self):
        """Execute 3 real EIS tests; first two are calibration passes, third is final."""
        try:
            for test_index in range(1, 4):
                if self.stop_requested:
                    self.log_message("Calibration sequence stopped by user.")
                    break

                self.log_calibration_stage_separator(test_index)
                self.root.after(0, self.clear_plot_status_overlays)
                if test_index < 3:
                    self.log_message(f"Calibration: Test {test_index}/3")
                    self.root.after(0, self.load_progress_lbl.config, {"text": f"Calibration: Test {test_index}/3"})
                    self.root.after(0, self.show_calibration_status_on_plots, f"CALIBRATING: TEST {test_index}/3")
                else:
                    self.log_message("Calibration complete. Running final test (3/3)...")
                    self.root.after(0, self.load_progress_lbl.config, {"text": "Calibration complete. Running final test (3/3)"})
                    self.root.after(0, self.show_calibration_status_on_plots, "FINAL TEST: 3/3")

                self.root.after(0, self.progress_var.set, 0.0)
                self.root.after(0, self._safe_set_shared_progress_text, "0%")

                self.run_real_eis_measurement(
                    manage_lifecycle=False,
                    is_calibration_stage=True,
                    calibration_stage=test_index,
                    calibration_total=3,
                    final_calibration_stage=(test_index == 3),
                )

                if self.stop_requested:
                    break

                if test_index < 3:
                    self.log_message(f"Calibration: Test {test_index}/3 complete.")
                    time.sleep(0.6)

            self.root.after(0, self.show_calibration_status_on_plots, None)
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            if self.connection_mode in ("sensit_bt", "sensit_usb") and self.ps_manager is not None:
                self.root.after(0, self.calibrate_btn.config, {"state": "normal"})
            else:
                self.root.after(0, self.calibrate_btn.config, {"state": "disabled"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})

            if self.stop_requested:
                self.root.after(0, self.load_progress_lbl.config, {"text": "Calibration sequence stopped"})
                self.root.after(0, self._safe_set_shared_progress_text, "Stopped")
                self.root.after(0, self._show_calibration_result_popup, False)
            else:
                self.root.after(0, self.load_progress_lbl.config, {"text": "Calibration sequence finished"})
                self.root.after(0, self.progress_var.set, 100.0)
                self.root.after(0, self._safe_set_shared_progress_text, "100%")
                self.root.after(0, self._show_calibration_result_popup, True)
                self.root.after(0, self.set_export_buttons_enabled, True)

            self.measurement_in_progress = False
            self.stop_requested = False
        except Exception as e:
            self.log_message(f"Real calibration sequence failed: {e}")
            self.root.after(0, self.show_calibration_status_on_plots, None)
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            if self.connection_mode in ("sensit_bt", "sensit_usb") and self.ps_manager is not None:
                self.root.after(0, self.calibrate_btn.config, {"state": "normal"})
            else:
                self.root.after(0, self.calibrate_btn.config, {"state": "disabled"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            self.root.after(0, self.progress_var.set, 0.0)
            self.root.after(0, self._safe_set_shared_progress_text, "0%")
            self.measurement_in_progress = False
            self.stop_requested = False

    def run_calibration_sequence(self, filepath):
        """Run simulated test three times; first two runs are calibration passes."""
        try:
            self.log_message(f"Starting calibration sequence using {filepath}")
            if not os.path.exists(filepath):
                self.log_message("ERROR: CSV file not found in project directory.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
                self.measurement_in_progress = False
                self.stop_requested = False
                return

            df = pd.read_csv(filepath)
            required_cols = {'Frequency (Hz)', "Z' (Ω)", "-Z'' (Ω)", "Z (Ω)", "-Phase (°)", "Time (s)"}
            if not required_cols.issubset(df.columns):
                self.log_message("ERROR: CSV file is missing required columns for calibration mode.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
                self.measurement_in_progress = False
                self.stop_requested = False
                return

            freq = pd.to_numeric(df['Frequency (Hz)'], errors='coerce').to_numpy(dtype=float)
            z_real = pd.to_numeric(df["Z' (Ω)"], errors='coerce').to_numpy(dtype=float)
            z_imag_neg = pd.to_numeric(df["-Z'' (Ω)"], errors='coerce').to_numpy(dtype=float)
            z_imag = -z_imag_neg

            valid = np.isfinite(freq) & np.isfinite(z_real) & np.isfinite(z_imag) & (freq > 0)
            freq = freq[valid]
            z_real = z_real[valid]
            z_imag = z_imag[valid]

            if len(freq) < 8:
                self.log_message("ERROR: Not enough valid points for calibration mode.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
                self.measurement_in_progress = False
                self.stop_requested = False
                return

            # Preserve simulated sweep direction (high -> low frequency)
            order = np.argsort(freq)[::-1]
            freq = freq[order]
            z_real = z_real[order]
            z_imag = z_imag[order]

            n = len(freq)
            total_time_per_test = 8.0
            interval = total_time_per_test / max(n, 1)

            for test_index in range(1, 4):
                if self.stop_requested:
                    self.log_message("Calibration sequence stopped by user.")
                    break

                self.log_calibration_stage_separator(test_index)
                self.root.after(0, self.clear_plot_status_overlays)

                if test_index < 3:
                    self.log_message(f"Calibration: Test {test_index}/3")
                    self.root.after(0, self.load_progress_lbl.config, {"text": f"Calibration: Test {test_index}/3"})
                    self.root.after(0, self.show_calibration_status_on_plots, f"CALIBRATING: TEST {test_index}/3")
                else:
                    self.log_message("Calibration complete. Running final test (3/3)...")
                    self.root.after(0, self.load_progress_lbl.config, {"text": "Calibration complete. Running final test (3/3)"})
                    self.root.after(0, self.show_calibration_status_on_plots, "FINAL TEST: 3/3")

                self.root.after(0, self.progress_var.set, 0.0)
                self.root.after(0, self._safe_set_shared_progress_text, "0%")

                x_buf, yr_buf, yi_buf = [], [], []
                for i in range(n):
                    if self.stop_requested:
                        self.log_message("Calibration sequence stopped by user.")
                        break

                    x_buf.append(freq[i])
                    yr_buf.append(z_real[i])
                    yi_buf.append(z_imag[i])

                    percent = (i + 1) / n * 100.0
                    if test_index < 3:
                        self.root.after(0, self.load_progress_lbl.config, {"text": f"Calibration: Test {test_index}/3 • {i+1}/{n} points"})
                    else:
                        self.root.after(0, self.load_progress_lbl.config, {"text": f"Final Test: 3/3 • {i+1}/{n} points"})
                    self.root.after(0, self.progress_var.set, percent)
                    self.root.after(0, self._safe_set_shared_progress_text, f"{percent:.0f}%")
                    self.root.after(0, self.update_plots_incremental, np.array(x_buf), np.array(yr_buf), np.array(yi_buf))

                    time.sleep(interval)

                if self.stop_requested:
                    break

                if test_index < 3:
                    self.log_message(f"Calibration: Test {test_index}/3 complete.")
                    self.root.after(0, self.progress_var.set, 100.0)
                    self.root.after(0, self._safe_set_shared_progress_text, "100%")
                    time.sleep(0.6)
                    continue

                # Final pass: run diagnosis and quality checks.
                self.root.after(0, self.show_calibration_status_on_plots, None)
                current_freq = np.array(x_buf)
                current_z_mag = np.sqrt(np.array(yr_buf) ** 2 + np.array(yi_buf) ** 2)
                diagnosis_result = self.diagnose_coating(current_z_mag, current_freq)
                self.log_message(f"Diagnosis: {diagnosis_result}")
                self.report_bode_data_quality(current_freq, current_z_mag)
                self.root.after(0, self.show_bode_threshold_indicator, current_freq, current_z_mag)
                self.root.after(0, self.show_diagnosis_on_plots, diagnosis_result)

            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            if self.stop_requested:
                self.root.after(0, self.show_calibration_status_on_plots, None)
                self.root.after(0, self.load_progress_lbl.config, {"text": "Calibration sequence stopped"})
                self.root.after(0, self._safe_set_shared_progress_text, "Stopped")
                self.root.after(0, self._show_calibration_result_popup, False)
            else:
                self.root.after(0, self.show_calibration_status_on_plots, None)
                self.root.after(0, self.load_progress_lbl.config, {"text": "Calibration sequence finished"})
                self.root.after(0, self.progress_var.set, 100.0)
                self.root.after(0, self._safe_set_shared_progress_text, "100%")
                self.root.after(0, self._show_calibration_result_popup, True)
                self.root.after(0, self.set_export_buttons_enabled, True)

            self.measurement_in_progress = False
            self.stop_requested = False
        except Exception as e:
            self.log_message(f"Calibration sequence failed: {e}")
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            self.root.after(0, self.progress_var.set, 0.0)
            self.root.after(0, self._safe_set_shared_progress_text, "0%")
            self.measurement_in_progress = False
            self.stop_requested = False

    def _send_stop_signal_to_instrument(self):
        """Best-effort stop/abort command dispatch for different SDK versions."""
        manager = self.ps_manager
        if manager is None:
            self.log_message("Stop request: no active instrument manager.")
            return

        stop_method_names = (
            'stop_measurement',
            'abort_measurement',
            'cancel_measurement',
            'break_measurement',
            'stop',
            'abort',
        )

        for method_name in stop_method_names:
            method = getattr(manager, method_name, None)
            if not callable(method):
                continue
            try:
                method()
                self.log_message(f"Stop signal sent using manager.{method_name}().")
                return
            except TypeError:
                try:
                    method(True)
                    self.log_message(f"Stop signal sent using manager.{method_name}(True).")
                    return
                except Exception as e:
                    self.log_message(f"Stop method {method_name} failed: {e}")
            except Exception as e:
                self.log_message(f"Stop method {method_name} failed: {e}")

        self.log_message("No supported stop method found on this SDK version; you may need to disconnect to force-stop.")

    def measurement_watchdog_tick(self):
        """Periodic UI/log heartbeat while EIS measurement is running."""
        if not self.measurement_in_progress:
            return

        now = time.time()
        elapsed = now - (self.measurement_start_time or now)
        since_last = now - (self.last_point_time or now)
        count = self.last_point_count
        n = max(self.expected_points, 1)

        # Keep the status label alive even when callback is quiet.
        self.load_progress_lbl.config(
            text=f"Measuring: {count}/{n} points • elapsed {elapsed:.0f}s • last point {since_last:.0f}s ago"
        )

        # Emit a periodic log only when callback has been quiet for a while.
        if since_last >= 20:
            self.log_message(
                f"Measurement still running: {count}/{n} points, no new point for {since_last:.0f}s (low-frequency points can take longer)."
            )

        self.root.after(5000, self.measurement_watchdog_tick)

    def build_eis_method(self):
        """Build EIS method from GUI parameters."""
        if ps is None:
            raise RuntimeError("PyPalmSens is not installed")

        start_freq = float(self.param_vars["Start Frequency (Hz)"].get())
        end_freq = float(self.param_vars["End Frequency (Hz)"].get())
        amplitude_mv = float(self.param_vars["Voltage Amplitude (mV)"].get())
        points_per_decade = float(self.param_vars["Points per Decade"].get())

        if start_freq <= 0 or end_freq <= 0:
            raise ValueError("Start/End frequency must be positive")
        if start_freq <= end_freq:
            raise ValueError("Start Frequency must be greater than End Frequency")
        if points_per_decade <= 0:
            raise ValueError("Points per Decade must be > 0")

        decades = np.log10(start_freq / end_freq)
        n_frequencies = max(2, int(round(decades * points_per_decade)) + 1)
        self.expected_points = n_frequencies

        method = ps.ElectrochemicalImpedanceSpectroscopy(
            max_frequency=start_freq,
            min_frequency=end_freq,
            n_frequencies=n_frequencies,
            ac_potential=amplitude_mv / 1000.0,
            frequency_type='scan',
            scan_type='fixed',  # Fixed frequency sweep
        )
        return method

    def run_real_eis_measurement(self, manage_lifecycle=True, is_calibration_stage=False, calibration_stage=1, calibration_total=1, final_calibration_stage=True):
        """Execute EIS measurement via PyPalmSens and stream callback data into plots."""
        freq_buf, zre_buf, zim_buf = [], [], []
        seen_points = set()
        last_freq_seen = [None]
        callback_call_count = [0]
        impedance_started = [False]  # Flag: True once we get valid impedance data
        buffered_by_index = {}  # Map index -> (freq, zre, zim) for frequency-only points
        replay_queue = []  # Queue of buffered points to replay gradually
        last_replay_time = [time.time()]  # Track last replay time

        def eis_callback(data):
            try:
                callback_call_count[0] += 1
                call_num = callback_call_count[0]

                if self.stop_requested:
                    return

                points = []

                # Log raw callback structure for first few processed calls (after skip)
                if call_num <= 8:
                    self.log_message(f"[DEBUG Callback #{call_num}] data type: {type(data).__name__}, dir: {[x for x in dir(data) if not x.startswith('_')]}")

                # Primary path: batched points from SDK callback
                try:
                    points = list(data.new_datapoints())
                    if points and call_num <= 8:
                        self.log_message(f"[DEBUG Callback #{call_num}] new_datapoints() returned {len(points)} point(s)")
                        for i, pt in enumerate(points[:1]):
                            self.log_message(f"[DEBUG Callback #{call_num}] point[0] keys: {list(pt.keys()) if isinstance(pt, dict) else 'N/A'}")
                            if isinstance(pt, dict):
                                # Log ALL fields from the point
                                for k in pt.keys():
                                    v = pt.get(k)
                                    self.log_message(f"[DEBUG Callback #{call_num}]   {k}={v}")
                except Exception as e:
                    if call_num <= 8:
                        self.log_message(f"[DEBUG Callback #{call_num}] new_datapoints() failed: {e}")
                    points = []

                # Fallback path: some SDK/transport paths may only expose last point
                if not points:
                    try:
                        last = data.last_datapoint()
                        if last and call_num <= 8:
                            self.log_message(f"[DEBUG Callback #{call_num}] last_datapoint() returned: {last}")
                        if last:
                            points = [last]
                    except Exception as e:
                        if call_num <= 8:
                            self.log_message(f"[DEBUG Callback #{call_num}] last_datapoint() failed: {e}")
                        points = []

                if not points and call_num <= 8:
                    self.log_message(f"[DEBUG Callback #{call_num}] WARNING: No points extracted.")

                # Gradually replay buffered frequencies (add a few from the queue each callback)
                current_time = time.time()
                while replay_queue and (current_time - last_replay_time[0] >= 0.3):  # Add ~3-4 points per second
                    buff_freq, buff_zre, buff_zim = replay_queue.pop(0)
                    
                    # Add to buffers with sequential deduplication
                    sig = (round(buff_freq, 8), round(buff_zre, 6), round(buff_zim, 6))
                    if sig not in seen_points:
                        seen_points.add(sig)
                        freq_buf.append(buff_freq)
                        zre_buf.append(buff_zre)
                        zim_buf.append(buff_zim)
                        last_freq_seen[0] = buff_freq
                        last_replay_time[0] = current_time

                for point in points:
                    if not isinstance(point, dict):
                        continue

                    freq = float(point.get('Frequency', np.nan))
                    zre = float(point.get('ZRe', np.nan))
                    zim = float(point.get('ZIm', np.nan))
                    z_mag = float(point.get('Z', np.nan))
                    
                    # If ZRe/ZIm are NaN but Z (magnitude) is available, derive them
                    # This handles SDK batches where component data lags behind magnitude data
                    if (np.isnan(zre) or np.isnan(zim)) and not np.isnan(z_mag):
                        # Use Z as magnitude, and phase to construct components if available
                        phase = float(point.get('Phase', np.nan))  # in degrees, typically
                        if not np.isnan(phase):
                            phase_rad = np.radians(phase)
                            zre = z_mag * np.cos(phase_rad)
                            zim = -z_mag * np.sin(phase_rad)  # negative for -Z''
                        else:
                            # Fall back to using Z as real part, zero imaginary
                            zre = z_mag
                            zim = 0.0
                    
                    # Reject if no frequency
                    if np.isnan(freq):
                        if call_num <= 8:
                            self.log_message(f"[DEBUG Callback #{call_num}] Skipped: freq=NaN (no frequency data)")
                        continue
                    
                    idx = point.get('index')
                    
                    # If impedance is missing, buffer this frequency point by index for later gradual replay
                    if (np.isnan(zre) or np.isnan(zim)):
                        if idx is not None and not impedance_started[0]:
                            if call_num <= 8:
                                self.log_message(f"[DEBUG Callback #{call_num}] Buffering: freq={freq:.2e} Hz, index={idx}")
                            buffered_by_index[idx] = (freq, float('nan'), float('nan'))
                        continue
                    
                    # Impedance arrived! Start plotting and queue buffered frequencies for gradual replay
                    if not impedance_started[0]:
                        impedance_started[0] = True
                        self.log_message(f"Impedance data detected at frequency {freq:.2e} Hz (index {idx}). Queuing buffered frequencies for gradual replay...")
                        
                        # Queue up all buffered frequencies (in sorted index order) for gradual replay
                        for buff_idx in sorted(buffered_by_index.keys()):
                            buff_freq, _, _ = buffered_by_index[buff_idx]
                            # Use current impedance as estimate
                            replay_queue.append((buff_freq, zre, zim))
                        buffered_by_index.clear()

                    # Some SDK paths can reuse/reshape point indices, so dedupe by data signature.
                    if idx is not None:
                        sig = (int(idx), round(freq, 8))
                    else:
                        sig = (round(freq, 8), round(zre, 6), round(zim, 6))
                    if sig in seen_points:
                        if call_num <= 8:
                            self.log_message(f"[DEBUG Callback #{call_num}] Dedupe skipped freq={freq:.2e} Hz (already seen)")
                        continue
                    seen_points.add(sig)

                    prev_freq = last_freq_seen[0]
                    if prev_freq is not None and freq < (prev_freq / 3.0):
                        self.log_message(
                            f"Frequency jump detected: {prev_freq:.2e} -> {freq:.2e} Hz (ratio {freq/prev_freq:.3f}, possible skipped callback payload)."
                        )
                    last_freq_seen[0] = freq

                    freq_buf.append(freq)
                    zre_buf.append(zre)
                    zim_buf.append(zim)
                    self.last_point_time = time.time()

                    i = len(freq_buf)
                    self.last_point_count = i
                    n = max(self.expected_points, 1)
                    percent = min(100.0, (i / n) * 100.0)
                    self.root.after(0, self.progress_var.set, percent)
                    self.root.after(0, self.load_progress_lbl.config, {"text": f"Measuring: {i}/{n} points"})
                    self.root.after(0, self._safe_set_shared_progress_text, f"{percent:.0f}%")
                    
                    # Throttle plot updates to prevent overwhelming the UI and blocking mouse events
                    current_time = time.time() * 1000  # ms
                    if current_time - self.last_plot_update_time >= 100:  # Update max every 100ms
                        self.last_plot_update_time = current_time
                        self.root.after(0, self.update_plots_incremental, np.array(freq_buf), np.array(zre_buf), np.array(zim_buf))
            except Exception as cb_err:
                self.log_message(f"Callback error: {cb_err}")

        try:
            method = self.build_eis_method()
            if is_calibration_stage:
                self.log_message(
                    f"Running calibration stage {calibration_stage}/{calibration_total} over Bluetooth: "
                    f"fmax={method.max_frequency:.2e} Hz, fmin={method.min_frequency:.2e} Hz, "
                    f"n={method.n_frequencies}, Vac={method.ac_potential:.3f} V"
                )
            else:
                self.log_message(
                    f"Running EIS over Bluetooth: fmax={method.max_frequency:.2e} Hz, "
                    f"fmin={method.min_frequency:.2e} Hz, n={method.n_frequencies}, "
                    f"Vac={method.ac_potential:.3f} V"
                )

            measurement = self.ps_manager.measure(method, callback=eis_callback)
            self.log_message(f"Measurement finished: {measurement.title}")

            # Flush any remaining queued replay points
            for buff_freq, buff_zre, buff_zim in replay_queue:
                sig = (round(buff_freq, 8), round(buff_zre, 6), round(buff_zim, 6))
                if sig not in seen_points:
                    seen_points.add(sig)
                    freq_buf.append(buff_freq)
                    zre_buf.append(buff_zre)
                    zim_buf.append(buff_zim)
            replay_queue.clear()

            if len(freq_buf) > 0:
                z_mag = np.sqrt(np.array(zre_buf) ** 2 + np.array(zim_buf) ** 2)
                # Final plot update to show all data
                self.root.after(0, self.update_plots_incremental, np.array(freq_buf), np.array(zre_buf), np.array(zim_buf))
                if (not is_calibration_stage) or final_calibration_stage:
                    diagnosis_result = self.diagnose_coating(z_mag, np.array(freq_buf))
                    self.log_message(f"Diagnosis: {diagnosis_result}")
                    self.report_bode_data_quality(np.array(freq_buf), z_mag)
                    self.root.after(0, self.show_bode_threshold_indicator, np.array(freq_buf), z_mag)
                    self.root.after(0, self.show_diagnosis_on_plots, diagnosis_result)

            if self.stop_requested:
                self.root.after(0, self.load_progress_lbl.config, {"text": "Measurement stopped"})
                self.log_message("Measurement stopped by user.")
            else:
                self.root.after(0, self.progress_var.set, 100.0)
                self.root.after(0, self.load_progress_lbl.config, {"text": "Test finished"})
                self.root.after(0, self.set_export_buttons_enabled, True)
            self.root.after(0, self._safe_set_shared_progress_text, "100%" if not self.stop_requested else "Stopped")
        except Exception as e:
            if self.stop_requested:
                self.log_message("Measurement stop completed.")
                self.root.after(0, self.load_progress_lbl.config, {"text": "Measurement stopped"})
            else:
                self.log_message(f"Real EIS measurement failed: {e}")
                self.root.after(0, self.progress_var.set, 0.0)
                self.root.after(0, self.load_progress_lbl.config, {"text": "Measurement failed"})
            self.root.after(0, self._safe_set_shared_progress_text, "0%" if not self.stop_requested else "Stopped")
        finally:
            if manage_lifecycle:
                self.measurement_in_progress = False
                self.stop_requested = False
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                if self.connection_mode in ("sensit_bt", "sensit_usb") and self.ps_manager is not None:
                    self.root.after(0, self.calibrate_btn.config, {"state": "normal"})
                else:
                    self.root.after(0, self.calibrate_btn.config, {"state": "disabled"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})

    def stream_load_data(self, filepath):
        """Reads CSV then streams data points progressively over the sample time."""
        try:
            self.log_message(f"Starting test using {filepath}")
            if not os.path.exists(filepath):
                self.log_message("ERROR: CSV file not found in project directory.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                return

            df = pd.read_csv(filepath)

            required_cols = {'Frequency (Hz)', "Z' (Ω)", "-Z'' (Ω)", "Z (Ω)", "-Phase (°)", "Time (s)"}
            if not required_cols.issubset(df.columns):
                self.log_message("ERROR: CSV file is missing required columns for streaming test.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                return

            freq = df['Frequency (Hz)'].to_numpy()
            z_real = df["Z' (Ω)"].to_numpy()
            z_imag_neg = df["-Z'' (Ω)"].to_numpy()
            z_imag = -z_imag_neg

            n = len(freq)
            # Streaming duration is fixed at 10 seconds total
            total_time = 10.0
            interval = total_time / max(n, 1)

            # Stream points
            x_buf = []
            yr_buf = []
            yi_buf = []

            for i in range(n):
                if self.stop_requested:
                    self.log_message("Simulated measurement stopped by user.")
                    break

                x_buf.append(freq[i])
                yr_buf.append(z_real[i])
                yi_buf.append(z_imag[i])

                # Update progress label on main thread
                percent = (i+1) / n * 100.0
                self.root.after(0, self.load_progress_lbl.config, {"text": f"Loading: {i+1}/{n} points"})
                # Update progress variable
                self.root.after(0, self.progress_var.set, percent)
                # Update shared progress label if present
                self.root.after(0, self._safe_set_shared_progress_text, f"{percent:.0f}%")

                # Update plots with current subset
                self.root.after(0, self.update_plots_incremental, np.array(x_buf), np.array(yr_buf), np.array(yi_buf))

                time.sleep(interval)

            if len(x_buf) > 0:
                self.log_message("Test complete. Full data loaded." if not self.stop_requested else "Test stopped.")
                # Determine coating health based on the measured impedance magnitude
                try:
                    current_z_mag = np.sqrt(np.array(yr_buf)**2 + np.array(yi_buf)**2)
                    current_freq = np.array(x_buf)
                    diagnosis_result = self.diagnose_coating(current_z_mag, current_freq)
                    self.log_message(f"Diagnosis: {diagnosis_result}")
                    self.report_bode_data_quality(current_freq, current_z_mag)
                    self.root.after(0, self.show_bode_threshold_indicator, current_freq, current_z_mag)
                    # Show diagnosis visually on plots
                    self.root.after(0, self.show_diagnosis_on_plots, diagnosis_result)
                except Exception as e:
                    self.log_message(f"Diagnosis failed: {e}")

            # destroy shared progress UI if present
            self.root.after(0, self._destroy_shared_progress_ui)
            # Re-enable button and reset progress
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            self.root.after(0, self.load_progress_lbl.config, {"text": "Test finished" if not self.stop_requested else "Measurement stopped"})
            if not self.stop_requested:
                self.root.after(0, self.progress_var.set, 100.0)
                self.root.after(0, self.set_export_buttons_enabled, True)
            self.root.after(0, self._safe_set_shared_progress_text, "100%" if not self.stop_requested else "Stopped")
            self.measurement_in_progress = False
            self.stop_requested = False

        except Exception as e:
            self.log_message(f"Error streaming test data: {e}")
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            self.root.after(0, self.progress_var.set, 0.0)
            self.root.after(0, self._safe_set_shared_progress_text, "0%")
            self.measurement_in_progress = False
            self.stop_requested = False

    def stream_load_data_messy(self, filepath):
        """Stream a noisy/distorted variant of built-in simulated data for setup quality testing."""
        try:
            self.log_message(f"Starting messy-data test using {filepath}")
            if not os.path.exists(filepath):
                self.log_message("ERROR: CSV file not found in project directory.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
                self.measurement_in_progress = False
                self.stop_requested = False
                return

            df = pd.read_csv(filepath)
            required_cols = {'Frequency (Hz)', "Z' (Ω)", "-Z'' (Ω)", "Z (Ω)", "-Phase (°)", "Time (s)"}
            if not required_cols.issubset(df.columns):
                self.log_message("ERROR: CSV file is missing required columns for messy streaming test.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
                self.measurement_in_progress = False
                self.stop_requested = False
                return

            freq = pd.to_numeric(df['Frequency (Hz)'], errors='coerce').to_numpy(dtype=float)
            z_real_base = pd.to_numeric(df["Z' (Ω)"], errors='coerce').to_numpy(dtype=float)
            z_imag_neg_base = pd.to_numeric(df["-Z'' (Ω)"], errors='coerce').to_numpy(dtype=float)
            z_imag_base = -z_imag_neg_base

            valid = np.isfinite(freq) & np.isfinite(z_real_base) & np.isfinite(z_imag_base) & (freq > 0)
            freq = freq[valid]
            z_real_base = z_real_base[valid]
            z_imag_base = z_imag_base[valid]

            if len(freq) < 8:
                self.log_message("ERROR: Not enough valid points to generate messy data.")
                self.root.after(0, self.run_test_btn.config, {"state": "normal"})
                self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
                self.measurement_in_progress = False
                self.stop_requested = False
                return

            rng = np.random.default_rng(20260226)
            # Keep sweep direction high -> low frequency for consistency with real/simulated test.
            order = np.argsort(freq)[::-1]
            freq = freq[order]
            z_real_base = z_real_base[order]
            z_imag_base = z_imag_base[order]

            # Build stronger "bad connection" artifacts: heavy ripple, drift, broadband noise, and spikes.
            normalized = np.linspace(0.0, 1.0, len(freq))
            ripple = 0.42 * np.sin(2 * np.pi * 5.1 * normalized)
            drift = 0.55 * (normalized - 0.5)
            burst = 0.18 * np.sign(np.sin(2 * np.pi * 9.0 * normalized + 0.7))
            mag_scale = 1.0 + ripple + drift
            mag_scale = np.clip(mag_scale + burst, 0.20, 2.8)

            z_real = z_real_base * mag_scale * (1.0 + rng.normal(0.0, 0.22, size=len(freq)))
            z_imag = z_imag_base * mag_scale * (1.0 + rng.normal(0.0, 0.26, size=len(freq)))

            # Inject a handful of outlier spikes to emulate loose leads/contact issues.
            n_spikes = max(6, len(freq) // 7)
            spike_idx = rng.choice(len(freq), size=n_spikes, replace=False)
            spike_factor = rng.uniform(2.2, 5.0, size=n_spikes)
            z_real[spike_idx] *= spike_factor
            z_imag[spike_idx] *= spike_factor

            # Add occasional dip-outs (partial contact loss) for stronger mismatch.
            n_dips = max(3, len(freq) // 14)
            dip_idx = rng.choice(len(freq), size=n_dips, replace=False)
            dip_factor = rng.uniform(0.18, 0.45, size=n_dips)
            z_real[dip_idx] *= dip_factor
            z_imag[dip_idx] *= dip_factor

            # Preserve sign conventions and avoid zeros.
            z_real = np.sign(z_real_base) * np.maximum(np.abs(z_real), 1e-3)
            z_imag = np.sign(z_imag_base) * np.maximum(np.abs(z_imag), 1e-3)

            n = len(freq)
            total_time = 10.0
            interval = total_time / max(n, 1)

            x_buf, yr_buf, yi_buf = [], [], []

            for i in range(n):
                if self.stop_requested:
                    self.log_message("Messy simulated measurement stopped by user.")
                    break

                x_buf.append(freq[i])
                yr_buf.append(z_real[i])
                yi_buf.append(z_imag[i])

                percent = (i + 1) / n * 100.0
                self.root.after(0, self.load_progress_lbl.config, {"text": f"Loading messy data: {i+1}/{n} points"})
                self.root.after(0, self.progress_var.set, percent)
                self.root.after(0, self._safe_set_shared_progress_text, f"{percent:.0f}%")
                self.root.after(0, self.update_plots_incremental, np.array(x_buf), np.array(yr_buf), np.array(yi_buf))

                time.sleep(interval)

            if len(x_buf) > 0:
                self.log_message("Messy-data test complete." if not self.stop_requested else "Messy-data test stopped.")
                try:
                    current_z_mag = np.sqrt(np.array(yr_buf) ** 2 + np.array(yi_buf) ** 2)
                    current_freq = np.array(x_buf)
                    diagnosis_result = self.diagnose_coating(current_z_mag, current_freq)
                    self.log_message(f"Diagnosis: {diagnosis_result}")
                    self.report_bode_data_quality(current_freq, current_z_mag)
                    self.root.after(0, self.show_bode_threshold_indicator, current_freq, current_z_mag)
                    self.root.after(0, self.show_diagnosis_on_plots, diagnosis_result)
                except Exception as e:
                    self.log_message(f"Diagnosis failed: {e}")

            self.root.after(0, self._destroy_shared_progress_ui)
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            self.root.after(0, self.load_progress_lbl.config, {"text": "Messy-data test finished" if not self.stop_requested else "Measurement stopped"})
            if not self.stop_requested:
                self.root.after(0, self.progress_var.set, 100.0)
                self.root.after(0, self.set_export_buttons_enabled, True)
            self.root.after(0, self._safe_set_shared_progress_text, "100%" if not self.stop_requested else "Stopped")
            self.measurement_in_progress = False
            self.stop_requested = False

        except Exception as e:
            self.log_message(f"Error streaming messy test data: {e}")
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_test_btn.config, {"state": "disabled"})
            self.root.after(0, self.progress_var.set, 0.0)
            self.root.after(0, self._safe_set_shared_progress_text, "0%")
            self.measurement_in_progress = False
            self.stop_requested = False

    def update_plots_incremental(self, freq_subset, z_real_subset, z_imag_subset):
        """Update Nyquist and Bode plots with partial data during streaming."""
        try:
            self.latest_plot_data = {
                'frequency': np.asarray(freq_subset),
                'z_real': np.asarray(z_real_subset),
                'z_imag': np.asarray(z_imag_subset),
            }
            z_mag = np.sqrt(z_real_subset**2 + z_imag_subset**2)
            z_imag_neg = -z_imag_subset

            # --- Nyquist: create or update a persistent Line2D to avoid clearing the axes ---
            if self.nyquist_line is None:
                # create the initial line on existing axes (axes already initialized in init_nyquist_plot)
                (self.nyquist_line,) = self.nyquist_ax.plot(z_real_subset, z_imag_neg, 'o-', markersize=4, color=self.theme["accent"])
                self.nyquist_ax.plot_data = (z_real_subset, z_imag_neg)
                try:
                    self.nyquist_ax.axis('equal')
                except Exception:
                    pass
            else:
                self.nyquist_line.set_data(z_real_subset, z_imag_neg)
                self.nyquist_ax.plot_data = (z_real_subset, z_imag_neg)
                # update view limits without reinitializing the axes
                try:
                    self.nyquist_ax.relim()
                    self.nyquist_ax.autoscale_view()
                except Exception:
                    pass
            self.nyquist_canvas.draw_idle()

            # --- Bode: update persistent line on log-scaled axes ---
            safe_freq = np.where(freq_subset <= 0, 1e-6, freq_subset)
            if self.bode_line is None:
                (self.bode_line,) = self.bode_ax_mag.plot(safe_freq, z_mag, 'o-', markersize=4, color=self.theme["accent"], zorder=10)
                self.bode_ax_mag.plot_data = (safe_freq, z_mag)
            else:
                # update data in-place
                self.bode_line.set_xdata(safe_freq)
                self.bode_line.set_ydata(z_mag)
                self.bode_ax_mag.plot_data = (safe_freq, z_mag)
                try:
                    self.bode_ax_mag.relim()
                    self.bode_ax_mag.autoscale_view()
                except Exception:
                    pass
            self.bode_canvas.draw_idle()

        except Exception as e:
            self.log_message(f"Plot update failed during streaming: {e}")

    def show_diagnosis_on_plots(self, diagnosis_text):
        """Display the diagnosis on both Nyquist and Bode plots with a colored badge."""
        try:
            # Decide color based on keywords
            txt = diagnosis_text.lower()
            if 'healthy' in txt or 'pass' in txt:
                face = self.theme['diag_pass']
                fg = 'white'
            elif 'monitor' in txt or 'caution' in txt or 'medium' in txt:
                face = self.theme['diag_caution']
                fg = 'white'
            else:
                face = self.theme['diag_fail']
                fg = 'white'

            # Prepare display text (shortened)
            short = diagnosis_text

            # Nyquist: place in upper-left corner
            try:
                if self.nyquist_diag_text is None:
                    self.nyquist_diag_text = self.nyquist_ax.text(
                        0.12, 0.95, short,
                        transform=self.nyquist_ax.transAxes,
                        fontsize=12, fontweight='bold', color=fg,
                        verticalalignment='top', bbox=dict(boxstyle='round', facecolor=face, alpha=0.9)
                    )
                else:
                    self.nyquist_diag_text.set_text(short)
                    self.nyquist_diag_text.set_bbox(dict(boxstyle='round', facecolor=face, alpha=0.9))
                    self.nyquist_diag_text.set_color(fg)
            except Exception:
                pass

            # Bode: place in upper-left corner
            try:
                if self.bode_diag_text is None:
                    self.bode_diag_text = self.bode_ax_mag.text(
                        0.12, 0.95, short,
                        transform=self.bode_ax_mag.transAxes,
                        fontsize=12, fontweight='bold', color=fg,
                        verticalalignment='top', bbox=dict(boxstyle='round', facecolor=face, alpha=0.9)
                    )
                else:
                    self.bode_diag_text.set_text(short)
                    self.bode_diag_text.set_bbox(dict(boxstyle='round', facecolor=face, alpha=0.9))
                    self.bode_diag_text.set_color(fg)
            except Exception:
                pass

            # Redraw canvases
            try:
                self.nyquist_canvas.draw_idle()
            except Exception:
                pass
            try:
                self.bode_canvas.draw_idle()
            except Exception:
                pass

        except Exception as e:
            self.log_message(f"Failed to display diagnosis on plots: {e}")

    # --- Plot Export Function ---
    def _build_export_dataframe(self, plot_type):
        """Build export dataframe and metadata for Nyquist/Bode plot."""
        freq = np.asarray(self.latest_plot_data.get('frequency', np.array([])))
        z_real = np.asarray(self.latest_plot_data.get('z_real', np.array([])))
        z_imag = np.asarray(self.latest_plot_data.get('z_imag', np.array([])))

        if freq.size == 0 and z_real.size == 0:
            return None, None, None

        if plot_type == 'nyquist':
            z_imag_neg = -z_imag
            export_df = pd.DataFrame({
                "Z_real (Ohm)": z_real,
                "-Z_imaginary (Ohm)": z_imag_neg,
            })
            if freq.size == z_real.size:
                export_df.insert(0, "Frequency (Hz)", freq)
            return export_df, "nyquist_plot.csv", "Export Nyquist CSV As..."

        if plot_type == 'bode':
            z_mag = np.sqrt(z_real**2 + z_imag**2)
            export_df = pd.DataFrame({
                "Frequency (Hz)": freq,
                "|Z| (Ohm)": z_mag,
            })
            return export_df, "bode_plot.csv", "Export Bode CSV As..."

        return None, None, None

    def _save_export_dataframe(self, export_df, default_name, dialog_title):
        """Prompt for save location and write CSV; returns filepath or None."""
        filetypes = [
            ('CSV File', '*.csv'),
            ('All Files', '*.*')
        ]

        filepath = filedialog.asksaveasfilename(
            title=dialog_title,
            initialfile=default_name,
            defaultextension=".csv",
            filetypes=filetypes
        )

        if not filepath:
            self.log_message("Export cancelled.")
            return None

        try:
            export_df.to_csv(filepath, index=False)
            self.log_message(f"CSV exported to: {filepath}")
            return filepath
        except Exception as e:
            self.log_message(f"Error exporting CSV: {e}")
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{e}")
            return None

    def _open_cloud_target(self, cloud_target):
        """Open selected cloud provider target for upload."""
        target = cloud_target.strip().lower()

        if target in ('1', 'onedrive', 'one drive'):
            onedrive_path = os.environ.get('OneDrive')
            if onedrive_path and os.path.isdir(onedrive_path):
                os.startfile(onedrive_path)
                return "OneDrive"
            webbrowser.open("https://onedrive.live.com")
            return "OneDrive"

        if target in ('2', 'google drive', 'gdrive', 'drive'):
            webbrowser.open("https://drive.google.com/drive/my-drive")
            return "Google Drive"

        if target in ('3', 'dropbox'):
            webbrowser.open("https://www.dropbox.com/home")
            return "Dropbox"

        return None

    def _generate_export_filename(self, base_name):
        """Generate timestamped filename to avoid collisions for cloud upload flow."""
        stem, ext = os.path.splitext(base_name)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        return f"{stem}_{timestamp}{ext}"

    def _cloud_export_path_for_choice(self, cloud_choice, default_name):
        """Resolve automatic save path for cloud upload without a local save dialog."""
        choice = cloud_choice.strip().lower()
        filename = self._generate_export_filename(default_name)

        if choice in ('1', 'onedrive', 'one drive'):
            onedrive_path = os.environ.get('OneDrive')
            if onedrive_path and os.path.isdir(onedrive_path):
                return os.path.join(onedrive_path, filename)

        queue_dir = os.path.join(os.path.dirname(__file__), "cloud_upload_queue")
        os.makedirs(queue_dir, exist_ok=True)
        return os.path.join(queue_dir, filename)

    def export_plot(self, plot_type):
        """Opens a 'Save As' dialog to export Nyquist/Bode plotted data as CSV."""

        export_df, default_name, dialog_title = self._build_export_dataframe(plot_type)
        if export_df is None:
            self.log_message("No plot data available to export.")
            messagebox.showwarning("No Data", "No plot data available to export.")
            return

        self._save_export_dataframe(export_df, default_name, dialog_title)

    def export_plot_to_cloud(self, plot_type):
        """Export CSV then open selected cloud target for upload."""
        export_df, default_name, dialog_title = self._build_export_dataframe(plot_type)
        if export_df is None:
            self.log_message("No plot data available to export.")
            messagebox.showwarning("No Data", "No plot data available to export.")
            return

        prompt = (
            "Choose cloud destination:\n"
            "1 = OneDrive\n"
            "2 = Google Drive\n"
            "3 = Dropbox"
        )
        cloud_choice = simpledialog.askstring("Cloud Save", prompt, parent=self.root)
        if not cloud_choice:
            self.log_message("Cloud target selection cancelled.")
            return

        filepath = self._cloud_export_path_for_choice(cloud_choice, default_name)
        try:
            export_df.to_csv(filepath, index=False)
            self.log_message(f"Cloud upload file prepared: {filepath}")
        except Exception as e:
            self.log_message(f"Error preparing cloud upload file: {e}")
            messagebox.showerror("Cloud Save Error", f"Failed to prepare CSV for upload:\n{e}")
            return

        cloud_name = self._open_cloud_target(cloud_choice)
        if cloud_name is None:
            messagebox.showwarning("Cloud Save", "Unknown cloud option. Use 1, 2, or 3.")
            self.log_message("Cloud save cancelled: unknown cloud option.")
            return

        messagebox.showinfo(
            "Cloud Save",
            f"CSV saved to:\n{filepath}\n\n{cloud_name} has been opened. Upload this file there.",
        )
        self.log_message(f"Cloud export ready: {cloud_name} opened for upload.")

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = EisAnalysisTool(root)
    
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit? (Connection will remain active if device is paired)"):
            # Keep Bluetooth connection alive on close—don't force disconnect
            # This allows the app to reconnect immediately on restart
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
