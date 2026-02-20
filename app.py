import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import time
import os
import re
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
        self.root.title("PC Test - Potentiostat GUI")
        self.root.geometry("900x700")

        # --- Style for larger UI elements ---
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 12))
        style.configure("TButton", font=("Helvetica", 12), padding=10)
        style.configure("TEntry", font=("Helvetica", 12))
        style.configure("TFrame", padding=10)
        style.configure("Header.TLabel", font=("Helvetica", 14, "bold"))
        style.configure("TNotebook.Tab", font=("Helvetica", 11, "bold"), padding=[10, 5])
        style.configure("Status.TLabel", font=("Helvetica", 12, "italic"))
        style.configure("Connect.TButton", font=("Helvetica", 10), padding=5)
        # Progressbar style with green fill for determinate mode
        style.configure("Green.Horizontal.TProgressbar", troughcolor='white', background='#4CAF50')

        # --- Connection Status Frame ---
        connect_frame = ttk.Frame(root, relief="groove", borderwidth=2)
        connect_frame.pack(fill="x", padx=10, pady=10)
        connect_frame.columnconfigure(1, weight=1)

        self.connect_btn = ttk.Button(connect_frame, text="Connect", command=self.start_connect_thread, style="Connect.TButton")
        self.connect_btn.grid(row=0, column=0, sticky="w", padx=(10, 5))

        self.disconnect_btn = ttk.Button(connect_frame, text="Disconnect", command=self.disconnect_device, style="Connect.TButton", state="disabled")
        self.disconnect_btn.grid(row=0, column=1, sticky="w", padx=5)

        self.status_label = ttk.Label(
            connect_frame, 
            text="Status: Disconnected", 
            foreground="red", 
            style="Status.TLabel"
        )
        self.status_label.grid(row=0, column=2, sticky="e", padx=10)
        
        # --- Main Control Area (using a Notebook/Tabs) ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        # Shared progress variable for progress bars (defined before any bar uses it)
        self.progress_var = tk.DoubleVar(value=0.0)

        # PyPalmSens runtime state
        self.ps_manager = None
        self.ps_instrument = None
        self.expected_points = 0
        self.measurement_in_progress = False
        self.target_mac = "00:16:A4:79:4E:03"
        self.measurement_start_time = None
        self.last_point_time = None
        self.last_point_count = 0
        self.stop_requested = False
        self.callback_debug_count = 0
        self.last_plot_update_time = 0  # Throttle plot updates (milliseconds)

        # --- Tab 1: Load Measurement ---
        self.eis_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.eis_frame, text='Measurement Setup') # <-- Renamed Tab

        # --- NEW: Load Data Button ---
        # The parameter fields have been removed.
        load_frame = ttk.Frame(self.eis_frame)
        load_frame.pack(expand=True)

        # --- Editable EIS parameter fields ---
        params = [
            ("Start Frequency (Hz)", "1e5"),
            ("End Frequency (Hz)", "1e-1"),
            ("Voltage Amplitude (mV)", "50"),
            ("Points per Decade", "5"),
        ]

        self.param_vars = {}
        params_frame = ttk.Frame(load_frame)
        params_frame.pack(pady=20)

        for i, (label_text, default) in enumerate(params):
            lbl = ttk.Label(params_frame, text=label_text)
            lbl.grid(row=i, column=0, sticky="e", padx=(0,8), pady=4)
            var = tk.StringVar(value=default)
            ent = ttk.Entry(params_frame, textvariable=var, width=20)
            ent.grid(row=i, column=1, sticky="w", pady=4)
            self.param_vars[label_text] = var

        # Run Test button - uses the CSV in project directory and streams data
        self.run_test_btn = ttk.Button(
            load_frame,
            text="Run Test",
            command=self.start_run_test_thread,
            state="disabled"
        )
        self.run_test_btn.pack(pady=(10,40), ipady=8, ipadx=10)

        self.stop_test_btn = ttk.Button(
            load_frame,
            text="Stop Test",
            command=self.request_stop_measurement,
            state="disabled"
        )
        self.stop_test_btn.pack(pady=(0, 10), ipady=6, ipadx=10)

        # Small label to show streaming progress
        self.load_progress_lbl = ttk.Label(load_frame, text="No test running")
        self.load_progress_lbl.pack()
        # --- END OF CHANGES TO TAB 1 ---

        # --- Tab 2: Nyquist Plot ---
        nyquist_tab = ttk.Frame(self.notebook)
        self.notebook.add(nyquist_tab, text='Nyquist Plot')
        self.nyquist_fig = Figure(figsize=(6, 4), dpi=100)
        self.nyquist_ax = self.nyquist_fig.add_subplot(111)
        self.nyquist_canvas = FigureCanvasTkAgg(self.nyquist_fig, master=nyquist_tab)
        self.nyquist_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        
        self.export_nyquist_btn = ttk.Button(nyquist_tab, text="Save Nyquist Plot", command=lambda: self.export_plot('nyquist'))
        self.export_nyquist_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5,0))

        self.nyquist_ax.plot_data = ([], []) 
        self.nyquist_line = None
        self.nyquist_diag_text = None
        
        # --- Tab 3: Bode Plot (Magnitude Only) ---
        bode_tab = ttk.Frame(self.notebook)
        self.notebook.add(bode_tab, text='Bode Plot')
        
        self.bode_fig = Figure(figsize=(6, 4), dpi=100)
        
        self.bode_ax_mag = self.bode_fig.add_axes([0.15, 0.15, 0.8, 0.75], zorder=1)
        self.bode_cbar_ax = self.bode_fig.add_axes([0.15, 0.15, 0.03, 0.75], zorder=2)
        
        self.bode_ax_mag.patch.set_alpha(0) 
        self.bode_ax_mag.plot_data = ([], []) 

        self.bode_line = None
        self.bode_diag_text = None

        self.bode_canvas = FigureCanvasTkAgg(self.bode_fig, master=bode_tab)
        self.bode_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        self.export_bode_btn = ttk.Button(bode_tab, text="Save Bode Plot", command=lambda: self.export_plot('bode'))
        self.export_bode_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5,0))

        # --- Tab 4: Output Log ---
        log_tab = ttk.Frame(self.notebook)
        self.notebook.add(log_tab, text='Output Log')
        # Progress bar shown while a test is running
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(log_tab, orient='horizontal', mode='determinate', variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', padx=5, pady=(5,3))

        self.output_text = scrolledtext.ScrolledText(log_tab, height=10, state="disabled", font=("Courier New", 10))
        self.output_text.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Initialize Plots & Annotations ---
        self.init_nyquist_plot()
        self.init_bode_plot()

        self.nyquist_annot = self.nyquist_ax.annotate("", xy=(0,0), xytext=(15,15),
            textcoords="offset points",
            bbox=dict(boxstyle="round", fc="w", alpha=0.7),
            arrowprops=dict(arrowstyle="->"))
        self.nyquist_annot.set_visible(False)

        self.bode_annot = self.bode_ax_mag.annotate("", xy=(0,0), xytext=(15,15),
            textcoords="offset points",
            bbox=dict(boxstyle="round", fc="w", alpha=0.7),
            arrowprops=dict(arrowstyle="->"))
        self.bode_annot.set_visible(False)

        self.nyquist_canvas.mpl_connect("motion_notify_event", self.on_plot_hover)
        self.bode_canvas.mpl_connect("motion_notify_event", self.on_plot_hover)

        # Try fast reconnect on startup (skip discovery if already connected and MAC is known)
        if self.target_mac:
            self.root.after(500, self.start_connect_thread)

    # --- REMOVED _create_param_entry ---

    # --- Connection Logic (Simulated) ---

    def start_connect_thread(self):
        """Connect to a PalmSens instrument over Bluetooth using PyPalmSens."""
        self.connect_btn.config(state="disabled")
        self.disconnect_btn.config(state="disabled")
        self.run_test_btn.config(state="disabled")
        self.status_label.config(text="Status: Connecting (Bluetooth)...", foreground="orange")
        threading.Thread(target=self.connect_device, daemon=True).start()

    def connect_device(self):
        """Discover and connect to first available Bluetooth PalmSens instrument."""
        if ps is None:
            self.log_message("ERROR: PyPalmSens is not installed. Install with: pip install pypalmsens")
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

            # Try fast reconnect by discovery (SDK will pick the device)
            # if MAC is known, discovery should return it first
            self.log_message("Scanning for Bluetooth PalmSens instruments...")
            instruments = ps.discover(
                ftdi=False,
                usbcdc=False,
                winusb=False,
                bluetooth=True,
                serial=False,
                ignore_errors=True,
            )
            if not instruments:
                self.log_message("No Bluetooth PalmSens instruments found.")
                self.root.after(0, self._set_disconnected_ui)
                return

            self.log_message(f"Discovered {len(instruments)} device(s):")
            for idx, inst in enumerate(instruments, start=1):
                self.log_message(f"  {idx}. {self._describe_instrument(inst)}")

            # Fixed MAC targeting (hard-coded in __init__)
            selected = None
            if self.target_mac:
                for inst in instruments:
                    if self._instrument_matches_mac(inst, self.target_mac):
                        selected = inst
                        break
                if selected is None:
                    self.log_message(f"Configured MAC {self.target_mac} not found in discovered devices.")
                    self.root.after(0, self._set_disconnected_ui)
                    return
                self.log_message(f"Using configured MAC match: {self._describe_instrument(selected)}")
            else:
                selected = instruments[0]

            self.ps_instrument = selected

            # Retry once for transient socket/binding issues
            last_err = None
            for attempt in range(1, 3):
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
            self.log_message(f"Connected to {self.ps_instrument.name} (Serial: {serial})")
            self.root.after(0, self._set_connected_ui)
        except Exception as e:
            self.log_message(f"Bluetooth connection failed: {e}")
            self.ps_manager = None
            self.ps_instrument = None
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
            self._set_disconnected_ui()
            self.log_message("Device disconnected.")

    def _set_connected_ui(self):
        self.status_label.config(text="Status: Connected", foreground="green")
        self.disconnect_btn.config(state="normal")
        self.run_test_btn.config(state="normal")

    def _set_disconnected_ui(self):
        self.connect_btn.config(state="normal")
        self.disconnect_btn.config(state="disabled")
        self.run_test_btn.config(state="disabled")
        self.status_label.config(text="Status: Disconnected", foreground="red")

    # --- Plotting Initialization ---

    def init_nyquist_plot(self):
        self.nyquist_ax.clear()
        self.nyquist_ax.set_xlabel('Z_real (Ohm)')
        self.nyquist_ax.set_ylabel('-Z_imaginary (Ohm)')
        self.nyquist_ax.set_title("Nyquist Plot")
        self.nyquist_ax.grid(True)
        
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
        self.bode_ax_mag.set_ylim(1e2, 1e10)
        self.bode_ax_mag.set_xlim(1e-2, 1e5)
        self.bode_ax_mag.set_ylabel('|Z| (Ohm)')
        self.bode_ax_mag.set_xlabel('Frequency (Hz)')
        self.bode_ax_mag.set_title("Bode Plot")
        self.bode_ax_mag.grid(True, which='both')
        self.bode_ax_mag.set_yscale('log')
        self.bode_ax_mag.set_xscale('log')
        
        self.bode_cbar_ax.clear()
        self.bode_cbar_ax.set_yscale('log')
        self.bode_cbar_ax.set_ylim(1e2, 1e10)
        
        self.bode_cbar_ax.axhspan(1e2, 1e5, facecolor='#FF8A80', alpha=0.4) # Red
        self.bode_cbar_ax.axhspan(1e5, 1e7, facecolor='#FFFF8D', alpha=0.4) # Yellow
        self.bode_cbar_ax.axhspan(1e7, 1e10, facecolor='#B9F6CA', alpha=0.4) # Green

        self.bode_cbar_ax.set_xticks([])
        self.bode_cbar_ax.set_yticks([])
        self.bode_cbar_ax.set_yticklabels([])
        
        self.bode_cbar_ax.patch.set_alpha(0) 
        
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
    
    def draw_plots(self, data):
        try:
            freq, z_real, z_imag = data['frequency'], data['z_real'], data['z_imag']
            z_mag = np.sqrt(z_real**2 + z_imag**2)
            z_imag_neg = -z_imag
            
            # --- 1. Nyquist Plot ---
            self.init_nyquist_plot() # Clears, sets formatters
            self.nyquist_ax.plot(z_real, z_imag_neg, 'o-', markersize=4, color='blue')
            self.nyquist_ax.plot_data = (z_real, z_imag_neg)
            self.nyquist_ax.relim()
            self.nyquist_ax.autoscale_view()
            
            # --- NEW: Add axis('equal') back ---
            # This makes the plot scales 1:1, just like MATLAB
            self.nyquist_ax.axis('equal') 
            
            self.nyquist_canvas.draw()

            # --- 2. Bode Plot ---
            self.init_bode_plot() 
            self.bode_ax_mag.loglog(freq, z_mag, 'o-', markersize=4, color='blue', zorder=10)
            self.bode_ax_mag.plot_data = (freq, z_mag)
            self.bode_canvas.draw()
            
            self.log_message("Plots updated.")
            self.notebook.select(self.notebook.tabs()[2]) # Switch to Bode plot

        except Exception as e:
            self.log_message(f"Failed to draw plots: {e}")

    # --- New: Streaming load for Run Test ---
    def start_run_test_thread(self):
        """Starts a real EIS measurement using the connected PalmSens instrument."""
        if self.ps_manager is None:
            self.log_message("ERROR: No instrument connected. Connect via Bluetooth first.")
            return
        if self.measurement_in_progress:
            self.log_message("Measurement already in progress.")
            return

        self.measurement_in_progress = True
        self.stop_requested = False
        self.measurement_start_time = time.time()
        self.last_point_time = self.measurement_start_time
        self.last_point_count = 0
        # Disable button and switch to log
        self.run_test_btn.config(state="disabled")
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
            try:
                if hasattr(self, 'shared_progress_frame') and self.shared_progress_frame.winfo_exists():
                    self.shared_progress_frame.destroy()
            except Exception:
                pass

            self.shared_progress_frame = ttk.Frame(bode_tab_widget)
            # pack it before the save button so it appears above
            self.shared_progress = ttk.Progressbar(self.shared_progress_frame, style='Green.Horizontal.TProgressbar', orient='horizontal', mode='determinate', variable=self.progress_var, maximum=100)
            self.shared_progress.pack(side='left', fill='x', expand=True, padx=(0,8))
            self.shared_progress_label = ttk.Label(self.shared_progress_frame, text='0%')
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
        self.load_progress_lbl.config(text="Starting EIS measurement...")
        self.root.after(5000, self.measurement_watchdog_tick)
        threading.Thread(target=self.run_real_eis_measurement, daemon=True).start()

    def request_stop_measurement(self):
        """Request cancellation of the active EIS measurement."""
        if not self.measurement_in_progress:
            self.log_message("No active measurement to stop.")
            return

        self.stop_requested = True
        self.stop_test_btn.config(state="disabled")
        self.load_progress_lbl.config(text="Stopping measurement...")
        self.log_message("Stop requested. Sending stop signal to potentiostat...")
        threading.Thread(target=self._send_stop_signal_to_instrument, daemon=True).start()

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

    def run_real_eis_measurement(self):
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
                    try:
                        self.root.after(0, self.shared_progress_label.config, {"text": f"{percent:.0f}%"})
                    except Exception:
                        pass
                    
                    # Throttle plot updates to prevent overwhelming the UI and blocking mouse events
                    current_time = time.time() * 1000  # ms
                    if current_time - self.last_plot_update_time >= 100:  # Update max every 100ms
                        self.last_plot_update_time = current_time
                        self.root.after(0, self.update_plots_incremental, np.array(freq_buf), np.array(zre_buf), np.array(zim_buf))
            except Exception as cb_err:
                self.log_message(f"Callback error: {cb_err}")

        try:
            method = self.build_eis_method()
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
                diagnosis_result = self.diagnose_coating(z_mag, np.array(freq_buf))
                self.log_message(f"Diagnosis: {diagnosis_result}")
                self.root.after(0, self.show_diagnosis_on_plots, diagnosis_result)

            if self.stop_requested:
                self.root.after(0, self.load_progress_lbl.config, {"text": "Measurement stopped"})
                self.log_message("Measurement stopped by user.")
            else:
                self.root.after(0, self.progress_var.set, 100.0)
                self.root.after(0, self.load_progress_lbl.config, {"text": "Test finished"})
            try:
                self.root.after(0, self.shared_progress_label.config, {"text": "100%" if not self.stop_requested else "Stopped"})
            except Exception:
                pass
        except Exception as e:
            if self.stop_requested:
                self.log_message("Measurement stop completed.")
                self.root.after(0, self.load_progress_lbl.config, {"text": "Measurement stopped"})
            else:
                self.log_message(f"Real EIS measurement failed: {e}")
                self.root.after(0, self.progress_var.set, 0.0)
                self.root.after(0, self.load_progress_lbl.config, {"text": "Measurement failed"})
            try:
                self.root.after(0, self.shared_progress_label.config, {"text": "0%" if not self.stop_requested else "Stopped"})
            except Exception:
                pass
        finally:
            self.measurement_in_progress = False
            self.stop_requested = False
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
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
                x_buf.append(freq[i])
                yr_buf.append(z_real[i])
                yi_buf.append(z_imag[i])

                # Update progress label on main thread
                percent = (i+1) / n * 100.0
                self.root.after(0, self.load_progress_lbl.config, {"text": f"Loading: {i+1}/{n} points"})
                # Update progress variable
                self.root.after(0, self.progress_var.set, percent)
                # Update shared progress label if present
                try:
                    self.root.after(0, self.shared_progress_label.config, {"text": f"{percent:.0f}%"})
                except Exception:
                    pass

                # Update plots with current subset
                self.root.after(0, self.update_plots_incremental, np.array(x_buf), np.array(yr_buf), np.array(yi_buf))

                time.sleep(interval)

            self.log_message("Test complete. Full data loaded.")
            # Determine coating health based on the last impedance magnitude
            try:
                # Full magnitude using original arrays
                full_z_mag = np.sqrt(z_real**2 + z_imag**2)
                diagnosis_result = self.diagnose_coating(full_z_mag, freq)
                self.log_message(f"Diagnosis: {diagnosis_result}")
                # Show diagnosis visually on plots
                self.root.after(0, self.show_diagnosis_on_plots, diagnosis_result)
            except Exception as e:
                self.log_message(f"Diagnosis failed: {e}")
            finally:
                # destroy shared progress UI if present
                try:
                    if hasattr(self, 'shared_progress_frame') and self.shared_progress_frame.winfo_exists():
                        self.root.after(0, self.shared_progress_frame.destroy)
                except Exception:
                    pass
            # Re-enable button and reset progress
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.load_progress_lbl.config, {"text": "Test finished"})
            self.root.after(0, self.progress_var.set, 100.0)
            try:
                self.root.after(0, self.shared_progress_label.config, {"text": "100%"})
            except Exception:
                pass

        except Exception as e:
            self.log_message(f"Error streaming test data: {e}")
            self.root.after(0, self.run_test_btn.config, {"state": "normal"})
            self.root.after(0, self.progress_var.set, 0.0)
            try:
                self.root.after(0, self.shared_progress_label.config, {"text": "0%"})
            except Exception:
                pass

    def update_plots_incremental(self, freq_subset, z_real_subset, z_imag_subset):
        """Update Nyquist and Bode plots with partial data during streaming."""
        try:
            z_mag = np.sqrt(z_real_subset**2 + z_imag_subset**2)
            z_imag_neg = -z_imag_subset

            # --- Nyquist: create or update a persistent Line2D to avoid clearing the axes ---
            if self.nyquist_line is None:
                # create the initial line on existing axes (axes already initialized in init_nyquist_plot)
                (self.nyquist_line,) = self.nyquist_ax.plot(z_real_subset, z_imag_neg, 'o-', markersize=4, color='blue')
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
                (self.bode_line,) = self.bode_ax_mag.plot(safe_freq, z_mag, 'o-', markersize=4, color='blue', zorder=10)
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
                face = '#4CAF50'  # green
                fg = 'white'
            elif 'monitor' in txt or 'caution' in txt or 'medium' in txt:
                face = '#FFB300'  # amber
                fg = 'black'
            else:
                face = '#E53935'  # red
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

    # --- Plot Export Function (Unchanged) ---
    def export_plot(self, plot_type):
        """Opens a 'Save As' dialog to export the specified plot."""
        
        if plot_type == 'nyquist':
            fig_to_save = self.nyquist_fig
            default_name = "nyquist_plot.png"
            dialog_title = "Save Nyquist Plot As..."
        elif plot_type == 'bode':
            fig_to_save = self.bode_fig
            default_name = "bode_plot.png"
            dialog_title = "Save Bode Plot As..."
        else:
            return
        
        filetypes = [
            ('PNG Image', '*.png'),
            ('JPEG Image', '*.jpg'),
            ('PDF Document', '*.pdf'),
            ('All Files', '*.*')
        ]
        
        filepath = filedialog.asksaveasfilename(
            title=dialog_title,
            initialfile=default_name,
            defaultextension=".png",
            filetypes=filetypes
        )
        
        if not filepath:
            self.log_message("Export cancelled.")
            return
        
        try:
            fig_to_save.savefig(filepath, dpi=300, bbox_inches='tight')
            self.log_message(f"Plot saved to: {filepath}")
        except Exception as e:
            self.log_message(f"Error saving plot: {e}")
            messagebox.showerror("Save Error", f"Failed to save plot:\n{e}")

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
