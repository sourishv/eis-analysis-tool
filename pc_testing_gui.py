import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import time
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class PalmSensApp:
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

        # --- Connection Status Frame ---
        connect_frame = ttk.Frame(root, relief="groove", borderwidth=2)
        connect_frame.pack(fill="x", padx=10, pady=10)
        connect_frame.columnconfigure(1, weight=1)

        self.connect_btn = ttk.Button(connect_frame, text="Connect", command=self.start_connect_thread, style="Connect.TButton")
        self.connect_btn.grid(row=0, column=0, sticky="w", padx=(10, 5))

        self.disconnect_btn = ttk.Button(connect_frame, text="Disconnect", command=self.simulate_disconnect, style="Connect.TButton", state="disabled")
        self.disconnect_btn.grid(row=0, column=1, sticky="w", padx=5)

        self.status_label = ttk.Label(
            connect_frame, 
            text="Status: Disconnected (Mock)", 
            foreground="red", 
            style="Status.TLabel"
        )
        self.status_label.grid(row=0, column=2, sticky="e", padx=10)
        
        # --- Main Control Area (using a Notebook/Tabs) ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Tab 1: EIS Parameters ---
        self.eis_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.eis_frame, text='Run Measurement')

        params_frame = ttk.Frame(self.eis_frame, relief="sunken", borderwidth=2)
        params_frame.pack(fill="x", pady=10, padx=20)
        params_frame.columnconfigure(1, weight=1)
        ttk.Label(params_frame, text="EIS Parameters", style="Header.TLabel").grid(row=0, column=0, columnspan=3, pady=10)

        self.eis_params = {}
        self._create_param_entry(params_frame, "OCP (sec):", "ocp_time", "15", 1)
        self._create_param_entry(params_frame, "Start Freq (Hz):", "start_freq", "100000", 2) # 10^5 Hz
        self._create_param_entry(params_frame, "End Freq (Hz):", "end_freq", "0.01", 3)      # 10^-2 Hz
        self._create_param_entry(params_frame, "Amplitude (mV):", "amplitude", "50", 4)
        self._create_param_entry(params_frame, "Points/Decade:", "points_per_decade", "5", 5)

        self.run_eis_btn = ttk.Button(self.eis_frame, text="Run EIS Measurement", command=self.start_eis_thread, state="disabled")
        self.run_eis_btn.pack(pady=20)

        # --- Tab 2: Nyquist Plot ---
        nyquist_tab = ttk.Frame(self.notebook)
        self.notebook.add(nyquist_tab, text='Nyquist Plot')
        self.nyquist_fig = Figure(figsize=(6, 4), dpi=100)
        self.nyquist_ax = self.nyquist_fig.add_subplot(111)
        self.nyquist_canvas = FigureCanvasTkAgg(self.nyquist_fig, master=nyquist_tab)
        self.nyquist_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        
        # --- NEW: Export button for Nyquist ---
        export_nyquist_btn = ttk.Button(nyquist_tab, text="Save Nyquist Plot", command=lambda: self.export_plot('nyquist'))
        export_nyquist_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5,0))
        
        # Store plot data for hover
        self.nyquist_ax.plot_data = ([], []) 
        
        # --- Tab 3: Bode Plot (Magnitude Only) ---
        bode_tab = ttk.Frame(self.notebook)
        self.notebook.add(bode_tab, text='Bode Plot')
        
        self.bode_fig = Figure(figsize=(6, 4), dpi=100)
        
        # Manually add axes to control layout
        # [left, bottom, width, height]
        # Main plot axes (zorder=1, behind)
        self.bode_ax_mag = self.bode_fig.add_axes([0.15, 0.15, 0.8, 0.75], zorder=1)
        # Color bar axes (zorder=2, on top)
        self.bode_cbar_ax = self.bode_fig.add_axes([0.15, 0.15, 0.03, 0.75], zorder=2)
        
        # Make main plot background transparent so grid shows through
        self.bode_ax_mag.patch.set_alpha(0) 
        
        # Store plot data for hover
        self.bode_ax_mag.plot_data = ([], []) 

        self.bode_canvas = FigureCanvasTkAgg(self.bode_fig, master=bode_tab)
        self.bode_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

        # --- NEW: Export button for Bode ---
        export_bode_btn = ttk.Button(bode_tab, text="Save Bode Plot", command=lambda: self.export_plot('bode'))
        export_bode_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5,0))

        # --- Tab 4: Output Log ---
        log_tab = ttk.Frame(self.notebook)
        self.notebook.add(log_tab, text='Output Log')
        self.output_text = scrolledtext.ScrolledText(log_tab, height=10, state="disabled", font=("Courier New", 10))
        self.output_text.pack(fill="both", expand=True, padx=5, pady=5)

        # --- Initialize Plots & Annotations ---
        self.init_nyquist_plot()
        self.init_bode_plot()

        # Create annotation objects (initially invisible)
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

        # Connect hover events
        self.nyquist_canvas.mpl_connect("motion_notify_event", self.on_plot_hover)
        self.bode_canvas.mpl_connect("motion_notify_event", self.on_plot_hover)

        # --- Auto-connect on startup ---
        self.root.after(500, self.start_connect_thread)

    def _create_param_entry(self, parent, label_text, param_key, default_value, row):
        label = ttk.Label(parent, text=label_text)
        label.grid(row=row, column=0, sticky="w", padx=10, pady=5)
        entry = ttk.Entry(parent, width=15)
        entry.grid(row=row, column=1, sticky="w", padx=10, pady=5)
        entry.insert(0, default_value)
        self.eis_params[param_key] = entry
        return entry

    # --- Connection Logic (Simulated) ---

    def start_connect_thread(self):
        """Simulates starting a connection."""
        self.connect_btn.config(state="disabled")
        self.disconnect_btn.config(state="disabled")
        self.run_eis_btn.config(state="disabled")
        self.status_label.config(text="Status: Connecting... (Mock)", foreground="orange")
        # Simulate connection delay
        self.root.after(1000, self.simulate_connection)

    def simulate_connection(self):
        """Simulates a successful connection."""
        self.status_label.config(text="Status: Connected (Mock Device)", foreground="green")
        self.disconnect_btn.config(state="normal")
        self.run_eis_btn.config(state="normal")
        self.log_message("Mock device connected.")

    def simulate_disconnect(self):
        """Simulates a disconnect."""
        self.connect_btn.config(state="normal")
        self.disconnect_btn.config(state="disabled")
        self.run_eis_btn.config(state="disabled")
        self.status_label.config(text="Status: Disconnected (Mock)", foreground="red")
        self.log_message("Mock device disconnected.")

    # --- Plotting Initialization ---

    def init_nyquist_plot(self):
        self.nyquist_ax.clear()
        self.nyquist_ax.set_xlabel('Z_real (Ohm)')
        self.nyquist_ax.set_ylabel('-Z_imaginary (Ohm)') # Updated label
        self.nyquist_ax.set_title("Nyquist Plot")
        self.nyquist_ax.grid(True)
        self.nyquist_fig.tight_layout()
        self.nyquist_canvas.draw()

    def init_bode_plot(self):
        """Initializes Bode plot with color bar and fixed axes."""
        
        # --- 1. Main Plot (self.bode_ax_mag) ---
        self.bode_ax_mag.clear()
        self.bode_ax_mag.set_ylim(1e2, 1e10)
        self.bode_ax_mag.set_xlim(1e-2, 1e5) # Updated X-axis limit
        self.bode_ax_mag.set_ylabel('|Z| (Ohm)')
        self.bode_ax_mag.set_xlabel('Frequency (Hz)')
        self.bode_ax_mag.set_title("Bode Plot")
        self.bode_ax_mag.grid(True, which='both')
        self.bode_ax_mag.set_yscale('log')
        self.bode_ax_mag.set_xscale('log')
        
        # --- 2. Color Bar Plot (self.bode_cbar_ax) ---
        self.bode_cbar_ax.clear()
        self.bode_cbar_ax.set_yscale('log')
        self.bode_cbar_ax.set_ylim(1e2, 1e10)
        
        # Add colored spans to the thin bar
        self.bode_cbar_ax.axhspan(1e2, 1e5, facecolor='#FF8A80', alpha=0.4) # Red
        self.bode_cbar_ax.axhspan(1e5, 1e7, facecolor='#FFFF8D', alpha=0.4) # Yellow
        self.bode_cbar_ax.axhspan(1e7, 1e10, facecolor='#B9F6CA', alpha=0.4) # Green

        # Hide all ticks and labels on the color bar
        self.bode_cbar_ax.set_xticks([])
        self.bode_cbar_ax.set_yticks([])
        self.bode_cbar_ax.set_yticklabels([])
        
        # Make the color bar's own background transparent
        self.bode_cbar_ax.patch.set_alpha(0) 
        
        self.bode_fig.tight_layout(rect=[0.05, 0, 1, 1])
        self.bode_canvas.draw()

    # --- Plot Hover Logic ---
    
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
        # Determine which annotation and data to use
        if event.inaxes == self.bode_ax_mag:
            annot = self.bode_annot
            data_x, data_y = self.bode_ax_mag.plot_data
            is_log_x = True
        elif event.inaxes == self.nyquist_ax:
            annot = self.nyquist_annot
            data_x, data_y = self.nyquist_ax.plot_data
            is_log_x = False
        else:
            # Hide annotations if not on a main axes
            if self.bode_annot.get_visible():
                self.bode_annot.set_visible(False)
                self.bode_canvas.draw_idle()
            if self.nyquist_annot.get_visible():
                self.nyquist_annot.set_visible(False)
                self.nyquist_canvas.draw_idle()
            return

        # Check if data exists
        if len(data_x) == 0:
            return

        # Find the nearest point
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return

        if is_log_x:
            log_dist_x = (np.log10(data_x) - np.log10(x))**2
            log_dist_y = (np.log10(data_y) - np.log10(y))**2
            dist = log_dist_x + log_dist_y
        else:
            x_range = np.max(data_x) - np.min(data_x)
            y_range = np.max(data_y) - np.min(data_y)
            if x_range == 0 or y_range == 0: return # Avoid division by zero
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

    # --- Measurement Logic ---

    def log_message(self, msg):
        def _log():
            self.output_text.configure(state="normal")
            self.output_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {msg}\n")
            self.output_text.see(tk.END)
            self.output_text.configure(state="disabled")
        self.root.after(0, _log)

    def get_params_from_gui(self):
        params = {}
        try:
            params['ocp_time'] = float(self.eis_params['ocp_time'].get())
            params['start_freq'] = float(self.eis_params['start_freq'].get())
            params['end_freq'] = float(self.eis_params['end_freq'].get())
            params['amplitude_v'] = float(self.eis_params['amplitude'].get()) / 1000.0
            params['points_per_decade'] = int(self.eis_params['points_per_decade'].get())
            
            log_f_range = np.abs(np.log10(params['start_freq']) - np.log10(params['end_freq']))
            params['n_points'] = int(log_f_range * params['points_per_decade']) + 1
        except Exception as e:
            params['error'] = e
        return params

    def start_eis_thread(self):
        self.run_eis_btn.config(state="disabled")
        self.notebook.select(self.notebook.tabs()[-1])
        threading.Thread(target=self.run_eis_measurement, daemon=True).start()

    def generate_fake_data(self, params):
        """
        Generates fake EIS data by interpolating between key points.
        Phase is forced to 0 at ends to satisfy Nyquist plot shape.
        """
        self.log_message("Generating fake data (Piecewise Log-Log Model)...")

        # --- Define Breakpoints (Frequency, Magnitude) ---
        bp_freqs = np.array([
            1e-2,  # P1: (10^-2 Hz, 10^6 Ohm)
            500,   # P2: (500 Hz, 2e5 Ohm)
            1e5    # P3: (10^5 Hz, 10^3 Ohm)
        ])
        bp_mags = np.array([
            1e6,   # |Z| at P1
            2e5,   # |Z| at P2 ("a little above 10^5")
            1e3    # |Z| at P3
        ])
        
        # --- Define Plausible Phase (in Degrees) ---
        # Phase is forced to 0 at the ends to make Z_imag = 0.
        bp_phase_deg = np.array([
            0.0,   # Phase = 0 at low freq (Resistive)
            -85.0, # Max capacitive phase at the "knee"
            0.0    # Phase = 0 at high freq (Resistive)
        ])

        # --- Get Query Frequencies ---
        freq_query = np.logspace(
            np.log10(params['start_freq']),
            np.log10(params['end_freq']),
            params['n_points']
        )
        
        # --- Interpolate in Log-Log Space ---
        log_mags_interp = np.interp(
            np.log10(freq_query),   # x (query points)
            np.log10(bp_freqs),     # xp (breakpoint x-values)
            np.log10(bp_mags)       # fp (breakpoint y-values)
        )
        z_mag = 10**log_mags_interp
        
        z_phase_deg = np.interp(
            np.log10(freq_query),
            np.log10(bp_freqs),
            bp_phase_deg
        )
        
        # --- Reconstruct Complex Impedance ---
        z_phase_rad = z_phase_deg * np.pi / 180.0
        Z_total = z_mag * (np.cos(z_phase_rad) + 1j * np.sin(z_phase_rad))
        
        # --- Add Noise ---
        noise_level = 0.01
        noise_real = np.random.normal(0, np.abs(Z_total) * noise_level, freq_query.shape)
        noise_imag = np.random.normal(0, np.abs(Z_total) * noise_level, freq_query.shape)
        Z_total_noisy = Z_total + noise_real + 1j * noise_imag

        return {
            'frequency': freq_query,
            'z_real': Z_total_noisy.real,
            'z_imag': Z_total_noisy.imag
        }

    def diagnose_coating(self, z_mag_data, freq_data):
        """Analyzes impedance data to provide a coating diagnosis."""
        try:
            # Find the impedance magnitude at the lowest frequency
            # Note: freq_data is high-to-low, so lowest freq is the last point
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

    def run_eis_measurement(self):
        try:
            self.log_message("Starting EIS measurement...")
            params_event = threading.Event()
            params = {}
            def _get_params():
                nonlocal params
                params = self.get_params_from_gui()
                params_event.set()
            self.root.after(0, _get_params)
            params_event.wait()

            if 'error' in params:
                raise ValueError(f"Invalid parameter input: {params['error']}")
            
            self.log_message(f"Method configured: OCP for {params['ocp_time']}s")
            self.log_message(f"  Freq: {params['start_freq']} Hz to {params['end_freq']} Hz")
            
            self.log_message("Running (simulated) measurement...")
            time.sleep(params['ocp_time'])
            self.log_message(f"OCP scan complete ({params['ocp_time']}s). Starting EIS...")
            time.sleep(2.0)
            
            # --- 3. Generate Data ---
            data = self.generate_fake_data(params)
            self.log_message(f"Acquired {len(data['frequency'])} data points.")
            
            # --- 4. Run Diagnosis ---
            z_mag = np.sqrt(data['z_real']**2 + data['z_imag']**2)
            diagnosis_result = self.diagnose_coating(z_mag, data['frequency'])
            self.log_message(f"Diagnosis: {diagnosis_result}")
            
            # --- 5. Draw Plots ---
            self.root.after(0, self.draw_plots, data)

        except Exception as e:
            self.log_message(f"Measurement error: {e}")
        finally:
            self.root.after(0, self.run_eis_btn.config, {"state": "normal"})
    
    def draw_plots(self, data):
        try:
            freq, z_real, z_imag = data['frequency'], data['z_real'], data['z_imag']
            z_mag = np.sqrt(z_real**2 + z_imag**2)
            z_imag_neg = -z_imag
            
            # --- 1. Nyquist Plot ---
            self.init_nyquist_plot()
            self.nyquist_ax.plot(z_real, z_imag_neg, 'o-', markersize=4, color='blue')
            # Store data for hover
            self.nyquist_ax.plot_data = (z_real, z_imag_neg)
            self.nyquist_ax.relim()
            self.nyquist_ax.autoscale_view()
            self.nyquist_canvas.draw()

            # --- 2. Bode Plot ---
            # Re-initialize plot (clears, sets limits, draws zones)
            self.init_bode_plot() 
            self.bode_ax_mag.loglog(freq, z_mag, 'o-', markersize=4, color='blue', zorder=10) # Plot on top
            # Store data for hover
            self.bode_ax_mag.plot_data = (freq, z_mag)
            self.bode_canvas.draw()
            
            self.log_message("Plots updated.")
            self.notebook.select(self.notebook.tabs()[2]) # Switch to Bode plot

        except Exception as e:
            self.log_message(f"Failed to draw plots: {e}")

    # --- NEW: Plot Export Function ---
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
            return # Should not happen
        
        filetypes = [
            ('PNG Image', '*.png'),
            ('JPEG Image', '*.jpg'),
            ('PDF Document', '*.pdf'),
            ('All Files', '*.*')
        ]
        
        # Ask the user for a file path
        filepath = filedialog.asksaveasfilename(
            title=dialog_title,
            initialfile=default_name,
            defaultextension=".png",
            filetypes=filetypes
        )
        
        if not filepath:
            # User cancelled the dialog
            self.log_message("Export cancelled.")
            return
        
        try:
            # Save the figure to the chosen path
            # Use high DPI for better quality and bbox_inches='tight'
            fig_to_save.savefig(filepath, dpi=300, bbox_inches='tight')
            self.log_message(f"Plot saved to: {filepath}")
        except Exception as e:
            self.log_message(f"Error saving plot: {e}")
            messagebox.showerror("Save Error", f"Failed to save plot:\n{e}")

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = PalmSensApp(root)
    
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

