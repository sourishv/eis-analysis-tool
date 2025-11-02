import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import time
import numpy as np
import pandas as pd
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

        # --- Tab 1: Load Measurement ---
        self.eis_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.eis_frame, text='Load Measurement') # <-- Renamed Tab

        # --- NEW: Load Data Button ---
        # The parameter fields have been removed.
        load_frame = ttk.Frame(self.eis_frame)
        load_frame.pack(expand=True)
        
        self.load_data_btn = ttk.Button(
            load_frame, 
            text="Load EIS Data File (.csv)", 
            command=self.start_load_data_thread, 
            state="disabled"
        )
        self.load_data_btn.pack(pady=50, ipady=10, ipadx=10)
        # --- END OF CHANGES TO TAB 1 ---

        # --- Tab 2: Nyquist Plot ---
        nyquist_tab = ttk.Frame(self.notebook)
        self.notebook.add(nyquist_tab, text='Nyquist Plot')
        self.nyquist_fig = Figure(figsize=(6, 4), dpi=100)
        self.nyquist_ax = self.nyquist_fig.add_subplot(111)
        self.nyquist_canvas = FigureCanvasTkAgg(self.nyquist_fig, master=nyquist_tab)
        self.nyquist_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        
        export_nyquist_btn = ttk.Button(nyquist_tab, text="Save Nyquist Plot", command=lambda: self.export_plot('nyquist'))
        export_nyquist_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(5,0))
        
        self.nyquist_ax.plot_data = ([], []) 
        
        # --- Tab 3: Bode Plot (Magnitude Only) ---
        bode_tab = ttk.Frame(self.notebook)
        self.notebook.add(bode_tab, text='Bode Plot')
        
        self.bode_fig = Figure(figsize=(6, 4), dpi=100)
        
        self.bode_ax_mag = self.bode_fig.add_axes([0.15, 0.15, 0.8, 0.75], zorder=1)
        self.bode_cbar_ax = self.bode_fig.add_axes([0.15, 0.15, 0.03, 0.75], zorder=2)
        
        self.bode_ax_mag.patch.set_alpha(0) 
        self.bode_ax_mag.plot_data = ([], []) 

        self.bode_canvas = FigureCanvasTkAgg(self.bode_fig, master=bode_tab)
        self.bode_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

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

        self.root.after(500, self.start_connect_thread)

    # --- REMOVED _create_param_entry ---

    # --- Connection Logic (Simulated) ---

    def start_connect_thread(self):
        """Simulates starting a connection."""
        self.connect_btn.config(state="disabled")
        self.disconnect_btn.config(state="disabled")
        self.load_data_btn.config(state="disabled") # <-- CHANGED
        self.status_label.config(text="Status: Connecting... (Mock)", foreground="orange")
        self.root.after(1000, self.simulate_connection)

    def simulate_connection(self):
        """Simulates a successful connection."""
        self.status_label.config(text="Status: Connected (Mock Device)", foreground="green")
        self.disconnect_btn.config(state="normal")
        self.load_data_btn.config(state="normal") # <-- CHANGED
        self.log_message("Mock device connected. Ready to load data.")

    def simulate_disconnect(self):
        """Simulates a disconnect."""
        self.connect_btn.config(state="normal")
        self.disconnect_btn.config(state="disabled")
        self.load_data_btn.config(state="disabled") # <-- CHANGED
        self.status_label.config(text="Status: Disconnected (Mock)", foreground="red")
        self.log_message("Mock device disconnected.")

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
        self.load_data_btn.config(state="disabled")
        self.notebook.select(self.notebook.tabs()[-1]) # Switch to log

        filepath = filedialog.askopenfilename(
            title="Select EIS Data File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not filepath:
            self.log_message("File load cancelled.")
            self.load_data_btn.config(state="normal")
            return
            
        # Start the file processing in a separate thread
        threading.Thread(target=self.process_data_file, args=(filepath,), daemon=True).start()

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
            
            # --- Draw Plots (on main thread) ---
            self.root.after(0, self.draw_plots, data)

        except Exception as e:
            self.log_message(f"Error processing file: {e}")
        finally:
            # Re-enable the button from the main thread
            self.root.after(0, self.load_data_btn.config, {"state": "normal"})

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
    app = PalmSensApp(root)
    
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
