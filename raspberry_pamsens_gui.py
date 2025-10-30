import tkinter as tk
from tkinter import messagebox, scrolledtext
import pypalmsens as ps
import threading

class PalmSensAppUSB:
    def __init__(self, root):
        self.root = root
        self.root.title("PalmSens USB Connector (on Pi)")
        self.root.geometry("400x400")

        self.connected_device_manager = None

        # --- GUI Elements ---
        
        # Connection Frame
        conn_frame = tk.Frame(root, pady=10)
        conn_frame.pack(fill="x")
        
        self.connect_btn = tk.Button(conn_frame, text="Connect via USB", command=self.connect_usb)
        self.connect_btn.pack(pady=10)

        self.status_label = tk.Label(conn_frame, text="Status: Not Connected", fg="red")
        self.status_label.pack()

        # Measurement Frame
        measure_frame = tk.Frame(root, pady=10)
        measure_frame.pack(fill="both", expand=True)

        tk.Label(measure_frame, text="Device Output:").pack(anchor="w", padx=10)
        self.output_text = scrolledtext.ScrolledText(measure_frame, height=10, state="disabled")
        self.output_text.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.measure_btn = tk.Button(measure_frame, text="Run Test Measurement", command=self.start_measurement_thread, state="disabled")
        self.measure_btn.pack(pady=5)

    def log_message(self, msg):
        """Helper function to print to the GUI text box."""
        self.output_text.configure(state="normal")
        self.output_text.insert(tk.END, msg + "\n")
        self.output_text.see(tk.END)
        self.output_text.configure(state="disabled")

    def set_status(self, msg, color):
        """Helper function to update the status label."""
        self.status_label.config(text=f"Status: {msg}", fg=color)

    def connect_usb(self):
        """Connect to the first available USB device."""
        try:
            self.set_status("Connecting via USB...", "orange")
            self.connect_btn.config(state="disabled")

            # This is the key change:
            # ps.connect() with no arguments finds the first device (usually USB)
            self.connected_device_manager = ps.connect()
            
            # Use .name attribute to get device name
            device_name = self.connected_device_manager.name
            
            self.log_message(f"Successfully connected to {device_name}!")
            self.set_status(f"Connected to {device_name}", "green")
            self.measure_btn.config(state="normal")
            self.connect_btn.config(text="Disconnect", state="normal", command=self.disconnect_device)

        except Exception as e:
            self.set_status(f"Connection failed: {e}", "red")
            self.log_message(f"Connection failed: {e}")
            self.connect_btn.config(state="normal")

    def disconnect_device(self):
        """Disconnects from the device."""
        if self.connected_device_manager:
            try:
                self.connected_device_manager.disconnect()
            except Exception as e:
                self.log_message(f"Error during disconnect: {e}")
            
            self.connected_device_manager = None
            self.set_status("Not Connected", "red")
            self.log_message("Disconnected.")
            self.measure_btn.config(state="disabled")
            self.connect_btn.config(text="Connect via USB", command=self.connect_usb)

    def start_measurement_thread(self):
        """
        Runs the measurement in a separate thread to avoid
        freezing the GUI.
        """
        threading.Thread(target=self.run_measurement, daemon=True).start()

    def run_measurement(self):
        """Run a simple Chronoamperometry measurement."""
        if not self.connected_device_manager:
            self.log_message("Not connected to any device.")
            return

        try:
            # Update GUI from the main thread
            self.root.after(0, self.log_message, "Running test measurement...")
            self.root.after(0, self.measure_btn.config, {"state": "disabled"})

            # Define a simple Chronoamperometry method
            method = ps.ChronoAmperometry(
                potential=0.1,  # 0.1 V
                run_time=5.0      # 5 seconds
            )

            # Run the measurement (this is a blocking call)
            measurement = self.connected_device_manager.measure(method)
            
            # --- Schedule GUI updates back on the main thread ---
            def update_gui_on_complete():
                self.log_message("Measurement complete.")
                self.log_message(f"Data points: {len(measurement.data['current'])}")
                
                # Show first 5 data points
                for i in range(min(5, len(measurement.data['current']))):
                    t = measurement.data['time'][i]
                    i_val = measurement.data['current'][i]
                    self.log_message(f"  t={t:.2f}s, i={i_val:.2e}A")
                
                self.measure_btn.config(state="normal")

            self.root.after(0, update_gui_on_complete)
            # --- End of main thread update ---

        except Exception as e:
            def update_gui_on_error():
                self.log_message(f"Measurement error: {e}")
                self.measure_btn.config(state="normal")
            
            self.root.after(0, update_gui_on_error)

# --- Main execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = PalmSensAppUSB(root)
    
    # Set a protocol to handle window close
    def on_closing():
        if app.connected_device_manager:
            app.disconnect_device()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
