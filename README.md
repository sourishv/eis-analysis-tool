# **Potentiostat EIS Software Analysis Tool**

This project is a Python-based PC application for simulating and visualizing Electrochemical Impedance Spectroscopy (EIS) data. It serves as a **PC testing tool** for developing and refining the user interface and features before deploying the final application to a Raspberry Pi with a real potentiostat.

This application is built with tkinter, matplotlib, and numpy.

## **Key Features**

* **Simulated EIS Measurement:** Runs a complete, simulated EIS test.
* **Realistic Fake Data:** Generates "fake" data by interpolating between key breakpoints in log-log space, precisely matching the shape of a target Bode plot.  
* **Interactive Plotting:**  
  * **Bode Plot:** Displays Impedance Magnitude |Z| vs. Frequency with a log-log scale.  
  * **Nyquist Plot:** Displays Z\_real vs. \-Z\_imaginary with a uniform, realistic semi-circular shape.  
* **Visual Diagnosis:**  
  * The Bode plot features a permanent color bar on the left (Green, Yellow, Red) to indicate coating health based on impedance magnitude.  
  * The Output Log provides a text-based diagnosis (e.g., "Healthy Coating," "Coating needs monitoring," "Defective Coating") based on the low-frequency impedance value.  
* **Data Export:** Both the Nyquist and Bode plots can be exported and saved as high-quality .png, .jpg, or .pdf files.  
* **Responsive UI:** Features a tabbed notebook interface to cleanly separate the measurement setup, plots, and output log. A simulated connection bar shows the device status.

## **Setup and Installation**

This project is intended to be run from a virtual environment.

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/sourishv/eis-analysis-tool.git
    cd eis-analysis-tool
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    # Create the venv
    python -m venv .venv

    # Activate on Windows (PowerShell)
    .\.venv\Scripts\Activate.ps1

    # Activate on macOS/Linux
    source .venv/bin/activate
    ```

3.  **Install the required libraries:**
    ```bash
    pip install -r requirements.txt
    ```

## How to Run the App

With your virtual environment active and packages installed, simply run the app:

```bash
python app.py
```

1. The application will launch and auto-connect to the "Mock Device."  
2. All EIS parameters are pre-filled with test values. Click **"Run EIS Measurement"** to start the simulation.  
3. The app will switch to the "Output Log" tab and show the process (scan, diagnosis).  
4. When complete, the app will switch to the "Bode Plot" tab to display the results.  
5. You can now switch between the "Nyquist Plot" and "Bode Plot" tabs, and use the "Save Plot" buttons to export the images.

## **Future Goals (Raspberry Pi Deployment)**

The ultimate goal of this project is to run on a Raspberry Pi. The next steps will involve:

1. **Hardware Integration:** Replace the simulated connection and generate\_fake\_data function with the actual pypalmsens library to connect to the potentiostat via USB.  
2. **Data Acquisition:** Call the real manager.measure() function and retrieve the *actual* data.  
3. **Porting:** Transfer the app.py script to the Raspberry Pi, install the dependencies, and run it in a full-screen or kiosk mode for the touchscreen.

## **Screenshots**

<img width="891" height="725" alt="main_screen" src="https://github.com/user-attachments/assets/1109b487-f85d-4021-b370-b269951c07dc" />

<img width="887" height="722" alt="bode" src="https://github.com/user-attachments/assets/8a963a24-4006-470a-942a-2e53384f2c91" />

<img width="892" height="723" alt="nyquist" src="https://github.com/user-attachments/assets/b809d4be-8fbb-4337-9f3d-bf2f5633a7ef" />

<img width="896" height="723" alt="output_log" src="https://github.com/user-attachments/assets/c11a8867-0b33-4a97-8474-0b416aad61e7" />
