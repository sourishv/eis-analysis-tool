# **Potentiostat EIS GUI (PC Test Harness)**

This project is a Python-based PC application for simulating and visualizing Electrochemical Impedance Spectroscopy (EIS) data. It serves as a **PC test harness** for developing and refining the user interface and features before deploying the final application to a Raspberry Pi with a real potentiostat.

This application is built with tkinter, ttk, matplotlib, and numpy.

## **Key Features**

* **Simulated EIS Measurement:** Runs a complete, simulated EIS test thread, including a mock OCP (Open Circuit Potential) wait time.  
* **Realistic Fake Data:** Generates complex, multi-segment "fake" data by interpolating between key breakpoints in log-log space, precisely matching the shape of a target Bode plot.  
* **Interactive Plotting:**  
  * **Bode Plot:** Displays Impedance Magnitude |Z| vs. Frequency with a log-log scale.  
  * **Nyquist Plot:** Displays Z\_real vs. \-Z\_imaginary with a uniform, realistic semi-circular shape.  
  * **Hover Tooltips:** Hovering the mouse over any data point on either plot reveals a tooltip with its exact (Freq, |Z|) or (Z\_real, \-Z\_imag) coordinates.  
* **Visual Diagnosis:**  
  * The Bode plot features a permanent color bar on the left (Green, Yellow, Red) to indicate coating health based on impedance magnitude.  
  * The Output Log provides a text-based diagnosis (e.g., "Healthy Coating," "Coating needs monitoring," "Defective Coating") based on the low-frequency impedance value.  
* **Data Export:** Both the Nyquist and Bode plots can be exported and saved as high-quality .png, .jpg, or .pdf files.  
* **Responsive UI:** Features a tabbed notebook interface to cleanly separate the measurement setup, plots, and output log. A simulated connection bar shows the device status.

## **Screenshot**

*\[A screenshot of the application's main window, showing the tabs and plots, would go here.\]*

## **Setup and Installation**

This project is intended to be run from a virtual environment.

1. **Clone the repository:**  
   git clone \[Your-Repository-URL\]  
   cd \[Your-Repository-Name\]

2. **Create and activate a virtual environment:**  
   \# Create the venv  
   python \-m venv .venv

   \# Activate on Windows (PowerShell)  
   .\\.venv\\Scripts\\Activate.ps1

   \# Activate on macOS/Linux  
   source .venv/bin/activate

3. Install the required libraries:  
   First, ensure you have a requirements.txt file. You can generate one from your working environment using:  
   pip freeze \> requirements.txt

   Then, anyone (including your future self on the Pi) can install the dependencies:  
   pip install \-r requirements.txt

   *If you do not have a requirements.txt file, manually install the necessary packages:*  
   pip install matplotlib numpy

## **How to Run the Test App**

With your virtual environment active and packages installed, simply run the PC test script:

python pc\_potentiostat\_gui\_test.py

1. The application will launch and auto-connect to the "Mock Device."  
2. All EIS parameters are pre-filled with test values. Click **"Run EIS Measurement"** to start the simulation.  
3. The app will switch to the "Output Log" tab and show the process (OCP, scan, diagnosis).  
4. When complete, the app will switch to the "Bode Plot" tab to display the results.  
5. You can now switch between the "Nyquist Plot" and "Bode Plot" tabs, hover over data points, and use the "Save Plot" buttons to export the images.

## **Future Goals (Raspberry Pi Deployment)**

The ultimate goal of this project is to run on a Raspberry Pi. The next steps will involve:

1. **Hardware Integration:** Replace the simulated connection and generate\_fake\_data function with the actual pypalmsens library to connect to the potentiostat via USB.  
2. **Data Acquisition:** Call the real manager.measure() function and retrieve the *actual* data.  
3. **Porting:** Transfer the pc\_potentiostat\_gui\_test.py script (renamed to app.py or similar) to the Raspberry Pi, install the dependencies, and run it in a full-screen or kiosk mode for the touchscreen.