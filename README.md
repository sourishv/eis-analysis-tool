# Potentiostat EIS Software Analysis Tool (PC test)

This is a small PC testing GUI for Electrochemical Impedance Spectroscopy (EIS).
It is intended to exercise the UI and plotting workflows on a desktop before
moving the code to a Raspberry Pi and a real potentiostat.

The current app (run with `python app.py`) is centered around loading EIS
CSV files, plotting Nyquist and Bode views, exporting high-quality images,
and showing a simple text diagnosis based on low-frequency impedance.

Implemented with: tkinter, ttk, matplotlib, numpy, and pandas.

## What changed vs earlier drafts

- The app now loads EIS data from a CSV file (no hardware measurement/simulation
  button in this PC-test build).
- The UI tabs are: `Load Measurement`, `Nyquist Plot`, `Bode Plot`, `Output Log`.
- Plots are interactive (hover shows point info) and can be exported via
  "Save Nyquist Plot" / "Save Bode Plot" buttons.

## Key features

- Load EIS data from CSV and plot:
  - Nyquist: Z_real vs -Z_imaginary
  - Bode (magnitude): |Z| vs Frequency (log-log)
- Permanent color bar beside the Bode magnitude axis indicating coating health
  bands (red / yellow / green).
- Simple automated diagnosis based on the low-frequency |Z| value.
- Export plots as PNG / JPG / PDF from the GUI.

## Required CSV format

The app expects CSV files with these columns (exact header names required):

- `Frequency (Hz)`
- `Z' (Ω)`   (real part)
- `-Z'' (Ω)` (negative imaginary part as in many instrument exports)
- `Z (Ω)`    (magnitude)
- `-Phase (°)`
- `Time (s)`

The code computes Z_imag = -(`-Z'' (Ω)`) so the Nyquist plot uses `Z' (Ω)` vs
`-Z'' (Ω)` (flipped sign internally). If columns are missing the app will log
an error in the Output Log.

## Diagnosis thresholds

The built-in diagnosis (displayed in the Output Log) uses the low-frequency
magnitude value and the following thresholds:

- |Z| >= 1e7  → "Healthy Coating (Pass)"
- |Z| >= 1e5  → "Coating needs monitoring (Caution)"
- otherwise   → "Defective Coating, needs maintenance (Fail)"

These values are simple heuristics for the PC test app and can be adjusted in
`app.py::diagnose_coating()`.

## Setup and installation

1. Clone the repository:

```powershell
git clone https://github.com/sourishv/eis-analysis-tool.git
cd eis-analysis-tool
```

2. Create and activate a virtual environment (PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

3. Install dependencies:

```powershell
pip install -r requirements.txt
```

## Run the application

With the venv active:

```powershell
python app.py
```

Behavior:

- The app auto-starts a mock connection and enables the "Load EIS Data File"
  button when the (simulated) device is connected.
- Click `Load EIS Data File (.csv)` and select a CSV that matches the
  required column names above.
- The Output Log tab will show loading/diagnosis messages. When processing
  completes the Bode and Nyquist tabs are updated and the app switches to
  the Bode tab.
- Use the "Save Nyquist Plot" / "Save Bode Plot" buttons to export images.

## Troubleshooting

- Missing columns: the Output Log will show which required columns are not
  present. Make sure the CSV headers match exactly.
- No plots / blank: verify the CSV contains numeric data and that the venv is
  active so required packages (pandas, numpy, matplotlib) are available.
- Activation errors on Windows: when creating/activating `.venv` in PowerShell
  you may need to run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

## Exporting plots

- Use the save buttons on each plot tab; supported formats: PNG, JPG, PDF.
- Images are saved at 300 DPI by default.

## Future work / Raspberry Pi goals

- Replace the mock connection and CSV workflow with real hardware calls
  (for example, pypalmsens or other potentiostat libraries).
- Add a reproducible dependency lock (e.g., `requirements.lock` or `pip-tools`).
- Improve diagnosis algorithms and add unit tests for data parsing & plotting.
