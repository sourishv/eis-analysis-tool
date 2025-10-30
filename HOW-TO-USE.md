# How to use — pi-eis-analyzer

Welcome! This guide shows how to get the pi-eis-analyzer project running on your Windows PC (using VS Code) and how to save your changes back to GitHub.

> Prerequisite: Visual Studio Code is already installed.

Table of contents
- [Part 1 — Run the app for the first time](#part-1---run-the-app-for-the-first-time)
  - [Step 1 — Install required software (Python & Git)](#step-1---install-required-software-python--git)
  - [Step 2 — Get the project files (clone the repo)](#step-2---get-the-project-files-clone-the-repo)
  - [Step 3 — Set up the virtual environment](#step-3---set-up-the-virtual-environment)
  - [Step 4 — Install libraries & run the app](#step-4---install-libraries--run-the-app)
- [Part 2 — Save & upload your changes (push to GitHub)](#part-2---save--upload-your-changes-push-to-github)
  - [The 3-step process (add → commit → push)](#the-3-step-process-add--commit--push)
  - [Pro tip: check what changed](#pro-tip-check-what-changed)
- [Troubleshooting](#troubleshooting)
- [Quick command reference](#quick-command-reference)

---

## Part 1 — Run the app for the first time

This is a one-time setup process to install the project and its dependencies.

### Step 1 — Install required software (Python & Git)

1. Install Python 3.11.9:
   - Visit: https://www.python.org/downloads/release/python-3119/
   - Download **Windows installer (64-bit)** and run it.
   - IMPORTANT: On the first installer page, check **Add python.exe to PATH**.
   - Click **Install Now** and finish.

2. Install Git:
   - Visit: https://git-scm.com/downloads
   - Download and run the installer. Default options are fine.

Why Git? Git is used to clone the repository from GitHub and to push your changes later.

3. Restart VS Code
   - Close and re-open VS Code so it detects the newly installed tools.

---

### Step 2 — Get the project files (clone the repo)

1. Open a terminal in VS Code:  
   Terminal → New Terminal

2. Change to the directory where you store projects, for example:
   ```powershell
   cd C:\Users\YourName\Documents\GitHub
   ```

3. Clone the project:
   ```bash
   git clone https://github.com/YourName/pi-eis-analyzer.git
   ```
   Replace `https://github.com/YourName/pi-eis-analyzer.git` with the actual repository URL.

4. Open the project in VS Code:
   - File → Open Folder... → select the `pi-eis-analyzer` folder.
   - If prompted with "Do you trust the authors...?", click **Yes**.

Alternative (not recommended): If you prefer not to install Git and only want to run the app (no pushes), you can download the repository ZIP from GitHub (Code → Download ZIP) and extract it.

---

### Step 3 — Set up the virtual environment

You only need to do this once per machine/project.

1. Open a terminal in the project folder (Terminal → New Terminal).

2. Create the virtual environment:
   ```powershell
   python -m venv .venv
   ```

3. Activate the virtual environment:
   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```
   After activation, your prompt should start with `(.venv)`.

Troubleshooting: If you see an execution policy error (scripts disabled), run this one-time command in PowerShell and then try activation again:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

---

### Step 4 — Install libraries & run the app

1. With the venv active, install requirements:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the application:
   ```bash
   python app.py
   ```

The app should start. To run the app in the future, reactivate the venv (Step 3.3) and run `python app.py`.

---

## Part 2 — Save & upload your changes (push to GitHub)

This section assumes you installed Git and cloned the repo (Part 1, Step 2).

The typical workflow: edit files → stage → commit → push.

### The 3-step process (add → commit → push)

1. Stage your changes (add):
   ```bash
   git add .
   ```
   The `.` stages all modified files.

2. Commit your changes:
   ```bash
   git commit -m "A short message describing your change"
   ```
   Example:
   ```bash
   git commit -m "Updated the Bode plot labels"
   ```

3. Push to GitHub:
   ```bash
   git push
   ```
   The first time you push, GitHub may prompt you to log in.

### Pro tip: check what changed

Before staging and committing, see what’s different:
```bash
git status
```
This shows modified, staged, and untracked files.

---

## Troubleshooting

- Virtual environment activation failure:
  - Run PowerShell as administrator or use the `Set-ExecutionPolicy` command shown above.
- Wrong Python version:
  - Ensure `python --version` reports 3.11.9 (or the installed 3.11.x). If not, check PATH or the installer options.
- pip install errors:
  - Make sure the venv is active and you have internet access. For platform-specific build errors, read the error and install any missing system packages (e.g., build tools).
- git push issues:
  - If authentication fails, follow GitHub’s prompts to authenticate (PAT or browser login). Alternatively configure SSH.

---

## Quick command reference

- Create venv:
  ```bash
  python -m venv .venv
  ```
- Activate venv (PowerShell):
  ```powershell
  .\.venv\Scripts\Activate.ps1
  ```
- Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```
- Run app:
  ```bash
  python app.py
  ```
- Git basics:
  ```bash
  git add .
  git commit -m "message"
  git push
  git status
  ```
