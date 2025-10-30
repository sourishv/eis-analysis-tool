Welcome! This file explains how to get the pi-eis-analyzer project running
on your computer for the first time, and how to save your changes to GitHub.

This guide assumes you are on a Windows PC with VS Code already installed.

PART 1: HOW TO RUN THE APP FOR THE FIRST TIME

This is a one-time setup process to get the project and all its
dependencies installed on your computer.

STEP 1: INSTALL REQUIRED SOFTWARE (PYTHON & GIT)

1. Install Python (v3.11.9):

Go to: https://www.python.org/downloads/release/python-3119/

Scroll down and click on "Windows installer (64-bit)".

Run the installer file you just downloaded.

IMPORTANT: On the first page of the installer, check the box at the bottom that says "Add python.exe to PATH". This is critical.

Click "Install Now" and finish the installation.

2. Install Git:

Go to: https://git-scm.com/downloads

Download and run the installer. You can safely click "Next" through all the default options.

(Why? Git is the tool that downloads ("clones") the project from GitHub. It is required for Part 2 if you want to upload ("push") your own changes.)

3. Restart VS Code:

Completely close and re-open VS Code to make sure it finds the new software you just installed.

STEP 2: GET THE PROJECT FILES (CLONE THE REPO)

1. Open a Terminal in VS Code:

At the top of the VS Code window, click Terminal > New Terminal.

You will now have a command prompt (likely PowerShell) at the bottom of your screen.

2. Navigate to where you store projects:

Type cd (change directory) to where you want to store the project.

Example: cd C:\Users\YourName\Documents\GitHub

3. Clone the project:

Copy and paste the following command into your terminal and press Enter.

git clone https://github.com/YourName/pi-eis-analyzer.git

(Replace "YourName/pi-eis-analyzer.git" with your repository's actual URL)

4. Open the project in VS Code:

In VS Code, go to File > Open Folder...

Find and select the new pi-eis-analyzer folder that you just downloaded.

VS Code will reload and open the project.

If it asks "Do you trust the authors...?", click "Yes".

(Alternative - Not Recommended):

If you refuse to install Git and only want to run the app (never update it or push changes), you can go to the GitHub repo URL, click the green <> Code button, and select "Download ZIP". Unzip this file and open that folder in VS Code.

STEP 3: SET UP THE VIRTUAL ENVIRONMENT

You only need to do this once.

1. Re-open the Terminal:

Go to Terminal > New Terminal again. It should now be in your project folder.

2. Create the virtual environment:

python -m venv .venv

(This uses the Python 3.11.9 you installed)

3. Activate the virtual environment:

.\.venv\Scripts\Activate.ps1

Your terminal prompt should now have (.venv) at the beginning.

TROUBLESHOOTING (Permission Error):

If you get a red error about "execution of scripts is disabled", run this one-time command and then try activating again:

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

STEP 4: INSTALL LIBRARIES & RUN THE APP

1. Install all libraries:

Make sure your venv is active ((.venv) is visible).

pip install -r requirements.txt

2. Run the application:

python app.py

The app should now launch!

To run the app in the future: You only need to repeat STEP 3 (item 3, activate) and STEP 4 (item 2, run).

PART 2: HOW TO SAVE & UPLOAD YOUR CHANGES (PUSH TO GITHUB)

This part requires you to have installed Git and used git clone in Step 2.

After you've made edits to the code, here is how you save them back to GitHub.

THE 3-STEP PROCESS (ADD, COMMIT, PUSH)

1. Stage Your Changes (Add):

In your VS Code terminal (with the venv active), type:

git add .

(The . means "add all files I have changed")

2. Save Your Changes (Commit):

This takes a "snapshot" of your staged files.

git commit -m "A short message describing your change"

Example: git commit -m "Updated the Bode plot labels"

3. Upload Your Changes (Push):

This sends your saved "commit" up to GitHub.

git push

Note: The first time you push, GitHub may ask you to log in.

Pro-Tip: Check Your Work

Before you add and commit, it's a good idea to see what you've changed.
Run this command at any time:

git status

It will show you a list of all the files you have modified.