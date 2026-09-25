# Argus

Argus is a toolkit of 6 small programs for PCC verification work — downloading photos, checking application status, merging documents, and more. This guide explains how to set it up, step by step, even if you've never used Git or a terminal before.

## Step 1: Install two programs (one-time)

1. Search for **PowerShell** in your Start menu and open it. A blue window will appear — this is where you type commands. (Every Windows computer already has this, nothing to download for this part.)
2. Type this and press Enter:
   ```
   winget install --id Git.Git -e
   ```
3. Then type this and press Enter:
   ```
   winget install --id Python.Python.3.12 -e
   ```
   If either one asks you to agree to something, type `Y` and press Enter.
4. Close this PowerShell window once both finish.

If Git and Python are already installed on your computer, skip this step.

## Step 2: Download Argus

1. Search for **Git Bash** in your Start menu and open it. A black window will appear — this is where you type commands.
2. Type this and press Enter:
   ```
   git clone https://github.com/2006-Z/Argus.git
   ```
   This downloads Argus into a new folder named `Argus`.
3. Type this and press Enter, to go inside that folder:
   ```
   cd Argus
   ```

## Step 3: Run the setup

Type this and press Enter:
```
./start
```
This installs everything Argus needs to work. It can take a few minutes — just wait until you see the message **"Setup complete."**

## Step 4: Start using Argus

1. Close this Git Bash window completely.
2. Open a **new** Git Bash window (or Command Prompt, or PowerShell).
3. Type:
   ```
   argus
   ```
   and press Enter. Argus will open.

From now on, you can open Argus from anywhere, anytime, just by typing `argus`.

## If something ever goes wrong

Go back into the `Argus` folder and run `./start` again. It checks itself and fixes anything that's broken.
