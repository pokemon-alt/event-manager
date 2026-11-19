# Event Manager

This standalone copy includes the current Event Coordinator sheet, People and Addresses, Retail Accounts, Cards, imports, CSV/Excel/AYCD JSON exports, and optional AI editing. It starts with a fresh workspace. For AI editing, enter your API key in Settings.

## Run locally

Use Python 3.9 or newer. From this folder on macOS or Linux:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python server.py
```

On Windows:

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe server.py
```

Open [Event Manager](http://127.0.0.1:8789). Press Ctrl+C in the terminal to stop it.

The server listens on the computer running it. Use `--port 8770` if the default port is busy, or `--data-dir <folder>` to choose where it saves. Your encrypted workspace database and its key are saved together in `.event-manager` under your home folder by default. Keep both files together when backing up or moving the workspace.
