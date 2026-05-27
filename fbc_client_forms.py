"""
FBC Client Forms Manager  — v3  (Google Sheets Sync Edition)
─────────────────────────────────────────────────────────────
Standalone desktop app for sending VFEX / ZSE account-opening
forms to clients via Outlook, tracking replies, firing
reminders, and marking clients as completed.

Shared real-time data via Google Sheets (free).

Requirements:
    pip install google-auth google-auth-httplib2 google-api-python-client --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org

Build EXE:
    pyinstaller --onefile --windowed --name "FBC-Client-Forms" fbc_client_forms.py
"""

# ════════════════════════════════════════════════════════════════════════════
#  AUTO-UPDATE
# ════════════════════════════════════════════════════════════════════════════
import sys, os, subprocess, urllib.request, threading

VERSION       = 5                         # ← bump this each release, then rebuild EXE
GITHUB_USER   = "Anashe-Masomeke"
GITHUB_REPO   = "fbc-client-forms"
GITHUB_BRANCH = "main"
EXE_NAME      = "FBC-Client-Forms.exe"   # ← must match the asset name in the GitHub Release exactly

_EXE = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download/{EXE_NAME}"
_VER = (f"https://raw.githubusercontent.com/"
        f"{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/version.txt")


def _remote_ver():
    """Fetch the plain-integer version from version.txt on GitHub main branch.
    Returns -1 if the file is missing, unreachable, or not a valid integer."""
    try:
        with urllib.request.urlopen(_VER, timeout=6) as r:
            return int(r.read().decode().strip())
    except Exception:
        return -1


def check_and_apply_update():
    rv = _remote_ver()
    if rv <= VERSION:
        return   # already up-to-date, or couldn't reach GitHub — silent skip

    import tkinter as tk
    from tkinter import messagebox, ttk

    # ── Ask the user ────────────────────────────────────────────────────────
    root = tk.Tk()
    root.withdraw()
    ok = messagebox.askyesno(
        "FBC Client Forms — Update Available",
        f"A new version is available  (v{rv}).\n"
        f"Your version: v{VERSION}\n\n"
        "Download and install now?\n\n"
        "(The new version will open automatically when the download finishes.)",
        icon="info")
    root.destroy()
    if not ok:
        return

    # ── File paths ──────────────────────────────────────────────────────────
    current_exe  = os.path.abspath(sys.argv[0])
    exe_dir      = os.path.dirname(current_exe)
    new_exe_path = os.path.join(exe_dir, f"FBC-Client-Forms-v{rv}.exe")
    bat_path     = os.path.join(exe_dir, "_fbc_cf_updater.bat")

    # ── Progress window ─────────────────────────────────────────────────────
    prog = tk.Tk()
    prog.title("Downloading Update…")
    prog.resizable(False, False)
    prog.attributes("-topmost", True)
    w, h = 420, 115
    prog.geometry(
        f"{w}x{h}+{(prog.winfo_screenwidth()-w)//2}+{(prog.winfo_screenheight()-h)//2}")

    tk.Label(prog, text=f"Downloading FBC Client Forms v{rv}…",
             font=("Segoe UI", 10, "bold"), pady=12).pack()
    bar = ttk.Progressbar(prog, mode="indeterminate", length=360)
    bar.pack(padx=30)
    bar.start(12)
    lbl_kb = tk.Label(prog, text="Starting…", font=("Segoe UI", 8), fg="#607080")
    lbl_kb.pack(pady=6)
    prog.update()

    error_holder = [None]

    def _do_download():
        try:
            MIN_SIZE = 5 * 1024 * 1024   # 5 MB — adjust if your EXE grows larger
            downloaded = 0
            with urllib.request.urlopen(_EXE, timeout=180) as resp:
                with open(new_exe_path, "wb") as f:
                    while True:
                        chunk = resp.read(65536)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        try:
                            lbl_kb.config(text=f"Downloaded: {downloaded // 1024:,} KB")
                        except Exception:
                            pass

            size = os.path.getsize(new_exe_path)
            if size < MIN_SIZE:
                os.remove(new_exe_path)
                raise Exception(
                    f"Download seems incomplete ({size // 1024} KB).\n"
                    "Please check your internet connection and try again.")

            # Write launcher bat: waits for old process to close, starts new EXE, self-deletes
            bat_lines = [
                "@echo off",
                "ping 127.0.0.1 -n 4 > nul",
                f'start "" "{new_exe_path}"',
                "ping 127.0.0.1 -n 2 > nul",
                'del "%~f0"',
            ]
            with open(bat_path, "w") as f:
                f.write("\n".join(bat_lines) + "\n")

        except Exception as e:
            error_holder[0] = str(e)
            for fp in [new_exe_path, bat_path]:
                try:
                    os.remove(fp)
                except Exception:
                    pass
        finally:
            try:
                prog.after(0, prog.quit)
            except Exception:
                pass

    t = threading.Thread(target=_do_download, daemon=True)
    t.start()
    prog.mainloop()   # blocks here until download thread calls prog.quit()
    prog.destroy()
    t.join()

    # ── Handle result ────────────────────────────────────────────────────────
    if error_holder[0]:
        err_root = tk.Tk()
        err_root.withdraw()
        messagebox.showerror(
            "Update Failed",
            f"Could not update:\n\n{error_holder[0]}\n\n"
            "Please download manually from:\n"
            f"github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest")
        err_root.destroy()
        return   # continue launching the current (old) version

    # ── Launch new EXE via bat and exit current process ──────────────────────
    subprocess.Popen(
        ["cmd.exe", "/c", bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True)
    sys.exit(0)


# ════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ════════════════════════════════════════════════════════════════════════════
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, date

# ════════════════════════════════════════════════════════════════════════════
#  COLOURS
# ════════════════════════════════════════════════════════════════════════════
FBC_DARK        = "#003B6F"
FBC_MID         = "#0066B3"
FBC_ACCENT      = "#00A3E0"
GREEN_DARK      = "#1A6B3A"
GREEN_LIGHT_BG  = "#F0FFF4"
AMBER           = "#B45309"
AMBER_BG        = "#FFF8E7"
WHITE           = "#FFFFFF"
BG              = "#F0F4F8"
CARD_BG         = "#FFFFFF"
SEP_CLR         = "#D0DAE8"
SIDEBAR_BG      = "#001F3F"
SIDEBAR_TEXT    = "#B0C8E8"

# ════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ════════════════════════════════════════════════════════════════════════════
APP_PASSWORD      = "anesu"
MAX_ATTEMPTS      = 6

DEFAULT_VFEX_FORM = ""
DEFAULT_ZSE_FORM  = ""

# Local fallback cache (used when offline)
STATE_FILE    = os.path.join(os.path.expanduser("~"), ".fbc_client_forms.json")
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".fbc_cf_settings.json")

# Google Sheets config
SHEET_NAME    = "FBC_Clients"   # Tab name inside the spreadsheet
SHEET_RANGE   = "FBC_Clients"   # Used for read/write range

# Sheet columns (A→K)
COL_HEADERS = [
    "id", "name", "email", "form_type", "sent_date",
    "sent_by", "reminders", "done", "done_date", "reminder_dates", "notes"
]

# Email templates
SUBJ_SEND = "Account Opening Forms — Action Required"
BODY_VFEX = (
    "Dear {{client}},\r\n\r\n"
    "Please find attached the VFEX account opening form for your review.\r\n\r\n"
    "Kindly complete and return the signed form within 2 business days.\r\n\r\n"
    "Should you have any questions, please do not hesitate to contact me.\r\n\r\n"
    "Regards,\r\n{{sender}}\r\nFBC Securities"
)
BODY_ZSE = (
    "Dear {{client}},\r\n\r\n"
    "Please find attached the ZSE account opening form for your review.\r\n\r\n"
    "Kindly complete and return the signed form within 2 business days.\r\n\r\n"
    "Should you have any questions, please do not hesitate to contact me.\r\n\r\n"
    "Regards,\r\n{{sender}}\r\nFBC Securities"
)
BODY_BOTH = (
    "Dear {{client}},\r\n\r\n"
    "Please find attached the VFEX and ZSE account opening forms for your review.\r\n\r\n"
    "Kindly complete and return the signed forms within 2 business days.\r\n\r\n"
    "Should you have any questions, please do not hesitate to contact me.\r\n\r\n"
    "Regards,\r\n{{sender}}\r\nFBC Securities"
)
SUBJ_REMINDER = "Reminder — Outstanding Account Opening Forms"
BODY_REMINDER = (
    "Dear {{client}},\r\n\r\n"
    "This is a friendly reminder that we are still awaiting your completed "
    "account opening form(s).\r\n\r\n"
    "Please return the signed form(s) at your earliest convenience so we can "
    "proceed with opening your account without further delay.\r\n\r\n"
    "Regards,\r\n{{sender}}\r\nFBC Securities"
)


# ════════════════════════════════════════════════════════════════════════════
#  SETTINGS
# ════════════════════════════════════════════════════════════════════════════
def load_settings():
    try:
        with open(SETTINGS_FILE) as f:
            return json.load(f)
    except Exception:
        return {"sender_name": "", "sheet_id": "", "key_file": ""}

def save_settings(s):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(s, f, indent=2)


# ════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS BACKEND
# ════════════════════════════════════════════════════════════════════════════
_sheets_service = None   # global, initialised once

def _get_service(key_file_path):
    """Build and cache the Sheets API service object."""
    global _sheets_service
    if _sheets_service:
        return _sheets_service
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = service_account.Credentials.from_service_account_file(
        key_file_path, scopes=SCOPES)
    _sheets_service = build("sheets", "v4", credentials=creds)
    return _sheets_service


def _ensure_header(service, sheet_id):
    """Make sure the header row exists; create sheet tab if missing."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"{SHEET_NAME}!A1:K1"
        ).execute()
        vals = result.get("values", [])
        if not vals or vals[0] != COL_HEADERS:
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"{SHEET_NAME}!A1",
                valueInputOption="RAW",
                body={"values": [COL_HEADERS]}
            ).execute()
    except Exception as e:
        err = str(e)
        # Tab doesn't exist — create it
        if "Unable to parse range" in err or "notFound" in err.lower():
            body = {"requests": [{"addSheet": {"properties": {"title": SHEET_NAME}}}]}
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id, body=body).execute()
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"{SHEET_NAME}!A1",
                valueInputOption="RAW",
                body={"values": [COL_HEADERS]}
            ).execute()


def _row_to_client(row):
    """Convert a sheet row (list of strings) to a client dict."""
    def g(i, default=""):
        return row[i] if i < len(row) else default
    return {
        "id":             g(0),
        "name":           g(1),
        "email":          g(2),
        "form_type":      g(3),
        "sent_date":      g(4),
        "sent_by":        g(5),
        "reminders":      int(g(6, "0") or "0"),
        "done":           g(7, "FALSE").upper() == "TRUE",
        "done_date":      g(8),
        "reminder_dates": g(9),
        "notes":          g(10),
    }

def _client_to_row(c):
    return [
        c.get("id", ""),
        c.get("name", ""),
        c.get("email", ""),
        c.get("form_type", ""),
        c.get("sent_date", ""),
        c.get("sent_by", ""),
        str(c.get("reminders", 0)),
        "TRUE" if c.get("done") else "FALSE",
        c.get("done_date", ""),
        c.get("reminder_dates", ""),
        c.get("notes", ""),
    ]


class SheetsDB:
    """
    Thin wrapper around Google Sheets API.
    Falls back to local JSON cache when offline.
    """
    def __init__(self, key_file, sheet_id):
        self.key_file  = key_file
        self.sheet_id  = sheet_id
        self._online   = False
        self._service  = None
        self._connect()

    def _connect(self):
        try:
            self._service = _get_service(self.key_file)
            _ensure_header(self._service, self.sheet_id)
            self._online = True
        except Exception as e:
            self._online = False
            print(f"[SheetsDB] Offline — {e}")

    @property
    def online(self):
        return self._online

    def read_all(self):
        """Return list of client dicts from Sheet (or local cache)."""
        if self._online:
            try:
                result = self._service.spreadsheets().values().get(
                    spreadsheetId=self.sheet_id,
                    range=f"{SHEET_NAME}!A2:K"
                ).execute()
                rows = result.get("values", [])
                clients = [_row_to_client(r) for r in rows if r and r[0]]
                # Update local cache
                with open(STATE_FILE, "w") as f:
                    json.dump(clients, f, indent=2)
                return clients
            except Exception as e:
                print(f"[SheetsDB] read failed: {e}")
                self._online = False
        # Fallback to local cache
        try:
            with open(STATE_FILE) as f:
                return json.load(f)
        except Exception:
            return []

    def _find_row_num(self, client_id):
        """Return 1-based row number for a given client id, or None."""
        try:
            result = self._service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id,
                range=f"{SHEET_NAME}!A:A"
            ).execute()
            col = result.get("values", [])
            for i, cell in enumerate(col):
                if cell and cell[0] == client_id:
                    return i + 1   # 1-based
        except Exception:
            pass
        return None

    def append_client(self, client):
        """Add a new client row to the Sheet."""
        if self._online:
            try:
                self._service.spreadsheets().values().append(
                    spreadsheetId=self.sheet_id,
                    range=f"{SHEET_NAME}!A1",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": [_client_to_row(client)]}
                ).execute()
                return True
            except Exception as e:
                print(f"[SheetsDB] append failed: {e}")
                self._online = False
        return False

    def update_client(self, client):
        """Overwrite an existing row for client['id']."""
        if self._online:
            try:
                row_num = self._find_row_num(client["id"])
                if row_num:
                    self._service.spreadsheets().values().update(
                        spreadsheetId=self.sheet_id,
                        range=f"{SHEET_NAME}!A{row_num}:K{row_num}",
                        valueInputOption="RAW",
                        body={"values": [_client_to_row(client)]}
                    ).execute()
                    return True
            except Exception as e:
                print(f"[SheetsDB] update failed: {e}")
                self._online = False
        return False

    def delete_client(self, client_id):
        """Delete the row for client_id."""
        if self._online:
            try:
                row_num = self._find_row_num(client_id)
                if row_num:
                    # Get sheet gid (tab id)
                    meta = self._service.spreadsheets().get(
                        spreadsheetId=self.sheet_id).execute()
                    sheet_gid = None
                    for s in meta["sheets"]:
                        if s["properties"]["title"] == SHEET_NAME:
                            sheet_gid = s["properties"]["sheetId"]
                            break
                    if sheet_gid is not None:
                        self._service.spreadsheets().batchUpdate(
                            spreadsheetId=self.sheet_id,
                            body={"requests": [{"deleteDimension": {
                                "range": {
                                    "sheetId":    sheet_gid,
                                    "dimension":  "ROWS",
                                    "startIndex": row_num - 1,
                                    "endIndex":   row_num
                                }
                            }}]}
                        ).execute()
                        return True
            except Exception as e:
                print(f"[SheetsDB] delete failed: {e}")
                self._online = False
        return False


# ════════════════════════════════════════════════════════════════════════════
#  LOCAL HELPERS (used as fallback / in-memory cache)
# ════════════════════════════════════════════════════════════════════════════
import uuid as _uuid

def new_id():
    return str(_uuid.uuid4())[:8]

def load_clients_local():
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except Exception:
        return []

def save_clients_local(clients):
    with open(STATE_FILE, "w") as f:
        json.dump(clients, f, indent=2)

def days_since(date_str):
    try:
        return (date.today() - datetime.strptime(date_str, "%Y-%m-%d").date()).days
    except Exception:
        return 0

def status_of(c):
    if c.get("done"):
        return "Completed", GREEN_DARK, "#EAF7EF"
    d = days_since(c.get("sent_date", ""))
    if d >= 2:
        return f"Overdue ({d}d)", "#B71C1C", "#FDECEA"
    return "Awaiting reply", FBC_MID, "#EAF4FB"

def open_outlook(to, subject, body, attachments=None):
    try:
        import win32com.client as win32
        outlook = win32.Dispatch("outlook.application")
        mail    = outlook.CreateItem(0)
        mail.To      = to
        mail.Subject = subject
        mail.Body    = body
        for fp in (attachments or []):
            if fp and os.path.exists(fp):
                mail.Attachments.Add(fp)
        mail.Display(True)
        return True, ""
    except ImportError:
        return False, "pywin32 not installed.\nRun:  pip install pywin32"
    except Exception as e:
        return False, str(e)

def fill_template(template, client_name, sender_name):
    return (template
            .replace("{{client}}", client_name)
            .replace("{{sender}}", sender_name or "FBC Securities"))


# ════════════════════════════════════════════════════════════════════════════
#  REUSABLE WIDGET HELPERS
# ════════════════════════════════════════════════════════════════════════════
def flat_entry(parent, var, **kw):
    return tk.Entry(parent, textvariable=var, font=("Segoe UI", 10),
                    bg="#F7FAFC", fg="#1A2B3C", relief="flat",
                    highlightbackground=SEP_CLR, highlightthickness=1, **kw)

def flat_text(parent, height=7, **kw):
    return tk.Text(parent, font=("Segoe UI", 10),
                   bg="#F7FAFC", fg="#1A2B3C", relief="flat",
                   highlightbackground=SEP_CLR, highlightthickness=1,
                   height=height, wrap="word", **kw)

def section_label(parent, text):
    tk.Label(parent, text=text, bg=BG, fg=FBC_DARK,
             font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(12, 3))

def card_frame(parent, **kw):
    return tk.Frame(parent, bg=CARD_BG, padx=14, pady=10,
                    highlightbackground=SEP_CLR, highlightthickness=1, **kw)


# ════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS SETUP DIALOG  (shown on first run)
# ════════════════════════════════════════════════════════════════════════════
class SheetsSetupDialog(tk.Toplevel):
    """
    Shown when sheet_id or key_file is missing.
    Lets the user paste their Sheet ID and browse for the JSON key.
    """
    def __init__(self, parent, settings, on_save):
        super().__init__(parent)
        self.settings = settings
        self.on_save  = on_save
        self.title("Google Sheets Setup")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.grab_set()
        self._sid  = tk.StringVar(value=settings.get("sheet_id", ""))
        self._key  = tk.StringVar(value=settings.get("key_file", ""))
        self._build()
        self.update_idletasks()
        w = 580
        # Cap height to 90% of screen so buttons always visible
        screen_h = self.winfo_screenheight()
        h = min(520, int(screen_h * 0.88))
        if parent:
            px = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
            py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        else:
            px = (self.winfo_screenwidth()  - w) // 2
            py = (self.winfo_screenheight() - h) // 2
        # Never position below 10px from top
        py = max(10, py)
        self.geometry(f"{w}x{h}+{max(px,0)}+{py}")

    def _build(self):
        tk.Frame(self, bg=FBC_ACCENT, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="☁  Connect to Google Sheets",
                 bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=16, anchor="w")
        tk.Label(hdr,
                 text="One-time setup — both users must complete this.",
                 bg=FBC_DARK, fg=SIDEBAR_TEXT,
                 font=("Segoe UI", 9)).pack(padx=16, anchor="w", pady=(0, 4))

        body = tk.Frame(self, bg=BG, padx=24, pady=12)
        body.pack(fill="both", expand=True)

        # Sheet ID
        tk.Label(body, text="Google Sheet ID", bg=BG, fg=FBC_DARK,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(body,
                 text="From the Sheet URL — the long code between /d/ and /edit",
                 bg=BG, fg="#607080",
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 4))
        flat_entry(body, self._sid).pack(fill="x", ipady=6, pady=(0, 14))

        # JSON key file
        tk.Label(body, text="Service Account JSON Key File", bg=BG, fg=FBC_DARK,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(body, text="The .json file downloaded from Google Cloud Console",
                 bg=BG, fg="#607080",
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 4))
        key_row = tk.Frame(body, bg=BG)
        key_row.pack(fill="x", pady=(0, 10))
        flat_entry(key_row, self._key).pack(side="left", fill="x",
                                             expand=True, ipady=6, padx=(0, 6))
        tk.Button(key_row, text="📂 Browse", font=("Segoe UI", 9),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK,
                  command=self._browse).pack(side="left")

        self._lbl_err = tk.Label(body, text="", bg=BG, fg="#B71C1C",
                                 font=("Segoe UI", 9))
        self._lbl_err.pack(anchor="w", pady=(0, 6))

        # Buttons INSIDE body so they are always visible
        btn_bar = tk.Frame(body, bg=BG)
        btn_bar.pack(fill="x", pady=(4, 0))
        tk.Button(btn_bar, text="Skip (offline mode)",
                  font=("Segoe UI", 9), bg=BG, fg="#607080",
                  relief="flat", cursor="hand2",
                  activebackground=SEP_CLR,
                  command=self._skip).pack(side="right", padx=(6, 0))
        tk.Button(btn_bar, text="  ✅  Connect & Save  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK,
                  command=self._save).pack(side="right")

    def _browse(self):
        p = filedialog.askopenfilename(
            title="Select Google Service Account JSON key",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if p:
            self._key.set(p)

    def _save(self):
        sid = self._sid.get().strip()
        key = self._key.get().strip()
        if not sid:
            self._lbl_err.config(text="❌  Please paste your Google Sheet ID.")
            return
        if not key or not os.path.exists(key):
            self._lbl_err.config(text="❌  Please select the JSON key file.")
            return
        self._lbl_err.config(text="⏳  Testing connection…")
        self.update()
        try:
            svc = _get_service(key)
            _ensure_header(svc, sid)
        except Exception as e:
            self._lbl_err.config(text=f"❌  Connection failed: {e}")
            global _sheets_service
            _sheets_service = None
            return
        self.settings["sheet_id"] = sid
        self.settings["key_file"] = key
        save_settings(self.settings)
        self.on_save(sid, key)
        self.destroy()

    def _skip(self):
        self.on_save("", "")
        self.destroy()


# ════════════════════════════════════════════════════════════════════════════
#  LOGIN DIALOG
# ════════════════════════════════════════════════════════════════════════════
class LoginDialog(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FBC Client Forms — Login")
        self.resizable(False, False)
        self.configure(bg=SIDEBAR_BG)
        self.authenticated = False
        self.sender_name   = ""
        self._attempts     = 0
        self._settings     = load_settings()
        self._build()
        self.update_idletasks()
        w, h = 400, 440
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build(self):
        hdr = tk.Frame(self, bg=FBC_ACCENT, pady=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text="FBC", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI", 20, "bold"), padx=12, pady=6).pack()
        tk.Label(hdr, text="Client Forms Manager", bg=FBC_ACCENT, fg=WHITE,
                 font=("Segoe UI", 11)).pack(pady=(2, 0))

        body = tk.Frame(self, bg=SIDEBAR_BG, padx=36, pady=20)
        body.pack(fill="both", expand=True)

        saved_name = self._settings.get("sender_name", "").strip()
        if saved_name:
            self._name_var = tk.StringVar(value=saved_name)
            greet = tk.Frame(body, bg="#0D2B4E", padx=12, pady=10,
                             highlightbackground=FBC_MID, highlightthickness=1)
            greet.pack(fill="x", pady=(0, 16))
            tk.Label(greet, text=f"👤  Welcome back, {saved_name}",
                     bg="#0D2B4E", fg=WHITE,
                     font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(greet,
                     text="Not you? Clear your name in Settings after login.",
                     bg="#0D2B4E", fg="#607080",
                     font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))
            name_entry = None
        else:
            tk.Label(body, text="Your Name", bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                     font=("Segoe UI", 10, "bold")).pack(anchor="w")
            tk.Label(body,
                     text="Appears in the email sign-off (Regards, ...)",
                     bg=SIDEBAR_BG, fg="#607080",
                     font=("Segoe UI", 8)).pack(anchor="w", pady=(0, 4))
            self._name_var = tk.StringVar()
            name_entry = tk.Entry(body, textvariable=self._name_var,
                                  font=("Segoe UI", 11), bg="#0D2B4E", fg=WHITE,
                                  insertbackground=WHITE, relief="flat",
                                  highlightbackground=FBC_MID, highlightthickness=1)
            name_entry.pack(fill="x", ipady=7, pady=(0, 16))
            name_entry.focus()

        tk.Label(body, text="Password", bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")

        pw_row = tk.Frame(body, bg=SIDEBAR_BG)
        pw_row.pack(fill="x", pady=(4, 0))

        self._pw   = tk.StringVar()
        self._show = False
        self._entry = tk.Entry(pw_row, textvariable=self._pw, show="●",
                               font=("Segoe UI", 11), bg="#0D2B4E", fg=WHITE,
                               insertbackground=WHITE, relief="flat",
                               highlightbackground=FBC_MID, highlightthickness=1)
        self._entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 4))

        self._eye = tk.Button(pw_row, text="👁", command=self._toggle,
                              bg="#0D2B4E", fg=SIDEBAR_TEXT, relief="flat",
                              font=("Segoe UI", 12), cursor="hand2",
                              activebackground=FBC_MID, activeforeground=WHITE,
                              padx=6)
        self._eye.pack(side="left")

        self._lbl_err = tk.Label(body, text="", bg=SIDEBAR_BG, fg="#FF6B6B",
                                 font=("Segoe UI", 9))
        self._lbl_err.pack(anchor="w", pady=(6, 0))
        self._lbl_att = tk.Label(body, text="", bg=SIDEBAR_BG, fg="#607080",
                                 font=("Segoe UI", 8))
        self._lbl_att.pack(anchor="w")

        tk.Button(body, text="  🔓  Login  ", command=self._attempt,
                  bg=FBC_MID, fg=WHITE, relief="flat",
                  font=("Segoe UI", 11, "bold"), cursor="hand2",
                  pady=10, activebackground=FBC_ACCENT).pack(fill="x", pady=(14, 0))

        if name_entry:
            name_entry.bind("<Return>", lambda _: self._attempt())
        else:
            self._entry.focus()
        self._entry.bind("<Return>", lambda _: self._attempt())

        # Sync status hint
        sid = self._settings.get("sheet_id", "")
        sync_txt  = "☁  Cloud sync: configured" if sid else "⚠  Cloud sync: not set up yet"
        sync_clr  = "#6EE7B7" if sid else "#FFB347"
        tk.Label(body, text=sync_txt, bg=SIDEBAR_BG, fg=sync_clr,
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(16, 0))

        tk.Label(self, text=f"v{VERSION}", bg=SIDEBAR_BG, fg="#2A4A6A",
                 font=("Segoe UI", 8)).pack(side="bottom", pady=6)

    def _toggle(self):
        self._show = not self._show
        self._entry.config(show="" if self._show else "●")
        self._eye.config(text="🙈" if self._show else "👁")

    def _attempt(self):
        sender = self._name_var.get().strip()
        if not sender:
            self._lbl_err.config(text="❌  Please enter your name.")
            return
        if self._pw.get().strip().lower() == APP_PASSWORD.lower():
            self.sender_name   = sender
            self.authenticated = True
            save_settings({**self._settings, "sender_name": sender})
            self.destroy()
            return
        self._attempts += 1
        rem = MAX_ATTEMPTS - self._attempts
        if rem <= 0:
            messagebox.showerror("Access Denied",
                "Too many incorrect attempts. The app will now close.")
            self.destroy()
            return
        self._lbl_err.config(text="❌  Incorrect password.")
        self._lbl_att.config(text=f"  {rem} attempt{'s' if rem>1 else ''} remaining")
        self._pw.set("")
        self._entry.focus()
        self._shake()

    def _shake(self, n=6, d=8):
        x0, y0 = self.winfo_x(), self.winfo_y()
        def step(i):
            if i == 0: self.geometry(f"+{x0}+{y0}"); return
            self.geometry(f"+{x0 + (d if i%2==0 else -d)}+{y0}")
            self.after(40, lambda: step(i-1))
        step(n)

    def _close(self):
        self.authenticated = False
        self.destroy()


# ════════════════════════════════════════════════════════════════════════════
#  SEND FORM DIALOG
# ════════════════════════════════════════════════════════════════════════════
class SendFormDialog(tk.Toplevel):
    def __init__(self, parent, sender_name, on_sent):
        super().__init__(parent)
        self.sender_name = sender_name
        self.on_sent     = on_sent
        self.title("Send Account Opening Form")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.grab_set()
        self._vfex   = tk.StringVar(value=DEFAULT_VFEX_FORM)
        self._zse    = tk.StringVar(value=DEFAULT_ZSE_FORM)
        self._canvas = None
        self._build()
        self.update_idletasks()
        w, h = 600, 700
        px = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _close(self):
        self._unbind_scroll()
        self.destroy()

    def _unbind_scroll(self):
        try:
            if self._canvas and self._canvas.winfo_exists():
                self._canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass

    def _build(self):
        tk.Frame(self, bg=FBC_ACCENT, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✉  Send Account Opening Form",
                 bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI", 12, "bold")).pack(padx=16, anchor="w")

        outer  = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        self._canvas = canvas
        sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(canvas, bg=BG, padx=20, pady=4)
        cid  = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(cid, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        sender_strip = tk.Frame(body, bg="#E8F1FB", padx=12, pady=7,
                                highlightbackground="#A8C4E0", highlightthickness=1)
        sender_strip.pack(fill="x", pady=(8, 0))
        tk.Label(sender_strip,
                 text=f"✏  Sending as:  {self.sender_name}",
                 bg="#E8F1FB", fg=FBC_DARK,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")

        section_label(body, "CLIENT DETAILS")
        c1 = card_frame(body); c1.pack(fill="x")
        self._name_var  = tk.StringVar()
        self._email_var = tk.StringVar()
        for lbl, var in [("Client name",  self._name_var),
                         ("Email address", self._email_var)]:
            row = tk.Frame(c1, bg=CARD_BG); row.pack(fill="x", pady=3)
            tk.Label(row, text=lbl, bg=CARD_BG, fg="#607080",
                     font=("Segoe UI", 9), width=14, anchor="w").pack(side="left")
            flat_entry(row, var).pack(side="left", fill="x", expand=True,
                                      ipady=5, padx=(4, 0))

        section_label(body, "FORM TYPE")
        c2 = card_frame(body); c2.pack(fill="x")
        self._ftype = tk.StringVar(value="VFEX")
        for val, txt in [("VFEX", "VFEX Account Opening"),
                         ("ZSE",  "ZSE Account Opening"),
                         ("BOTH", "VFEX + ZSE (Both)")]:
            tk.Radiobutton(c2, text=txt, variable=self._ftype, value=val,
                           command=self._on_type_change,
                           bg=CARD_BG, fg=FBC_DARK, selectcolor=CARD_BG,
                           font=("Segoe UI", 10), cursor="hand2",
                           activebackground=CARD_BG).pack(anchor="w", pady=2)

        section_label(body, "FORM ATTACHMENTS")
        c3 = card_frame(body); c3.pack(fill="x")
        def att_row(parent, lbl, var, attr):
            f = tk.Frame(parent, bg=CARD_BG); f.pack(fill="x", pady=3)
            tk.Label(f, text=lbl, bg=CARD_BG, fg="#607080",
                     font=("Segoe UI", 9), width=10, anchor="w").pack(side="left")
            tk.Label(f, textvariable=var, bg=CARD_BG, fg=FBC_MID,
                     font=("Segoe UI", 9), anchor="w").pack(
                         side="left", fill="x", expand=True, padx=4)
            tk.Button(f, text="📂 Override", font=("Segoe UI", 9),
                      bg=BG, fg=FBC_DARK, relief="flat", cursor="hand2",
                      command=lambda v=var: self._pick(v),
                      activebackground=SEP_CLR).pack(side="right")
            setattr(self, attr, f)
        att_row(c3, "VFEX form:", self._vfex, "_row_vfex")
        att_row(c3, "ZSE form:",  self._zse,  "_row_zse")

        section_label(body, "EMAIL MESSAGE")
        c4 = card_frame(body); c4.pack(fill="x")
        tk.Label(c4, text="Subject:", bg=CARD_BG, fg="#607080",
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._subj = tk.StringVar(value=SUBJ_SEND)
        flat_entry(c4, self._subj).pack(fill="x", ipady=5, pady=(2, 10))
        tk.Label(c4, text="Body:", bg=CARD_BG, fg="#607080",
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._body_txt = flat_text(c4, height=9)
        self._body_txt.pack(fill="x", pady=(2, 0))

        tk.Frame(body, bg=BG, height=8).pack()

        btn_bar = tk.Frame(self, bg=BG, padx=20, pady=10)
        btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="Cancel", font=("Segoe UI", 10),
                  bg=BG, fg="#607080", relief="flat", cursor="hand2",
                  command=self._close,
                  activebackground=SEP_CLR).pack(side="right", padx=(6, 0))
        tk.Button(btn_bar, text="  ✉  Open in Outlook  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK,
                  command=self._send).pack(side="right")

        self._on_type_change()

    def _on_type_change(self):
        t = self._ftype.get()
        if t == "VFEX":
            self._row_vfex.pack(fill="x", pady=3)
            self._row_zse.pack_forget()
            tmpl = BODY_VFEX
        elif t == "ZSE":
            self._row_vfex.pack_forget()
            self._row_zse.pack(fill="x", pady=3)
            tmpl = BODY_ZSE
        else:
            self._row_vfex.pack(fill="x", pady=3)
            self._row_zse.pack(fill="x", pady=3)
            tmpl = BODY_BOTH
        preview = tmpl.replace("{{sender}}", self.sender_name or "Your Name")
        self._body_txt.delete("1.0", "end")
        self._body_txt.insert("1.0", preview)

    def _pick(self, var):
        p = filedialog.askopenfilename(
            title="Select form PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if p:
            var.set(p)

    def _send(self):
        client_name = self._name_var.get().strip()
        email       = self._email_var.get().strip()
        if not client_name:
            messagebox.showwarning("Missing", "Please enter the client name.", parent=self)
            return
        if not email or "@" not in email:
            messagebox.showwarning("Missing", "Please enter a valid email address.", parent=self)
            return
        t   = self._ftype.get()
        sub = self._subj.get().strip()
        raw = self._body_txt.get("1.0", "end").strip()
        body = fill_template(raw, client_name, self.sender_name)
        attachments = []
        if t in ("VFEX", "BOTH"):
            vp = self._vfex.get().strip()
            if not vp or not os.path.exists(vp):
                messagebox.showwarning("No VFEX form",
                    "Please click '📂 Override' to attach the VFEX PDF.", parent=self)
                return
            attachments.append(vp)
        if t in ("ZSE", "BOTH"):
            zp = self._zse.get().strip()
            if not zp or not os.path.exists(zp):
                messagebox.showwarning("No ZSE form",
                    "Please click '📂 Override' to attach the ZSE PDF.", parent=self)
                return
            attachments.append(zp)
        ok, err = open_outlook(email, sub, body, attachments)
        if not ok:
            messagebox.showerror("Outlook error", err, parent=self)
            return
        self._unbind_scroll()
        self.on_sent({
            "id":          new_id(),
            "name":        client_name,
            "email":       email,
            "form_type":   t,
            "sent_date":   date.today().isoformat(),
            "sent_by":     self.sender_name,
            "reminders":   0,
            "done":        False,
            "done_date":   "",
            "reminder_dates": "",
            "notes":       "",
        })
        self.destroy()


# ════════════════════════════════════════════════════════════════════════════
#  REMINDER DIALOG
# ════════════════════════════════════════════════════════════════════════════
class ReminderDialog(tk.Toplevel):
    def __init__(self, parent, client, sender_name, on_sent):
        super().__init__(parent)
        self.client      = client
        self.sender_name = sender_name
        self.on_sent     = on_sent
        self.title("Send Reminder")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.grab_set()
        self._build()
        self.update_idletasks()
        w, h = 520, 430
        px = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

    def _build(self):
        tk.Frame(self, bg=AMBER, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"🔔  Reminder — {self.client['name']}",
                 bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI", 11, "bold")).pack(padx=16, anchor="w")
        body = tk.Frame(self, bg=BG, padx=20, pady=12)
        body.pack(fill="both", expand=True)
        days = days_since(self.client.get("sent_date", ""))
        info = tk.Frame(body, bg=AMBER_BG, padx=12, pady=8,
                        highlightbackground="#D4A017", highlightthickness=1)
        info.pack(fill="x", pady=(0, 10))
        tk.Label(info, text=f"⚠  Form sent {days} day(s) ago — no reply recorded.",
                 bg=AMBER_BG, fg=AMBER,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(info, text=f"To: {self.client['email']}",
                 bg=AMBER_BG, fg="#607080",
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(body, text="Subject:", bg=BG, fg="#607080",
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._subj = tk.StringVar(value=SUBJ_REMINDER)
        flat_entry(body, self._subj).pack(fill="x", ipady=5, pady=(2, 10))
        tk.Label(body, text="Message:", bg=BG, fg="#607080",
                 font=("Segoe UI", 9)).pack(anchor="w")
        self._msg = flat_text(body, height=9)
        self._msg.pack(fill="x", pady=(2, 0))
        preview = fill_template(BODY_REMINDER, self.client["name"], self.sender_name)
        self._msg.insert("1.0", preview)
        btn_bar = tk.Frame(self, bg=BG, padx=20, pady=10)
        btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="Cancel", font=("Segoe UI", 10),
                  bg=BG, fg="#607080", relief="flat", cursor="hand2",
                  command=self.destroy,
                  activebackground=SEP_CLR).pack(side="right", padx=(6, 0))
        tk.Button(btn_bar, text="  🔔  Open in Outlook  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=AMBER, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground="#92400E",
                  command=self._send).pack(side="right")

    def _send(self):
        raw  = self._msg.get("1.0", "end").strip()
        body = fill_template(raw, self.client["name"], self.sender_name)
        ok, err = open_outlook(self.client["email"],
                               self._subj.get().strip(), body)
        if not ok:
            messagebox.showerror("Outlook error", err, parent=self)
            return
        self.on_sent()
        self.destroy()


# ════════════════════════════════════════════════════════════════════════════
#  CONFIRM DIALOG
# ════════════════════════════════════════════════════════════════════════════
class ConfirmDialog(tk.Toplevel):
    def __init__(self, parent, client, on_confirm):
        super().__init__(parent)
        self.client     = client
        self.on_confirm = on_confirm
        self.title("Confirm — Forms Received")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.grab_set()
        self._build()
        self.update_idletasks()
        w, h = 420, 260
        px = parent.winfo_rootx() + (parent.winfo_width()  - w) // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

    def _build(self):
        tk.Frame(self, bg=GREEN_DARK, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✅  Confirm Forms Received",
                 bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI", 11, "bold")).pack(padx=16, anchor="w")
        body = tk.Frame(self, bg=BG, padx=24, pady=20)
        body.pack(fill="both", expand=True)
        tk.Label(body,
                 text="Mark this client as completed?\n"
                      "They will be removed from the active dashboard.",
                 bg=BG, fg="#374151",
                 font=("Segoe UI", 10), justify="left").pack(anchor="w")
        info = tk.Frame(body, bg=GREEN_LIGHT_BG, padx=12, pady=10,
                        highlightbackground="#6EE7B7", highlightthickness=1)
        info.pack(fill="x", pady=12)
        tk.Label(info, text=self.client["name"],
                 bg=GREEN_LIGHT_BG, fg=GREEN_DARK,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Label(info,
                 text=f"{self.client['email']}  ·  {self.client['form_type']}",
                 bg=GREEN_LIGHT_BG, fg="#607080",
                 font=("Segoe UI", 9)).pack(anchor="w")
        btn_bar = tk.Frame(self, bg=BG, padx=20, pady=10)
        btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="Cancel", font=("Segoe UI", 10),
                  bg=BG, fg="#607080", relief="flat", cursor="hand2",
                  command=self.destroy,
                  activebackground=SEP_CLR).pack(side="right", padx=(6, 0))
        tk.Button(btn_bar, text="  ✅  Confirm & Remove  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=GREEN_DARK, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground="#155724",
                  command=self._confirm).pack(side="right")

    def _confirm(self):
        self.on_confirm()
        self.destroy()


# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self, sender_name, settings):
        super().__init__()
        self.sender_name = sender_name
        self.settings    = settings
        self.title(f"FBC Client Forms Manager  v{VERSION}")
        self.state("zoomed")
        self.configure(bg=BG)

        # Initialise DB
        sid = settings.get("sheet_id", "")
        key = settings.get("key_file", "")
        if sid and key and os.path.exists(key):
            self.db = SheetsDB(sid, key)
        else:
            self.db = None

        self.clients = []
        self._filter  = "all"
        self._sync_job = None
        self._build()
        self._full_sync()          # Initial load
        self._schedule_auto_sync() # Poll every 30 s

    # ── build UI ──────────────────────────────────────────────────────────
    def _build(self):
        top = tk.Frame(self, bg=FBC_DARK)
        top.pack(fill="x")
        tk.Frame(top, bg=FBC_ACCENT, width=6).pack(side="left", fill="y")
        tk.Label(top, text="📋  FBC Client Forms Manager",
                 bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI", 14, "bold"),
                 pady=14, padx=16).pack(side="left")

        # Sync status badge
        self._sync_lbl = tk.Label(top, text="", bg=FBC_DARK,
                                  font=("Segoe UI", 9))
        self._sync_lbl.pack(side="left", padx=10)

        tk.Button(top, text="Change",
                  font=("Segoe UI", 8), bg=FBC_DARK, fg="#4A7AAA",
                  relief="flat", cursor="hand2",
                  activebackground=FBC_DARK, activeforeground=SIDEBAR_TEXT,
                  command=self._clear_name).pack(side="right", padx=(0, 4))
        tk.Label(top, text=f"👤  {self.sender_name}",
                 bg=FBC_DARK, fg=SIDEBAR_TEXT,
                 font=("Segoe UI", 9)).pack(side="right", padx=(0, 4))
        tk.Label(top, text=f"v{VERSION}", bg=FBC_DARK, fg="#2A5A8A",
                 font=("Segoe UI", 9)).pack(side="right", padx=10)

        tk.Button(top, text="☁  Sheets Setup",
                  font=("Segoe UI", 9),
                  bg=FBC_DARK, fg=SIDEBAR_TEXT, relief="flat", cursor="hand2",
                  activebackground=FBC_MID,
                  command=self._open_sheets_setup).pack(side="right", padx=4, pady=6)

        tk.Button(top, text="  ✉  Send New Form  ",
                  font=("Segoe UI", 10, "bold"),
                  bg=FBC_ACCENT, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_MID, pady=8, padx=4,
                  command=self._open_send).pack(side="right", padx=12, pady=6)

        # Metrics strip
        metric_bg  = tk.Frame(self, bg=SIDEBAR_BG, pady=10)
        metric_bg.pack(fill="x")
        metrics_row = tk.Frame(metric_bg, bg=SIDEBAR_BG)
        metrics_row.pack(padx=20)

        def mcard(lbl, colour):
            f = tk.Frame(metrics_row, bg="#002855", padx=20, pady=10,
                         highlightbackground="#0A3A6A", highlightthickness=1)
            f.pack(side="left", padx=(0, 10))
            tk.Label(f, text=lbl, bg="#002855", fg=SIDEBAR_TEXT,
                     font=("Segoe UI", 9)).pack(anchor="w")
            v = tk.Label(f, text="0", bg="#002855", fg=colour,
                         font=("Segoe UI", 22, "bold"))
            v.pack(anchor="w")
            return v

        self._m_total   = mcard("Total sent",        FBC_ACCENT)
        self._m_pending = mcard("Awaiting reply",     FBC_ACCENT)
        self._m_overdue = mcard("Overdue (>2 days)",  "#FF4444")
        self._m_done    = mcard("Completed",          "#6EE7B7")

        # Filter tabs
        tab_row = tk.Frame(self, bg=BG, padx=20, pady=8)
        tab_row.pack(fill="x")
        self._tabs = {}
        for key, lbl in [("all","All Clients"), ("overdue","Overdue"),
                          ("pending","Awaiting"), ("done","Completed")]:
            b = tk.Button(tab_row, text=lbl, font=("Segoe UI", 9),
                          relief="flat", cursor="hand2", padx=16, pady=6,
                          command=lambda k=key: self._set_filter(k))
            b.pack(side="left", padx=(0, 4))
            self._tabs[key] = b
        self._paint_tabs()

        tk.Frame(self, bg=SEP_CLR, height=1).pack(fill="x", padx=20)

        # Scrollable list
        list_outer = tk.Frame(self, bg=BG)
        list_outer.pack(fill="both", expand=True, padx=20, pady=10)
        self._canvas = tk.Canvas(list_outer, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(list_outer, orient="vertical",
                          command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._inner    = tk.Frame(self._canvas, bg=BG)
        self._inner_id = self._canvas.create_window(
            (0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._inner_id, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1*(e.delta//120), "units"))

    def _paint_tabs(self):
        for k, b in self._tabs.items():
            if k == self._filter:
                b.config(bg=FBC_DARK, fg=WHITE)
            else:
                b.config(bg=BG, fg="#607080",
                         activebackground=SEP_CLR, activeforeground=FBC_DARK)

    def _set_filter(self, key):
        self._filter = key
        self._paint_tabs()
        self._refresh()

    # ── sync ─────────────────────────────────────────────────────────────
    def _update_sync_badge(self):
        if self.db and self.db.online:
            self._sync_lbl.config(text="☁  Synced", fg="#6EE7B7")
        elif self.db:
            self._sync_lbl.config(text="⚠  Offline (local cache)", fg="#FFB347")
        else:
            self._sync_lbl.config(text="⚠  No sync configured", fg="#607080")

    def _full_sync(self):
        """Read all clients from Sheet (background thread)."""
        def _do():
            if self.db:
                clients = self.db.read_all()
            else:
                clients = load_clients_local()
            self.after(0, lambda: self._apply_clients(clients))
        threading.Thread(target=_do, daemon=True).start()

    def _apply_clients(self, clients):
        self.clients = clients
        self._update_sync_badge()
        self._refresh()

    def _schedule_auto_sync(self):
        """Auto-refresh every 30 seconds."""
        self._full_sync()
        self._sync_job = self.after(30_000, self._schedule_auto_sync)

    # ── refresh UI ───────────────────────────────────────────────────────
    def _refresh(self):
        all_c   = self.clients
        active  = [c for c in all_c if not c.get("done")]
        done    = [c for c in all_c if c.get("done")]
        overdue = [c for c in active if days_since(c.get("sent_date","")) >= 2]

        self._m_total.config(  text=str(len(all_c)))
        self._m_pending.config(text=str(len(active)))
        self._m_overdue.config(text=str(len(overdue)))
        self._m_done.config(   text=str(len(done)))

        if   self._filter == "overdue": visible = overdue
        elif self._filter == "pending": visible = active
        elif self._filter == "done":    visible = done
        else:                           visible = all_c

        for w in self._inner.winfo_children():
            w.destroy()

        if not visible:
            msg = ("No clients yet. Click  '✉ Send New Form'  to get started."
                   if self._filter == "all"
                   else "No clients in this category.")
            tk.Label(self._inner, text=msg, bg=BG, fg="#8096B0",
                     font=("Segoe UI", 11), pady=40).pack()
            return

        for client in visible:
            self._client_card(self._inner, client)

    def _client_card(self, parent, client):
        done = client.get("done", False)
        st, sc, sbg = status_of(client)
        days    = days_since(client.get("sent_date", ""))
        rems    = client.get("reminders", 0)
        sent_by = client.get("sent_by", "")

        card = tk.Frame(parent, bg=CARD_BG, padx=14, pady=12,
                        highlightbackground=SEP_CLR, highlightthickness=1)
        card.pack(fill="x", pady=4)

        initials = "".join(w[0] for w in client["name"].split()[:2]).upper()
        av_bg = GREEN_DARK if done else FBC_DARK
        tk.Label(card, text=initials, bg=av_bg, fg=WHITE,
                 font=("Segoe UI", 11, "bold"),
                 width=3, padx=4, pady=6).pack(side="left", padx=(0, 14))

        info = tk.Frame(card, bg=CARD_BG)
        info.pack(side="left", fill="x", expand=True)
        tk.Label(info, text=client["name"], bg=CARD_BG, fg=FBC_DARK,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w")
        meta = (f"{client['email']}   ·   {client['form_type']}"
                f"   ·   Sent: {client.get('sent_date','?')}"
                f"   ·   {days}d ago"
                + (f"   ·   By: {sent_by}" if sent_by else "")
                + (f"   ·   Reminders: {rems}" if rems else ""))
        tk.Label(info, text=meta, bg=CARD_BG, fg="#607080",
                 font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 4))
        tk.Label(info, text=f"  {st}  ", bg=sbg, fg=sc,
                 font=("Segoe UI", 8, "bold"),
                 padx=4, pady=2).pack(anchor="w")

        right = tk.Frame(card, bg=CARD_BG)
        right.pack(side="right")
        if not done:
            tk.Button(right, text="🔔  Reminder",
                      font=("Segoe UI", 9), bg=AMBER_BG, fg=AMBER,
                      relief="flat", cursor="hand2",
                      activebackground="#FDE68A",
                      command=lambda c=client: self._open_reminder(c)
                      ).pack(side="left", padx=(0, 6))
            tk.Button(right, text="✅  Confirm",
                      font=("Segoe UI", 9), bg=GREEN_LIGHT_BG, fg=GREEN_DARK,
                      relief="flat", cursor="hand2",
                      activebackground="#BBF7D0",
                      command=lambda c=client: self._open_confirm(c)
                      ).pack(side="left", padx=(0, 6))
        tk.Button(right, text="🗑",
                  font=("Segoe UI", 10), bg=CARD_BG, fg="#CBD5E1",
                  relief="flat", cursor="hand2",
                  activebackground=BG,
                  command=lambda c=client: self._delete(c)
                  ).pack(side="left")

    # ── actions ───────────────────────────────────────────────────────────
    def _open_sheets_setup(self):
        SheetsSetupDialog(self, self.settings, self._on_sheets_saved)

    def _on_sheets_saved(self, sid, key):
        if sid and key:
            self.settings["sheet_id"] = sid
            self.settings["key_file"] = key
            global _sheets_service
            _sheets_service = None
            self.db = SheetsDB(sid, key)
            self._full_sync()
        self._update_sync_badge()

    def _clear_name(self):
        if messagebox.askyesno("Change User",
            "This will clear your saved name.\n\n"
            "You will be asked for your name next time you open the app.\n\nContinue?",
            parent=self):
            save_settings({**self.settings, "sender_name": ""})
            messagebox.showinfo("Done",
                "Name cleared. Please restart the app.", parent=self)

    def _open_send(self):
        SendFormDialog(self, self.sender_name, self._on_sent)

    def _on_sent(self, rec):
        # Write to Sheet immediately (background)
        def _do():
            if self.db:
                self.db.append_client(rec)
            self.clients.append(rec)
            save_clients_local(self.clients)
            self.after(0, self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _open_reminder(self, client):
        ReminderDialog(self, client, self.sender_name,
                       lambda c=client: self._on_reminder(c))

    def _on_reminder(self, client):
        client["reminders"] = client.get("reminders", 0) + 1
        existing = client.get("reminder_dates", "")
        today = date.today().isoformat()
        client["reminder_dates"] = (existing + "," + today).strip(",")
        def _do():
            if self.db:
                self.db.update_client(client)
            save_clients_local(self.clients)
            self.after(0, self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _open_confirm(self, client):
        ConfirmDialog(self, client, lambda c=client: self._on_confirm(c))

    def _on_confirm(self, client):
        client["done"]      = True
        client["done_date"] = date.today().isoformat()
        def _do():
            if self.db:
                self.db.update_client(client)
            save_clients_local(self.clients)
            self.after(0, self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _delete(self, client):
        if not messagebox.askyesno("Delete record",
            f"Permanently delete the record for {client['name']}?\n\nThis cannot be undone.",
            parent=self):
            return
        def _do():
            if self.db:
                self.db.delete_client(client["id"])
            if client in self.clients:
                self.clients.remove(client)
            save_clients_local(self.clients)
            self.after(0, self._refresh)
        threading.Thread(target=_do, daemon=True).start()


# ════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    check_and_apply_update()

    login = LoginDialog()
    login.mainloop()

    if not login.authenticated:
        sys.exit(0)

    settings = load_settings()

    # Show Sheets setup if not yet configured
    if not settings.get("sheet_id") or not settings.get("key_file"):
        # Temporary root to host the setup dialog
        _tmp = tk.Tk(); _tmp.withdraw()
        _saved = {}
        def _on_setup(sid, key):
            _saved["sheet_id"] = sid
            _saved["key_file"]  = key
        dlg = SheetsSetupDialog(_tmp, settings, _on_setup)
        _tmp.wait_window(dlg)
        _tmp.destroy()
        if _saved:
            settings.update(_saved)
            save_settings(settings)

    app = App(sender_name=login.sender_name, settings=settings)
    app.mainloop()
