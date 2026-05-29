"""
FBC Client Forms Manager  — v28
─────────────────────────────────────────────────────────────
Changes in v28:
  - CC address is now a saved setting (not hardcoded)
  - CC field shown in Send Form and Reminder dialogs
  - CC can be overridden per-send (multiple addresses comma-separated)
  - Default CC editable via ⚙ Sheets Setup / Settings area
  - First-run default CC pre-filled from previous hardcoded value

All v27 features retained:
  - Forms bundled inside EXE via PyInstaller --add-data
  - Google Sheets live sync
  - Auto-update from GitHub releases
  - Overdue aging badges, Reply Received workflow
"""

import os, sys
import subprocess, urllib.request, threading

# ════════════════════════════════════════════════════════════════════════════
#  AUTO-UPDATE
# ════════════════════════════════════════════════════════════════════════════
VERSION       = 18
GITHUB_USER   = "Anashe-Masomeke"
GITHUB_REPO   = "client-forms"
GITHUB_BRANCH = "main"
EXE_NAME      = "FBC-Client-Forms.exe"

_EXE = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download/{EXE_NAME}"
_VER = (f"https://raw.githubusercontent.com/"
        f"{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/version.txt")

def _remote_ver():
    try:
        with urllib.request.urlopen(_VER, timeout=6) as r:
            return int(r.read().decode().strip())
    except Exception:
        return -1

def check_and_apply_update():
    rv = _remote_ver()
    if rv <= VERSION:
        return
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
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
    current_exe  = os.path.abspath(sys.argv[0])
    exe_dir      = os.path.dirname(current_exe)
    downloads    = os.path.join(os.path.expanduser("~"), "Downloads")
    save_dir     = downloads if os.path.isdir(downloads) else os.environ.get("TEMP", exe_dir)
    new_exe_path = os.path.join(save_dir, f"FBC-Client-Forms-v{rv}.exe")
    bat_path     = os.path.join(save_dir, "_fbc_cf_updater.bat")
    prog = tk.Tk()
    prog.title("Downloading Update…")
    prog.resizable(False, False)
    prog.attributes("-topmost", True)
    w, h = 420, 115
    prog.geometry(f"{w}x{h}+{(prog.winfo_screenwidth()-w)//2}+{(prog.winfo_screenheight()-h)//2}")
    tk.Label(prog, text=f"Downloading FBC Client Forms v{rv}…",
             font=("Segoe UI", 10, "bold"), pady=12).pack()
    import tkinter.ttk as ttk
    bar = ttk.Progressbar(prog, mode="indeterminate", length=360)
    bar.pack(padx=30); bar.start(12)
    lbl_kb = tk.Label(prog, text="Starting…", font=("Segoe UI", 8), fg="#607080")
    lbl_kb.pack(pady=6); prog.update()
    error_holder = [None]
    def _do_download():
        try:
            downloaded = 0
            with urllib.request.urlopen(_EXE, timeout=180) as resp:
                with open(new_exe_path, "wb") as f:
                    while True:
                        chunk = resp.read(65536)
                        if not chunk: break
                        f.write(chunk); downloaded += len(chunk)
                        try: lbl_kb.config(text=f"Downloaded: {downloaded//1024:,} KB")
                        except Exception: pass
            if os.path.getsize(new_exe_path) < 5*1024*1024:
                os.remove(new_exe_path)
                raise Exception("Download incomplete.")
            with open(bat_path, "w") as f:
                f.write("\n".join(["@echo off","ping 127.0.0.1 -n 6 > nul",
                    f'start "" "{new_exe_path}"',"ping 127.0.0.1 -n 2 > nul",'del "%~f0"',""]))
        except Exception as e:
            error_holder[0] = str(e)
            for fp in [new_exe_path, bat_path]:
                try: os.remove(fp)
                except: pass
        finally:
            try: prog.after(0, prog.quit)
            except: pass
    t = threading.Thread(target=_do_download, daemon=True)
    t.start(); prog.mainloop(); prog.destroy(); t.join()
    if error_holder[0]:
        r2 = tk.Tk(); r2.withdraw()
        from tkinter import messagebox as mb
        mb.showerror("Update Failed", error_holder[0]); r2.destroy(); return
    subprocess.Popen(["cmd.exe","/c",bat_path],
                     creationflags=subprocess.CREATE_NO_WINDOW, close_fds=True)
    sys.exit(0)

# ════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ════════════════════════════════════════════════════════════════════════════
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, date

# ════════════════════════════════════════════════════════════════════════════
#  COLOURS
# ════════════════════════════════════════════════════════════════════════════
FBC_DARK       = "#003B6F"
FBC_MID        = "#0066B3"
FBC_ACCENT     = "#00A3E0"
GREEN_DARK     = "#1A6B3A"
GREEN_LIGHT_BG = "#F0FFF4"
AMBER          = "#B45309"
AMBER_BG       = "#FFF8E7"
WHITE          = "#FFFFFF"
BG             = "#F0F4F8"
CARD_BG        = "#FFFFFF"
SEP_CLR        = "#D0DAE8"
SIDEBAR_BG     = "#001F3F"
SIDEBAR_TEXT   = "#B0C8E8"
OVERDUE_BG     = "#FFF0F0"
OVERDUE_CLR    = "#B71C1C"
AGING_WARN_BG  = "#FFF8E7"
AGING_WARN_CLR = "#B45309"

# ════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ════════════════════════════════════════════════════════════════════════════
APP_PASSWORD  = "anesu"
MAX_ATTEMPTS  = 6
OVERDUE_DAYS  = 2

# ── Default CC — shown on first run, user can change and save ────────────────
DEFAULT_CC = "Norman.Chirima@fbc.co.zw, Manatsa.Tagwirei@fbc.co.zw"

STATE_FILE    = os.path.join(os.path.expanduser("~"), ".fbc_client_forms.json")
SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".fbc_cf_settings.json")
SHEET_NAME    = "FBC_Clients"

COL_HEADERS = [
    "id","name","email","form_type","sent_date",
    "sent_by","reminders","done","done_date","reminder_dates","notes",
    "reply_date","reply_entry_id","reply_has_attach","sent_datetime"
]

SUBJ_SEND     = "SHARE TRADING ACCOUNT OPENING"
SUBJ_REMINDER = "Reminder — Outstanding Account Opening Forms"

_BODY_COMMON_INTRO = (
    "Good day\r\n\r\n"
    "Thank you for your interest in investment with Fbc Securities. "
    "Please find attached herein account opening forms for {{exchanges}} "
    "for your completion, scan and return via same e-mail address. "
    "Kindly also note that we have included account opening forms for Individual accounts.\r\n\r\n"
    "We have 2 stock exchanges namely the Zimbabwe Stock Exchange (ZSE) which transacts in "
    "ZiG currency and the Victoria Falls Exchange (VFEX) which transacts in the USD currency.\r\n\r\n"
    "Requirements For Individual Account Opening\r\n"
    "  •  Have attached account forms to open both ZSE and VFEX accounts.\r\n"
    "  •  The ZSE forms require your local banking account whilst the VFEX requires your FCA Nostro.\r\n"
    "  •  Kindly state your source of income for this investment.\r\n"
    "  •  Certified copy of your national identification document / valid passport page.\r\n"
    "  •  3 months bank statement.\r\n"
    "  •  Passport size photo.\r\n"
    "  •  Proof of residence\r\n\r\n"
    "For any further clarifications please do not hesitate to contact the undersigned.\r\n\r\n"
    "Regards,\r\n{{sender}}\r\nFBC Securities"
)
BODY_VFEX = _BODY_COMMON_INTRO.replace("{{exchanges}}", "the Victoria Falls Exchange (VFEX)")
BODY_ZSE  = _BODY_COMMON_INTRO.replace("{{exchanges}}", "the Zimbabwe Stock Exchange (ZSE)")
BODY_BOTH = _BODY_COMMON_INTRO.replace("{{exchanges}}",
    "both Zimbabwe Stock Exchange (ZSE) and the Victoria Falls Exchange (VFEX)")
BODY_REMINDER = (
    "Dear {{client}},\r\n\r\n"
    "This is a friendly reminder that we are still awaiting your completed "
    "account opening form(s).\r\n\r\n"
    "Please return the signed form(s) at your earliest convenience so we can "
    "proceed with opening your account without further delay.\r\n\r\n"
    "Regards,\r\n{{sender}}\r\nFBC Securities"
)

# ════════════════════════════════════════════════════════════════════════════
#  BUNDLED FORMS HELPER
# ════════════════════════════════════════════════════════════════════════════
def bundled_form(filename):
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "forms", filename)

_F_VFEX_INDIV = "chengetedzai individual 1.pdf"
_F_ZSE_INDIV  = "Individual Account Opening Form. 2024.doc"
_F_CUST_ID    = "FBC Securities Customer Identity Form 2021.doc"
_F_VFEX_CSD   = "VFEX CSD SECURITIES ACCOUNT OPENING FORM.PDF"
_F_ZSE_CSD    = "CSD 1 - SECURITIES ACCOUNT OPENING FORM - Copy.pdf"
_F_FAQ        = "Frequently asked Questions FBC Securities.docx"

_DEFAULT_VFEX = [bundled_form(_F_VFEX_INDIV), bundled_form(_F_CUST_ID),
                 bundled_form(_F_VFEX_CSD),   bundled_form(_F_FAQ)]
_DEFAULT_ZSE  = [bundled_form(_F_ZSE_INDIV),  bundled_form(_F_CUST_ID),
                 bundled_form(_F_ZSE_CSD),    bundled_form(_F_FAQ)]
_DEFAULT_BOTH = [bundled_form(_F_VFEX_INDIV), bundled_form(_F_ZSE_INDIV),
                 bundled_form(_F_CUST_ID),    bundled_form(_F_VFEX_CSD),
                 bundled_form(_F_ZSE_CSD),    bundled_form(_F_FAQ)]

# ════════════════════════════════════════════════════════════════════════════
#  HTML EMAIL
# ════════════════════════════════════════════════════════════════════════════
def _html_wrap(plain_body: str) -> str:
    safe = plain_body.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
    safe = safe.replace("\r\n","<br>").replace("\n","<br>")
    safe = safe.replace("  •  ","&nbsp;&nbsp;&bull;&nbsp;&nbsp;")
    return ('<html><body style="font-family: Georgia, \'Times New Roman\', serif; '
            'font-size: 11pt; color: #1a1a1a; line-height: 1.6;">'
            f"{safe}</body></html>")

# ════════════════════════════════════════════════════════════════════════════
#  SETTINGS  — now includes "cc_address"
# ════════════════════════════════════════════════════════════════════════════
def load_settings():
    try:
        with open(SETTINGS_FILE) as f:
            s = json.load(f)
        for key, default in [
            ("default_attachments_vfex", _DEFAULT_VFEX),
            ("default_attachments_zse",  _DEFAULT_ZSE),
            ("default_attachments_both", _DEFAULT_BOTH),
            ("cc_address",               DEFAULT_CC),     # ← new field
        ]:
            if not s.get(key):
                s[key] = default
        return s
    except Exception:
        return {
            "sender_name": "", "sheet_id": "", "key_file": "",
            "cc_address":  DEFAULT_CC,                    # ← new field
            "default_attachments_vfex": _DEFAULT_VFEX,
            "default_attachments_zse":  _DEFAULT_ZSE,
            "default_attachments_both": _DEFAULT_BOTH,
        }

def save_settings(s):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(s, f, indent=2)

# ════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS BACKEND
# ════════════════════════════════════════════════════════════════════════════
def _row_to_client(row):
    def g(i, default=""):
        return row[i] if i < len(row) else default
    def safe_int(val, default=0):
        try: return int(val or default)
        except: return default
    return {
        "id": g(0), "name": g(1), "email": g(2), "form_type": g(3),
        "sent_date": g(4), "sent_by": g(5),
        "reminders": safe_int(g(6,"0")),
        "done": g(7,"FALSE").upper() == "TRUE",
        "done_date": g(8), "reminder_dates": g(9), "notes": g(10),
        "reply_date": g(11), "reply_entry_id": g(12),
        "reply_has_attach": str(g(13,"FALSE")).upper() == "TRUE",
        "sent_datetime": g(14),
    }

def _client_to_row(c):
    return [
        c.get("id",""),       c.get("name",""),      c.get("email",""),
        c.get("form_type",""),c.get("sent_date",""),  c.get("sent_by",""),
        str(c.get("reminders",0)),
        "TRUE" if c.get("done") else "FALSE",
        c.get("done_date",""),    c.get("reminder_dates",""),
        c.get("notes",""),        c.get("reply_date",""),
        c.get("reply_entry_id",""),
        "TRUE" if c.get("reply_has_attach") else "FALSE",
        c.get("sent_datetime",""),
    ]

def _open_worksheet(key_file, sheet_id):
    import gspread
    from google.oauth2.service_account import Credentials
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = Credentials.from_service_account_file(key_file, scopes=SCOPES)
    gc     = gspread.authorize(creds)
    sh     = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=15)
        ws.append_row(COL_HEADERS)
    existing = ws.row_values(1)
    if existing != COL_HEADERS:
        if len(existing) == 14 and existing == COL_HEADERS[:14]:
            ws.update_cell(1, 15, "sent_datetime")
        else:
            ws.delete_rows(1); ws.insert_row(COL_HEADERS, 1)
    return ws

def test_sheets_connection(key_file, sheet_id):
    try: _open_worksheet(key_file, sheet_id); return True, ""
    except Exception as e: return False, str(e)

class SheetsDB:
    def __init__(self, key_file, sheet_id):
        self.key_file = key_file; self.sheet_id = sheet_id
        self._online = False; self._ws = None; self._connect()

    def _connect(self):
        try:
            self._ws = _open_worksheet(self.key_file, self.sheet_id)
            self._online = True
            print("[SheetsDB] Connected OK")
        except Exception as e:
            self._online = False; print(f"[SheetsDB] Offline — {e}")

    def _reconnect_if_needed(self):
        if not self._online: self._connect()

    @property
    def online(self): return self._online

    def read_all(self):
        self._reconnect_if_needed()
        if self._online:
            try:
                rows = self._ws.get_all_values()
                data = [r for r in rows if r and r[0] and r[0].strip().lower() != "id"]
                clients = [_row_to_client(r) for r in data]
                with open(STATE_FILE,"w") as f: json.dump(clients,f,indent=2)
                return clients
            except Exception as e:
                print(f"[SheetsDB] read failed: {e}"); self._online = False
        try:
            with open(STATE_FILE) as f: return json.load(f)
        except: return []

    def _find_row_num(self, client_id):
        try:
            col = self._ws.col_values(1)
            for i, val in enumerate(col):
                if val == client_id: return i+1
        except: pass
        return None

    def append_client(self, client):
        self._reconnect_if_needed()
        if self._online:
            try: self._ws.append_row(_client_to_row(client), value_input_option="RAW"); return True
            except Exception as e: print(f"[SheetsDB] append failed: {e}"); self._online = False
        return False

    def update_client(self, client):
        self._reconnect_if_needed()
        if self._online:
            try:
                rn = self._find_row_num(client["id"])
                if rn:
                    self._ws.update(values=[_client_to_row(client)],
                                    range_name=f"A{rn}:O{rn}",
                                    value_input_option="RAW"); return True
            except Exception as e: print(f"[SheetsDB] update failed: {e}"); self._online = False
        return False

    def delete_client(self, client_id):
        self._reconnect_if_needed()
        if self._online:
            try:
                rn = self._find_row_num(client_id)
                if rn: self._ws.delete_rows(rn); return True
            except Exception as e: print(f"[SheetsDB] delete failed: {e}"); self._online = False
        return False

# ════════════════════════════════════════════════════════════════════════════
#  LOCAL HELPERS
# ════════════════════════════════════════════════════════════════════════════
import uuid as _uuid

def new_id(): return str(_uuid.uuid4())[:8]

def load_clients_local():
    try:
        with open(STATE_FILE) as f: return json.load(f)
    except: return []

def save_clients_local(clients):
    with open(STATE_FILE,"w") as f: json.dump(clients,f,indent=2)

def days_since(date_str):
    try: return (date.today()-datetime.strptime(date_str,"%Y-%m-%d").date()).days
    except: return 0

def aging_label(client):
    sent = client.get("sent_date","")
    if not sent: return None
    d = days_since(sent)
    if d >= OVERDUE_DAYS:   return (f"⏰ OVERDUE — {d} days since sent", OVERDUE_CLR, OVERDUE_BG)
    elif d >= 1:            return (f"🕐 {d} day{'s' if d!=1 else ''} since sent", AGING_WARN_CLR, AGING_WARN_BG)
    else:                   return ("🕐 Sent today", FBC_MID, "#EAF4FB")

def status_of(c):
    if c.get("done"): return "✅ Completed", GREEN_DARK, "#EAF7EF"
    reply = c.get("reply_date","")
    if reply:
        d = days_since(reply); ha = c.get("reply_has_attach",True)
        if not ha: return (f"⚠ Reply received — NO ATTACHMENT ({d}d)", AMBER, AMBER_BG)
        if d >= 2: return (f"⚠ Back office overdue ({d}d since reply)", "#B71C1C", "#FDECEA")
        return (f"📨 Reply received — processing ({d}d)", "#6B21A8", "#F3E8FF")
    d = days_since(c.get("sent_date",""))
    if d >= OVERDUE_DAYS: return (f"⏰ Overdue — {d}d awaiting client", OVERDUE_CLR, OVERDUE_BG)
    if d >= 2: return (f"Awaiting client ({d}d — send reminder)", AMBER, AMBER_BG)
    return "Awaiting client reply", FBC_MID, "#EAF4FB"

def open_outlook(to, subject, body_plain, attachments=None, cc=""):
    try:
        import win32com.client as win32
        outlook = win32.Dispatch("outlook.application")
        mail    = outlook.CreateItem(0)
        mail.To = to
        if cc: mail.CC = cc
        mail.Subject  = subject
        mail.HTMLBody = _html_wrap(body_plain)
        for fp in (attachments or []):
            if fp and os.path.exists(fp): mail.Attachments.Add(fp)
        mail.Display(True); return True, ""
    except ImportError:
        return False, "pywin32 not installed.\nRun:  pip install pywin32"
    except Exception as e:
        return False, str(e)

def open_outlook_email_by_entry_id(entry_id):
    try:
        import win32com.client as win32
        outlook = win32.Dispatch("outlook.application")
        ns = outlook.GetNamespace("MAPI")
        ns.GetItemFromID(entry_id).Display(True); return True, ""
    except ImportError: return False, "pywin32 not installed."
    except Exception as e: return False, str(e)

def fill_template(template, client_name, sender_name):
    return (template.replace("{{client}}", client_name)
                    .replace("{{sender}}", sender_name or "FBC Securities"))

# ════════════════════════════════════════════════════════════════════════════
#  WIDGET HELPERS
# ════════════════════════════════════════════════════════════════════════════
def flat_entry(parent, var, **kw):
    return tk.Entry(parent, textvariable=var, font=("Segoe UI",10),
                    bg="#F7FAFC", fg="#1A2B3C", relief="flat",
                    highlightbackground=SEP_CLR, highlightthickness=1, **kw)

def flat_text(parent, height=7, **kw):
    return tk.Text(parent, font=("Segoe UI",10),
                   bg="#F7FAFC", fg="#1A2B3C", relief="flat",
                   highlightbackground=SEP_CLR, highlightthickness=1,
                   height=height, wrap="word", **kw)

def section_label(parent, text):
    tk.Label(parent, text=text, bg=BG, fg=FBC_DARK,
             font=("Segoe UI",9,"bold")).pack(anchor="w", pady=(12,3))

def card_frame(parent, **kw):
    return tk.Frame(parent, bg=CARD_BG, padx=14, pady=10,
                    highlightbackground=SEP_CLR, highlightthickness=1, **kw)

# ════════════════════════════════════════════════════════════════════════════
#  CC SETTINGS DIALOG  — standalone, launched from ⚙ CC in top bar
# ════════════════════════════════════════════════════════════════════════════
class CCSettingsDialog(tk.Toplevel):
    """Edit the default CC address saved in settings."""
    def __init__(self, parent, settings, on_save):
        super().__init__(parent)
        self.settings = settings; self.on_save = on_save
        self.title("⚙  Default CC Address")
        self.resizable(False, False); self.configure(bg=BG); self.grab_set()
        self._cc_var = tk.StringVar(value=settings.get("cc_address", DEFAULT_CC))
        self._build()
        self.update_idletasks()
        w, h = 560, 360
        px = parent.winfo_rootx() + (parent.winfo_width()  - w)//2
        py = parent.winfo_rooty() + (parent.winfo_height() - h)//2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

    def _build(self):
        tk.Frame(self, bg=FBC_MID, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text="⚙  Default CC Address", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",12,"bold")).pack(padx=16, anchor="w")
        tk.Label(hdr, text="Applied to every email unless overridden per-send",
                 bg=FBC_DARK, fg=SIDEBAR_TEXT,
                 font=("Segoe UI",9)).pack(padx=16, anchor="w", pady=(0,4))

        body = tk.Frame(self, bg=BG, padx=24, pady=20); body.pack(fill="both", expand=True)

        tk.Label(body, text="CC address(es)", bg=BG, fg=FBC_DARK,
                 font=("Segoe UI",10,"bold")).pack(anchor="w")
        tk.Label(body,
                 text="Separate multiple addresses with a comma  e.g.  alice@fbc.co.zw, bob@fbc.co.zw",
                 bg=BG, fg="#607080", font=("Segoe UI",8)).pack(anchor="w", pady=(0,6))
        flat_entry(body, self._cc_var).pack(fill="x", ipady=8)

        # info strip
        info = tk.Frame(body, bg="#E8F4FB", padx=12, pady=7,
                        highlightbackground="#90CAF9", highlightthickness=1)
        info.pack(fill="x", pady=(12,0))
        tk.Label(info,
                 text="ℹ  You can also override the CC per-send inside the Send Form dialog.",
                 bg="#E8F4FB", fg=FBC_DARK, font=("Segoe UI",9)).pack(anchor="w")

        btn_bar = tk.Frame(self, bg=BG, padx=24, pady=12); btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="Cancel", font=("Segoe UI",10),
                  bg=BG, fg="#607080", relief="flat", cursor="hand2",
                  command=self.destroy, activebackground=SEP_CLR
                  ).pack(side="right", padx=(6,0))
        tk.Button(btn_bar, text="  💾  Save  ", font=("Segoe UI",10,"bold"),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK, command=self._save).pack(side="right")

    def _save(self):
        self.settings["cc_address"] = self._cc_var.get().strip()
        save_settings(self.settings)
        self.on_save(self.settings["cc_address"])
        self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  GOOGLE SHEETS SETUP DIALOG
# ════════════════════════════════════════════════════════════════════════════
class SheetsSetupDialog(tk.Toplevel):
    def __init__(self, parent, settings, on_save):
        super().__init__(parent)
        self.settings = settings; self.on_save = on_save
        self.title("Google Sheets Setup")
        self.resizable(True, True); self.configure(bg=BG); self.grab_set()
        self._sid = tk.StringVar(value=settings.get("sheet_id",""))
        self._key = tk.StringVar(value=settings.get("key_file",""))
        self._build()
        self.update_idletasks()
        w = 580; screen_h = self.winfo_screenheight()
        h = min(520, int(screen_h*0.88))
        if parent:
            px = parent.winfo_rootx() + (parent.winfo_width()  - w)//2
            py = parent.winfo_rooty() + (parent.winfo_height() - h)//2
        else:
            px = (self.winfo_screenwidth() - w)//2
            py = (self.winfo_screenheight() - h)//2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(10,py)}")

    def _build(self):
        tk.Frame(self, bg=FBC_ACCENT, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text="☁  Connect to Google Sheets", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",12,"bold")).pack(padx=16, anchor="w")
        tk.Label(hdr, text="One-time setup — both users must complete this.",
                 bg=FBC_DARK, fg=SIDEBAR_TEXT,
                 font=("Segoe UI",9)).pack(padx=16, anchor="w", pady=(0,4))
        body = tk.Frame(self, bg=BG, padx=24, pady=12); body.pack(fill="both", expand=True)
        tk.Label(body, text="Google Sheet ID", bg=BG, fg=FBC_DARK,
                 font=("Segoe UI",10,"bold")).pack(anchor="w")
        tk.Label(body, text="From the Sheet URL — the long code between /d/ and /edit",
                 bg=BG, fg="#607080", font=("Segoe UI",8)).pack(anchor="w", pady=(0,4))
        flat_entry(body, self._sid).pack(fill="x", ipady=6, pady=(0,14))
        tk.Label(body, text="Service Account JSON Key File", bg=BG, fg=FBC_DARK,
                 font=("Segoe UI",10,"bold")).pack(anchor="w")
        tk.Label(body, text="The .json file downloaded from Google Cloud Console",
                 bg=BG, fg="#607080", font=("Segoe UI",8)).pack(anchor="w", pady=(0,4))
        kr = tk.Frame(body, bg=BG); kr.pack(fill="x", pady=(0,10))
        flat_entry(kr, self._key).pack(side="left", fill="x", expand=True, ipady=6, padx=(0,6))
        tk.Button(kr, text="📂 Browse", font=("Segoe UI",9),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK, command=self._browse).pack(side="left")
        self._lbl_err = tk.Label(body, text="", bg=BG, fg="#B71C1C", font=("Segoe UI",9))
        self._lbl_err.pack(anchor="w", pady=(0,6))
        bb = tk.Frame(body, bg=BG); bb.pack(fill="x", pady=(4,0))
        tk.Button(bb, text="Skip (offline mode)", font=("Segoe UI",9),
                  bg=BG, fg="#607080", relief="flat", cursor="hand2",
                  activebackground=SEP_CLR, command=self._skip).pack(side="right", padx=(6,0))
        tk.Button(bb, text="  ✅  Connect & Save  ", font=("Segoe UI",10,"bold"),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK, command=self._save).pack(side="right")

    def _browse(self):
        p = filedialog.askopenfilename(title="Select Google Service Account JSON key",
            filetypes=[("JSON files","*.json"),("All files","*.*")])
        if p: self._key.set(p)

    def _save(self):
        sid = self._sid.get().strip(); key = self._key.get().strip()
        if not sid: self._lbl_err.config(text="❌  Please paste your Google Sheet ID."); return
        if not key: self._lbl_err.config(text="❌  Please select the JSON key file."); return
        if not os.path.exists(key):
            self._lbl_err.config(text=f"❌  File not found: {key}")
            messagebox.showerror("File Not Found",
                f"Cannot find the JSON file at:\n{key}", parent=self); return
        self._lbl_err.config(text="⏳  Testing connection…"); self.update()
        ok, err = test_sheets_connection(key, sid)
        if not ok:
            self._lbl_err.config(text=f"❌  {err[:120]}")
            messagebox.showerror("Connection Error", err, parent=self); return
        self.settings["sheet_id"] = sid; self.settings["key_file"] = key
        save_settings(self.settings); self.on_save(sid, key); self.destroy()

    def _skip(self): self.on_save("",""); self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  LOGIN DIALOG
# ════════════════════════════════════════════════════════════════════════════
class LoginDialog(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FBC Client Forms — Login")
        self.resizable(False, False); self.configure(bg=SIDEBAR_BG)
        self.authenticated = False; self.sender_name = ""
        self._attempts = 0; self._settings = load_settings()
        self._build()
        self.update_idletasks()
        w = 400; sh = self.winfo_screenheight()
        h = min(self.winfo_reqheight()+20, int(sh*0.90))
        x = (self.winfo_screenwidth()-w)//2; y = max(20,(sh-h)//2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build(self):
        hdr = tk.Frame(self, bg=FBC_ACCENT, pady=18); hdr.pack(fill="x")
        tk.Label(hdr, text="FBC", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",20,"bold"), padx=12, pady=6).pack()
        tk.Label(hdr, text="Client Forms Manager", bg=FBC_ACCENT, fg=WHITE,
                 font=("Segoe UI",11)).pack(pady=(2,0))
        body = tk.Frame(self, bg=SIDEBAR_BG, padx=36, pady=20); body.pack(fill="both", expand=True)
        saved_name = self._settings.get("sender_name","").strip()
        if saved_name:
            self._name_var = tk.StringVar(value=saved_name)
            greet = tk.Frame(body, bg="#0D2B4E", padx=12, pady=10,
                             highlightbackground=FBC_MID, highlightthickness=1)
            greet.pack(fill="x", pady=(0,16))
            tk.Label(greet, text=f"👤  Welcome back, {saved_name}",
                     bg="#0D2B4E", fg=WHITE, font=("Segoe UI",11,"bold")).pack(anchor="w")
            tk.Label(greet, text="Not you? Click 'Change' in the app after login.",
                     bg="#0D2B4E", fg="#607080", font=("Segoe UI",8)).pack(anchor="w", pady=(2,0))
            name_entry = None
        else:
            tk.Label(body, text="Your Name", bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                     font=("Segoe UI",10,"bold")).pack(anchor="w")
            tk.Label(body, text="Appears in the email sign-off (Regards, ...)",
                     bg=SIDEBAR_BG, fg="#607080", font=("Segoe UI",8)).pack(anchor="w", pady=(0,4))
            self._name_var = tk.StringVar()
            name_entry = tk.Entry(body, textvariable=self._name_var, font=("Segoe UI",11),
                                  bg="#0D2B4E", fg=WHITE, insertbackground=WHITE, relief="flat",
                                  highlightbackground=FBC_MID, highlightthickness=1)
            name_entry.pack(fill="x", ipady=7, pady=(0,16)); name_entry.focus()
        tk.Label(body, text="Password", bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                 font=("Segoe UI",10,"bold")).pack(anchor="w")
        pw_row = tk.Frame(body, bg=SIDEBAR_BG); pw_row.pack(fill="x", pady=(4,0))
        self._pw = tk.StringVar(); self._show = False
        self._entry = tk.Entry(pw_row, textvariable=self._pw, show="●", font=("Segoe UI",11),
                               bg="#0D2B4E", fg=WHITE, insertbackground=WHITE, relief="flat",
                               highlightbackground=FBC_MID, highlightthickness=1)
        self._entry.pack(side="left", fill="x", expand=True, ipady=7, padx=(0,4))
        self._eye = tk.Button(pw_row, text="👁", command=self._toggle, bg="#0D2B4E",
                              fg=SIDEBAR_TEXT, relief="flat", font=("Segoe UI",12),
                              cursor="hand2", activebackground=FBC_MID,
                              activeforeground=WHITE, padx=6)
        self._eye.pack(side="left")
        self._lbl_err = tk.Label(body, text="", bg=SIDEBAR_BG, fg="#FF6B6B", font=("Segoe UI",9))
        self._lbl_err.pack(anchor="w", pady=(6,0))
        self._lbl_att = tk.Label(body, text="", bg=SIDEBAR_BG, fg="#607080", font=("Segoe UI",8))
        self._lbl_att.pack(anchor="w")
        tk.Button(body, text="  🔓  Login  ", command=self._attempt, bg=FBC_MID, fg=WHITE,
                  relief="flat", font=("Segoe UI",11,"bold"), cursor="hand2",
                  pady=10, activebackground=FBC_ACCENT).pack(fill="x", pady=(14,0))
        if name_entry:
            name_entry.bind("<Return>", lambda _: self._attempt())
        else:
            self._entry.focus()
        self._entry.bind("<Return>", lambda _: self._attempt())
        sid = self._settings.get("sheet_id","")
        tk.Label(body,
                 text="☁  Cloud sync: configured" if sid else "⚠  Cloud sync: not set up yet",
                 bg=SIDEBAR_BG, fg="#6EE7B7" if sid else "#FFB347",
                 font=("Segoe UI",8)).pack(anchor="w", pady=(16,0))
        tk.Label(body, text=f"v{VERSION}", bg=SIDEBAR_BG, fg="#2A4A6A",
                 font=("Segoe UI",8)).pack(anchor="w", pady=(4,0))

    def _toggle(self):
        self._show = not self._show
        self._entry.config(show="" if self._show else "●")
        self._eye.config(text="🙈" if self._show else "👁")

    def _attempt(self):
        sender = self._name_var.get().strip()
        if not sender: self._lbl_err.config(text="❌  Please enter your name."); return
        if self._pw.get().strip().lower() == APP_PASSWORD.lower():
            self.sender_name = sender; self.authenticated = True
            save_settings({**self._settings, "sender_name": sender})
            self.destroy(); return
        self._attempts += 1; rem = MAX_ATTEMPTS - self._attempts
        if rem <= 0:
            messagebox.showerror("Access Denied",
                "Too many incorrect attempts. The app will now close.")
            self.destroy(); return
        self._lbl_err.config(text="❌  Incorrect password.")
        self._lbl_att.config(text=f"  {rem} attempt{'s' if rem>1 else ''} remaining")
        self._pw.set(""); self._entry.focus(); self._shake()

    def _shake(self, n=6, d=8):
        x0, y0 = self.winfo_x(), self.winfo_y()
        def step(i):
            if i == 0: self.geometry(f"+{x0}+{y0}"); return
            self.geometry(f"+{x0+(d if i%2==0 else -d)}+{y0}")
            self.after(40, lambda: step(i-1))
        step(n)

    def _close(self): self.authenticated = False; self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  ATTACHMENT DEFAULTS DIALOG
# ════════════════════════════════════════════════════════════════════════════
class AttachmentDefaultsDialog(tk.Toplevel):
    def __init__(self, parent, settings, on_save):
        super().__init__(parent)
        self.settings = settings; self.on_save = on_save
        self.title("⚙  Attachment Defaults")
        self.resizable(True, True); self.configure(bg=BG); self.grab_set()
        self._lists = {
            "vfex": list(settings.get("default_attachments_vfex",[])),
            "zse":  list(settings.get("default_attachments_zse", [])),
            "both": list(settings.get("default_attachments_both",[])),
        }
        self._frames = {}; self._build()
        self.update_idletasks()
        w,h = 680,600
        px = parent.winfo_rootx()+(parent.winfo_width() -w)//2
        py = parent.winfo_rooty()+(parent.winfo_height()-h)//2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

    def _build(self):
        tk.Frame(self, bg=FBC_ACCENT, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text="⚙  Default Attachment Bundles", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",12,"bold")).pack(padx=16, anchor="w")
        body = tk.Frame(self, bg=BG, padx=20, pady=10); body.pack(fill="both", expand=True)
        for key, label, colour in [
            ("vfex","VFEX — default attachments",FBC_MID),
            ("zse", "ZSE  — default attachments","#1A6B3A"),
            ("both","BOTH (VFEX + ZSE) — default attachments","#6B21A8"),
        ]:
            tk.Label(body, text=label, bg=BG, fg=colour,
                     font=("Segoe UI",9,"bold")).pack(anchor="w", pady=(10,3))
            frame = tk.Frame(body, bg=CARD_BG,
                             highlightbackground=SEP_CLR, highlightthickness=1, padx=10, pady=8)
            frame.pack(fill="x"); self._frames[key] = frame; self._render_list(key)
        bb = tk.Frame(self, bg=BG, padx=20, pady=10); bb.pack(fill="x")
        tk.Button(bb, text="Cancel", font=("Segoe UI",10), bg=BG, fg="#607080",
                  relief="flat", cursor="hand2", command=self.destroy,
                  activebackground=SEP_CLR).pack(side="right", padx=(6,0))
        tk.Button(bb, text="  💾  Save Defaults  ", font=("Segoe UI",10,"bold"),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK, command=self._save).pack(side="right")

    def _render_list(self, key):
        frame = self._frames[key]
        for w in frame.winfo_children(): w.destroy()
        files = self._lists[key]
        if not files:
            tk.Label(frame, text="No default files set — click ➕ to add.",
                     bg=CARD_BG, fg="#8096B0", font=("Segoe UI",9,"italic")).pack(anchor="w")
        else:
            for i, fp in enumerate(files):
                row = tk.Frame(frame, bg=CARD_BG); row.pack(fill="x", pady=1)
                exists = os.path.exists(fp)
                tk.Label(row, text=f"{'📄' if exists else '❌'}  {os.path.basename(fp)}",
                         bg=CARD_BG, fg=FBC_DARK if exists else "#B71C1C",
                         font=("Segoe UI",9), anchor="w").pack(side="left", fill="x", expand=True)
                tk.Button(row, text="✕", font=("Segoe UI",8), bg=CARD_BG, fg="#CBD5E1",
                          relief="flat", cursor="hand2", activebackground=BG,
                          command=lambda k=key,idx=i: self._remove(k,idx)).pack(side="right")
        ar = tk.Frame(frame, bg=CARD_BG); ar.pack(anchor="w", pady=(4,0))
        tk.Button(ar, text="➕ Add File", font=("Segoe UI",9), bg=BG, fg=FBC_MID,
                  relief="flat", cursor="hand2", activebackground=SEP_CLR,
                  command=lambda k=key: self._add(k)).pack(side="left")

    def _add(self, key):
        paths = filedialog.askopenfilenames(
            title=f"Select files for {key.upper()} bundle",
            filetypes=[("Documents","*.pdf *.doc *.docx *.jpg *.png"),("All files","*.*")])
        if paths: self._lists[key].extend(paths); self._render_list(key)

    def _remove(self, key, idx):
        self._lists[key].pop(idx); self._render_list(key)

    def _save(self):
        self.settings["default_attachments_vfex"] = self._lists["vfex"]
        self.settings["default_attachments_zse"]  = self._lists["zse"]
        self.settings["default_attachments_both"] = self._lists["both"]
        save_settings(self.settings); self.on_save(); self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  SEND FORM DIALOG  — CC field with override
# ════════════════════════════════════════════════════════════════════════════
class SendFormDialog(tk.Toplevel):
    def __init__(self, parent, sender_name, settings, on_sent):
        super().__init__(parent)
        self.sender_name = sender_name; self.settings = settings; self.on_sent = on_sent
        self.title("Send Account Opening Form")
        self.resizable(False, False); self.configure(bg=BG); self.grab_set()
        self._canvas = None; self._extra_files = []
        # CC starts from saved default, user can edit per-send
        self._cc_var = tk.StringVar(value=settings.get("cc_address", DEFAULT_CC))
        self._build()
        self.update_idletasks()
        w, h = 640, 800
        px = parent.winfo_rootx()+(parent.winfo_width() -w)//2
        py = parent.winfo_rooty()+(parent.winfo_height()-h)//2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _close(self): self._unbind_scroll(); self.destroy()

    def _unbind_scroll(self):
        try:
            if self._canvas and self._canvas.winfo_exists():
                self._canvas.unbind_all("<MouseWheel>")
        except: pass

    def _get_default_attachments(self):
        t = self._ftype.get() if hasattr(self,"_ftype") else "BOTH"
        return list(self.settings.get(f"default_attachments_{t.lower()}",[]))

    def _all_attachments(self):
        return self._get_default_attachments() + self._extra_files

    def _build(self):
        tk.Frame(self, bg=FBC_ACCENT, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text="✉  Send Account Opening Form", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",12,"bold")).pack(padx=16, anchor="w")

        outer = tk.Frame(self, bg=BG); outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0); self._canvas = canvas
        sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        body = tk.Frame(canvas, bg=BG, padx=20, pady=4)
        cid = canvas.create_window((0,0), window=body, anchor="nw")
        body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cid, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120),"units"))

        # Sender strip
        strip = tk.Frame(body, bg="#E8F1FB", padx=12, pady=7,
                         highlightbackground="#A8C4E0", highlightthickness=1)
        strip.pack(fill="x", pady=(8,0))
        tk.Label(strip, text=f"✏  Sending as:  {self.sender_name}",
                 bg="#E8F1FB", fg=FBC_DARK, font=("Segoe UI",9,"bold")).pack(anchor="w")

        # ── Client details ────────────────────────────────────────────────
        section_label(body, "CLIENT DETAILS")
        c1 = card_frame(body); c1.pack(fill="x")
        self._name_var = tk.StringVar(); self._email_var = tk.StringVar()
        for lbl, var in [("Client name",self._name_var),("Email address",self._email_var)]:
            row = tk.Frame(c1, bg=CARD_BG); row.pack(fill="x", pady=3)
            tk.Label(row, text=lbl, bg=CARD_BG, fg="#607080",
                     font=("Segoe UI",9), width=14, anchor="w").pack(side="left")
            flat_entry(row, var).pack(side="left", fill="x", expand=True, ipady=5, padx=(4,0))

        # ── CC address ────────────────────────────────────────────────────
        section_label(body, "CC ADDRESS")
        cc_card = card_frame(body); cc_card.pack(fill="x")

        cc_top = tk.Frame(cc_card, bg=CARD_BG); cc_top.pack(fill="x")
        tk.Label(cc_top, text="CC:", bg=CARD_BG, fg="#607080",
                 font=("Segoe UI",9), width=6, anchor="w").pack(side="left")
        flat_entry(cc_top, self._cc_var).pack(side="left", fill="x", expand=True,
                                               ipady=5, padx=(4,0))

        cc_hint = tk.Frame(cc_card, bg=CARD_BG); cc_hint.pack(fill="x", pady=(4,0))
        tk.Label(cc_hint,
                 text="Separate multiple addresses with commas.  Clear the field to send with no CC.",
                 bg=CARD_BG, fg="#8096B0", font=("Segoe UI",8)).pack(side="left")
        tk.Button(cc_hint, text="↩ Reset to default", font=("Segoe UI",8),
                  bg=CARD_BG, fg=FBC_MID, relief="flat", cursor="hand2",
                  activebackground=BG,
                  command=lambda: self._cc_var.set(
                      self.settings.get("cc_address", DEFAULT_CC))
                  ).pack(side="right")

        # ── Form type ─────────────────────────────────────────────────────
        section_label(body, "FORM TYPE")
        c2 = card_frame(body); c2.pack(fill="x")
        self._ftype = tk.StringVar(value="BOTH")
        for val, txt in [("VFEX","VFEX Account Opening"),
                         ("ZSE", "ZSE Account Opening"),
                         ("BOTH","VFEX + ZSE (Both)  ← standard")]:
            tk.Radiobutton(c2, text=txt, variable=self._ftype, value=val,
                           command=self._on_type_change, bg=CARD_BG, fg=FBC_DARK,
                           selectcolor=CARD_BG, font=("Segoe UI",10), cursor="hand2",
                           activebackground=CARD_BG).pack(anchor="w", pady=2)

        # ── Attachments ───────────────────────────────────────────────────
        section_label(body, "ATTACHMENTS")
        self._att_card = card_frame(body); self._att_card.pack(fill="x")
        self._render_attachments()

        # ── Email message ─────────────────────────────────────────────────
        section_label(body, "EMAIL MESSAGE  (sent in Georgia font)")
        c4 = card_frame(body); c4.pack(fill="x")
        tk.Label(c4, text="Subject:", bg=CARD_BG, fg="#607080",
                 font=("Segoe UI",9)).pack(anchor="w")
        self._subj = tk.StringVar(value=SUBJ_SEND)
        flat_entry(c4, self._subj).pack(fill="x", ipady=5, pady=(2,10))
        tk.Label(c4, text="Body:", bg=CARD_BG, fg="#607080",
                 font=("Segoe UI",9)).pack(anchor="w")
        self._body_txt = flat_text(c4, height=11)
        self._body_txt.pack(fill="x", pady=(2,0))
        tk.Frame(body, bg=BG, height=8).pack()

        # ── Buttons ───────────────────────────────────────────────────────
        btn_bar = tk.Frame(self, bg=BG, padx=20, pady=10); btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="Cancel", font=("Segoe UI",10), bg=BG, fg="#607080",
                  relief="flat", cursor="hand2", command=self._close,
                  activebackground=SEP_CLR).pack(side="right", padx=(6,0))
        tk.Button(btn_bar, text="  ✉  Open in Outlook  ", font=("Segoe UI",10,"bold"),
                  bg=FBC_MID, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_DARK, command=self._send).pack(side="right")

        self._on_type_change()

    def _render_attachments(self):
        card = self._att_card
        for w in card.winfo_children(): w.destroy()
        defaults  = self._get_default_attachments()
        all_files = defaults + self._extra_files
        if not all_files:
            warn = tk.Frame(card, bg="#FFF8E7", padx=10, pady=8,
                            highlightbackground="#D4A017", highlightthickness=1)
            warn.pack(fill="x", pady=(0,6))
            tk.Label(warn, text="⚠  No default attachments configured.",
                     bg="#FFF8E7", fg=AMBER, font=("Segoe UI",9)).pack(anchor="w")
        else:
            if defaults:
                tk.Label(card, text="Default files (bundled forms):", bg=CARD_BG, fg="#607080",
                         font=("Segoe UI",8,"italic")).pack(anchor="w", pady=(0,2))
                for fp in defaults:
                    row = tk.Frame(card, bg=CARD_BG); row.pack(fill="x", pady=1)
                    exists = os.path.exists(fp)
                    tk.Label(row, text=f"{'📄' if exists else '❌'}  {os.path.basename(fp)}",
                             bg=CARD_BG, fg="#374151" if exists else "#B71C1C",
                             font=("Segoe UI",9)).pack(side="left", fill="x", expand=True)
                    tk.Label(row, text="default", bg=CARD_BG, fg="#A0B0C0",
                             font=("Segoe UI",7)).pack(side="right")
            if self._extra_files:
                tk.Label(card, text="Extra files (this send only):", bg=CARD_BG, fg="#607080",
                         font=("Segoe UI",8,"italic")).pack(anchor="w", pady=(6,2))
                for i, fp in enumerate(self._extra_files):
                    row = tk.Frame(card, bg=CARD_BG); row.pack(fill="x", pady=1)
                    exists = os.path.exists(fp)
                    tk.Label(row, text=f"{'📎' if exists else '❌'}  {os.path.basename(fp)}",
                             bg=CARD_BG, fg=FBC_MID if exists else "#B71C1C",
                             font=("Segoe UI",9)).pack(side="left", fill="x", expand=True)
                    tk.Button(row, text="✕ Remove", font=("Segoe UI",8), bg=CARD_BG, fg="#CBD5E1",
                              relief="flat", cursor="hand2", activebackground=BG,
                              command=lambda idx=i: self._remove_extra(idx)).pack(side="right")
        br = tk.Frame(card, bg=CARD_BG); br.pack(anchor="w", pady=(8,0))
        tk.Button(br, text="➕ Add File", font=("Segoe UI",9), bg=BG, fg=FBC_MID,
                  relief="flat", cursor="hand2", activebackground=SEP_CLR,
                  command=self._add_extra).pack(side="left", padx=(0,8))
        tk.Button(br, text="⚙ Edit Defaults", font=("Segoe UI",9), bg=BG, fg="#607080",
                  relief="flat", cursor="hand2", activebackground=SEP_CLR,
                  command=self._open_defaults).pack(side="left")
        missing = [f for f in all_files if not os.path.exists(f)]
        summary = f"{len(all_files)} file(s) attached"
        if missing: summary += f"  —  ⚠ {len(missing)} file(s) missing!"
        tk.Label(card, text=summary, bg=CARD_BG,
                 fg="#B71C1C" if missing else GREEN_DARK,
                 font=("Segoe UI",8,"bold")).pack(anchor="w", pady=(6,0))

    def _add_extra(self):
        paths = filedialog.askopenfilenames(
            title="Add extra file(s) for this send",
            filetypes=[("Documents","*.pdf *.doc *.docx *.jpg *.png"),("All files","*.*")])
        if paths: self._extra_files.extend(paths); self._render_attachments()

    def _remove_extra(self, idx):
        self._extra_files.pop(idx); self._render_attachments()

    def _open_defaults(self):
        def _on_saved():
            fresh = load_settings(); self.settings.update(fresh); self._render_attachments()
        AttachmentDefaultsDialog(self, self.settings, _on_saved)

    def _on_type_change(self):
        t = self._ftype.get()
        tmpl = BODY_VFEX if t=="VFEX" else (BODY_ZSE if t=="ZSE" else BODY_BOTH)
        preview = tmpl.replace("{{sender}}", self.sender_name or "Your Name")
        self._body_txt.delete("1.0","end"); self._body_txt.insert("1.0", preview)
        self._render_attachments()

    def _send(self):
        client_name = self._name_var.get().strip()
        email       = self._email_var.get().strip()
        if not client_name:
            messagebox.showwarning("Missing","Please enter the client name.",parent=self); return
        if not email or "@" not in email:
            messagebox.showwarning("Missing","Please enter a valid email address.",parent=self); return
        attachments = self._all_attachments()
        missing = [f for f in attachments if not os.path.exists(f)]
        if missing:
            names = "\n".join(f"  • {os.path.basename(f)}" for f in missing)
            if not messagebox.askyesno("Missing files",
                    f"The following attachment(s) could not be found:\n\n{names}\n\n"
                    "They will be skipped. Continue sending without them?",parent=self):
                return
        attachments = [f for f in attachments if os.path.exists(f)]
        cc       = self._cc_var.get().strip()
        sub      = self._subj.get().strip()
        raw      = self._body_txt.get("1.0","end").strip()
        body_txt = fill_template(raw, client_name, self.sender_name)
        ok, err  = open_outlook(email, sub, body_txt, attachments, cc=cc)
        if not ok:
            messagebox.showerror("Outlook error", err, parent=self); return
        self._unbind_scroll()
        now = datetime.now()
        self.on_sent({
            "id": new_id(), "name": client_name, "email": email,
            "form_type": self._ftype.get(),
            "sent_date": now.date().isoformat(),
            "sent_datetime": now.strftime("%Y-%m-%dT%H:%M:%S"),
            "sent_by": self.sender_name,
            "reminders": 0, "done": False,
            "done_date":"","reminder_dates":"","notes":"",
            "reply_date":"","reply_entry_id":"","reply_has_attach": False,
        })
        self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  REMINDER DIALOG  — CC field with override
# ════════════════════════════════════════════════════════════════════════════
class ReminderDialog(tk.Toplevel):
    def __init__(self, parent, client, sender_name, settings, on_sent):
        super().__init__(parent)
        self.client = client; self.sender_name = sender_name
        self.settings = settings; self.on_sent = on_sent
        self.title("Send Reminder")
        self.resizable(False, False); self.configure(bg=BG); self.grab_set()
        self._cc_var = tk.StringVar(value=settings.get("cc_address", DEFAULT_CC))
        self._build()
        self.update_idletasks()
        w, h = 540, 520
        px = parent.winfo_rootx()+(parent.winfo_width() -w)//2
        py = parent.winfo_rooty()+(parent.winfo_height()-h)//2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

    def _build(self):
        tk.Frame(self, bg=AMBER, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text=f"🔔  Reminder — {self.client['name']}",
                 bg=FBC_DARK, fg=WHITE, font=("Segoe UI",11,"bold")).pack(padx=16, anchor="w")

        body = tk.Frame(self, bg=BG, padx=20, pady=12); body.pack(fill="both", expand=True)

        days = days_since(self.client.get("sent_date",""))
        info = tk.Frame(body, bg=AMBER_BG, padx=12, pady=8,
                        highlightbackground="#D4A017", highlightthickness=1)
        info.pack(fill="x", pady=(0,10))
        tk.Label(info, text=f"⚠  Form sent {days} day(s) ago — no reply recorded.",
                 bg=AMBER_BG, fg=AMBER, font=("Segoe UI",9,"bold")).pack(anchor="w")
        tk.Label(info, text=f"To: {self.client['email']}",
                 bg=AMBER_BG, fg="#607080", font=("Segoe UI",9)).pack(anchor="w")

        # CC row
        cc_card = card_frame(body); cc_card.pack(fill="x", pady=(0,6))
        cc_row = tk.Frame(cc_card, bg=CARD_BG); cc_row.pack(fill="x")
        tk.Label(cc_row, text="CC:", bg=CARD_BG, fg="#607080",
                 font=("Segoe UI",9), width=6, anchor="w").pack(side="left")
        flat_entry(cc_row, self._cc_var).pack(side="left", fill="x", expand=True,
                                               ipady=5, padx=(4,0))
        tk.Button(cc_row, text="↩", font=("Segoe UI",9), bg=CARD_BG, fg=FBC_MID,
                  relief="flat", cursor="hand2", activebackground=BG,
                  command=lambda: self._cc_var.set(
                      self.settings.get("cc_address", DEFAULT_CC))
                  ).pack(side="right", padx=(4,0))
        tk.Label(cc_card, text="Clear to send with no CC.  ↩ resets to default.",
                 bg=CARD_BG, fg="#8096B0", font=("Segoe UI",8)).pack(anchor="w", pady=(3,0))

        tk.Label(body, text="Subject:", bg=BG, fg="#607080",
                 font=("Segoe UI",9)).pack(anchor="w")
        self._subj = tk.StringVar(value=SUBJ_REMINDER)
        flat_entry(body, self._subj).pack(fill="x", ipady=5, pady=(2,10))
        tk.Label(body, text="Message:", bg=BG, fg="#607080",
                 font=("Segoe UI",9)).pack(anchor="w")
        self._msg = flat_text(body, height=9); self._msg.pack(fill="x", pady=(2,0))
        self._msg.insert("1.0", fill_template(BODY_REMINDER, self.client["name"], self.sender_name))

        btn_bar = tk.Frame(self, bg=BG, padx=20, pady=10); btn_bar.pack(fill="x")
        tk.Button(btn_bar, text="Cancel", font=("Segoe UI",10), bg=BG, fg="#607080",
                  relief="flat", cursor="hand2", command=self.destroy,
                  activebackground=SEP_CLR).pack(side="right", padx=(6,0))
        tk.Button(btn_bar, text="  🔔  Open in Outlook  ", font=("Segoe UI",10,"bold"),
                  bg=AMBER, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground="#92400E", command=self._send).pack(side="right")

    def _send(self):
        raw  = self._msg.get("1.0","end").strip()
        body = fill_template(raw, self.client["name"], self.sender_name)
        cc   = self._cc_var.get().strip()
        ok, err = open_outlook(self.client["email"], self._subj.get().strip(), body, cc=cc)
        if not ok: messagebox.showerror("Outlook error", err, parent=self); return
        self.on_sent(); self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  CONFIRM DIALOG
# ════════════════════════════════════════════════════════════════════════════
class ConfirmDialog(tk.Toplevel):
    def __init__(self, parent, client, on_confirm):
        super().__init__(parent)
        self.client = client; self.on_confirm = on_confirm
        self.title("Confirm — Forms Received")
        self.resizable(False,False); self.configure(bg=BG); self.grab_set()
        self._build()
        self.update_idletasks()
        w,h=420,260
        px=parent.winfo_rootx()+(parent.winfo_width() -w)//2
        py=parent.winfo_rooty()+(parent.winfo_height()-h)//2
        self.geometry(f"{w}x{h}+{max(px,0)}+{max(py,0)}")

    def _build(self):
        tk.Frame(self, bg=GREEN_DARK, height=4).pack(fill="x")
        hdr = tk.Frame(self, bg=FBC_DARK, pady=10); hdr.pack(fill="x")
        tk.Label(hdr, text="✅  Confirm Forms Received", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",11,"bold")).pack(padx=16, anchor="w")
        body = tk.Frame(self, bg=BG, padx=24, pady=20); body.pack(fill="both", expand=True)
        tk.Label(body, text="Mark this client as completed?\n"
                            "They will be removed from the active dashboard.",
                 bg=BG, fg="#374151", font=("Segoe UI",10), justify="left").pack(anchor="w")
        info = tk.Frame(body, bg=GREEN_LIGHT_BG, padx=12, pady=10,
                        highlightbackground="#6EE7B7", highlightthickness=1)
        info.pack(fill="x", pady=12)
        tk.Label(info, text=self.client["name"],
                 bg=GREEN_LIGHT_BG, fg=GREEN_DARK, font=("Segoe UI",11,"bold")).pack(anchor="w")
        tk.Label(info, text=f"{self.client['email']}  ·  {self.client['form_type']}",
                 bg=GREEN_LIGHT_BG, fg="#607080", font=("Segoe UI",9)).pack(anchor="w")
        bb = tk.Frame(self, bg=BG, padx=20, pady=10); bb.pack(fill="x")
        tk.Button(bb, text="Cancel", font=("Segoe UI",10), bg=BG, fg="#607080",
                  relief="flat", cursor="hand2", command=self.destroy,
                  activebackground=SEP_CLR).pack(side="right", padx=(6,0))
        tk.Button(bb, text="  ✅  Confirm & Remove  ", font=("Segoe UI",10,"bold"),
                  bg=GREEN_DARK, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground="#155724",
                  command=lambda: (self.on_confirm(), self.destroy())).pack(side="right")

# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self, sender_name, settings):
        super().__init__()
        self.sender_name = sender_name; self.settings = settings
        self.title(f"FBC Client Forms Manager  v{VERSION}")
        self.state("zoomed"); self.configure(bg=BG)
        sid = settings.get("sheet_id",""); key = settings.get("key_file","")
        self.db = SheetsDB(key,sid) if (sid and key and os.path.exists(key)) else None
        self.clients=[]; self._filter="all"; self._build()
        self._full_sync(); self._schedule_auto_sync()

    def _build(self):
        top = tk.Frame(self, bg=FBC_DARK); top.pack(fill="x")
        tk.Frame(top, bg=FBC_ACCENT, width=6).pack(side="left", fill="y")
        tk.Label(top, text="📋  FBC Client Forms Manager", bg=FBC_DARK, fg=WHITE,
                 font=("Segoe UI",14,"bold"), pady=14, padx=16).pack(side="left")
        self._sync_lbl = tk.Label(top, text="", bg=FBC_DARK, font=("Segoe UI",9))
        self._sync_lbl.pack(side="left", padx=10)

        # right side buttons
        tk.Button(top, text="  ✉  Send New Form  ", font=("Segoe UI",10,"bold"),
                  bg=FBC_ACCENT, fg=WHITE, relief="flat", cursor="hand2",
                  activebackground=FBC_MID, pady=8, padx=4,
                  command=self._open_send).pack(side="right", padx=12, pady=6)
        tk.Button(top, text="☁  Sheets", font=("Segoe UI",9), bg=FBC_DARK,
                  fg=SIDEBAR_TEXT, relief="flat", cursor="hand2", activebackground=FBC_MID,
                  command=self._open_sheets_setup).pack(side="right", padx=4, pady=6)
        tk.Button(top, text="✉  CC", font=("Segoe UI",9), bg=FBC_DARK,
                  fg=SIDEBAR_TEXT, relief="flat", cursor="hand2", activebackground=FBC_MID,
                  command=self._open_cc_settings).pack(side="right", padx=4, pady=6)
        tk.Button(top, text="Change", font=("Segoe UI",8), bg=FBC_DARK, fg="#4A7AAA",
                  relief="flat", cursor="hand2", activebackground=FBC_DARK,
                  command=self._clear_name).pack(side="right", padx=(0,4))
        tk.Label(top, text=f"👤  {self.sender_name}", bg=FBC_DARK, fg=SIDEBAR_TEXT,
                 font=("Segoe UI",9)).pack(side="right", padx=(0,4))
        tk.Label(top, text=f"v{VERSION}", bg=FBC_DARK, fg="#2A5A8A",
                 font=("Segoe UI",9)).pack(side="right", padx=10)

        # metrics
        metric_bg = tk.Frame(self, bg=SIDEBAR_BG, pady=10); metric_bg.pack(fill="x")
        mr = tk.Frame(metric_bg, bg=SIDEBAR_BG); mr.pack(padx=20)
        def mcard(lbl, colour):
            f = tk.Frame(mr, bg="#002855", padx=20, pady=10,
                         highlightbackground="#0A3A6A", highlightthickness=1)
            f.pack(side="left", padx=(0,10))
            tk.Label(f, text=lbl, bg="#002855", fg=SIDEBAR_TEXT, font=("Segoe UI",9)).pack(anchor="w")
            v = tk.Label(f, text="0", bg="#002855", fg=colour, font=("Segoe UI",22,"bold"))
            v.pack(anchor="w"); return v
        self._m_total      = mcard("Total sent",                        FBC_ACCENT)
        self._m_waiting    = mcard("Awaiting client reply",             FBC_ACCENT)
        self._m_overdue_cl = mcard(f"Overdue (>{OVERDUE_DAYS}d)",       "#FF4444")
        self._m_replied    = mcard("Reply received",                    "#A855F7")
        self._m_overdue_bo = mcard("Back office overdue",               "#FF4444")
        self._m_done       = mcard("Completed",                         "#6EE7B7")

        # tabs
        tab_row = tk.Frame(self, bg=BG, padx=20, pady=8); tab_row.pack(fill="x")
        self._tabs = {}
        for key, lbl in [("all","All Clients"),("waiting","Awaiting Client"),
                          ("overdue_cl",f"Overdue (>{OVERDUE_DAYS}d)"),
                          ("replied","Reply Received"),
                          ("overdue_bo","Back Office Overdue"),("done","Completed")]:
            b = tk.Button(tab_row, text=lbl, font=("Segoe UI",9), relief="flat",
                          cursor="hand2", padx=12, pady=6,
                          command=lambda k=key: self._set_filter(k))
            b.pack(side="left", padx=(0,4)); self._tabs[key] = b
        self._paint_tabs()
        tk.Frame(self, bg=SEP_CLR, height=1).pack(fill="x", padx=20)

        # list
        lo = tk.Frame(self, bg=BG)
        lo.pack(fill="both", expand=True, padx=20, pady=10)
        self._canvas = tk.Canvas(lo, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(lo, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); self._canvas.pack(side="left", fill="both", expand=True)
        self._inner = tk.Frame(self._canvas, bg=BG)
        self._inner_id = self._canvas.create_window((0,0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._inner_id, width=e.width))
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1*(e.delta//120),"units"))

    def _paint_tabs(self):
        for k,b in self._tabs.items():
            if k==self._filter: b.config(bg=FBC_DARK, fg=WHITE)
            else: b.config(bg=BG, fg="#607080",
                           activebackground=SEP_CLR, activeforeground=FBC_DARK)

    def _set_filter(self, key):
        self._filter=key; self._paint_tabs(); self._refresh()

    def _update_sync_badge(self):
        if self.db and self.db.online:
            self._sync_lbl.config(text="☁  Synced", fg="#6EE7B7")
        elif self.db:
            self._sync_lbl.config(text="⚠  Offline (local cache)", fg="#FFB347")
        else:
            self._sync_lbl.config(text="⚠  No sync configured", fg="#607080")

    def _full_sync(self):
        def _do():
            clients = self.db.read_all() if self.db else load_clients_local()
            self.after(0, lambda: self._apply_clients(clients))
        threading.Thread(target=_do, daemon=True).start()

    def _apply_clients(self, clients):
        self.clients=clients; self._update_sync_badge(); self._refresh()

    def _schedule_auto_sync(self):
        self._full_sync()
        self.after(30_000, self._schedule_auto_sync)

    def _refresh(self):
        all_c      = self.clients
        done       = [c for c in all_c if c.get("done")]
        active     = [c for c in all_c if not c.get("done")]
        waiting    = [c for c in active if not c.get("reply_date")]
        overdue_cl = [c for c in waiting if days_since(c.get("sent_date",""))>=OVERDUE_DAYS]
        replied    = [c for c in active if c.get("reply_date")]
        overdue_bo = [c for c in replied
                      if c.get("reply_has_attach",True)
                      and days_since(c.get("reply_date",""))>=2]
        self._m_total.config(     text=str(len(all_c)))
        self._m_waiting.config(   text=str(len(waiting)))
        self._m_overdue_cl.config(text=str(len(overdue_cl)))
        self._m_replied.config(   text=str(len(replied)))
        self._m_overdue_bo.config(text=str(len(overdue_bo)))
        self._m_done.config(      text=str(len(done)))
        if   self._filter=="waiting":    visible=waiting
        elif self._filter=="overdue_cl": visible=overdue_cl
        elif self._filter=="replied":    visible=replied
        elif self._filter=="overdue_bo": visible=overdue_bo
        elif self._filter=="done":       visible=done
        else:                            visible=all_c
        for w in self._inner.winfo_children(): w.destroy()
        if not visible:
            msg = ("No clients yet. Click  '✉ Send New Form'  to get started."
                   if self._filter=="all" else "No clients in this category.")
            tk.Label(self._inner, text=msg, bg=BG, fg="#8096B0",
                     font=("Segoe UI",11), pady=40).pack()
            return
        for client in visible:
            self._client_card(self._inner, client)

    def _client_card(self, parent, client):
        done=client.get("done",False); reply_date=client.get("reply_date","")
        has_attach=client.get("reply_has_attach",True); entry_id=client.get("reply_entry_id","")
        st,sc,sbg=status_of(client)
        days=days_since(client.get("sent_date","")); rems=client.get("reminders",0)
        sent_by=client.get("sent_by","")
        is_overdue_cl=(not done and not reply_date and days>=OVERDUE_DAYS)
        card=tk.Frame(parent, bg=OVERDUE_BG if is_overdue_cl else CARD_BG, padx=14, pady=12,
                      highlightbackground=OVERDUE_CLR if is_overdue_cl else SEP_CLR,
                      highlightthickness=1)
        card.pack(fill="x", pady=4)
        cbg=OVERDUE_BG if is_overdue_cl else CARD_BG
        initials="".join(w[0] for w in client["name"].split()[:2]).upper()
        av_bg=(GREEN_DARK if done else ("#B71C1C" if is_overdue_cl
               else ("#B45309" if (reply_date and not has_attach)
               else ("#6B21A8" if reply_date else FBC_DARK))))
        tk.Label(card, text=initials, bg=av_bg, fg=WHITE,
                 font=("Segoe UI",11,"bold"), width=3, padx=4, pady=6
                 ).pack(side="left", padx=(0,14))
        info=tk.Frame(card, bg=cbg); info.pack(side="left", fill="x", expand=True)
        tk.Label(info, text=client["name"], bg=cbg, fg=FBC_DARK,
                 font=("Segoe UI",11,"bold")).pack(anchor="w")
        meta=(f"{client['email']}   ·   {client['form_type']}"
              f"   ·   Sent: {client.get('sent_date','?')}   ·   {days}d ago"
              +(f"   ·   By: {sent_by}" if sent_by else "")
              +(f"   ·   📨 Replied: {reply_date}" if reply_date else "")
              +("   ·   ⚠ No attachment" if (reply_date and not has_attach) else "")
              +(f"   ·   Reminders: {rems}" if rems else ""))
        tk.Label(info, text=meta, bg=cbg, fg="#607080",
                 font=("Segoe UI",8)).pack(anchor="w", pady=(2,4))
        tk.Label(info, text=f"  {st}  ", bg=sbg, fg=sc,
                 font=("Segoe UI",8,"bold"), padx=4, pady=2).pack(anchor="w")
        if not done:
            ag=aging_label(client)
            if ag:
                al,ac,ab=ag
                tk.Label(info, text=f"  {al}  ", bg=ab, fg=ac,
                         font=("Segoe UI",8), padx=4, pady=1).pack(anchor="w", pady=(2,0))
        right=tk.Frame(card, bg=cbg); right.pack(side="right")
        if not done:
            if not reply_date:
                tk.Button(right, text="🔔  Reminder", font=("Segoe UI",9),
                          bg=AMBER_BG, fg=AMBER, relief="flat", cursor="hand2",
                          activebackground="#FDE68A",
                          command=lambda c=client: self._open_reminder(c)
                          ).pack(side="left", padx=(0,6))
                tk.Button(right, text="📨  Reply Received", font=("Segoe UI",9),
                          bg="#F3E8FF", fg="#6B21A8", relief="flat", cursor="hand2",
                          activebackground="#E9D5FF",
                          command=lambda c=client: self._mark_replied(c)
                          ).pack(side="left", padx=(0,6))
            else:
                if entry_id:
                    tk.Button(right, text="📧  See Email", font=("Segoe UI",9),
                              bg="#EEF2FF", fg="#3730A3", relief="flat", cursor="hand2",
                              activebackground="#C7D2FE",
                              command=lambda eid=entry_id: self._open_reply_email(eid)
                              ).pack(side="left", padx=(0,6))
                if not has_attach:
                    tk.Button(right, text="🔔  Reminder", font=("Segoe UI",9),
                              bg=AMBER_BG, fg=AMBER, relief="flat", cursor="hand2",
                              activebackground="#FDE68A",
                              command=lambda c=client: self._open_reminder(c)
                              ).pack(side="left", padx=(0,6))
                else:
                    tk.Button(right, text="✅  Mark Completed", font=("Segoe UI",9),
                              bg=GREEN_LIGHT_BG, fg=GREEN_DARK, relief="flat", cursor="hand2",
                              activebackground="#BBF7D0",
                              command=lambda c=client: self._open_confirm(c)
                              ).pack(side="left", padx=(0,6))
        tk.Button(right, text="🗑", font=("Segoe UI",10), bg=cbg, fg="#CBD5E1",
                  relief="flat", cursor="hand2", activebackground=BG,
                  command=lambda c=client: self._delete(c)).pack(side="left")

    # ── actions ───────────────────────────────────────────────────────────
    def _open_reply_email(self, entry_id):
        ok, err = open_outlook_email_by_entry_id(entry_id)
        if not ok:
            messagebox.showerror("Could not open email",
                f"{err}\n\nThe email may have been moved or deleted.", parent=self)

    def _open_sheets_setup(self):
        SheetsSetupDialog(self, self.settings, self._on_sheets_saved)

    def _on_sheets_saved(self, sid, key):
        if sid and key:
            self.settings["sheet_id"]=sid; self.settings["key_file"]=key
            self.db=SheetsDB(key,sid); self._full_sync()
        self._update_sync_badge()

    def _open_cc_settings(self):
        def _on_saved(new_cc):
            self.settings["cc_address"] = new_cc
        CCSettingsDialog(self, self.settings, _on_saved)

    def _clear_name(self):
        if messagebox.askyesno("Change User",
            "This will clear your saved name.\n\n"
            "You will be asked for your name next time you open the app.\n\nContinue?",
            parent=self):
            save_settings({**self.settings,"sender_name":""})
            messagebox.showinfo("Done","Name cleared. Please restart the app.",parent=self)

    def _open_send(self):
        SendFormDialog(self, self.sender_name, self.settings, self._on_sent)

    def _on_sent(self, rec):
        def _do():
            if self.db: self.db.append_client(rec)
            self.clients.append(rec); save_clients_local(self.clients)
            self.after(0, self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _open_reminder(self, client):
        ReminderDialog(self, client, self.sender_name, self.settings,
                       lambda c=client: self._on_reminder(c))

    def _on_reminder(self, client):
        client["reminders"]=client.get("reminders",0)+1
        existing=client.get("reminder_dates",""); today=date.today().isoformat()
        client["reminder_dates"]=(existing+","+today).strip(",")
        def _do():
            if self.db: self.db.update_client(client)
            save_clients_local(self.clients); self.after(0,self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _mark_replied(self, client):
        if not messagebox.askyesno("Mark Reply Received",
                f"Confirm that {client['name']} has returned their signed form(s)?\n\n"
                "The 2-day back office processing clock will start from today.",
                parent=self): return
        client["reply_date"]=date.today().isoformat()
        client["reply_has_attach"]=True; client["reply_entry_id"]=""
        def _do():
            if self.db: self.db.update_client(client)
            save_clients_local(self.clients); self.after(0,self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _open_confirm(self, client):
        ConfirmDialog(self, client, lambda c=client: self._on_confirm(c))

    def _on_confirm(self, client):
        client["done"]=True; client["done_date"]=date.today().isoformat()
        def _do():
            if self.db: self.db.update_client(client)
            save_clients_local(self.clients); self.after(0,self._refresh)
        threading.Thread(target=_do, daemon=True).start()

    def _delete(self, client):
        if not messagebox.askyesno("Delete record",
            f"Permanently delete the record for {client['name']}?\n\nThis cannot be undone.",
            parent=self): return
        def _do():
            if self.db: self.db.delete_client(client["id"])
            if client in self.clients: self.clients.remove(client)
            save_clients_local(self.clients); self.after(0,self._refresh)
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

    if (not settings.get("sheet_id") or not settings.get("key_file")
            or not os.path.exists(settings.get("key_file",""))):
        _tmp=tk.Tk(); _tmp.withdraw()
        _saved={}
        def _on_setup(sid,key):
            _saved["sheet_id"]=sid; _saved["key_file"]=key
        dlg=SheetsSetupDialog(_tmp, settings, _on_setup)
        _tmp.wait_window(dlg); _tmp.destroy()
        if _saved: settings.update(_saved); save_settings(settings)

    settings = load_settings()
    app = App(sender_name=login.sender_name, settings=settings)
    app.mainloop()
