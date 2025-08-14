import os, sys, json, subprocess, ctypes
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_NAME = "BloatGuard"
# Standard data dir for all users
PROGRAM_DATA = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / APP_NAME
PROGRAM_DATA.mkdir(parents=True, exist_ok=True)

CONFIG_PATH = PROGRAM_DATA / "bloatguard.config.json"
LOG_PATH = PROGRAM_DATA / "bloatguard.log"
TASK_NAME = "BloatGuard_Enforce"

def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def relaunch_as_admin():
    params = " ".join([f'"{arg}"' if " " in arg else arg for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)

def run(cmd, shell=False):
    try:
        p = subprocess.run(cmd, shell=shell, capture_output=True, text=True)
        return p.returncode, p.stdout.strip(), p.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def run_ps(cmd):
    full = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd]
    return run(full)

def log(msg: str):
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(msg.rstrip() + "\n")

def winget_available() -> bool:
    rc, out, _ = run(["winget", "--version"])
    return rc == 0

def detect_edge() -> bool:
    rc, out, _ = run(["winget", "list", "--id", "Microsoft.Edge", "--source", "winget"])
    return ("Microsoft Edge" in out) and rc == 0

def uninstall_edge():
    if not winget_available():
        return False, "winget not available."
    rc, out, err = run(["winget","uninstall","--id","Microsoft.Edge","--source","winget","--silent","--accept-package-agreements","--accept-source-agreements"])
    return rc == 0, out + ("\n" + err if err else "")

def detect_store() -> bool:
    rc, out, err = run_ps("Get-AppxPackage Microsoft.WindowsStore | Select-Object -ExpandProperty PackageFullName")
    return (rc == 0) and out.strip()

def uninstall_store():
    cmd = "Get-AppxPackage -AllUsers Microsoft.WindowsStore | Remove-AppxPackage"
    rc, out, err = run_ps(cmd)
    return rc == 0, (out + ("\n" + err if err else ""))

def detect_office() -> bool:
    if not winget_available(): 
        return False
    rc, out, _ = run(["winget", "list"])
    if rc != 0: return False
    markers = ["Microsoft 365", "Microsoft Office", " Word ", " Excel "]
    low = out.lower()
    return any(m.lower() in low for m in markers)

def uninstall_office():
    if not winget_available():
        return False, "winget not available."
    rc, out, err = run(["winget","uninstall","--id","Microsoft.Office","--silent","--accept-package-agreements","--accept-source-agreements"])
    outputs = []
    attempted = []
    if rc == 0:
        outputs.append(out)
        attempted.append("Microsoft.Office (ID)")
    rc2, list_out, _ = run(["winget","list"])
    if rc2 == 0:
        for line in list_out.splitlines():
            name = line.split("  ")[0].strip()
            if name.lower().startswith("microsoft") and any(w in line.lower() for w in ["office"," 365"," word "," excel "]):
                rc3, out3, err3 = run(["winget","uninstall",name,"--silent","--accept-package-agreements","--accept-source-agreements"])
                outputs.append(out3 if out3 else err3)
                attempted.append(name)
    return bool(attempted), "Attempted: " + ", ".join(attempted) + "\n" + "\n".join(outputs)

DEFAULT_CONFIG = {"edge": False, "store": False, "office": False, "enforce": False}

def load_config():
    try:
        if CONFIG_PATH.exists():
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")

def enforce(cfg):
    log(f"== {APP_NAME} enforce run ==")
    if cfg.get("edge") and detect_edge():
        ok, msg = uninstall_edge(); log(f"Edge uninstall ok={ok}\n{msg}")
    if cfg.get("store") and detect_store():
        ok, msg = uninstall_store(); log(f"Store uninstall ok={ok}\n{msg}")
    if cfg.get("office") and detect_office():
        ok, msg = uninstall_office(); log(f"Office uninstall ok={ok}\n{msg}")

def create_task():
    script = f'"{sys.executable}" "{Path(__file__).resolve()}" --enforce'
    cmd = ["schtasks","/Create","/TN",TASK_NAME,"/TR",script,"/SC","ONLOGON","/RL","HIGHEST","/F"]
    rc, out, err = run(cmd); return rc == 0, out + ("\n" + err if err else "")

def delete_task():
    rc, out, err = run(["schtasks","/Delete","/TN",TASK_NAME,"/F"])
    return rc == 0, out + ("\n" + err if err else "")

def task_exists():
    rc, out, err = run(["schtasks","/Query","/TN",TASK_NAME])
    return rc == 0

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME); self.geometry("600x460"); self.resizable(False, False)
        self.cfg = load_config()

        topbar = ttk.Frame(self, padding=10); topbar.pack(fill="x")
        self.admin_label = ttk.Label(topbar, text="Running as Administrator ✅" if is_admin() else "Not elevated ❌ — click Elevate")
        self.admin_label.pack(side="left")
        ttk.Button(topbar, text="Elevate", command=self.on_elevate).pack(side="right")

        status = ttk.LabelFrame(self, text="Detection", padding=10); status.pack(fill="x", padx=10, pady=6)
        self.edge_status = ttk.Label(status, text="Edge: (checking...)")
        self.store_status = ttk.Label(status, text="Microsoft Store: (checking...)")
        self.office_status = ttk.Label(status, text="Office/365 (Word/Excel): (checking...)")
        for w in (self.edge_status,self.store_status,self.office_status): w.pack(anchor="w")

        choices = ttk.LabelFrame(self, text="Select items to remove (consent-based)", padding=10); choices.pack(fill="x", padx=10, pady=6)
        self.var_edge = tk.BooleanVar(value=self.cfg.get("edge", False))
        self.var_store = tk.BooleanVar(value=self.cfg.get("store", False))
        self.var_office = tk.BooleanVar(value=self.cfg.get("office", False))
        ttk.Checkbutton(choices, text="Microsoft Edge", variable=self.var_edge).pack(anchor="w")
        ttk.Checkbutton(choices, text="Microsoft Store (not recommended; may break features)", variable=self.var_store).pack(anchor="w")
        ttk.Checkbutton(choices, text="Microsoft 365/Office (includes Word & Excel)", variable=self.var_office).pack(anchor="w")

        enforce_frame = ttk.LabelFrame(self, text="Ongoing enforcement (optional, transparent)", padding=10); enforce_frame.pack(fill="x", padx=10, pady=6)
        self.var_enforce = tk.BooleanVar(value=self.cfg.get("enforce", False))
        ttk.Checkbutton(enforce_frame, text="Re-check at logon and re-remove selected items if they reappear", variable=self.var_enforce).pack(anchor="w")

        actions = ttk.Frame(self, padding=10); actions.pack(fill="x")
        ttk.Button(actions, text="Save Choices", command=self.on_save).pack(side="left")
        ttk.Button(actions, text="Uninstall Selected Now", command=self.on_uninstall_now).pack(side="left", padx=10)
        ttk.Button(actions, text="Open Data Folder", command=lambda: os.startfile(PROGRAM_DATA)).pack(side="right")
        ttk.Button(actions, text="Remove Enforce Task", command=self.on_remove_task).pack(side="right", padx=8)

        note = ttk.Label(self, wraplength=560, foreground="#444",
            text="Heads-up: Removing Microsoft Store can impact Windows features.\n"
                 "Edge/Office are reinstallable via winget. Store requires additional steps.")
        note.pack(padx=10, pady=6)

        self.after(200, self.refresh_detection)

    def ensure_admin(self):
        if is_admin(): return True
        messagebox.showwarning(APP_NAME, "Please run elevated (Administrator) to proceed."); return False

    def on_elevate(self):
        relaunch_as_admin(); self.after(800, self.destroy)

    def refresh_detection(self):
        e = "✅ Present" if detect_edge() else "✔️ Not installed"
        s = "✅ Present" if detect_store() else "✔️ Not installed"
        o = "✅ Present" if detect_office() else "✔️ Not installed"
        self.edge_status.config(text=f"Edge: {e}")
        self.store_status.config(text=f"Microsoft Store: {s}")
        self.office_status.config(text=f"Office/365 (Word/Excel): {o}")

    def on_save(self):
        self.cfg.update({"edge": self.var_edge.get(), "store": self.var_store.get(), "office": self.var_office.get(), "enforce": self.var_enforce.get()})
        save_config(self.cfg)
        if self.cfg["enforce"]:
            if not self.ensure_admin(): return
            ok, msg = create_task()
            messagebox.showinfo(APP_NAME, "Enforce task created." if ok else f"Failed to create task:\n{msg}")
        else:
            if task_exists():
                ok, msg = delete_task()
                if not ok: messagebox.showerror(APP_NAME, f"Failed to remove enforce task:\n{msg}")
        messagebox.showinfo(APP_NAME, "Configuration saved.")

    def on_uninstall_now(self):
        if not self.ensure_admin(): return
        actions = []
        if self.var_edge.get(): actions.append(("Microsoft Edge", uninstall_edge))
        if self.var_store.get(): actions.append(("Microsoft Store", uninstall_store))
        if self.var_office.get(): actions.append(("Microsoft 365/Office", uninstall_office))
        if not actions:
            messagebox.showinfo(APP_NAME, "No items selected."); return
        summary = []
        for name, fn in actions:
            ok, msg = fn(); summary.append(f"[{name}] success={ok}\n{msg}\n")
        messagebox.showinfo(APP_NAME, "\n".join(summary)); self.refresh_detection()

    def on_remove_task(self):
        if not task_exists():
            messagebox.showinfo(APP_NAME, "No enforce task found."); return
        if not self.ensure_admin(): return
        ok, msg = delete_task()
        messagebox.showinfo(APP_NAME, "Enforce task removed." if ok else f"Failed to remove task:\n{msg}")

def main():
    if "--enforce" in sys.argv:
        enforce(load_config()); return
    if os.name != "nt":
        print("Windows only."); return
    if not winget_available():
        messagebox.showwarning(APP_NAME, "winget is required for most actions. Please install/enable it first.")
    app = App(); app.mainloop()

if __name__ == "__main__":
    main()
