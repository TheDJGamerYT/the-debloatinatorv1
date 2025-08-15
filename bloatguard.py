import os, sys, json, subprocess, ctypes
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_NAME = "BloatGuard"
PROGRAM_DATA = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / APP_NAME
PROGRAM_DATA.mkdir(parents=True, exist_ok=True)

CONFIG_PATH = PROGRAM_DATA / "bloatguard.config.json"
LOG_PATH = PROGRAM_DATA / "bloatguard.log"
TASK_NAME = "BloatGuard_Enforce"

# ---------------------------
# Silent subprocess helpers
# ---------------------------
CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0

def run(cmd, shell=False):
    """Run command hidden; return (rc, stdout, stderr)."""
    try:
        p = subprocess.run(
            cmd,
            shell=shell,
            capture_output=True,
            text=True,
            creationflags=CREATE_NO_WINDOW
        )
        return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()
    except Exception as e:
        return 1, "", str(e)

def run_ps(ps_cmd):
    full = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd]
    return run(full)

def log(msg: str):
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(msg.rstrip() + "\n")

# ---------------------------
# Admin helpers
# ---------------------------
def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def relaunch_as_admin():
    params = " ".join([f'"{arg}"' if " " in arg else arg for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)

# ---------------------------
# winget / detection
# ---------------------------
def winget_available() -> bool:
    rc, _, _ = run(["winget", "--version"])
    return rc == 0

def detect_edge() -> bool:
    rc, out, _ = run(["winget", "list", "--id", "Microsoft.Edge", "--source", "winget"])
    return rc == 0 and ("Microsoft Edge" in out)

def uninstall_edge():
    if not winget_available():
        return False, "winget not available."
    rc, out, err = run(["winget","uninstall","--id","Microsoft.Edge","--source","winget","--silent","--accept-package-agreements","--accept-source-agreements"])
    return rc == 0, (out + ("\n"+err if err else ""))

def detect_store() -> bool:
    rc, out, _ = run_ps("Get-AppxPackage Microsoft.WindowsStore | Select-Object -ExpandProperty PackageFullName")
    return rc == 0 and out.strip() != ""

def uninstall_store():
    rc, out, err = run_ps("Get-AppxPackage -AllUsers Microsoft.WindowsStore | Remove-AppxPackage")
    return rc == 0, (out + ("\n"+err if err else ""))

def detect_office() -> bool:
    if not winget_available(): 
        return False
    rc, out, _ = run(["winget", "list"])
    if rc != 0: return False
    low = out.lower()
    return any(x in low for x in ["microsoft 365","microsoft office"," word "," excel "])

def uninstall_office():
    if not winget_available():
        return False, "winget not available."
    outputs, attempted, ok = [], [], False
    rc, out, err = run(["winget","uninstall","--id","Microsoft.Office","--silent","--accept-package-agreements","--accept-source-agreements"])
    if rc == 0: outputs.append(out); attempted.append("Microsoft.Office (ID)"); ok = True
    rc2, list_out, _ = run(["winget","list"])
    if rc2 == 0:
        for line in list_out.splitlines():
            name = line.split("  ")[0].strip()
            low = line.lower()
            if name.lower().startswith("microsoft") and any(w in low for w in ["office"," 365"," word "," excel "]):
                rc3, out3, err3 = run(["winget","uninstall",name,"--silent","--accept-package-agreements","--accept-source-agreements"])
                outputs.append(out3 or err3); attempted.append(name); ok = ok or (rc3 == 0)
    return ok, "Attempted: " + ", ".join(attempted) + "\n" + "\n".join(outputs)

# ---------------------------
# Copilot (disable/remove)
# ---------------------------
def detect_copilot_present() -> bool:
    # presence heuristic: Web Experience Pack or taskbar button setting
    rc, out, _ = run_ps("Get-AppxPackage *WebExperience* | Select-Object -ExpandProperty Name")
    if rc == 0 and "WebExperience" in out: 
        return True
    rc2, out2, _ = run_ps(r'Get-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name ShowCopilotButton -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ShowCopilotButton')
    return rc2 == 0 and out2.strip() == "1"

def disable_copilot():
    # Policy: Turn off Windows Copilot + hide taskbar button
    cmds = [
        r'New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "WindowsCopilot" -Force | Out-Null',
        r'Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Type DWord -Value 1',
        r'Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowCopilotButton" -Type DWord -Value 0',
        r'Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue'  # refresh taskbar
    ]
    ok_all, outs = True, []
    for c in cmds:
        rc, out, err = run_ps(c); ok_all = ok_all and (rc == 0); outs.append(out or err)
    return ok_all, "\n".join(outs)

def remove_copilot_webxp():
    # Uninstall Web Experience Pack (Store app that powers Copilot UI)
    rc, out, err = run_ps("Get-AppxPackage -AllUsers *WebExperience* | Remove-AppxPackage")
    return rc == 0, (out + ("\n"+err if err else ""))

# ---------------------------
# Config & enforcement
# ---------------------------
DEFAULT_CONFIG = {
    "edge": False,
    "store": False,
    "office": False,
    "copilot_disable": False,
    "copilot_remove_webxp": False,
    "enforce": False,
    "binder_enabled": False,
    "binder_target": ""  # full path to exe or file to open
}

def load_config():
    if CONFIG_PATH.exists():
        try:
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
    if cfg.get("copilot_disable"):
        ok, msg = disable_copilot(); log(f"Copilot disable ok={ok}\n{msg}")
    if cfg.get("copilot_remove_webxp"):
        # try remove if still present
        if detect_copilot_present():
            ok, msg = remove_copilot_webxp(); log(f"WebXP remove ok={ok}\n{msg}")

# ---------------------------
# Scheduled task management
# ---------------------------
def create_task():
    script = f'"{sys.executable}" "{Path(__file__).resolve()}" --enforce'
    cmd = ["schtasks","/Create","/TN",TASK_NAME,"/TR",script,"/SC","ONLOGON","/RL","HIGHEST","/F"]
    rc, out, err = run(cmd); return rc == 0, out + ("\n"+err if err else "")

def delete_task():
    rc, out, err = run(["schtasks","/Delete","/TN",TASK_NAME,"/F"])
    return rc == 0, out + ("\n"+err if err else "")

def task_exists():
    rc, _, _ = run(["schtasks","/Query","/TN",TASK_NAME])
    return rc == 0

# ---------------------------
# GUI
# ---------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME); self.geometry("660x600"); self.resizable(False, False)
        self.cfg = load_config()

        # Admin banner
        top = ttk.Frame(self, padding=10); top.pack(fill="x")
        ttk.Label(top, text="Running as Administrator ✅" if is_admin() else "Not elevated ❌ — click Elevate").pack(side="left")
        ttk.Button(top, text="Elevate", command=self.on_elevate).pack(side="right")

        # Detection
        det = ttk.LabelFrame(self, text="Detection", padding=10); det.pack(fill="x", padx=10, pady=6)
        self.edge_status  = ttk.Label(det, text="Edge: (checking...)"); self.edge_status.pack(anchor="w")
        self.store_status = ttk.Label(det, text="Microsoft Store: (checking...)"); self.store_status.pack(anchor="w")
        self.office_status= ttk.Label(det, text="Office/365 (Word/Excel): (checking...)"); self.office_status.pack(anchor="w")
        self.copilot_status=ttk.Label(det, text="Windows Copilot: (checking...)"); self.copilot_status.pack(anchor="w")

        # Choices
        box = ttk.LabelFrame(self, text="Select items to remove / disable (consent-based)", padding=10); box.pack(fill="x", padx=10, pady=6)
        self.var_edge   = tk.BooleanVar(value=self.cfg.get("edge", False))
        self.var_store  = tk.BooleanVar(value=self.cfg.get("store", False))
        self.var_office = tk.BooleanVar(value=self.cfg.get("office", False))
        self.var_copilot_disable = tk.BooleanVar(value=self.cfg.get("copilot_disable", False))
        self.var_copilot_remove  = tk.BooleanVar(value=self.cfg.get("copilot_remove_webxp", False))

        ttk.Checkbutton(box, text="Microsoft Edge", variable=self.var_edge).pack(anchor="w")
        ttk.Checkbutton(box, text="Microsoft Store (not recommended; may break features)", variable=self.var_store).pack(anchor="w")
        ttk.Checkbutton(box, text="Microsoft 365/Office (includes Word & Excel)", variable=self.var_office).pack(anchor="w")
        ttk.Checkbutton(box, text="Windows Copilot: disable (policy + hide button)", variable=self.var_copilot_disable).pack(anchor="w")
        ttk.Checkbutton(box, text="Windows Copilot: remove Web Experience Pack (advanced)", variable=self.var_copilot_remove).pack(anchor="w")

        # Binder UI
        bind = ttk.LabelFrame(self, text="Copilot key binder (Win+C → your app)", padding=10); bind.pack(fill="x", padx=10, pady=6)
        self.var_binder = tk.BooleanVar(value=self.cfg.get("binder_enabled", False))
        self.binder_path = tk.StringVar(value=self.cfg.get("binder_target", ""))

        row = ttk.Frame(bind); row.pack(fill="x")
        ttk.Checkbutton(row, text="Enable binder", variable=self.var_binder).pack(side="left")
        ttk.Entry(row, textvariable=self.binder_path, width=60).pack(side="left", padx=6)
        ttk.Button(row, text="Browse…", command=self.pick_app).pack(side="left")

        # Enforce
        enf = ttk.LabelFrame(self, text="Ongoing enforcement (optional, transparent)", padding=10); enf.pack(fill="x", padx=10, pady=6)
        self.var_enforce = tk.BooleanVar(value=self.cfg.get("enforce", False))
        ttk.Checkbutton(enf, text="Re-check at logon and re-apply selected items", variable=self.var_enforce).pack(anchor="w")

        # Actions
        actions = ttk.Frame(self, padding=10); actions.pack(fill="x")
        ttk.Button(actions, text="Save Choices", command=self.on_save).pack(side="left")
        ttk.Button(actions, text="Apply Now", command=self.on_apply_now).pack(side="left", padx=10)
        ttk.Button(actions, text="Open Data Folder", command=lambda: os.startfile(PROGRAM_DATA)).pack(side="right")
        ttk.Button(actions, text="Remove Enforce Task", command=self.on_remove_task).pack(side="right", padx=8)

        note = ttk.Label(self, wraplength=620, foreground="#444",
            text="Removing Microsoft Store may impact Windows features. Copilot removal uses policy + (optionally) uninstalls the Web Experience Pack. Binder requires admin and runs in the Agent.")
        note.pack(padx=10, pady=6)

        self.after(200, self.refresh_detection)

    def on_elevate(self):
        relaunch_as_admin(); self.after(800, self.destroy)

    def ensure_admin(self):
        if is_admin(): return True
        messagebox.showwarning(APP_NAME, "Please run elevated (Administrator) to proceed."); return False

    def pick_app(self):
        path = filedialog.askopenfilename(title="Select application or file to open")
        if path: self.binder_path.set(path)

    def refresh_detection(self):
        self.edge_status.config(text=f"Edge: {'✅ Present' if detect_edge() else '✔️ Not installed'}")
        self.store_status.config(text=f"Microsoft Store: {'✅ Present' if detect_store() else '✔️ Not installed'}")
        self.office_status.config(text=f"Office/365: {'✅ Present' if detect_office() else '✔️ Not installed'}")
        self.copilot_status.config(text=f"Windows Copilot: {'✅ Enabled/Present' if detect_copilot_present() else '✔️ Disabled/Not detected'}")

    def on_save(self):
        self.cfg.update({
            "edge": self.var_edge.get(),
            "store": self.var_store.get(),
            "office": self.var_office.get(),
            "copilot_disable": self.var_copilot_disable.get(),
            "copilot_remove_webxp": self.var_copilot_remove.get(),
            "enforce": self.var_enforce.get(),
            "binder_enabled": self.var_binder.get(),
            "binder_target": self.binder_path.get().strip(),
        })
        save_config(self.cfg)

        if self.cfg["enforce"]:
            if not self.ensure_admin(): return
            ok, msg = create_task()
            messagebox.showinfo(APP_NAME, "Enforce task created." if ok else f"Failed to create task:\n{msg}")
        else:
            if task_exists():
                ok, msg = delete_task()
                if not ok: messagebox.showerror(APP_NAME, f"Failed to remove task:\n{msg}")

        messagebox.showinfo(APP_NAME, "Configuration saved.")

    def on_apply_now(self):
        if not self.ensure_admin(): return
        ops = []
        if self.var_edge.get():   ops.append(("Microsoft Edge", uninstall_edge))
        if self.var_store.get():  ops.append(("Microsoft Store", uninstall_store))
        if self.var_office.get(): ops.append(("Microsoft 365/Office", uninstall_office))
        if self.var_copilot_disable.get(): ops.append(("Windows Copilot (disable)", disable_copilot))
        if self.var_copilot_remove.get():  ops.append(("Windows Copilot WebXP (remove)", remove_copilot_webxp))
        if not ops:
            messagebox.showinfo(APP_NAME, "No items selected."); return
        summary = []
        for name, fn in ops:
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
        messagebox.showwarning(APP_NAME, "winget is required for some actions. Please install/enable it first.")
    app = App(); app.mainloop()

if __name__ == "__main__":
    main()
