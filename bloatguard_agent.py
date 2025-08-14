import sys, os, json
from pathlib import Path
import subprocess

APP_NAME = "BloatGuard"
PROGRAM_DATA = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / APP_NAME
CONFIG_PATH = PROGRAM_DATA / "bloatguard.config.json"
LOG_PATH = PROGRAM_DATA / "bloatguard.log"

def log(msg):
    PROGRAM_DATA.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(msg.rstrip() + "\n")

def run(cmd):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True)
        return p.returncode, p.stdout.strip(), p.stderr.strip()
    except Exception as e:
        return 1, "", str(e)

def run_ps(cmd):
    full = ["powershell","-NoProfile","-ExecutionPolicy","Bypass","-Command",cmd]
    return run(full)

def winget_available():
    rc, out, _ = run(["winget","--version"])
    return rc == 0

def detect_edge():
    rc, out, _ = run(["winget","list","--id","Microsoft.Edge","--source","winget"])
    return ("Microsoft Edge" in out) and rc == 0

def uninstall_edge():
    if not winget_available(): return False, "winget unavailable"
    rc, out, err = run(["winget","uninstall","--id","Microsoft.Edge","--source","winget","--silent","--accept-package-agreements","--accept-source-agreements"])
    return rc == 0, out + ("\n"+err if err else "")

def detect_store():
    rc, out, err = run_ps("Get-AppxPackage Microsoft.WindowsStore | Select-Object -ExpandProperty PackageFullName")
    return (rc == 0) and out.strip()

def uninstall_store():
    rc, out, err = run_ps("Get-AppxPackage -AllUsers Microsoft.WindowsStore | Remove-AppxPackage")
    return rc == 0, (out + ("\n"+err if err else ""))

def detect_office():
    if not winget_available(): return False
    rc, out, _ = run(["winget","list"])
    if rc != 0: return False
    low = out.lower()
    return any(x in low for x in ["microsoft 365","microsoft office"," word "," excel "])

def uninstall_office():
    ok = False; details = []
    rc, out, err = run(["winget","uninstall","--id","Microsoft.Office","--silent","--accept-package-agreements","--accept-source-agreements"])
    if rc == 0: ok = True; details.append(out)
    rc2, list_out, _ = run(["winget","list"])
    if rc2 == 0:
        for line in list_out.splitlines():
            name = line.split("  ")[0].strip()
            if name.lower().startswith("microsoft") and any(w in line.lower() for w in ["office"," 365"," word "," excel "]):
                rc3, out3, err3 = run(["winget","uninstall",name,"--silent","--accept-package-agreements","--accept-source-agreements"])
                ok = ok or (rc3 == 0); details.append(out3 or err3)
    return ok, "\n".join(details)

def load_cfg():
    if CONFIG_PATH.exists():
        try: return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except: pass
    return {"edge": False, "store": False, "office": False}

def main():
    cfg = load_cfg()
    log("== BloatGuard agent start ==")
    if cfg.get("edge") and detect_edge():
        ok, msg = uninstall_edge(); log(f"Edge ok={ok}\n{msg}")
    if cfg.get("store") and detect_store():
        ok, msg = uninstall_store(); log(f"Store ok={ok}\n{msg}")
    if cfg.get("office") and detect_office():
        ok, msg = uninstall_office(); log(f"Office ok={ok}\n{msg}")
    log("== BloatGuard agent end ==")

if __name__ == "__main__":
    main()
