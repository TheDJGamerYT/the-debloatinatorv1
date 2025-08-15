import sys, os, json, subprocess, time
from pathlib import Path

APP_NAME = "BloatGuard"
PROGRAM_DATA = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / APP_NAME
CONFIG_PATH = PROGRAM_DATA / "bloatguard.config.json"
LOG_PATH = PROGRAM_DATA / "bloatguard.log"

CREATE_NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0

def log(msg):
    PROGRAM_DATA.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(msg.rstrip() + "\n")

def run(cmd):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, creationflags=CREATE_NO_WINDOW)
        return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()
    except Exception as e:
        return 1, "", str(e)

def run_ps(ps_cmd):
    full = ["powershell","-NoProfile","-ExecutionPolicy","Bypass","-Command",ps_cmd]
    return run(full)

def winget_available(): rc,_,_=run(["winget","--version"]); return rc==0
def detect_edge(): rc,out,_=run(["winget","list","--id","Microsoft.Edge","--source","winget"]); return rc==0 and ("Microsoft Edge" in out)
def uninstall_edge(): 
    if not winget_available(): return False,"winget unavailable"
    rc,out,err=run(["winget","uninstall","--id","Microsoft.Edge","--source","winget","--silent","--accept-package-agreements","--accept-source-agreements"]); 
    return rc==0,(out+("\n"+err if err else ""))

def detect_store(): rc,out,_=run_ps("Get-AppxPackage Microsoft.WindowsStore | Select-Object -ExpandProperty PackageFullName"); return rc==0 and out.strip()!=""
def uninstall_store(): rc,out,err=run_ps("Get-AppxPackage -AllUsers Microsoft.WindowsStore | Remove-AppxPackage"); return rc==0,(out+("\n"+err if err else ""))

def detect_office():
    if not winget_available(): return False
    rc,out,_=run(["winget","list"]); 
    if rc!=0: return False
    low=out.lower(); return any(x in low for x in ["microsoft 365","microsoft office"," word "," excel "])
def uninstall_office():
    ok=False; details=[]
    rc,out,err=run(["winget","uninstall","--id","Microsoft.Office","--silent","--accept-package-agreements","--accept-source-agreements"])
    if rc==0: ok=True; details.append(out)
    rc2,out2,_=run(["winget","list"])
    if rc2==0:
        for line in out2.splitlines():
            name=line.split("  ")[0].strip(); low=line.lower()
            if name.lower().startswith("microsoft") and any(w in low for w in ["office"," 365"," word "," excel "]):
                rc3,o3,e3=run(["winget","uninstall",name,"--silent","--accept-package-agreements","--accept-source-agreements"])
                ok = ok or (rc3==0); details.append(o3 or e3)
    return ok,"\n".join(details)

# Copilot
def disable_copilot():
    cmds = [
        r'New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows" -Name "WindowsCopilot" -Force | Out-Null',
        r'Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" -Name "TurnOffWindowsCopilot" -Type DWord -Value 1',
        r'Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowCopilotButton" -Type DWord -Value 0'
    ]
    ok=True; outs=[]
    for c in cmds:
        rc,out,err=run_ps(c); ok = ok and (rc==0); outs.append(out or err)
    return ok,"\n".join(outs)

def remove_copilot_webxp():
    rc,out,err = run_ps("Get-AppxPackage -AllUsers *WebExperience* | Remove-AppxPackage")
    return rc==0,(out+("\n"+err if err else ""))

def load_cfg():
    if CONFIG_PATH.exists():
        try: return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except: pass
    return {"edge":False,"store":False,"office":False,"copilot_disable":False,"copilot_remove_webxp":False,"binder_enabled":False,"binder_target":""}

def launch_file(path):
    try:
        os.startfile(path)
        return True,""
    except Exception as e:
        return False,str(e)

def run_enforcement(cfg):
    log("== BloatGuard agent start ==")
    if cfg.get("edge")   and detect_edge():  ok,msg=uninstall_edge();  log(f"Edge ok={ok}\n{msg}")
    if cfg.get("store")  and detect_store(): ok,msg=uninstall_store(); log(f"Store ok={ok}\n{msg}")
    if cfg.get("office") and detect_office():ok,msg=uninstall_office();log(f"Office ok={ok}\n{msg}")
    if cfg.get("copilot_disable"): ok,msg=disable_copilot(); log(f"Copilot disable ok={ok}\n{msg}")
    if cfg.get("copilot_remove_webxp"): ok,msg=remove_copilot_webxp(); log(f"Copilot WebXP remove ok={ok}\n{msg}")
    log("== BloatGuard agent end ==")

def run_binder_loop(target):
    # Requires package 'keyboard' and admin privileges
    try:
        import keyboard
    except Exception as e:
        log(f"Binder unavailable (keyboard module missing): {e}")
        return
    log(f"Binder active: Win+C -> {target}")
    try:
        keyboard.clear_all_hotkeys()
        keyboard.add_hotkey('windows+c', lambda: launch_file(target))
        while True:
            time.sleep(0.5)
    except Exception as e:
        log(f"Binder error: {e}")

def main():
    cfg = load_cfg()
    # Fire-and-forget enforcement
    run_enforcement(cfg)
    # Optional binder stays resident
    if cfg.get("binder_enabled") and cfg.get("binder_target"):
        run_binder_loop(cfg["binder_target"])

if __name__ == "__main__":
    main()
