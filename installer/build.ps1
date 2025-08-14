$ErrorActionPreference = "Stop"

# Path to Inno Setup Compiler (adjust if installed elsewhere)
$ISCC = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (-not (Test-Path $ISCC)) {
  throw "Could not find Inno Setup compiler at: $ISCC"
}

# Verify files exist
if (-not (Test-Path "..\dist\BloatGuard\BloatGuard.exe")) {
  throw "Missing ..\dist\BloatGuard\BloatGuard.exe"
}
if (-not (Test-Path "..\dist\BloatGuardAgent\BloatGuardAgent.exe")) {
  throw "Missing ..\dist\BloatGuardAgent\BloatGuardAgent.exe"
}

# Build the installer
& "$ISCC" "bloatguard.iss"

Write-Host "✓ Build complete. Check installer\Output\BloatGuard-Setup.exe"
