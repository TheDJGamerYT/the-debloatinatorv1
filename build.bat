pyinstaller --noconfirm --clean --windowed ^
  --name BloatGuard ^
  --icon=assets\bloatguard.ico ^
  bloatguard.py
pause
pyinstaller --noconfirm --clean --console ^
  --name BloatGuardAgent ^
  bloatguard_agent.py
pause