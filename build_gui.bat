@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --console --name doubao_voice_bridge_cli --hidden-import=win32timezone --hidden-import=win32api --hidden-import=win32con --exclude-module=pyautogui --exclude-module=pyscreeze --exclude-module=PIL --exclude-module=cv2 --exclude-module=numpy --exclude-module=PyQt5 --exclude-module=PySide6 --exclude-module=matplotlib --exclude-module=IPython --exclude-module=cryptography --exclude-module=OpenSSL --exclude-module=bcrypt feishu_voice_bridge.py
if not exist tools mkdir tools
copy /Y dist\doubao_voice_bridge_cli.exe tools\doubao_voice_bridge_cli.exe >nul
copy /Y config.example.json tools\config.example.json >nul
python -m PyInstaller --noconfirm --clean doubao_voice_bridge_gui.spec
echo.
echo Build finished:
echo %~dp0dist\DouBaoVoiceBridge.exe

