@echo off
setlocal
pushd "%~dp0"

echo =========================================================
echo   Build Windows EXE from the Menu Selector app
echo =========================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [Error] Python is not installed on this system.
    echo Please install Python first, and during setup make sure
    echo to check "Add Python to PATH":
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/4] Installing required packages ...
python -m pip install --upgrade pip >nul
python -m pip install openpyxl pyinstaller pillow
if errorlevel 1 (
    echo [Error] Failed to install required Python packages.
    echo Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo [2/4] Cleaning previous builds (avoids stale cached icon) ...
if exist build rd /s /q build
if exist dist rd /s /q dist
if exist MotorcycleMenuSelector.spec del /q MotorcycleMenuSelector.spec

echo.
echo [3/4] Building the EXE (no black cmd window, with icon) ...
echo A detailed log is being saved to build_log.txt
if exist icon.ico (
    python -m PyInstaller --noconfirm --onefile --windowed ^
        --name "MotorcycleMenuSelector" ^
        --icon "icon.ico" ^
        --add-data "tree_data.json;." ^
        --add-data "icon.ico;." ^
        --add-data "brand_icons;brand_icons" ^
        --add-data "ui_icons;ui_icons" ^
        motorcycle_menu_app.py > build_log.txt 2>&1
) else (
    python -m PyInstaller --noconfirm --onefile --windowed ^
        --name "MotorcycleMenuSelector" ^
        --add-data "tree_data.json;." ^
        --add-data "brand_icons;brand_icons" ^
        --add-data "ui_icons;ui_icons" ^
        motorcycle_menu_app.py > build_log.txt 2>&1
)

echo.
if not exist "dist\MotorcycleMenuSelector.exe" (
    echo =========================================================
    echo   [ERROR] Build FAILED - the exe was not created.
    echo =========================================================
    echo.
    echo Most common cause: your antivirus / Windows Defender deleted
    echo the file right after PyInstaller created it. Check:
    echo   Windows Security - Virus ^& threat protection - Protection history
    echo If you find it there, add this folder as an exclusion under
    echo   Virus ^& threat protection settings - Exclusions
    echo and run this file again.
    echo.
    echo Full build details were saved to build_log.txt in this folder.
    echo If the problem continues, send that build_log.txt file so it
    echo can be checked.
    echo.
    popd
    pause
    exit /b 1
)

echo =========================================================
echo   [4/4] Done - build succeeded.
echo =========================================================
echo The final executable is here:
echo   dist\MotorcycleMenuSelector.exe
echo You can run this file standalone on any Windows PC, no Python needed.
echo.
echo To create a desktop shortcut: right-click the exe file, choose
echo "Send to", then "Desktop (create shortcut)".
echo.
echo Note about the icon: if you already built an exe or shortcut
echo before, Windows caches the old icon. To see the new one:
echo   1) Delete the old shortcut and exe,
echo   2) Make a new shortcut from this fresh exe (inside the dist folder).
echo If the old icon still shows, restart Windows once
echo (or restart Explorer from the Task Manager) to clear the icon cache.
echo.
popd
pause
