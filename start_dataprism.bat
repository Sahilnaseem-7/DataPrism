@echo off
setlocal
cd /d "%~dp0"
set "APP_PORT=8785"
where py >nul 2>nul
if errorlevel 1 (set "PYRUN=python") else (set "PYRUN=py")
%PYRUN% --version >nul 2>nul
if errorlevel 1 (
    echo Python 3.10 or later is required. Install Python, then run this file again.
    pause
    exit /b 1
)
%PYRUN% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing the XLSX reader needed for spreadsheet uploads...
    %PYRUN% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Could not install the dependency. See README.md for manual setup.
        pause
        exit /b 1
    )
)
start "DataPrism server" /min %PYRUN% -u platform_server.py %APP_PORT%
powershell -NoProfile -Command "$ready=$false; for($i=0;$i -lt 30;$i++){try{$r=Invoke-WebRequest -UseBasicParsing -TimeoutSec 1 http://127.0.0.1:%APP_PORT%/api/health; if($r.StatusCode -eq 200){$ready=$true; break}}catch{}; Start-Sleep -Milliseconds 250}; if(-not $ready){exit 1}"
if errorlevel 1 (
    echo DataPrism did not start. Check that Python is installed and port 8785 is available.
    pause
    exit /b 1
)
where chrome >nul 2>nul
if not errorlevel 1 (
    start "" chrome --new-window "http://127.0.0.1:%APP_PORT%"
) else if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe" --new-window "http://127.0.0.1:%APP_PORT%"
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" --new-window "http://127.0.0.1:%APP_PORT%"
) else if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    start "" "%LocalAppData%\Google\Chrome\Application\chrome.exe" --new-window "http://127.0.0.1:%APP_PORT%"
) else (
    start "" "http://127.0.0.1:%APP_PORT%"
)
echo DataPrism is open at http://127.0.0.1:%APP_PORT%
echo Keep the minimized DataPrism server window open while you use it.
pause
