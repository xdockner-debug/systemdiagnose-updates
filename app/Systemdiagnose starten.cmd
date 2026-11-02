@echo off
setlocal
title Systemdiagnose (nur lesend)
set "SKRIPT=%~dp0Systemdiagnose.ps1"

:menue
cls
echo.
echo   ============================================================
echo    Systemdiagnose  -  prueft Windows, aendert nichts
echo   ============================================================
echo.
echo    [1]  Schnell      ohne Adminrechte, ca. 20-30 Sekunden
echo    [2]  Tief         mit Adminrechten: SMART-Werte, TPM, DISM,
echo                      Wiederherstellungspunkte - ca. 1 Minute
echo    [3]  Gruendlich   wie Tief + Systemdatei-Pruefung (SFC)
echo                      - ca. 5-20 Minuten
echo    [4]  Sicherheits-Selbsttest anzeigen
echo    [5]  Berichtsordner oeffnen
echo    [0]  Beenden
echo.
choice /c 123450 /n /m "   Auswahl: "
if errorlevel 6 goto :eof
if errorlevel 5 goto ordner
if errorlevel 4 goto selbsttest
if errorlevel 3 goto gruendlich
if errorlevel 2 goto tief

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SKRIPT%" -Modus Schnell
echo.
pause
goto :eof

:tief
echo.
echo   Windows fragt gleich nach Adminrechten - die Diagnose oeffnet sich in einem neuen Fenster.
powershell.exe -NoProfile -Command "Start-Process powershell.exe -Verb RunAs -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File \"' + $env:SKRIPT + '\" -Modus Tief -Pause')"
goto :eof

:gruendlich
echo.
echo   Windows fragt gleich nach Adminrechten - die Diagnose oeffnet sich in einem neuen Fenster.
powershell.exe -NoProfile -Command "Start-Process powershell.exe -Verb RunAs -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File \"' + $env:SKRIPT + '\" -Modus Gruendlich -Pause')"
goto :eof

:selbsttest
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SKRIPT%" -Selbsttest
echo.
pause
goto menue

:ordner
if exist "%~dp0Berichte" (
    start "" "%~dp0Berichte"
) else (
    echo.
    echo   Noch keine Berichte vorhanden.
    pause
)
goto menue
