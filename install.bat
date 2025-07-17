@echo off
echo ===================================
echo    REGAIN AGENT INSTALLER
echo ===================================
echo.

cd /d "%~dp0"

:: ----------------------------------------------------------------------------------------------
echo [..] Controllo privilegi amministratore...
net session >nul 2>&1
if %errorlevel% NEQ 0 (
    echo [KO] Esecuzione senza diritti amministrativi.
    echo [!!] Elevazione richiesta. Rilancio come amministratore...
    pause
    echo.
    powershell -Command "Start-Process '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs"
    exit /b
)
echo [OK] Esecuzione come amministratore confermata.
pause
echo.
echo [??] Cartella di installazione regain agent: %cd%
pause
echo.

:: ----------------------------------------------------------------------------------------------
echo [..] Verifica Powershell installata sul sistema con versione 5.1...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "if ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -eq 1) { exit 0 } else { exit 1 }"
if %errorlevel% EQU 0 (
    echo [OK] Versione PowerShell 5.1 confermata.
    pause
    echo.
) else (
    echo [KO] La versione PowerShell non è installata o non è compatibile. Necessaria la versione 5.1. Installazione interrotta.
    pause
    echo.
    exit /b
)

:: ----------------------------------------------------------------------------------------------
echo [..] Esecuzione script di installazione powershell...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" "VALID"
set "ps_exit_code=%errorlevel%"

echo.
echo [??] Script PowerShell terminato con codice: %ps_exit_code%

if %ps_exit_code% EQU 0 (
    echo [OK] Installazione completata con successo
) else (
    echo [KO] Installazione fallita
)

echo.
echo [??] Ritorno codice di errore finale: %ps_exit_code%
exit /b %ps_exit_code%