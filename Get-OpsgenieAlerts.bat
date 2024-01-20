@echo off
NET SESSION > nul 2>&1
IF %ERRORLEVEL% NEQ 0 GOTO ELEVATE
GOTO ADMIN

:ELEVATE
MSHTA "javascript: var shell = new ActiveXObject('shell.application'); shell.ShellExecute('%~nx0', '', '', 'runas', 1);close();"
EXIT

:ADMIN
powershell.exe -NoLogo -Sta -NoProfile -NoExit -ExecutionPolicy Bypass -File "%~dp0Get-OpsgenieAlerts.ps1"
