@echo off
echo === Exchange/Outlook - testmodus ===
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0MailRydder-Exchange.ps1" -DryRun
pause
