@echo off
echo === Gmail - testmodus ===
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0MailRydder.ps1" -DryRun
pause
