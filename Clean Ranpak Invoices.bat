@echo off
rem Drag one or more Ranpak Excel backup files onto this file.
python "%~dp0ranpak_clean.py" %*
pause
