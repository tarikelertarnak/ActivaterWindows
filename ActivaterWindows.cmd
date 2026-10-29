@echo off
title ActivaterWindows
rem ActivaterWindows v3.0.2 - Terminal (cift tik baslatici)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ActivaterWindows.ps1" %*
pause
