@echo off
rem ---------------------------------------------------------------
rem  Keep this file ASCII-only.
rem  Korean text inside a .bat breaks: cmd reads the file with the
rem  active console codepage, and when that does not match the file
rem  encoding the byte offsets drift and leading characters of the
rem  following lines get eaten. The active codepage differs depending
rem  on how the file is launched, so it cannot be pinned reliably.
rem  All real work lives in build.ps1 (UTF-8 with BOM).
rem ---------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
pause
