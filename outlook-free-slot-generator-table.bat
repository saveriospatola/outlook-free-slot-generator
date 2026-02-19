@echo off
REM --- Lancia lo script PowerShell ignorando temporaneamente la policy ---
REM powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\outlook-free-slot-generator-table.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\outlook-free-slot-generator.ps1" -Formato Tabella