@echo off
cd /d C:\Users\jehaz\OneDrive\Escritorio\zebra-printer\zebra-print-app

start "Zebra Print Server" /min cmd /k "npm start"

timeout /t 2 >nul

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    start "" "%ProgramFiles%\Google\Chrome\Application\chrome.exe" --app=http://localhost:3000
) else (
    start "" http://localhost:3000
)
