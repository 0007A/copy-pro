@echo off
title Copy Pro - Smart MCQ Extractor
cd /d "g:\My Drive\All project\copy Pro"
echo Starting Copy Pro...
"C:\Users\biswa\AppData\Local\Programs\Python\Python312\python.exe" "g:\My Drive\All project\copy Pro\main.py"
if errorlevel 1 (
    echo.
    echo ========================================================
    echo An error occurred while running Copy Pro.
    echo ========================================================
    pause
)
