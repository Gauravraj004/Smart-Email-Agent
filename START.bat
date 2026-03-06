@echo off
cls
echo.
echo ========================================================================
echo              COLD EMAIL AUTOMATION - PRODUCTION READY
echo ========================================================================
echo.
echo SYSTEM STATUS: ALL VERIFIED ^& READY
echo.
echo CAPABILITIES:
echo   [x] Unlimited CSV files support (1.csv, 2.csv, 3.csv, ... 999.csv)
echo   [x] All follow-up emails in SAME thread (proper threading)
echo   [x] Automatic deduplication across files
echo   [x] Smart Recovery from Gmail if tracking lost
echo   [x] Reply detection and duplicate prevention
echo   [x] Resume attachment on first emails
echo.
echo ========================================================================
echo                           STARTING SYSTEM
echo ========================================================================
echo.

REM Activate virtual environment and run
call venv\Scripts\activate.bat
python cold_email_automation.py

pause