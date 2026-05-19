@echo off
cd /d "%~dp0"
set LOG_DIR=%~dp0logs
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
set LOG_FILE=%LOG_DIR%\job_reporter_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%date% %time%] Starting job reporter >> "%LOG_FILE%"
"%~dp0.venv\Scripts\python.exe" "%~dp0jsearch_job_reporter.py" >> "%LOG_FILE%" 2>&1
echo [%date% %time%] Finished with exit code %ERRORLEVEL% >> "%LOG_FILE%"
