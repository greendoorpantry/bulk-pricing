@echo off
REM ============================================================
REM  Bulk Pricing Auto-Update - Runs daily at 8am via Task Scheduler
REM ============================================================

cd /d C:\BulkPricing

REM Log output to file for debugging
echo. >> logs\daily_update.log
echo ======================================== >> logs\daily_update.log
echo %date% %time% - Starting bulk pricing update >> logs\daily_update.log
echo ======================================== >> logs\daily_update.log

REM Create logs folder if it doesn't exist
if not exist logs mkdir logs

REM Run the Python script
python generate_bulk_pricing_simple.py >> logs\daily_update.log 2>&1

echo %date% %time% - Finished (exit code: %errorlevel%) >> logs\daily_update.log
