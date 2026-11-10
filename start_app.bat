@echo off
title EBR High-Throughput Big-Data Server Engine (50 Lakh Rows Architecture)
color 0A
echo ================================================================
echo   EBR Big-Data Engine (50,00,000 Rows Streaming Pipeline)
echo ================================================================
echo Starting backend server on http://127.0.0.1:8000 ...
start "" "http://127.0.0.1:8000"
python server.py
pause
