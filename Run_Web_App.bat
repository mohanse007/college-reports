@echo off
echo ========================================================
echo  St. Ann's College for Women (A) - Report Generator Web
echo ========================================================
echo Starting local web portal at http://localhost:8000 ...
start "" "http://localhost:8000"
python -m http.server 8000
pause
