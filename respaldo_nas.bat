@echo off
cd /d "%~dp0"
if not exist logs mkdir logs
python respaldo_nas.py >> logs\respaldo_nas.log 2>&1
