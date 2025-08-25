@echo off
cd /d "%~dp0"
python Data2.py && python FNO.py && python Indices.py
pause