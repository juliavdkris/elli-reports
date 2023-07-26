@echo off

IF EXIST ".venv" rmdir /s /q .venv
IF EXIST "out\*.docxd" del out\*.docx

python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python main.py

PAUSE
