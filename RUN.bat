@echo off

IF EXIST "out\*.docx" del out\*.docx

python -m venv .venv
.venv\Scripts\pip install --upgrade -r requirements.txt
.venv\Scripts\python main.py

PAUSE
