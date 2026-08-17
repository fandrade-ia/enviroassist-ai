@echo off
cd /d "%~dp0"
call venv\Scripts\activate
streamlit run src\app_ieet.py
pause
