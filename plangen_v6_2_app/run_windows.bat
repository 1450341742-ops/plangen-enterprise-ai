@echo off
cd /d %~dp0
if not exist .venv (
    python -m venv .venv
)
call .venv\Scripts\activate.bat
pip install -r requirements.txt
python -m streamlit run app.py
pause

REM 建议先执行：ollama pull qwen2.5:3b
