\
@echo off
cd /d "%~dp0"
call .venv\Scripts\activate
python ppt_youtube_agent.py --auth
pause
