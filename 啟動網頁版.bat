@echo off
rem 在這台電腦開報到網頁，同一個網路的其他電腦用瀏覽器開 http://本機IP:8501
cd /d %~dp0
python -m streamlit run checkin_web.py --server.address 0.0.0.0 --server.port 8501
pause
