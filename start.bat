@echo off
cd /d "%~dp0"
echo Web検品ツールを起動します...
powershell -ExecutionPolicy Bypass -NoExit -Command "python -m pip install -r requirements.txt -q; Write-Host 'ブラウザが開きます (http://localhost:8501)'; python -m streamlit run app.py --server.port 8501"
