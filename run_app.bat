@echo off
echo 🚀 Starting File Converter App...
echo.

REM Activate virtual environment
call venv\Scripts\activate

REM Check if requirements are installed
python -c "import streamlit" 2>nul
if errorlevel 1 (
    echo 📦 Installing requirements...
    pip install -r requirements.txt
)

REM Start the application
echo 🌐 Starting Streamlit server...
echo 📋 The app will open in your default browser
echo 🛑 Press Ctrl+C to stop the server
echo.
streamlit run app.py

pause