@echo off

cd /d "%~dp0"

echo ========================================
echo Current folder: %cd%
echo ========================================
echo.

echo Checking embedded Python...
"%~dp0python_embedded\python.exe" --version

echo.
echo Checking Streamlit...
"%~dp0python_embedded\python.exe" -m streamlit --version

echo.
echo Checking app.py...

if exist "%~dp0app.py" (
    echo app.py FOUND
) else (
    echo app.py NOT FOUND
)

echo.
echo ========================================
echo Launching LPU Monitor...
echo ========================================

"%~dp0python_embedded\python.exe" -m streamlit run "%~dp0app.py" --server.port=8501 --server.headless=false

echo.
echo ========================================
echo If you see an error above, that is the problem.
echo ========================================

pause