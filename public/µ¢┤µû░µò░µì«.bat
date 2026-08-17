@echo off
cd /d "%~dp0"
chcp 65001 >nul

echo ================================
echo   HK IPO Cornerstone Investor DB
echo   Data Update Tool
echo ================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo   Download: https://www.python.org/downloads/
    echo   Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: Install python-pptx if needed
python -c "from pptx import Presentation" >nul 2>&1
if errorlevel 1 (
    echo [INSTALL] First run, installing python-pptx...
    pip install python-pptx
    echo.
)

:: Check PPT files
dir /b ppts\*.pptx >nul 2>&1
if errorlevel 1 (
    echo [HINT] No .pptx files found in ppts/ folder.
    echo   Please copy the new PPT into ppts/ folder first.
    pause
    exit /b 1
)

echo [1/4] Extracting PPT data...
python scripts/extract_pptx.py
if errorlevel 1 (
    echo [ERROR] PPT extraction failed. Check PPT format.
    pause
    exit /b 1
)

echo [2/4] Fetching company details...
python scripts/fetch_company_info.py
if errorlevel 1 (
    echo [WARNING] Company data fetch partially failed, using cache.
)

echo [3/4] Compiling IPO details...
python scripts\compile_ipo_details.py
if errorlevel 1 (
    echo [WARNING] IPO details compile failed, core function unaffected.
)

echo [4/4] Building standalone file...
python scripts\bundle_standalone.py
if errorlevel 1 (
    echo [WARNING] Standalone build failed, local use unaffected.
)

echo [5/5] Generating name cleaning checklist...
python scripts\generate_checklist.py
if errorlevel 1 (
    echo [WARNING] Checklist generation failed, skipping.
)

echo [DONE] All tasks completed!
echo.
echo   [Local use]: hk_cornerstone_investors.html
echo   [Share]:      hk_cornerstone_investors_standalone.html
echo.
echo Data updated! Open hk_cornerstone_investors.html to view.
echo.
echo To share with your manager, send:
echo   hk_cornerstone_investors_standalone.html
echo.
pause
