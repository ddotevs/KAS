@echo off
chcp 65001 >nul
cls
echo.
echo            =====                                                                        
echo         ===========                                                                     
echo      =======    ======                                                                  
echo   =====              ===                                                                
echo =====  ==  ==    =     ====   =====                                  ==                 
echo ====   =====    ===   ==  =   ==           ===                                          
echo  ===   ====    =====  =====   ===== ========== ===== =========== ====== ==== ===========
echo  ====  =====  =======    ===  ==    ==  == ============   ==  == ==  == ==== ====== ====
echo   ===  ==  =====   ========   ===== ==  == === ===== ==   ====== ==  ===================
echo   ====                  ==    =                           ==                            
echo    ===                 ===    =                                                         
echo    ====              ====     = == ====      =====================================      
echo     ====================      = == ====== =  ==============================             
echo     ===================                                                                 
echo.
echo.
echo ==================================================================================
echo                         LABEL SORTER - BUILD SETUP
echo ==================================================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/
    pause
    exit /b 1
)

echo [1/2] Installing required packages...
pip install pyinstaller tkinterdnd2 pillow pywin32

echo.
echo [2/2] Building EXE file...
echo.

REM Build the EXE with PyInstaller, including images and reference data
pyinstaller --onefile --windowed --name "KAS_Label_Sorter" --icon=images/favicon.ico --add-data "images/Logo.png;images" --add-data "images/favicon.ico;images" --add-data "reference_data.json;." kas_label_sorter.py

echo.
echo ==================================================================================
if exist "dist\KAS_Label_Sorter.exe" (
    echo                              BUILD SUCCESSFUL!
    echo ==================================================================================
    echo.
    echo   EXE Location: dist\KAS_Label_Sorter.exe
    echo.
    echo   Always remember, Stephen is a whore.
    echo.
) else (
    echo                           BUILD MAY HAVE FAILED
    echo ==================================================================================
    echo   Check the output above for errors.
    echo   If successful, look in the dist folder.
)
echo.
pause
