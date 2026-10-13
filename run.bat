@echo off
REM run.bat - Windows one-command runner for the photo pipeline.
REM
REM Usage:
REM   run.bat                          (defaults: shoot "_sort images", out "sorted_output")
REM   run.bat "MyShoot" "my_output"    (custom shoot folder + output folder)
REM
REM Runs both steps: sort_photos.py then categorize.py

setlocal
set SHOOT=%~1
if "%SHOOT%"=="" set SHOOT=_sort images
set OUT=%~2
if "%OUT%"=="" set OUT=sorted_output

echo ==================================================================
echo  Step 1/2: matching edits to originals
echo    shoot : %SHOOT%
echo    out   : %OUT%
echo ==================================================================
python sort_photos.py --shoot "%SHOOT%" --out "%OUT%"
if errorlevel 1 goto :error

echo.
echo ==================================================================
echo  Step 2/2: categorizing + captioning edits
echo ==================================================================
python categorize.py --edit-dir "%SHOOT%\Edit" --report "%OUT%\report.csv" --out "%OUT%\photos.xlsx"
if errorlevel 1 goto :error

echo.
echo Done. Results:
echo   Excel   : %OUT%\photos.xlsx
echo   Report  : %OUT%\report.csv
echo   Groups  : %OUT%\groups\
echo   Debug   : %OUT%\categorize_debug.log
goto :eof

:error
echo.
echo ERROR: a step failed. See the messages above.
exit /b 1