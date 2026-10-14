@echo off
REM ===========================================================================
REM run.bat - Windows one-command runner for the photo matching + categorizing
REM           pipeline. Mirrors run.py so the same commands work on Windows.
REM
REM Usage (run from the project folder):
REM   run.bat                              (default shoot "_sort images")
REM   run.bat "input_folder"               (every shoot inside it)
REM   run.bat "MyShoot" "my_output"        (single shoot + custom output)
REM   run.bat "input_folder" "my_output"   (batch + custom output base)
REM   run.bat "input_folder" --provider openai                 (use OpenAI for captions)
REM   run.bat "input_folder" --provider openai --model gpt-4o  (override the model)
REM
REM Delegates to run.py (all arguments are forwarded verbatim via %*), which
REM detects whether the path is a single shoot or a parent folder of shoots,
REM then runs both steps: sort_photos.py + categorize.py.
REM ===========================================================================

setlocal enableextensions

REM --- Force the console to UTF-8 so non-ASCII filenames never crash ----------
chcp 65001 >nul 2>&1
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

REM --- Locate a Python interpreter (prefer the py launcher if present) --------
set "PYTHON="
where py >nul 2>&1 && set "PYTHON=py -3"
if not defined PYTHON (
    where python >nul 2>&1 && set "PYTHON=python"
)
if not defined PYTHON (
    echo ERROR: Python was not found on PATH.
    echo Install it from https://www.python.org/downloads/ and tick
    echo "Add Python to PATH", then reopen this window.
    goto :error
)

REM --- Run the pipeline (all arguments are forwarded as-is) -------------------
%PYTHON% run.py %*
if errorlevel 1 goto :error

echo.
echo Done.
exit /b 0

:error
echo.
echo ERROR: a step failed. See the messages above.
REM Keep the window open so the message is readable when double-clicked.
pause
exit /b 1