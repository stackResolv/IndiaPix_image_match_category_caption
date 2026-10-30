@echo off
REM ===========================================================================
REM run.bat - Windows one-command runner for the photo matching + categorizing
REM           pipeline. Mirrors run.py so the same commands work on Windows.
REM
REM Usage:
REM   run.bat                              (default shoot "_sort images")
REM   run.bat "E:\path\to\batch"           (or drag a folder onto this .bat)
REM   run.bat "input_folder"               (every shoot inside it, any depth)
REM   run.bat "MyShoot" "my_output"        (single shoot + custom output)
REM   run.bat "input_folder" "my_output"   (batch + custom output base)
REM   run.bat "input_folder" --provider openai                 (use OpenAI for captions)
REM   run.bat "input_folder" --provider openai --model gpt-4o  (override the model)
REM   run.bat "input_folder" --dry-run     (pair the sets, write report.csv, no API)
REM
REM Two input layouts are handled automatically - no flag needed:
REM   SHOOT      somewhere under the path there is a "JPEG" and/or "Edit" folder.
REM              Photos are matched to their originals by capture timestamp.
REM   SETS       photos but NO "JPEG"/"Edit" folders anywhere (one folder per
REM              set, e.g. ...\lense correction\INDIAPIX_BATCH_001\SET_000001\)
REM              Each photo's role is read from its own metadata instead.
REM
REM Delegates to run.py (all arguments forwarded verbatim via %*). Any nesting
REM depth is searched, so a deeply nested client delivery just works.
REM ===========================================================================

setlocal enableextensions

REM --- Always run from the folder this .bat lives in -------------------------
REM run.py resolves relative paths (e.g. "input_folder", "input_folder/JPEG")
REM against the current directory, so without this a run started from another
REM folder - or via a full path - would look for run.py and the shoot in the
REM wrong place. %~dp0 is this .bat's own directory, trailing backslash
REM included; /d also switches drive when the project is on another one.
cd /d "%~dp0"

REM --- Force the console to UTF-8 so non-ASCII filenames never crash ----------
chcp 65001 >nul 2>&1
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

REM --- Tell Python how many columns this window has --------------------------
REM A default cmd window is 80 wide. Without this, run.py would guess and the
REM step banners / explanations would wrap badly. Remove the next line to let
REM Python detect the width itself (it falls back to 80 if unsure).
if not defined PIPELINE_WIDTH set "PIPELINE_WIDTH=100"

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
REM %* keeps the user's quoting intact, so paths with spaces work whether the
REM call came from cmd.exe, PowerShell, or Explorer. The quoted script path is
REM safe even if the project folder itself contains spaces.
%PYTHON% "%~dp0run.py" %*
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