#!/bin/bash
# =============================================================================
# run.sh - macOS / Linux one-command runner for the photo matching + cataloging
#          pipeline. Mirrors run.bat so the same commands work everywhere.
#
# Usage:
#   ./run.sh                              # default shoot "_sort images"
#   ./run.sh "input_folder"               # every shoot inside it
#   ./run.sh "MyShoot" "my_output"        # single shoot + custom output
#   ./run.sh "input_folder" "my_output"   # batch + custom output base
#
# Delegates to run.py, which detects whether the path is a single shoot or a
# parent folder of shoots, then runs both steps: sort_photos.py + categorize.py.
# =============================================================================

set -e

# Force UTF-8 for the console and for any child Python processes.
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8

# Let the caller pin the wrap width when piping into a log viewer that reports
# no width (e.g. `./run.sh ... | tee run.log`). Leave unset for auto-detect.
if [ -t 1 ] && [ -z "${PIPELINE_WIDTH:-}" ] && command -v tput >/dev/null 2>&1; then
    cols=$(tput cols 2>/dev/null || echo "")
    case "$cols" in
        ''|*[!0-9]*) : ;;
        *) [ "$cols" -ge 40 ] && export PIPELINE_WIDTH="$cols" ;;
    esac
fi

# Prefer python3, fall back to python.
if command -v python3 >/dev/null 2>&1; then
    PY=python3
else
    PY=python
fi

exec "$PY" run.py "$@"