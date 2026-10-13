#!/bin/bash
# run.sh — one-shot runner for the photo sorting + categorizing pipeline.
#
# Usage:
#   ./run.sh                          # defaults: shoot "_sort images", out "sorted_output"
#   ./run.sh "MyShoot" "my_output"    # custom shoot folder + output folder
#
# Runs both steps:
#   1. sort_photos.py  -> groups each edit with its matching original
#   2. categorize.py   -> adds category + caption, writes Excel

set -e

SHOOT="${1:-_sort images}"
OUT="${2:-sorted_output}"

echo "=================================================================="
echo " Step 1/2: matching edits to originals"
echo "   shoot : $SHOOT"
echo "   out   : $OUT"
echo "=================================================================="
python3 sort_photos.py --shoot "$SHOOT" --out "$OUT"

echo ""
echo "=================================================================="
echo " Step 2/2: categorizing + captioning edits"
echo "=================================================================="
python3 categorize.py \
  --edit-dir "$SHOOT/Edit" \
  --report "$OUT/report.csv" \
  --out "$OUT/photos.xlsx"

echo ""
echo "Done. Results:"
echo "  Excel   : $OUT/photos.xlsx"
echo "  Report  : $OUT/report.csv"
echo "  Groups  : $OUT/groups/"
echo "  Debug   : $OUT/categorize_debug.log"