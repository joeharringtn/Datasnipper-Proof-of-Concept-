#!/bin/bash
# run.sh — Run the workbook generator, log all output, and push to GitHub.
#
# Usage (from repo root):
#   bash scripts/run.sh
#   bash scripts/run.sh --source-pdf "my.pdf" --terms "Joe" --pages "1-3"
#
# All arguments are forwarded to generate_workbook.py.
# Output is printed to the terminal AND appended to errors/run_log.txt,
# which is committed and pushed so it's visible from any machine.

LOG_FILE="errors/run_log.txt"
DIVIDER="============================================================"
TIMESTAMP=$(date '+%Y-%m-%d %H:%M:%S')

# Write run header to log
{
  echo ""
  echo "$DIVIDER"
  echo "[$TIMESTAMP]"
  echo "Command: python scripts/generate_workbook.py $*"
  echo "$DIVIDER"
} >> "$LOG_FILE"

# Run the script — output goes to terminal AND log file simultaneously
python scripts/generate_workbook.py "$@" 2>&1 | tee -a "$LOG_FILE"

# Capture exit code before doing anything else
EXIT_CODE=${PIPESTATUS[0]}

echo "" >> "$LOG_FILE"
echo "Exit code: $EXIT_CODE" >> "$LOG_FILE"

# Push log + any new output files to GitHub so Claude can read them
git add "$LOG_FILE"
git add output/*.xlsx 2>/dev/null || true
git commit -m "log: run $TIMESTAMP"
git push

echo ""
echo "Log and output files pushed. Claude can now read errors/run_log.txt and output/."
