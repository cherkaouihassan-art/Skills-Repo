#!/usr/bin/env bash
# Weekly run: put the 5 source files in a folder and call
#   ./run_week.sh <folder> <as-of YYYY-MM-DD> <ar-date YYYY-MM-DD> <prev-ar-date YYYY-MM-DD> [dso.json]
# Files are found by name pattern inside <folder> (case-insensitive):
#   AR_PRE*.xlsx            raw AR line list of the week (sheet 'AR')
#   CASH*flow*.xlsx         cash flow workbook ('EXEC Summary' + 'REPORT - Cashflow by BU'; tracking log / priority / LC ongoing optional)
#   AP*.xlsx | Supplier*    supplier invoice extract (sheet '4411 - Parked - Downpayment')
#   *LC*SITUATION*.xlsx     bank LC situation (sheet 'LC avec echéance')
#   UR*Tracking*.xlsx       unrecognized revenue consolidation
# The beginning-of-month AR file and its date are read from inputs/month_start.txt (two lines: path, date).
set -euo pipefail
cd "$(dirname "$0")"
SRC="$1"; ASOF="$2"; ARDATE="$3"; PREVDATE="$4"; DSO="${5:-}"
find1() { find "$SRC" -maxdepth 1 -iname "$1" | head -1; }
AR=$(find1 'AR_PRE*.xlsx'); CASH=$(find1 'CASH*flow*.xlsx'); AP=$(find1 'AP*.xlsx'); [ -z "$AP" ] && AP=$(find1 'Supplier*.xlsx')
LC=$(find1 '*LC*SITUATION*.xlsx'); UR=$(find1 'UR*Tracking*.xlsx')
for v in AR CASH AP LC UR; do [ -n "${!v}" ] || { echo "missing source file for $v in $SRC"; exit 1; }; done
MS_FILE=$(sed -n 1p inputs/month_start.txt); MS_DATE=$(sed -n 2p inputs/month_start.txt)
PREV=$(ls -t inputs/history/weekly_inputs_*.xlsx | head -1)
echo "AR=$AR"; echo "CASH=$CASH"; echo "AP=$AP"; echo "LC=$LC"; echo "UR=$UR"; echo "previous week=$PREV"; echo "month start=$MS_FILE ($MS_DATE)"
DSOARG=(); [ -n "$DSO" ] && DSOARG=(--dso "$DSO")
python3 -W ignore extract_inputs.py --ar "$AR" --prev-inputs "$PREV" --ar-baseline "$MS_FILE" --ar-baseline-date "$MS_DATE" \
  --cash "$CASH" --lc "$LC" --ap "$AP" --ur "$UR" "${DSOARG[@]}" \
  --as-of "$ASOF" --ar-date "$ARDATE" --ar-prev-date "$PREVDATE" --out inputs/weekly_inputs.xlsx
python3 build_deck.py --inputs inputs/weekly_inputs.xlsx --out "output/JESA_AR_Cash_Position_${ASOF}.pptx"
cp inputs/weekly_inputs.xlsx "inputs/history/weekly_inputs_${ASOF}.xlsx"
echo "done: output/JESA_AR_Cash_Position_${ASOF}.pptx"
