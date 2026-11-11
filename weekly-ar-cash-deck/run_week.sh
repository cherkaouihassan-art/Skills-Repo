#!/usr/bin/env bash
# Weekly run: put the 5 source files in a folder and call
#   ./run_week.sh <folder> <as-of YYYY-MM-DD> <ar-date YYYY-MM-DD> <prev-ar-date YYYY-MM-DD>
# Files are found by name pattern inside <folder> (case-insensitive):
#   AR*.xlsx | AR*.zip | AR*.csv   raw AR line list of the week (sheet 'AR', or the sheet saved as CSV, optionally zipped)
#   CASH*flow*.xlsx         cash flow workbook ('EXEC Summary' + 'REPORT - Cashflow by BU'; tracking log / priority / LC ongoing optional)
#   AP*.xlsx | Supplier*    supplier invoice extract (sheet '4411 - Parked - Downpayment')
#   *LC*SITUATION*.xlsx     bank LC situation (sheet 'LC avec echéance')
#   UR*Tracking*.xlsx       unrecognized revenue consolidation
#   OCP*Payments*.xlsx      OCP payments weekly workbook ('Summary by BU'): monthly collections per OCP payment report (kept in inputs/ocp_payments/)
#   *DSO*.xlsx              DSO calculation workbook ('12 Months Trailing' revenue per project, 'Trend'); if absent, the
#                           latest inputs/dso_calc/*.xlsx is used (the revenue side only changes monthly)
# The beginning-of-month AR file and its date are read from inputs/month_start.txt (two lines: path, date).
set -euo pipefail
cd "$(dirname "$0")"
SRC="$1"; ASOF="$2"; ARDATE="$3"; PREVDATE="$4"
find1() { find "$SRC" -maxdepth 1 -iname "$1" | head -1; }
AR=$(find1 'AR_PRE*.xlsx'); [ -z "$AR" ] && AR=$(find1 'AR*.zip'); [ -z "$AR" ] && AR=$(find1 'AR*.csv'); [ -z "$AR" ] && AR=$(find1 'AR*.xlsx'); CASH=$(find1 'CASH*flow*.xlsx'); AP=$(find1 'AP*.xlsx'); [ -z "$AP" ] && AP=$(find1 'Supplier*.xlsx')
LC=$(find1 '*LC*SITUATION*.xlsx'); [ -z "$LC" ] && LC=$(find1 '*LC*.zip'); UR=$(find1 'UR*Tracking*.xlsx'); DSO=$(find1 '*DSO*.xlsx'); OCP=$(find1 'OCP*Payments*.xlsx'); SUSP=$(find1 '*Suspension*.xlsx')
mkdir -p inputs/suspension_log; if [ -n "$SUSP" ]; then cp "$SUSP" "inputs/suspension_log/Suspension_Log_${ASOF}.xlsx"; fi
SUSP=$(ls -t inputs/suspension_log/*.xlsx 2>/dev/null | head -1); SUSPARG=(); [ -n "$SUSP" ] && SUSPARG=(--suspension-log "$SUSP")
mkdir -p inputs/ocp_payments; if [ -n "$OCP" ]; then cp "$OCP" "inputs/ocp_payments/OCP_Payments_Weekly_${ASOF}.xlsx"; fi
OCP=$(ls -t inputs/ocp_payments/*.xlsx 2>/dev/null | head -1); OCPARG=(); [ -n "$OCP" ] && OCPARG=(--ocp-payments "$OCP")
# Cash flow workbook: the master maintained by the cash flow refresh kit in this repository (branch CASH_BRANCH),
# newest delivered output first, else the master; a CASH*.xlsx dropped in the source folder overrides it.
CASH_BRANCH="${CASH_BRANCH:-claude/laughing-wright-x02plq}"
if [ -z "$CASH" ]; then
  git fetch -q origin "$CASH_BRANCH" 2>/dev/null || true
  REF="origin/$CASH_BRANCH"; git rev-parse -q --verify "$REF" >/dev/null 2>&1 || REF="origin/main"
  CF_PATH=$(git ls-tree -r --name-only "$REF" -- cashflow-bu/outputs 2>/dev/null | grep -i '\.xlsx$' | sort | tail -1)
  [ -z "$CF_PATH" ] && CF_PATH="cashflow-bu/workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx"
  mkdir -p inputs/cash && CASH="inputs/cash/$(basename "$CF_PATH" | tr ' ' '_')"
  git show "$REF:$CF_PATH" > "$CASH" 2>/dev/null || { echo "cash flow workbook not found in $REF ($CF_PATH) and none in $SRC"; exit 1; }
  echo "cash flow workbook taken from $REF:$CF_PATH"
fi
for v in AR CASH AP LC UR; do [ -n "${!v}" ] || { echo "missing source file for $v in $SRC"; exit 1; }; done
mkdir -p inputs/dso_calc
# a weekly 'DSO Table' workbook (PU, 12 months trailing, AR, DSO, month revenue) is used as is; the full DSO calculation workbook is recomputed on the AR extract
if [ -n "$DSO" ]; then case "$(basename "$DSO" | tr 'A-Z' 'a-z')" in *table*) cp "$DSO" "inputs/dso_calc/DSO_Table_${ASOF}.xlsx";; *) cp "$DSO" "inputs/dso_calc/DSO_calc_${ASOF}.xlsx";; esac; fi
DSO=$(ls -t inputs/dso_calc/*.xlsx 2>/dev/null | head -1); [ -n "$DSO" ] || { echo "no DSO calculation workbook (inputs/dso_calc/ empty and none in $SRC)"; exit 1; }
case "$(basename "$DSO")" in DSO_Table*) DSOARG=(--dso-table "$DSO");; *) DSOARG=(--dso-calc "$DSO");; esac
MS_FILE=$(sed -n 1p inputs/month_start.txt); MS_DATE=$(sed -n 2p inputs/month_start.txt)
PREV=$(ls -t inputs/history/weekly_inputs_*.xlsx | head -1)
echo "AR=$AR"; echo "CASH=$CASH"; echo "AP=$AP"; echo "LC=$LC"; echo "UR=$UR"; echo "DSO calc=$DSO"; echo "previous week=$PREV"; echo "month start=$MS_FILE ($MS_DATE)"
python3 -W ignore extract_inputs.py --ar "$AR" --prev-inputs "$PREV" --ar-baseline "$MS_FILE" --ar-baseline-date "$MS_DATE" \
  --cash "$CASH" --lc "$LC" --ap "$AP" --ur "$UR" "${DSOARG[@]}" "${OCPARG[@]}" "${SUSPARG[@]}" --coll-month "$(python3 -c "import datetime as d; a=d.date.fromisoformat('$ASOF'); m=(a.replace(day=1)-d.timedelta(days=1)) if a.day<=7 else a; print(m.strftime('%Y-%m'))")" \
  --as-of "$ASOF" --ar-date "$ARDATE" --ar-prev-date "$PREVDATE" --out inputs/weekly_inputs.xlsx
WK=$(python3 -c "import datetime as d; a=d.date.fromisoformat('$ASOF'); print(a.strftime('%b')+'_W'+str((a.day-1)//7+1))")
OUT="output/JESA_AR_Cash_Position_${WK}_${ASOF}.pptx"
python3 build_deck.py --inputs inputs/weekly_inputs.xlsx --out "$OUT"
cp inputs/weekly_inputs.xlsx "inputs/history/weekly_inputs_${ASOF}.xlsx"
echo "done: $OUT"
