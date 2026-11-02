#!/usr/bin/env python3
"""Verify a loaded master without Excel: recalculate a COPY in LibreOffice (so helper formulas get values), then run
shadow_calc.py on that copy and compare with a baseline (state json). The copy is never delivered: LibreOffice 24.2
leaves #NAME? in XLOOKUP/LET cells, which shadow_calc.py treats as missing and re-derives.
Usage: verify.py <loaded_master.xlsx> --baseline cashflow-bu/state/refresh-state.json [--json out.json] [--keep]"""
import sys, os, shutil, subprocess, tempfile, json
here = os.path.dirname(os.path.abspath(__file__))
recalc = os.path.normpath(os.path.join(here, '..', '..', 'xlsx', 'scripts', 'recalc.py'))
src = sys.argv[1]; rest = sys.argv[2:]
tmp = tempfile.mkdtemp(prefix='cfverify_'); cp = os.path.join(tmp, 'recalc_copy.xlsx'); shutil.copy(src, cp)
print(f"Recalculating a copy in LibreOffice: {cp} (several minutes for this workbook) …", flush=True)
r = subprocess.run([sys.executable, recalc, cp, '900'], capture_output=True, text=True, cwd=os.path.dirname(recalc))
try:
    j = json.loads(r.stdout.strip().splitlines()[-1]); summ = j.get('error_summary', {})
    print("LibreOffice recalc:", j.get('status'), '| errors:', {k: v['count'] for k, v in summ.items()},
          "(#NAME? = XLOOKUP/LET not supported by LibreOffice — expected; #REF!/#VALUE! in legacy bridge tabs pre-exist)")
    bad = [loc for k in ('#REF!', '#VALUE!', '#DIV/0!') for loc in summ.get(k, {}).get('locations', [])
           if loc.split('!')[0].strip("'") in ('EXEC Summary', 'REPORT - Cashflow by BU', 'EXEC - BU & Integrity checks', 'REPORT - Oct Ongoing-Position')]
    if bad: print("!! formula errors on report tabs:", bad[:20])
except Exception:
    print(r.stdout[-2000:], r.stderr[-2000:])
rc = subprocess.call([sys.executable, os.path.join(here, 'shadow_calc.py'), cp] + [a for a in rest if a != '--keep'])
if '--keep' in rest: print('kept', cp)
else: shutil.rmtree(tmp, ignore_errors=True)
sys.exit(rc)
