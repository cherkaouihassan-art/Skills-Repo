#!/usr/bin/env python3
"""Verify a loaded master without Excel.
1. shadow_calc.py recomputes the report from the source tabs (formula helper columns re-derived) and compares the
   headline with the baseline state json -> the expected movement.
2. Optionally (--lo-scan) a COPY is recalculated in LibreOffice only to scan the report tabs for #REF!/#VALUE!/#DIV/0!.
   LibreOffice 24.2 cannot evaluate XLOOKUP/LET and mis-evaluates some helper columns, so its numbers are never used.
Usage: verify.py <loaded_master.xlsx> --baseline cashflow-bu/state/refresh-state.json [--json out.json] [--lo-scan]"""
import sys, os, shutil, subprocess, tempfile, json
here = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1]; rest = [a for a in sys.argv[2:] if a != '--lo-scan']
rc = subprocess.call([sys.executable, os.path.join(here, 'shadow_calc.py'), src] + rest)
if '--lo-scan' in sys.argv:
    recalc = os.path.normpath(os.path.join(here, '..', '..', 'xlsx', 'scripts', 'recalc.py'))
    tmp = tempfile.mkdtemp(prefix='cfverify_'); cp = os.path.join(tmp, 'recalc_copy.xlsx'); shutil.copy(src, cp)
    print(f"\nLibreOffice formula-error scan on a copy (several minutes) …", flush=True)
    r = subprocess.run([sys.executable, recalc, cp, '900'], capture_output=True, text=True, cwd=os.path.dirname(recalc))
    try:
        j = json.loads(r.stdout.strip().splitlines()[-1]); summ = j.get('error_summary', {})
        report_tabs = ('EXEC Summary', 'REPORT - Cashflow by BU', 'EXEC - BU & Integrity checks', 'REPORT - Oct Ongoing-Position', 'DD', 'SRC - LC Payments', 'October Actual Payments')
        bad = [loc for k in ('#REF!', '#VALUE!', '#DIV/0!') for loc in summ.get(k, {}).get('locations', []) if loc.split('!')[0].strip("'") in report_tabs]
        print("LibreOffice scan:", {k: v['count'] for k, v in summ.items()}, "(#NAME? = XLOOKUP/LET unsupported by LibreOffice: expected; #REF!/#VALUE! in legacy bridge tabs pre-exist)")
        print("Formula errors on report/source tabs:", bad[:30] if bad else 'none')
    except Exception:
        print(r.stdout[-1500:], r.stderr[-1500:])
    shutil.rmtree(tmp, ignore_errors=True)
sys.exit(rc)
