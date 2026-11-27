#!/usr/bin/env python3
"""Produce a deliverable with calculated values: make the report formulas LibreOffice-compatible, recalculate in LibreOffice,
re-enable Excel's full recalculation on open (for the dynamic-array view tabs LibreOffice cannot evaluate), then read back the
integrity checks and the headline from the calculated values.
Usage: deliver.py <master.xlsx> <out.xlsx>"""
import sys, os, subprocess, zipfile, shutil, re, json, tempfile
here = os.path.dirname(os.path.abspath(__file__)); src, out = sys.argv[1], sys.argv[2]
tmp = tempfile.mkdtemp(prefix='cfdeliver_'); conv = os.path.join(tmp, 'conv.xlsx')
subprocess.check_call([sys.executable, os.path.join(here, 'make_lo_compatible.py'), src, conv])
recalc = os.path.normpath(os.path.join(here, '..', '..', 'xlsx', 'scripts', 'recalc.py'))
print('LibreOffice recalculation (several minutes) …', flush=True)
r = subprocess.run([sys.executable, recalc, conv, '1500'], capture_output=True, text=True, cwd=os.path.dirname(recalc))
try:
    j = json.loads(r.stdout.strip().splitlines()[-1]); summ = j.get('error_summary', {})
    print('LibreOffice:', j.get('status'), {k: v['count'] for k, v in summ.items()})
    bad = [loc for k in summ for loc in summ[k]['locations'] if loc.split('!')[0].strip("'") in ('EXEC Summary', 'REPORT - Cashflow by BU', 'EXEC - BU & Integrity checks')]
    print('errors on report tabs:', bad[:40] if bad else 'none')
except Exception: print(r.stdout[-1500:], r.stderr[-1500:])
# re-enable full calc on load (LibreOffice drops the flag) by patching xl/workbook.xml inside the zip
zin = zipfile.ZipFile(conv); zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'xl/workbook.xml':
        x = data.decode('utf-8')
        if '<calcPr' in x: x = re.sub(r'<calcPr([^>]*?)/>', lambda m: '<calcPr' + re.sub(r'\sfullCalcOnLoad="[^"]*"', '', m.group(1)) + ' fullCalcOnLoad="1"/>', x, 1)
        else: x = x.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
        data = x.encode('utf-8')
    zout.writestr(item, data)
zout.close(); zin.close(); shutil.rmtree(tmp, ignore_errors=True)
print('written', out, round(os.path.getsize(out) / 1e6, 1), 'MB')
subprocess.call([sys.executable, os.path.join(here, 'snapshot.py'), out])
