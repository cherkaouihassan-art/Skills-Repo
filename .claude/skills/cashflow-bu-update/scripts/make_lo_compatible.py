#!/usr/bin/env python3
"""Rewrite the formulas that feed the report so that LibreOffice (and any Excel) can evaluate them:
  XLOOKUP(key, lookup, return[, not_found]) -> INDEX/MATCH ; LET(...) -> inline expansion.
Applied only to the tabs that feed 'REPORT - Cashflow by BU' / 'EXEC Summary' (the dynamic-array view tabs are left as they are;
Excel recalculates them on open). Usage: make_lo_compatible.py <in.xlsx> <out.xlsx> [--sheets a,b,c]"""
import sys, re
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

def split_args(s):
    """split a function argument string on top-level commas (respects parentheses, brackets and quotes)"""
    out, depth, cur, q = [], 0, '', False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '"':
            q = not q
        elif not q:
            if ch in '([{': depth += 1
            elif ch in ')]}': depth -= 1
            elif ch == ',' and depth == 0:
                out.append(cur); cur = ''; i += 1; continue
        cur += ch; i += 1
    out.append(cur); return out

def find_calls(f, name):
    """yield (start, end, args) for every call of `name(` in formula text f (end = index after the closing paren)"""
    i = 0
    while True:
        j = f.find(name + '(', i)
        if j < 0: return
        k = j + len(name) + 1; depth = 1; q = False
        while k < len(f) and depth:
            ch = f[k]
            if ch == '"': q = not q
            elif not q:
                if ch == '(': depth += 1
                elif ch == ')': depth -= 1
            k += 1
        yield j, k, f[j + len(name) + 1:k - 1]
        i = j + 1

def convert_xlookup(f):
    while True:
        calls = list(find_calls(f, '_xlfn.XLOOKUP'))
        if not calls: return f
        # innermost first: pick the call whose args contain no other XLOOKUP
        for j, k, a in calls:
            if '_xlfn.XLOOKUP(' not in a:
                args = split_args(a)
                key, lookup, ret = args[0], args[1], args[2]
                core = f"INDEX({ret},MATCH({key},{lookup},0))"
                rep = f"IFERROR({core},{args[3]})" if len(args) > 3 and args[3].strip() != '' else core
                f = f[:j] + rep + f[k:]
                break
        else:
            return f

def expand_let(f):
    while True:
        calls = list(find_calls(f, '_xlfn.LET'))
        if not calls: return f
        for j, k, a in calls:
            if '_xlfn.LET(' not in a:
                args = split_args(a); body = args[-1]; pairs = list(zip(args[0:-1:2], args[1:-1:2]))
                defs = {}
                for name, expr in pairs:
                    e = expr
                    for n2, e2 in defs.items(): e = re.sub(re.escape(n2) + r'(?![A-Za-z0-9_.])', '(' + e2 + ')', e)
                    defs[name.strip()] = e
                b = body
                for n2, e2 in sorted(defs.items(), key=lambda x: -len(x[0])): b = re.sub(re.escape(n2) + r'(?![A-Za-z0-9_.])', '(' + e2 + ')', b)
                f = f[:j] + b + f[k:]
                break
        else:
            return f

def convert_tier_let(f):
    """compact rewrite of the ongoing-table tier formula (LET with nat / raw / t over the Tiering Master)"""
    m = re.search(r"_xlfn\.XLOOKUP\((TRIM\([^)]*\)\)?),'Tiering Master'!\$A\$5:\$A\$10000,'Tiering Master'!\$B\$5:\$B\$10000\)", f)
    d = re.search(r'_xlpm\.raw=""\),("Unclassified"|\d+),_xlpm\.raw\)', f)
    if not m or not d: return None
    key, DEF = m.group(1), d.group(1)
    M = f"MATCH({key},'Tiering Master'!$A$5:$A$10000,0)"
    C = f"INDEX('Tiering Master'!$C$5:$C$10000,{M})"; B = f"INDEX('Tiering Master'!$B$5:$B$10000,{M})"
    T = f'IFERROR(IF(OR({C}=0,{C}=""),{DEF},{C}),{DEF})'
    return f'=IF(AND(IFERROR({B},"")="International",ISNUMBER({T})),"Intl "&{T},{T})'

def convert(f):
    if '_xlpm.nat' in f and '_xlpm.raw' in f and '_xlpm.t,' in f:
        g = convert_tier_let(f)
        if g: return g
    g = expand_let(f); g = convert_xlookup(g)
    return g

if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    sheets = sys.argv[sys.argv.index('--sheets') + 1].split(',') if '--sheets' in sys.argv else ['REPORT - Cashflow by BU', 'DD', 'LC Ongoing (19-10)', 'September Actual Payments', 'REPORT - Oct Ongoing-Position', 'EXEC - BU & Integrity checks', 'EXEC Summary', 'SRC - LC Payments']
    wb = load_workbook(src); n = 0; maxlen = 0
    for s in sheets:
        if s not in wb.sheetnames: continue
        ws = wb[s]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula):
                    t = v.text
                    if '_xlfn.XLOOKUP' in t or '_xlfn.LET' in t:
                        nt = convert(t); c.value = ArrayFormula(v.ref, nt); n += 1; maxlen = max(maxlen, len(nt))
                elif isinstance(v, str) and v.startswith('=') and ('_xlfn.XLOOKUP' in v or '_xlfn.LET' in v):
                    nv = convert(v); c.value = nv; n += 1; maxlen = max(maxlen, len(nv))
    wb.save(dst); print(f'converted {n} formulas (longest {maxlen} chars) -> {dst}')
