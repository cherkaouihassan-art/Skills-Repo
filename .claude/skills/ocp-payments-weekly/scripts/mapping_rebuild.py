import openpyxl, sys, re, datetime, pickle
from openpyxl.styles import Font, PatternFill, Alignment
AR, WB = sys.argv[1], sys.argv[2]
def key(v):
    if v is None: return ''
    s=str(v).replace('\xa0','').strip()
    if re.fullmatch(r'\d+\.0', s): s=s[:-2]
    return s
ar=openpyxl.load_workbook(AR, read_only=True, data_only=True)['AR']
rows=ar.iter_rows(values_only=True); hdr=[str(h).strip() if h is not None else '' for h in next(rows)]
ix={h:hdr.index(h) for h in ['Business Unit','Invoice Number','Invoice Type','Type 2','PSR/PROC','Payment Date']}
assert [openpyxl.utils.get_column_letter(ix[h]+1) for h in ['Business Unit','Invoice Number','Invoice Type','Type 2','PSR/PROC']]==['O','P','AL','AW','BH'], ix
new={}; dups=[]; n=0; maxpay=None
for r in rows:
    inv=key(r[ix['Invoice Number']])
    if not inv: continue
    n+=1
    rec=dict(bu=(r[ix['Business Unit']] or '').strip() or None, it=r[ix['Invoice Type']], t2=r[ix['Type 2']], fl=(r[ix['PSR/PROC']] or None))
    pd=r[ix['Payment Date']]
    if isinstance(pd,datetime.datetime) and (maxpay is None or pd>maxpay): maxpay=pd
    if inv in new:
        o=new[inv]
        if (rec['bu'] and o['bu']!=rec['bu']) or (rec['fl'] and o['fl']!=rec['fl']): dups.append((inv,o,rec))
    else: new[inv]=rec
print(f'AR weekly: {n} rows with invoice, {len(new)} unique, {len(dups)} duplicates with differing BU/flag, max Payment Date {maxpay:%d/%m/%Y}')
wb=openpyxl.load_workbook(WB); M=wb['Mapping']
old={}
for r in M.iter_rows(min_row=2,values_only=True):
    if r[0]: old[key(r[0])]=dict(bu=r[1],it=r[2],t2=r[3],fl=r[4])
carried=0; filled=0
for inv,o in old.items():
    if inv not in new: new[inv]=dict(o); new[inv]['carried']=True; carried+=1
    elif not new[inv]['fl'] and o['fl']: new[inv]['fl']=o['fl']; filled+=1
chg_bu=[(i,old[i]['bu'],new[i]['bu']) for i in old if i in new and (old[i]['bu'] or '')!=(new[i]['bu'] or '') and new[i]['bu']]
chg_fl=[(i,old[i]['fl'],new[i]['fl']) for i in old if i in new and (old[i]['fl'] or '')!=(new[i]['fl'] or '') and new[i]['fl']]
added=[i for i in new if i not in old]
print(f'carried from old mapping (not in AR weekly): {carried}; blank flags filled from old: {filled}; new invoices: {len(added)}; BU changes: {len(chg_bu)}; flag changes: {len(chg_fl)}')
st=pickle.load(open('weekly_state.pkl','rb')); live={p['ref']:p for p in st['paid'] if p['live']}
print('BU changes affecting live-month payments:'); [print('  ',c, live[c[0]]['net'], live[c[0]]['pm'].date()) for c in chg_bu if c[0] in live]
print('flag changes affecting live-month payments:'); [print('  ',c, live[c[0]]['net'], live[c[0]]['pm'].date()) for c in chg_fl if c[0] in live]
print('newly mapped live-month invoices:'); [print('  ',i,new[i]['bu'],new[i]['fl'],live[i]['net'],live[i]['pm'].date()) for i in added if i in live]
print('dups sample:'); [print('  ',d) for d in dups[:8]]
# write Mapping
M.delete_rows(2, M.max_row)
r=2
for inv in sorted(new, key=lambda x:(len(x),x)):
    v=new[inv]; M.cell(r,1,inv); M.cell(r,2,v['bu']); M.cell(r,3,v['it']); M.cell(r,4,v['t2']); M.cell(r,5,v['fl'])
    if v.get('carried'): M.cell(r,6,'carried from previous mapping (not in AR weekly 26/10/2026)')
    r+=1
M.cell(1,6,'Note'); M.cell(1,6).font=Font(bold=True,color='FFFFFF'); M.cell(1,6).fill=PatternFill('solid',fgColor='1F4E78'); M.column_dimensions['F'].width=48
C=wb['Config']
for row in C.iter_rows(min_row=2):
    if row[0].value=='MappingSource': row[1].value=f'AR weekly.xlsx (Power BI AR extract, AR tab, {n:,} rows with invoice, max Payment Date {maxpay:%d/%m/%Y}) + {carried} invoices carried from the previous mapping'; row[2].value=f'Rebuilt 02/11/2026. First occurrence kept where an invoice appears twice ({len(dups)} duplicates with differing BU/flag). Blank flags filled from the previous mapping ({filled}).'
D=wb['Data Sources']; k=D.max_row+1
for j,v in enumerate([f"Mapping!A2:F{r-1}","AR weekly.xlsx (Power BI AR extract uploaded by user 02/11/2026)",f"AR tab, cols P / O / AL / AW / BH; {n:,} rows with invoice number, max Payment Date {maxpay:%d/%m/%Y}","Weekly run by Claude (ocp-payments-weekly skill)","User: check the BU of the 12 unmapped October invoices in the new AR file; Mapping rebuilt",
    f"Deduped on invoice (first occurrence); {carried} invoices absent from the new AR file carried over from the previous mapping (col F note); {filled} blank flags filled from previous mapping","2026-11-02","User",r-2,
    f"{len(added)} new invoices; {len(chg_bu)} BU changes and {len(chg_fl)} PSR/PROC flag changes vs previous mapping (of which live-month impact: BU {sum(1 for c in chg_bu if c[0] in live)}, flag {sum(1 for c in chg_fl if c[0] in live)})",
    "All 12 previously unmapped Oct-26 invoices resolve from this file: 11 AM/PROC, 90022524 AM/PSR"],1): D.cell(k,j,v)
wb.save(WB); print('mapping written', r-2)
pickle.dump(dict(chg_bu=chg_bu,chg_fl=chg_fl,added=added,carried=carried,filled=filled,n=n,dups=dups,maxpay=maxpay),open('remap_state.pkl','wb'))
