import openpyxl, pickle, datetime, sys, os, copy, re
from openpyxl.utils import get_column_letter as L
st=pickle.load(open('weekly_state.pkl','rb')); paid,unm,months,piv,pivc,gross,maxd=st['paid'],st['unm'],st['months'],st['piv'],st['pivc'],st['gross'],st['maxd']
WB="../out/OCP Payments Weekly.xlsx"; EXTRACT_NAME=os.environ['EXTRACT_NAME']; RUN=datetime.datetime.strptime(os.environ['RUN_DATE'],'%Y-%m-%d')
wb=openpyxl.load_workbook(WB); S=wb['Summary by BU']; bus=[S.cell(r,2).value for r in range(4,15)]
def copy_style(src,dst):
    dst.font=copy.copy(src.font); dst.fill=copy.copy(src.fill); dst.border=copy.copy(src.border); dst.alignment=copy.copy(src.alignment); dst.number_format=src.number_format; dst.protection=copy.copy(src.protection)
def shift_formula(f, from_col, to_col):
    # shift column letters >= from_col by (to_col-from_col) in simple A1 refs of this sheet
    d=to_col-from_col
    def rep(m):
        col=m.group(2); row=m.group(4)
        ci=openpyxl.utils.column_index_from_string(col)
        if ci>=from_col: col=L(ci+d)
        return f"{m.group(1)}{col}{m.group(3)}{row}"
    return re.sub(r"(\$?)([A-Z]{1,2})(\$?)(\d+)", rep, f)
def month_cols(hdr_row, c0, c1):
    return {S.cell(hdr_row,c).value:c for c in range(c0,c1+1) if isinstance(S.cell(hdr_row,c).value,datetime.datetime)}
def add_month_simple(hdr_row, first_row, last_row, total_row, m):
    """Blocks A/C/D: months then Total(+%) columns. Insert m before the first non-date column after the months."""
    cols=month_cols(hdr_row,3,30); last=max(cols.values()); new=last+1
    # tail columns to move right by one: from new .. max used
    maxc=max(c for r in range(hdr_row,total_row+1) for c in range(1,31) if S.cell(r,c).value is not None)
    for r in range(hdr_row-2, total_row+1):
        for c in range(maxc, new-1, -1):
            src=S.cell(r,c); dst=S.cell(r,c+1)
            v=src.value
            if isinstance(v,str) and v.startswith('='): v=shift_formula(v,new,new+1)
            dst.value=v; copy_style(src,dst); src.value=None
        # new column: style from previous month
        copy_style(S.cell(r,last), S.cell(r,new))
    S.column_dimensions[L(new+1)].width=S.column_dimensions[L(maxc)].width or 14
    S.column_dimensions[L(new)].width=S.column_dimensions[L(last)].width or 14
    S.cell(hdr_row,new).value=m
    S.cell(total_row,new).value=f"=SUM({L(new)}{first_row}:{L(new)}{last_row})"
    # fix Total column formulas (now at new+1) to include the new month
    tc=new+1
    for r in range(first_row,last_row+1):
        v=S.cell(r,tc).value
        if isinstance(v,str) and v.startswith('=SUM('):
            S.cell(r,tc).value=re.sub(r"=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", lambda k:f"=SUM({k.group(1)}{k.group(2)}:{L(new)}{k.group(4)})", v)
    return new
def add_month_group(hdr_row, first_row, last_row, total_row, m, pct_label):
    """Blocks E/F: groups of 3 (count, amount, %) per month; append a new group after the last."""
    cols=month_cols(hdr_row,3,40); last_amt=max(cols.values()); nc,na,npct=last_amt+2,last_amt+3,last_amt+4
    for r in range(hdr_row-3, total_row+1):
        for off,src_c in ((nc,last_amt-1),(na,last_amt),(npct,last_amt+1)): copy_style(S.cell(r,src_c), S.cell(r,off))
    for off,src_c in ((nc,last_amt-1),(na,last_amt),(npct,last_amt+1)): S.column_dimensions[L(off)].width=S.column_dimensions[L(src_c)].width or 12
    S.cell(hdr_row,nc).value='# Invoices'; S.cell(hdr_row,na).value=m; S.cell(hdr_row,npct).value=pct_label
    for r in range(first_row,last_row+1): S.cell(r,npct).value=f"=IF({L(na)}${total_row}=0,0,{L(na)}{r}/{L(na)}${total_row})"
    S.cell(total_row,nc).value=f"=SUM({L(nc)}{first_row}:{L(nc)}{last_row})"; S.cell(total_row,na).value=f"=SUM({L(na)}{first_row}:{L(na)}{last_row})"; S.cell(total_row,npct).value=f"=IF({L(na)}{total_row}=0,0,{L(na)}{total_row}/{L(na)}{total_row})"
    return nc,na
added=[]
for m in months:
    if m not in month_cols(3,3,30):
        add_month_simple(3,4,14,16,m); add_month_simple(37,38,48,49,m); add_month_simple(53,54,64,65,m)
        add_month_group(70,71,81,82,m,'% of PSR'); add_month_group(87,88,98,99,m,'% of Proc.')
        for r in (68,85):
            v=S.cell(r,2).value
            if isinstance(v,str): S.cell(r,2).value=v.replace('October',m.strftime('%B')).replace('November 2026',m.strftime('%B %Y'))
        added.append(m)
colA=month_cols(3,3,30); colC=month_cols(37,3,30); colD=month_cols(53,3,30); colE=month_cols(70,3,40); colF=month_cols(87,3,40)
for m in months:
    for i,bu in enumerate(bus):
        S.cell(4+i,colA[m]).value=piv.get((m,bu),[0,0])[1]
        S.cell(38+i,colC[m]).value=pivc.get((m,bu,'Proservices'),[0,0])[1]
        S.cell(54+i,colD[m]).value=pivc.get((m,bu,'Procurement'),[0,0])[1]
        S.cell(71+i,colE[m]-1).value=pivc.get((m,bu,'Proservices'),[0,0])[0]; S.cell(71+i,colE[m]).value=pivc.get((m,bu,'Proservices'),[0,0])[1]
        S.cell(88+i,colF[m]-1).value=pivc.get((m,bu,'Procurement'),[0,0])[0]; S.cell(88+i,colF[m]).value=pivc.get((m,bu,'Procurement'),[0,0])[1]
# ---- Payments Key: deltas then rewrite
pk=wb['Payments Key']; hdr=[c.value for c in pk[1]]; prev={}
for r in pk.iter_rows(min_row=2,values_only=True):
    d=dict(zip(hdr,r)); prev[(str(d['Pièce comptable']),str(d['Réf Paiement']),str(d['Réference']))]=d
new_live=[p for p in paid if p['live'] and (p['piece'],p['refpay'],p['ref']) not in prev]
prev_tot={}
for d in prev.values():
    if d['In live scope']=='Y' and d['BU']: prev_tot.setdefault(d['Pay Month'],[0,0.0]); prev_tot[d['Pay Month']][0]+=1; prev_tot[d['Pay Month']][1]+=float(d['Net à payer'])
pk.delete_rows(2, pk.max_row)
for i,p in enumerate(paid,2):
    vals=[p['row'],p['ref'],p['iddoc'],p['fourn'],p['service'],p['div'],p['piece'],p['datepay'],p['refpay'],p['net'],p['pm'],p['bu'] or None,p['cls'] or None,p['fl'],p['applied'] or None,'Y' if p['live'] else 'N',maxd]
    for j,v in enumerate(vals,1):
        c=pk.cell(i,j,v)
        if j in (8,11,17) and isinstance(v,datetime.datetime): c.number_format='dd/mm/yyyy'
        if j==10: c.number_format='#,##0.00'
# ---- Unmapped
U=wb['Unmapped Invoices']; n=U.max_row
for p in unm:
    n+=1
    sug='AM / Procurement (pattern: JESA on JFC/MP1/JPH/107 DEF)' if (p['div'] or '').startswith(('JFC','MP1','JPH','107')) else 'no pattern - user to decide'
    vals=[p['row'],p['ref'],p['iddoc'],p['fourn'],p['service'],p['div'],p['datedoc'],p['datepay'],p['pm'],p['refpay'],p['net'],None,sug,RUN,None]
    for j,v in enumerate(vals,1):
        c=U.cell(n,j,v)
        if j in (7,8,9,14) and isinstance(v,datetime.datetime): c.number_format='dd/mm/yyyy'
        if j==11: c.number_format='#,##0.00'
mapped_refs={p['ref'] for p in paid if p['live'] and p['bu']}
for rr in range(2, U.max_row+1):
    if U.cell(rr,15).value is None and str(U.cell(rr,2).value) in mapped_refs: U.cell(rr,15).value=f"Resolved {RUN:%d/%m/%Y}"
# ---- Config / Data Sources
C=wb['Config']
for r in C.iter_rows(min_row=2):
    if r[0].value=='LastExtract': r[1].value=f"{EXTRACT_NAME} - extract {maxd:%d/%m/%Y} - {st['nraw']:,} rows - {len(paid):,} paid rows kept in Payments Key"
    if r[0].value=='LiveMonths': r[1].value=', '.join(m.strftime('%Y-%m-%d') for m in months)
def tot(m): return (sum(piv.get((m,b),[0,0])[0] for b in bus), sum(piv.get((m,b),[0,0])[1] for b in bus), sum(pivc.get((m,b,'Proservices'),[0,0])[1] for b in bus), sum(pivc.get((m,b,'Procurement'),[0,0])[1] for b in bus))
summ='; '.join(f"{m:%b-%y}: gross {gross[m][0]} / {gross[m][1]:,.2f}, {sum(1 for p in unm if p['pm']==m)} unmapped, mapped {tot(m)[0]} / {tot(m)[1]:,.2f} = PSR {tot(m)[2]:,.0f} + PROC {tot(m)[3]:,.0f}" for m in months)
D=wb['Data Sources']; k=D.max_row+1
row=["'Summary by BU' live columns "+', '.join(f"{m:%b-%y}={L(colA[m])}" for m in months)+(f"; new column(s) inserted for {', '.join(x.strftime('%b-%y') for x in added)}" if added else '')+f"; 'Payments Key'!A2:Q{len(paid)+1}; 'Unmapped Invoices' rows {n-len(unm)+1}-{n}" if unm else "'Summary by BU' live columns "+', '.join(f"{m:%b-%y}={L(colA[m])}" for m in months)+(f"; new column(s) inserted for {', '.join(x.strftime('%b-%y') for x in added)}" if added else '')+f"; 'Payments Key'!A2:Q{len(paid)+1}",
     EXTRACT_NAME,f"SAP supplier-payment extract, 35 cols A:AI, {st['nraw']:,} rows","Weekly run by Claude (ocp-payments-weekly skill)",os.environ.get('RUN_NOTE','Weekly refresh'),
     "Layout check OK; Overrides > Mapping > fallback; Pay Month = 26th; frozen months untouched",RUN.strftime('%Y-%m-%d'),"User",len(paid),
     f"Re-date sweep: {len(st['redated'])} moved payments, {len(st['gone'])} missing. "+summ, os.environ.get('RUN_TAIL','')]
for j,v in enumerate(row,1): D.cell(k,j,v)
wb.save(WB); print('saved; new month columns:',[m.date() for m in added])
print('prev live totals:',{(k.date() if hasattr(k,'date') else k):(v[0],round(v[1],2)) for k,v in prev_tot.items()})
print('new live payments this week:',len(new_live), round(sum(p['net'] for p in new_live),2))
for p in sorted(new_live,key=lambda p:-p['net']): print(f"   {p['ref']:>12} {p['datepay']:%d/%m} {p['net']:>15,.2f}  {p['bu'] or 'UNMAPPED':26} {p['cls']:12} {p['div']}")
