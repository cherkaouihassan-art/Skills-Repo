import openpyxl, sys, datetime, re, pickle
from collections import defaultdict
WB, EX, WRITE = sys.argv[1], sys.argv[2], len(sys.argv)>3 and sys.argv[3]=='--write'
RUN=datetime.datetime(2026,11,2)
def key(v):
    if v is None: return ''
    s=str(v).replace('\xa0','').strip()
    if re.fullmatch(r'\d+\.0', s): s=s[:-2]
    return s
wb=openpyxl.load_workbook(WB)
cfg={r[0]:r[1] for r in wb['Config'].iter_rows(min_row=2,values_only=True) if r[0]}
FROZEN=cfg['FrozenThrough']; EXP_HDR=cfg['ExtractLayout'].split('|')
OV={}
for r in wb['Overrides'].iter_rows(min_row=2,values_only=True):
    if r[0]: OV[key(r[0])]=dict(bu=r[1],cls=r[2],pm=r[3],scope=r[4])
MAP={}
for r in wb['Mapping'].iter_rows(min_row=2,values_only=True):
    if r[0]: MAP[key(r[0])]=dict(bu=r[1],it=r[2],t2=r[3],fl=r[4])
PREV={}
pk_hdr=[c.value for c in wb['Payments Key'][1]]
for r in wb['Payments Key'].iter_rows(min_row=2,values_only=True):
    d=dict(zip(pk_hdr,r)); PREV[(key(d['Pièce comptable']),key(d['Réf Paiement']),key(d['Réference']))]=d
# ---- extract
ex=openpyxl.load_workbook(EX, read_only=True, data_only=True)
ws=ex['JESA'] if 'JESA' in ex.sheetnames else ex.worksheets[0]
rows=ws.iter_rows(values_only=True); hdr=[str(h).strip() if h is not None else '' for h in next(rows)]
assert hdr[:35]==EXP_HDR, ('LAYOUT CHANGED', [(i,a,b) for i,(a,b) in enumerate(zip(hdr,EXP_HDR)) if a!=b])
paid=[]; maxd=None; nraw=0
for i,row in enumerate(rows, start=2):
    if row is None or all(v in (None,'') for v in row[:35]): continue
    nraw+=1
    try: z=float(row[25] or 0)
    except: z=0
    if z==0: continue
    d=row[29]
    if isinstance(d,datetime.datetime) and (maxd is None or d>maxd): maxd=d
    paid.append(dict(row=i, ref=key(row[8]), iddoc=key(row[2]), fourn=row[6], service=row[4], div=row[20], datedoc=row[9],
                     piece=key(row[28]), datepay=d, refpay=key(row[30]), net=z))
print(f'extract rows {nraw}, paid rows {len(paid)}, max Date Paiement {maxd:%d/%m/%Y}')
def classify(p):
    o=OV.get(p['ref']); m=MAP.get(p['ref'])
    bu=(o['bu'] if o and o['bu'] else None) or ((m['bu'] or '').strip() if m and m['bu'] else '')
    fl,t2,it=(m['fl'],m['t2'],m['it']) if m else (None,None,None)
    if o and o['cls']: cls=o['cls']
    elif not bu: cls=''
    elif fl=='PSR': cls='Proservices'
    elif fl=='PROC': cls='Procurement'
    elif t2=='Services': cls='Proservices'
    elif t2 not in (None,''): cls='Procurement'
    elif it=='Labor': cls='Proservices'
    else: cls='Procurement'
    d=p['datepay']; pm=datetime.datetime(d.year,d.month,26) if isinstance(d,datetime.datetime) else None
    applied=''
    if o and o['pm']=='EXCLUDE': pm='EXCLUDE'; applied='Pay Month EXCLUDE'
    elif o and isinstance(o['pm'],datetime.datetime): pm=o['pm']; applied='Pay Month forced'
    if o and (o['bu'] or o['cls']): applied=(applied+'; ' if applied else '')+'BU/Class'
    live=isinstance(pm,datetime.datetime) and pm>FROZEN
    return bu,cls,pm,fl,applied,live
for p in paid: p['bu'],p['cls'],p['pm'],p['fl'],p['applied'],p['live']=classify(p)
# ---- re-date sweep
redated=[]; seen=set()
for p in paid:
    k=(p['piece'],p['refpay'],p['ref']); seen.add(k)
    q=PREV.get(k)
    if q:
        if q['Pay Month']!=p['pm'] or abs(float(q['Net à payer'])-p['net'])>0.005:
            redated.append((p['ref'],p['piece'],p['refpay'],q['Date Paiement'],p['datepay'],q['Pay Month'],p['pm'],float(q['Net à payer']),p['net']))
gone=[q for k,q in PREV.items() if k not in seen and q['In live scope']=='Y']
print('re-dated / amount-changed payments:',len(redated)); [print('  ',r) for r in redated]
print('live payments in last week\'s key but missing now:',len(gone)); [print('  ',q['Réference'],q['Net à payer'],q['Date Paiement']) for q in gone]
# ---- unmapped
unm=[p for p in paid if p['live'] and not p['bu']]
print('unmapped live rows:',len(unm), round(sum(p['net'] for p in unm),2))
for p in unm: print('  ',p['row'],p['ref'],p['fourn'],p['div'],p['net'],p['datepay'].date(),p['pm'].date())
# ---- pivots
months=sorted({p['pm'] for p in paid if p['live']})
print('live months:',[m.date() for m in months])
piv=defaultdict(lambda:[0,0.0]); pivc=defaultdict(lambda:[0,0.0]); gross=defaultdict(lambda:[0,0.0])
for p in paid:
    d=p['datepay']
    if isinstance(d,datetime.datetime) and datetime.datetime(d.year,d.month,26)>FROZEN: gross[datetime.datetime(d.year,d.month,26)][0]+=1; gross[datetime.datetime(d.year,d.month,26)][1]+=p['net']
    if p['live'] and p['bu']:
        piv[(p['pm'],p['bu'])][0]+=1; piv[(p['pm'],p['bu'])][1]+=p['net']; pivc[(p['pm'],p['bu'],p['cls'])][0]+=1; pivc[(p['pm'],p['bu'],p['cls'])][1]+=p['net']
S=wb['Summary by BU']; bus=[S.cell(r,2).value for r in range(4,15)]
for m in months:
    t=[0,0.0]; ps=[0,0.0]; pr=[0,0.0]
    for bu in bus:
        t[0]+=piv[(m,bu)][0]; t[1]+=piv[(m,bu)][1]; ps[1]+=pivc[(m,bu,'Proservices')][1]; pr[1]+=pivc[(m,bu,'Procurement')][1]; ps[0]+=pivc[(m,bu,'Proservices')][0]; pr[0]+=pivc[(m,bu,'Procurement')][0]
    exc=[p for p in paid if isinstance(p['datepay'],datetime.datetime) and datetime.datetime(p['datepay'].year,p['datepay'].month,26)==m and p['pm']!=m]
    forced=[p for p in paid if p['pm']==m and isinstance(p['datepay'],datetime.datetime) and datetime.datetime(p['datepay'].year,p['datepay'].month,26)!=m]
    print(f"\n{m:%b-%y}: gross {gross[m][0]} / {gross[m][1]:,.2f}  excluded {[(p['ref'],p['net']) for p in exc]}  forced-in {[(p['ref'],p['net']) for p in forced]}  unmapped {sum(1 for p in unm if p['pm']==m)}")
    print(f"   mapped {t[0]} / {t[1]:,.2f} = PSR {ps[0]} / {ps[1]:,.2f} + PROC {pr[0]} / {pr[1]:,.2f}  (check {ps[1]+pr[1]:,.2f})")
    for bu in bus: print(f"   {bu:28} {piv[(m,bu)][0]:4} {piv[(m,bu)][1]:>16,.0f}   PSR {pivc[(m,bu,'Proservices')][1]:>14,.0f}  PROC {pivc[(m,bu,'Procurement')][1]:>14,.0f}")
pickle.dump(dict(paid=paid,redated=redated,gone=gone,unm=unm,months=months,piv=dict(piv),pivc=dict(pivc),gross=dict(gross),maxd=maxd,nraw=nraw),open('weekly_state.pkl','wb'))
