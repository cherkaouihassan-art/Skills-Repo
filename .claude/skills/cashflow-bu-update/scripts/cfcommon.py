"""Shared constants and helpers for the CASH FLOW ANALYSIS - BUs workbook."""
import re, datetime as dt

BUS = ['AM', 'Advisory', 'Building & Capabilities', 'Building & Infrastructure', 'Energy Operations',
       'Fertilizers & Chemicals', 'Global Mining', 'JESA Technologies', 'SP2M Ad Hoc Program', 'Water & Environement']
BU_CODES = {'AMT': 'AM', 'ADS': 'Advisory', 'BIC': 'Building & Capabilities', 'BIN': 'Building & Infrastructure',
            'ENE': 'Energy Operations', 'FEC': 'Fertilizers & Chemicals', 'GMM': 'Global Mining', 'JET': 'JESA Technologies',
            'SP2M': 'SP2M Ad Hoc Program', 'SP2M-PAP': 'SP2M Ad Hoc Program', 'SP2M-OSBL': 'SP2M Ad Hoc Program',
            'PAP': 'SP2M Ad Hoc Program', 'MPH OSBL': 'SP2M Ad Hoc Program', 'WEN': 'Water & Environement'}
BU_ALIASES = {'water & environnement': 'Water & Environement', 'water': 'Water & Environement', 'w&e': 'Water & Environement',
              'f&c': 'Fertilizers & Chemicals', 'fertilizers': 'Fertilizers & Chemicals', 'mining': 'Global Mining',
              'gm': 'Global Mining', 'energy': 'Energy Operations', 'ene': 'Energy Operations', 'jet': 'JESA Technologies',
              'b&i': 'Building & Infrastructure', 'b&c': 'Building & Capabilities', 'sp2m': 'SP2M Ad Hoc Program',
              'am': 'AM', 'amt': 'AM', 'advisory': 'Advisory', 'ads': 'Advisory'}
TIERS = ['1', '2', '3', 'Intl 1', 'Intl 2', 'Intl 3', 'Etranger*', 'OCP et Filiale', 'Office/Publique']
UNCL_TIERS = ['Unclassified', 'Not on OCP Tier', 'Unclassified / Not on OCP', 'Etranger unclassified']
EXCLUDED_TIERS = ['Agency Local', 'Filiale']
BLOCK_ROWS = {'AM': 40, 'Advisory': 72, 'Building & Capabilities': 101, 'Building & Infrastructure': 130,
              'Energy Operations': 159, 'Fertilizers & Chemicals': 188, 'Global Mining': 217, 'JESA Technologies': 246,
              'SP2M Ad Hoc Program': 275, 'Water & Environement': 304}
LC_PROJECT = {'AM': 'AMT', 'Energy Operations': 'ENE', 'Fertilizers & Chemicals': 'FEC', 'JESA Technologies': 'JET',
              'SP2M Ad Hoc Program': 'SP2M', 'Water & Environement': 'WEN'}
# offsets from the block header row
OFF = dict(psr_rev=2, psr=3, less_psr=4, psr_res=5, net_psr=6, proc_rev=8, proc=9, div=10, uncl=11, cash_in=12,
           tier0=13, lc=23, customs=24, less_oh=25, less_pistelli=26, cash_out=27, net=28, oh=29)
MONTHS = [dt.date(2026, 7, 1), dt.date(2026, 8, 1), dt.date(2026, 9, 1), dt.date(2026, 10, 1)]
REPORT_COLS = {dt.date(2026, 7, 1): 'B', dt.date(2026, 8, 1): 'C', dt.date(2026, 9, 1): 'D', dt.date(2026, 10, 1): 'E'}

SHEETS = dict(exec='EXEC Summary', checks='EXEC - BU & Integrity checks', report='REPORT - Cashflow by BU',
              ongoing='REPORT - Oct Ongoing-Position', cashin='SRC - Cash In (AX Jul-Aug)',
              pay_jul='SRC - Payments Jul (Mapped)', pay_aug='SRC - Payments Aug (Mapped)',
              pay_sep='September Actual Payments', pay_oct='October Actual Payments', lc='SRC - LC Payments',
              lc_ongoing='LC Ongoing (19-10)', customs='DD', proc_comp='PROC Compensable (standing)',
              psr_res='PSR Reserve (provision)', sources='Data Sources', oh='OH Suppliers', tiering='Tiering Master')


def canon_bu(v):
    if v is None: return None
    s = str(v).strip()
    if s in BUS: return s
    if s.upper() in BU_CODES: return BU_CODES[s.upper()]
    return BU_ALIASES.get(s.lower(), s)


def norm_ref(v):
    """Reference key exactly as the workbook's helper formulas build it:
    UPPER(SUBSTITUTE(TRIM(x&""), " ", ""), "-", ""), "/", ""), ".", "")."""
    if v is None: return ''
    if isinstance(v, (dt.datetime, dt.date)): s = v.strftime('%Y%m%d')
    elif isinstance(v, float) and v.is_integer(): s = str(int(v))
    else: s = str(v)
    return s.strip().upper().replace(' ', '').replace('-', '').replace('/', '').replace('.', '')


def norm_supplier(v):
    """Supplier key as Excel COUNTIF/XLOOKUP compare it: trimmed, case-insensitive."""
    return str(v or '').strip().upper()


def tier_str(v):
    """Tier as the report writes it: 1/2/3 become '1','2','3'."""
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()


def to_date(v):
    if v is None or v == '': return None
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, (int, float)):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    s = str(v).strip()
    for f in ('%Y-%m-%d', '%d/%m/%Y', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S', '%d-%m-%Y'):
        try: return dt.datetime.strptime(s[:19], f).date()
        except ValueError: pass
    return None


def month_start(d):
    d = to_date(d)
    return dt.date(d.year, d.month, 1) if d else None


def num(v):
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace(' ', '').replace('\xa0', '').replace(' ', '').replace(',', '.')
    s = re.sub(r'\.(?=.*\.)', '', s)
    try: return float(s)
    except ValueError: return 0.0


def fmt(x):
    return f"{x:,.0f}" if isinstance(x, (int, float)) else str(x)
