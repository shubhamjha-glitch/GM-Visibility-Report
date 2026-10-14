"""Refresh the original workbook through Microsoft Excel, without rebuilding it."""
from __future__ import annotations
import calendar
from collections import defaultdict
import datetime as dt
import json
import re
import shutil
import time
from pathlib import Path

INPUTS = ('GRC_DATA', 'PPO_DATA', 'GM_TNA', 'MNL_PPO')
OUTPUTS = ('REGULAR_DEPT', 'REGULAR_ART', 'WINTER_DEPT', 'WINTER_ART')
LAYOUTS = {
    'REGULAR_DEPT': dict(last=192, keys=['J','K','L'], grc='U', tna='V', ppo='AM', months=['AN','AO','AP'], manual='AQ'),
    'REGULAR_ART': dict(last=409, keys=['J','K','L','M'], grc='V', tna='W', ppo='AN', months=['AO','AP','AQ'], manual='AR'),
    'WINTER_DEPT': dict(last=30, keys=['G','H','I'], grc='N', tna='O', ppo='X', months=[], manual='Y'),
    'WINTER_ART': dict(last=81, keys=['G','H','I','J'], grc='O', tna='P', ppo='Y', months=[], manual='Z'),
}
SOURCE_KEYS = {'GRC_DATA':['E','F','G','H'], 'PPO_DATA':['A','M','N','O'], 'GM_TNA':['A','B','C','D'], 'MNL_PPO':['C','D','E','G']}
REQUIRED = {
    'GRC_DATA': ['ENTRY DATE','DIVISION','SECTION','DEPARTMENT','ARTICLE NAME','GRC QTY'],
    'PPO_DATA': ['DIVISION','SECTION','DEPARTMENT','ARTCLE NAME','PENDING_Q','DELIVERY DATE'],
    'GM_TNA': ['DIVISION','SECTION','DEPARTMENT','ARTICLE_NAME','TNA_Q'],
    'MNL_PPO': ['DIVISION','SECTION','DEPARTMENT','ARTICLE_NAME','TTL ORDER QTY'],
}

def norm(v):
    return re.sub(r'[^A-Z0-9®]', '', str(v or '').upper())

ALIASES = {
    'ARTCLENAME':'ARTICLENAME', 'ARTICLE_NAME':'ARTICLENAME',
    'PENDINGQTY':'PENDINGQ', 'PENDINGQUANTITY':'PENDINGQ',
    'SCHEDULEDATE':'SCHEDULEDT', 'SCHEDULEMONTH':'SCHEDULEMNTH',
    'SCHEDULEMNTH':'SCHEDULEMNTH', 'GRCDATE':'ENTRYDATE',
    'TNAQTY':'TNAQ', 'TOTALORDERQTY':'TTLORDERQTY',
    'DELIVERYDT':'DELIVERYDATE',
}

def key(v):
    n=norm(v)
    return ALIASES.get(n,n)

def col(n):
    s=''
    while n:
        n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def month_add(day, offset):
    n=day.year*12+day.month-1+offset
    return dt.date(n//12,n%12+1,1)

def serial(day):
    if isinstance(day,dt.datetime): day=day.date()
    return (day-dt.date(1899,12,30)).days

def parse_date(v):
    if isinstance(v,(dt.datetime,dt.date)): return serial(v)
    if isinstance(v,(int,float)): return float(v)
    if v is None or str(v).strip()=='':return None
    t=str(v).strip()
    for fmt in ('%Y-%m-%d','%Y-%m-%d %H:%M:%S','%d-%m-%Y','%d/%m/%Y','%d/%m/%y','%d-%b-%Y','%d-%b-%y','%b-%y','%b-%Y','%B-%Y','%b %Y','%B %Y'):
        try:return serial(dt.datetime.strptime(t,fmt))
        except ValueError:pass
    raise ValueError(f'Unrecognized date {v!r}. Use Excel dates or DD-MM-YYYY.')

def matrix(value):
    if value is None:return []
    if not isinstance(value,(tuple,list)):return [[value]]
    if value and not isinstance(value[0],(tuple,list)):return [list(value)]
    return [list(r) for r in value]

def align_rows(raw, canonical, name, tna_field='TNA_Q '):
    """Match columns by header, retain original positions, validate measures/dates."""
    header_row=None
    required=[key(h) for h in REQUIRED[name]]
    if name=='GM_TNA':required[-1]=key(tna_field)
    for ri,row in enumerate(raw[:30]):
        found={key(h) for h in row if h is not None}
        if name=='PPO_DATA' and 'SCHEDULEDT' in found:found.add('DELIVERYDATE')
        if set(required)<=found:header_row=ri;break
    if header_row is None:
        raise ValueError(f'{name}: required headers not found in first 30 rows: {", ".join(REQUIRED[name])}' + (f'. Selected TNA field: {tna_field}' if name=='GM_TNA' else ''))
    headers=raw[header_row]
    lookup={}
    for i,h in enumerate(headers):
        if h is None or not str(h).strip():continue
        k=key(h)
        if k in lookup:raise ValueError(f'{name}: duplicate header {h!r}.')
        lookup[k]=i
    expected=[key(h) for h in canonical]
    rows=[]; ignored=0
    qty={'GRC_DATA':'GRCQTY','PPO_DATA':'PENDINGQ','GM_TNA':key(tna_field),'MNL_PPO':'TTLORDERQTY'}[name]
    source_key_names=['DIVISION','SECTION','DEPARTMENT','ARTICLENAME']
    for ri,rawrow in enumerate(raw[header_row+1:],header_row+2):
        if not any(v is not None and str(v).strip() for v in rawrow):continue
        row=[rawrow[lookup[k]] if k in lookup and lookup[k]<len(rawrow) else None for k in expected]
        # Ignore formatted empty tails, including formula-only amount cells in MNL_PPO.
        if not any(row[expected.index(k)] for k in source_key_names if k in expected):ignored+=1;continue
        for k in source_key_names:
            if k in expected:
                ci=expected.index(k)
                if row[ci] is None or not str(row[ci]).strip():raise ValueError(f'{name} row {ri}: missing {k}.')
                row[ci]=str(row[ci]).strip()
        if qty in expected:
            qi=expected.index(qty);v=row[qi]
            if v is None or v=='':raise ValueError(f'{name} row {ri}: {qty} is blank. Enter 0 for a genuine zero.')
            if not isinstance(v,(int,float)):
                try:row[qi]=float(str(v).replace(',','').strip())
                except ValueError:raise ValueError(f'{name} row {ri}: non-numeric quantity {v!r}.')
        date_fields={'GRC_DATA':['ENTRYDATE','ORDERDATE','GATEENTRYDATE','GRCMONTH'], 'PPO_DATA':['ORDERDATE','SCHEDULEDT','SCHEDULEMNTH','DELIVERYDATE']}.get(name,[])
        for k in date_fields:
            if k in expected:row[expected.index(k)]=parse_date(row[expected.index(k)])
        if name=='GRC_DATA':
            d=row[expected.index('ENTRYDATE')]
            if d is None:raise ValueError(f'GRC_DATA row {ri}: ENTRY DATE is blank.')
            day=dt.date(1899,12,30)+dt.timedelta(days=int(d))
            row[expected.index('GRCMONTH')]=serial(day.replace(day=1))
        if name=='PPO_DATA':
            di=expected.index('SCHEDULEDT');mi=expected.index('SCHEDULEMNTH')
            d=row[di] if row[di] is not None else row[mi]
            if d is not None:
                day=dt.date(1899,12,30)+dt.timedelta(days=int(d));row[mi]=serial(day.replace(day=1))
            delivery=row[expected.index('DELIVERYDATE')]
            schedule=row[di]
            if delivery is not None and schedule is not None and int(delivery)!=int(schedule):raise ValueError(f'PPO_DATA row {ri}: SCHEDULE DT and DELIVERY DATE differ. They represent the same date; correct the mismatch or provide only SCHEDULE DT.')
            delivery=schedule if schedule is not None else delivery
            if delivery is None or delivery<=0:raise ValueError(f'PPO_DATA row {ri}: SCHEDULE DT (Delivery Date) is missing or invalid.')
            row[expected.index('DELIVERYDATE')]=delivery
        rows.append(row)
    if not rows and name!='MNL_PPO':raise ValueError(f'{name}: no data rows found. Upload at least one row with quantity 0 if this input has no activity.')
    return rows,ignored

def sumifs(source, measure, source_keys, target_keys, row, dates=()):
    args=[f"'{source}'!${measure}:${measure}"]
    for s,t in zip(source_keys,target_keys):args.extend([f"'{source}'!${s}:${s}",f'${t}{row}'])
    for datecol,criterion in dates:args.extend([f"'{source}'!${datecol}:${datecol}",criterion])
    return '=SUMIFS('+','.join(args)+')'

def formula_patches(name,last,tna_col):
    l=LAYOUTS[name];patches={}
    def block(dest,source,measure,dates=()):
        patches[dest]=tuple((sumifs(source,measure,SOURCE_KEYS[source],l['keys'],r,dates),) for r in range(5,last+1))
    block(l['grc'],'GRC_DATA','AF', [('A','">="&\'REPORT_CONTROL\'!$B$2'),('A','"<"&MIN(EDATE(\'REPORT_CONTROL\'!$B$2,1),\'REPORT_CONTROL\'!$B$3+1)')])
    block(l['tna'],'GM_TNA',tna_col)
    block(l['manual'],'MNL_PPO','AL')
    if l['months']:
        for i,dest in enumerate(l['months']):
            dates=[('AN','">0"'),('AN','"<"&EDATE(\'REPORT_CONTROL\'!$B$2,1)')] if i==0 else [('AN',f'">="&EDATE(\'REPORT_CONTROL\'!$B$2,{i})'),('AN',f'"<"&EDATE(\'REPORT_CONTROL\'!$B$2,{i+1})')]
            block(dest,'PPO_DATA','AM',dates)
        patches[l['ppo']]=tuple((f'=SUM({l["months"][0]}{r}:{l["months"][-1]}{r})',) for r in range(5,last+1))
    else:block(l['ppo'],'PPO_DATA','AM')
    return patches

def source_totals(rows, report_month, as_of, tna_col):
    """Independent Python totals to check Excel's live formula calculation."""
    totals=defaultdict(float)
    start=serial(report_month);end=min(serial(month_add(report_month,1)),serial(as_of)+1)
    indexes={'GRC_DATA':[4,5,6,7], 'PPO_DATA':[0,12,13,14], 'GM_TNA':[0,1,2,3], 'MNL_PPO':[2,3,4,6]}
    for source,records in rows.items():
        for r in records:
            for length in (3,4):
                k=tuple(str(r[i] or '').strip().upper() for i in indexes[source][:length])
                if source=='GRC_DATA' and start<=r[0]<end:totals[('grc',k)]+=r[31]
                elif source=='GM_TNA':totals[('tna',k)]+=r[23 if tna_col=='X' else 22]
                elif source=='MNL_PPO':totals[('manual',k)]+=r[37]
                elif source=='PPO_DATA':
                    totals[('winter_ppo',k)]+=r[38]
                    for i in range(3):
                        delivery=r[39]
                        in_month=(0<delivery<serial(month_add(report_month,1))) if i==0 else (serial(month_add(report_month,i))<=delivery<serial(month_add(report_month,i+1)))
                        if in_month:
                            totals[(f'ppo_{i}',k)]+=r[38];totals[('regular_ppo',k)]+=r[38]
    return totals

def layout_signature(sheet, last):
    return {
        'columns':[(sheet.Columns(i).ColumnWidth,sheet.Columns(i).Hidden,sheet.Columns(i).OutlineLevel) for i in range(1,sheet.UsedRange.Columns.Count+1)],
        'rows':[(sheet.Rows(i).RowHeight,sheet.Rows(i).Hidden,sheet.Rows(i).OutlineLevel) for i in range(1,last+2)],
        'conditional_formats':sheet.UsedRange.FormatConditions.Count,
    }

