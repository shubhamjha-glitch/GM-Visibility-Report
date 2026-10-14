"""Division, department and article master uploads with live keyed lookups."""
import math
from report_engine import key, col, matrix, LAYOUTS
SCHEMAS={
 'DIVISION_MERCHANT':['DIVISION','MERCHANT'],
 'ARTICLE_MASTER':['DIVISION','SECTION','DEPARTMENT','ARTICLE NAME','ATTRIBUTE','ART_STATUS',"STH% MAMJ'26","STH% JA'26"],
 'DEPARTMENT_MASTER':['DIVISION','SECTION','DEPARTMENT','ATTRIBUTE','SALE RANK',"STH% MAMJ'26","STH% JA'26"],
}
KEY_LENGTH={'DIVISION_MERCHANT':1,'ARTICLE_MASTER':4,'DEPARTMENT_MASTER':3}
MAPS={
 'REGULAR_DEPT':{'I':('DIVISION_MERCHANT','MERCHANT'),'M':('DEPARTMENT_MASTER','ATTRIBUTE'),'N':('DEPARTMENT_MASTER',"STH% MAMJ'26"),'O':('DEPARTMENT_MASTER',"STH% JA'26"),'P':('DEPARTMENT_MASTER','SALE RANK')},
 'REGULAR_ART':{'I':('DIVISION_MERCHANT','MERCHANT'),'N':('ARTICLE_MASTER','ART_STATUS'),'O':('ARTICLE_MASTER',"STH% MAMJ'26"),'P':('ARTICLE_MASTER',"STH% JA'26"),'Q':('DEPARTMENT_MASTER','SALE RANK')},
 'WINTER_DEPT':{'F':('DIVISION_MERCHANT','MERCHANT'),'J':('DEPARTMENT_MASTER','ATTRIBUTE'),'K':('DEPARTMENT_MASTER','SALE RANK')},
 'WINTER_ART':{'F':('DIVISION_MERCHANT','MERCHANT'),'K':('ARTICLE_MASTER','ART_STATUS'),'L':('DEPARTMENT_MASTER','SALE RANK')},
}

def normalize_metadata(raw,name):
    headers=SCHEMAS[name];length=KEY_LENGTH[name];expected={key(h) for h in headers}
    if name=='ARTICLE_MASTER':expected.discard(key('ATTRIBUTE'))
    start=next((i for i,row in enumerate(raw[:30]) if expected<={key(v) for v in row}),None)
    if start is None:raise ValueError(f'{name}: required headers are '+', '.join(headers))
    indexes={}
    for i,h in enumerate(raw[start]):
        k=key(h)
        if not k:continue
        if k in indexes:raise ValueError(f'{name}: duplicate header {h}.')
        indexes[k]=i
    seen=set();result=[]
    for ri,row in enumerate(raw[start+1:],start+2):
        v=[row[indexes[key(h)]] if key(h) in indexes and indexes[key(h)]<len(row) else None for h in headers]
        if not any(x is not None and str(x).strip() for x in v):continue
        keys=tuple(str(x or '').strip().upper() for x in v[:length])
        if not all(keys):raise ValueError(f'{name} row {ri}: incomplete matching key.')
        if keys in seen:raise ValueError(f'{name} row {ri}: duplicate matching key: '+ ' / '.join(keys))
        seen.add(keys)
        v[:length]=[str(x).strip() for x in v[:length]]
        for i,h in enumerate(headers[length:],length):
            value=v[i]
            if value is None or str(value).strip()=='':v[i]=None;continue
            if h.startswith('STH%') or h=='SALE RANK':
                try:
                    t=str(value).strip().replace(',','');n=float(t[:-1])/100 if t.endswith('%') else float(t)
                except ValueError:
                    hint='use a whole-number rank or leave the cell blank' if h=='SALE RANK' else 'use 12.5% or 0.125 for a percentage'
                    raise ValueError(f'{name} row {ri}: {h} must be numeric ({hint}).')
                if not math.isfinite(n) or n<0 or (h=='SALE RANK' and not n.is_integer()):raise ValueError(f'{name} row {ri}: invalid {h}.')
                v[i]=n
            else:v[i]=str(value).strip()
        result.append(v+['\x1f'.join(keys)])
    return result,seen

def lookup_formula(name,dest,row):
    source,field=MAPS[name][dest];columns=LAYOUTS[name]['keys'][:KEY_LENGTH[source]]
    match='&CHAR(31)&'.join(f'${c}{row}' for c in columns)
    value_col=col(SCHEMAS[source].index(field)+1);key_col=col(len(SCHEMAS[source])+1)
    value=f"INDEX('{source}'!${value_col}:${value_col},MATCH({match},'{source}'!${key_col}:${key_col},0))"
    return f'=IFERROR(IF({value}="","",{value}),"")'

def metadata_patches(name,last):
    return {c:tuple((lookup_formula(name,c,r),) for r in range(5,last+1)) for c in MAPS[name]}

def configure_headers(book,lastrows):
    # Existing department Preference field is now explicitly the uploaded Attribute.
    for name,column in [('REGULAR_DEPT','M'),('WINTER_DEPT','J')]:
        sheet=book.Worksheets(name)
        if str(sheet.Range(column+'3').Value2 or '').strip().upper()=='PREFERENCE':sheet.Range(f'{column}5:{column}{lastrows[name]}').ClearContents()
        sheet.Range(column+'3').Value2='ATTRIBUTE'
    # Remove the optional columns added by previous portal versions if an old
    # generated workbook is used as the reference. Original report columns stay fixed.
    removable={
      'REGULAR_ART':[('AU','ATTRIBUTE')],
      'WINTER_DEPT':[('AC',"STH% JA'26"),('AB',"STH% MAMJ'26")],
      'WINTER_ART':[('AE',"STH% JA'26"),('AD',"STH% MAMJ'26"),('AC','ATTRIBUTE')],
    }
    for name,fields in removable.items():
        sheet=book.Worksheets(name)
        for column,label in fields:
            if key(sheet.Range(column+'3').Value2)==key(label):sheet.Columns(column).Delete()
