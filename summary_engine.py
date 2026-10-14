"""Generate checked, value-pasted Regular/Winter summaries and Bottom 15 MCs."""
import datetime as dt
from pathlib import Path
from report_engine import matrix, col, LAYOUTS, month_add

SPEC={
 'REGULAR_DEPT':dict(merchant='I',division='J',bp='Q',total='T',coverage='AG',rate='Z',rank='P',dept='L'),
 'REGULAR_ART':dict(merchant='I',division='J',bp='R',total='U',coverage='AH',rate='Y',rank='Q',dept='L'),
 'WINTER_DEPT':dict(merchant='F',division='G',bp='L',total='M',coverage='W',rate='Q',rank='K',dept='I'),
 'WINTER_ART':dict(merchant='F',division='G',bp='M',total='N',coverage='X',rate='R',rank='L',dept='I'),
}

def sum_formula(name,column,last,criterion=None):
    if last<5:return '0'
    value=f"'{name}'!${column}$5:${column}${last}"
    if criterion is None:return 'SUM('+value+')'
    field,match=criterion
    return f'SUMIFS({value},\'{name}\'!${field}$5:${field}${last},{match})'

def summary_formulas(name,last,criterion=None):
    s=SPEC[name]
    numerators=['A','B','C','D','E'];denominators=[s['bp'],s['total'],s['bp'],s['total'],s['total']]
    result=[]
    for i,(n,d) in enumerate(zip(numerators,denominators)):
        ns=sum_formula(name,n,last,criterion);ds=sum_formula(name,d,last,criterion)
        if i==4 and name.startswith('REGULAR'):f=f'=IF({ds}>0,IFERROR({ns}/{ds},0),1)'
        else:f=f'=IFERROR({ns}/{ds},0)'
        result.append(f)
    result.append('='+sum_formula(name,s['coverage'],last,criterion))
    return result

def group_names(records,names,field):
    nameset=set()
    for name in names:
        i=column_index(SPEC[name][field])
        for row in records[name]:nameset.add(str(row[i] or '').strip())
    return sorted(nameset,key=lambda x:(not bool(x),x.upper()))

def column_index(text):
    n=0
    for ch in text:n=n*26+ord(ch)-64
    return n-1

def bottom_selection(records,limit=15):
    s=SPEC['REGULAR_DEPT']
    candidates=[]
    for r,row in enumerate(records,5):
        plan=row[column_index(s['bp'])]
        if not isinstance(plan,(int,float)) or plan<=0:continue
        rate=row[column_index(s['rate'])];rank=row[column_index(s['rank'])]
        if not isinstance(rank,(int,float)) or isinstance(rank,bool) or not 1<=rank<=100:continue
        if not isinstance(rate,(int,float)):continue
        candidates.append((float(rate),rank if isinstance(rank,(int,float)) else float('inf'),str(row[column_index(s['dept'])] or ''),r,row))
    chosen=sorted(candidates,key=lambda x:x[:4])[:limit]
    # Preserve the reference's buyer/division presentation after selecting lowest rates.
    return sorted(chosen,key=lambda x:(str(x[4][column_index(s['merchant'])] or ''),str(x[4][column_index(s['division'])] or ''),x[2]))

