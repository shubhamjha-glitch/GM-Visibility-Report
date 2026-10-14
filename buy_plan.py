"""Buy-plan schemas, strict key matching, live links and period captions."""
import calendar
import math
import re
from report_engine import key, matrix, col, month_add, LAYOUTS, sumifs

PLAN_SHEETS={'REGULAR_DEPT':'REGULAR_DEPT_BP','REGULAR_ART':'REGULAR_ART_BP','WINTER_DEPT':'WINTER_DEPT_BP','WINTER_ART':'WINTER_ART_BP'}
PLAN_COLUMNS={
 'REGULAR_DEPT':['Q','R','S','T'], 'REGULAR_ART':['R','S','T','U'],
 'WINTER_DEPT':['L','M'], 'WINTER_ART':['M','N'],
}

def plan_headers(name,month):
    keys=['DIVISION','SECTION','DEPARTMENT']
    if name.endswith('ART'):keys.append('ARTICLE NAME')
    periods=[month_add(month,i).strftime('%b_%Y').upper()+'_BUY_PLAN' for i in range(3)]
    return keys+(periods if name.startswith('REGULAR') else [periods[0],'WINTER_BUY_PLAN'])

def match_key(values):return tuple(str(v or '').strip().upper() for v in values)

def normalize_plan(raw,name,month,required_keys):
    headers=plan_headers(name,month);required={key(h) for h in headers}
    header_row=None
    for i,row in enumerate(raw[:30]):
        if required<={key(v) for v in row}:header_row=i;break
    if header_row is None:raise ValueError(f'{PLAN_SHEETS[name]}: expected headers: '+', '.join(headers)+'. Download the template for the selected reporting month.')
    indexes={}
    for i,h in enumerate(raw[header_row]):
        k=key(h)
        if not k:continue
        if k in indexes:raise ValueError(f'{PLAN_SHEETS[name]}: duplicate column {h!r}.')
        indexes[k]=i
    length=len(LAYOUTS[name]['keys']);records=[];seen=set();values={}
    for ri,row in enumerate(raw[header_row+1:],header_row+2):
        aligned=[row[indexes[key(h)]] if indexes[key(h)]<len(row) else None for h in headers]
        if not any(v is not None and str(v).strip() for v in aligned):continue
        k=match_key(aligned[:length])
        if not all(k):raise ValueError(f'{PLAN_SHEETS[name]} row {ri}: category keys are incomplete.')
        if k in seen:raise ValueError(f'{PLAN_SHEETS[name]} row {ri}: duplicate category '+ ' / '.join(k))
        seen.add(k);numbers=[]
        for h,v in zip(headers[length:],aligned[length:]):
            if v is None or str(v).strip()=='':raise ValueError(f'{PLAN_SHEETS[name]} row {ri}: {h} is blank. Enter the latest plan or 0.')
            try:n=float(str(v).replace(',','').strip())
            except (TypeError,ValueError):raise ValueError(f'{PLAN_SHEETS[name]} row {ri}: {h} must be numeric.')
            if not math.isfinite(n) or n<0:raise ValueError(f'{PLAN_SHEETS[name]} row {ri}: {h} must be a finite, non-negative quantity.')
            numbers.append(n)
        records.append([str(v).strip() for v in aligned[:length]]+numbers);values[k]=numbers
    return records,values,seen-set(required_keys)

def report_populations(input_rows, plans, original_keys):
    """Union input articles and plan articles; original listing is classification only."""
    articles={'REGULAR':set(plans.get('REGULAR_ART',{})), 'WINTER':set(plans.get('WINTER_ART',{}))}
    sources={'GRC_DATA':([4,5,6,7],8),'PPO_DATA':([0,12,13,14],11),'GM_TNA':([0,1,2,3],4),'MNL_PPO':([2,3,4,6],7)}
    classified={k:g for g in ('REGULAR','WINTER') for k in articles[g]}
    original={k:g for g in ('REGULAR','WINTER') for k in original_keys.get(g+'_ART',set())}
    evidence={};display={}
    for source,rows in input_rows.items():
        indexes,attribute=sources[source]
        for row in rows:
            k=match_key([row[i] for i in indexes]);display.setdefault(k,[str(row[i]).strip() for i in indexes])
            seasonal='WINTER' in str(row[attribute] or '').upper() or 'WINTER' in k[1]
            evidence[k]=evidence.get(k,False) or seasonal
    for k,seasonal in evidence.items():
        if k in articles['REGULAR'] or k in articles['WINTER']:continue
        group=original.get(k,'WINTER' if seasonal else 'REGULAR')
        articles[group].add(k)
    result={g+'_ART':articles[g] for g in articles}
    for g in articles:
        result[g+'_DEPT']={k[:3] for k in articles[g]}|set(plans.get(g+'_DEPT',{}))
    return result,display

def rewrite_population(sheet,name,wanted,display,old_last):
    """Retain matching template rows, append source-only categories with copied formulas/styles."""
    l=LAYOUTS[name];keycols=l['keys'];width=sheet.UsedRange.Column+sheet.UsedRange.Columns.Count-1
    old_rows=[];merchant={};merchant_col='I' if name.startswith('REGULAR') else 'F'
    for r in range(5,old_last+1):
        k=match_key([sheet.Range(f'{c}{r}').Value2 for c in keycols]);old_rows.append((r,k))
        merchant.setdefault(k[:3],sheet.Range(f'{merchant_col}{r}').Value2)
    # A temporary sheet retains a prototype even if every old report row is removed.
    prototype=sheet.Parent.Worksheets.Add(After=sheet.Parent.Worksheets(sheet.Parent.Worksheets.Count))
    rowheight=sheet.Rows(5).RowHeight
    sheet.Range(f'A5:{col(width)}5').Copy(Destination=prototype.Range(f'A5:{col(width)}5'))
    try:
        keep={k for _,k in old_rows if k in wanted}
        for r,k in reversed(old_rows):
            if k not in wanted:sheet.Rows(r).Delete()
        last=4+len(keep)
        for k in sorted(wanted-keep):
            r=last+1;sheet.Rows(r).Insert()
            prototype.Range(f'A5:{col(width)}5').Copy(Destination=sheet.Range(f'A{r}:{col(width)}{r}'))
            sheet.Rows(r).RowHeight=rowheight;sheet.Rows(r).Hidden=False
            # Retain prototype calculations; clear all inherited static metadata and quantities.
            for ci,v in enumerate(matrix(sheet.Range(f'A{r}:{col(width)}{r}').Formula)[0],1):
                if not (isinstance(v,str) and v.startswith('=')):sheet.Cells(r,ci).ClearContents()
            labels=display.get(k,list(k))
            for c,v in zip(keycols,labels):sheet.Range(f'{c}{r}').Value2=v
            sheet.Range(f'{merchant_col}{r}').Value2=merchant.get(k[:3])
            last=r
        # Normalize subtotal bounds after deleting/inserting rows, including initially empty ranges.
        for r in (2,last+1):
            for ci in range(1,width+1):
                cell=sheet.Cells(r,ci);f=cell.Formula
                if isinstance(f,str) and 'SUBTOTAL(' in f.upper():
                    letter=col(ci);cell.Formula=f'=SUBTOTAL(9,{letter}5:{letter}{last})' if last>=5 else '=0'
        return last
    finally:prototype.Delete()

def plan_patches(name,last):
    source=PLAN_SHEETS[name];length=len(LAYOUTS[name]['keys']);destinations=PLAN_COLUMNS[name]
    direct=destinations[:3] if name.startswith('REGULAR') else destinations
    patches={}
    for i,dest in enumerate(direct):
        measure=col(length+i+1)
        patches[dest]=tuple((sumifs(source,measure,[col(j+1) for j in range(length)],LAYOUTS[name]['keys'],r),) for r in range(5,last+1))
    if name.startswith('REGULAR'):
        patches[destinations[-1]]=tuple((f'=SUM({destinations[0]}{r}:{destinations[2]}{r})',) for r in range(5,last+1))
    return patches

def period_caption(value,month):
    """Relabel a reference or previously generated header without chained substitutions."""
    if not isinstance(value,str):return value
    # Names in the bundled September template, and names in an earlier generated template,
    # are replaced by semantic column positions in the engine. This helper handles combined labels.
    combined=''.join(month_add(month,i).strftime('%b')[0].upper() for i in range(3))
    return re.sub(r'\b[A-Z] & [A-Z]{2}\b|\b[A-Z]{3} / [A-Z]{3} / [A-Z]{3}\b|\b(?:JFM|FMA|MAM|AMJ|MJJ|JJA|JAS|ASO|SON|OND|NDJ|DJF)\b',combined,value)

def header_for_column(name,column,original,month):
    """Month captions follow column roles even when the custom template already says OCT/NOV/DEC."""
    if not isinstance(original,str):return original
    l=LAYOUTS[name];plan=PLAN_COLUMNS[name];month_index=None
    if name.startswith('REGULAR'):
        if column in plan[:3]:month_index=plan[:3].index(column)
        if column in l['months']:month_index=l['months'].index(column)
        if column in (['AE','AI','AK','G'] if name=='REGULAR_DEPT' else ['AF','AJ','AL','G']):month_index=1
        if column in (['AF','AJ','AL','H'] if name=='REGULAR_DEPT' else ['AG','AK','AM','H']):month_index=2
        if column in (['A','C','F','U','X','Z','AD','AH'] if name=='REGULAR_DEPT' else ['A','C','F','V','Y','AA','AE','AI']):month_index=0
    else:
        if column in (['A','C','L','N','Q','S'] if name=='WINTER_DEPT' else ['A','C','M','O','R','T']):month_index=0
    if month_index is not None:
        target=month_add(month,month_index).strftime('%b').upper()
        original=re.sub(r'\b(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b',target,original)
    return period_caption(original,month)
