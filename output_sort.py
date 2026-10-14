"""Sort only report detail rows with the user's specified priority levels."""
from report_engine import matrix, LAYOUTS
from value_paste import imported_columns

SORT_FIELDS={
 'REGULAR_DEPT':[('I',1),('J',1),('K',1),('T',2)],
 'REGULAR_ART':[('I',1),('J',1),('K',1),('L',1),('U',2)],
 'WINTER_DEPT':[('F',1),('G',1),('H',1),('M',2)],
 'WINTER_ART':[('F',1),('G',1),('H',1),('I',1),('N',2)],
}
def index(column):
    n=0
    for c in column:n=n*26+ord(c)-64
    return n-1

def snapshot(rows,name):
    fields=sorted(imported_columns(name)|{c for c,_ in SORT_FIELDS[name]})
    result={}
    for row in rows:
        k=tuple(str(row[index(c)] or '').strip().upper() for c in LAYOUTS[name]['keys'])
        result[k]=tuple(None if row[index(c)] in (None,'') else row[index(c)] for c in fields)
    return result

def sort_output_sheets(book,lastrows,progress=lambda s:None):
    for name,last in lastrows.items():
        if last<6:continue
        sheet=book.Worksheets(name)
        width=sheet.UsedRange.Column+sheet.UsedRange.Columns.Count-1
        block=sheet.Range(sheet.Cells(5,1),sheet.Cells(last,width))
        before=snapshot(matrix(block.Value2),name)
        progress(f'Sorting {name}: Merchant → Division → Section → '+('Department → ' if name.endswith('ART') else '')+'Buy Plan descending')
        sorter=sheet.Sort;sorter.SortFields.Clear()
        for column,order in SORT_FIELDS[name]:
            sorter.SortFields.Add(Key=sheet.Range(f'{column}5:{column}{last}'),SortOn=0,Order=order,DataOption=0)
        sorter.SetRange(block);sorter.Header=2;sorter.MatchCase=False;sorter.Orientation=1;sorter.Apply()
        sheet.Calculate()
        after=matrix(block.Value2)
        if snapshot(after,name)!=before:raise ValueError(f'{name}: sorting changed category-linked input data or buy-plan totals.')
        # Buy Plan is descending within each hierarchy group, not globally.
        prefix=[index(c) for c,_ in SORT_FIELDS[name][:-1]]
        plan=index(SORT_FIELDS[name][-1][0])
        for a,b in zip(after,after[1:]):
            same=all(str(a[c] or '').strip().casefold()==str(b[c] or '').strip().casefold() for c in prefix)
            if same and a[plan]<b[plan]:raise ValueError(f'{name}: Buy Plan descending sort within its hierarchy group failed.')
