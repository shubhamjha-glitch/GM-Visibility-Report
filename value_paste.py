"""Freeze imported report fields while retaining business calculation formulas."""
from report_engine import LAYOUTS
from buy_plan import PLAN_COLUMNS
from metadata import MAPS

def imported_columns(name):
    l=LAYOUTS[name]
    fields={l['grc'],l['tna'],l['manual'],*l['months'],*MAPS[name]}
    fields.add(l['ppo'])
    plan=PLAN_COLUMNS[name]
    fields.update(plan)
    return fields

def freeze_imported_fields(book,lastrows):
    count=0
    for name,last in lastrows.items():
        if last<5:continue
        sheet=book.Worksheets(name)
        for column in sorted(imported_columns(name)):
            cells=sheet.Range(f'{column}5:{column}{last}')
            values=cells.Value2
            cells.Value2=values
            count+=last-4
            formulas=cells.Formula
            matrix=formulas if isinstance(formulas,(tuple,list)) else ((formulas,),)
            for row in matrix:
                for value in row if isinstance(row,(tuple,list)) else (row,):
                    if isinstance(value,str) and value.startswith('='):raise ValueError(f'{name}!{column}: imported value was stored as a formula.')
    return count
