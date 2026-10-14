"""Preserve report displays when conditional formatting changes cell colours."""
from copy import deepcopy
from openpyxl.formatting.formatting import ConditionalFormattingList

def keep_cell_number_formats(sheet):
    # The XLSB conversion introduced General into differential styles. Excel
    # applies that number format whenever a colour rule matches, masking 0%
    # and grouped quantities even though the ordinary cell style is correct.
    rules=[(str(cf.sqref),[deepcopy(rule) for rule in sheet.conditional_formatting[cf]]) for cf in sheet.conditional_formatting]
    sheet.conditional_formatting=ConditionalFormattingList()
    for address,items in rules:
        for rule in items:
            if rule.dxf is not None:rule.dxf.numFmt=None
            # A stale ID could still point to the original General override.
            rule.dxfId=None
            sheet.conditional_formatting.add(address,rule)

def format_report_numbers(sheet,last):
    for header in sheet[3]:
        text=str(header.value or '').upper()
        if '%' in text:format_code='0%'
        elif 'COVERAGE' in text and 'QTY' in text:format_code='#,##0'
        else:continue
        for row in (2,*range(5,last+2)):sheet.cell(row,header.column).number_format=format_code
    keep_cell_number_formats(sheet)
