"""Save a single Merchant cell selected in each output worksheet."""
from report_engine import OUTPUTS

def select_merchant_cells(book):
    for name in OUTPUTS:
        sheet=book.Worksheets(name);sheet.Visible=-1;sheet.Activate()
        sheet.Range('I5' if name.startswith('REGULAR') else 'F5').Select()
    book.Worksheets(OUTPUTS[0]).Activate()
