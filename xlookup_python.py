from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def insert_and_fill_values(
    path, sheet_name, start_row, interval, lookup_sheet_name=None
):
    wb = load_workbook(path)
    ws = wb[sheet_name]
    
    # 1) pick lookup sheet
    if lookup_sheet_name:
        ws_lookup = wb[lookup_sheet_name]
    else:
        idx = wb.sheetnames.index(sheet_name)
        ws_lookup = wb[wb.sheetnames[idx + 1]]
    
    # 2) build lookup dict: key = str(C)+str(A), value = E
    lookup = {}
    for row in ws_lookup.iter_rows(min_row=1,
                                   max_col=5,
                                   values_only=True):
        a, _, c, _, e = row[:5]   # row = (A, B, C, D, E)
        if a is not None and c is not None:
            lookup[str(c) + str(a)] = e
    
    # 3) find last row & last column in main sheet
    last_row = max(c.row for c in ws['A'] if c.value is not None)
    last_col = max(c.column for c in ws[1]  if c.value is not None)
    
    # 4) compute insertion points
    pts = [start_row + k * interval
           for k in range((last_row - start_row)//interval + 1)
           if start_row + k * interval <= last_row]
    
    # 5) do insert→copy→fill entirely in Python
    for r in reversed(pts):
        ws.insert_rows(r)
        # copy A:D from r+1 → r
        for col in range(1, 5):
            ws.cell(r, col).value = ws.cell(r+1, col).value
        # E = "blah"
        ws.cell(r, 5).value = "blah"
        # F–H left alone
        
        # now fill I→last_col by looking up C & header
        c_val = ws.cell(r, 3).value            # C<r>
        for col in range(9, last_col+1):       # 9 = column I
            hdr = ws.cell(1, col).value        # I1, J1, …
            key = str(c_val) + str(hdr)
            ws.cell(r, col).value = lookup.get(key, "")
    
    wb.save(path)
