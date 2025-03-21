from openpyxl import load_workbook, Workbook

MAX_ROWS_PER_SHEET = 1_000_000  # Set a safe limit (Excel limit is ~1,048,576)

def write_pairs_to_multiple_sheets(wb, pairs, base_sheet_name="Output"):
    """
    Writes the list of pairs to multiple sheets in the workbook.
    Each sheet will have at most MAX_ROWS_PER_SHEET rows (excluding header).
    """
    sheet_count = 1
    current_sheet = wb.create_sheet(f"{base_sheet_name}_{sheet_count}")
    current_sheet.append(["Value 1", "Value 2"])
    row_count = 1  # header already written

    for pair in pairs:
        if row_count >= MAX_ROWS_PER_SHEET:
            sheet_count += 1
            current_sheet = wb.create_sheet(f"{base_sheet_name}_{sheet_count}")
            current_sheet.append(["Value 1", "Value 2"])
            row_count = 1
        current_sheet.append(list(pair))
        row_count += 1

def main():
    input_file = "input.xlsx"
    wb = load_workbook(input_file)
    ws = wb.active

    # Read the column values from ws (assuming column A)
    values = [cell.value for cell in ws["A"] if cell.value is not None]
    
    # Grouping logic (as defined previously)
    groups = []
    current_group = []
    current_wildcard = None

    def get_wildcard(text):
        return text.rsplit('-', 1)[0] if text and '-' in text else text

    for value in values:
        if not current_group:
            current_group.append(value)
            current_wildcard = get_wildcard(value)
        else:
            if current_wildcard in value:
                current_group.append(value)
            else:
                if len(current_group) >= 2:
                    groups.append(current_group)
                current_group = [value]
                current_wildcard = get_wildcard(value)
    if len(current_group) >= 2:
        groups.append(current_group)
    
    # Revised pairing: create ordered pairs from each group
    def create_pairs(group):
        pairs = []
        n = len(group)
        for i in range(n):
            for j in range(n):
                if i != j:
                    pairs.append((group[i], group[j]))
        return pairs

    all_pairs = []
    for group in groups:
        all_pairs.extend(create_pairs(group))

    # Create a new workbook to write output
    out_wb = Workbook()
    # Remove the default sheet created by Workbook
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    
    write_pairs_to_multiple_sheets(out_wb, all_pairs, base_sheet_name="Output")
    
    output_file = "output_multiple_sheets.xlsx"
    out_wb.save(output_file)
    print(f"Output written to {output_file}")

if __name__ == "__main__":
    main()
