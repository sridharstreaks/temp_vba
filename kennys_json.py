import pandas as pd
from openpyxl import load_workbook
import requests
import json

# ----- Step 1: Read the Excel file and extract URLs from column A -----
excel_file = input("Path of the excel file: ").replace("\"","")  # Path to your Excel file
df = pd.read_excel(excel_file)
isbns = df['isbn'].tolist()  # Get all URLs from the second column

# ----- Step 2: Ask for start and end row and create the selected URLs list -----
# Note: Excel rows are 1-indexed so we adjust for Python's 0-indexing.  
start_row = int(input("Enter start row number: "))
end_row = int(input("Enter end row number: "))
if start_row == 1:
    isbns = isbns[start_row-1 : end_row]
else:
    isbns = isbns[start_row-2 : end_row-1]  # Slice the list based on user input

# ----- Step 3: extract data from kennys site -----
extracted_data = {}

for isbn in isbns:
    response = requests.get(f"https://www.kennys.ie/searchquery?size=12&from=0&source=%7B%22query%22%3A%7B%22bool%22%3A%7B%22must%22%3A%5B%7B%22terms%22%3A%7B%22state%22%3A%5B1%2C2%5D%7D%7D%2C%7B%22terms%22%3A%7B%22cat_state%22%3A%5B1%2C2%5D%7D%7D%2C%7B%22bool%22%3A%7B%22should%22%3A%5B%7B%22match%22%3A%7B%22title%22%3A%7B%22query%22%3A%22{isbn}%22%2C%22operator%22%3A%22and%22%2C%22boost%22%3A1.7%7D%7D%7D%2C%7B%22match%22%3A%7B%22description%22%3A%7B%22query%22%3A%22{isbn}%22%2C%22operator%22%3A%22and%22%2C%22boost%22%3A0.7%7D%7D%7D%2C%7B%22match%22%3A%7B%22path%22%3A%7B%22query%22%3A%22{isbn}%22%2C%22operator%22%3A%22and%22%2C%22boost%22%3A2%7D%7D%7D%5D%7D%7D%2C%7B%22bool%22%3A%7B%22should%22%3A%5B%7B%22range%22%3A%7B%22end_date%22%3A%7B%22gte%22%3A%222025-03-09+11%3A27%3A54%22%7D%7D%7D%2C%7B%22term%22%3A%7B%22end_date%22%3A%220%22%7D%7D%2C%7B%22bool%22%3A%7B%22must_not%22%3A%5B%7B%22exists%22%3A%7B%22field%22%3A%22end_date%22%7D%7D%5D%7D%7D%5D%7D%7D%2C%7B%22bool%22%3A%7B%22should%22%3A%5B%7B%22range%22%3A%7B%22publish_end_date%22%3A%7B%22gte%22%3A%222025-03-09+11%3A27%3A54%22%7D%7D%7D%2C%7B%22term%22%3A%7B%22publish_end_date%22%3A%220%22%7D%7D%2C%7B%22bool%22%3A%7B%22must_not%22%3A%5B%7B%22exists%22%3A%7B%22field%22%3A%22publish_end_date%22%7D%7D%5D%7D%7D%5D%7D%7D%2C%7B%22bool%22%3A%7B%22should%22%3A%5B%7B%22range%22%3A%7B%22publish_start_date%22%3A%7B%22lte%22%3A%222025-03-09+11%3A27%3A54%22%7D%7D%7D%2C%7B%22term%22%3A%7B%22publish_start_date%22%3A%220%22%7D%7D%2C%7B%22bool%22%3A%7B%22must_not%22%3A%5B%7B%22exists%22%3A%7B%22field%22%3A%22publish_start_date%22%7D%7D%5D%7D%7D%5D%7D%7D%5D%2C%22should%22%3A%5B%5D%2C%22must_not%22%3A%5B%5D%2C%22filter%22%3A%5B%5D%7D%7D%2C%22suggest%22%3A%7B%22kennysdidyoumean%22%3A%7B%22text%22%3A%22{isbn}%22%2C%22phrase%22%3A%7B%22field%22%3A%22description%22%7D%7D%7D%7D&source_content_type=application%2Fjson&type=")
    if response.status_code==200:
        data = response.content
    else:
        continue
    
    # If data is a string, parse it into a dict.
    if isinstance(data, bytes):
        data = json.loads(data)

    if data.get("_shards", {}).get("successful") != 0:
        # Get id
        ISBN = data.get("suggest", {}).get("kennysdidyoumean", [])[0].get("text")

        if ISBN == isbn:
        
            # Navigate to the list of hits.
            hits_list = data.get("hits", {}).get("hits", [])
            if not hits_list:
                extracted_result = "No results found"
            else:
                # Get the first hit.
                first_hit = hits_list[0]
                source = first_hit.get("_source", {})
                
                # Extract 'path' and 'title' (renamed as 'text').
                path_value = source.get("path")
                
                extracted_result = "https://www.kennys.ie/"+path_value
        else:
            extracted_result = "isbn not match"
    else:
        extracted_result = "Failed to retrieve data"
    
    extracted_data[isbn] = extracted_result

    print(isbn,extracted_result)
    
# Load the workbook with openpyxl
wb = load_workbook(excel_file)
ws = wb.active  # Assumes the data is in the active sheet

# Write each result to column C starting at the specified start_row.
current_row = start_row
for isbn in isbns:
    ws.cell(row=current_row, column=3, value=extracted_data.get(isbn, ""))
    current_row += 1

wb.save(excel_file)
print("Results have been written to Excel.")
