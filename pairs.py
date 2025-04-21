import pandas as pd

def filter_and_delete_rows(input_file, output_file, column_name, values_to_remove):
    # Load the Excel file
    df = pd.read_excel(input_file)

    # Filter the DataFrame to exclude rows with specific values in the target column
    filtered_df = df[~df[column_name].isin(values_to_remove)]

    # Save the result to a new Excel file
    filtered_df.to_excel(output_file, index=False)
    print(f"Filtered file saved as: {output_file}")

# Example usage
input_file = 'input_data.xlsx'
output_file = 'filtered_output.xlsx'
column_name = 'Status'  # Change this to your target column
values_to_remove = ['Inactive', 'Pending']  # List of values to filter out

filter_and_delete_rows(input_file, output_file, column_name, values_to_remove)
