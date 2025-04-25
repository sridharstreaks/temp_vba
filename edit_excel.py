import os
import sys
glob
from datetime import datetime, timedelta
import pandas as pd

def get_recent_files(folder, minutes=5):
    now = datetime.now()
    pattern = os.path.join(folder, '*.csv')
    files = glob.glob(pattern)
    return [f for f in files if now - datetime.fromtimestamp(os.path.getmtime(f)) <= timedelta(minutes=minutes)]


def process_files(folder):
    files = get_recent_files(folder)
    if len(files) != 2:
        raise Exception(f"Expected exactly 2 CSV files updated in the last 5 minutes, found {len(files)}.")
    files.sort()
    first_file = files[0]

    # Read CSV into DataFrame
    df = pd.read_csv(first_file)

    # Insert new column 'Segment' at position 0
    df.insert(0, 'Segment', '')

    # Fill header and segment values
    df.at[0, 'Segment'] = 'Segment'
    # Identify rows that have any data in other columns
    mask = df.iloc[:, 1:].notna().any(axis=1)
    df.loc[mask, 'Segment'] = 'yes'

    # Save back to CSV
    df.to_csv(first_file, index=False)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python script.py <folder_path>")
        sys.exit(1)
    process_files(sys.argv[1])
    print("CSV processing completed.")
