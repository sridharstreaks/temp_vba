import os
import glob
import argparse
import pandas as pd
import string

def extract_stats_and_series(df, col_idx):
    """
    Given a DataFrame df, extract:
     - min, avg, max from rows 4,5,6 (1-based) => iloc[3], [4], [5]
     - 29 values from rows 7–35 => iloc[6:35]
    """
    # Note: pandas is 0-based
    min_ = df.iat[3, col_idx]
    avg_ = df.iat[4, col_idx]
    max_ = df.iat[5, col_idx]
    series = df.iloc[6:35, col_idx].tolist()
    if len(series) != 29:
        raise ValueError(f"Expected 29 values in rows 7–35, got {len(series)}")
    return [min_, avg_, max_] + series

def get_dynamic_headers(df):
    labels = df.iloc[6:35, 0].tolist()  # Column A, rows 7-35
    if len(labels) != 29:
        raise ValueError("Expected 29 labels in Column A (rows 7–35)")
    return labels

def process_name(folder, name):
    """
    For a given base name (without .csv or _capacity.csv), return three rows of data.
    """
    rows = []
    base_path = os.path.join(folder, f"{name}.csv")
    cap_path  = os.path.join(folder, f"{name}_capacity.csv")

    # Load CSVs into DataFrames (no header assumed)
    df_base = pd.read_csv(base_path, header=None)
    df_cap  = pd.read_csv(cap_path,  header=None)

    # 1) "cac" => column B (index 1)
    stats = extract_stats_and_series(df_base, col_idx=1)
    rows.append({
        **{"head1": "", "head2": name, "head3": "cac"},
        **{f"head{4+i}": stats[i] for i in range(len(stats))}
    })

    # 2) "fre" => column C (index 2)
    stats = extract_stats_and_series(df_base, col_idx=2)
    rows.append({
        **{"head1": "", "head2": name, "head3": "fre"},
        **{f"head{4+i}": stats[i] for i in range(len(stats))}
    })

    # 3) "capacity" => column B of capacity file
    stats = extract_stats_and_series(df_cap, col_idx=1)
    rows.append({
        **{"head1": "", "head2": name, "head3": "capacity"},
        **{f"head{4+i}": stats[i] for i in range(len(stats))}
    })

    return rows

def main(folder, out_csv):
    # Discover all base CSVs (skip _capacity files)
    pattern = os.path.join(folder, "*.csv")
    all_csvs = glob.glob(pattern)
    names = sorted({
        os.path.splitext(os.path.basename(f))[0]
        for f in all_csvs
        if not f.endswith("_capacity.csv")
    })

    # Prepare list of dict rows
    all_rows = []
    for name in names:
        try:
            rows = process_name(folder, name)
            all_rows.extend(rows)
        except Exception as e:
            print(f"Skipping '{name}' due to error: {e}")

    # Build DataFrame
    # Headers: head1...head35
    headers = [f"head{i}" for i in range(1,36)]
    df_out = pd.DataFrame(all_rows, columns=headers)

    # Write to CSV
    df_out.to_csv(out_csv, index=False)
    print(f"Wrote aggregated data to {out_csv} ({len(df_out)} rows).")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Aggregate pairs of CSV files in a folder into one master CSV."
    )
    parser.add_argument("folder", help="Path to the folder containing .csv and _capacity.csv files")
    parser.add_argument(
        "--output", "-o", default="aggregated.csv",
        help="Output CSV filename (default: aggregated.csv)"
    )
    args = parser.parse_args()

    main(args.folder, args.output)
