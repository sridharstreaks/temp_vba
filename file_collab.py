import os
import glob
import argparse
import pandas as pd

def extract_stats_and_series(df, col_idx):
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
    base_path = os.path.join(folder, f"{name}.csv")
    cap_path  = os.path.join(folder, f"{name}_capacity.csv")

    df_base = pd.read_csv(base_path, header=None)
    df_cap  = pd.read_csv(cap_path,  header=None)

    dynamic_headers = get_dynamic_headers(df_base)
    rows = []

    # "cac"
    stats = extract_stats_and_series(df_base, col_idx=1)
    rows.append({
        **{"head1": "", "head2": name, "head3": "cac"},
        **{dynamic_headers[i]: stats[3+i] for i in range(29)},
        "head4": stats[0], "head5": stats[1], "head6": stats[2]
    })

    # "fre"
    stats = extract_stats_and_series(df_base, col_idx=2)
    rows.append({
        **{"head1": "", "head2": name, "head3": "fre"},
        **{dynamic_headers[i]: stats[3+i] for i in range(29)},
        "head4": stats[0], "head5": stats[1], "head6": stats[2]
    })

    # "capacity"
    stats = extract_stats_and_series(df_cap, col_idx=1)
    rows.append({
        **{"head1": "", "head2": name, "head3": "capacity"},
        **{dynamic_headers[i]: stats[3+i] for i in range(29)},
        "head4": stats[0], "head5": stats[1], "head6": stats[2]
    })

    return rows, dynamic_headers

def main(folder, out_csv):
    pattern = os.path.join(folder, "*.csv")
    all_csvs = glob.glob(pattern)
    names = sorted({
        os.path.splitext(os.path.basename(f))[0].replace("_capacity", "")
        for f in all_csvs if not f.endswith("_capacity.csv")
    })

    all_rows = []
    final_headers = []

    for name in names:
        try:
            rows, dynamic_headers = process_name(folder, name)
            all_rows.extend(rows)
            if not final_headers:
                final_headers = dynamic_headers
        except Exception as e:
            print(f"Skipping '{name}': {e}")

    # Construct final headers
    static_headers = ["head1", "head2", "head3", "head4", "head5", "head6"]
    full_headers = static_headers + final_headers

    df_out = pd.DataFrame(all_rows, columns=full_headers)
    df_out.to_csv(out_csv, index=False)
    print(f"Aggregated file written to: {out_csv}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("folder", help="Folder containing CSV files")
    parser.add_argument("--output", "-o", default="aggregated.csv", help="Output CSV file name")
    args = parser.parse_args()
    main(args.folder, args.output)
