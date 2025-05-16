import os import glob import argparse import pandas as pd

Mapping for the two-letter codes after hyphen to head1 values

HEAD1_MAPPING = { 'AB': 1, 'CD': 2, # Add more mappings as needed }

Default value when there's no code suffix

DEFAULT_HEAD1 = 9

def extract_stats_and_series(df, col_idx): min_ = df.iat[3, col_idx] avg_ = df.iat[4, col_idx] max_ = df.iat[5, col_idx]

# Fill NaNs in series with 0
series = df.iloc[6:35, col_idx].fillna(0).tolist()
if len(series) != 29:
    raise ValueError(f"Expected 29 values in rows 7–35, got {len(series)}")

# Replace NaN for min/avg/max
min_ = 0 if pd.isna(min_) else min_
avg_ = 0 if pd.isna(avg_) else avg_
max_ = 0 if pd.isna(max_) else max_

return [min_, avg_, max_] + series

def get_dynamic_headers(df): labels = df.iloc[6:35, 0].tolist()  # Column A, rows 7-35 if len(labels) != 29: raise ValueError("Expected 29 labels in Column A (rows 7–35)") return labels

def process_name(folder, name): # name may include a suffix like 'john-AB' if '-' in name: base_name, code = name.split('-', 1) head1_val = HEAD1_MAPPING.get(code, DEFAULT_HEAD1) else: base_name = name head1_val = DEFAULT_HEAD1

base_path = os.path.join(folder, f"{name}.csv")
cap_path  = os.path.join(folder, f"{name}_capacity.csv")

df_base = pd.read_csv(base_path, header=None)
df_cap  = pd.read_csv(cap_path,  header=None)

dynamic_headers = get_dynamic_headers(df_base)
rows = []

# Helper to build each row
def build_row(tag, stats):
    row = {"head1": head1_val,
           "head2": base_name,
           "head3": tag,
           "head4": stats[0],
           "head5": stats[1],
           "head6": stats[2]}
    # Add dynamic series columns
    for i, value in enumerate(stats[3:], start=0):
        row[dynamic_headers[i]] = value
    return row

# "cac" => column B (index 1)
stats = extract_stats_and_series(df_base, col_idx=1)
rows.append(build_row("cac", stats))

# "fre" => column C (index 2)
stats = extract_stats_and_series(df_base, col_idx=2)
rows.append(build_row("fre", stats))

# "capacity" => column B of capacity file
stats = extract_stats_and_series(df_cap, col_idx=1)
rows.append(build_row("capacity", stats))

return rows, dynamic_headers

def main(folder, out_csv): pattern = os.path.join(folder, "*.csv") all_csvs = glob.glob(pattern) # Collect unique name stems (remove '_capacity' and extension) names = sorted({ os.path.splitext(os.path.basename(f))[0].replace("_capacity", "") for f in all_csvs if not f.endswith("_capacity.csv") })

all_rows = []
final_dyn_headers = []

for name in names:
    try:
        rows, dyn_headers = process_name(folder, name)
        all_rows.extend(rows)
        if not final_dyn_headers:
            final_dyn_headers = dyn_headers
    except Exception as e:
        print(f"Skipping '{name}' due to error: {e}")

# Construct DataFrame headers
static_headers = ["head1", "head2", "head3", "head4", "head5", "head6"]
full_headers = static_headers + final_dyn_headers

df_out = pd.DataFrame(all_rows, columns=full_headers)
df_out.to_csv(out_csv, index=False)
print(f"Aggregated data written to {out_csv} ({len(df_out)} rows).")

if name == "main": parser = argparse.ArgumentParser( description="Aggregate CSV pairs with dynamic headers and code-based head1 value." ) parser.add_argument("folder", help="Path to folder with CSV files") parser.add_argument( "--output", "-o", default="aggregated.csv", help="Output CSV filename (default: aggregated.csv)" ) args = parser.parse_args() main(args.folder, args.output)

