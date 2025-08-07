#!/usr/bin/env python3
"""
Blank-grid solver with

  • exact row counts
  • exact column counts
  • exact totals in “… F” columns and “… en 100 g” columns
  •     PLUS four group-level quotas that remove the remaining ambiguity

The first grid satisfying everything is written to blank_grid.xlsx
"""

import itertools
from collections import defaultdict
import pandas as pd
import re
import numpy as np
import xlsxwriter as xlsxwriter
import random

# --- Step 0: Get the raw text ---
# The first line will be the file name, 6.8.2 Carnes importadas for example
# First to lines will be excluded
# Headers should be between lines 'Componente alimentario' and 'Nutriente Tagname Unidad...'

raw_text = """
6.9.2 Pescados
Rm-Pes-1 Rm-Pes-2 Rm-Pes-3 Rm-Pes-4 Rm-Pes-5 Rm-Pes-6 Rm-Pes-7
Componente alimentario 
Bagre, chihuil 
Besugo
Bonito, bacoreta, cachorra 
Boquerón crudo 
Boquerón cocido
Boquilla, ronco 
Cabrilla
Nutriente Tagname Unidad F En 100 g F En 100 g F En 100 g F En 100 g F En 100 g F En 100 g F En 100 g
Elementos principales
Energía ENERC kcal 103 103 162 114 190 80 81
kJ 433 433 677 477 795 334 340
Humedad WATER % 79.31 77.23 1 67.60 1 71.10 1 58.60 1 79.80 1 79.60
Fibra dietética FIBTG mg 1 0.00 0.00 1 0.00 1 0.00 1 0.00 1 0.00 1 0.00
Hidratos de C CHOCDF mg 1 0.00 1 0.00 1 0.00 R 0.30 1 0.00 1 0.00
Proteínas PROCNT mg 15.91 17.95 1 24.00 1 17.70 R 24.00 1 19.50 1 20.10
Lípidos tot FAT mg 1 2.70 1 7.30 1 4.80 1 10.30 1 0.20 1 0.10
Ác. grasos
Saturados FASAT mg 1 0.60 1 - 1 1.30 1 - 1 - 1 -
Monoinsaturados FAMS mg 1 1.00 1 - 1 1.20 1 - 1 - 1 -
Poliinsaturados FAPU mg 1 0.80 1 - 1 1.60 1 - 1 - 1 -
Linolénico F18D3N3 mg 1 0.10 1 - 1 - 1 - 1 - 1 -
Eicosapentaenoico F20D5N3 mg 12.34 6.80 1 - 1 0.50 1 - 1 - 1 -
Docosahexaenoico F22D6N3 mg 39.02 19.18 1 - 1 0.90 1 - 1 - 1 -
Colesterol CHOLE mg 20.27 9.70 1 - 1 - 1 - 1 - 1 -
minerales
Calcio CA mg 1 32.00 1 26.00 R - 1 168.00 1 10.00 1 15.00
Fósforo P mg 2 194.00 2 258.00 R - 2 - 2 167.00 2 183.00
Hierro FE mg 1 0.40 1 0.70 1 1.20 1 - R - 1 0.00
Magnesio MG mg 1 - 1 - 1 - 1 - 1 - 1 -
Sodio NA mg 1 60.00 1 40.00 1 - 1 - 1 - 1 -
Potasio K mg 1 330.00 1 293.00 1 - 1 - 1 - 1 -
Zinc ZN mg 1 0.80 1 0.30 1 0.50 1 1.80 1 - 1 -
Vitaminas
RAE (vit A) VITA µg 1 - 1 - 1 0.00 1 0.00 1 - 1 -
Ác. ascórbico ASCL mg 1 0.00 1 0.00 1 0.00 1 0.00 1 0.00 1 0.00
Tiamina THIA mg 1 0.04 1 0.02 1 0.01 1 - 1 0.03 1 0.05
Riboflavina RIBF mg 1 0.03 1 0.05 1 0.15 1 - 1 0.03 1 0.05
Niacina NIA mg 1 1.70 1 12.80 1 1.90 1 - 1 2.20 1 1.10
Piridoxina VITB6A mg 1 - 1 - 1 0.48 1 - 1 - 1 -
Ác. fólico FOL µg 1 - 1 - 1 8.00 1 - 1 - 1 -
Cianocobalamina VITB12 µg 1 - 1 - 1 28.00 1 - 1 - 1 -
Alimento crudo en peso neto P. comestible 50% P. comestible % P. comestible 51% P. comestible 80% P. comestible 100% P. comestible 51% P. comestible 51%
"""

# --- Step 0a: Extract file name from the first line and remove the first two lines ---
all_lines = raw_text.strip().splitlines()
# First line will be used for naming the file.
file_name_line = all_lines[0].strip()
# Keep all lines except the first two and the last one.
data_lines = all_lines[2:-1]
# Rebuild raw_text from remaining lines.
raw_text = "\n".join(data_lines)

# --- Step 0b: Replace specific words with their correct forms ---
replacements = {
    "Fibra dietética": "Fibra_dietética",
    "Hidratos de C": "Hidratos_de_C",
    "Lípidos tot": "Lípidos_tot",
    "RAE (vit A)": "RAE_(vit_A)",
    "Ác. ascórbico": "Ác._ascórbico",
    "Ác. fólico": "Ác._fólico"
}
for old, new in replacements.items():
    raw_text = raw_text.replace(old, new)

# --- Step 0c: Modify lines following any line that starts with "Energía ENERC" ---
lines = raw_text.strip().splitlines()
modified_lines = []
i = 0
while i < len(lines):
    current_line = lines[i]
    modified_lines.append(current_line)
    if current_line.strip().startswith("Energía ENERC"):
        if i + 1 < len(lines):
            next_line = lines[i+1]
            # If the next line doesn't already start with "Energía ENERC", prepend it.
            if not next_line.strip().startswith("Energía ENERC"):
                modified_lines.append("Energía ENERC " + next_line)
                i += 2
                continue
    i += 1
raw_text = "\n".join(modified_lines)

# --- Step A: Split raw_text into lines ---
lines = raw_text.strip().splitlines()

# --- Step B: Process header block to extract food names ---
# Get all non-empty lines between "Componente alimentario" and "Nutriente"
header_lines = []
capture = False
for line in lines:
    if "Componente alimentario" in line:
        capture = True
        # Remove "Componente alimentario" from the line if it is on the same line.
        line = line.replace("Componente alimentario", "").strip()
        if line:
            header_lines.append(line)
        continue
    if "Nutriente" in line:
        break
    if capture and line.strip():
        header_lines.append(line.strip())

# In this raw text, each food appears on one line.
food_count = len(header_lines)

# Create extended headers: for each food, generate "[food] F" and "[food] en 100 g"
final_headers = []
for h in header_lines:
    final_headers.append(h + " F")
    final_headers.append(h + " en 100 g")

# --- Step C: Build the full column headers ---
# Four fixed columns plus 2 columns per food item.
columns = ["Grupos", "Nutriente", "Tagname", "Unidad"] + final_headers

# --- Step D: Process nutrient rows ---
data_rows = []
expected_pair_count = food_count * 2  # Expect 2 values per food

# Find the starting index (the line after "Nutriente Tagname Unidad")
start_index = next(i for i, line in enumerate(lines) if "Nutriente Tagname Unidad" in line) + 1

# Define possible group names.
group_names = ["Elementos principales", "elementos principales", "Ác. grasos", "Minerales", "Vitaminas"]
current_group = None

for line in lines[start_index:]:
    line = line.strip()
    if not line:
        continue
    if line in group_names:
        current_group = line
        continue

    parts = line.split()
    if len(parts) < 4:
        continue

    # The first three tokens are nutrient info.
    nutrient = parts[0]
    tagname = parts[1]
    unit = parts[2]
    raw_values = parts[3:]
    
    # Process raw_values:
    # - If the count equals food_count, duplicate each value.
    # - If it equals expected_pair_count, assume they are flag–value pairs.
    # - Otherwise, try to pair them up.
    if len(raw_values) == food_count:
        processed_values = []
        for v in raw_values:
            processed_values.extend([v, v])
    elif len(raw_values) == expected_pair_count:
        processed_values = raw_values[:expected_pair_count]
    else:
        processed_values = []
        i_val = 0
        while i_val < len(raw_values) and len(processed_values) < expected_pair_count:
            processed_values.append(raw_values[i_val])
            if i_val + 1 < len(raw_values):
                processed_values.append(raw_values[i_val+1])
            i_val += 2

    # Pad if needed
    processed_values += [''] * (expected_pair_count - len(processed_values))
    
    row = [current_group, nutrient, tagname, unit] + processed_values
    data_rows.append(row)

# --- Step E: Create DataFrame and Export ---
df = pd.DataFrame(data_rows, columns=columns)

# ──────────────────────────────────────────────────────────────
# 0️⃣  𝙶𝚛𝚘𝚞𝚙 𝚍𝚎𝚏𝚒𝚗𝚒𝚝𝚒𝚘𝚗 (row ➜ group)
# ──────────────────────────────────────────────────────────────
row2group = {}
row2group.update({r: "Elementos principales" for r in [
    "Humedad", "Fibra_dietética", "Hidratos_de_C", "Proteínas", "Lípidos_tot"]})
row2group.update({r: "Ácidos grasos" for r in [
    "Saturados", "Monoinsaturados", "Poliinsaturados", "Linolénico",
    "Eicosapentaenoico", "Docosahexaenoico", "Colesterol"]})
row2group.update({r: "Minerales" for r in [
    "Calcio", "Fósforo", "Hierro", "Magnesio", "Sodio", "Potasio", "Zinc"]})
row2group.update({r: "Vitaminas" for r in [
    "RAE_(vit_A)", "Ác._ascórbico", "Tiamina", "Riboflavina",
    "Niacina", "Piridoxina", "Ác._fólico", "Cianocobalamina"]})

# ──────────────────────────────────────────────────────────────
# 1️⃣  Row & column totals (unchanged from your script)
# ──────────────────────────────────────────────────────────────
# Prompt user for row blanks, skipping first 2 rows
row_blanks = {}
print("=== HOW MANY BLANKS PER ROW? (Skipping first 2 rows) ===")
for row in df.index[2:]:
    nutrient_name = df.loc[row, "Nutriente"]
    while True:
        try:
            val = int(input(f"Enter blanks count for row '{nutrient_name}': "))
            row_blanks[nutrient_name] = val
            break
        except ValueError:
            print("❌ Please enter a valid integer.")


# Prompt user for each relevant column
col_blanks = {}
print("\n=== HOW MANY BLANKS PER COLUMN? (Only ' F' and 'en 100 g' columns) ===")
for col in df.columns:
    if col.endswith(" F") or col.endswith("en 100 g"):
        while True:
            try:
                val = int(input(f"Enter blanks count for column '{col}': "))
                col_blanks[col] = val
                break
            except ValueError:
                print("❌ Please enter a valid integer.")

# Automatically calculate grand totals
F_total_required = sum(v for k, v in col_blanks.items() if k.endswith(" F"))
g100_total_required = sum(v for k, v in col_blanks.items() if k.endswith("en 100 g"))

cols_only = [col for col in df.columns if col in col_blanks]
original_values = {
    df.loc[i, "Nutriente"]: [v for v in df.loc[i, cols_only].values if pd.notna(v) and str(v).strip() != ""]
    for i in df.index[2:]
    if df.loc[i, "Nutriente"] in row2group
}

# ──────────────────────────────────────────────────────────────
# 2️⃣  Group-column quotas (only for columns that have blanks)
# ──────────────────────────────────────────────────────────────

group_col_req = {}
groups = ["Elementos principales", "Ácidos grasos", "Minerales", "Vitaminas"]

# Only columns with > 0 blanks
relevant_cols_with_blanks = [col for col, val in col_blanks.items() if val > 0]

print("\n\n\n=== GROUP-COLUMN QUOTAS (asked per column) ===")
for col in relevant_cols_with_blanks:
    print(f"\n➡️  Now entering quotas for column: '{col}'")
    for group in groups:
        while True:
            try:
                val = input(f"  ➜ Number of X’s for ({group}, '{col}'), or leave blank to skip: ").strip()
                if val == "":
                    break
                num = int(val)
                group_col_req[(group, col)] = num
                break
            except ValueError:
                print("❌ Please enter a valid integer or leave blank to skip.")

# ──────────────────────────────────────────────────────────────
# 3️⃣  Manually specify known blank cells for validation
# ──────────────────────────────────────────────────────────────

print("\n\n\n=== MANUAL KNOWN BLANK LOCATIONS (one per column) ===")
print("You'll now be asked to identify ONE row that contains a blank for up to 4 columns.\n")


# Only sample from columns with > 0 blanks
nonzero_blank_cols = [col for col, val in col_blanks.items() if val > 0]
sampled_cols = random.sample(nonzero_blank_cols, min(4, len(nonzero_blank_cols)))

known_cells = set()

# Build mapping: Nutriente name ➜ internal row index
nutriente_to_rowkey = {df.loc[i, "Nutriente"]: df.loc[i, "Nutriente"] for i in df.index[2:] if df.loc[i, "Nutriente"] in row_blanks}
valid_nutrientes = list(nutriente_to_rowkey.keys())

for col in sampled_cols:
    print(f"\n➡️  For column: '{col}'")
    # print("   Valid nutrient names:")
    # for label in valid_nutrientes:
        # print(f"     • {label}")

    while True:
        input_line = input("  ➜ Type one or more nutrient names (space-separated) from the Nutriente list: ").strip()
        input_labels = input_line.split()

        invalids = [name for name in input_labels if name not in nutriente_to_rowkey]
        if invalids:
            print("❌ Invalid name(s):", ", ".join(invalids))
            print("   Valid nutrient names:")
            for label in valid_nutrientes:
                print(f"     • {label}")
            continue

        for name in input_labels:
            known_cells.add((nutriente_to_rowkey[name], col))
        break

            

# ──────────────────────────────────────────────────────────────
#     Preliminaries
# ──────────────────────────────────────────────────────────────
rows = list(row_blanks.keys())
cols = list(col_blanks.keys())

relevant_cols = [c for c, k in col_blanks.items() if k > 0]   # only columns that can contain blanks
rc = len(relevant_cols)

# pre-allocate helpers
row_patterns, row_F, row_100, row_gc = {}, {}, {}, {}
for r in rows:
    need = row_blanks[r]
    opts, f_cnt, g_cnt, gc_cnt = [], [], [], []
    for mask in itertools.product([0, 1], repeat=rc):
        pat = dict(zip(relevant_cols, mask))     # ← use relevant_cols

        if sum(pat.values()) != need:
            continue
        # honour the individually-fixed blanks
        if any(pat[col] == 0 for (row, col) in known_cells if row == r):
            continue
        opts.append(pat)
        f_cnt.append(sum(v for c,v in pat.items() if c.endswith(" F")))
        g_cnt.append(sum(v for c,v in pat.items() if c.endswith("en 100 g")))
        gc_cnt.append({(row2group[r], c): v
                       for c,v in pat.items()
                       if (row2group[r], c) in group_col_req})
    row_patterns[r], row_F[r], row_100[r], row_gc[r] = opts, f_cnt, g_cnt, gc_cnt

# ──────────────────────────────────────────────────────────────
# 4️⃣  Back-tracking with
#        • column totals
#        • grand totals
#        • group-column quotas
# ──────────────────────────────────────────────────────────────
targets     = {c: col_blanks[c] for c in relevant_cols}
col_totals  = {c: 0 for c in relevant_cols}
group_tot   = defaultdict(int)            # running (group,col) tally
assignment  = {}
solution    = None

def backtrack(i, F_so_far, G_so_far):
    global solution
    if solution: return
    if i == len(rows):
        if (all(col_totals[c] == targets[c] for c in relevant_cols)
                and F_so_far == F_total_required
                and G_so_far == g100_total_required
                and all(group_tot[key] == need
                        for key, need in group_col_req.items())):
            solution = assignment.copy()
        return

    r = rows[i]
    for k, pat in enumerate(row_patterns[r]):
        f_inc, g_inc = row_F[r][k], row_100[r][k]
        # grand totals
        if F_so_far + f_inc > F_total_required:    continue
        if G_so_far + g_inc > g100_total_required: continue
        # column totals
        if any(col_totals[c] + pat[c] > targets[c] for c in relevant_cols):
            continue
        # group-column quotas
        gc_inc = row_gc[r][k]
        if any(group_tot[key] + v > group_col_req[key] for key, v in gc_inc.items()):
            continue

        # choose
        assignment[r] = pat
        for c in relevant_cols: col_totals[c] += pat[c]
        for key, v in gc_inc.items(): group_tot[key] += v

        backtrack(i + 1, F_so_far + f_inc, G_so_far + g_inc)

        # undo
        for c in relevant_cols: col_totals[c] -= pat[c]
        for key, v in gc_inc.items(): group_tot[key] -= v
        del assignment[r]

backtrack(0, 0, 0)
if solution is None:
    raise RuntimeError("❌  No grid satisfies all constraints.")

# ──────────────────────────────────────────────────────────────
# 5️⃣  Build DataFrame with a “Group” column & save
# ──────────────────────────────────────────────────────────────
grid = [[0]*len(cols) for _ in rows]
cindex = {c:j for j,c in enumerate(cols)}
for r, pat in solution.items():
    i = rows.index(r)
    for c, v in pat.items(): grid[i][cindex[c]] = v

df = pd.DataFrame(
    {
        "Group": [row2group[r] for r in rows],
        **{c: ["X" if grid[i][cindex[c]] else "" for i in range(len(rows))]
           for c in cols}
    },
    index=rows,
)

# Build output DataFrame: "X" where blanks, original values elsewhere
filled_data = []
for i, row in enumerate(rows):
    row_data = []
    original_row_values = iter(original_values[row])
    for c in cols:
        if grid[i][cindex[c]]:  # is an X
            row_data.append("-")
        else:
            val = next(original_row_values, "")  # Use empty string if exhausted
            row_data.append(val)
    filled_data.append(row_data)

# Build the final DataFrame with group and index
final_df = pd.DataFrame(
    {
        "Group": [row2group[r] for r in rows],
        **{c: [filled_data[i][j] for i in range(len(rows))] for j, c in enumerate(cols)}
    },
    index=rows,
)

# Extract first two skipped rows from original df
skipped_rows = df.iloc[:2].copy()
skipped_rows["Group"] = "Elementos principales"

# Align with final_df structure (same column order)
ordered_columns = ["Group"] + cols
skipped_rows = skipped_rows[ordered_columns]

# ✅ Prepend skipped rows to final_df
final_df = pd.concat([skipped_rows, final_df], axis=0)

# Save to Excel
final_df.to_excel("final_grid_2.xlsx", index=True)
print("✅  Grid saved to final_grid_2.xlsx")

