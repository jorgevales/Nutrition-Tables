#!/usr/bin/env python3
# nutri_blanks_fixed.py
# --------------------------------------------------------------------------- #
#  Build initial table  ➜  apply hard‐coded blanks  ➜  save final Excel file   #
# --------------------------------------------------------------------------- #

import pandas as pd
import numpy as np
import re
from copy import deepcopy
from itertools import combinations, product

# --------------------------------------------------------------------------- #``
#  (1)  ------------  PARSE RAW TABLE  -------------------------------------- #
# --------------------------------------------------------------------------- #
RAW_TEXT = r"""
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

# text substitutions (diacritics‐safe column names)
replacements = {
    "Fibra dietética": "Fibra_dietética",
    "Hidratos de C":   "Hidratos_de_C",
    "Lípidos tot":     "Lípidos_tot",
    "RAE (vit A)":     "RAE_(vit_A)",
    "Ác. ascórbico":   "Ác._ascórbico",
    "Ác. fólico":      "Ác._fólico",
}

# Keep everything except the first 2 and the last line
lines = RAW_TEXT.strip().splitlines()
data_lines = lines[2:-1]
for i, line in enumerate(data_lines):
    for old, new in replacements.items():
        data_lines[i] = data_lines[i].replace(old, new)
RAW_TEXT = "\n".join(data_lines)

# Fix missing “Energía ENERC” prefix on the next line
tmp, fixed = RAW_TEXT.splitlines(), []
i = 0
while i < len(tmp):
    fixed.append(tmp[i])
    if tmp[i].lstrip().startswith("Energía ENERC") and i + 1 < len(tmp):
        nxt = tmp[i + 1]
        if not nxt.lstrip().startswith("Energía ENERC"):
            fixed.append("Energía ENERC " + nxt)
            i += 2
            continue
    i += 1
RAW_TEXT = "\n".join(fixed)

# ----------- build headers (Food 1 F, Food 1 en 100 g, Food 2 F, …) ----------
lines = RAW_TEXT.splitlines()
foods = []
cap = False
for ln in lines:
    if "Componente alimentario" in ln:
        cap = True
        ln = ln.replace("Componente alimentario", "").strip()
        if ln:
            foods.append(ln)
        continue
    if "Nutriente" in ln:
        break
    if cap and ln.strip():
        foods.append(ln.strip())

mov_cols = [col for food in foods for col in (f"{food} F", f"{food} en 100 g")]
columns  = ["Grupos", "Nutriente", "Tagname", "Unidad"] + mov_cols

# ----------------- numeric rows ----------------------------------------------
start = next(i for i, ln in enumerate(lines) if "Nutriente Tagname Unidad" in ln) + 1
group_lbls = ["Elementos principales", "elementos principales", "Ác. grasos", "Minerales", "Vitaminas"]
rows, grp = [], None
need = len(foods) * 2       # flag/value per food

for ln in lines[start:]:
    if not (ln := ln.strip()):
        continue
    if ln in group_lbls:
        grp = ln
        continue
    parts = ln.split()
    if len(parts) < 4:
        continue
    nutr, tag, unit, *vals = parts
    if len(vals) == len(foods):
        # only one number per food → duplicate to (F, En 100 g)
        vals = [v for v in vals for _ in (0, 1)]
    elif len(vals) < need:
        pad, j = [], 0
        while j < len(vals):
            pad.append(vals[j])
            pad.append(vals[j + 1] if j + 1 < len(vals) else "")
            j += 2
        vals = (pad + [""])[:need]
    rows.append([grp, nutr, tag, unit] + vals[:need])

df = pd.DataFrame(rows, columns=columns)
df.to_excel("original_table.xlsx", index=False)
print("✅  Original table saved to 'original_table.xlsx'")

# --------------------------------------------------------------------------- #
#  (2)  ------------  CLEAR DATA AND APPLY HARD-CODED 'X'  ------------------- #
# --------------------------------------------------------------------------- #

# (2.1) Clear everything in the “F” / “en 100 g” columns
for col in mov_cols:
    df[col] = ""

# (2.2) Define how many blanks each nutrient-row needs
nutrient_to_blank_row = {
    'Humedad': 2,
    'Fibra_dietética': 1,
    'Hidratos_de_C': 2,
    'Proteínas': 2,
    'Lípidos_tot': 2,
    'Saturados': 2,
    'Monoinsaturados': 2,
    'Poliinsaturados': 2,
    'Linolénico': 2,
    'Eicosapentaenoico': 2,
    'Docosahexaenoico': 2,
    'Colesterol': 2,
    'Calcio': 2,
    'Fósforo': 2,
    'Hierro': 2,
    'Magnesio': 2,
    'Sodio': 2,
    'Potasio': 2,
    'Zinc': 2,
    'RAE_(vit_A)': 2,
    'Ác._ascórbico': 2,
    'Tiamina': 2,
    'Riboflavina': 2,
    'Niacina': 2,
    'Piridoxina': 2,
    'Ác._fólico': 2,
    'Cianocobalamina': 2
}

# (2.3) Define how many blanks each column needs
col_targets = {
    'Bagre, chihuil F': 5,
    'Bagre, chihuil en 100 g': 0,
    'Besugo F': 27,
    'Besugo en 100 g': 21,
    'Bonito, bacoreta, cachorra F': 0,
    'Bonito, bacoreta, cachorra en 100 g': 0,
    'Boquerón crudo F': 0,
    'Boquerón crudo en 100 g': 0,
    'Boquerón cocido F': 0,
    'Boquerón cocido en 100 g': 0,
    'Boquilla, ronco F': 0,
    'Boquilla, ronco en 100 g': 0,
    'Cabrilla F': 0,
    'Cabrilla en 100 g': 0,
}

groups_blanks = {'Elementos principales': 8}

elementos_principales_targets = {
    "F": 6,
    "en 100 g": 2
}

# New subgroup blank targets inside group
group_column_type_targets = {
    'Elementos principales': {'F': 6, 'en 100 g': 2}
}

# --- Setup ---
modifiable_cols = [col for col in df.columns if col in col_targets]
remaining_col_blanks = deepcopy(col_targets)
group_column_type_used = {
    g: {'F': 0, 'en 100 g': 0} for g in df['Grupos'].unique() if isinstance(g, str)
}

# Step 1 – Pre-fill fully blank columns (27 blanks)
FULL_BLANK = 27
for col, target in col_targets.items():
    if target == FULL_BLANK:
        for i, row in df.iterrows():
            nutrient = row['Nutriente']
            group = row['Grupos']
            if nutrient in nutrient_to_blank_row and nutrient_to_blank_row[nutrient] > 0:
                df.at[i, col] = "X"
                remaining_col_blanks[col] -= 1
                nutrient_to_blank_row[nutrient] -= 1
                if group in groups_blanks:
                    groups_blanks[group] -= 1
                if group in group_column_type_used:
                    if col.endswith("F"):
                        group_column_type_used[group]["F"] += 1
                    elif col.endswith("en 100 g"):
                        group_column_type_used[group]["en 100 g"] += 1

# Step 2 – Backtracking group assignment
def try_group_assignment(group_name, group_rows):
    row_indices = group_rows.index.tolist()
    needed_blanks = [nutrient_to_blank_row.get(df.at[i, "Nutriente"], 0) for i in row_indices]
    
    col_combos_per_row = []
    for i, blanks in zip(row_indices, needed_blanks):
        valid_cols = [col for col in modifiable_cols if remaining_col_blanks[col] > 0]
        combos = list(combinations(valid_cols, blanks))
        col_combos_per_row.append(combos)

    for row_combo_set in product(*col_combos_per_row):
        # Start trial
        trial_df = df.copy()
        trial_col_blanks = deepcopy(remaining_col_blanks)
        trial_row_blanks = deepcopy(nutrient_to_blank_row)
        trial_group_blanks = deepcopy(groups_blanks)
        trial_type_used = deepcopy(group_column_type_used)

        valid = True
        for idx, combo in zip(row_indices, row_combo_set):
            nutrient = trial_df.at[idx, "Nutriente"]
            if nutrient not in trial_row_blanks:
                continue  # Skip rows like 'Energía' that are not part of the blanking rules

            counts = {'F': 0, 'en 100 g': 0}

            for c in combo:
                if trial_col_blanks[c] <= 0:
                    valid = False
                    break
                if c.endswith("F"):
                    counts["F"] += 1
                elif c.endswith("en 100 g"):
                    counts["en 100 g"] += 1
            if not valid:
                break
            if group_name in group_column_type_targets:
                limits = group_column_type_targets[group_name]
                used = trial_type_used[group_name]
                if any(used[t] + counts[t] > limits[t] for t in counts):
                    valid = False
                    break
            if trial_group_blanks.get(group_name, 0) < sum(counts.values()):
                valid = False
                break

            for c in combo:
                trial_df.at[idx, c] = "X"
                trial_col_blanks[c] -= 1
            trial_row_blanks[nutrient] -= sum(counts.values())
            trial_group_blanks[group_name] -= sum(counts.values())
            for t in counts:
                trial_type_used[group_name][t] += counts[t]

        if valid:
            return trial_df, trial_col_blanks, trial_row_blanks, trial_group_blanks, trial_type_used

    return None, None, None, None, None

# Loop over groups
for group in df['Grupos'].unique():
    if group not in groups_blanks:
        continue
    group_rows = df[df['Grupos'] == group]
    result = try_group_assignment(group, group_rows)
    if result[0] is not None:
        df, remaining_col_blanks, nutrient_to_blank_row, groups_blanks, group_column_type_used = result
    else:
        print(f"⚠️ No valid configuration found for group: {group}")

output_path = "all_blanks_file.xlsx"
df.to_excel(output_path, index=False)
print(f"✅ Final distributed table saved to '{output_path}'")
