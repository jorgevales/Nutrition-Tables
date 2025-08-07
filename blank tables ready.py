import pandas as pd
import re
import numpy as np
import xlsxwriter as xlsxwriter

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

# === EXPORT TO EXCEL ===
output_filename = file_name_line + "_1.xlsx"
df.to_excel(output_filename, index=False)
print(f"\n✅ File saved as '{output_filename}'")


# === ASK FOR BLANKS & MODIFY DATAFRAME ===

for col in final_headers:
    response = input(f"\nHow many blanks in column '{col}'? (enter 'A' for all, or number): ").strip()

    if response.upper() == "A":
        # Fill all except the first two rows
        for idx in range(2, len(df)):
            val_index = df.columns.get_loc(col)
            current_val = df.iat[idx, val_index]
            if pd.isna(current_val) or current_val == "":
                df.iat[idx, val_index] = "-"
            else:
                # Push everything in the row to the right
                row_vals = df.iloc[idx, 4:].tolist()
                rel_idx = val_index - 4
                for j in range(len(row_vals) - 1, rel_idx, -1):
                    row_vals[j] = row_vals[j - 1]
                row_vals[rel_idx] = "-"
                df.iloc[idx, 4:] = row_vals
        continue

    try:
        remaining_blanks = int(response)
    except ValueError:
        print("❌ Invalid input. Skipping column.")
        continue

    if remaining_blanks == 0:
        continue  # Column fully filled, skip to next

    apply_rest = False  # Set when user types 'R'

    for idx in range(2, len(df)):  # Skip first 2 rows
        if remaining_blanks <= 0:
            break

        row_name = df.loc[idx, "Nutriente"]

        if not apply_rest:
            row_input = input(f"How many blanks in row '{row_name}' for column '{col}'? (Remaining: {remaining_blanks}): ").strip()

            if row_input.upper() == 'R':
                apply_rest = True
                row_blanks = 1
            else:
                try:
                    row_blanks = int(row_input)
                except ValueError:
                    print("❌ Invalid input. Skipping row.")
                    continue
        else:
            row_blanks = 1

        if row_blanks <= 0:
            continue

        val_index = df.columns.get_loc(col)

        for _ in range(min(row_blanks, remaining_blanks)):
            current_val = df.iat[idx, val_index]
            if pd.isna(current_val) or current_val == "":
                df.iat[idx, val_index] = "-"
            else:
                # Push everything from this position to the right
                row_vals = df.iloc[idx, 4:].tolist()
                rel_idx = val_index - 4
                for j in range(len(row_vals) - 1, rel_idx, -1):
                    row_vals[j] = row_vals[j - 1]
                row_vals[rel_idx] = "-"
                df.iloc[idx, 4:] = row_vals

        remaining_blanks -= row_blanks

# === EXPORT TO EXCEL ===
output_filename = file_name_line + "_2.xlsx"
df.to_excel(output_filename, index=False)
print(f"\n✅ File saved as '{output_filename}'")