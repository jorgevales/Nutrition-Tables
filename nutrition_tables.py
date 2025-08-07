import pandas as pd
import os

# --- Step 0: Get the raw text ---
# The first line will be the file name, 6.8.2 Carnes importadas for example
# First two lines will be excluded
# Headers should be between lines 'Componente alimentario' and 'Nutriente Tagname Unidad...'

raw_text = """
6.9.3 Pescados enlatados
Rm-Pen-1 Rm-Pen-2 Rm-Pen-3 Rm-Pen-4
Componente alimentario 
Atún en aceite
Salmón en aceite
Sardinas en aceite
Sardinas en tomate
Nutriente Tagname Unidad F En 100 g F En 100 g F En 100 g F En 100 g
Elementos principales
Energía ENERC kcal 281 302 290 197
kJ 1177 1263 1213 801
Humedad WATER % 1 60.60 1 64.20 3 59.60 1 64.30
Fibra dietética FIBTG g 1 0.00 1 0.00 1 0.00 1 0.00
Hidratos de C CHOCDF g 1 0.00 1 0.00 R 0.60 R 1.70
Proteínas PROCNT g 1 24.20 1 21.70 1 18.80 1 18.70
Lípidos tot FAT g 1 20.50 R 23.90 1 23.60 1 12.20
Ác. grasos
Saturados FASAT g 1 5.00 R 6.00 1 - 1 2.80
Monoinsaturados FAMS g 1 4.00 R 7.40 1 - 1 2.60
Poliinsaturados FAPU g 1 8.00 R 8.70 1 - 1 4.40
Linolénico F18D3N3 g 1 0.05 1 0.1 1 - 1 0.1
Eicosapentaenoico F20D5N3 g 1 0.3 1 0.9 1 - 1 -
Docosahexaenoico F22D6N3 g 1 1 1 1.6 1 - 1 -
Colesterol CHOLE mg 1 55.00 1 80.00 1 - 1 120.00
minerales
Calcio CA mg 1 7.00 1 79.00 1 303.00 1 449.00
Fósforo P mg 2 294.00 2 305.00 2 434.00 2 478.00
Hierro FE mg 1 1.20 1 0.90 1 5.20 1 4.10
Magnesio MG mg 1 23.00 1 32.00 1 - 1 -
Sodio NA mg 1 800.00 1 473.00 1 510.00 1 400.00
Potasio K mg 1 301.00 1 126.00 1 560.00 1 320.00
Zinc ZN mg 1 0.40 1 0.90 1 3.00 1 2.70
Vitaminas
RAE (vit A) VITA µg 1 6.00 1 18.00 1 9.00 1 9.00
Ác. ascórbico ASCL mg 1 0.00 1 0.00 1 0.00 1 0.00
Tiamina THIA mg 1 0.04 1 - 1 0.01 1 0.01
Riboflavina RIBF mg 1 0.10 1 0.08 1 0.27 1 0.27
Niacina NIA mg 1 11.10 1 7.20 1 5.30 1 5.30
Piridoxina VITB6A mg 1 0.44 1 0.75 1 0.48 1 0.48
Ác. fólico FOL µg 1 15.00 1 26.00 1 8.00 1 8.00
Cianocobalamina VITB12 µg 1 5.00 1 5.00 1 28.00 1 28.00
"""

# --- Step 0a: Extract file name from the first line and remove the first two lines ---
all_lines = raw_text.strip().splitlines()
# First line will be used for naming the file.
file_name_line = all_lines[0].strip()
# Keep all lines except the first two and the last one.
data_lines = all_lines[2:]
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

# --- Step F: Export with overwrite protection ---
output_filename = file_name_line.strip() + ".xlsx"

while os.path.exists(output_filename):
    print(f"⚠️  The file '{output_filename}' already exists.")
    new_name = input("Please type a new name for the file (without extension): ").strip()
    if new_name:
        output_filename = new_name + ".xlsx"
    else:
        print("❌ No valid name provided. Export cancelled.")
        exit()

df.to_excel(output_filename, index=False)
print(f"✅ File exported as '{output_filename}'")