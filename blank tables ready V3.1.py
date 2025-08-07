import pandas as pd
import tkinter as tk
import os

# ─────────────────────────────────────────────
# GUI: Multi-select nutrient rows for one column
# ─────────────────────────────────────────────
def select_blanks_for_column(col_name, row_labels):
    selection = []

    def confirm():
        selection.extend([row_labels[i] for i in listbox.curselection()])
        root.destroy()

    def select_all():
        listbox.select_set(0, tk.END)

    def select_none():
        listbox.select_clear(0, tk.END)

    root = tk.Tk()
    root.title(f"Select blank rows for: {col_name}")
    root.attributes("-fullscreen", True)

    title_label = tk.Label(
        root, text=f"Column:\n{col_name}", font=("Helvetica", 32, "bold"), pady=30
    )
    title_label.pack()

    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    tk.Button(button_frame, text="🧼 Select NONE", font=("Helvetica", 14), command=select_none).grid(row=0, column=0, padx=20)
    tk.Button(button_frame, text="✅ Confirm Selection", font=("Helvetica", 18), command=confirm).grid(row=0, column=1, padx=20)
    tk.Button(button_frame, text="📄 Select ALL", font=("Helvetica", 14), command=select_all).grid(row=0, column=2, padx=20)

    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=60, pady=20)

    listbox = tk.Listbox(
        frame, selectmode=tk.MULTIPLE, font=("Courier", 16),
        height=30, width=80, exportselection=False
    )
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    listbox.config(yscrollcommand=scrollbar.set)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    for row in row_labels:
        listbox.insert(tk.END, row)

    root.mainloop()
    return set(selection)

# ─────────────────────────────────────────────
# Insert blank logic
# ─────────────────────────────────────────────
def insert_blank_in_column(df, row_idx, col_name):
    # 1. Get full row values (only the food values, starting at col 4)
    full_values = df.iloc[row_idx, 4:].tolist()

    # 2. Identify the position (within food columns) of the column to blank
    col_pos = df.columns.get_loc(col_name) - 4  # offset inside food columns

    # 3. Shift values: insert "-" and push everything right
    new_values = full_values[:col_pos] + ["-"] + full_values[col_pos:]
    new_values = new_values[:len(full_values)]  # Trim back to original length

    # 4. Replace in DataFrame
    df.iloc[row_idx, 4:] = new_values



# ─────────────────────────────────────────────
# Raw text (your full dataset)
# ─────────────────────────────────────────────
raw_text = """
6.9.4 mariscos frescos y enlatados (continuación 2)
Rm-mFe-13 Rm-mFe-14 Rm-mFe-15 Rm-mFe-16
Componente alimentario
Ostión del Golfo de México
Ostión del Pacífico
Ostiones sin concha
Pulpo crudo
Nutriente Tagname Unidad F En 100 g F En 100 g F En 100 g F En 100 g
elementos principales
Energía ENERC kcal 50 50 78 59
kJ 209 209 326 249
Humedad WATER % 90.57 93.95 1 81.80 1 82.40
Fibra dietética FIBTG mg 0.00 0.00 1 0.00 1 0.00
Hidratos de C CHOCDF mg R 4.90 1 0.00
Proteínas PROCNT mg 9 3.85 1 9.40 1 12.60
Lípidos tot FAT mg 2 1 2.30 1 1.00
Ác. grasos
Saturados FASAT mg 1 0.50 1 0.30
Monoinsaturados FAMS mg 1 0.40 1 0.10
Poliinsaturados FAPU mg 1 0.90 1 0.30
Linolénico F18D3N3 mg 1 0.01 1 -
Eicosapentaenoico F20D5N3 mg 3.57 10.1 1 0.4 1 0.1
Docosahexaenoico F22D6N3 mg 14.28 6.62 1 0.2 1 0.1
Colesterol CHOLE mg 9.29 7.95 1 38.00 1 -
minerales
Calcio CA mg 1 91.00 1 39.00
Fósforo P mg R - 2 109.00
Hierro FE mg 1 6.30 1 2.50
Magnesio MG mg 1 32.00 1 -
Sodio NA mg 1 200.00 1 89.00
Potasio K mg 1 175.00 1 274.00
Zinc ZN mg 1 74.70 1 1.70
Vitaminas
RAE (vit A) VITA µg 1 - 1 -
Ác. ascórbico ASCL mg 1 5.00 1 0.00
Tiamina THIA mg 1 0.13 1 0.02
Riboflavina RIBF mg 1 0.09 1 0.07
Niacina NIA mg 1 1.90 1 1.30
Piridoxina VITB6A mg 1 - 1 -
Ác. fólico FOL µg 1 - 1 13.00
Cianocobalamina VITB12 µg 1 - 1 2.00
"""

# ─────────────────────────────────────────────
# Preprocessing
# ─────────────────────────────────────────────
lines = raw_text.strip().splitlines()
file_name = lines[0].strip()
data_lines = lines[2:]

replacements = {
    "Fibra dietética": "Fibra_dietética",
    "Hidratos de C": "Hidratos_de_C",
    "Lípidos tot": "Lípidos_tot",
    "RAE (vit A)": "RAE_(vit_A)",
    "Acetato de vit A": "Acetato_de_vit_A",
    "Ác. ascórbico": "Ác._ascórbico",
    "Ác. fólico": "Ác._fólico",
    "Vit E (α tocoferol)": "Vit_E_(α_tocoferol)",
    "Vit D (colecalciferol)": "Vit_D_(colecalciferol)",
    "Etapa de consumo": "Etapa_de_consumo"
}

for old, new in replacements.items():
    data_lines = [line.replace(old, new) for line in data_lines]

def duplicate_energia(lines):
    result = []
    i = 0
    while i < len(lines):
        result.append(lines[i])
        if lines[i].strip().startswith("Energía ENERC") and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if not next_line.startswith("Energía ENERC"):
                result.append("Energía ENERC " + next_line)
                i += 2
                continue
        i += 1
    return result

lines = duplicate_energia(data_lines)

header_lines = []
capture = False
for line in lines:
    if "Componente alimentario" in line:
        capture = True
        line = line.replace("Componente alimentario", "").strip()
        if line:
            header_lines.append(line)
        continue
    if "Nutriente" in line:
        break
    if capture and line.strip():
        header_lines.append(line.strip())

original_column_pairs = []
for h in header_lines:
    original_column_pairs.append(h + " F")
    original_column_pairs.append(h + " en 100 g")

columns = ["Grupos", "Nutriente", "Tagname", "Unidad"] + original_column_pairs
food_count = len(header_lines)
expected_pair_count = food_count * 2

group_names = ["Elementos principales", "elementos principales", "Ác. grasos", "Minerales", "Vitaminas"]
data_rows = []
current_group = None
start_index = next(i for i, line in enumerate(lines) if "Nutriente Tagname Unidad" in line) + 1

special_rows = []
for line in lines[start_index:]:
    line = line.strip()
    if not line:
        continue
    if line.lower() in [g.lower() for g in group_names]:
        current_group = next(g for g in group_names if g.lower() == line.lower())
        continue

    parts = line.split()
    if len(parts) < 4:
        continue

    nutrient, tagname, unit = parts[:3]
    raw_values = parts[3:]

    if nutrient in ["Orden", "Etapa_de_consumo"]:
        filled = [""] * len(original_column_pairs)
        for i, val in enumerate(raw_values):
            if 4 + i < len(columns):
                filled[i] = val
        special_rows.append([current_group, nutrient, tagname, unit] + filled)
        continue

    if len(raw_values) == food_count:
        processed_values = [v for val in raw_values for v in (val, val)]
    elif len(raw_values) == expected_pair_count:
        processed_values = raw_values[:expected_pair_count]
    else:
        processed_values = []
        i = 0
        while i < len(raw_values) and len(processed_values) < expected_pair_count:
            processed_values.append(raw_values[i])
            if i + 1 < len(raw_values):
                processed_values.append(raw_values[i + 1])
            i += 2
    processed_values += [''] * (expected_pair_count - len(processed_values))
    data_rows.append([current_group, nutrient, tagname, unit] + processed_values)

df = pd.DataFrame(data_rows, columns=columns)

# ─────────────────────────────────────────────
# Ask blanks column-by-column in original order
# ─────────────────────────────────────────────
nutrient_names = [df.loc[i, 'Nutriente'].replace("_", " ") for i in range(2, len(df))]
title_to_idx = {name.upper(): i for i, name in enumerate(nutrient_names, start=2)}

for col in original_column_pairs:
    selected_rows = select_blanks_for_column(col, nutrient_names)
    for row in selected_rows:
        idx = title_to_idx[row.upper()]
        insert_blank_in_column(df, idx, col)

# ─────────────────────────────────────────────
# Add special rows at the end of the DataFrame
# ─────────────────────────────────────────────
for row in special_rows:
    df.loc[len(df)] = row

# ─────────────────────────────────────────────
# Export with overwrite protection
# ─────────────────────────────────────────────
output_filename = file_name.strip() + ".xlsx"

while os.path.exists(output_filename):
    print(f"⚠️  The file '{output_filename}' already exists.")
    new_name = input("Please type a new name for the file (without extension): ").strip()
    if new_name:
        output_filename = new_name + ".xlsx"
    else:
        print("❌ No valid name provided. Export cancelled.")
        exit()

df.to_excel(output_filename, index=False)
print(f"\n✅ File saved as '{output_filename}'")
