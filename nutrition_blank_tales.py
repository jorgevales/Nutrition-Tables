import re
import pandas as pd

# ================================
# User-supplied multi‐section raw text.
# Sections are separated by a full blank line.
# ================================
raw_text = """
6.9.1 algas
Rm-am-R-3 Rm-am-c-1 Rm-am-c-2 Rm-am-c-3 Rm-am-c-4 Rm-am-c-5
Componente alimentario
Hypnea valentiae
Sargassum sinicola
Sargassum herporizum
Padina durvillaei
Globito
Hydroclathrus clathratus
Nutriente Tagname Unidad F En 100 g F En 100 g F En 100 g F En 100 g F En 100 g F En 100 g
elementos principales
Energía ENERC kcal 158 181 195 198 99 105
kJ 660 758 815 828 414 441
Humedad WATER % 16 8.36 16 9.34 16 8.18 16 7.89 16 5.48 16 4.66
Fibra cruda mg 16 3.97 16 6.46 16 5.82 16 7.57 16 6.60 16 4.73
Fibra soluble 16 32.87 16 38.27 16 43.55 16 44.18 16 21.62 16 22.78
Proteína cruda mg 16 6.57 16 6.97 16 5.12 16 5.24 16 3.13 16 3.57
Lípidos tot FAT mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Ác. grasos
Saturados FASAT mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Monoinsaturados FAMS mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Poliinsaturados FAPU mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Linolénico F18D3N3 mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Eicosapentaenoico F20D5N3 mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Docosahexaenoico F22D6N3 mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
Colesterol CHOLE mg 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00 16 0.00
minerales
Calcio CA mg 16 1.80 16 7.28 16 6.74 16 5.56 16 6.82 16 6.96
Fósforo P mg 16 0.51 16 0.50 16 0.53 16 0.51 16 0.51 16 0.51
Hierro FE mg 16 417.00 16 419.00 16 458.00 16 455.00 16 397.00 16 204.00
Magnesio MG mg 16 1.01 16 1.39 16 1.40 16 1.79 16 1.59 16 1.27
Sodio NA mg 16 15.77 16 3.20 16 3.44 16 2.30 16 14.75 16 12.69
Potasio K mg 16 3.37 16 5.59 16 3.91 16 6.54 16 20.32 16 20.56
Zinc ZN mg 16 11.00 16 32.00 16 50.00 16 11.00 16 9.00 16 30.00
Vitaminas
RAE (vit A) VITA µg - - - - - -
Ác. ascórbico ASCL mg - - - - - -
Tiamina THIA mg - - - - - -
Riboflavina RIBF mg - - - - - -
Niacina NIA mg - - - - - -
Piridoxina VITB6A mg - - - - - -
Ác. fólico FOL µg - - - - - -
Cianocobalamina VITB12 µg - - - - - -
Alimento crudo en peso neto P. comestible - P. comestible - P. comestible - P. comestible - P. comestible - P. comestible -
"""  # raw_text ends here

# ===============================
# Process sections separated by a full blank line
# ===============================
sections = [s.strip() for s in raw_text.split("\n\n") if s.strip()]

# We'll process only sections that contain both "Componente alimentario" and "Nutriente Tagname Unidad"
valid_sections = [sec for sec in sections if ("Componente alimentario" in sec and "Nutriente Tagname Unidad" in sec)]

for sec in valid_sections:
    # Split the section into individual lines.
    sec_lines = sec.splitlines()
    
    # --- Step 1: Ignore the last 4 lines of the section ---
    if len(sec_lines) >= 4:
        sec_lines = sec_lines[:-4]
    
    # --- Step 2: Extract file name from the first line and remove the first two lines ---
    file_name_line = sec_lines[0].strip()  # Title for the output Excel file.
    data_lines = sec_lines[2:]              # Skip the first two lines (title and extra header)
    section_text = "\n".join(data_lines)
    
    # --- Step 3: Replace specific words with their correct forms ---
    replacements = {
        "Fibra dietética": "Fibra_dietética",
        "Hidratos de C": "Hidratos_de_C",
        "Lípidos tot": "Lípidos_tot",
        "RAE (vit A)": "RAE_(vit_A)",
        "Ác. ascórbico": "Ác._ascórbico",
        "Ác. fólico": "Ác._fólico",
        "Fibra cruda": "Fibra_cruda",
        "Fibra soluble": "Fibra_soluble",
        "Proteína cruda": "Proteína_cruda"
    }
    for old, new in replacements.items():
        section_text = section_text.replace(old, new)
    
    # --- Step 4: Modify lines following any line that starts with "Energía ENERC" ---
    sec_lines_mod = section_text.splitlines()
    mod_lines = []
    i = 0
    while i < len(sec_lines_mod):
        curr_line = sec_lines_mod[i]
        mod_lines.append(curr_line)
        if curr_line.strip().startswith("Energía ENERC"):
            if i + 1 < len(sec_lines_mod):
                next_line = sec_lines_mod[i+1]
                if not next_line.strip().startswith("Energía ENERC"):
                    mod_lines.append("Energía ENERC " + next_line)
                    i += 2
                    continue
        i += 1
    section_text = "\n".join(mod_lines)
    
    # --- Step 5: Split section_text into lines ---
    sec_lines = section_text.splitlines()
    
    # --- Step 6: Process header block to extract food names ---
    header_lines = []
    capture = False
    for line in sec_lines:
        if "Componente alimentario" in line:
            capture = True
            # Remove "Componente alimentario" from line if present.
            line = line.replace("Componente alimentario", "").strip()
            if line:
                header_lines.append(line)
            continue
        if "Nutriente" in line:
            break
        if capture and line.strip():
            header_lines.append(line.strip())
    
    food_count = len(header_lines)
    final_headers = []
    for h in header_lines:
        final_headers.append(h + " F")
        final_headers.append(h + " en 100 g")
    
    # --- Step 7: Build the full column headers ---
    full_columns = ["Grupos", "Nutriente", "Tagname", "Unidad"] + final_headers
    
    # --- Step 8: Process nutrient rows ---
    data_rows = []
    expected_pair_count = food_count * 2
    start_index = next(i for i, line in enumerate(sec_lines) if "Nutriente Tagname Unidad" in line) + 1
    group_names_list = ["Elementos principales", "elementos principales", "Ác. grasos", "Minerales", "minerales", "Vitaminas", "vitaminas"]
    current_group = None
    for line in sec_lines[start_index:]:
        line = line.strip()
        if not line:
            continue
        if line in group_names_list:
            current_group = line.title()
            continue
        parts = line.split()
        if len(parts) < 4:
            continue
        # The first three tokens are: Nutriente, Tagname, Unidad.
        nutrient = parts[0]
        tagname = parts[1]
        unit = parts[2]
        # If nutrient starts with a digit or a dot, then it is blank.
        if re.match(r'^[\d\.]', nutrient):
            nutrient = "-"
        # Tagname must be fully uppercase; if not, use "N/A"
        if not re.match(r'^[A-Z]+$', tagname):
            tagname = "N/A"
        # Unidad must start with a lowercase letter or allowed symbol (%, µ, etc.)
        if not re.match(r'^[a-z%µ]', unit):
            unit = "N/A"
        # For specific nutrient names, force tagname to be blank.
        if nutrient in ["Fibra_cruda", "Fibra_soluble", "Proteína_cruda"]:
            tagname = ""
        
        raw_values = parts[3:]
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
        # Pad missing cells with "-" instead of empty string.
        processed_values += ['-'] * (expected_pair_count - len(processed_values))
        
        # Validate each flag–value pair.
        for j in range(0, len(processed_values), 2):
            flag = processed_values[j]
            value = processed_values[j+1]
            # If flag starts with a digit or dot, set to "-"
            if re.match(r'^[\d\.]', flag):
                flag = "-"
            # If value does not match a number (integer or decimal), set to "-"
            if not re.match(r'^\d+(\.\d+)?$', value):
                value = "-"
            processed_values[j] = flag
            processed_values[j+1] = value
        
        row = [current_group, nutrient, tagname, unit] + processed_values
        data_rows.append(row)
    
    # --- Step 9: Create DataFrame and Export ---
    df = pd.DataFrame(data_rows, columns=full_columns)
    output_filename = file_name_line + ".xlsx"
    df.to_excel(output_filename, index=False)
    print("✅ File exported as '{}'".format(output_filename))