import xlrd
import json

file_path = "PV_S3.xls"
matricule_to_find = "23068"

workbook = xlrd.open_workbook(file_path)
sheet = workbook.sheet_by_index(0)

# Extract headers (matières names)
headers = []
for cell in sheet.row_values(4):
    if cell != '':
        headers.append(cell)

# Actual structure of your modules (number of matieres per module)
modules_structure = [4, 2, 3, 4]

# Convert strings like '12,34' to float
def parse_float(value):
    try:
        if isinstance(value, float):
            return value
        return float(str(value).replace(',', '.'))
    except:
        return None

result_found = False

for row_idx in range(6, sheet.nrows):
    row = sheet.row_values(row_idx)

    # Safely extract the student's ID
    try:
        stud_id = str(int(float(row[1]))).strip()
    except (ValueError, TypeError):
        continue

    if stud_id == matricule_to_find:
        result_found = True
        raw = row[4:]  # Skip first 4 columns (meta info)

        matieres_dict = {}
        modules_dict = {}
        i = 0
        m_idx = 0
        mod_num =1
        mod_name = f"Module {mod_num}"
        modules_dict[mod_name] = {'matieres': [], 'moyenne': None, 'decision': None}

        for _ in range(13):
                if m_idx == 13:
                    break
                print(m_idx)
                name = headers[m_idx]
                d = parse_float(raw[i])
                sn = parse_float(raw[i + 1])
                sr = parse_float(raw[i + 2])
                m = parse_float(raw[i + 3])
                dec = raw[i + 4]

                matieres_dict[name] = {
                    'devoir': d,
                    'exam_sn': sn,
                    'exam_sr': sr,
                    'moyenne': m,
                    'decision': dec
                }
                modules_dict[mod_name]['matieres'].append(name)

                i += 5
                m_idx += 1
                element = raw[i + 1]
                if element in ['V','NV','VC']:
                    modules_dict[mod_name]['moyenne'] = parse_float(raw[i])
                    modules_dict[mod_name]['decision'] = raw[i + 1]
                    i += 2
                    mod_num += 1
                    if(m_idx != 13):
                        mod_name = f"Module {mod_num}"
                        modules_dict[mod_name] = {'matieres': [], 'moyenne': None, 'decision': None}



            # Add module average and decision


        # Extract semester summary
        semestre = {
            'moyenne_total': parse_float(raw[i]),
            'credit_total': int(raw[i + 1]) if str(raw[i + 1]).isdigit() else raw[i + 1],
            'decision': raw[i + 2]
        }
        modules_dict['Semestre'] = semestre

        # Print results
        print(f"Résultats pour matricule {matricule_to_find}:\n")
        print("Matières:")
        print(json.dumps(matieres_dict, indent=4, ensure_ascii=False))
        print("\nModules:")
        print(json.dumps(modules_dict, indent=4, ensure_ascii=False))
        break

if not result_found:
    print(f"⚠️ Matricule {matricule_to_find} non trouvé.")

