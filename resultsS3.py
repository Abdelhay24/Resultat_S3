import xlrd
import json
matricule = input("matriculak chenhou ?   ")
DSI_file ="PV_S3_DSI.xls"
RSS_file = "PV_S3_RSS.xls"
CNM_file="PV_S3_DWM.xls"
S1_file = "PV_S1.xls"
option = input("choose an option and type it ? : ['CNM_S3','DSI_S3','RSS_S3','S1']   :    ")
matricule_to_find = matricule
if (option=="CNM_S3"):
    file_path = CNM_file
elif(option=="DSI_S3"):
    file_path =DSI_file
elif(option=="RSS_S3"):
    file_path = RSS_file
elif(option=="S1") :
    file_path = S1_file
else:
    print("ta5assouss mahou 5aleg")
    exit()
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

# Define exam requirements for each subject

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
        stud_name = row[2] + " "+ row[3]
        raw = row[4:]  # Skip first 4 columns (meta info)

        matieres_dict = {}
        modules_dict = {}
        i = 0
        m_idx = 0
        mod_num =1
        mod_name = f"Module {mod_num}"
        modules_dict[mod_name] = {'matieres': [], 'moyenne': None, 'decision': None}

        for _ in range(13):
                if (option == "S1") and (m_idx==11) :
                    break
                elif m_idx == 13:
                    break
                name = headers[m_idx]
                d = parse_float(raw[i])
                sn = parse_float(raw[i + 1])
                sr = parse_float(raw[i + 2])
                m = parse_float(raw[i + 3])
                dec = raw[i + 4]

                # Calculate weighted average and points needed
                if d is not None and sn is not None:
                    weighted_avg = (d * 0.4) + (sn * 0.6)
                    points_needed = round(max(0, (10 - weighted_avg) * 1.66666667), 2)
                else:
                    points_needed = None

                matieres_dict[name] = {
                    'devoir': d,
                    'exam_sn': sn,
                    'exam_sr': sr,
                    'moyenne': m,
                    'decision': dec,
                    'Les points que tu dois ajouter lors du rattrapage pour avoir une moyenne de 10': points_needed,
                    'status': 'Validé' if m and m >= 10 else 'Non Validé'
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

        # Extract semester summary
        semestre = {
            'moyenne_total': parse_float(raw[i]),
            'credit_total': int(raw[i + 1]) if str(raw[i + 1]).isdigit() else raw[i + 1],
            'decision': raw[i + 2]
        }
        modules_dict['Semestre'] = semestre

        # Extract non validated matieres
        non_validees = [matiere for matiere, data in matieres_dict.items() if data['status'] == 'Non Validé']

        # Display results
        print(f"\n")
        print(f"Résultats pour {stud_name} de matricule {matricule_to_find}:\n")
        print("Matières:")
        print(json.dumps(matieres_dict, indent=4, ensure_ascii=False))
        print("\nModules:")
        print(json.dumps(modules_dict, indent=4, ensure_ascii=False))
        print("\nLes matières non validées:")
        print(json.dumps(non_validees, indent=4, ensure_ascii=False))

if not result_found:
    print(f"⚠️ Matricule {matricule_to_find} non trouvé.")

