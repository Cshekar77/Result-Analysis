import pandas as pd

# -------------------------------
# Step 1: Robust split_marks() for incomplete marks
# -------------------------------
def split_marks(cell):
    val = str(cell).strip()
    if val.lower() == 'absent':
        return None, None
    parts = val.split('+')
    first = int(parts[0].strip()) if parts[0].strip().isdigit() else 0
    second = int(parts[1].strip()) if len(parts) > 1 and parts[1].strip().isdigit() else 0
    return first, second

# -------------------------------
# Step 2: Load Excel
# -------------------------------
file_path = r"Your file path"
raw = pd.read_excel(file_path, header=None)

# -------------------------------
# Step 3: Find subject header row
# -------------------------------
sub_row = None
for r in range(len(raw)):
    if raw.iloc[r].astype(str).str.contains("Sub", na=False).any():
        sub_row = raw.iloc[r]
        break
if sub_row is None:
    raise ValueError("Subject row not found!")

# -------------------------------
# Step 4: Map subjects
# -------------------------------
subjects = {
    "ENG": sub_row[sub_row=="SC2ENG"].index[0],
    "KAN": sub_row[sub_row=="BCAKN2"].index[0],
    "DS": sub_row[sub_row=="BCA21T"].index[0],
    "JAVA": sub_row[sub_row=="BCA22T"].index[0],
    "OS": sub_row[sub_row=="BCA23T"].index[0],
    "DSL": sub_row[sub_row=="BCA21P"].index[0],
    "LNX": sub_row[sub_row=="BCA23P"].index[0],
    "OOPL": sub_row[sub_row=="BCA22P"].index[0],
    "CA": sub_row[sub_row=="SECCA1"].index[0]
}

# -------------------------------
# Step 5: Process all students
# -------------------------------
all_students = []

for i in range(len(raw)):
    row = raw.iloc[i].astype(str)
    
    if row.str.contains("U03", na=False).any():
        usn = row[row.str.contains("U03")].values[0]
        name_row = raw.iloc[i+1]
        name = str(name_row.dropna().values[0])
        th = row
        
        # Practical row (look only next 10 rows)
        pr = None
        for j in range(i+1, min(i+10, len(raw))):
            if raw.iloc[j].astype(str).str.contains("Pr", na=False).any():
                pr = raw.iloc[j]
                break
        if pr is None:
            pr = pd.Series([0]*len(row))
        
        # Result / Overall / SGPA / CGPA rows (look only next 15 rows)
        result_row = pd.Series(["Pass"]*len(row))
        overall_result = "Pass"
        sgpa = None
        cgpa = None
        
        for k in range(i+1, min(i+15, len(raw))):
            row_str = raw.iloc[k].astype(str).str.lower()
            if row_str.str.contains("pass|fail|absent").any():
                result_row = raw.iloc[k]
            if row_str.str.contains("result:").any():
                r_vals = raw.iloc[k].dropna().values
                for val in r_vals:
                    if str(val).lower().startswith("result:"):
                        overall_result = str(val).split(":")[-1].strip().capitalize()
                        break
            if row_str.str.contains("sgpa").any():
                sgpa_vals = [v for v in raw.iloc[k] if str(v).replace('.','',1).isdigit()]
                if sgpa_vals:
                    sgpa = float(sgpa_vals[0])
            if row_str.str.contains("cgpa").any():
                cgpa_vals = [v for v in raw.iloc[k] if str(v).replace('.','',1).isdigit()]
                if cgpa_vals:
                    cgpa = float(cgpa_vals[0])
        
        # Extract marks
        ENG_T, ENG_I = split_marks(th.get(subjects["ENG"], 0))
        KAN_T, KAN_I = split_marks(th.get(subjects["KAN"], 0))
        DS_T, DS_I = split_marks(th.get(subjects["DS"], 0))
        JAVA_T, JAVA_I = split_marks(th.get(subjects["JAVA"], 0))
        OS_T, OS_I = split_marks(th.get(subjects["OS"], 0))
        CA_T, CA_I = split_marks(th.get(subjects["CA"], 0))
        DSL_E, DSL_I = split_marks(pr.get(subjects["DSL"], 0))
        LNX_E, LNX_I = split_marks(pr.get(subjects["LNX"], 0))
        OOPL_E, OOPL_I = split_marks(pr.get(subjects["OOPL"], 0))
        
        # Totals
        ENG_TOTAL = (ENG_T or 0) + (ENG_I or 0)
        KAN_TOTAL = (KAN_T or 0) + (KAN_I or 0)
        DS_TOTAL = (DS_T or 0) + (DS_I or 0)
        JAVA_TOTAL = (JAVA_T or 0) + (JAVA_I or 0)
        OS_TOTAL = (OS_T or 0) + (OS_I or 0)
        CA_TOTAL = (CA_T or 0) + (CA_I or 0)
        DSL_TOTAL = (DSL_E or 0) + (DSL_I or 0)
        LNX_TOTAL = (LNX_E or 0) + (LNX_I or 0)
        OOPL_TOTAL = (OOPL_E or 0) + (OOPL_I or 0)
        
        GRAND_TOTAL = ENG_TOTAL + KAN_TOTAL + DS_TOTAL + JAVA_TOTAL + OS_TOTAL + CA_TOTAL + DSL_TOTAL + LNX_TOTAL + OOPL_TOTAL
        
        # Use M.C.No Result row directly
        ENG_R = result_row.get(subjects["ENG"], "Pass")
        KAN_R = result_row.get(subjects["KAN"], "Pass")
        DS_R = result_row.get(subjects["DS"], "Pass")
        JAVA_R = result_row.get(subjects["JAVA"], "Pass")
        OS_R = result_row.get(subjects["OS"], "Pass")
        CA_R = result_row.get(subjects["CA"], "Pass")
        DSL_R = result_row.get(subjects["DSL"], "Pass")
        LNX_R = result_row.get(subjects["LNX"], "Pass")
        OOPL_R = result_row.get(subjects["OOPL"], "Pass")
        
        # Append
        all_students.append({
            "USN": usn, "NAME": name,
            "ENG(THEORY)": ENG_T, "ENG(INTERNAL)": ENG_I, "ENG(TOTAL)": ENG_TOTAL, "ENG_RESULT": ENG_R,
            "KAN(THEORY)": KAN_T, "KAN(INTERNAL)": KAN_I, "KAN(TOTAL)": KAN_TOTAL, "KAN_RESULT": KAN_R,
            "DS LAB(EXTERNAL)": DSL_E, "DS LAB(INTERNAL)": DSL_I, "DS LAB(TOTAL)": DSL_TOTAL, "DS LAB_RESULT": DSL_R,
            "OOP LAB(EXTERNAL)": OOPL_E, "OOP LAB(INTERNAL)": OOPL_I, "OOP LAB(TOTAL)": OOPL_TOTAL, "OOP LAB_RESULT": OOPL_R,
            "LINUX LAB(EXTERNAL)": LNX_E, "LINUX LAB(INTERNAL)": LNX_I, "LINUX LAB(TOTAL)": LNX_TOTAL, "LINUX LAB_RESULT": LNX_R,
            "CA(THEORY)": CA_T, "CA(INTERNAL)": CA_I, "CA(TOTAL)": CA_TOTAL, "CA_RESULT": CA_R,
            "JAVA(THEORY)": JAVA_T, "JAVA(INTERNAL)": JAVA_I, "JAVA(TOTAL)": JAVA_TOTAL, "JAVA_RESULT": JAVA_R,
            "DS(THEORY)": DS_T, "DS(INTERNAL)": DS_I, "DS(TOTAL)": DS_TOTAL, "DS_RESULT": DS_R,
            "OS(THEORY)": OS_T, "OS(INTERNAL)": OS_I, "OS(TOTAL)": OS_TOTAL, "OS_RESULT": OS_R,
            "GRAND_TOTAL": GRAND_TOTAL,
            "OVERALL_RESULT": overall_result,
            "SGPA": sgpa,
            "CGPA": cgpa
        })

# -------------------------------
# Step 6: Export all students
# -------------------------------
df_students = pd.DataFrame(all_students)
output_path = r"Path of output file"
df_students.to_excel(output_path, index=False)

print(f"✅ Exported all students results to {output_path}")

