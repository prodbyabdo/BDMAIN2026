import csv
from datetime import datetime

input_file = r"c:\Users\ben.arthur\Desktop\BD MAIN 2026\Provider of Services File - Clinical Laboratories\Provider of Services File - Clinical Laboratories\2025-Q4\CLIA.DATA.Q4_2025.csv"
output_file = r"c:\Users\ben.arthur\Desktop\BD MAIN 2026\Provider of Services File - Clinical Laboratories\Provider of Services File - Clinical Laboratories\2025-Q4\filtered_clia.csv"

today_str = datetime.now().strftime("%Y%m%d")

print("Filtering CSV...")

target_codes = {'400', '110', '900', '500'}
keywords = ['GENOMICS', 'MOLECULAR', 'GENETICS', 'SEQUENCING', 'PATHOLOGY', 'PRECISION', 'IMMUNOLOGY']
valid_types = {'1', '3', '9'}

count_in = 0
count_out = 0

with open(input_file, mode='r', encoding='utf-8-sig') as fin, open(output_file, mode='w', encoding='utf-8', newline='') as fout:
    reader = csv.reader(fin)
    writer = csv.writer(fout)
    
    headers = next(reader)
    writer.writerow(headers)
    
    # Map column indices safely
    try:
        type_cd_idx = headers.index('CRTFCT_TYPE_CD')
        trm_cd_idx = headers.index('CLIA_TRMNTN_CD')
        exprtn_dt_idx = headers.index('TRMNTN_EXPRTN_DT')
        fac_name_idx = headers.index('FAC_NAME')
        cap_idx = headers.index('CAP_ACRDTD_Y_MATCH_SW')
        cola_idx = headers.index('COLA_ACRDTD_Y_MATCH_SW')
        jcaho_idx = headers.index('JCAHO_ACRDTD_Y_MATCH_SW')
        
        spec_indices = [headers.index(f'CLIA_LAB_CLASSIFICATION_CD_{i}') for i in range(1, 11)]
    except ValueError as e:
        print(f"Missing column: {e}")
        exit(1)

    for row in reader:
        count_in += 1
        if len(row) <= exprtn_dt_idx:
            continue
            
        # 1. Certificate Type
        if row[type_cd_idx] not in valid_types: continue
        
        # 4. Active Status
        if row[trm_cd_idx] != '00': continue
        
        # 6. Future Termination Date (format YYYYMMDD)
        dt = row[exprtn_dt_idx].strip()
        if len(dt) != 8 or not dt.isdigit() or dt <= today_str: continue
        
        # 2 & 5. Specialty OR Keyword
        has_spec = False
        for idx in spec_indices:
            if idx < len(row) and row[idx].strip() in target_codes:
                has_spec = True
                break
                
        has_kw = False
        fac_name = row[fac_name_idx].upper()
        for kw in keywords:
            if kw in fac_name:
                has_kw = True
                break
                
        if not (has_spec or has_kw): continue
        
        # 3. Accreditation (CAP or COLA or JCAHO)
        is_cap = (row[cap_idx] == 'Y') if cap_idx < len(row) else False
        is_cola = (row[cola_idx] == 'Y') if cola_idx < len(row) else False
        is_jcaho = (row[jcaho_idx] == 'Y') if jcaho_idx < len(row) else False
        
        if not (is_cap or is_cola or is_jcaho): continue
        
        writer.writerow(row)
        count_out += 1

print(f"Processed {count_in} records. Kept {count_out} records.")
