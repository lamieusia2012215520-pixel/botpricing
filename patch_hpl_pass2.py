import re

with open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """        # PASS 2: CTng Freight Charges/Freight Surcharges vA m?i charge cA3 nhAn Rail.
        # Rail vn lA mTt ph n c a c>c v-n chuyn dA1 HPL `t nA3 trong Destination section.
        for table in all_tables:
            try:
                all_ths  = table.find_elements(By.CSS_SELECTOR, "thead th")
                th_texts = [th.text.strip() for th in all_ths]
                section  = th_texts[0] if th_texts else ""
            except:
                continue

            col_map = {}
            for i, txt in enumerate(th_texts):
                if txt in ("20STD", "40STD", "40HC"):
                    col_map[i] = txt
                    avail_cols.add(txt)

            for row in table.find_elements(By.CSS_SELECTOR, "tbody tr"):
                try:
                    tds = row.find_elements(By.CSS_SELECTOR, "td")
                    if len(tds) < 2: continue
                    full_charge_name = tds[0].text.strip()
                    charge_name = full_charge_name.split("\\n")[0]
                    unit_text   = tds[1].text.strip()
                    currency    = tds[2].text.strip()
                    charge_scope = f"{section} {full_charge_name} {unit_text}".upper()
                    is_freight_charge = "FREIGHT" in section.upper()"""

replacement = """        # PASS 2: Cong Freight Charges/Freight Surcharges
        for table in all_tables:
            try:
                all_ths  = table.find_elements(By.CSS_SELECTOR, "thead th")
                th_texts = [th.text.strip() for th in all_ths]
                # Find section by looking at the first non-empty text before 'Unit'
                section = ""
                for t in th_texts:
                    if t.upper() in ("UNIT", "CURR.", "CURR"): break
                    if t: section = t
            except:
                continue
                
            unit_idx = -1
            curr_idx = -1
            for i, txt in enumerate(th_texts):
                txt_up = txt.upper()
                if txt_up == "UNIT": unit_idx = i
                elif txt_up in ("CURR.", "CURR"): curr_idx = i

            charge_idx = unit_idx - 1 if unit_idx > 0 else 0

            col_map = {}
            for i, txt in enumerate(th_texts):
                if txt in ("20STD", "40STD", "40HC"):
                    col_map[i] = txt
                    avail_cols.add(txt)

            for row in table.find_elements(By.CSS_SELECTOR, "tbody tr"):
                try:
                    tds = row.find_elements(By.CSS_SELECTOR, "td")
                    if len(tds) <= max(charge_idx, unit_idx, curr_idx): continue
                    
                    full_charge_name = tds[charge_idx].text.strip()
                    charge_name = full_charge_name.split("\\n")[0]
                    unit_text   = tds[unit_idx].text.strip() if unit_idx >= 0 else ""
                    currency    = tds[curr_idx].text.strip() if curr_idx >= 0 else ""
                    
                    charge_scope = f"{section} {full_charge_name} {unit_text}".upper()
                    is_freight_charge = "FREIGHT" in section.upper()"""

# I need to handle encoding issues and replacing the text safely
import codecs
try:
    with codecs.open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # We will use regex to replace to avoid encoding mismatches in the target string
    pattern = r"# PASS 2.*?is_freight_charge = \"FREIGHT\" in section\.upper\(\)"
    
    new_content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    
    with codecs.open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
        f.write(new_content)
    print("Patched successfully!")
except Exception as e:
    print(f"Error: {e}")
