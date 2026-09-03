import re

with open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
    content = f.read()

replacement = r"""        # PASS 2: Cong Freight Charges/Freight Surcharges
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
                    charge_name = full_charge_name.split("\n")[0]
                    unit_text   = tds[unit_idx].text.strip() if unit_idx >= 0 else ""
                    currency    = tds[curr_idx].text.strip() if curr_idx >= 0 else ""
                    
                    charge_scope = f"{section} {full_charge_name} {unit_text}".upper()
                    is_freight_charge = "FREIGHT" in section.upper()"""

pattern = r"# PASS 2.*?is_freight_charge = \"FREIGHT\" in section\.upper\(\)"

# Replace using string find/replace to avoid regex escape issues with the replacement
idx1 = content.find("# PASS 2")
idx2 = content.find('is_freight_charge = "FREIGHT" in section.upper()') + len('is_freight_charge = "FREIGHT" in section.upper()')
if idx1 != -1 and idx2 != -1:
    new_content = content[:idx1] + replacement + content[idx2:]
    with open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
        f.write(new_content)
    print("Patched successfully!")
else:
    print("Could not find boundaries")
