import re

with open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
    content = f.read()

replacement = """                # Transit time - Robust extraction
                try:
                    tt_days = 999
                    card_text = card.text.strip()
                    tt_match = re.search(r'\\b(\\d+)\\s*[Dd]ays?\\b', card_text, re.IGNORECASE)
                    if tt_match:
                        tt_days = int(tt_match.group(1))
                    else:
                        for div_idx in (4, 5, 6, 7):
                            try:
                                div_text = card.find_element(By.XPATH, f"./div[{div_idx}]").text.strip()
                                if re.search(r'202\\d', div_text) or re.search(r'^[a-zA-Z]+', div_text): 
                                    continue
                                m = re.search(r'^(\\d+)$', div_text)
                                if m:
                                    val = int(m.group(1))
                                    if val < 2000:
                                        tt_days = val
                                        break
                            except: pass
                except: tt_days = 999"""

content = re.sub(r' +(?=# Transit time - Robust extraction).+?except: tt_days = 999', replacement, content, flags=re.DOTALL)

with open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
    f.write(content)
print("Successfully fixed indentation!")
