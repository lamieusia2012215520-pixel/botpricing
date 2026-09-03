import re

with open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """                  # Transit time từ div[5]
                  try:
                      tt_div   = card.find_element(By.XPATH, "./div[5]")
                      tt_match = re.search(r'(\d+)', tt_div.text.strip())
                      tt_days  = int(tt_match.group(1)) if tt_match else 999
                  except:
                      tt_days = 999"""

replacement = """                  # Transit time - Robust extraction
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

if target in content:
    content = content.replace(target, replacement)
    with open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_HPL.py transit time logic!")
else:
    print("Target not found! Let me try generic replace")
    
    # Try generic replace
    idx = content.find('# Transit time')
    if idx != -1:
        end_idx = content.find('tt_days = 999', idx) + len('tt_days = 999')
        if end_idx != -1:
            content = content[:idx] + replacement + content[end_idx:]
            with open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
                f.write(content)
            print("Successfully patched bot_HPL.py transit time logic via substring match!")
        else:
            print("End not found")
