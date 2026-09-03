import re

with open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """                # Arrival date từ div[4]"""
replacement = """                print(f"   [Tab{tab_idx}] [DEBUG CARD TEXT]\\n{card.text}")
                # Arrival date từ div[4]"""

if target in content:
    content = content.replace(target, replacement)
    with open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_HPL.py to print card text!")
else:
    print("Target not found!")
