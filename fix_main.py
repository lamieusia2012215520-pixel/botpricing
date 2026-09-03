import re

with open(r"c:\He_thong_Bot\main.py", "r", encoding="utf-8") as f:
    content = f.read()

# Restore the missing block
bad_content = '''    "ZIM":          "bot_zim.py",
    "bot_KMTC.py":     "KMTC",
    "bot_msc.py":      "MSC",'''

fixed_content = '''    "ZIM":          "bot_zim.py",
    "ZIM LINE":     "bot_zim.py",
    "ZIM LINES":    "bot_zim.py",
}

# Bot short names (for display & temp file naming)
BOT_DISPLAY = {
    "bot_cma.py":      "CMA / ANL / CNC / APL",
    "bot_COSCO.py":    "COSCO",
    "bot_cul.py":      "CUL",
    "bot_EMC.py":      "EMC",
    "bot_HPL.py":      "HAPAG LLOYD",
    "bot_hmm.py":      "HMM",
    "bot_KMTC.py":     "KMTC",
    "bot_msc.py":      "MSC",'''

content = content.replace(bad_content, fixed_content)

with open(r"c:\He_thong_Bot\main.py", "w", encoding="utf-8") as f:
    f.write(content)
