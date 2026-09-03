import re

with open(r"c:\He_thong_Bot\main.py", "r", encoding="utf-8") as f:
    content = f.read()

pause_code = """
import psutil
import msvcrt

_is_paused = False
def _pause_listener_thread():
    global _is_paused
    while True:
        try:
            if msvcrt.kbhit():
                key = msvcrt.getch()
                if key.lower() == b'p':
                    _is_paused = not _is_paused
                    try:
                        current_proc = psutil.Process(os.getpid())
                        children = current_proc.children(recursive=True)
                        if _is_paused:
                            print("\\n[PAUSE] === D\u00c3 T\u1ea0M D\u1eeaNG C\u00c1C BOT PYTHON === (Tr\u00ecnh duy\u1ec7t v\u1eabn ho\u1ea1t \u0111\u1ed9ng. B\u1ea5m 'P' \u0111\u1ec3 ti\u1ebfp t\u1ee5c)")
                            for child in children:
                                try:
                                    if "python" in child.name().lower():
                                        child.suspend()
                                except: pass
                        else:
                            print("\\n[RESUME] === D\u00c3 TI\u1ebeP T\u1ee4C CH\u1ea0Y C\u00c1C BOT ===")
                            for child in children:
                                try:
                                    if "python" in child.name().lower():
                                        child.resume()
                                except: pass
                    except Exception as e:
                        print(f"\\n[PAUSE ERROR] L\u1ed7i khi pause/resume: {e}")
        except Exception:
            pass
        time.sleep(0.1)

# ===================================================================================
# ENTRY POINT
# ===================================================================================
if __name__ == "__main__":
    threading.Thread(target=_pause_listener_thread, daemon=True).start()
"""

target = """# ===================================================================================
# ENTRY POINT
# ===================================================================================
if __name__ == "__main__":"""

if target in content:
    content = content.replace(target, pause_code)
    with open(r"c:\He_thong_Bot\main.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully added pause listener to main.py!")
else:
    print("Failed to find entry point in main.py!")
