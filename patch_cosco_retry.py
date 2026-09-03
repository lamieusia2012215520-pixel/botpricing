import re

with open(r"c:\He_thong_Bot\bot_COSCO.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """# Wrap webdriver.Edge() trong thread timeout d? trnh treo v h?n
_driver_holder = {"driver": None, "error": None}
def _create_session():
    try:
        _driver_holder["driver"] = webdriver.Edge(service=service, options=edge_options)
    except Exception as ex:
        _driver_holder["error"] = ex

_t = threading.Thread(target=_create_session, daemon=True)
_t.start()
_t.join(timeout=30)
if _t.is_alive() or _driver_holder["driver"] is None:
    if _driver_holder["error"]:
        raise _driver_holder["error"]
    raise TimeoutException("webdriver.Edge() treo qu 30s")
driver = _driver_holder["driver"]"""

replacement = """# Wrap webdriver.Edge() trong thread timeout + RETRY d? trnh treo v h?n
driver = None
for attempt in range(1, 4):
    _driver_holder = {"driver": None, "error": None}
    def _create_session():
        try:
            _driver_holder["driver"] = webdriver.Edge(service=service, options=edge_options)
        except Exception as ex:
            _driver_holder["error"] = ex

    _t = threading.Thread(target=_create_session, daemon=True)
    _t.start()
    _t.join(timeout=30)
    
    if _t.is_alive() or _driver_holder["driver"] is None:
        print(f"[WARN] webdriver.Edge() b? treo (attempt {attempt}/3)...")
        if attempt == 3:
            if _driver_holder["error"]:
                raise _driver_holder["error"]
            raise TimeoutException("webdriver.Edge() treo qu 30s sau 3 l?n th?")
        continue # Th? l?i
        
    driver = _driver_holder["driver"]
    break # Thnh cng"""

if target in content:
    content = content.replace(target, replacement)
    with open(r"c:\He_thong_Bot\bot_COSCO.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_COSCO.py with RETRY")
else:
    print("Could not find target string in bot_COSCO.py!")
