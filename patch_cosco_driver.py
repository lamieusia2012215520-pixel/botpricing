import re

with open(r"c:\He_thong_Bot\bot_COSCO.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """edge_options = Options()
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9523")
service = Service(executable_path=driver_path)
driver = webdriver.Edge(service=service, options=edge_options)"""

replacement = """import threading
from selenium.common.exceptions import TimeoutException

edge_options = Options()
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9523")
service = Service(executable_path=driver_path)

# Wrap webdriver.Edge() trong thread timeout d? trnh treo v h?n
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

if target in content:
    content = content.replace(target, replacement)
    with open(r"c:\He_thong_Bot\bot_COSCO.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_COSCO.py")
else:
    print("Could not find the target string in bot_COSCO.py!")
