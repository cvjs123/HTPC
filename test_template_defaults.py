import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location("web_app_module", "web_app.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

base = Path(__file__).resolve().parent
assert hasattr(module, "resolve_template_path"), "resolve_template_path chưa tồn tại"

for name in ["Tong.xlsx", "M1.xlsx", "M2.xlsx", "PL.xlsx", "SPC1.xlsx", "SPCTH.xlsx", "M1TH.xlsx", "XMTH.xlsx"]:
    path = module.resolve_template_path(name)
    assert path is not None, f"Không tìm thấy template mặc định: {name}"
    assert path.exists(), f"File mẫu không tồn tại: {path}"

print("template defaults OK")
