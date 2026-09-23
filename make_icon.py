import base64, re
from pathlib import Path
from PIL import Image

src = Path("GameVault_app.py")
text = src.read_text(encoding="utf-8")
m = re.search(r'_ICON_PNG_B64\s*=\s*"([^"]+)"', text)
if not m:
    raise SystemExit("Could not find _ICON_PNG_B64 in GameVault_app.py")
raw = base64.b64decode(m.group(1))
Path("icon.png").write_bytes(raw)
img = Image.open(Path("icon.png")).convert("RGBA")
img.save("icon.ico", format="ICO", sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])
print("Created icon.png and icon.ico")
