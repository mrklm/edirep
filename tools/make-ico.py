from pathlib import Path
from PIL import Image

SRC = Path("assets/logo.png")
DST = Path("assets/logo.ico")

if not SRC.exists():
    raise FileNotFoundError(f"Logo source introuvable: {SRC}")

img = Image.open(SRC).convert("RGBA")

# Tailles standards Windows (explorateur, taskbar, raccourcis, etc.)
sizes = [
    (16, 16),
    (24, 24),
    (32, 32),
    (48, 48),
    (64, 64),
    (128, 128),
    (256, 256),
]

img.save(DST, format="ICO", sizes=sizes)

print(f"Icône générée avec succès : {DST}")
