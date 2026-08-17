#!/usr/bin/env python
"""Build the deployable static site without exposing source PPT files."""
from pathlib import Path
import shutil

BASE = Path(__file__).resolve().parent.parent
PUBLIC = BASE / "public"

if PUBLIC.exists():
    shutil.rmtree(PUBLIC)
(PUBLIC / "data").mkdir(parents=True, exist_ok=True)

shutil.copy2(BASE / "index.html", PUBLIC / "index.html")
for name in ("data.js", "company_info.js", "ipo_details.js"):
    shutil.copy2(BASE / "data" / name, PUBLIC / "data" / name)

print(f"Built deployable site at {PUBLIC}")
