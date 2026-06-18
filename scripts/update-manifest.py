#!/usr/bin/env python3
"""Scan image folders and update carousel manifests + JS image lists."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

FOLDERS = [
    {
        "img_dir": "images/crmef",
        "manifest": "images/crmef/manifest.json",
        "js_out": "images/crmef/images.js",
        "var": "CRMEF_IMAGES",
    },
    {
        "img_dir": "images/moments",
        "manifest": "images/moments/manifest.json",
        "js_out": "images/moments/images.js",
        "var": "MOMENTS_IMAGES",
    },
    {
        "img_dir": "images/msp",
        "manifest": "images/msp/manifest.json",
        "js_out": "images/msp/images.js",
        "var": "MSP_IMAGES",
    },
]

for cfg in FOLDERS:
    dir_path = ROOT / cfg["img_dir"]
    if not dir_path.is_dir():
        print(f"SKIP {cfg['img_dir']} — folder not found")
        continue

    images = sorted(
        str(p.relative_to(ROOT))
        for p in sorted(dir_path.iterdir())
        if p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp", ".gif")
    )

    # JSON manifest (for reference)
    (ROOT / cfg["manifest"]).write_text(
        json.dumps({"images": images}, indent=2), encoding="utf-8"
    )

    # JS file with array
    js_content = "var " + cfg["var"] + " = " + json.dumps(images) + ";\n"
    (ROOT / cfg["js_out"]).write_text(js_content, encoding="utf-8")

    print(f"✓ {cfg['img_dir']} → {len(images)} images")

print("Done.")
