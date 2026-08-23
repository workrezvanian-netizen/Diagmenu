# -*- coding: utf-8 -*-
"""
ساخت / به‌روزرسانی excel_versions.json از روی پوشه‌های محلی Diag_Menu و Diag_Database.

استفاده:
  python make_excel_versions.py
  python make_excel_versions.py --diag "D:\\path\\Diag_Menu" --db "D:\\path\\Diag_Database" --version 1.2.0

خروجی: excel_versions.json در پوشهٔ جاری (آماده برای commit روی گیت‌هاب)
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from datetime import datetime, timezone


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def collect(folder: str, prefix: str) -> dict:
    out = {}
    if not folder or not os.path.isdir(folder):
        return out
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
            continue
        full = os.path.join(folder, name)
        if not os.path.isfile(full):
            continue
        rel = f"{prefix}/{name}".replace("\\", "/")
        out[rel] = {
            "version": None,  # بعداً پر می‌شود
            "sha256": sha256_file(full),
            "size": os.path.getsize(full),
        }
    return out


def main():
    p = argparse.ArgumentParser(description="Generate excel_versions.json for GitHub auto-update")
    p.add_argument("--diag", default="Diag_Menu", help="مسیر پوشهٔ Diag_Menu")
    p.add_argument("--db", default="Diag_Database", help="مسیر پوشهٔ Diag_Database")
    p.add_argument("--version", default="1.0.0", help="شماره نسخهٔ این انتشار")
    p.add_argument("-o", "--output", default="excel_versions.json", help="نام فایل خروجی")
    args = p.parse_args()

    files = {}
    files.update(collect(args.diag, "Diag_Menu"))
    files.update(collect(args.db, "Diag_Database"))

    for meta in files.values():
        meta["version"] = args.version

    manifest = {
        "version": args.version,
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "files": files,
    }

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"Wrote {args.output} with {len(files)} file(s), version={args.version}")
    for rel in files:
        print(f"  - {rel}")


if __name__ == "__main__":
    main()
