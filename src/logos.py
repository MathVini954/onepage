# src/logos.py
from __future__ import annotations
from pathlib import Path

LOGO_EXTS = [".png", ".jpg", ".jpeg", ".webp"]

def _candidates(name: str) -> list[str]:
    n = name.strip()
    return list(dict.fromkeys([
        n,
        n.upper(),
        n.lower(),
        n.replace(" ", "_"),
        n.replace(" ", "-"),
        n.upper().replace(" ", "_"),
        n.upper().replace(" ", "-"),
    ]))

def find_logo_path(sheet_name: str, logos_dir: str = "assets/logos") -> str | None:
    base = Path(logos_dir)
    if not base.exists():
        return None

    for cand in _candidates(sheet_name):
        for ext in LOGO_EXTS:
            p = base / f"{cand}{ext}"
            if p.exists():
                return str(p)
    return None
