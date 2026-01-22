from __future__ import annotations

from pathlib import Path


def find_logo_path(sheet_name: str, logos_dir: str = "assets/logos") -> str | None:
    """
    Procura uma logo com o MESMO nome da aba (case-insensitive).
    Ex: aba "BOSSA" -> assets/logos/BOSSA.png (ou .jpg/.jpeg/.webp)
    """
    base = Path(logos_dir)
    if not base.exists():
        return None

    target = sheet_name.strip().lower()
    exts = [".png", ".jpg", ".jpeg", ".webp"]

    for p in base.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() not in exts:
            continue
        if p.stem.strip().lower() == target:
            return str(p)

    return None
