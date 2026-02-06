import json
from pathlib import Path
from typing import Dict, List


def add_boss_alias(base_dir: Path, alias: str, canonical: str) -> None:
    path = base_dir / "boss_aliases.json"
    alias = alias.lower()
    canonical = canonical.lower()
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
    else:
        data = []

    updated = False
    for item in data:
        if alias in item:
            item[alias] = canonical
            updated = True
            break

    if not updated:
        data.append({alias: canonical})

    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def add_name_alias(base_dir: Path, alias: str, canonical: str) -> None:
    path = base_dir / "name_aliases.json"
    alias = alias.lower()
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
    else:
        data = {}

    data[alias] = canonical
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def add_points_value(base_dir: Path, boss: str, points: int) -> None:
    path = base_dir / "points.json"
    boss = boss.lower()
    if path.exists():
        data = json.loads(path.read_text(encoding="utf-8"))
    else:
        data = {}

    data[boss] = int(points)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
