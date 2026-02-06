import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Union

MODIFIERS = ["brucybonus", "double", "doublepoints", "fail", "comp"]


@dataclass
class Tier:
    level: int
    point_5: int
    point_6: int


PointValue = Union[int, Dict[str, int], List[Tier]]


class PointsStore:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = base_dir
        self.points_map: Dict[str, PointValue] = {}
        self.prios: List[str] = []
        self._bosses: List[str] = []
        self._bosses_re: Optional[re.Pattern] = None
        self._load()

    @property
    def bosses(self) -> List[str]:
        return self._bosses

    def _load(self) -> None:
        points_path = self.base_dir / "points.json"
        prios_path = self.base_dir / "prios.json"

        with points_path.open("r", encoding="utf-8") as f:
            raw = json.load(f)

        points_map: Dict[str, PointValue] = {}
        for key, value in raw.items():
            if isinstance(value, int):
                points_map[key] = value
            elif isinstance(value, dict):
                points_map[key] = {str(k): int(v) for k, v in value.items()}
            elif isinstance(value, list):
                tiers = []
                for item in value:
                    tiers.append(
                        Tier(
                            level=int(item["level"]),
                            point_5=int(item["5"]),
                            point_6=int(item["6"]),
                        )
                    )
                points_map[key] = tiers
            else:
                raise ValueError(f"Unknown point value type for {key}")

        with prios_path.open("r", encoding="utf-8") as f:
            prios = json.load(f)

        self.points_map = points_map
        self.prios = [str(p) for p in prios]
        self._bosses = list(points_map.keys())
        self._bosses_re = self._build_bosses_re()

    def _build_bosses_re(self) -> re.Pattern:
        ring_match = r"/?rings([1-4])x([5-6])"
        legacy_match = r"/?legacy(\d+)\.([5-6])"
        root_match = r"/?root\d*"

        boss_values: List[str] = []
        for boss, value in self.points_map.items():
            if isinstance(value, int):
                if re.match(r"^\d+\.\d+$", boss):
                    boss_values.append(r"/?" + re.escape(boss))
                elif boss.startswith("/"):
                    boss_values.append(r"/?" + re.escape(boss.lstrip("/")))
                else:
                    boss_values.append(re.escape(boss))

        boss_values.append(ring_match)
        boss_values.append(legacy_match)
        boss_values.append(root_match)

        boss_union = "|".join(boss_values)
        modifier_union = "|".join(MODIFIERS)
        pattern = rf"^(?P<boss>{boss_union})(\((?P<modifier>{modifier_union})\))?$"
        return re.compile(pattern)

    def get_points(self, boss: str) -> Optional[int]:
        if self._bosses_re is None:
            self._bosses_re = self._build_bosses_re()

        rings_capture = re.compile(r"^/?rings(?P<num>[1-4])x(?P<star>[5-6])$")
        legacy_capture = re.compile(r"^/?legacy(?P<level>\d+)\.(?P<star>[5-6])$")
        root_re = re.compile(r"^/?root\d*$")

        match = self._bosses_re.match(boss)
        if not match:
            return None

        points = 0
        double_points = False
        half_points = False

        stripped_boss = match.group("boss")
        if stripped_boss not in self.points_map:
            if f"/{stripped_boss}" in self.points_map:
                stripped_boss = f"/{stripped_boss}"
            elif stripped_boss.startswith("/") and stripped_boss[1:] in self.points_map:
                stripped_boss = stripped_boss[1:]
        modifier = match.group("modifier")

        if modifier:
            if modifier == "brucybonus":
                points += 5
            elif modifier in {"double", "doublepoints"}:
                double_points = True
            elif modifier == "fail":
                half_points = True
            elif modifier == "comp":
                if not any(prio in boss for prio in self.prios):
                    return None

        if stripped_boss in self.points_map:
            value = self.points_map[stripped_boss]
            if isinstance(value, int):
                points += value
            else:
                return None
        else:
            ring_match = rings_capture.match(stripped_boss)
            if ring_match:
                num = int(ring_match.group("num"))
                star = ring_match.group("star")
                ring_map = self.points_map.get("/rings")
                if ring_map is None:
                    ring_map = self.points_map.get("rings")
                if isinstance(ring_map, dict):
                    points += int(ring_map[star]) * num
            else:
                legacy_match = legacy_capture.match(stripped_boss)
                if legacy_match:
                    level = int(legacy_match.group("level"))
                    star = legacy_match.group("star")
                    tiers = self.points_map.get("/legacy")
                    if tiers is None:
                        tiers = self.points_map.get("legacy")
                    if isinstance(tiers, list):
                        for tier in tiers:
                            if level >= tier.level:
                                points += tier.point_5 if star == "5" else tier.point_6
                                break
                elif root_re.match(stripped_boss):
                    points += 4
                else:
                    return None

        if double_points:
            points *= 2
        if half_points:
            points = (points + 1) // 2

        return points
