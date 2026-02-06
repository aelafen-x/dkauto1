import json
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import re

from .points import MODIFIERS, PointsStore

Line = Tuple[int, str]
MULTI_NOT_MARKER = "__multinot__"


@dataclass
class ValidationErrors:
    date_lines: List[int]
    boss_lines: List[int]
    at_lines: List[int]
    single_char_lines: List[int]
    incorrect_not_lines: List[int]
    ambiguous_not_boss_lines: List[int]
    general_lines: List[int]
    unknown_bosses: Dict[str, List[int]]

    def any(self) -> bool:
        return any(
            [
                self.date_lines,
                self.boss_lines,
                self.at_lines,
                self.single_char_lines,
                self.incorrect_not_lines,
                self.ambiguous_not_boss_lines,
                self.general_lines,
            ]
        )


@dataclass
class SanityCheck:
    first_entry: Optional[str]
    last_entry: Optional[str]
    total_lines: int


def _load_boss_aliases(base_dir: Path) -> List[Tuple[str, str]]:
    aliases_path = base_dir / "boss_aliases.json"
    with aliases_path.open("r", encoding="utf-8") as f:
        raw = json.load(f)
    pairs: List[Tuple[str, str]] = []
    for item in raw:
        for key, value in item.items():
            pairs.append((key, value))
    return pairs


def sanitize_line(
    raw_line: str,
    aliases: Optional[List[Tuple[str, str]]] = None,
    base_dir: Optional[Path] = None,
) -> str:
    if aliases is None:
        if base_dir is None:
            raise ValueError("Either aliases or base_dir must be provided.")
        aliases = _load_boss_aliases(base_dir)

    tokens = [
        "".join(ch for ch in token if ch.isascii()).lower()
        for token in raw_line.strip().split()
    ]
    joined = " ".join(tokens)
    if not joined:
        return ""

    updated = joined

    updated = re.sub(r"\(double\s+points?\)", "(doublepoints)", updated)
    updated = re.sub(r"(^|\s)/\s+", r"\1/", updated)
    updated = re.sub(r"\brootx(\d+)\b", r"root\1", updated)
    updated = re.sub(r"/faction\b", "/factions", updated)
    updated = re.sub(r"/nerco\b", "/necro", updated)
    updated = re.sub(r"/hrugn\b", "/hrung", updated)
    updated = re.sub(r"/mordis\b", "/mord", updated)
    updated = re.sub(r"/mords\b", "/mord", updated)
    updated = re.sub(r"\baggy/\s*", "/aggy ", updated)

    alias_map = {original.lower(): replacement.lower() for original, replacement in aliases}
    if ":" in updated:
        prefix, entry = updated.rsplit(":", 1)
        entry = entry.strip()
        if entry:
            parts = entry.split()
            if parts:
                boss = parts[0]
                modifier = ""
                if boss.endswith(")") and "(" in boss:
                    boss, modifier = boss.split("(", 1)
                    modifier = "(" + modifier
                boss = alias_map.get(boss, boss)
                parts[0] = f"{boss}{modifier}"
                entry = " ".join(parts)
                updated = f"{prefix}:{entry}"
    else:
        parts = updated.split()
        if parts:
            boss = parts[0]
            modifier = ""
            if boss.endswith(")") and "(" in boss:
                boss, modifier = boss.split("(", 1)
                modifier = "(" + modifier
            boss = alias_map.get(boss, boss)
            parts[0] = f"{boss}{modifier}"
            updated = " ".join(parts)

    return updated


def preprocess_lines(timers_path: Path, base_dir: Path) -> List[Line]:
    aliases = _load_boss_aliases(base_dir)

    with timers_path.open("r", encoding="utf-8", errors="ignore") as f:
        raw_lines = [line.rstrip("\n") for line in f.readlines()]

    processed: List[Line] = []
    for index, line in enumerate(raw_lines, start=1):
        updated = sanitize_line(line, aliases=aliases)
        if not updated:
            continue
        processed.append((index, updated))

    return processed


def get_date(line: str) -> Optional[datetime]:
    patterns = [
        (r"^(?P<date>\d{1,2} [A-Za-z]{3} \d{4} at \d{2}:\d{2})", "%d %b %Y at %H:%M"),
        (r"^(?P<date>[A-Za-z]{3} \d{1,2}, \d{4} at \d{1,2}:\d{2} [AP]M)", "%b %d, %Y at %I:%M %p"),
        (r"^(?P<date>[A-Za-z]+ \d{1,2}, \d{4} \d{1,2}:\d{2} [AP]M)", "%B %d, %Y %I:%M %p"),
        (r"^(?P<date>[A-Za-z]{3} \d{1,2}, \d{4} \d{1,2}:\d{2} [AP]M)", "%b %d, %Y %I:%M %p"),
    ]

    for pattern, fmt in patterns:
        match = re.match(pattern, line, flags=re.IGNORECASE)
        if not match:
            continue
        date_part = match.group("date").rstrip(":").strip()
        try:
            return datetime.strptime(date_part, fmt)
        except ValueError:
            continue

    return None


def _first_index_of_boss(line: str, bosses: Iterable[str]) -> int:
    min_index = None
    for boss in bosses:
        idx = line.find(boss)
        if idx != -1:
            min_index = idx if min_index is None else min(min_index, idx)
    return min_index if min_index is not None else 0


def slice_by_date(lines: List[Line], start: datetime, end: Optional[datetime] = None) -> List[Line]:
    if end is None:
        end = start + timedelta(days=7)

    def local_tzinfo():
        return datetime.now().astimezone().tzinfo or timezone.utc

    def to_utc(value: datetime) -> datetime:
        if value.tzinfo is None:
            value = value.replace(tzinfo=local_tzinfo())
        return value.astimezone(timezone.utc)

    start_utc = to_utc(start)
    end_utc = to_utc(end)

    start_index = 0
    end_index = 0

    for index, (_, line) in list(enumerate(lines))[::-1]:
        as_date = get_date(line)
        if as_date is None:
            continue
        as_utc = to_utc(as_date)
        if end_index == 0:
            if as_utc <= end_utc:
                end_index = index
        elif as_utc < start_utc:
            start_index = index + 1
            break

    if end_index < start_index:
        return []

    sliced = lines[start_index : end_index + 1]
    return list(sliced)


def validate_lines(
    lines: List[Line],
    points_store: PointsStore,
) -> Tuple[List[Tuple[int, List[str]]], ValidationErrors]:
    error_date_lines: List[int] = []
    error_boss_lines: List[int] = []
    error_at_lines: List[int] = []
    error_single_character_lines: List[int] = []
    incorrect_not_lines: List[int] = []
    ambiguous_not_boss_lines: List[int] = []
    general_error_lines: List[int] = []
    unknown_bosses: Dict[str, List[int]] = {}

    for index, line in lines:
        if get_date(line) is None:
            error_date_lines.append(index)

    boss_lines: List[Tuple[int, List[str]]] = []
    for index, line in lines:
        if ":" not in line:
            general_error_lines.append(index)
            continue
        segment = line.rsplit(":", 1)[1].strip()
        if not segment:
            general_error_lines.append(index)
            continue
        boss_lines.append((index, segment.split()))

    formatted_lines: List[Tuple[int, List[str]]] = []
    for index, tokens in boss_lines:
        full_line = list(tokens)
        if len(full_line) < 2:
            general_error_lines.append(index)
            continue

        modifier = full_line.pop(1)

        is_valid_modifier = False
        for test in MODIFIERS:
            if modifier == f"({test})":
                is_valid_modifier = True
                break

        if is_valid_modifier:
            full_line[0] = f"{full_line[0]}{modifier}"
        else:
            full_line.insert(1, modifier)

        boss = full_line.pop(0)
        allow_multi_not = False
        if MULTI_NOT_MARKER in full_line:
            allow_multi_not = True
            full_line = [t for t in full_line if t != MULTI_NOT_MARKER]

        if boss in {"/legacy", "legacy"} and full_line:
            if re.match(r"^\d+\.[56]$", full_line[0]):
                boss = f"{boss}{full_line.pop(0)}"

        if boss in {"/rings", "rings"} and full_line:
            if re.match(r"^[1-4]x[5-6]$", full_line[0]):
                boss = f"{boss}{full_line.pop(0)}"

        if re.match(r"^\d{3}$", boss) and full_line:
            if full_line[0] in {"4", "5", "6"}:
                candidate = f"{boss}.{full_line[0]}"
                if points_store.get_points(candidate) is not None:
                    boss = candidate
                    full_line.pop(0)

        if re.match(r"^\d{4}\.?$", boss):
            candidate = f"{boss[:3]}.{boss[3]}"
            if points_store.get_points(candidate) is not None:
                boss = candidate

        points = points_store.get_points(boss)
        if points is None:
            if "not" in full_line:
                ambiguous_not_boss_lines.append(index)
                unknown_bosses.setdefault(boss, []).append(index)
                continue
            if full_line:
                alt_boss = full_line[0]
                alt_points = points_store.get_points(alt_boss)
                if alt_points is not None:
                    full_line = [boss] + full_line[1:]
                    formatted_lines.append((index, [alt_boss] + full_line))
                    continue
            unknown_bosses.setdefault(boss, []).append(index)
            error_boss_lines.append(index)
            continue

        if "at" in full_line:
            error_at_lines.append(index)

        if "not" in full_line:
            if allow_multi_not:
                if len(full_line) >= 3 and full_line[1] == "not":
                    pass
                elif len(full_line) >= 2 and full_line[0] == "not":
                    pass
                else:
                    incorrect_not_lines.append(index)
            else:
                if len(full_line) == 3 and full_line[1] == "not":
                    pass
                elif len(full_line) == 2 and full_line[0] == "not":
                    pass
                else:
                    incorrect_not_lines.append(index)

        error_single_character_lines.extend(
            [index for name in full_line if len(name) == 1]
        )

        formatted_lines.append((index, [boss] + full_line))

    errors = ValidationErrors(
        date_lines=error_date_lines,
        boss_lines=error_boss_lines,
        at_lines=error_at_lines,
        single_char_lines=error_single_character_lines,
        incorrect_not_lines=incorrect_not_lines,
        ambiguous_not_boss_lines=ambiguous_not_boss_lines,
        general_lines=general_error_lines,
        unknown_bosses=unknown_bosses,
    )

    return formatted_lines, errors


def build_sanity_check(lines: List[Line]) -> SanityCheck:
    if not lines:
        return SanityCheck(first_entry=None, last_entry=None, total_lines=0)
    return SanityCheck(
        first_entry=lines[0][1],
        last_entry=lines[-1][1],
        total_lines=len(lines),
    )
