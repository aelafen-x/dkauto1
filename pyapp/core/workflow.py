import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Set, Tuple

from .autocorrect import Autocorrecter
from .points import PointsStore
from .sanitise import (
    SanityCheck,
    ValidationErrors,
    build_sanity_check,
    get_date,
    preprocess_lines,
    slice_by_date,
    validate_lines,
    MULTI_NOT_MARKER,
)
from .aliases import add_name_alias
from .sheets import get_names_from_sheets


@dataclass
class Resolution:
    names: List[str]
    cache_original: bool
    add_new: bool = False
    persist_alias: bool = False
    merge_with_prev: bool = False
    merge_with_next: bool = False
    reprocess: bool = False


@dataclass
class CalculationResult:
    totals: List[Tuple[str, int]]
    sanity: SanityCheck
    errors: ValidationErrors
    boss_counts: Dict[str, Dict[str, int]]
    boss_list: List[str]
    events: List["EventRecord"]


@dataclass
class EventEntry:
    name: str
    delta: int


@dataclass
class EventRecord:
    event_time: datetime
    boss: str
    points: int
    entries: List[EventEntry]
    source_line: str


ResolveCallback = Callable[
    [str, List[str], str, str, str, str, str], Optional[Resolution]
]


def build_aliases(names: Iterable[str], base_dir: Path) -> Dict[str, str]:
    aliases_path = base_dir / "name_aliases.json"
    with aliases_path.open("r", encoding="utf-8") as f:
        aliases: Dict[str, str] = {k.lower(): v for k, v in (json.load(f)).items()}

    for name in names:
        if " " in name:
            parts = name.split()
            aliases[parts[0].lower()] = name
            joined = "".join(parts).lower()
            aliases[joined] = name
            aliases[joined.rstrip("0123456789")] = name
        else:
            aliases[name.lower()] = name
            if not name.isnumeric():
                aliases[name.rstrip("0123456789").lower()] = name

    aliases["nekotin"] = "NEKOTIN"
    aliases["nekotin2"] = "NEKOTIN2"
    return aliases


def calculate_points(
    timers_path: Path,
    start_date: Optional[datetime],
    end_date: Optional[datetime],
    use_all_entries: bool,
    spreadsheet_id: str,
    range_name: str,
    credentials_path: Path,
    token_path: Path,
    base_dir: Path,
    resolve_unknown: ResolveCallback,
) -> CalculationResult:
    strict_prefix = "__strict__"
    points_store = PointsStore(base_dir)

    def normalize_boss_key(raw_boss: str) -> str:
        cleaned = raw_boss.strip()
        if "(" in cleaned and cleaned.endswith(")"):
            cleaned = cleaned.split("(", 1)[0]
        if cleaned.startswith("/"):
            cleaned = cleaned[1:]
        return cleaned

    lines = preprocess_lines(timers_path, base_dir)
    if not use_all_entries and start_date and end_date:
        lines = slice_by_date(lines, start_date, end_date)

    sanity = build_sanity_check(lines)

    line_map = {idx: line for idx, line in lines}
    formatted_lines, errors = validate_lines(lines, points_store)
    if errors.any():
        return CalculationResult(
            totals=[],
            sanity=sanity,
            errors=errors,
            boss_counts={},
            boss_list=[],
            events=[],
        )

    names = get_names_from_sheets(
        spreadsheet_id=spreadsheet_id,
        range_name=range_name,
        credentials_path=credentials_path,
        token_path=token_path,
    )

    aliases = build_aliases(names, base_dir)
    sheet_lookup = {name.lower(): name for name in names}
    autocorrecter = Autocorrecter(names)
    discard: Set[str] = set()

    dkp_count: Dict[str, int] = {}
    boss_counts: Dict[str, Dict[str, int]] = {}
    boss_set: Set[str] = set()
    events: List[EventRecord] = []

    for line_index, tokens in formatted_lines:
        if not tokens:
            continue
        boss = tokens[0]
        points = points_store.get_points(boss)
        if points is None:
            errors.boss_lines.append(line_index)
            continue
        boss_key = normalize_boss_key(boss)

        actual_names: List[str] = []

        line_text_raw = line_map.get(line_index, "")
        line_prefix = ""
        if ":" in line_text_raw:
            line_prefix = line_text_raw.rsplit(":", 1)[0]

        name_tokens = list(tokens[1:])
        i = 0
        last_appended_token: Optional[str] = None

        while i < len(name_tokens):
            raw_token = name_tokens[i]
            strict = False
            if raw_token.startswith(strict_prefix):
                strict = True
                name = raw_token[len(strict_prefix) :]
            else:
                name = raw_token
            if name == MULTI_NOT_MARKER:
                i += 1
                continue
            if name == "not":
                actual_names.append("not")
                last_appended_token = None
                i += 1
                continue

            if name in discard or len(name) <= 1:
                i += 1
                continue

            actual_name = aliases.get(name)
            if actual_name:
                actual_names.append(actual_name)
                last_appended_token = name
                i += 1
                continue
            if strict and name in sheet_lookup:
                actual_names.append(sheet_lookup[name])
                last_appended_token = name
                i += 1
                continue

            suggestions = autocorrecter.correct(name)

            prev_token_raw = name_tokens[i - 1] if i - 1 >= 0 else ""
            next_token_raw = name_tokens[i + 1] if i + 1 < len(name_tokens) else ""
            if prev_token_raw.startswith(strict_prefix):
                prev_token_raw = prev_token_raw[len(strict_prefix) :]
            if next_token_raw.startswith(strict_prefix):
                next_token_raw = next_token_raw[len(strict_prefix) :]

            prev_token = ""
            if (
                prev_token_raw
                and prev_token_raw not in {"not", MULTI_NOT_MARKER}
                and len(prev_token_raw) > 1
                and prev_token_raw not in discard
                and last_appended_token == prev_token_raw
            ):
                prev_token = prev_token_raw

            next_token = ""
            if (
                next_token_raw
                and next_token_raw not in {"not", MULTI_NOT_MARKER}
                and len(next_token_raw) > 1
                and next_token_raw not in discard
            ):
                next_token = next_token_raw

            display_tokens = []
            for token in name_tokens:
                if token.startswith(strict_prefix):
                    display_tokens.append(token[len(strict_prefix) :])
                else:
                    display_tokens.append(token)
            entry = " ".join([boss] + display_tokens)
            line_text = f"{line_prefix}:{entry}" if line_prefix else entry

            prev_line_raw = line_map.get(line_index - 1, "")
            next_line_raw = line_map.get(line_index + 1, "")
            resolution = resolve_unknown(
                name,
                suggestions,
                line_text,
                prev_token,
                next_token,
                prev_line_raw,
                next_line_raw,
            )
            if resolution is None:
                discard.add(name)
                i += 1
                continue

            if resolution.reprocess:
                replacement_tokens = [
                    f"{strict_prefix}{n.lower()}" for n in resolution.names if n
                ]
                if not replacement_tokens:
                    discard.add(name)
                    i += 1
                    continue

                if resolution.merge_with_prev and prev_token and i - 1 >= 0:
                    if actual_names:
                        actual_names.pop()
                    start = i - 1
                    name_tokens[start : i + 1] = replacement_tokens
                    i = max(start, 0)
                    last_appended_token = None
                    continue

                if resolution.merge_with_next and next_token and i + 1 < len(name_tokens):
                    name_tokens[i : i + 2] = replacement_tokens
                    last_appended_token = None
                    continue

                name_tokens[i : i + 1] = replacement_tokens
                last_appended_token = None
                continue

            if resolution.merge_with_prev and prev_token:
                if actual_names:
                    actual_names.pop()

            if resolution.add_new:
                new_name = resolution.names[0]
                aliases[name] = new_name
                aliases[new_name.lower()] = new_name
                autocorrecter.add_word(new_name)
                actual_names.append(new_name)
                if resolution.merge_with_prev and prev_token:
                    last_appended_token = f"{prev_token}{name}"
                elif resolution.merge_with_next and next_token:
                    last_appended_token = f"{name}{next_token}"
                else:
                    last_appended_token = name
                if resolution.persist_alias:
                    add_name_alias(base_dir, name, new_name)
                if resolution.merge_with_next and next_token:
                    i += 2
                else:
                    i += 1
                continue

            resolved_names: List[str] = []
            for correction in resolution.names:
                correction_key = correction.lower()
                resolved = aliases.get(correction_key)
                if resolved is None:
                    aliases[correction_key] = correction
                    autocorrecter.add_word(correction)
                    resolved = correction

                if resolution.cache_original:
                    aliases[name] = resolved
                    if resolution.persist_alias and len(resolution.names) == 1:
                        add_name_alias(base_dir, name, resolved)

                resolved_names.append(resolved)
                actual_names.append(resolved)

            if resolved_names:
                if resolution.merge_with_prev and prev_token:
                    last_appended_token = f"{prev_token}{name}"
                elif resolution.merge_with_next and next_token:
                    last_appended_token = f"{name}{next_token}"
                else:
                    last_appended_token = name

            if resolution.merge_with_next and next_token:
                i += 2
            else:
                i += 1

        handled = False
        if "not" in actual_names:
            if len(actual_names) >= 3 and actual_names[1] == "not":
                dkp_count[actual_names[0]] = dkp_count.get(actual_names[0], 0) + points
                for name in actual_names[2:]:
                    dkp_count[name] = dkp_count.get(name, 0) - points
                handled = True
            elif len(actual_names) >= 2 and actual_names[0] == "not":
                for name in actual_names[1:]:
                    dkp_count[name] = dkp_count.get(name, 0) - points
                handled = True

        if not handled:
            cleaned = {n for n in actual_names if n != "not"}
            for name in cleaned:
                dkp_count[name] = dkp_count.get(name, 0) + points
        else:
            cleaned = set()

        event_entries: List[EventEntry] = []
        if handled:
            if len(actual_names) >= 3 and actual_names[1] == "not":
                event_entries.append(EventEntry(actual_names[0], points))
                for name in actual_names[2:]:
                    if name != "not":
                        event_entries.append(EventEntry(name, -points))
            elif len(actual_names) >= 2 and actual_names[0] == "not":
                for name in actual_names[1:]:
                    if name != "not":
                        event_entries.append(EventEntry(name, -points))
        else:
            for name in cleaned:
                event_entries.append(EventEntry(name, points))

        if event_entries:
            if any(entry.delta > 0 for entry in event_entries):
                boss_set.add(boss_key)
            for entry in event_entries:
                if not entry.name:
                    continue
                counts = boss_counts.setdefault(entry.name, {})
                current = counts.get(boss_key, 0)
                if entry.delta > 0:
                    counts[boss_key] = current + 1
                elif entry.delta < 0 and current > 0:
                    new_value = current - 1
                    if new_value > 0:
                        counts[boss_key] = new_value
                    else:
                        counts.pop(boss_key, None)
                        if not counts:
                            boss_counts.pop(entry.name, None)

        event_time = get_date(line_text_raw)
        if event_time and event_entries:
            events.append(
                EventRecord(
                    event_time=event_time,
                    boss=boss_key,
                    points=points,
                    entries=event_entries,
                    source_line=line_text_raw,
                )
            )

    totals = [(name, points) for name, points in dkp_count.items() if points > 0]
    totals.sort(key=lambda item: item[0].lower())

    boss_list = sorted(boss_set, key=str.lower)
    return CalculationResult(
        totals=totals,
        sanity=sanity,
        errors=errors,
        boss_counts=boss_counts,
        boss_list=boss_list,
        events=events,
    )


def estimate_unknown_count(
    timers_path: Path,
    start_date: Optional[datetime],
    end_date: Optional[datetime],
    use_all_entries: bool,
    spreadsheet_id: str,
    range_name: str,
    credentials_path: Path,
    token_path: Path,
    base_dir: Path,
) -> Tuple[int, List[str]]:
    points_store = PointsStore(base_dir)

    lines = preprocess_lines(timers_path, base_dir)
    if not use_all_entries and start_date and end_date:
        lines = slice_by_date(lines, start_date, end_date)

    formatted_lines, errors = validate_lines(lines, points_store)
    if errors.any():
        return 0, []

    names = get_names_from_sheets(
        spreadsheet_id=spreadsheet_id,
        range_name=range_name,
        credentials_path=credentials_path,
        token_path=token_path,
    )

    aliases = build_aliases(names, base_dir)
    seen: Set[str] = set()
    discard: Set[str] = set()
    count = 0

    for _, tokens in formatted_lines:
        for name in tokens[1:]:
            if name in {MULTI_NOT_MARKER, "not"}:
                continue
            if name in discard or len(name) <= 1:
                continue
            if name in aliases:
                continue
            if name in seen:
                continue
            seen.add(name)
            count += 1

    return count, names
