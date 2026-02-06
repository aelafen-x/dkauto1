import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional


def _runs_path(base_dir: Path) -> Path:
    return base_dir / "runs" / "events.json"


def _ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _local_tzinfo():
    return datetime.now().astimezone().tzinfo or timezone.utc


def _isoformat_utc(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=_local_tzinfo())
    return dt.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def _parse_iso(value: str) -> datetime:
    if value.endswith("Z"):
        value = value.replace("Z", "+00:00")
    return datetime.fromisoformat(value)


def load_run_store(base_dir: Path) -> Dict[str, Any]:
    path = _runs_path(base_dir)
    if not path.exists():
        return {"version": 1, "runs": [], "events": []}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"version": 1, "runs": [], "events": []}


def save_run_store(base_dir: Path, data: Dict[str, Any]) -> None:
    path = _runs_path(base_dir)
    _ensure_parent(path)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def save_run(
    base_dir: Path,
    run_meta: Dict[str, Any],
    events: List[Dict[str, Any]],
) -> None:
    data = load_run_store(base_dir)
    data.setdefault("version", 1)
    data.setdefault("runs", [])
    data.setdefault("events", [])

    start = _parse_iso(run_meta["start_utc"])
    end = _parse_iso(run_meta["end_utc"])
    run_id = run_meta["run_id"]

    for event in data["events"]:
        if not event.get("active", True):
            continue
        event_time_raw = event.get("event_time_utc")
        if not event_time_raw:
            continue
        event_time = _parse_iso(event_time_raw)
        if start <= event_time <= end:
            event["active"] = False
            event["replaced_by"] = run_id

    data["runs"].append(run_meta)
    data["events"].extend(events)
    save_run_store(base_dir, data)


def build_run_meta(
    run_id: str,
    created_utc: datetime,
    start_utc: datetime,
    end_utc: datetime,
    event_count: int,
    timers_path: Optional[str] = None,
) -> Dict[str, Any]:
    meta = {
        "run_id": run_id,
        "created_utc": _isoformat_utc(created_utc),
        "start_utc": _isoformat_utc(start_utc),
        "end_utc": _isoformat_utc(end_utc),
        "event_count": int(event_count),
    }
    if timers_path:
        meta["timers_path"] = timers_path
    return meta


def normalize_event(
    run_id: str,
    created_utc: datetime,
    event_time: datetime,
    boss: str,
    points: int,
    entries: List[Dict[str, Any]],
    source_line: str,
) -> Dict[str, Any]:
    return {
        "run_id": run_id,
        "created_utc": _isoformat_utc(created_utc),
        "event_time_utc": _isoformat_utc(event_time),
        "boss": boss,
        "points": int(points),
        "entries": entries,
        "source_line": source_line,
        "active": True,
        "replaced_by": None,
    }


def iter_active_events(base_dir: Path) -> List[Dict[str, Any]]:
    data = load_run_store(base_dir)
    events = data.get("events", [])
    return [event for event in events if event.get("active", True)]


def iso_to_dt(value: str) -> datetime:
    return _parse_iso(value)
