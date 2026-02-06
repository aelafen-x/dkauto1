import json
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from platformdirs import user_config_dir


@dataclass
class AppConfig:
    spreadsheet_id: str = ""
    range_name: str = "DKP Sheet!B3:B"
    last_timers_path: str = ""
    last_credentials_path: str = ""
    use_all_entries: bool = True
    start_date_iso: str = ""
    end_date_iso: str = ""
    use_native_dialog: bool = True
    activity_a_threshold: int = 70
    activity_aplus_threshold: int = 300


def config_path() -> Path:
    base = Path(user_config_dir("dkp_automator_gui"))
    return base / "settings.json"


def token_path() -> Path:
    base = Path(user_config_dir("dkp_automator_gui"))
    return base / "token.json"


def load_config() -> AppConfig:
    path = config_path()
    if not path.exists():
        return AppConfig()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return AppConfig()

    return AppConfig(
        spreadsheet_id=data.get("spreadsheet_id", ""),
        range_name=data.get("range_name", "DKP Sheet!B3:B"),
        last_timers_path=data.get("last_timers_path", ""),
        last_credentials_path=data.get("last_credentials_path", ""),
        use_all_entries=bool(data.get("use_all_entries", True)),
        start_date_iso=data.get("start_date_iso", ""),
        end_date_iso=data.get("end_date_iso", ""),
        use_native_dialog=bool(data.get("use_native_dialog", True)),
        activity_a_threshold=int(data.get("activity_a_threshold", 70)),
        activity_aplus_threshold=int(data.get("activity_aplus_threshold", 300)),
    )


def save_config(cfg: AppConfig) -> None:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "spreadsheet_id": cfg.spreadsheet_id,
        "range_name": cfg.range_name,
        "last_timers_path": cfg.last_timers_path,
        "last_credentials_path": cfg.last_credentials_path,
        "use_all_entries": cfg.use_all_entries,
        "start_date_iso": cfg.start_date_iso,
        "end_date_iso": cfg.end_date_iso,
        "use_native_dialog": cfg.use_native_dialog,
        "activity_a_threshold": int(cfg.activity_a_threshold),
        "activity_aplus_threshold": int(cfg.activity_aplus_threshold),
    }
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
