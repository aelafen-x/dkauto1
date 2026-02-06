import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from pyapp.core.sanitise import preprocess_lines, validate_lines
from pyapp.core.workflow import Resolution, calculate_points


def _write_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


class WorkflowPointsTests(unittest.TestCase):
    def _setup_base_dir(self, base_dir: Path) -> None:
        _write_json(base_dir / "points.json", {"boss1": 10, "boss2": 5})
        _write_json(base_dir / "prios.json", [])
        _write_json(base_dir / "boss_aliases.json", [])
        _write_json(base_dir / "name_aliases.json", {})

    def _write_timers(self, base_dir: Path, lines: list[str]) -> Path:
        timers_path = base_dir / "timers.txt"
        timers_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return timers_path

    def _validate_lines(self, base_dir: Path, lines: list[str]):
        timers_path = self._write_timers(base_dir, lines)
        processed = preprocess_lines(timers_path, base_dir)
        points_store = None
        from pyapp.core.points import PointsStore

        points_store = PointsStore(base_dir)
        formatted, errors = validate_lines(processed, points_store)
        return formatted, errors

    @patch("pyapp.core.workflow.get_names_from_sheets")
    def test_basic_points_and_boss_counts(self, mock_get_names) -> None:
        mock_get_names.return_value = ["alice", "bob"]
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                [
                    "01 Jan 2026 at 20:00: boss1 alice bob",
                    "02 Jan 2026 at 20:00: boss2 alice",
                ],
            )

            result = calculate_points(
                timers_path=timers_path,
                start_date=None,
                end_date=None,
                use_all_entries=True,
                spreadsheet_id="dummy",
                range_name="dummy",
                credentials_path=base_dir / "credentials.json",
                token_path=base_dir / "token.json",
                base_dir=base_dir,
                resolve_unknown=lambda *_: None,
            )

            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 15)
            self.assertEqual(totals["bob"], 10)
            self.assertEqual(result.boss_list, ["boss1", "boss2"])
            boss_counts = {name.lower(): counts for name, counts in result.boss_counts.items()}
            self.assertEqual(boss_counts["alice"]["boss1"], 1)
            self.assertEqual(boss_counts["alice"]["boss2"], 1)
            self.assertEqual(boss_counts["bob"]["boss1"], 1)

    @patch("pyapp.core.workflow.get_names_from_sheets")
    def test_not_logic_counts_only_positive_name(self, mock_get_names) -> None:
        mock_get_names.return_value = ["Alice", "Bob", "Carl"]
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alice not bob carl __multinot__"],
            )

            result = calculate_points(
                timers_path=timers_path,
                start_date=None,
                end_date=None,
                use_all_entries=True,
                spreadsheet_id="dummy",
                range_name="dummy",
                credentials_path=base_dir / "credentials.json",
                token_path=base_dir / "token.json",
                base_dir=base_dir,
                resolve_unknown=lambda *_: None,
            )

            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 10)
            self.assertNotIn("bob", totals)
            self.assertNotIn("carl", totals)
            boss_counts = {name.lower(): counts for name, counts in result.boss_counts.items()}
            self.assertEqual(boss_counts["alice"]["boss1"], 1)
            self.assertNotIn("bob", boss_counts)
            self.assertNotIn("carl", boss_counts)

    @patch("pyapp.core.workflow.get_names_from_sheets")
    def test_date_slice_respects_start_date(self, mock_get_names) -> None:
        mock_get_names.return_value = ["Alice", "Bob"]
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                [
                    "25 Dec 2025 at 20:00: boss1 bob",
                    "01 Jan 2026 at 20:00: boss1 alice",
                ],
            )

            result = calculate_points(
                timers_path=timers_path,
                start_date=datetime(2026, 1, 1, 0, 0, tzinfo=timezone.utc),
                end_date=datetime(2026, 1, 8, 0, 0, tzinfo=timezone.utc),
                use_all_entries=False,
                spreadsheet_id="dummy",
                range_name="dummy",
                credentials_path=base_dir / "credentials.json",
                token_path=base_dir / "token.json",
                base_dir=base_dir,
                resolve_unknown=lambda *_: None,
            )

            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 10)
            self.assertNotIn("bob", totals)

    def test_multinot_detection(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            _, errors = self._validate_lines(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alice not bob carl __multinot__"],
            )
            self.assertEqual(errors.incorrect_not_lines, [])

    def test_multinot_points_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alice not bob carl __multinot__"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice", "Bob", "Carl"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals.get("alice"), 10)
            self.assertNotIn("bob", totals)
            self.assertNotIn("carl", totals)

    def test_not_detection(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            _, errors = self._validate_lines(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alice not bob"],
            )
            self.assertEqual(errors.incorrect_not_lines, [])

            _, errors_valid = self._validate_lines(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 not alice"],
            )
            self.assertEqual(errors_valid.incorrect_not_lines, [])

            _, errors_invalid = self._validate_lines(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alice not bob carl"],
            )
            self.assertEqual(errors_invalid.incorrect_not_lines, [1])

    def test_single_not_subtraction(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alice not bob"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice", "Bob"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals.get("alice"), 10)
            self.assertNotIn("bob", totals)

    def test_subtraction_offsets_existing_points(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                [
                    "01 Jan 2026 at 20:00: boss1 bob",
                    "01 Jan 2026 at 20:05: boss1 alice not bob",
                ],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice", "Bob"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals.get("alice"), 10)
            self.assertNotIn("bob", totals)

    def test_boss_points_and_alias_points(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            _write_json(base_dir / "points.json", {"boss1": 10})
            _write_json(base_dir / "prios.json", [])
            _write_json(base_dir / "boss_aliases.json", [{"b1": "boss1"}])
            _write_json(base_dir / "name_aliases.json", {})
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: b1 alice"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 10)

    def test_player_alias_points_calculation(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            _write_json(base_dir / "name_aliases.json", {"ali": "Alice"})
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 ali"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 10)

    def test_resolve_unknown_fixed_value_points(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 alcie"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice"]

                def resolver(name, suggestions, *_args):
                    return Resolution(names=["Alice"], cache_original=True)

                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=resolver,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 10)

    def test_autocorrect_add_new_name_points(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 zed"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice"]

                def resolver(name, suggestions, *_args):
                    return Resolution(names=["Zed"], cache_original=True, add_new=True)

                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=resolver,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["zed"], 10)

    def test_autocorrect_unknown_without_alias_discards(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                ["01 Jan 2026 at 20:00: boss1 zed"],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertNotIn("zed", totals)

    def test_modifier_points_double_and_brucybonus(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            base_dir = Path(tmpdir)
            self._setup_base_dir(base_dir)
            timers_path = self._write_timers(
                base_dir,
                [
                    "01 Jan 2026 at 20:00: boss1(doublepoints) alice",
                    "01 Jan 2026 at 20:00: boss1(brucybonus) bob",
                ],
            )
            with patch("pyapp.core.workflow.get_names_from_sheets") as mock_get_names:
                mock_get_names.return_value = ["Alice", "Bob"]
                result = calculate_points(
                    timers_path=timers_path,
                    start_date=None,
                    end_date=None,
                    use_all_entries=True,
                    spreadsheet_id="dummy",
                    range_name="dummy",
                    credentials_path=base_dir / "credentials.json",
                    token_path=base_dir / "token.json",
                    base_dir=base_dir,
                    resolve_unknown=lambda *_: None,
                )
            totals = {name.lower(): points for name, points in result.totals}
            self.assertEqual(totals["alice"], 20)
            self.assertEqual(totals["bob"], 15)


if __name__ == "__main__":
    unittest.main()
