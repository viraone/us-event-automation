"""Regression coverage for the configured four-month scraper workflow."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from main import (
    DAILY_GIG_HEADERS,
    MONTH_CONFIGS,
    TARGET_MONTHS,
    MonthConfig,
    merge_daily_month_rows,
    normalize_month_events,
    verify_month_stage,
    write_month_stage,
)


RUN_ONE = "2026-08-11T10:00:00"
RUN_TWO = "2026-08-12T10:00:00"


def calendar_event(event_date: str, raw: str, public_id: str) -> dict:
    return {
        "date": event_date,
        "metadataDate": event_date,
        "dayCellDate": event_date,
        "rawText": raw,
        "displayedTime": raw.split(" ", 1)[0],
        "start": f"{event_date}T11:00:00",
        "eventId": f"render-{public_id}",
        "publicId": public_id,
        "color": "",
        "url": "",
        "extendedProps": {},
    }


def temporary_config(directory: Path, year: int, month: int, name: str) -> MonthConfig:
    return MonthConfig(
        year=year,
        month=month,
        label=f"{name} {year}",
        xlsx_path=directory / f"US_Event_{name}_{year}_Gigs.xlsx",
        csv_path=directory / f"US_Event_{name}_{year}_Gigs.csv",
        sheet_name=f"{name} {year} Gigs",
        table_name=f"{name}{year}Gigs",
    )


class MultiMonthConfigurationTests(unittest.TestCase):
    def test_four_months_and_output_names_are_configured(self) -> None:
        self.assertEqual(
            TARGET_MONTHS,
            ((2026, 8), (2026, 9), (2026, 10), (2026, 11)),
        )
        self.assertEqual(
            [config.prefix for config in MONTH_CONFIGS],
            ["2026-08", "2026-09", "2026-10", "2026-11"],
        )
        self.assertEqual(
            [config.xlsx_path.name for config in MONTH_CONFIGS],
            [
                "US_Event_August_2026_Gigs.xlsx",
                "US_Event_September_2026_Gigs.xlsx",
                "US_Event_October_2026_Gigs.xlsx",
                "US_Event_November_2026_Gigs.xlsx",
            ],
        )

    def test_october_and_november_exclude_trailing_calendar_dates(self) -> None:
        cases = (
            (MONTH_CONFIGS[2], "2026-09-30", "2026-10-01", "2026-11-01"),
            (MONTH_CONFIGS[3], "2026-10-31", "2026-11-01", "2026-12-01"),
        )
        for config, previous_date, included_date, next_date in cases:
            with self.subTest(month=config.prefix):
                rows = normalize_month_events(
                    [
                        calendar_event(previous_date, "11a WA|| Previous 1 (Tacoma)", "previous"),
                        calendar_event(included_date, "11a WA|| Safeway 1142 (Kirkland)", "included"),
                        calendar_event(next_date, "11a WA|| Next 1 (Seattle)", "next"),
                    ],
                    config,
                    RUN_ONE,
                )
                self.assertEqual([row["date"] for row in rows], [included_date])
                self.assertEqual(rows[0]["preferred_area"], "YES")
                self.assertEqual(rows[0]["location_priority"], 2)

    def test_new_removed_and_baseline_history_work_for_new_months(self) -> None:
        for config in MONTH_CONFIGS[2:]:
            with self.subTest(month=config.prefix):
                day = f"{config.prefix}-01"
                first_rows = normalize_month_events(
                    [calendar_event(day, "11a WA|| Safeway 1142 (Kirkland)", "old")],
                    config,
                    RUN_ONE,
                )
                baseline, baseline_summary = merge_daily_month_rows(
                    first_rows, False, [], RUN_ONE
                )
                self.assertEqual(baseline[0]["status"], "EXISTING")
                self.assertEqual(baseline_summary["new"], 0)

                new_rows = normalize_month_events(
                    [calendar_event(day, "1p WA|| Costco 1 (Seattle)", "new")],
                    config,
                    RUN_TWO,
                )
                merged, summary = merge_daily_month_rows(
                    new_rows, True, baseline, RUN_TWO
                )
                self.assertEqual(summary, {"current": 1, "new": 1, "existing": 0, "removed": 1})
                self.assertEqual(
                    {row["public_id"]: row["status"] for row in merged},
                    {"old": "REMOVED", "new": "NEW"},
                )
                self.assertEqual(
                    next(row for row in merged if row["public_id"] == "old")["first_seen"],
                    RUN_ONE,
                )

    def test_october_and_november_xlsx_and_csv_use_shared_schema(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            for month, name in ((10, "October"), (11, "November")):
                with self.subTest(month=name):
                    config = temporary_config(directory, 2026, month, name)
                    current = normalize_month_events(
                        [calendar_event(f"2026-{month:02d}-01", "11a WA|| Safeway 1142 (Kirkland)", name.lower())],
                        config,
                        RUN_ONE,
                    )
                    rows, _ = merge_daily_month_rows(current, False, [], RUN_ONE)
                    write_month_stage(rows, config, config.xlsx_path, config.csv_path)
                    verify_month_stage(rows, config, config.xlsx_path, config.csv_path)
                    self.assertTrue(config.xlsx_path.exists())
                    self.assertTrue(config.csv_path.exists())
                    self.assertEqual(
                        config.csv_path.read_text(encoding="utf-8-sig").splitlines()[0].split(","),
                        DAILY_GIG_HEADERS,
                    )


if __name__ == "__main__":
    unittest.main()
