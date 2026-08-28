"""Regression tests for the local, read-only US Event dashboard.

The tests build disposable workbooks with ``openpyxl`` so they exercise the
same file-discovery and schema-normalization paths as the real scraper output.
No production workbook is ever written by this suite.
"""

from __future__ import annotations

from datetime import date, datetime, time
import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from urllib.parse import quote

from openpyxl import Workbook

from app import create_app
from dashboard_data import discover_workbooks, load_dashboard_data


PROJECT_DIR = Path(__file__).resolve().parents[1]


STANDARD_HEADERS = [
    "Date",
    "Day",
    "Start Time",
    "Start Time Sort",
    "State",
    "Store / Account",
    "Store Number",
    "City",
    "Preferred Area",
    "Location Priority",
    "Raw Calendar Listing",
    "Calendar Color",
    "Event URL",
    "Status",
    "First Seen",
    "Last Seen",
    "Scraped At",
    "Event ID",
]


def write_workbook(
    path: Path,
    rows: list[list[object]],
    *,
    headers: list[str] | None = None,
    preamble: list[list[object]] | None = None,
    sheet_name: str = "Gigs",
) -> Path:
    """Write a small fixture workbook and return its path."""

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name
    for preamble_row in preamble or []:
        worksheet.append(preamble_row)
    worksheet.append(headers or STANDARD_HEADERS)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)
    workbook.close()
    return path


def standard_row(
    *,
    gig_date: object = "2026-08-14",
    start_time: object = "4p",
    state: object = "WA",
    account: object = "Safeway",
    store_number: object = "3265",
    city: object = "Chelan",
    preferred: object = "NO",
    location_priority: object = 9,
    raw_listing: object = "4p WA|| Safeway 3265 (Chelan)",
    status: object = "EXISTING",
    first_seen: object = "2026-08-10 09:00:00",
    last_seen: object = "2026-08-11 09:00:00",
    scraped_at: object = "2026-08-11 09:00:00",
    event_id: object = "event-1",
) -> list[object]:
    """Return a row matching the production workbook schema."""

    return [
        gig_date,
        "",
        start_time,
        "",
        state,
        account,
        store_number,
        city,
        preferred,
        location_priority,
        raw_listing,
        "#808080",
        "https://example.test/event",
        status,
        first_seen,
        last_seen,
        scraped_at,
        event_id,
    ]


def file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(64 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


class TemporaryWorkbookTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp_directory = tempfile.TemporaryDirectory()
        self.data_dir = Path(self._temp_directory.name)

    def tearDown(self) -> None:
        self._temp_directory.cleanup()


class DiscoveryAndSchemaTests(TemporaryWorkbookTestCase):
    def test_rows_receive_stable_opaque_gig_keys_for_ui_selection(self) -> None:
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [
                standard_row(event_id="stable-event-1"),
                standard_row(
                    gig_date="2026-08-15",
                    start_time="2p",
                    raw_listing="2p WA|| Safeway 1600 (Bellevue)",
                    city="Bellevue",
                    event_id="",
                ),
            ],
        )

        first = load_dashboard_data(self.data_dir, strict=True)["rows"]
        second = load_dashboard_data(self.data_dir, strict=True)["rows"]
        first_keys = [row["gig_key"] for row in first]

        self.assertEqual(first_keys, [row["gig_key"] for row in second])
        self.assertEqual(len(first_keys), len(set(first_keys)))
        self.assertTrue(all(key.startswith("gig-") and len(key) == 28 for key in first_keys))
        self.assertTrue(all("stable-event-1" not in key for key in first_keys))

    def test_discovers_current_and_future_months_dynamically(self) -> None:
        august = write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [standard_row()],
        )
        october = write_workbook(
            self.data_dir / "US_Event_October_2027_Gigs.xlsx",
            [
                standard_row(
                    gig_date="2027-10-03",
                    raw_listing="11a WA|| Costco 1 (Seattle)",
                    start_time="11a",
                    account="Costco",
                    store_number=1,
                    city="Seattle",
                    event_id="future-1",
                )
            ],
        )
        backup_dir = self.data_dir / "backup"
        backup_dir.mkdir()
        write_workbook(
            backup_dir / "US_Event_September_2026_Gigs.xlsx",
            [standard_row(event_id="backup-only")],
        )
        (self.data_dir / "US_Event_August_2026_Gigs.csv").write_text(
            "Date,City\n2026-08-14,Seattle\n", encoding="utf-8"
        )

        discovered = discover_workbooks(self.data_dir)
        # ``TemporaryDirectory`` may expose ``/var`` while ``Path.resolve``
        # canonicalizes it to macOS's ``/private/var``.  Names and ordering are
        # the discovery contract; both paths still identify the same files.
        self.assertEqual(
            [path.name for path in discovered], [august.name, october.name]
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        self.assertEqual(snapshot["discovered_file_count"], 2)
        self.assertEqual(
            [month["key"] for month in snapshot["months"]],
            ["2026-08", "2027-10"],
        )
        self.assertEqual(
            [month["label"] for month in snapshot["months"]],
            ["August 2026", "October 2027"],
        )
        self.assertEqual(len(snapshot["rows"]), 2)
        self.assertEqual(snapshot["filter_options"]["months"][-1]["value"], "2027-10")

    def test_evolving_schema_and_missing_optional_columns_are_safe(self) -> None:
        headers = [
            "Gig Date",
            "Time",
            "Region",
            "Account",
            "Store No",
            "Town",
            "Preferred",
            "Raw Listing",
            "Gig Status",
            "Event Identifier",
            "Assignment Note",
        ]
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [
                [
                    "08/14/2026",
                    "4:30p",
                    "wa",
                    "Safeway",
                    3265,
                    "BELLEVUE",
                    "no",
                    "4:30p WA|| Safeway 3265 (Bellevue)",
                    "added",
                    "semantic-1",
                    "Bring table",
                ]
            ],
            headers=headers,
            preamble=[["US Event Management export"], []],
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        row = snapshot["rows"][0]
        self.assertEqual(row["date_iso"], "2026-08-14")
        self.assertEqual(row["day"], "Friday")
        self.assertEqual(row["start_time"], "4:30 PM")
        self.assertEqual(row["start_time_sort"], "16:30")
        self.assertEqual(row["state"], "WA")
        self.assertEqual(row["store_display"], "Safeway 3265")
        self.assertEqual(row["city"], "Bellevue")
        self.assertEqual(row["preferred_area"], "YES")
        self.assertEqual(row["location_priority"], 1)
        self.assertEqual(row["status"], "NEW")
        self.assertEqual(row["event_id"], "semantic-1")
        self.assertEqual(row["event_url"], "")
        self.assertEqual(row["calendar_color"], "")
        self.assertEqual(row["first_seen"], "")
        self.assertEqual(row["extra_fields"], {"Assignment Note": "Bring table"})
        self.assertEqual(snapshot["errors"], [])

    def test_sparse_supported_schema_uses_filename_month_without_inventing_values(self) -> None:
        write_workbook(
            self.data_dir / "US_Event_December_2028_Gigs.xlsx",
            [[None, "Tacoma"], [date(2028, 12, 8), "Seattle"]],
            headers=["Date", "City"],
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        self.assertEqual(len(snapshot["rows"]), 2)
        # Valid dates sort before sparse undated rows.
        dated, undated = snapshot["rows"]
        self.assertEqual(undated["date_iso"], "")
        self.assertEqual(undated["month_key"], "2028-12")
        self.assertEqual(undated["start_time"], "")
        self.assertEqual(undated["status"], "")
        self.assertEqual(undated["preferred_area"], "NO")
        self.assertEqual(dated["date_iso"], "2028-12-08")
        self.assertEqual(dated["preferred_area"], "YES")

    def test_status_and_preferred_city_normalization(self) -> None:
        cases = [
            ("BELLEVUE", "added", "YES", 1, "NEW"),
            ("kirkland", "active", "YES", 2, "EXISTING"),
            ("Seattle", "deleted", "YES", 3, "REMOVED"),
            ("Tacoma", "missing", "NO", 9, "REMOVED"),
        ]
        rows = []
        for index, (city, status, _, _, _) in enumerate(cases, start=1):
            rows.append(
                standard_row(
                    city=city,
                    status=status,
                    preferred="YES",  # City semantics remain authoritative.
                    event_id=f"normal-{index}",
                    raw_listing=f"4p WA|| Store {index} ({city})",
                )
            )
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx", rows
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        by_city = {row["city"]: row for row in snapshot["rows"]}
        for city, _, expected_preferred, expected_priority, expected_status in cases:
            normalized_city = city.title() if city.isupper() or city.islower() else city
            row = by_city[normalized_city]
            self.assertEqual(row["preferred_area"], expected_preferred)
            self.assertEqual(row["location_priority"], expected_priority)
            self.assertEqual(row["status"], expected_status)

        self.assertEqual(snapshot["summary"]["total_rows"], 4)
        self.assertEqual(snapshot["summary"]["total_current"], 2)
        self.assertEqual(snapshot["summary"]["new"], 1)
        self.assertEqual(snapshot["summary"]["preferred"], 2)
        self.assertEqual(snapshot["summary"]["removed"], 2)
        self.assertEqual(snapshot["filter_options"]["statuses"], ["NEW", "EXISTING", "REMOVED"])


class DeduplicationAndSortTests(TemporaryWorkbookTestCase):
    def test_event_id_dedupe_prefers_fresh_state_and_backfills_blanks(self) -> None:
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [
                standard_row(
                    city="Seattle",
                    status="EXISTING",
                    last_seen="2026-08-10 09:00:00",
                    scraped_at="2026-08-10 09:00:00",
                    event_id="SHARED-ID",
                )
            ],
        )
        write_workbook(
            self.data_dir / "US_Event_September_2026_Gigs.xlsx",
            [
                standard_row(
                    gig_date="2026-08-14",
                    city="",
                    account="",
                    store_number="",
                    status="REMOVED",
                    last_seen="2026-08-12 09:00:00",
                    scraped_at="2026-08-12 09:00:00",
                    event_id="shared-id",
                )
            ],
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        self.assertEqual(len(snapshot["rows"]), 1)
        row = snapshot["rows"][0]
        self.assertEqual(row["status"], "REMOVED")
        self.assertEqual(row["city"], "Seattle")
        self.assertEqual(row["store_display"], "Safeway 3265")
        self.assertEqual(
            row["source_files"],
            [
                "US_Event_August_2026_Gigs.xlsx",
                "US_Event_September_2026_Gigs.xlsx",
            ],
        )

    def test_composite_dedupe_uses_date_time_and_normalized_raw_listing(self) -> None:
        raw = "4p WA|| Safeway 3265 (Chelan)"
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [
                standard_row(event_id="", raw_listing=raw),
                standard_row(
                    event_id="",
                    raw_listing="4p WA|| Safeway 9999 (Chelan)",
                    store_number="9999",
                ),
            ],
        )
        write_workbook(
            self.data_dir / "US_Event_September_2026_Gigs.xlsx",
            [
                standard_row(
                    event_id="",
                    raw_listing="  4p   WA|| Safeway 3265 (Chelan)  ",
                    status="NEW",
                    scraped_at="2026-08-12 10:00:00",
                )
            ],
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        self.assertEqual(len(snapshot["rows"]), 2)
        deduped = next(row for row in snapshot["rows"] if "3265" in row["raw_calendar_listing"])
        self.assertEqual(deduped["status"], "NEW")
        self.assertEqual(len(deduped["source_files"]), 2)

    def test_rows_sort_by_date_then_time_then_priority_then_city(self) -> None:
        specifications = [
            ("2026-08-14", "4p", "Mountlake Terrace", "mt"),
            ("2026-08-14", "4p", "Seattle", "sea"),
            ("2026-08-14", "4p", "Chelan", "chelan"),
            ("2026-08-14", "4p", "Kirkland", "kirk"),
            ("2026-08-14", "4p", "Bellevue", "bell"),
            ("2026-08-14", "2p", "Kirkland", "two"),
            ("2026-08-14", "1p", "Seattle", "one"),
            ("2026-08-13", "5p", "Tacoma", "prior-date"),
            ("2026-08-14", "", "Bellevue", "no-time"),
        ]
        rows = [
            standard_row(
                gig_date=gig_date,
                start_time=start,
                city=city,
                raw_listing=f"{start} WA|| Store ({city})",
                event_id=event_id,
            )
            for gig_date, start, city, event_id in specifications
        ]
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx", rows
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        order = [row["event_id"] for row in snapshot["rows"]]
        self.assertEqual(
            order,
            [
                "prior-date",
                "one",
                "two",
                "bell",
                "kirk",
                "sea",
                "chelan",
                "mt",
                "no-time",
            ],
        )
        self.assertEqual(
            snapshot["filter_options"]["start_times"],
            ["1:00 PM", "2:00 PM", "4:00 PM", "5:00 PM"],
        )


class RealWorkbookTests(unittest.TestCase):
    def test_real_august_and_september_workbooks_load_read_only(self) -> None:
        required = {
            "US_Event_August_2026_Gigs.xlsx",
            "US_Event_September_2026_Gigs.xlsx",
        }
        sources = {path.name: path for path in discover_workbooks(PROJECT_DIR)}
        self.assertTrue(required.issubset(sources), sorted(sources))
        hashes_before = {name: file_digest(sources[name]) for name in required}

        snapshot = load_dashboard_data(PROJECT_DIR, strict=True)

        hashes_after = {name: file_digest(sources[name]) for name in required}
        self.assertEqual(hashes_after, hashes_before)
        self.assertEqual(snapshot["errors"], [])
        month_keys = {month["key"] for month in snapshot["months"]}
        self.assertTrue({"2026-08", "2026-09"}.issubset(month_keys))
        for month_key in ("2026-08", "2026-09"):
            self.assertTrue(
                any(row["month_key"] == month_key for row in snapshot["rows"]),
                f"expected at least one real row for {month_key}",
            )
        self.assertEqual(
            snapshot["summary"]["total_rows"], len(snapshot["rows"])
        )
        self.assertEqual(
            snapshot["summary"]["total_current"],
            sum(row["status"] != "REMOVED" for row in snapshot["rows"]),
        )
        self.assertTrue(
            all("_freshness" not in row for row in snapshot["rows"])
        )
        # The public snapshot must remain immediately serializable for Flask.
        json.dumps(snapshot)


class FlaskRouteTests(TemporaryWorkbookTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.workbook_path = write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [standard_row()],
        )
        self.application = create_app(self.data_dir)
        self.application.config.update(TESTING=True)
        self.client = self.application.test_client()

    def test_dashboard_api_and_health_routes(self) -> None:
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"US Event Gig Dashboard", response.data)
        self.assertIn(b"Safeway 3265", response.data)

        api_response = self.client.get("/api/gigs")
        self.assertEqual(api_response.status_code, 200)
        payload = api_response.get_json()
        self.assertEqual(len(payload["rows"]), 1)
        self.assertEqual(payload["rows"][0]["city"], "Chelan")
        self.assertEqual(
            payload["files"][0]["download_url"],
            "/spreadsheets/US_Event_August_2026_Gigs.xlsx",
        )

        health_response = self.client.get("/health")
        self.assertEqual(health_response.status_code, 200)
        self.assertEqual(
            health_response.get_json(),
            {"status": "ok", "workbooks": 1, "rows": 1},
        )

    def test_recruiter_selection_markup_uses_stable_keys_and_excludes_removed(self) -> None:
        write_workbook(
            self.workbook_path,
            [
                standard_row(event_id="selectable-event"),
                standard_row(
                    gig_date="2026-08-15",
                    raw_listing="2p WA|| Removed Store 2 (Seattle)",
                    status="REMOVED",
                    event_id="removed-event",
                ),
            ],
        )

        snapshot = load_dashboard_data(self.data_dir, strict=True)
        rows_by_status = {row["status"]: row for row in snapshot["rows"]}
        response = self.client.get("/")
        html = response.get_data(as_text=True)

        self.assertIn('id="generate-recruiter-text"', html)
        self.assertIn('id="selected-gig-count"', html)
        self.assertIn('id="recruiter-drawer"', html)
        self.assertIn('role="dialog"', html)
        self.assertIn('id="recruiter-message"', html)
        self.assertIn('id="copy-recruiter-text"', html)
        self.assertIn(
            f'data-gig-id="{rows_by_status["EXISTING"]["gig_key"]}"', html
        )
        self.assertIn(
            f'value="{rows_by_status["EXISTING"]["gig_key"]}"', html
        )
        self.assertEqual(html.count('class="gig-select-checkbox"'), 1)
        self.assertIn("Removed gig cannot be selected", html)

    def test_checkbox_selection_is_isolated_client_side(self) -> None:
        script = (PROJECT_DIR / "static" / "js" / "app.js").read_text(
            encoding="utf-8"
        )
        response = self.client.get("/")
        html = response.get_data(as_text=True)

        self.assertIn('class="gig-select-checkbox"', html)
        self.assertIn('type="checkbox"', html)
        self.assertRegex(
            html,
            r'id="generate-recruiter-text"[\s\S]*?type="button"',
        )
        self.assertIn("selectedGigIds: new Set()", script)
        self.assertRegex(
            script,
            r'row\.selectCheckbox\.addEventListener\("click", \(event\) => \{\s*event\.stopPropagation\(\);',
        )
        self.assertRegex(
            script,
            r'row\.selectCheckbox\.addEventListener\("change", \(event\) => \{\s*event\.stopPropagation\(\);',
        )

    def test_download_allows_only_exact_discovered_workbooks(self) -> None:
        expected_bytes = self.workbook_path.read_bytes()
        response = self.client.get(
            f"/spreadsheets/{quote(self.workbook_path.name)}"
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, expected_bytes)
        self.assertIn("attachment", response.headers["Content-Disposition"])
        self.assertIn(self.workbook_path.name, response.headers["Content-Disposition"])
        response.close()

        secret = self.data_dir / "secret.xlsx"
        write_workbook(secret, [standard_row(event_id="secret")])
        self.assertEqual(self.client.get("/spreadsheets/secret.xlsx").status_code, 404)
        self.assertEqual(
            self.client.get("/spreadsheets/../secret.xlsx").status_code, 404
        )
        self.assertEqual(
            self.client.get("/spreadsheets/%2E%2E%2Fsecret.xlsx").status_code,
            404,
        )
        self.assertEqual(
            self.client.get("/spreadsheets/US_Event_Missing_2026.xlsx").status_code,
            404,
        )

    def test_each_request_rereads_workbooks_from_disk(self) -> None:
        first_response = self.client.get("/api/gigs")
        first_payload = first_response.get_json()
        first_response.close()
        self.assertEqual(len(first_payload["rows"]), 1)

        write_workbook(
            self.workbook_path,
            [
                standard_row(),
                standard_row(
                    gig_date="2026-08-15",
                    city="Seattle",
                    start_time=time(13, 0),
                    raw_listing="1p WA|| Costco 1 (Seattle)",
                    account="Costco",
                    store_number=1,
                    event_id="event-2",
                ),
            ],
        )

        second_payload = self.client.get("/api/gigs").get_json()
        self.assertEqual(len(second_payload["rows"]), 2)
        self.assertEqual(second_payload["summary"]["total_current"], 2)
        self.assertEqual(
            {row["event_id"] for row in second_payload["rows"]},
            {"event-1", "event-2"},
        )
        self.assertEqual(self.client.get("/health").get_json()["rows"], 2)

    def test_dashboard_query_parameters_are_validated_server_side(self) -> None:
        valid = self.client.get("/?view=new&month=2026-08")
        self.assertEqual(valid.status_code, 200)
        self.assertIn(b"New gigs", valid.data)

        invalid = self.client.get("/?view=javascript%3Aalert(1)&month=2099-99")
        self.assertEqual(invalid.status_code, 200)
        self.assertIn(b"Dashboard", invalid.data)
        self.assertNotIn(b"javascript:alert(1)", invalid.data)


class ErrorToleranceTests(TemporaryWorkbookTestCase):
    def test_one_unreadable_matching_file_does_not_hide_valid_data(self) -> None:
        write_workbook(
            self.data_dir / "US_Event_August_2026_Gigs.xlsx",
            [standard_row()],
        )
        broken = self.data_dir / "US_Event_September_2026_Gigs.xlsx"
        broken.write_bytes(b"not an xlsx archive")

        snapshot = load_dashboard_data(self.data_dir)
        self.assertEqual(len(snapshot["rows"]), 1)
        self.assertEqual(snapshot["discovered_file_count"], 2)
        self.assertEqual(len(snapshot["errors"]), 1)
        self.assertEqual(snapshot["errors"][0]["file"], broken.name)
        self.assertIn("BadZipFile", snapshot["errors"][0]["error"])

        with self.assertRaises(Exception):
            load_dashboard_data(self.data_dir, strict=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
