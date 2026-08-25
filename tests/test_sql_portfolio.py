from pathlib import Path
import unittest

import duckdb


class SqlPortfolioTests(unittest.TestCase):
    def execute_example(self, filename):
        repository = Path(__file__).resolve().parents[1]
        sql_directory = repository / "sql"
        connection = duckdb.connect(":memory:")
        self.addCleanup(connection.close)
        connection.execute((sql_directory / "00_demo_schema_and_data.sql").read_text(encoding="utf-8"))
        cursor = connection.execute((sql_directory / filename).read_text(encoding="utf-8"))
        return [column[0] for column in cursor.description], cursor.fetchall()

    def test_material_code_substitution_example_uses_only_fictional_candidates(self):
        _, rows = self.execute_example("07_material_code_substitution_candidates.sql")

        selected_codes = {request_id: proposed_code for request_id, _, proposed_code, *_ in rows}
        self.assertEqual(
            {
                "REQ-001": "MC-A2",
                "REQ-002": "MC-B2",
                "REQ-003": "MC-C2",
            },
            selected_codes,
        )

    def test_leads_funnel_example_preserves_the_source_formula(self):
        columns, rows = self.execute_example("08_leads_funnel_analysis.sql")
        result = {
            (row[0], row[1]): dict(zip(columns[2:], row[2:]))
            for row in rows
        }

        self.assertEqual(1, result[("MODEL-A", "test drive")]["2026-01-10"])
        self.assertIsNone(result[("MODEL-A", "test drive")]["2026-01-11"])
        self.assertEqual(1, result[("MODEL-B", "test drive request")]["2026-01-11"])

    def test_free_car_reconciliation_surfaces_only_active_inconsistencies(self):
        _, rows = self.execute_example("09_free_cars_reconciliation.sql")

        self.assertEqual(
            {
                ("VEH-FC-001", "ASSIGNED_OR_LINKED_IN_TRACKING"),
                ("VEH-FC-002", "FREE_CAR_NOTE_TAG_MISMATCH"),
                ("VEH-FC-003", "OFFLINE_OR_MISSING_VESSEL"),
                ("VEH-FC-004", "MISSING_FROM_VEHICLE_TRACKING"),
                ("VEH-FC-005", "ALSO_IN_RESERVATIONS"),
            },
            set(rows),
        )
