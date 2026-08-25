from pathlib import Path
import unittest

import duckdb


class SqlPortfolioTests(unittest.TestCase):
    def test_material_code_substitution_example_uses_only_fictional_candidates(self):
        repository = Path(__file__).resolve().parents[1]
        sql_directory = repository / "sql"
        connection = duckdb.connect(":memory:")
        try:
            connection.execute((sql_directory / "00_demo_schema_and_data.sql").read_text(encoding="utf-8"))
            rows = connection.execute(
                (sql_directory / "07_material_code_substitution_candidates.sql").read_text(
                    encoding="utf-8"
                )
            ).fetchall()
        finally:
            connection.close()

        selected_codes = {request_id: proposed_code for request_id, _, proposed_code, *_ in rows}
        self.assertEqual(
            {
                "REQ-001": "MC-A2",
                "REQ-002": "MC-B2",
                "REQ-003": "MC-C2",
            },
            selected_codes,
        )
