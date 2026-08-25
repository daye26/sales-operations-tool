from pathlib import Path
import tempfile
import types
import unittest

from openpyxl import Workbook

import allocation_excel
import asignaciones_excel
import sales_operations_app
import tabular_normalization as tabular
import vehicle_tracking_loader


class EngineInfrastructureTests(unittest.TestCase):
    def test_tabular_helpers_normalize_generic_headers_and_keys(self):
        self.assertEqual(tabular.normalize_header("  Purchase-ETA_value  "), "purchase eta value")
        self.assertEqual(tabular.text_key("  Éxample\n value "), "EXAMPLE VALUE")
        self.assertEqual(
            tabular.build_indexes(
                ["VIN #", "Purchase ETA"],
                {"vin": ["vin #"], "eta": ["purchase eta"]},
                ("vin",),
                ("eta",),
            ),
            {"vin": 0, "eta": 1},
        )

    def test_vehicle_tracking_cache_does_not_hide_missing_required_source_fields(self):
        with tempfile.TemporaryDirectory() as directory:
            source_path = Path(directory) / "vehicle_tracking.xlsx"
            cache_path = Path(directory) / "vehicle_tracking.pkl"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["VIN #", "Purchase ETA", "Current warehouse", "Gate in date", "Current status"])
            worksheet.append(["TEST-VEHICLE-01", "2030-01-01", "PORT-A", "2030-01-02", "Available"])
            workbook.save(source_path)
            workbook.close()

            messages = []
            loader = lambda message: messages.append(message)
            vehicle_tracking_loader.load_vehicle_tracking(
                source_path,
                cache_path,
                loader,
                required_fields=("vin", "eta", "port", "gate_in", "status"),
            )

            with self.assertRaisesRegex(ValueError, "Missing column material_code"):
                vehicle_tracking_loader.load_vehicle_tracking(
                    source_path,
                    cache_path,
                    loader,
                    required_fields=("vin", "material_code"),
                )
            self.assertTrue(
                any("required source fields are unavailable" in message for message in messages)
            )

    def test_priority_references_support_or_ow_and_sales_orders(self):
        self.assertEqual(asignaciones_excel.priority_reference(" OR-00001 "), ("or_ow", "OR-00001"))
        self.assertEqual(asignaciones_excel.priority_reference(12345), ("sales_order", "12345"))
        self.assertTrue(
            allocation_excel.priority_reference_matches_reservation(
                {"or_number_key": "OR-00001", "so_number_key": ""},
                {"or_ow": {"OR-00001"}, "sales_order": set()},
            )
        )
        self.assertTrue(
            allocation_excel.priority_reference_matches_reservation(
                {"or_number_key": "", "so_number_key": "12345"},
                {"or_ow": set(), "sales_order": {"12345"}},
            )
        )

    def test_temporary_engine_configuration_restores_state_after_an_error(self):
        engine = types.SimpleNamespace(
            EXCEL_PATHS={"source": Path("original.xlsx")},
            OUTPUT_XLSX_PATH=Path("original-output.xlsx"),
            PROGRESS_CALLBACK="original callback",
        )

        with self.assertRaisesRegex(RuntimeError, "expected failure"):
            with sales_operations_app.temporary_engine_configuration(
                engine,
                {"source": Path("temporary.xlsx")},
                "temporary-output.xlsx",
                "temporary callback",
            ):
                self.assertEqual(engine.EXCEL_PATHS, {"source": Path("temporary.xlsx")})
                raise RuntimeError("expected failure")

        self.assertEqual(engine.EXCEL_PATHS, {"source": Path("original.xlsx")})
        self.assertEqual(engine.OUTPUT_XLSX_PATH, Path("original-output.xlsx"))
        self.assertEqual(engine.PROGRESS_CALLBACK, "original callback")
