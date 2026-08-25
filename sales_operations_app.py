import contextlib
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
import json
import os
from pathlib import Path
import queue
import re
import threading
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk

import asignaciones_excel as allocation_engine
import allocation_excel as reservation_allocation_engine
import check_free_cars_excel as check_free_cars_engine
import dealer_stock_excel as dealer_stock_engine
import leads_analysis_excel as leads_analysis_engine
import popup_positioning
import vehicle_preallocation_excel as vehicle_preallocation_engine


engine = allocation_engine

APP_VERSION = "2.1"
APP_TITLE = f"Sales Operations Tool {APP_VERSION}"
CONFIG_FILENAME = "sales_operations_tool_local_config.json"
PROCESS_ALLOCATION = "allocation"
PROCESS_RESERVATION_ALLOCATION = "reservation_allocation"
PROCESS_DEALER_STOCK = "dealer_stock"
PROCESS_VEHICLE_PREALLOCATION = "vehicle_preallocation"
PROCESS_CHECK_FREE_CARS = "check_free_cars"
PROCESS_LEADS_ANALYSIS = "leads_analysis"
PROCESS_LABELS = {
    PROCESS_ALLOCATION: "Preallocation",
    PROCESS_RESERVATION_ALLOCATION: "Vehicle Allocation",
    PROCESS_DEALER_STOCK: "Dealer Stock",
    PROCESS_VEHICLE_PREALLOCATION: "Vehicle Preallocation",
    PROCESS_CHECK_FREE_CARS: "Check Free Cars",
    PROCESS_LEADS_ANALYSIS: "Leads Analysis",
}
LABEL_TO_PROCESS = {label: key for key, label in PROCESS_LABELS.items()}
SHARED_FILE_SPECS = [
    ("sototal", "SO Enquiry by VC", "sototal.xlsx"),
    ("mc_norm", "Material codes", "material code.xlsx"),
    ("dealer_info", "Dealer info", "Dealer Info.xlsx"),
    ("not_allocated", "Cars not allocated", "Cars not allocated.xlsx"),
    ("vehicle_tracking", "Vehicle tracking", "VehicleTracking.xlsx"),
    ("newport", "New port", "NEWport.xlsx"),
    ("reservation", "Reservations", "Vehicle_Reservation.xlsx"),
    ("priority_orders", "Priority orders", "quick allocate.xlsx"),
    ("orders", "Dealer orders", "Dealer Orders.xlsx"),
    ("registration", "Registration", "ANT.xlsx"),
    ("unavailable", "Unavailable", "unavailable.xlsx"),
    ("derogation", "Derogation list", "derogation_list.xlsx"),
    ("logistics_db", "Logistics database", "BASE DE DATOS LOGISTICA.xlsx"),
    ("leads_sp", "Leads Spain", "Leads_SP.csv"),
    ("leads_pt", "Leads Portugal", "Leads_PT.csv"),
    ("model_eq", "Model equivalence", "model_eq.xlsx"),
]
SHARED_FILE_SPECS_BY_KEY = {key: (key, label, file_name) for key, label, file_name in SHARED_FILE_SPECS}
ALLOCATION_FILE_KEYS = [
    "sototal",
    "mc_norm",
    "dealer_info",
    "not_allocated",
    "vehicle_tracking",
    "newport",
    "logistics_db",
    "reservation",
    "priority_orders",
]
RESERVATION_ALLOCATION_FILE_KEYS = [
    "reservation",
    "vehicle_tracking",
    "mc_norm",
    "newport",
    "logistics_db",
    "priority_orders",
]
DEALER_STOCK_FILE_KEYS = [
    "vehicle_tracking",
    "mc_norm",
    "dealer_info",
    "not_allocated",
    "orders",
    "registration",
    "unavailable",
    "derogation",
    "reservation",
    "logistics_db",
]
VEHICLE_PREALLOCATION_FILE_KEYS = [
    "reservation",
    "vehicle_tracking",
    "mc_norm",
    "logistics_db",
]
CHECK_FREE_CARS_FILE_KEYS = [
    "vehicle_tracking",
    "mc_norm",
    "not_allocated",
    "reservation",
    "unavailable",
    "logistics_db",
]
LEADS_ANALYSIS_FILE_KEYS = ["leads_sp", "leads_pt", "model_eq"]
PROCESS_CONFIGS = {
    PROCESS_ALLOCATION: {
        "label": PROCESS_LABELS[PROCESS_ALLOCATION],
        "output_filename": "preallocation_result.xlsx",
        "file_keys": ALLOCATION_FILE_KEYS,
        "engine_name": "preallocation",
    },
    PROCESS_RESERVATION_ALLOCATION: {
        "label": PROCESS_LABELS[PROCESS_RESERVATION_ALLOCATION],
        "output_filename": "vehicle_allocation_result.xlsx",
        "file_keys": RESERVATION_ALLOCATION_FILE_KEYS,
        "engine_name": "allocation",
    },
    PROCESS_DEALER_STOCK: {
        "label": PROCESS_LABELS[PROCESS_DEALER_STOCK],
        "output_filename": "dealer_stock_result.xlsx",
        "file_keys": DEALER_STOCK_FILE_KEYS,
        "engine_name": "dealer stock",
    },
    PROCESS_VEHICLE_PREALLOCATION: {
        "label": PROCESS_LABELS[PROCESS_VEHICLE_PREALLOCATION],
        "output_filename": "vehicle_preallocation_result.xlsx",
        "file_keys": VEHICLE_PREALLOCATION_FILE_KEYS,
        "engine_name": "vehicle preallocation",
    },
    PROCESS_CHECK_FREE_CARS: {
        "label": PROCESS_LABELS[PROCESS_CHECK_FREE_CARS],
        "output_filename": "check_free_cars_result.xlsx",
        "file_keys": CHECK_FREE_CARS_FILE_KEYS,
        "engine_name": "check free cars",
    },
    PROCESS_LEADS_ANALYSIS: {
        "label": PROCESS_LABELS[PROCESS_LEADS_ANALYSIS],
        "output_filename": "leads_analysis_result.xlsx",
        "file_keys": LEADS_ANALYSIS_FILE_KEYS,
        "engine_name": "leads analysis",
    },
}
FILE_SPECS = SHARED_FILE_SPECS
LEADS_SOURCE_FILE_KEYS = set(LEADS_ANALYSIS_FILE_KEYS)
HIDDEN_RESULT_SECTIONS = {"MATERIAL_CODE_MISMATCHES", "FINAL_MATCHES"}
LOG_SECTION_HEADERS = {"INPUT_SUMMARY", "PROCESS_SUMMARY"}
WARNING_PREFIXES = ("WARNING:", "SOFT WARNING:")
MANUAL_SETTING_KEYS = [
    "fixed_material_code_sales_orders",
    "step4_blocked_p_series",
    "excluded_created_by_users",
    "excluded_dsn_contains",
    "model_year_cut_rules",
    "dealer_stock_old_derogation_cutoff",
]
PREALLOCATION_MODE_LABELS = {
    "eta_days": "ETA within X days",
    "gate_in": "Gate in only",
}
PREALLOCATION_LABEL_TO_MODE = {
    label: mode
    for mode, label in PREALLOCATION_MODE_LABELS.items()
}


class FilteredRunLog:
    """Collect only the log content that the application displays."""

    def __init__(self):
        self._parts = []
        self._pending = ""
        self._skipping_result_section = False

    def write(self, text):
        text = str(text)
        self._pending += text
        while "\n" in self._pending:
            line, self._pending = self._pending.split("\n", 1)
            self._append_line(line, newline=True)
        return len(text)

    def flush(self):
        return None

    def getvalue(self):
        if self._pending:
            self._append_line(self._pending, newline=False)
            self._pending = ""
        return "".join(self._parts)

    def _append_line(self, line, newline):
        stripped = line.strip()
        if stripped in HIDDEN_RESULT_SECTIONS:
            self._skipping_result_section = True
            return

        if self._skipping_result_section:
            if not stripped:
                self._skipping_result_section = False
            return

        self._parts.append(line)
        if newline:
            self._parts.append("\n")


def config_path():
    appdata = os.environ.get("APPDATA")
    base = Path(appdata) if appdata else Path.home() / ".sales_operations_tool"
    return base / CONFIG_FILENAME


def default_manual_settings_text():
    return {
        "fixed_material_code_sales_orders": "\n".join(sorted(engine.FIXED_MATERIAL_CODE_SALES_ORDERS)),
        "step4_blocked_p_series": "\n".join(sorted(engine.STEP4_BLOCKED_P_SERIES)),
        "excluded_created_by_users": "\n".join(sorted(engine.EXCLUDED_CREATED_BY_USERS)),
        "excluded_dsn_contains": "\n".join(sorted(engine.EXCLUDED_DSN_CONTAINS)),
        "model_year_cut_rules": "\n".join(
            (
                f"{rule['p_series']};{rule['aggr_contains']};"
                f"{rule['min_new_model_year']}"
            )
            for rule in engine.MODEL_YEAR_CUT_RULES
        ),
        "dealer_stock_old_derogation_cutoff": (
            dealer_stock_engine.OLD_DEROGATION_CUTOFF.strftime("%d/%m/%Y")
            if dealer_stock_engine.OLD_DEROGATION_CUTOFF is not None
            else ""
        ),
    }


def default_preallocation_settings():
    return {
        "mode": engine.PREALLOCATION_WINDOW_MODE,
        "days": str(engine.PREALLOCATION_WINDOW_DAYS),
    }


def normalize_preallocation_settings(settings):
    defaults = default_preallocation_settings()
    if not isinstance(settings, dict):
        return defaults

    mode = settings.get("mode", defaults["mode"])
    if mode not in PREALLOCATION_MODE_LABELS:
        mode = defaults["mode"]

    try:
        days = int(settings.get("days", defaults["days"]))
    except (TypeError, ValueError):
        days = int(defaults["days"])
    days = max(days, 0)

    return {
        "mode": mode,
        "days": str(days),
    }


def clean_expected_file_name(value):
    file_name = str(value or "").strip().strip('"')
    if not file_name:
        raise ValueError("Expected file name cannot be empty.")

    path = Path(file_name)
    if path.is_absolute() or path.name != file_name:
        raise ValueError("Enter only a file name, not a folder or full path.")

    return file_name


def split_manual_values(text):
    return [value.strip() for value in re.split(r"[\n,;]+", text) if value.strip()]


def normalize_manual_settings_text(settings):
    defaults = default_manual_settings_text()
    normalized = defaults.copy()
    for key in MANUAL_SETTING_KEYS:
        if key in settings and isinstance(settings[key], str):
            normalized[key] = settings[key]

    normalized["excluded_created_by_users"] = normalize_excluded_created_by_text(
        normalized["excluded_created_by_users"]
    )
    return normalized


def created_by_exclusion_key(value):
    return re.sub(r"[^A-Z0-9]", "", value.strip().upper())


def normalize_excluded_created_by_text(text):
    default_labels_by_key = {
        created_by_exclusion_key(value): value
        for value in split_manual_values(default_manual_settings_text()["excluded_created_by_users"])
    }
    seen = set()
    values = []

    for value in split_manual_values(text):
        key = created_by_exclusion_key(value)
        if key in seen:
            continue

        seen.add(key)
        values.append(default_labels_by_key.get(key, value))

    return "\n".join(values)


def parse_model_year_cut_rules(text):
    rules = []
    seen_rules = set()
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        separator = ";" if ";" in line else ","
        parts = [part.strip() for part in line.split(separator)]
        if len(parts) != 3 or any(not part for part in parts):
            raise ValueError(
                "Model year cut rules must use one rule per line: "
                "p_series;aggr_contains;min_new_model_year "
                f"(line {line_number}: {raw_line})"
            )

        year_text = parts[2].upper()
        if year_text.startswith("MY"):
            year_text = year_text[2:]

        try:
            min_new_model_year = Decimal(year_text)
        except InvalidOperation as exc:
            raise ValueError(
                f"Model year cut rule line {line_number} has invalid model year: {parts[2]}"
            ) from exc

        rule_key = (parts[0].upper(), parts[1].upper())
        if rule_key in seen_rules:
            raise ValueError(
                "Duplicated model year cut rule: "
                f"{parts[0]};{parts[1]} (line {line_number})"
            )
        seen_rules.add(rule_key)

        rules.append(
            {
                "p_series": parts[0].upper(),
                "aggr_contains": parts[1].upper(),
                "min_new_model_year": min_new_model_year,
            }
        )

    return rules


def parse_manual_settings(settings_text):
    return {
        "fixed_material_code_sales_orders": split_manual_values(
            settings_text["fixed_material_code_sales_orders"]
        ),
        "step4_blocked_p_series": split_manual_values(settings_text["step4_blocked_p_series"]),
        "excluded_created_by_users": split_manual_values(settings_text["excluded_created_by_users"]),
        "excluded_dsn_contains": split_manual_values(settings_text["excluded_dsn_contains"]),
        "model_year_cut_rules": parse_model_year_cut_rules(settings_text["model_year_cut_rules"]),
    }


def parse_dealer_stock_settings(settings_text):
    cutoff_text = settings_text["dealer_stock_old_derogation_cutoff"].strip()
    if not cutoff_text:
        return {"old_derogation_cutoff": None}

    for date_format in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            cutoff = datetime.strptime(cutoff_text, date_format).date()
            break
        except ValueError:
            cutoff = None
    if cutoff is None:
        raise ValueError("Old derogation cutoff must use dd/mm/yyyy.")

    return {
        "old_derogation_cutoff": cutoff,
    }


def apply_manual_settings(settings):
    engine.FIXED_MATERIAL_CODE_SALES_ORDERS = set(settings["fixed_material_code_sales_orders"])
    engine.FIXED_MATERIAL_CODE_SALES_ORDER_KEYS = {
        str(sales_order).strip().upper()
        for sales_order in engine.FIXED_MATERIAL_CODE_SALES_ORDERS
    }

    engine.STEP4_BLOCKED_P_SERIES = set(settings["step4_blocked_p_series"])
    engine.STEP4_BLOCKED_P_SERIES_KEYS = {
        str(p_series).strip().upper()
        for p_series in engine.STEP4_BLOCKED_P_SERIES
    }

    engine.EXCLUDED_CREATED_BY_USERS = set(settings["excluded_created_by_users"])
    engine.EXCLUDED_CREATED_BY_KEYS = {
        re.sub(r"[^A-Z0-9]", "", str(user).strip().upper())
        for user in engine.EXCLUDED_CREATED_BY_USERS
    }

    engine.EXCLUDED_DSN_CONTAINS = set(settings["excluded_dsn_contains"])
    engine.EXCLUDED_DSN_CONTAINS_KEYS = {
        str(dsn_text).strip().upper()
        for dsn_text in engine.EXCLUDED_DSN_CONTAINS
        if str(dsn_text).strip()
    }

    engine.MODEL_YEAR_CUT_RULES = list(settings["model_year_cut_rules"])


def apply_dealer_stock_settings(settings):
    dealer_stock_engine.OLD_DEROGATION_CUTOFF = settings["old_derogation_cutoff"]


def run_engine(file_paths, output_path, manual_settings=None, preallocation_settings=None, progress_callback=None):
    old_paths = engine.EXCEL_PATHS.copy()
    old_output = engine.OUTPUT_XLSX_PATH
    old_preallocation_mode = engine.PREALLOCATION_WINDOW_MODE
    old_preallocation_days = engine.PREALLOCATION_WINDOW_DAYS
    old_non_or_eta_limit = engine.NON_OR_ETA_BEFORE_LIMIT
    old_reservation_eta_limit = engine.RESERVATION_ETA_LIMIT
    old_progress_callback = engine.PROGRESS_CALLBACK
    old_fixed_orders = engine.FIXED_MATERIAL_CODE_SALES_ORDERS.copy()
    old_fixed_order_keys = engine.FIXED_MATERIAL_CODE_SALES_ORDER_KEYS.copy()
    old_blocked_p_series = engine.STEP4_BLOCKED_P_SERIES.copy()
    old_blocked_p_series_keys = engine.STEP4_BLOCKED_P_SERIES_KEYS.copy()
    old_excluded_created_by = engine.EXCLUDED_CREATED_BY_USERS.copy()
    old_excluded_created_by_keys = engine.EXCLUDED_CREATED_BY_KEYS.copy()
    old_excluded_dsn_contains = engine.EXCLUDED_DSN_CONTAINS.copy()
    old_excluded_dsn_contains_keys = engine.EXCLUDED_DSN_CONTAINS_KEYS.copy()
    old_model_year_cut_rules = [rule.copy() for rule in engine.MODEL_YEAR_CUT_RULES]

    try:
        mapped_paths = {
            "sototal": Path(file_paths["sototal"]),
            "mc_norm": Path(file_paths["mc_norm"]),
            "dealer_info": Path(file_paths["dealer_info"]),
            "not_allocated": Path(file_paths["not_allocated"]),
            "vehicle_tracking": Path(file_paths["vehicle_tracking"]),
            "newport": Path(file_paths["newport"]),
            "logistics_db": Path(file_paths["logistics_db"]),
            "reservations": Path(file_paths["reservation"]),
            "priority_orders": Path(file_paths["priority_orders"]),
        }
        engine.EXCEL_PATHS.clear()
        engine.EXCEL_PATHS.update(mapped_paths)
        engine.OUTPUT_XLSX_PATH = Path(output_path)
        engine.PROGRESS_CALLBACK = progress_callback
        if manual_settings is not None:
            apply_manual_settings(manual_settings)

        preallocation_settings = normalize_preallocation_settings(preallocation_settings)
        engine.configure_preallocation_window(
            mode=preallocation_settings["mode"],
            days=int(preallocation_settings["days"]),
            today=date.today(),
        )

        engine.main()
    finally:
        engine.EXCEL_PATHS.clear()
        engine.EXCEL_PATHS.update(old_paths)
        engine.OUTPUT_XLSX_PATH = old_output
        engine.PREALLOCATION_WINDOW_MODE = old_preallocation_mode
        engine.PREALLOCATION_WINDOW_DAYS = old_preallocation_days
        engine.NON_OR_ETA_BEFORE_LIMIT = old_non_or_eta_limit
        engine.RESERVATION_ETA_LIMIT = old_reservation_eta_limit
        engine.PROGRESS_CALLBACK = old_progress_callback
        engine.FIXED_MATERIAL_CODE_SALES_ORDERS = old_fixed_orders
        engine.FIXED_MATERIAL_CODE_SALES_ORDER_KEYS = old_fixed_order_keys
        engine.STEP4_BLOCKED_P_SERIES = old_blocked_p_series
        engine.STEP4_BLOCKED_P_SERIES_KEYS = old_blocked_p_series_keys
        engine.EXCLUDED_CREATED_BY_USERS = old_excluded_created_by
        engine.EXCLUDED_CREATED_BY_KEYS = old_excluded_created_by_keys
        engine.EXCLUDED_DSN_CONTAINS = old_excluded_dsn_contains
        engine.EXCLUDED_DSN_CONTAINS_KEYS = old_excluded_dsn_contains_keys
        engine.MODEL_YEAR_CUT_RULES = old_model_year_cut_rules


def run_dealer_stock_engine(file_paths, output_path, dealer_stock_settings=None, progress_callback=None):
    old_paths = dealer_stock_engine.EXCEL_PATHS.copy()
    old_output = dealer_stock_engine.OUTPUT_XLSX_PATH
    old_progress_callback = dealer_stock_engine.PROGRESS_CALLBACK
    old_derogation_cutoff = dealer_stock_engine.OLD_DEROGATION_CUTOFF

    try:
        mapped_paths = {
            "vehicle_tracking": Path(file_paths["vehicle_tracking"]),
            "mc_norm": Path(file_paths["mc_norm"]),
            "dealer_info": Path(file_paths["dealer_info"]),
            "not_allocated": Path(file_paths["not_allocated"]),
            "orders": Path(file_paths["orders"]),
            "registration": Path(file_paths["registration"]),
            "unavailable": Path(file_paths["unavailable"]),
            "derogation": Path(file_paths["derogation"]),
            "reservation": Path(file_paths["reservation"]),
            "in_port": Path(file_paths["logistics_db"]),
            "already_gate_out": Path(file_paths["logistics_db"]),
        }
        dealer_stock_engine.EXCEL_PATHS.clear()
        dealer_stock_engine.EXCEL_PATHS.update(mapped_paths)
        dealer_stock_engine.OUTPUT_XLSX_PATH = Path(output_path)
        dealer_stock_engine.PROGRESS_CALLBACK = progress_callback
        if dealer_stock_settings is not None:
            apply_dealer_stock_settings(dealer_stock_settings)
        dealer_stock_engine.main()
    finally:
        dealer_stock_engine.EXCEL_PATHS.clear()
        dealer_stock_engine.EXCEL_PATHS.update(old_paths)
        dealer_stock_engine.OUTPUT_XLSX_PATH = old_output
        dealer_stock_engine.PROGRESS_CALLBACK = old_progress_callback
        dealer_stock_engine.OLD_DEROGATION_CUTOFF = old_derogation_cutoff


def run_vehicle_preallocation_engine(file_paths, output_path, progress_callback=None):
    old_paths = vehicle_preallocation_engine.EXCEL_PATHS.copy()
    old_output = vehicle_preallocation_engine.OUTPUT_XLSX_PATH
    old_progress_callback = vehicle_preallocation_engine.PROGRESS_CALLBACK

    try:
        mapped_paths = {
            "reservation": Path(file_paths["reservation"]),
            "vehicle_tracking": Path(file_paths["vehicle_tracking"]),
            "mc_norm": Path(file_paths["mc_norm"]),
            "logistics_db": Path(file_paths["logistics_db"]),
        }
        vehicle_preallocation_engine.EXCEL_PATHS.clear()
        vehicle_preallocation_engine.EXCEL_PATHS.update(mapped_paths)
        vehicle_preallocation_engine.OUTPUT_XLSX_PATH = Path(output_path)
        vehicle_preallocation_engine.PROGRESS_CALLBACK = progress_callback
        vehicle_preallocation_engine.main()
    finally:
        vehicle_preallocation_engine.EXCEL_PATHS.clear()
        vehicle_preallocation_engine.EXCEL_PATHS.update(old_paths)
        vehicle_preallocation_engine.OUTPUT_XLSX_PATH = old_output
        vehicle_preallocation_engine.PROGRESS_CALLBACK = old_progress_callback


def run_check_free_cars_engine(file_paths, output_path, progress_callback=None):
    old_paths = check_free_cars_engine.EXCEL_PATHS.copy()
    old_output = check_free_cars_engine.OUTPUT_XLSX_PATH
    old_progress_callback = check_free_cars_engine.PROGRESS_CALLBACK

    try:
        mapped_paths = {
            "vehicle_tracking": Path(file_paths["vehicle_tracking"]),
            "mc_norm": Path(file_paths["mc_norm"]),
            "not_allocated": Path(file_paths["not_allocated"]),
            "reservation": Path(file_paths["reservation"]),
            "unavailable": Path(file_paths["unavailable"]),
            "logistics_db": Path(file_paths["logistics_db"]),
        }
        check_free_cars_engine.EXCEL_PATHS.clear()
        check_free_cars_engine.EXCEL_PATHS.update(mapped_paths)
        check_free_cars_engine.OUTPUT_XLSX_PATH = Path(output_path)
        check_free_cars_engine.PROGRESS_CALLBACK = progress_callback
        check_free_cars_engine.main()
    finally:
        check_free_cars_engine.EXCEL_PATHS.clear()
        check_free_cars_engine.EXCEL_PATHS.update(old_paths)
        check_free_cars_engine.OUTPUT_XLSX_PATH = old_output
        check_free_cars_engine.PROGRESS_CALLBACK = old_progress_callback


def run_leads_analysis_engine(
    source_paths,
    output_path,
    start_date,
    end_date,
    progress_callback=None,
    test_drive_formula=leads_analysis_engine.DEFAULT_TEST_DRIVE_FORMULA,
):
    old_paths = leads_analysis_engine.EXCEL_PATHS.copy()
    old_output = leads_analysis_engine.OUTPUT_XLSX_PATH
    old_progress_callback = leads_analysis_engine.PROGRESS_CALLBACK

    try:
        leads_analysis_engine.EXCEL_PATHS.clear()
        leads_analysis_engine.EXCEL_PATHS.update(
            {
                "leads_sp": Path(source_paths["leads_sp"]),
                "leads_pt": Path(source_paths["leads_pt"]),
                "model_eq": Path(source_paths["model_eq"]),
            }
        )
        leads_analysis_engine.OUTPUT_XLSX_PATH = Path(output_path)
        leads_analysis_engine.PROGRESS_CALLBACK = progress_callback
        leads_analysis_engine.main(start_date, end_date, test_drive_formula)
    finally:
        leads_analysis_engine.EXCEL_PATHS.clear()
        leads_analysis_engine.EXCEL_PATHS.update(old_paths)
        leads_analysis_engine.OUTPUT_XLSX_PATH = old_output
        leads_analysis_engine.PROGRESS_CALLBACK = old_progress_callback


def run_reservation_allocation_engine(
    file_paths,
    output_path,
    preallocation_settings=None,
    progress_callback=None,
):
    old_paths = reservation_allocation_engine.EXCEL_PATHS.copy()
    old_output = reservation_allocation_engine.OUTPUT_XLSX_PATH
    old_progress_callback = reservation_allocation_engine.PROGRESS_CALLBACK
    old_window_mode = reservation_allocation_engine.ALLOCATION_WINDOW_MODE
    old_window_days = reservation_allocation_engine.ALLOCATION_WINDOW_DAYS
    old_eta_limit = reservation_allocation_engine.ALLOCATION_ETA_LIMIT

    try:
        mapped_paths = {
            "reservation": Path(file_paths["reservation"]),
            "vehicle_tracking": Path(file_paths["vehicle_tracking"]),
            "mc_norm": Path(file_paths["mc_norm"]),
            "newport": Path(file_paths["newport"]),
            "logistics_db": Path(file_paths["logistics_db"]),
            "priority_orders": Path(file_paths["priority_orders"]),
        }
        reservation_allocation_engine.EXCEL_PATHS.clear()
        reservation_allocation_engine.EXCEL_PATHS.update(mapped_paths)
        reservation_allocation_engine.OUTPUT_XLSX_PATH = Path(output_path)
        reservation_allocation_engine.PROGRESS_CALLBACK = progress_callback
        preallocation_settings = normalize_preallocation_settings(preallocation_settings)
        reservation_allocation_engine.configure_allocation_window(
            mode=preallocation_settings["mode"],
            days=int(preallocation_settings["days"]),
            today=date.today(),
        )
        reservation_allocation_engine.main()
    finally:
        reservation_allocation_engine.EXCEL_PATHS.clear()
        reservation_allocation_engine.EXCEL_PATHS.update(old_paths)
        reservation_allocation_engine.OUTPUT_XLSX_PATH = old_output
        reservation_allocation_engine.PROGRESS_CALLBACK = old_progress_callback
        reservation_allocation_engine.ALLOCATION_WINDOW_MODE = old_window_mode
        reservation_allocation_engine.ALLOCATION_WINDOW_DAYS = old_window_days
        reservation_allocation_engine.ALLOCATION_ETA_LIMIT = old_eta_limit


def filter_result_sections(log):
    filtered_lines = []
    skipping_result_section = False

    for line in log.splitlines():
        stripped = line.strip()
        if stripped in HIDDEN_RESULT_SECTIONS:
            skipping_result_section = True
            continue

        if skipping_result_section:
            if stripped == "":
                skipping_result_section = False
            continue

        filtered_lines.append(line)

    if not filtered_lines:
        return ""

    return "\n".join(filtered_lines).rstrip() + "\n"


def filter_app_log(log, include_warning_details=False):
    log_without_result_sections = filter_result_sections(log)
    if include_warning_details:
        return log_without_result_sections

    filtered_lines = []
    skipping_warning_details = False
    for line in log_without_result_sections.splitlines():
        stripped = line.strip()
        if is_warning_line(stripped):
            filtered_lines.append(line)
            skipping_warning_details = True
            continue

        if skipping_warning_details:
            if is_log_boundary(stripped):
                skipping_warning_details = False
            else:
                continue

        filtered_lines.append(line)

    if not filtered_lines:
        return ""

    return "\n".join(filtered_lines).rstrip() + "\n"


def is_warning_line(line):
    return line.strip().startswith(WARNING_PREFIXES)


def is_log_boundary(line):
    stripped = line.strip()
    return (
        stripped == ""
        or stripped in LOG_SECTION_HEADERS
        or stripped.startswith("OUTPUT_XLSX:")
        or is_warning_line(stripped)
    )


def selectable_filetypes_for_key(key):
    if key in LEADS_SOURCE_FILE_KEYS:
        return [
            ("Excel and CSV files", "*.xlsx *.xlsm *.xltx *.xltm *.csv"),
            ("All files", "*.*"),
        ]
    return [("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")]


def extract_warning_blocks(log):
    blocks = []
    current_block = None

    for line in filter_app_log(log, include_warning_details=True).splitlines():
        stripped = line.strip()
        if is_warning_line(stripped):
            if current_block is not None:
                blocks.append(current_block)
            current_block = {"header": line, "details": []}
            continue

        if current_block is None:
            continue

        if is_log_boundary(stripped):
            blocks.append(current_block)
            current_block = None
            continue

        current_block["details"].append(line)

    if current_block is not None:
        blocks.append(current_block)

    return blocks


class SalesOperationsApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x700")
        self.minsize(900, 560)
        self.configure_style()

        self.worker = None
        self.progress_queue = queue.Queue()
        self.working_folder_var = tk.StringVar()
        self.process_var = tk.StringVar(value=PROCESS_ALLOCATION)
        self.process_label_var = tk.StringVar(value=PROCESS_LABELS[PROCESS_ALLOCATION])
        self.status_var = tk.StringVar(value="Ready")
        self.warning_summary_var = tk.StringVar(value="No warnings.")
        self.overrides = {}
        self.file_names = {}
        self.manual_settings_text = default_manual_settings_text()
        self.preallocation_settings = default_preallocation_settings()
        self.leads_start_date = ""
        self.leads_end_date = ""
        self.leads_test_drive_formula = leads_analysis_engine.DEFAULT_TEST_DRIVE_FORMULA
        self.setting_text_widgets = {}
        self.preallocation_mode_var = None
        self.preallocation_days_var = None
        self.dealer_stock_old_derogation_cutoff_var = None
        self.warning_text = None
        self.file_action_buttons = []

        self.load_config()
        self.process_label_var.set(self.current_process_label())
        self.build_ui()
        self.clear_warning_details()
        self.refresh_file_table()

    def configure_style(self):
        for font_name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            font = tkfont.nametofont(font_name)
            size = int(font.cget("size"))
            if 0 < size < 10:
                font.configure(size=10)

        style = ttk.Style(self)
        style.configure("TNotebook", tabmargins=(8, 6, 8, 0))
        style.configure("TNotebook.Tab", padding=(18, 8))
        style.configure("TButton", padding=(10, 5))
        style.configure("Treeview", rowheight=26)
        style.configure("Treeview.Heading", padding=(6, 6))

    def load_config(self):
        path = config_path()
        default_folder = ""
        if not path.exists():
            self.working_folder_var.set(default_folder)
            return

        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            self.working_folder_var.set(default_folder)
            return

        self.working_folder_var.set(data.get("working_folder") or default_folder)
        selected_process = data.get("selected_process")
        if selected_process in PROCESS_CONFIGS:
            self.process_var.set(selected_process)
        self.manual_settings_text = normalize_manual_settings_text(data.get("manual_settings", {}))
        self.preallocation_settings = normalize_preallocation_settings(data.get("preallocation_window", {}))
        self.leads_start_date = str(data.get("leads_start_date") or "").strip()
        self.leads_end_date = str(data.get("leads_end_date") or "").strip()
        self.leads_test_drive_formula = (
            str(data.get("leads_test_drive_formula") or "").strip()
            or leads_analysis_engine.DEFAULT_TEST_DRIVE_FORMULA
        )
        self.overrides = self.normalize_overrides(data.get("overrides", {}))
        self.file_names = self.normalize_file_names(data.get("file_names", {}))

    def normalize_overrides(self, overrides):
        normalized = {}
        if not isinstance(overrides, dict):
            return normalized

        valid_keys = {spec[0] for spec in SHARED_FILE_SPECS}
        legacy_key_map = {
            "reservations": "reservation",
        }

        def add_override(key, value):
            canonical_key = legacy_key_map.get(key, key)
            if canonical_key in valid_keys and value and canonical_key not in normalized:
                normalized[canonical_key] = value

        if any(process in overrides for process in PROCESS_CONFIGS):
            selected_process = self.process_var.get()
            process_order = [selected_process] + [
                process for process in PROCESS_CONFIGS if process != selected_process
            ]
            for process in process_order:
                process_overrides = overrides.get(process, {})
                if process not in PROCESS_CONFIGS or not isinstance(process_overrides, dict):
                    continue
                for key, value in process_overrides.items():
                    add_override(key, value)
            return normalized

        for key, value in overrides.items():
            add_override(key, value)
        return normalized

    def normalize_file_names(self, file_names):
        normalized = {}
        if not isinstance(file_names, dict):
            return normalized

        valid_keys = {spec[0] for spec in SHARED_FILE_SPECS}
        legacy_key_map = {
            "reservations": "reservation",
        }
        for key, value in file_names.items():
            canonical_key = legacy_key_map.get(key, key)
            if canonical_key not in valid_keys:
                continue

            try:
                file_name = clean_expected_file_name(value)
            except ValueError:
                continue

            if file_name != SHARED_FILE_SPECS_BY_KEY[canonical_key][2]:
                normalized[canonical_key] = file_name

        return normalized

    def save_config(self):
        path = config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            preallocation_window = self.current_preallocation_settings()
        except ValueError:
            preallocation_window = normalize_preallocation_settings(self.preallocation_settings)

        data = {
            "working_folder": self.working_folder_var.get(),
            "selected_process": self.process_var.get(),
            "overrides": self.overrides,
            "file_names": self.file_names,
            "manual_settings": normalize_manual_settings_text(self.current_manual_settings_text()),
            "preallocation_window": preallocation_window,
            "leads_start_date": self.leads_start_date,
            "leads_end_date": self.leads_end_date,
            "leads_test_drive_formula": self.leads_test_drive_formula,
        }
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        folder_frame = ttk.Frame(self, padding=(16, 16, 16, 8))
        folder_frame.grid(row=0, column=0, sticky="ew")
        folder_frame.columnconfigure(1, weight=1)

        ttk.Label(folder_frame, text="Working folder").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.working_folder_var)
        self.folder_entry.grid(row=0, column=1, sticky="ew")
        self.select_folder_button = ttk.Button(
            folder_frame,
            text="Select folder",
            command=self.select_working_folder,
        )
        self.select_folder_button.grid(
            row=0, column=2, padx=(8, 0)
        )

        main_notebook = ttk.Notebook(self)
        main_notebook.grid(row=1, column=0, sticky="nsew", padx=16, pady=(4, 16))

        run_frame = ttk.Frame(main_notebook, padding=12)
        run_frame.columnconfigure(0, weight=1)
        run_frame.rowconfigure(1, weight=1)
        main_notebook.add(run_frame, text="Run")

        run_actions = ttk.Frame(run_frame)
        run_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        run_actions.columnconfigure(5, weight=1)
        ttk.Label(run_actions, text="Process").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.process_combo = ttk.Combobox(
            run_actions,
            textvariable=self.process_label_var,
            values=list(PROCESS_LABELS.values()),
            state="readonly",
            width=18,
        )
        self.process_combo.grid(row=0, column=1, padx=(0, 8))
        self.process_combo.bind("<<ComboboxSelected>>", self.on_process_changed)
        self.run_button = ttk.Button(run_actions, text="Run", command=self.run_allocation)
        self.run_button.grid(row=0, column=2, padx=(0, 8))
        self.open_result_button = ttk.Button(run_actions, text="Open result", command=self.open_result)
        self.open_result_button.grid(row=0, column=3, padx=(0, 8))
        ttk.Label(run_actions, textvariable=self.status_var).grid(row=0, column=5, sticky="e")

        log_frame = ttk.LabelFrame(run_frame, text="Run log", padding=8)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.grid(row=1, column=0, sticky="nsew")

        self.log_text = tk.Text(log_frame, wrap="none", padx=8, pady=6)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        default_font = tkfont.nametofont("TkTextFont")
        warning_font = default_font.copy()
        warning_font.configure(weight="bold")
        self.log_text.tag_configure("warning", foreground="#9a3412", font=warning_font)
        self.log_text.tag_configure("warning_detail", foreground="#7c2d12")
        self.log_text.tag_configure(
            "warning_banner",
            background="#fff3cd",
            foreground="#7c2d12",
            font=warning_font,
            spacing1=4,
            spacing3=4,
        )
        self.log_text.tag_configure("error", foreground="#b91c1c", font=warning_font)
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        warnings_frame = ttk.Frame(main_notebook, padding=12)
        warnings_frame.columnconfigure(0, weight=1)
        warnings_frame.rowconfigure(1, weight=1)
        main_notebook.add(warnings_frame, text="Warnings")

        ttk.Label(warnings_frame, textvariable=self.warning_summary_var).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.warning_text = tk.Text(warnings_frame, wrap="none", padx=8, pady=6)
        self.warning_text.grid(row=1, column=0, sticky="nsew")
        self.warning_text.tag_configure("warning", foreground="#9a3412", font=warning_font)
        self.warning_text.tag_configure("warning_detail", foreground="#7c2d12")
        warning_scrollbar = ttk.Scrollbar(
            warnings_frame,
            orient=tk.VERTICAL,
            command=self.warning_text.yview,
        )
        warning_scrollbar.grid(row=1, column=1, sticky="ns")
        self.warning_text.configure(yscrollcommand=warning_scrollbar.set)

        files_frame = ttk.Frame(main_notebook, padding=12)
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        main_notebook.add(files_frame, text="Files")

        files_actions = ttk.Frame(files_frame)
        files_actions.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        files_actions.columnconfigure(5, weight=1)
        change_file_button = ttk.Button(
            files_actions,
            text="Change selected file",
            command=self.change_selected_file,
        )
        change_file_button.grid(
            row=0, column=0, padx=(0, 8)
        )
        edit_name_button = ttk.Button(
            files_actions,
            text="Edit expected name",
            command=self.edit_expected_name_for_selected,
        )
        edit_name_button.grid(
            row=0, column=1, padx=(0, 8)
        )
        use_default_button = ttk.Button(
            files_actions,
            text="Use default for selected",
            command=self.use_default_for_selected,
        )
        use_default_button.grid(
            row=0, column=2, padx=(0, 8)
        )
        reset_files_button = ttk.Button(
            files_actions,
            text="Reset files from folder",
            command=self.reset_files_from_folder,
        )
        reset_files_button.grid(
            row=0, column=3, padx=(0, 8)
        )
        clear_cache_button = ttk.Button(files_actions, text="Clear cache", command=self.clear_cache)
        clear_cache_button.grid(
            row=0, column=4, padx=(0, 8)
        )
        self.file_action_buttons = [
            change_file_button,
            edit_name_button,
            use_default_button,
            reset_files_button,
            clear_cache_button,
        ]

        columns = ("file", "needed", "expected", "status", "source", "modified", "path")
        self.file_tree = ttk.Treeview(files_frame, columns=columns, show="headings", height=9)
        self.file_tree.heading("file", text="File")
        self.file_tree.heading("needed", text="Needed")
        self.file_tree.heading("expected", text="Expected name")
        self.file_tree.heading("status", text="Status")
        self.file_tree.heading("source", text="Source")
        self.file_tree.heading("modified", text="Modified")
        self.file_tree.heading("path", text="Path")
        self.file_tree.column("file", width=140, minwidth=110, anchor="w")
        self.file_tree.column("needed", width=65, minwidth=60, anchor="center", stretch=False)
        self.file_tree.column("expected", width=165, minwidth=120, anchor="w")
        self.file_tree.column("status", width=78, minwidth=70, anchor="center", stretch=False)
        self.file_tree.column("source", width=95, minwidth=80, anchor="center", stretch=False)
        self.file_tree.column("modified", width=130, minwidth=115, anchor="center", stretch=False)
        self.file_tree.column("path", width=320, minwidth=180, anchor="w")
        self.file_tree.grid(row=1, column=0, sticky="nsew")
        self.file_tree.bind("<Double-1>", lambda _event: self.edit_expected_name_for_selected())

        files_scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        files_scrollbar.grid(row=1, column=1, sticky="ns")
        files_x_scrollbar = ttk.Scrollbar(files_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview)
        files_x_scrollbar.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        self.file_tree.configure(
            yscrollcommand=files_scrollbar.set,
            xscrollcommand=files_x_scrollbar.set,
        )

        settings_frame = ttk.Frame(main_notebook, padding=12)
        settings_frame.columnconfigure(0, weight=1)
        settings_frame.rowconfigure(1, weight=1)
        main_notebook.add(settings_frame, text="Manual settings")

        settings_header = ttk.Frame(settings_frame)
        settings_header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        settings_header.columnconfigure(0, weight=1)
        ttk.Label(
            settings_header,
            text="Values are saved when a process starts. Use one item per line, or comma/semicolon separated.",
        ).grid(row=0, column=0, sticky="w")

        self.settings_notebook = ttk.Notebook(settings_frame)
        self.settings_notebook.grid(row=1, column=0, sticky="nsew")
        self.create_settings_tab(
            "Fixed material",
            "fixed_material_code_sales_orders",
            "Sales orders that must keep the exact same material_code.",
        )
        self.create_settings_tab(
            "Blocked model",
            "step4_blocked_p_series",
            "Non-OR orders for these models are not included.",
        )
        self.create_settings_tab(
            "Excluded users",
            "excluded_created_by_users",
            "Created By users excluded from the clean order pool.",
        )
        self.create_settings_tab(
            "Excluded DSNs",
            "excluded_dsn_contains",
            "Orders whose DSN contains one of these texts are excluded from the clean order pool.",
        )
        self.create_settings_tab(
            "MY cuts",
            "model_year_cut_rules",
            "One rule per line: p_series;aggr_contains;min_new_model_year",
        )
        self.create_dealer_stock_settings_tab()
        self.create_preallocation_settings_tab()

    def create_settings_tab(self, title, key, description):
        frame = ttk.Frame(self.settings_notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        self.settings_notebook.add(frame, text=title)

        ttk.Label(frame, text=description).grid(row=0, column=0, sticky="w", pady=(0, 6))

        text = tk.Text(frame, height=7, wrap="none", padx=8, pady=6)
        text.grid(row=1, column=0, sticky="nsew")
        text.insert("1.0", self.manual_settings_text.get(key, ""))
        self.setting_text_widgets[key] = text

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)

    def create_preallocation_settings_tab(self):
        frame = ttk.Frame(self.settings_notebook, padding=12)
        frame.columnconfigure(1, weight=1)
        self.settings_notebook.add(frame, text="Preallocation")

        ttk.Label(
            frame,
            text=(
                "Controls the availability window used by Preallocation "
                "and Vehicle Allocation."
            ),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.preallocation_mode_var = tk.StringVar(
            value=PREALLOCATION_MODE_LABELS[self.preallocation_settings["mode"]]
        )
        self.preallocation_days_var = tk.StringVar(value=self.preallocation_settings["days"])

        ttk.Label(frame, text="Window").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        self.preallocation_mode_combo = ttk.Combobox(
            frame,
            textvariable=self.preallocation_mode_var,
            values=list(PREALLOCATION_MODE_LABELS.values()),
            state="readonly",
            width=22,
        )
        self.preallocation_mode_combo.grid(row=1, column=1, sticky="w", pady=(0, 8))

        ttk.Label(frame, text="Days").grid(row=2, column=0, sticky="w", padx=(0, 8))
        self.preallocation_days_spinbox = ttk.Spinbox(
            frame,
            from_=0,
            to=90,
            textvariable=self.preallocation_days_var,
            width=8,
        )
        self.preallocation_days_spinbox.grid(row=2, column=1, sticky="w")

        def refresh_days_state(*_):
            mode = PREALLOCATION_LABEL_TO_MODE.get(self.preallocation_mode_var.get(), "eta_days")
            self.preallocation_days_spinbox.configure(
                state=tk.NORMAL if mode == "eta_days" else tk.DISABLED
            )

        self.preallocation_mode_combo.bind("<<ComboboxSelected>>", refresh_days_state)
        self.refresh_preallocation_days_state = refresh_days_state
        refresh_days_state()

    def create_dealer_stock_settings_tab(self):
        frame = ttk.Frame(self.settings_notebook, padding=12)
        frame.columnconfigure(1, weight=1)
        self.settings_notebook.add(frame, text="Old derogations")

        ttk.Label(
            frame,
            text="Dealer Stock only: derogations before this date appear in OLD_DEROGATIONS.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        ttk.Label(frame, text="Cutoff date").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.dealer_stock_old_derogation_cutoff_var = tk.StringVar(
            value=self.manual_settings_text.get("dealer_stock_old_derogation_cutoff", "")
        )
        self.dealer_stock_old_derogation_cutoff_entry = ttk.Entry(
            frame,
            textvariable=self.dealer_stock_old_derogation_cutoff_var,
            width=14,
        )
        self.dealer_stock_old_derogation_cutoff_entry.grid(
            row=1, column=1, sticky="w"
        )
        ttk.Label(frame, text="dd/mm/yyyy").grid(row=1, column=2, sticky="w", padx=(8, 0))

    def current_manual_settings_text(self):
        if not self.setting_text_widgets:
            return self.manual_settings_text

        settings = {
            key: widget.get("1.0", tk.END).strip()
            for key, widget in self.setting_text_widgets.items()
        }
        if self.dealer_stock_old_derogation_cutoff_var is not None:
            settings["dealer_stock_old_derogation_cutoff"] = (
                self.dealer_stock_old_derogation_cutoff_var.get().strip()
            )
        return settings

    def current_preallocation_settings(self):
        if self.preallocation_mode_var is None or self.preallocation_days_var is None:
            return normalize_preallocation_settings(self.preallocation_settings)

        mode = PREALLOCATION_LABEL_TO_MODE.get(self.preallocation_mode_var.get(), "eta_days")
        try:
            days = int(self.preallocation_days_var.get())
        except ValueError as exc:
            raise ValueError("Preallocation days must be a whole number.") from exc
        if days < 0:
            raise ValueError("Preallocation days must be 0 or greater.")

        return normalize_preallocation_settings({"mode": mode, "days": days})

    def collect_manual_settings(self):
        settings_text = normalize_manual_settings_text(self.current_manual_settings_text())
        return parse_manual_settings(settings_text)

    def collect_preallocation_settings(self):
        return self.current_preallocation_settings()

    def collect_dealer_stock_settings(self):
        settings_text = normalize_manual_settings_text(self.current_manual_settings_text())
        return parse_dealer_stock_settings(settings_text)

    def working_folder(self):
        folder_text = self.working_folder_var.get().strip()
        return Path(folder_text).expanduser() if folder_text else None

    def current_process(self):
        process = self.process_var.get()
        return process if process in PROCESS_CONFIGS else PROCESS_ALLOCATION

    def current_process_label(self):
        return PROCESS_CONFIGS[self.current_process()]["label"]

    def current_required_file_keys(self):
        return set(PROCESS_CONFIGS[self.current_process()]["file_keys"])

    def current_overrides(self):
        return self.overrides

    def default_file_name(self, key):
        return SHARED_FILE_SPECS_BY_KEY[key][2]

    def expected_file_name(self, key):
        return self.file_names.get(key, self.default_file_name(key))

    def default_path(self, key):
        file_name = self.expected_file_name(key)
        folder = self.working_folder()
        return (folder / file_name) if folder is not None else Path(file_name)

    def resolved_path(self, key):
        overrides = self.current_overrides()
        return Path(overrides[key]) if key in overrides else self.default_path(key)

    def output_path(self):
        folder = self.working_folder()
        output_filename = PROCESS_CONFIGS[self.current_process()]["output_filename"]
        return (folder / output_filename) if folder is not None else Path(output_filename)

    def on_process_changed(self, *_):
        selected_label = self.process_combo.get()
        selected_process = LABEL_TO_PROCESS.get(selected_label, PROCESS_ALLOCATION)
        self.process_var.set(selected_process)
        self.process_label_var.set(PROCESS_LABELS[selected_process])
        self.status_var.set(f"Ready: {self.current_process_label()}")
        self.save_config()
        self.refresh_file_table()

    def select_working_folder(self):
        current_folder = self.working_folder()
        selected = filedialog.askdirectory(
            title="Select working folder",
            initialdir=(
                str(current_folder)
                if current_folder is not None and current_folder.exists()
                else str(Path.home())
            ),
        )
        if not selected:
            return

        self.working_folder_var.set(selected)
        self.overrides.clear()
        self.save_config()
        self.refresh_file_table()

    def reset_files_from_folder(self):
        self.overrides.clear()
        self.save_config()
        self.refresh_file_table()

    def selected_key(self):
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showinfo(APP_TITLE, "Select a file row first.")
            return None

        return selection[0]

    def change_selected_file(self):
        key = self.selected_key()
        if key is None:
            return

        initial = self.resolved_path(key)
        working_folder = self.working_folder() or Path.home()
        selected = filedialog.askopenfilename(
            title="Select source file",
            initialdir=str(initial.parent if initial.parent.exists() else working_folder),
            filetypes=selectable_filetypes_for_key(key),
        )
        if not selected:
            return

        self.current_overrides()[key] = selected
        self.save_config()
        self.refresh_file_table()

    def edit_expected_name_for_selected(self):
        key = self.selected_key()
        if key is None:
            return

        _, label, _ = SHARED_FILE_SPECS_BY_KEY[key]
        current_name = self.expected_file_name(key)
        new_name = simpledialog.askstring(
            APP_TITLE,
            f"Expected file name for {label}:",
            initialvalue=current_name,
            parent=self,
        )
        if new_name is None:
            return

        try:
            clean_name = clean_expected_file_name(new_name)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return

        if clean_name == self.default_file_name(key):
            self.file_names.pop(key, None)
        else:
            self.file_names[key] = clean_name

        self.current_overrides().pop(key, None)
        self.save_config()
        self.refresh_file_table()

    def use_default_for_selected(self):
        key = self.selected_key()
        if key is None:
            return

        self.current_overrides().pop(key, None)
        self.file_names.pop(key, None)
        self.save_config()
        self.refresh_file_table()

    def clear_cache(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, "Wait until the current run finishes before clearing the cache.")
            return

        removed = engine.clear_vehicle_tracking_cache()
        removed = reservation_allocation_engine.clear_vehicle_tracking_cache() or removed
        removed = dealer_stock_engine.clear_vehicle_tracking_cache() or removed
        removed = vehicle_preallocation_engine.clear_vehicle_tracking_cache() or removed
        message = "VehicleTracking cache cleared." if removed else "VehicleTracking cache was already empty."
        self.append_log_line(message)
        messagebox.showinfo(APP_TITLE, message)

    def set_configuration_controls_enabled(self, enabled):
        ttk_state = tk.NORMAL if enabled else tk.DISABLED
        self.folder_entry.configure(state=ttk_state)
        self.select_folder_button.configure(state=ttk_state)
        for button in self.file_action_buttons:
            button.configure(state=ttk_state)
        self.file_tree.state(["!disabled"] if enabled else ["disabled"])
        self.settings_notebook.state(["!disabled"] if enabled else ["disabled"])

        for text in self.setting_text_widgets.values():
            text.configure(state=tk.NORMAL if enabled else tk.DISABLED)

        if enabled:
            self.preallocation_mode_combo.configure(state="readonly")
            self.refresh_preallocation_days_state()
            self.dealer_stock_old_derogation_cutoff_entry.configure(state=tk.NORMAL)
        else:
            self.preallocation_mode_combo.configure(state=tk.DISABLED)
            self.preallocation_days_spinbox.configure(state=tk.DISABLED)
            self.dealer_stock_old_derogation_cutoff_entry.configure(state=tk.DISABLED)

    def refresh_file_table(self):
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        overrides = self.current_overrides()
        required_keys = self.current_required_file_keys()
        rows = []
        for key, label, _ in SHARED_FILE_SPECS:
            path = self.resolved_path(key)
            exists = path.exists()
            expected_name = self.expected_file_name(key)
            if key in overrides:
                source = "Manual path"
            elif key in self.file_names:
                source = "Custom name"
            else:
                source = "Auto"
            needed = "Yes" if key in required_keys else ""
            modified = ""
            if exists:
                modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
            rows.append(
                (
                    key,
                    (
                        label,
                        needed,
                        expected_name,
                        "Found" if exists else "Missing",
                        source,
                        modified,
                        str(path),
                    ),
                )
            )

        rows.sort(key=lambda row: row[0] not in required_keys)
        path_font = tkfont.nametofont("TkDefaultFont")
        longest_path_width = max(
            (path_font.measure(values[-1]) + 16 for _, values in rows),
            default=320,
        )
        self.file_tree.column("path", width=max(320, longest_path_width))

        for key, values in rows:
            self.file_tree.insert(
                "",
                tk.END,
                iid=key,
                values=values,
            )

        self.open_result_button.configure(state=tk.NORMAL if self.output_path().exists() else tk.DISABLED)

    def validate_files(self):
        if self.working_folder() is None:
            return {}, ["Select a working folder before running the process."]

        missing = []
        paths = {}
        for key in PROCESS_CONFIGS[self.current_process()]["file_keys"]:
            _, label, _ = SHARED_FILE_SPECS_BY_KEY[key]
            path = self.resolved_path(key)
            paths[key] = path
            if not path.exists():
                missing.append(f"{label}: {path}")

        return paths, missing

    def prompt_leads_period(self):
        selected_period = None
        window = tk.Toplevel(self)
        window.title("Leads analysis period")
        window.resizable(False, False)
        window.transient(self)
        window.grab_set()

        ttk.Label(window, text="Start date (dd/mm/yyyy)").grid(
            row=0, column=0, sticky="w", padx=(14, 8), pady=(14, 8)
        )
        start_var = tk.StringVar(value=self.leads_start_date)
        start_entry = ttk.Entry(window, textvariable=start_var, width=18)
        start_entry.grid(row=0, column=1, padx=(0, 14), pady=(14, 8))
        ttk.Label(window, text="End date (dd/mm/yyyy)").grid(
            row=1, column=0, sticky="w", padx=(14, 8), pady=(0, 8)
        )
        end_var = tk.StringVar(value=self.leads_end_date)
        end_entry = ttk.Entry(window, textvariable=end_var, width=18)
        end_entry.grid(row=1, column=1, padx=(0, 14), pady=(0, 8))

        ttk.Label(window, text="Test drive formula").grid(
            row=2, column=0, sticky="w", padx=(14, 8), pady=(0, 6)
        )
        formula_var = tk.StringVar(value=self.leads_test_drive_formula)
        formula_entry = ttk.Entry(window, textvariable=formula_var, width=68)
        formula_entry.grid(row=2, column=1, sticky="ew", padx=(0, 14), pady=(0, 6))

        def insert_formula_token(token):
            try:
                formula_entry.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass
            formula_entry.insert(tk.INSERT, token)
            formula_entry.focus_set()

        ttk.Label(window, text="Allowed variables").grid(
            row=3, column=0, columnspan=2, sticky="w", padx=14, pady=(2, 4)
        )
        variables_frame = ttk.Frame(window)
        variables_frame.grid(row=4, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 8))
        for column, variable in enumerate(leads_analysis_engine.TEST_DRIVE_FORMULA_VARIABLES):
            ttk.Button(
                variables_frame,
                text=variable,
                command=lambda token=variable: insert_formula_token(token),
            ).grid(row=0, column=column, padx=(0, 6))

        ttk.Label(window, text="Allowed operations").grid(
            row=5, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 4)
        )
        operations_frame = ttk.Frame(window)
        operations_frame.grid(row=6, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 10))
        for column, token in enumerate(("+", "-", "*", "/", "(", ")")):
            ttk.Button(
                operations_frame,
                text=token,
                width=4,
                command=lambda token=token: insert_formula_token(token),
            ).grid(row=0, column=column, padx=(0, 4))
        ttk.Button(
            operations_frame,
            text="Clear",
            command=lambda: formula_var.set(""),
        ).grid(row=0, column=6, padx=(4, 0))

        def confirm_period():
            nonlocal selected_period
            try:
                start_date = leads_analysis_engine.parse_period_date(start_var.get())
                end_date = leads_analysis_engine.parse_period_date(end_var.get())
                if start_date > end_date:
                    raise ValueError("Start date cannot be after end date.")
                test_drive_formula, _ = leads_analysis_engine.parse_test_drive_formula(
                    formula_var.get()
                )
            except ValueError as exc:
                messagebox.showerror(APP_TITLE, str(exc), parent=window)
                return

            self.leads_start_date = start_date.strftime("%d/%m/%Y")
            self.leads_end_date = end_date.strftime("%d/%m/%Y")
            self.leads_test_drive_formula = test_drive_formula
            self.save_config()
            selected_period = (start_date, end_date, test_drive_formula)
            window.destroy()

        actions = ttk.Frame(window)
        actions.grid(row=7, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 14))
        ttk.Button(actions, text="Cancel", command=window.destroy).grid(row=0, column=0)
        ttk.Button(actions, text="Run", command=confirm_period).grid(row=0, column=1, padx=(8, 0))
        window.protocol("WM_DELETE_WINDOW", window.destroy)
        popup_positioning.place_over_parent(window, self)
        start_entry.focus_set()
        self.wait_window(window)
        return selected_period

    def run_allocation(self):
        if self.worker and self.worker.is_alive():
            return

        process = self.current_process()
        leads_period = None
        if process == PROCESS_LEADS_ANALYSIS:
            leads_period = self.prompt_leads_period()
            if leads_period is None:
                return

        paths, missing = self.validate_files()
        if missing:
            messagebox.showerror(APP_TITLE, "Missing required files:\n\n" + "\n".join(missing))
            return

        manual_settings = None
        preallocation_settings = None
        dealer_stock_settings = None
        if process in {PROCESS_ALLOCATION, PROCESS_RESERVATION_ALLOCATION}:
            try:
                preallocation_settings = self.collect_preallocation_settings()
                if process == PROCESS_ALLOCATION:
                    manual_settings = self.collect_manual_settings()
            except ValueError as exc:
                messagebox.showerror(APP_TITLE, str(exc))
                return
        elif process == PROCESS_DEALER_STOCK:
            try:
                dealer_stock_settings = self.collect_dealer_stock_settings()
            except ValueError as exc:
                messagebox.showerror(APP_TITLE, str(exc))
                return

        self.save_config()
        self.log_text.delete("1.0", tk.END)
        self.clear_warning_details()
        self.status_var.set(f"Running {self.current_process_label()}...")
        self.run_button.configure(state=tk.DISABLED)
        self.process_combo.configure(state=tk.DISABLED)
        self.open_result_button.configure(state=tk.DISABLED)
        self.set_configuration_controls_enabled(False)

        self.worker = threading.Thread(
            target=self.run_allocation_worker,
            args=(
                process,
                {key: str(path) for key, path in paths.items()},
                str(self.output_path()),
                manual_settings,
                preallocation_settings,
                dealer_stock_settings,
                leads_period,
            ),
            daemon=True,
        )
        self.worker.start()
        self.poll_progress_queue()

    def run_allocation_worker(
        self,
        process,
        paths,
        output_path,
        manual_settings,
        preallocation_settings,
        dealer_stock_settings,
        leads_period,
    ):
        buffer = FilteredRunLog()
        try:
            with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
                if process == PROCESS_ALLOCATION:
                    run_engine(
                        paths,
                        output_path,
                        manual_settings=manual_settings,
                        preallocation_settings=preallocation_settings,
                        progress_callback=self.progress_queue.put,
                    )
                elif process == PROCESS_DEALER_STOCK:
                    run_dealer_stock_engine(
                        paths,
                        output_path,
                        dealer_stock_settings=dealer_stock_settings,
                        progress_callback=self.progress_queue.put,
                    )
                elif process == PROCESS_RESERVATION_ALLOCATION:
                    run_reservation_allocation_engine(
                        paths,
                        output_path,
                        preallocation_settings=preallocation_settings,
                        progress_callback=self.progress_queue.put,
                    )
                elif process == PROCESS_VEHICLE_PREALLOCATION:
                    run_vehicle_preallocation_engine(
                        paths,
                        output_path,
                        progress_callback=self.progress_queue.put,
                    )
                elif process == PROCESS_CHECK_FREE_CARS:
                    run_check_free_cars_engine(
                        paths,
                        output_path,
                        progress_callback=self.progress_queue.put,
                    )
                elif process == PROCESS_LEADS_ANALYSIS:
                    run_leads_analysis_engine(
                        paths,
                        output_path,
                        start_date=leads_period[0],
                        end_date=leads_period[1],
                        test_drive_formula=leads_period[2],
                        progress_callback=self.progress_queue.put,
                    )
                else:
                    raise ValueError(f"Unknown process: {process}")
            result = ("success", process, buffer.getvalue(), output_path)
        except Exception:
            result = ("error", process, buffer.getvalue() + "\n" + traceback.format_exc(), output_path)

        self.after(0, lambda: self.finish_run(result))

    def append_log_line(self, line):
        timestamp = datetime.now().strftime("%H:%M:%S")
        tag = "warning" if is_warning_line(line) else None
        if tag:
            self.log_text.insert(tk.END, f"[{timestamp}] {line}\n", tag)
        else:
            self.log_text.insert(tk.END, f"[{timestamp}] {line}\n")
        self.log_text.see(tk.END)
        self.status_var.set(line)

    def clear_warning_details(self):
        self.warning_summary_var.set("No warnings.")
        if self.warning_text is None:
            return

        self.warning_text.delete("1.0", tk.END)
        self.warning_text.insert(tk.END, "No warnings for the current run.\n")

    def set_warning_details(self, warning_blocks):
        self.warning_text.delete("1.0", tk.END)
        if not warning_blocks:
            self.warning_summary_var.set("No warnings.")
            self.warning_text.insert(tk.END, "No warnings for the current run.\n")
            return

        self.warning_summary_var.set(
            f"{len(warning_blocks)} warning(s). Cases are grouped by warning below."
        )
        for index, block in enumerate(warning_blocks, start=1):
            self.warning_text.insert(
                tk.END,
                f"{index}. {block['header']}\n",
                "warning",
            )
            if block["details"]:
                for detail in block["details"]:
                    self.warning_text.insert(tk.END, f"   {detail}\n", "warning_detail")
            else:
                self.warning_text.insert(tk.END, "   No detail rows were reported.\n")
            self.warning_text.insert(tk.END, "\n")

        self.warning_text.see("1.0")

    def append_log_block(self, log):
        warning_blocks = extract_warning_blocks(log)
        warning_count = len(warning_blocks)
        self.set_warning_details(warning_blocks)

        filtered_log = filter_app_log(log, include_warning_details=False)
        if warning_count:
            self.log_text.insert(
                tk.END,
                f"\nWARNINGS FOUND: {warning_count}. Open the Warnings tab to review cases.\n",
                "warning_banner",
            )

        for line in filtered_log.splitlines():
            stripped = line.strip()
            if is_warning_line(stripped):
                self.log_text.insert(tk.END, line + "\n", "warning")
            elif stripped.startswith(("Traceback", "Error:", "Exception:")):
                self.log_text.insert(tk.END, line + "\n", "error")
            else:
                self.log_text.insert(tk.END, line + "\n")

        self.log_text.see(tk.END)
        return warning_count

    def poll_progress_queue(self):
        while True:
            try:
                message = self.progress_queue.get_nowait()
            except queue.Empty:
                break

            self.append_log_line(message)

        if self.worker and self.worker.is_alive():
            self.after(100, self.poll_progress_queue)

    def finish_run(self, result):
        status, process, log, output_path = result
        process_label = PROCESS_CONFIGS.get(process, PROCESS_CONFIGS[PROCESS_ALLOCATION])["label"]
        self.poll_progress_queue()
        warning_count = 0
        if log:
            warning_count = self.append_log_block(log)
        self.run_button.configure(state=tk.NORMAL)
        self.process_combo.configure(state="readonly")
        self.set_configuration_controls_enabled(True)
        self.refresh_file_table()

        if status == "success":
            if warning_count:
                self.status_var.set(f"Finished with {warning_count} warning(s). Output: {output_path}")
                messagebox.showwarning(
                    APP_TITLE,
                    (
                        f"{process_label} finished with {warning_count} warning(s).\n\n"
                        f"Output file:\n{output_path}\n\n"
                        "Open the Warnings tab to review the cases."
                    ),
                )
            else:
                self.status_var.set(f"Finished. Output: {output_path}")
                messagebox.showinfo(APP_TITLE, f"{process_label} finished.\n\nOutput file:\n{output_path}")
        else:
            self.status_var.set("Failed")
            messagebox.showerror(APP_TITLE, f"{process_label} failed. Check the run log.")

    def open_result(self):
        path = self.output_path()
        if not path.exists():
            messagebox.showinfo(APP_TITLE, "The output file does not exist yet.")
            return

        os.startfile(path)


if __name__ == "__main__":
    SalesOperationsApp().mainloop()
