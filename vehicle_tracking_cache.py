"""Shared, validated cache for normalized VehicleTracking rows."""

from datetime import date, datetime, time, timedelta
import os
import pickle
from pathlib import Path
import tempfile
import unicodedata


CACHE_VERSION = 8
CACHE_SCHEMA = "sales_operations.vehicle_tracking"
RECORD_FIELDS = (
    "vin",
    "material_code",
    "description",
    "eta",
    "port",
    "vessel_name",
    "dsn",
    "sap",
    "gate_in",
    "gate_out",
    "production_date",
    "status",
    "invoice_date",
    "country",
    "address",
    "city",
    "tag",
    "related_order",
    "reserved_so",
    "dn_create_time",
    "allocation_date",
)


class RestrictedCacheUnpickler(pickle.Unpickler):
    """Allow only the date classes used by the normalized cache payload."""

    ALLOWED_GLOBALS = {
        ("datetime", "date"): date,
        ("datetime", "datetime"): datetime,
        ("datetime", "time"): time,
        ("datetime", "timedelta"): timedelta,
    }

    def find_class(self, module, name):
        allowed_class = self.ALLOWED_GLOBALS.get((module, name))
        if allowed_class is None:
            raise pickle.UnpicklingError(f"unsupported cache type: {module}.{name}")
        return allowed_class


def default_cache_dir():
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "Sales Operations Tool" / "cache"
    return Path.home() / ".sales_operations_tool" / "cache"


def vehicle_tracking_file_signature(path):
    path = Path(path)
    stat = path.stat()
    return {
        "path": os.path.normcase(str(path.resolve())),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def clear_cache(cache_path):
    try:
        Path(cache_path).unlink()
        return True
    except FileNotFoundError:
        return False


def load_cache(cache_path, signature, report_progress):
    cache_path = Path(cache_path)
    if not cache_path.exists():
        report_progress("VehicleTracking cache miss: no cache found")
        return None

    try:
        with cache_path.open("rb") as file:
            payload = RestrictedCacheUnpickler(file).load()
    except Exception:
        clear_cache(cache_path)
        report_progress("VehicleTracking cache invalid: cache cleared")
        return None

    if not isinstance(payload, dict):
        clear_cache(cache_path)
        report_progress("VehicleTracking cache invalid: cache cleared")
        return None

    if payload.get("version") != CACHE_VERSION:
        clear_cache(cache_path)
        report_progress("VehicleTracking cache miss: cache version changed, cache cleared")
        return None

    if payload.get("schema") != CACHE_SCHEMA or payload.get("record_fields") != RECORD_FIELDS:
        clear_cache(cache_path)
        report_progress("VehicleTracking cache miss: cache schema changed, cache cleared")
        return None

    if payload.get("signature") != signature:
        clear_cache(cache_path)
        report_progress("VehicleTracking cache miss: source file changed, cache cleared")
        return None

    by_vin = payload.get("by_vin")
    if not _has_expected_record_schema(by_vin):
        clear_cache(cache_path)
        report_progress("VehicleTracking cache invalid: cache cleared")
        return None

    report_progress(f"VehicleTracking cache hit: {len(by_vin):,} VINs loaded from cache")
    return by_vin


def write_cache(cache_path, signature, by_vin, report_progress):
    cache_path = Path(cache_path)
    temporary_path = None
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="wb",
            dir=cache_path.parent,
            prefix=f".{cache_path.stem}.",
            suffix=".tmp",
            delete=False,
        ) as file:
            temporary_path = Path(file.name)
            pickle.dump(
                {
                    "version": CACHE_VERSION,
                    "schema": CACHE_SCHEMA,
                    "record_fields": RECORD_FIELDS,
                    "signature": signature,
                    "by_vin": by_vin,
                },
                file,
                protocol=pickle.HIGHEST_PROTOCOL,
            )
        os.replace(temporary_path, cache_path)
        temporary_path = None
        report_progress("VehicleTracking cache written")
    except Exception:
        report_progress("VehicleTracking cache could not be written; continuing without cache")
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink()
            except FileNotFoundError:
                pass


def _has_expected_record_schema(by_vin):
    if not isinstance(by_vin, dict):
        return False
    if not by_vin:
        return True

    vin, record = next(iter(by_vin.items()))
    return isinstance(vin, str) and isinstance(record, dict) and tuple(record) == RECORD_FIELDS


def vessel_key(value):
    text = "" if value is None else str(value).strip()
    text = unicodedata.normalize("NFD", text.upper())
    text = "".join(character for character in text if unicodedata.category(character) != "Mn")
    return " ".join(text.split())


def apply_shipping_eta_overrides(by_vin, overrides, report_progress):
    """Return an in-memory ETA-adjusted copy without changing the cached source."""
    if not overrides:
        return by_vin

    normalized_overrides = {
        vessel_key(vessel): eta
        for vessel, eta in overrides.items()
        if vessel_key(vessel) and eta is not None
    }
    if not normalized_overrides:
        return by_vin

    matched_vessels = set()
    adjusted = None
    changed_vins = 0
    for vin, record in by_vin.items():
        vessel = vessel_key(record.get("vessel_name") or record.get("vessel"))
        if vessel not in normalized_overrides:
            continue
        if str(record.get("status") or "").strip().upper() != "SHIPPING":
            continue

        matched_vessels.add(vessel)
        if record.get("eta") == normalized_overrides[vessel]:
            continue

        if adjusted is None:
            adjusted = by_vin.copy()
        adjusted_record = record.copy()
        adjusted_record["eta"] = normalized_overrides[vessel]
        adjusted[vin] = adjusted_record
        changed_vins += 1

    if changed_vins:
        report_progress(
            f"Shipping ETA overrides applied: {changed_vins:,} VINs across {len(matched_vessels):,} vessel(s)"
        )

    unmatched_vessels = sorted(set(normalized_overrides) - matched_vessels)
    if unmatched_vessels:
        report_progress(
            "WARNING: shipping ETA overrides without Shipping VINs: " + "; ".join(unmatched_vessels[:20])
        )

    return adjusted if adjusted is not None else by_vin
