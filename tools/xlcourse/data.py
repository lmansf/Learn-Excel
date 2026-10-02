"""Typed loaders for the course CSV files in ``data/``.

    from xlcourse import data
    enc = data.load("encounters")          # list[dict] with real date/datetime/float/int values
    enc_2025 = [e for e in enc if e["AdmitDateTime"].year == 2025]

Values are converted by column name:
  * date columns      -> datetime.date
  * datetime columns  -> datetime.datetime
  * numeric columns   -> int or float
  * everything else   -> str  (IDs, MRN, ZIP and Phone stay text on purpose)
  * empty cells       -> None
"""
from __future__ import annotations

import csv
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
AS_OF = date(2025, 12, 31)

DATE_COLS = {
    "HireDate", "DOB", "RegistrationDate", "DeceasedDate", "ServiceDate", "SubmitDate", "PaidDate", "DischargeDate",
    "SurveyReceivedDate", "CensusDate", "TermDate", "ShiftDate", "LastCountDate", "ExpirationDate", "MonthStart",
}
DATETIME_COLS = {
    "AdmitDateTime", "DischargeDateTime", "ArrivalDateTime", "TriageDateTime", "ProviderSeenDateTime",
    "DispositionDateTime", "DepartureDateTime", "OrderDateTime", "CollectedDateTime", "ResultedDateTime",
    "ScheduledStart", "ScheduledEnd", "ClockIn", "ClockOut",
}
INT_COLS = {
    "LicensedBeds", "OpenedYear", "StaffedBeds", "CostCenter", "AvgDaysToPay", "TimelyFilingDays", "ESILevel",
    "HeartRate", "RespRate", "SystolicBP", "DiastolicBP", "SpO2", "PainScore", "DosesDispensed", "NurseCommunication",
    "DoctorCommunication", "StaffResponsiveness", "Cleanliness", "Quietness", "MedicationExplained", "OverallRating",
    "Admissions", "Discharges", "MidnightCensus", "ScheduledHours", "QtyOnHand", "ParLevel", "ReorderPoint", "ReorderQty",
    "LeadTimeDays", "FiscalYear", "FiscalMonth", "BudgetAmount", "ActualAmount",
}
FLOAT_COLS = {
    "AvgAllowedPctOfCharges", "FTE", "HeightIn", "WeightLb", "TotalCharges", "BilledAmount", "AllowedAmount",
    "PatientResponsibility", "PaidAmount", "ResultValue", "RefLow", "RefHigh", "UnitCost", "TempF", "HourlyRate",
    "WorkedHours", "HoursOverSchedule", "ExpectedLOS",
}

TABLES = [
    "facilities", "departments", "providers", "payers", "diagnoses", "patients", "encounters", "ed_visits", "claims",
    "lab_results", "medications", "patient_satisfaction", "daily_census", "employees", "shifts", "supply_inventory", "budget",
]


def _convert(col: str, v: str):
    if v == "":
        return None
    if col in DATE_COLS:
        return date.fromisoformat(v)
    if col in DATETIME_COLS:
        return datetime.strptime(v, "%Y-%m-%d %H:%M")
    if col in INT_COLS:
        return int(v)
    if col in FLOAT_COLS:
        return float(v)
    return v


@lru_cache(maxsize=None)
def _load(name: str) -> tuple:
    path = DATA / f"{name}.csv"
    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return tuple({k: _convert(k, v) for k, v in row.items()} for row in reader)


def load(name: str) -> list[dict]:
    """Return a fresh list of row dicts (safe to mutate) for a dataset name like 'encounters'."""
    return [dict(r) for r in _load(name)]


def load_raw(relpath: str) -> list[dict]:
    """Load any CSV under data/ as plain strings (e.g. 'messy/patient_registrations_raw')."""
    path = DATA / (relpath if relpath.endswith(".csv") else relpath + ".csv")
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def columns(name: str) -> list[str]:
    with (DATA / f"{name}.csv").open(newline="", encoding="utf-8") as f:
        return next(csv.reader(f))


def index(rows: list[dict], key: str) -> dict:
    """Map key -> row (for lookups). Keys must be unique."""
    return {r[key]: r for r in rows}


def excel_serial(d) -> float:
    """Excel serial number for a date/datetime (1900 date system)."""
    base = datetime(1899, 12, 30)
    if isinstance(d, datetime):
        return (d - base).total_seconds() / 86400
    return float((d - base.date()).days)
