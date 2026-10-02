#!/usr/bin/env python3
"""Generate the synthetic Bluestone Health System datasets used throughout the course.

    python tools/generate_data.py

Writes CSV files to ``data/`` (and answer-key truth files to ``tools/_truth/``).
The output is fully deterministic: the same SEED always produces byte-identical
files, so every lesson's answer key stays correct when the data is regenerated.

All people, organizations, places, and numbers are fictional. ICD-10-CM codes are
real public codes used only for realism.
"""
from __future__ import annotations

import csv
import math
import random
from bisect import bisect_right
from datetime import date, datetime, timedelta
from itertools import accumulate
from pathlib import Path

import catalogs as C

SEED = 20240101
START = date(2024, 1, 1)
END = date(2025, 12, 31)
AS_OF = date(2025, 12, 31)
ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data"
TRUTH = Path(__file__).resolve().parent / "_truth"


def R(name: str) -> random.Random:
    """Independent, deterministic random stream per dataset (string seeds are hashed with SHA-512)."""
    return random.Random(f"{SEED}:{name}")


def wchoice(r: random.Random, items, weights):
    return r.choices(items, weights=weights, k=1)[0]


def clamp(x, lo, hi):
    return max(lo, min(hi, x))


def lognorm(r: random.Random, median: float, sigma: float) -> float:
    return median * math.exp(r.gauss(0, sigma))


def fmt_dt(x: datetime | None) -> str:
    return "" if x is None else x.strftime("%Y-%m-%d %H:%M")


def fmt_d(x: date | None) -> str:
    return "" if x is None else x.strftime("%Y-%m-%d")


def money(x: float) -> str:
    return f"{x:.2f}"


def daterange(a: date, b: date):
    d = a
    while d <= b:
        yield d
        d += timedelta(days=1)


def season_mult(month: int, winter: float, summer: float) -> float:
    if month in (12, 1, 2):
        return winter
    if month in (11, 3):
        return 1 + (winter - 1) / 2
    if month in (6, 7, 8):
        return summer
    if month in (5, 9):
        return 1 + (summer - 1) / 2
    return 1.0


def age_on(dob: date, on: date) -> int:
    return on.year - dob.year - ((on.month, on.day) < (dob.month, dob.day))


# ---------------------------------------------------------------------------
# Facilities, departments, payers
# ---------------------------------------------------------------------------
DEPT = {d[0]: d for d in C.DEPARTMENTS}
PAYER = {p[0]: p for p in C.PAYERS}


def write_csv(path: Path, header: list[str], rows: list[list]):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        for row in rows:
            w.writerow(["" if v is None else v for v in row])


# ---------------------------------------------------------------------------
# Providers
# ---------------------------------------------------------------------------
PROVIDER_PLAN = [
    # specialty, credentials (weights), dept, count
    ("Emergency Medicine", {"MD": 3, "DO": 1}, "D100", 10),
    ("Emergency Medicine", {"PA": 1, "NP": 1}, "D100", 4),
    ("Emergency Medicine", {"MD": 3, "DO": 1}, "D200", 5),
    ("Emergency Medicine", {"PA": 1, "NP": 1}, "D200", 1),
    ("Emergency Medicine", {"MD": 3, "DO": 1}, "D300", 6),
    ("Emergency Medicine", {"PA": 1, "NP": 1}, "D300", 2),
    ("Hospital Medicine", {"MD": 3, "DO": 1}, "D110", 7),
    ("Hospital Medicine", {"MD": 3, "DO": 1}, "D111", 6),
    ("Hospital Medicine", {"NP": 1}, "D195", 2),
    ("Hospital Medicine", {"MD": 3, "DO": 1}, "D210", 4),
    ("Hospital Medicine", {"MD": 3, "DO": 1}, "D310", 6),
    ("Cardiology", {"MD": 4, "DO": 1}, "D120", 4),
    ("Cardiology", {"MD": 4, "DO": 1}, "D320", 2),
    ("Cardiology", {"MD": 4, "DO": 1}, "D410", 3),
    ("Pulmonary & Critical Care", {"MD": 4, "DO": 1}, "D130", 5),
    ("Pulmonary & Critical Care", {"MD": 4, "DO": 1}, "D230", 2),
    ("Pulmonary & Critical Care", {"MD": 4, "DO": 1}, "D330", 3),
    ("Orthopedic Surgery", {"MD": 3, "DO": 1}, "D140", 4),
    ("Orthopedic Surgery", {"MD": 3, "DO": 1}, "D340", 2),
    ("Orthopedic Surgery", {"MD": 3, "DO": 1}, "D420", 2),
    ("Orthopedic Surgery", {"PA": 1}, "D420", 2),
    ("General Surgery", {"MD": 4, "DO": 1}, "D196", 5),
    ("General Surgery", {"MD": 4, "DO": 1}, "D210", 1),
    ("General Surgery", {"MD": 4, "DO": 1}, "D310", 2),
    ("Obstetrics & Gynecology", {"MD": 3, "DO": 1}, "D170", 5),
    ("Obstetrics & Gynecology", {"CNM": 1}, "D170", 2),
    ("Obstetrics & Gynecology", {"MD": 3, "DO": 1}, "D270", 2),
    ("Pediatrics", {"MD": 4, "DO": 1}, "D180", 4),
    ("Pediatrics", {"MD": 4, "DO": 1}, "D450", 4),
    ("Pediatrics", {"NP": 1}, "D450", 1),
    ("Hematology & Oncology", {"MD": 1}, "D150", 3),
    ("Hematology & Oncology", {"MD": 1}, "D430", 3),
    ("Hematology & Oncology", {"NP": 1}, "D430", 1),
    ("Neurology", {"MD": 4, "DO": 1}, "D160", 3),
    ("Psychiatry", {"MD": 3, "DO": 1}, "D190", 4),
    ("Psychiatry", {"NP": 1}, "D190", 1),
    ("Family Medicine", {"MD": 2, "DO": 1}, "D400", 8),
    ("Family Medicine", {"NP": 1}, "D400", 3),
    ("Family Medicine", {"PA": 1}, "D400", 2),
    ("Internal Medicine", {"MD": 3, "DO": 1}, "D400", 4),
    ("Endocrinology", {"MD": 1}, "D440", 2),
    ("Endocrinology", {"NP": 1}, "D440", 1),
    ("Radiology", {"MD": 1}, "D510", 3),
    ("Pathology", {"MD": 1}, "D500", 1),
]


def gen_providers():
    r = R("providers")
    rows = []
    used = set()
    n = 0
    for spec, creds, dept, count in PROVIDER_PLAN:
        for _ in range(count):
            n += 1
            while True:
                sex = "F" if r.random() < 0.5 else "M"
                first = r.choice(C.FEMALE_FIRST if sex == "F" else C.MALE_FIRST)
                last = r.choice(C.LAST_NAMES)
                if (first, last) not in used:
                    used.add((first, last))
                    break
            cred = wchoice(r, list(creds), list(creds.values()))
            hire = date(1998, 1, 1) + timedelta(days=r.randint(0, (date(2023, 12, 31) - date(1998, 1, 1)).days))
            fte = wchoice(r, [1.0, 0.9, 0.8, 0.6, 0.5], [70, 8, 10, 7, 5])
            rows.append({
                "ProviderID": f"PRV{1000 + n}", "FirstName": first, "LastName": last, "Credential": cred,
                "Specialty": spec, "PrimaryDeptID": dept, "FacilityID": DEPT[dept][2],
                "HireDate": hire, "FTE": fte,
            })
    return rows


# ---------------------------------------------------------------------------
# Patients
# ---------------------------------------------------------------------------
N_PATIENTS = 4000


def chronic_profile(r, age, sex):
    cond = []
    def p(x):
        return r.random() < x
    if age >= 40 and p(0.60 if age >= 65 else 0.32):
        cond.append("HTN")
    if age >= 30 and p(0.25 if age >= 65 else 0.12):
        cond.append("DM")
    if age >= 55 and p(0.15 if age >= 75 else 0.06):
        cond.append("HF")
    if age >= 45 and p(0.12 if age >= 65 else 0.05):
        cond.append("COPD")
    if age >= 60 and p(0.13 if age >= 75 else 0.06):
        cond.append("AFIB")
    if age >= 50 and p(0.12 if age >= 70 else 0.05):
        cond.append("CKD")
    if 3 <= age <= 70 and p(0.08):
        cond.append("ASTHMA")
    if age >= 35 and p(0.06 if age >= 60 else 0.025):
        cond.append("CANCER")
    if 13 <= age <= 85 and p(0.10):
        cond.append("PSYCH")
    if 21 <= age <= 80 and p(0.045 if sex == "M" else 0.02):
        cond.append("ETOH")
    if age >= 50 and p(0.12):
        cond.append("OA")
    if p(0.012):
        cond.append("SEIZURE")
    return cond


def gen_patients(providers):
    r = R("patients")
    pcp_adult = [p["ProviderID"] for p in providers if p["PrimaryDeptID"] == "D400"]
    pcp_peds = [p["ProviderID"] for p in providers if p["PrimaryDeptID"] == "D450"]
    towns = C.TOWNS
    town_w = [t[3] for t in towns]
    mrns = set()
    patients = []
    for i in range(N_PATIENTS):
        sex = "F" if r.random() < 0.53 else "M"
        band = wchoice(r, [(0, 17), (18, 44), (45, 64), (65, 79), (80, 99)], [12, 25, 28, 23, 12])
        age = r.randint(*band)
        by = 2025 - age
        dob = date(by, 1, 1) + timedelta(days=r.randint(0, 364))
        first = r.choice(C.FEMALE_FIRST if sex == "F" else C.MALE_FIRST)
        last = r.choice(C.LAST_NAMES)
        town = wchoice(r, towns, town_w)
        if age < 18:
            payer = wchoice(r, ["PY03", "PY04", "PY05", "PY06", "PY07"], [45, 20, 16, 14, 5])
        elif age < 65:
            payer = wchoice(r, ["PY04", "PY05", "PY06", "PY03", "PY07", "PY01"], [24, 20, 16, 23, 9, 4])
        else:
            payer = wchoice(r, ["PY01", "PY02", "PY04", "PY03"], [55, 37, 4, 4])
        pcp = None
        if r.random() > 0.08:
            pcp = r.choice(pcp_peds if age < 18 else pcp_adult)
        while True:
            mrn = f"{r.randint(1, 9_999_999):08d}"
            if mrn not in mrns:
                mrns.add(mrn)
                break
        # anthropometrics
        if age >= 18:
            h = r.gauss(69.3, 2.9) if sex == "M" else r.gauss(63.9, 2.7)
            bmi = clamp(r.gauss(29.2, 6.2), 16.5, 58)
        else:
            h = clamp(20 + age * 2.6 + r.gauss(0, 1.8), 19, 74) if age < 13 else r.gauss(64 if sex == "F" else 66, 3)
            bmi = clamp(r.gauss(16.5 + age * 0.35, 2.8), 12.5, 40)
        w = bmi * h * h / 703
        lang = wchoice(r, [l for l, _ in C.LANGUAGES], [w_ for _, w_ in C.LANGUAGES])
        cond = chronic_profile(r, age, sex)
        email = None
        if r.random() < 0.68 and age >= 16:
            email = f"{first}.{last}{r.randint(1, 99)}@example.com".lower().replace("'", "")
        phone = f"(555) {r.randint(200, 989)}-{r.randint(0, 9999):04d}"
        patients.append({
            "PatientID": f"PT{10001 + i}", "MRN": mrn, "FirstName": first, "LastName": last, "Sex": sex,
            "DOB": dob, "City": town[0], "State": "OH", "ZIP": town[1], "HomeFacility": town[2],
            "Phone": phone, "Email": email, "PreferredLanguage": lang, "PrimaryPayerID": payer,
            "PCPProviderID": pcp, "HeightIn": round(h, 1), "WeightLb": round(w, 1),
            "Chronic": cond, "RegistrationDate": None, "DeceasedDate": None,
        })
    return patients


# ---------------------------------------------------------------------------
# Encounters, ED visits
# ---------------------------------------------------------------------------
TYPES = ("Inpatient", "Observation", "Emergency", "Outpatient")
BASE_RATE = {"Inpatient": 6.6, "Observation": 1.5, "Emergency": 9.3, "Outpatient": 13.8}
DOW_FACTOR = {
    "Inpatient": [1.08, 1.06, 1.05, 1.05, 1.04, 0.86, 0.84],
    "Observation": [1.05, 1.0, 1.0, 1.0, 1.0, 0.98, 0.97],
    "Emergency": [1.13, 1.03, 1.0, 0.99, 1.0, 0.94, 0.92],
    "Outpatient": [1.05, 1.0, 1.0, 1.0, 0.92, 0.0, 0.0],
}
ED_HOUR_W = [2.2, 1.8, 1.5, 1.3, 1.2, 1.3, 1.8, 2.8, 4.0, 5.0, 5.6, 5.8, 5.8, 5.7, 5.6, 5.6, 5.6, 5.6, 5.5, 5.2, 4.6, 3.9, 3.2, 2.6]
CROWD = [0.75, 0.7, 0.68, 0.66, 0.66, 0.7, 0.78, 0.88, 0.98, 1.08, 1.18, 1.25, 1.3, 1.32, 1.34, 1.36, 1.36, 1.34, 1.3, 1.24, 1.15, 1.02, 0.92, 0.82]


def dx_eligible(dx, age, sex):
    d = C.DX[dx]
    amin, amax = d[4]
    return amin <= age <= amax and (d[5] is None or d[5] == sex)


def build_patient_pools(patients):
    """Static per-diagnosis sampling pools (eligible patients + cumulative weights)."""
    chronic_dx = {}
    for key, codes in C.CHRONIC_LINKS.items():
        for c in codes:
            chronic_dx[c] = key
    pools = {}
    for dx in C.DX:
        ids, cum, total = [], [], 0.0
        for idx, p in enumerate(patients):
            age = age_on(p["DOB"], date(2025, 1, 1))
            if not dx_eligible(dx, age, p["Sex"]):
                continue
            if dx in chronic_dx:
                w = 4.0 if chronic_dx[dx] in p["Chronic"] else 0.4
            else:
                w = 1.0 + 0.35 * len(p["Chronic"])
            total += w
            ids.append(idx)
            cum.append(total)
        pools[dx] = (ids, cum, total)
    return pools


def pick_from_pool(r, pool):
    ids, cum, total = pool
    x = r.random() * total
    return ids[min(bisect_right(cum, x), len(ids) - 1)]


def choose_facility(r, home, route, enc_type):
    if enc_type == "Outpatient":
        return "F04"
    if route in ("onc", "neuro", "peds", "psych"):
        return "F01"
    x = r.random()
    fac = home if x < 0.76 else ("F01" if x < 0.94 else r.choice(["F01", "F02", "F03"]))
    if route in ("ortho", "ortho_elective") and fac == "F02":
        fac = "F01" if r.random() < 0.7 else "F03"
    if route == "ob" and fac == "F03":
        fac = "F01"
    return fac


def providers_by(providers):
    by_dept, by_fac_spec = {}, {}
    for p in providers:
        by_dept.setdefault(p["PrimaryDeptID"], []).append(p)
        by_fac_spec.setdefault((p["FacilityID"], p["Specialty"]), []).append(p)
    return by_dept, by_fac_spec


def ed_esi(r, dx, admitted):
    acuity = C.DX[dx][15]
    base = acuity + r.gauss(0, 0.6) - (0.3 if admitted else 0)
    for level, cut in ((1, 1.25), (2, 2.45), (3, 3.55), (4, 4.5)):
        if base < cut:
            return level
    return 5


def vitals_for(r, dx, age, esi):
    temp, hr, rr, sbp, spo2, pain = r.gauss(98.3, 0.45), r.gauss(84, 11), r.gauss(16.5, 2), r.gauss(132, 17), r.gauss(97.6, 1.1), r.choice([0, 0, 0, 1, 2, 3])
    if age < 1:
        hr, rr, sbp = r.gauss(135, 14), r.gauss(38, 6), r.gauss(85, 8)
    elif age < 6:
        hr, rr, sbp = r.gauss(112, 12), r.gauss(26, 4), r.gauss(96, 8)
    elif age < 13:
        hr, rr, sbp = r.gauss(96, 11), r.gauss(21, 3), r.gauss(105, 9)
    cat = C.DX[dx][1]
    if dx == "A41.9":
        temp = r.gauss(101.9, 1.3) if r.random() < 0.82 else r.gauss(96.2, 0.6)
        hr, rr, sbp, spo2 = r.gauss(116, 13), r.gauss(24, 4), r.gauss(98, 15), r.gauss(94, 2.2)
    elif dx in ("J18.9", "U07.1"):
        temp, hr, rr, spo2 = r.gauss(101.0, 1.0), r.gauss(104, 11), r.gauss(23, 3), r.gauss(92, 2.6)
    elif dx in ("J44.1", "J96.01", "J21.9"):
        hr, rr, spo2 = hr + r.gauss(18, 8), rr + r.gauss(9, 3), r.gauss(88.5, 3)
    elif dx == "J45.909":
        hr, rr, spo2 = hr + r.gauss(14, 6), rr + r.gauss(6, 2), r.gauss(94, 2)
    elif dx == "I50.9":
        rr, spo2, sbp = r.gauss(23, 3), r.gauss(91.5, 2.5), r.gauss(152, 22)
    elif dx in ("I21.4", "R07.9"):
        hr, sbp, pain = r.gauss(90, 14), r.gauss(150, 22), r.randint(3, 9)
    elif dx == "I48.91":
        hr = r.gauss(128, 16)
    elif dx == "I63.9":
        sbp = r.gauss(176, 22)
    elif dx == "K92.2":
        hr, sbp = r.gauss(112, 12), r.gauss(102, 14)
    elif dx in ("I10",):
        sbp = r.gauss(182, 16)
    elif dx in ("E86.0", "A08.4"):
        hr = hr + r.gauss(16, 6)
    elif dx in ("N39.0", "L03.115", "H66.90"):
        temp = r.gauss(100.4, 1.1)
    if cat in ("Injury", "Musculoskeletal") or dx in ("R10.9", "K35.80", "K80.20", "K56.609", "G43.909", "M54.50"):
        pain = r.randint(4, 10)
    if esi == 1:
        sbp = min(sbp, r.gauss(88, 12)) if r.random() < 0.5 else sbp
        spo2 = min(spo2, r.gauss(87, 4))
    sbp = clamp(sbp, 60, 240)
    dbp = clamp(sbp * r.uniform(0.52, 0.68), 30, 140)
    return (round(clamp(temp, 94.0, 105.8), 1), int(clamp(hr, 38, 210)), int(clamp(rr, 8, 60)),
            int(sbp), int(dbp), int(clamp(round(spo2), 70, 100)), int(pain))


def gen_encounters(patients, providers):
    r = R("encounters")
    rp = R("encounter-patients")
    pools = build_patient_pools(patients)
    by_dept, by_fac_spec = providers_by(providers)
    busy_until = [datetime(2000, 1, 1)] * len(patients)
    deceased = [None] * len(patients)
    scheduled = {}  # date -> list of (patient_idx, dx)
    encounters, ed_rows = [], []
    enc_n = ed_n = 0
    dx_codes = list(C.DX)

    def pick_provider(fac, dept, route, dx, enc_type, pat):
        def pool(spec):
            return by_fac_spec.get((fac, spec), [])
        if enc_type == "Emergency":
            cands = by_dept[C.ED_DEPT[fac]]
        elif enc_type == "Outpatient":
            cands = by_dept[dept]
            if pat["PCPProviderID"] and dept in ("D400", "D450") and r.random() < 0.7:
                return pat["PCPProviderID"]
        elif route in ("cardiac",) and fac in ("F01", "F03"):
            cands = pool("Cardiology") if r.random() < 0.6 else pool("Hospital Medicine")
        elif dept in C.ICU_DEPT.values():
            cands = pool("Pulmonary & Critical Care")
        elif route == "surgery":
            cands = pool("General Surgery")
        elif route in ("ortho", "ortho_elective"):
            cands = pool("Orthopedic Surgery")
        elif route == "ob":
            cands = pool("Obstetrics & Gynecology")
            if dx != "O80":
                cands = [c for c in cands if c["Credential"] != "CNM"]
        elif route == "peds":
            cands = by_dept["D180"]
        elif route == "onc":
            cands = by_dept["D150"]
        elif route == "neuro" and fac == "F01":
            cands = by_dept["D160"]
        elif route == "psych":
            cands = by_dept["D190"]
        elif enc_type == "Observation" and fac == "F01":
            cands = by_dept["D195"] + by_dept["D110"] if r.random() < 0.35 else by_dept["D110"] + by_dept["D111"]
        else:
            cands = pool("Hospital Medicine")
        cands = [c for c in cands if c["FacilityID"] == fac or enc_type == "Outpatient"] or cands
        weights = [3 if c["Credential"] in ("MD", "DO") else 1.4 for c in cands]
        return wchoice(r, cands, weights)["ProviderID"]

    def ip_dept(fac, route, dx, enc_type):
        if enc_type == "Observation":
            return C.OBS_DEPT[fac]
        if route == "cardiac":
            return C.CARDIAC_DEPT.get(fac) or C.MEDSURG_DEPT[fac][0]
        if route == "icu":
            p_icu = {"A41.9": 0.42, "J96.01": 0.8}.get(dx, 0.3)
            return C.ICU_DEPT[fac] if r.random() < p_icu else r.choice(C.MEDSURG_DEPT[fac])
        if route == "medsurg_icu":
            return C.ICU_DEPT[fac] if r.random() < 0.3 else r.choice(C.MEDSURG_DEPT[fac])
        if route in ("ortho", "ortho_elective"):
            return C.ORTHO_DEPT[fac]
        if route == "ob":
            return C.OB_DEPT[fac]
        if route == "peds":
            return "D180"
        if route == "onc":
            return "D150"
        if route == "neuro":
            return "D160" if fac == "F01" else r.choice(C.MEDSURG_DEPT[fac])
        if route == "psych":
            return "D190"
        return r.choice(C.MEDSURG_DEPT[fac])

    def clinic_dept(dx, age):
        key = C.DX[dx][16] or "primary"
        if age < 18 and key in ("primary", "peds"):
            return "D450"
        return {"primary": "D400", "cardio": "D410", "ortho": "D420", "onc": "D430", "endo": "D440", "peds": "D450"}[key]

    def disposition(dx, age, dept, enc_type, fac, n_chronic):
        route = C.DX[dx][8]
        mort = 0.65 * C.DX[dx][13] * (2.0 if age >= 80 else 1.3 if age >= 65 else 0.3 if age < 40 else 1.0)
        if dept in C.ICU_DEPT.values():
            mort *= 1.5
        if enc_type == "Observation":
            opts = {"Home/Self-Care": 93, "Home Health": 3, "Skilled Nursing Facility": 1.5, "Left AMA": 2, "Transfer to Another Hospital": 0.5}
        elif route == "ob" or route == "peds":
            opts = {"Home/Self-Care": 97.5, "Home Health": 1.5, "Left AMA": 0.5, "Transfer to Another Hospital": 0.5}
        elif dx == "S72.001A":
            opts = {"Home/Self-Care": 10, "Home Health": 22, "Skilled Nursing Facility": 44, "Inpatient Rehab": 18, "Hospice": 1.5}
        elif route == "ortho_elective":
            opts = {"Home/Self-Care": 45, "Home Health": 45, "Skilled Nursing Facility": 7, "Inpatient Rehab": 3}
        elif dx == "I63.9":
            opts = {"Home/Self-Care": 30, "Home Health": 20, "Skilled Nursing Facility": 20, "Inpatient Rehab": 25, "Hospice": 2}
        elif route == "psych":
            opts = {"Home/Self-Care": 86, "Left AMA": 4, "Transfer to Another Hospital": 6, "Home Health": 1}
        else:
            opts = {"Home/Self-Care": 74, "Home Health": 11, "Skilled Nursing Facility": 7, "Inpatient Rehab": 1.5,
                    "Hospice": 1.0, "Left AMA": 1.8, "Transfer to Another Hospital": 1.5}
            if dx == "F10.239":
                opts["Left AMA"] = 9
            if age >= 75:
                opts["Skilled Nursing Facility"] *= 2.5
                opts["Home Health"] *= 1.6
                opts["Hospice"] = opts.get("Hospice", 0) * 2
            if fac == "F02":
                opts["Transfer to Another Hospital"] *= 3
        if r.random() < mort:
            return "Expired"
        return wchoice(r, list(opts), list(opts.values()))

    def charges(enc_type, dx, dept, los_days, esi, route):
        noise = math.exp(r.gauss(0, 0.22))
        if enc_type == "Inpatient":
            unit_route = "icu" if dept in C.ICU_DEPT.values() else route if route in C.ROOM_RATE else "medsurg"
            if route == "ortho_elective":
                unit_route = "ortho"
            rate = C.ROOM_RATE.get(unit_route, 3200)
            ancillary = r.uniform(1800, 9000) + (4000 if unit_route == "icu" else 0)
            proc = C.DX[dx][11]
            if dx == "I21.4" and r.random() < 0.55:
                proc += 24000  # PCI
            return (rate * max(los_days, 1) + ancillary + proc) * noise
        if enc_type == "Observation":
            return (C.ROOM_RATE["obs"] * max(los_days, 0.5) + r.uniform(2500, 7500)) * noise
        if enc_type == "Emergency":
            base = {1: 9200, 2: 5600, 3: 3300, 4: 1450, 5: 720}[esi]
            return base * noise
        # outpatient
        if dx == "Z51.11":
            return r.uniform(6500, 14500) * noise
        if dx in ("C34.90", "C50.911"):
            return r.uniform(420, 980) * noise
        if dx == "Z23":
            return r.uniform(95, 145)
        if dx in ("Z00.00", "Z00.129"):
            return r.uniform(260, 390) * noise
        return r.uniform(170, 460) * noise

    def ed_timeline(arrival, esi, admitted, dx):
        crowd = CROWD[arrival.hour] * (1.1 if arrival.weekday() == 0 else 1.0) * (1.12 if arrival.month in (12, 1, 2) else 1.0)
        triage = arrival + timedelta(minutes=r.uniform(0, 3) if esi == 1 else clamp(lognorm(r, 8, 0.5), 2, 40))
        med = {1: 2, 2: 14, 3: 38, 4: 52, 5: 48}[esi]
        d2p = r.uniform(0, 6) if esi == 1 else clamp(lognorm(r, med * crowd, 0.6), 3, 420)
        lwbs = False
        if not admitted and esi >= 3 and d2p > 120 and r.random() < min(0.6, (d2p - 120) / 220 + 0.18):
            lwbs = True
        if lwbs:
            depart = arrival + timedelta(minutes=d2p * r.uniform(0.5, 0.95))
            return triage, None, depart, depart, "LWBS"
        provider = arrival + timedelta(minutes=max(d2p, (triage - arrival).total_seconds() / 60 + 1))
        dec = {1: (90, 240), 2: (90, 240), 3: (80, 210), 4: (25, 95), 5: (20, 70)}[esi]
        dispo_t = provider + timedelta(minutes=r.uniform(*dec))
        if admitted:
            depart = dispo_t + timedelta(minutes=clamp(lognorm(r, 150, 0.7), 25, 1100))
            return triage, provider, dispo_t, depart, "Admitted"
        x = r.random()
        if x < 0.016:
            depart = provider + timedelta(minutes=r.uniform(20, 120))
            return triage, provider, depart, depart, "Left AMA"
        if x < 0.016 + (0.06 if esi <= 2 else 0.006):
            depart = dispo_t + timedelta(minutes=r.uniform(60, 300))
            return triage, provider, dispo_t, depart, "Transferred"
        if esi == 1 and x > 0.97:
            return triage, provider, dispo_t, dispo_t, "Expired"
        depart = dispo_t + timedelta(minutes=r.uniform(8, 45))
        return triage, provider, dispo_t, depart, "Discharged"

    def arrival_mode(esi, age, dx):
        if C.DX[dx][1] == "Mental & Behavioral" and r.random() < 0.15:
            return "Police"
        p_amb = {1: 0.85, 2: 0.45, 3: 0.2, 4: 0.06, 5: 0.03}[esi] * (1.6 if age >= 75 else 1)
        if esi == 1 and r.random() < 0.06:
            return "Air Transport"
        return "Ambulance" if r.random() < p_amb else "Walk-In"

    def complaint(dx):
        return C.DX[dx][14] if r.random() > 0.06 else r.choice(C.CHIEF_COMPLAINT_EXTRA)

    def add_ed_row(enc_id, pat, fac, arrival, triage, provider, dispo_t, depart, ed_dispo, esi, dx, age):
        nonlocal ed_n
        ed_n += 1
        v = vitals_for(r, dx, age, esi)
        ed_rows.append({
            "EDVisitID": f"ED{200000 + ed_n}", "EncounterID": enc_id, "PatientID": pat["PatientID"],
            "FacilityID": fac, "ArrivalDateTime": arrival, "TriageDateTime": triage,
            "ProviderSeenDateTime": provider, "DispositionDateTime": dispo_t, "DepartureDateTime": depart,
            "ArrivalMode": arrival_mode(esi, age, dx), "ESILevel": esi, "ChiefComplaint": complaint(dx),
            "EDDisposition": ed_dispo, "TempF": v[0], "HeartRate": v[1], "RespRate": v[2],
            "SystolicBP": v[3], "DiastolicBP": v[4], "SpO2": v[5], "PainScore": v[6],
        })

    def pick_patient(dx, start):
        for _ in range(80):
            idx = pick_from_pool(rp, pools[dx])
            if deceased[idx] is None and busy_until[idx] < start - timedelta(hours=2):
                return idx
        return None

    for day in daterange(START, END):
        growth = 1.04 if day.year == 2025 else 1.0
        slots = []
        for t, ti in zip(TYPES, range(4)):
            w = [C.DX[d][3][ti] * season_mult(day.month, C.DX[d][6], C.DX[d][7]) for d in dx_codes]
            base_w = sum(C.DX[d][3][ti] for d in dx_codes)
            lam = BASE_RATE[t] * growth * DOW_FACTOR[t][day.weekday()] * (sum(w) / base_w)
            # Poisson draw via Knuth (lam is small)
            L, k, p = math.exp(-lam), 0, 1.0
            while True:
                p *= r.random()
                if p <= L:
                    break
                k += 1
            for _ in range(k):
                dx = wchoice(r, dx_codes, w)
                if C.DX[dx][8] == "ortho_elective" and day.weekday() >= 5:
                    continue
                slots.append((t, dx, None))
        for (pidx, dx) in scheduled.pop(day, []):
            slots.append(("Inpatient", dx, pidx))

        # assign start times
        timed = []
        for t, dx, forced in slots:
            route = C.DX[dx][8]
            if t == "Outpatient":
                start = datetime.combine(day, datetime.min.time()) + timedelta(minutes=8 * 60 + 15 * r.randint(0, 34))
                src = None
            elif t == "Inpatient" and route == "ortho_elective":
                start = datetime.combine(day, datetime.min.time()) + timedelta(minutes=r.randint(330, 480))
                src = "Elective"
            elif t == "Inpatient" and route == "ob":
                start = datetime.combine(day, datetime.min.time()) + timedelta(minutes=r.randint(0, 1439))
                src = "Direct Admit" if r.random() < 0.75 else "Emergency Department"
            else:
                hour = wchoice(r, list(range(24)), ED_HOUR_W)
                start = datetime.combine(day, datetime.min.time()) + timedelta(minutes=hour * 60 + r.randint(0, 59))
                if t in ("Emergency",):
                    src = None
                elif t == "Observation":
                    src = "Emergency Department" if r.random() < 0.86 else "Direct Admit"
                else:
                    x = r.random()
                    if route == "surgery" and dx == "K80.20" and x < 0.4:
                        src = "Elective"
                    elif forced is not None:
                        src = "Emergency Department" if x < 0.86 else "Direct Admit"
                    else:
                        src = "Emergency Department" if x < 0.74 else ("Direct Admit" if x < 0.95 else "Transfer")
            timed.append((start, t, dx, forced, src))
        timed.sort(key=lambda z: (z[0], z[1], z[2]))

        for start, t, dx, forced, src in timed:
            if forced is not None:
                pidx = forced
                if deceased[pidx] is not None or busy_until[pidx] >= start:
                    continue
            else:
                pidx = pick_patient(dx, start)
                if pidx is None:
                    continue
            pat = patients[pidx]
            age = age_on(pat["DOB"], day)
            route = C.DX[dx][8]
            fac = choose_facility(r, pat["HomeFacility"], route, t)
            enc_n += 1
            enc_id = f"ENC{100000 + enc_n}"
            esi = None
            ed_info = None
            if t == "Emergency":
                dept = C.ED_DEPT[fac]
                esi = ed_esi(r, dx, False)
                triage, provider, dispo_t, depart, ed_dispo = ed_timeline(start, esi, False, dx)
                admit_dt, disch_dt = start, depart
                dispo = {"Discharged": "Home/Self-Care", "LWBS": "Left Without Being Seen", "Left AMA": "Left AMA",
                         "Transferred": "Transfer to Another Hospital", "Expired": "Expired"}[ed_dispo]
                ed_info = (start, triage, provider, dispo_t, depart, ed_dispo)
                los_days = (disch_dt - admit_dt).total_seconds() / 86400
            elif t == "Outpatient":
                dept = clinic_dept(dx, age)
                dur = r.randint(150, 330) if dx == "Z51.11" else r.randint(15, 55)
                admit_dt, disch_dt = start, start + timedelta(minutes=dur)
                dispo = "Home/Self-Care"
                los_days = 0
            else:
                dept = ip_dept(fac, route, dx, t)
                if src == "Emergency Department":
                    esi = ed_esi(r, dx, True)
                    triage, provider, dispo_t, depart, ed_dispo = ed_timeline(start, esi, True, dx)
                    ed_info = (start, triage, provider, dispo_t, depart, "Observation" if t == "Observation" else "Admitted")
                    admit_dt = dispo_t
                    start_floor = depart
                else:
                    admit_dt = start
                    start_floor = start
                if t == "Observation":
                    hours = clamp(lognorm(r, 21, 0.38), 8, 46)
                else:
                    mean_los = C.DX[dx][9] * (1.35 if dept in C.ICU_DEPT.values() else 1.0)
                    if age > 65:
                        mean_los *= 1 + (age - 65) * 0.008
                    hours = clamp(lognorm(r, mean_los * 24 * 0.9, 0.45), 14, 24 * 45)
                raw_end = start_floor + timedelta(hours=hours)
                if t == "Inpatient":
                    # discharges mostly happen late morning to early evening
                    end_day = raw_end.date()
                    tod = clamp(r.gauss(14.2, 2.2), 8.5, 21.5)
                    disch_dt = datetime.combine(end_day, datetime.min.time()) + timedelta(hours=tod)
                    if disch_dt <= start_floor + timedelta(hours=10):
                        disch_dt = start_floor + timedelta(hours=10 + r.uniform(0, 6))
                    disch_dt = disch_dt.replace(second=0, microsecond=0)
                else:
                    disch_dt = raw_end.replace(second=0, microsecond=0)
                dispo = disposition(dx, age, dept, t, fac, len(pat["Chronic"]))
                los_days = (disch_dt - admit_dt).total_seconds() / 86400
            admit_dt = admit_dt.replace(second=0, microsecond=0)
            disch_dt = disch_dt.replace(second=0, microsecond=0)
            if disch_dt > datetime(2025, 12, 31, 23, 59):
                # still in house at extract time -> drop (keeps the extract simple: discharged encounters only)
                enc_n -= 1
                continue
            attending = pick_provider(fac, dept, route, dx, t, pat)
            payer = pat["PrimaryPayerID"]
            if dx in ("S93.401A", "M54.50", "S06.0X0A") and 18 <= age <= 64 and r.random() < 0.14:
                payer = "PY08"
            chg = charges(t, dx, dept, round(los_days), esi or 3, route)
            enc = {
                "EncounterID": enc_id, "PatientID": pat["PatientID"], "EncounterType": t, "FacilityID": fac,
                "DeptID": dept, "AttendingProviderID": attending, "AdmitSource": src, "AdmitDateTime": admit_dt,
                "DischargeDateTime": disch_dt, "PrimaryDxCode": dx, "DischargeDisposition": dispo,
                "PayerID": payer, "TotalCharges": round(chg, 2), "_pidx": pidx, "_age": age, "_esi": esi,
            }
            encounters.append(enc)
            busy_until[pidx] = disch_dt
            if ed_info:
                a, tr, pv, dt_, dep, edd = ed_info
                busy_until[pidx] = max(disch_dt, dep)
                add_ed_row(enc_id, pat, fac, a, tr, pv, dt_, dep, edd, esi, dx, age)
                enc["_ed_departure"] = dep
            if dispo == "Expired":
                deceased[pidx] = disch_dt
                pat["DeceasedDate"] = disch_dt.date()
            # schedule a readmission
            if t == "Inpatient" and dispo not in ("Expired", "Hospice", "Transfer to Another Hospital"):
                risk = 0.55 * C.DX[dx][12] * (1.25 if age >= 75 else 1.0) * (1.3 if len(pat["Chronic"]) >= 3 else 1.0)
                risk *= 1.5 if dispo == "Left AMA" else 1.2 if dispo == "Skilled Nursing Facility" else 1.0
                if r.random() < risk:
                    while True:
                        gap = 1 + int(r.expovariate(1 / 9))
                        if gap <= 30:
                            break
                    rdx_opts = [dx] * 9 + ["A41.9", "J18.9", "N39.0", "N17.9", "E87.1", "K92.2", "E86.0"] * 1
                    for key, codes in C.CHRONIC_LINKS.items():
                        if key in pat["Chronic"] and key in ("HF", "COPD", "DM"):
                            rdx_opts += codes * 2
                    rdx = r.choice(rdx_opts)
                    if C.DX[rdx][3][0] == 0 or not dx_eligible(rdx, age, pat["Sex"]):
                        rdx = "J18.9" if dx_eligible("J18.9", age, pat["Sex"]) else dx
                    scheduled.setdefault(disch_dt.date() + timedelta(days=gap), []).append((pidx, rdx))

    # registration dates
    rr = R("registration")
    first_seen = {}
    for e in encounters:
        first_seen.setdefault(e["_pidx"], e["AdmitDateTime"].date())
    for idx, p in enumerate(patients):
        anchor = first_seen.get(idx, date(2025, 12, 31) - timedelta(days=rr.randint(0, 700)))
        earliest = max(p["DOB"], date(2008, 1, 1))
        span = (anchor - earliest).days
        p["RegistrationDate"] = anchor - timedelta(days=int(span * rr.random() ** 1.6)) if span > 0 else anchor

    # readmission flags
    by_pat = {}
    for e in encounters:
        if e["EncounterType"] == "Inpatient":
            by_pat.setdefault(e["PatientID"], []).append(e)
    for e in encounters:
        e["Readmit30"] = None
    for pid, lst in by_pat.items():
        lst.sort(key=lambda z: z["AdmitDateTime"])
        for i, e in enumerate(lst):
            flag = "N"
            if e["DischargeDisposition"] != "Expired":
                for nxt in lst[i + 1:]:
                    gap = (nxt["AdmitDateTime"].date() - e["DischargeDateTime"].date()).days
                    if nxt["AdmitDateTime"] > e["DischargeDateTime"] and gap <= 30:
                        flag = "Y"
                        break
                    if gap > 30:
                        break
            e["Readmit30"] = flag
    return encounters, ed_rows


# ---------------------------------------------------------------------------
# Claims
# ---------------------------------------------------------------------------
def gen_claims(encounters):
    r = R("claims")
    rows = []
    reasons = [x for x, _ in C.DENIAL_REASONS]
    rweights = [w for _, w in C.DENIAL_REASONS]
    for i, e in enumerate(encounters):
        pid = e["PayerID"]
        _, pname, ptype, rate, days_pay, timely, denial = PAYER[pid]
        svc = e["DischargeDateTime"].date()
        lag = int(clamp(lognorm(r, 4.5, 0.55), 1, 25))
        if r.random() < 0.035:
            lag = r.randint(35, 140)
        submit = svc + timedelta(days=lag)
        billed = e["TotalCharges"]
        allowed = round(billed * rate * math.exp(r.gauss(0, 0.07)), 2)
        if ptype == "Commercial":
            pt_resp = round(allowed * clamp(r.gauss(0.11, 0.05), 0.0, 0.3), 2)
        elif ptype in ("Government",) and pid == "PY01":
            pt_resp = round(allowed * clamp(r.gauss(0.09, 0.04), 0.0, 0.2), 2)
        elif ptype == "Medicare Advantage":
            pt_resp = round(allowed * clamp(r.gauss(0.06, 0.03), 0.0, 0.2), 2)
        elif ptype == "Self-Pay":
            pt_resp = allowed
        else:
            pt_resp = 0.0
        status, reason, paid, paid_date = "Paid", None, 0.0, None
        expected_pay = round(allowed - pt_resp, 2)
        pay_days = int(clamp(lognorm(r, days_pay, 0.35), 5, 180))
        if ptype == "Self-Pay":
            x = r.random()
            pay_date = submit + timedelta(days=pay_days)
            if pay_date > AS_OF or x < 0.35:
                status = "Pending"
            elif x < 0.55:
                status, paid, paid_date = "Paid", allowed, pay_date
            else:
                status, paid, paid_date = "Partially Paid", round(allowed * r.uniform(0.1, 0.8), 2), pay_date
            expected_pay = 0.0
        else:
            late = timely and lag > timely
            deny_p = denial * (1.6 if e["EncounterType"] == "Inpatient" else 1.0)
            pay_date = submit + timedelta(days=pay_days)
            if late:
                status, reason = "Denied", "Timely Filing"
                paid_date = None
            elif r.random() < deny_p:
                status = "Denied"
                reason = wchoice(r, reasons, rweights)
                if reason == "Timely Filing":
                    reason = "Authorization Required"
            elif pay_date > AS_OF:
                status = "Pending"
            elif r.random() < 0.07:
                status = "Partially Paid"
                reason = wchoice(r, ["Coding Error", "Medical Necessity", "Missing Documentation"], [5, 3, 2])
                paid = round(expected_pay * r.uniform(0.35, 0.9), 2)
                paid_date = pay_date
            else:
                paid, paid_date = expected_pay, pay_date
            if status == "Denied" and r.random() < 0.3 and submit + timedelta(days=pay_days + 30) <= AS_OF:
                status = "Appealed"
        rows.append({
            "ClaimID": f"CLM{500000 + i + 1}", "EncounterID": e["EncounterID"], "PatientID": e["PatientID"],
            "PayerID": pid, "ServiceDate": svc, "SubmitDate": submit, "BilledAmount": billed,
            "AllowedAmount": allowed if status not in ("Denied", "Appealed") else 0.0,
            "PatientResponsibility": pt_resp if status not in ("Denied", "Appealed") else 0.0,
            "PaidAmount": paid, "ClaimStatus": status, "DenialReason": reason, "PaidDate": paid_date,
        })
    return rows


# ---------------------------------------------------------------------------
# Labs and medications
# ---------------------------------------------------------------------------
LAB = {t[0]: t for t in C.LAB_TESTS}


def lab_value(r, code, dx, pat, age):
    t = LAB[code]
    mean, sd = t[5], t[6]
    chronic = pat["Chronic"]
    v = r.gauss(mean, sd)
    if code == "GLU":
        if dx == "E11.65":
            v = r.gauss(330, 95)
        elif "DM" in chronic:
            v = r.gauss(178, 55)
        elif dx == "A41.9":
            v = r.gauss(152, 38)
        else:
            v = r.gauss(104, 18)
    elif code == "NA" and dx == "E87.1":
        v = r.gauss(125.5, 4)
    elif code == "K" and dx == "N17.9":
        v = r.gauss(5.4, 0.6)
    elif code == "CREAT":
        if dx == "N17.9":
            v = r.gauss(3.1, 1.1)
        elif "CKD" in chronic:
            v = r.gauss(2.0, 0.55)
        elif dx == "A41.9":
            v = r.gauss(1.6, 0.55)
        if age < 12:
            v = r.gauss(0.45, 0.1)
        v = max(v, 0.2)
    elif code == "WBC":
        v = {"A41.9": r.gauss(17, 5.5) if r.random() > 0.1 else r.gauss(3.1, 0.6),
             "J18.9": r.gauss(14, 4), "K35.80": r.gauss(15.2, 3.5), "D70.1": abs(r.gauss(0.8, 0.4)) + 0.1,
             "N39.0": r.gauss(12.2, 3.0), "L03.115": r.gauss(12.8, 3.2)}.get(dx, v)
    elif code == "HGB":
        v = r.gauss(13.2, 1.1) if pat["Sex"] == "F" else r.gauss(14.8, 1.2)
        if dx == "K92.2":
            v = r.gauss(8.1, 1.4)
        elif "CANCER" in chronic:
            v = r.gauss(10.6, 1.3)
        elif "CKD" in chronic:
            v = r.gauss(10.9, 1.2)
    elif code == "PLT" and dx == "D70.1":
        v = abs(r.gauss(62, 35)) + 8
    elif code == "TROP":
        if dx == "I21.4":
            v = lognorm(r, 480, 1.0)
        elif dx in ("I50.9", "A41.9", "N17.9") and r.random() < 0.5:
            v = r.uniform(30, 95)
        elif r.random() < 0.06:
            v = r.uniform(35, 70)
        else:
            v = abs(r.gauss(6, 4))
    elif code == "BNP":
        if dx == "I50.9":
            v = lognorm(r, 960, 0.6)
        elif dx in ("I48.91",):
            v = r.uniform(140, 420)
        else:
            v = abs(r.gauss(42, 24))
    elif code == "LACT":
        if dx == "A41.9":
            v = lognorm(r, 3.0, 0.45)
        elif dx in ("J96.01", "K92.2"):
            v = lognorm(r, 2.2, 0.35)
        else:
            v = abs(r.gauss(1.2, 0.3)) + 0.3
    elif code == "A1C":
        if dx == "E11.65":
            v = r.gauss(10.2, 1.6)
        elif "DM" in chronic:
            v = r.gauss(8.0, 1.3)
        else:
            v = r.gauss(5.4, 0.3)
    elif code == "LDL":
        v = r.gauss(152, 30) if dx == "E78.5" else r.gauss(112, 30)
    elif code == "INR":
        if "AFIB" in chronic and r.random() < 0.4:
            v = r.gauss(2.5, 0.6) if dx != "K92.2" else r.gauss(3.9, 1.1)
        else:
            v = r.gauss(1.03, 0.08)
    elif code == "TSH":
        v = abs(r.gauss(1.9, 0.95)) + 0.05 if r.random() > 0.08 else r.choice([r.uniform(0.02, 0.35), r.uniform(4.5, 14)])
    elif code == "BUN":
        pass
    v = max(v, 0.0)
    dec = t[7]
    v = round(v, dec) if dec else float(round(v))
    return v


def lab_flag(code, v):
    t = LAB[code]
    lo, hi, clo, chi = t[3], t[4], t[8], t[9]
    if clo is not None and v < clo:
        return "LL"
    if chi is not None and v > chi:
        return "HH"
    if v < lo:
        return "L"
    if v > hi:
        return "H"
    return "N"


def gen_labs(encounters, patients_by_id):
    r = R("labs")
    rows = []
    n = 0
    for e in encounters:
        t, dx = e["EncounterType"], e["PrimaryDxCode"]
        pat = patients_by_id[e["PatientID"]]
        age = e["_age"]
        tests = []
        if t in ("Inpatient", "Observation"):
            tests = ["GLU", "CREAT", "WBC", "HGB"]
            if r.random() < 0.25 or dx in ("E87.1", "N17.9", "E11.65", "F10.239"):
                tests += ["NA", "K"]
            if r.random() < 0.12:
                tests += ["BUN", "PLT"]
        elif t == "Emergency":
            if (C.DX[dx][15] <= 3.4 and r.random() < 0.6) or r.random() < 0.12:
                tests = ["GLU", "CREAT", "WBC", "HGB"][: r.randint(2, 4)]
        else:
            if dx in ("E11.9", "E11.65"):
                tests = ["A1C"] + (["GLU", "CREAT"] if r.random() < 0.5 else [])
            elif dx in ("Z00.00", "E78.5"):
                tests = ["LDL"] + (["A1C"] if "DM" in pat["Chronic"] or r.random() < 0.25 else []) + (["TSH"] if r.random() < 0.3 else [])
            elif dx == "I48.91" and "AFIB" in pat["Chronic"] and r.random() < 0.5:
                tests = ["INR"]
            elif dx == "Z51.11":
                tests = ["WBC", "HGB", "PLT"]
        if dx in ("I21.4", "R07.9") or (dx == "I50.9" and t != "Outpatient"):
            tests.append("TROP")
        if dx == "I50.9" and t != "Outpatient":
            tests.append("BNP")
        if dx in ("A41.9", "J96.01") or (dx == "K92.2" and r.random() < 0.5):
            tests.append("LACT")
        if dx in ("E11.65",) and t != "Outpatient":
            tests.append("A1C")
        if dx in ("K92.2", "I48.91", "I63.9") and t != "Outpatient":
            tests.append("INR")
        seen = []
        for code in tests:
            if code in seen:
                continue
            seen.append(code)
            n += 1
            start = e["AdmitDateTime"] if t != "Emergency" else e["AdmitDateTime"]
            order = start + timedelta(minutes=r.randint(5, 75) if t != "Outpatient" else r.randint(5, 30))
            stat = t in ("Emergency", "Inpatient", "Observation") and (code in ("TROP", "LACT", "K") or r.random() < 0.62)
            collected = order + timedelta(minutes=r.randint(4, 35))
            if stat:
                tat = clamp(lognorm(r, 46, 0.4), 14, 240)
            elif code in ("A1C", "LDL", "TSH"):
                tat = clamp(lognorm(r, 14 * 60, 0.5), 180, 3000)
            else:
                tat = clamp(lognorm(r, 170, 0.5), 40, 900)
            resulted = collected + timedelta(minutes=tat)
            v = lab_value(r, code, dx, pat, age)
            lt = LAB[code]
            rows.append({
                "LabResultID": f"LAB{700000 + n}", "EncounterID": e["EncounterID"], "PatientID": e["PatientID"],
                "TestCode": code, "TestName": lt[1], "ResultValue": v, "Units": lt[2], "RefLow": lt[3], "RefHigh": lt[4],
                "AbnormalFlag": lab_flag(code, v), "Priority": "STAT" if stat else "Routine",
                "OrderDateTime": order.replace(second=0, microsecond=0), "CollectedDateTime": collected.replace(second=0, microsecond=0),
                "ResultedDateTime": resulted.replace(second=0, microsecond=0),
            })
    return rows


MED = {m[0]: m for m in C.MEDICATIONS}
MED_MAP = {
    "A41.9": [("Piperacillin-Tazobactam", .7), ("Vancomycin", .6), ("Ceftriaxone", .25), ("Lactated Ringer's", .85), ("Sodium Chloride 0.9%", .4), ("Acetaminophen", .6), ("Heparin", .45), ("Pantoprazole", .3)],
    "J18.9": [("Ceftriaxone", .8), ("Azithromycin", .7), ("Sodium Chloride 0.9%", .35), ("Acetaminophen", .6), ("Enoxaparin", .55)],
    "U07.1": [("Methylprednisolone", .55), ("Enoxaparin", .7), ("Acetaminophen", .7), ("Ceftriaxone", .2)],
    "J44.1": [("Ipratropium-Albuterol", .85), ("Methylprednisolone", .6), ("Prednisone", .35), ("Azithromycin", .45), ("Enoxaparin", .5)],
    "J96.01": [("Ipratropium-Albuterol", .6), ("Methylprednisolone", .55), ("Piperacillin-Tazobactam", .4), ("Heparin", .5), ("Pantoprazole", .4)],
    "J45.909": [("Albuterol", .9), ("Ipratropium-Albuterol", .5), ("Prednisone", .7), ("Methylprednisolone", .25)],
    "I50.9": [("Furosemide", .95), ("Metoprolol Tartrate", .6), ("Lisinopril", .4), ("Atorvastatin", .4), ("Enoxaparin", .5), ("Potassium Chloride", .45)],
    "I21.4": [("Aspirin", .95), ("Heparin", .8), ("Atorvastatin", .9), ("Metoprolol Tartrate", .8), ("Morphine", .3)],
    "I48.91": [("Diltiazem", .5), ("Metoprolol Tartrate", .7), ("Apixaban", .55), ("Warfarin", .15)],
    "R07.9": [("Aspirin", .75), ("Acetaminophen", .35), ("Sodium Chloride 0.9%", .3)],
    "I63.9": [("Aspirin", .8), ("Atorvastatin", .8), ("Heparin", .5), ("Sodium Chloride 0.9%", .4)],
    "N39.0": [("Ceftriaxone", .7), ("Cephalexin", .45), ("Acetaminophen", .5)],
    "L03.115": [("Cephalexin", .5), ("Vancomycin", .5), ("Ceftriaxone", .3), ("Acetaminophen", .5)],
    "E11.65": [("Insulin Lispro", .85), ("Insulin Glargine", .6), ("Sodium Chloride 0.9%", .6), ("Potassium Chloride", .35)],
    "E86.0": [("Sodium Chloride 0.9%", .9), ("Lactated Ringer's", .3), ("Ondansetron", .7)],
    "A08.4": [("Sodium Chloride 0.9%", .8), ("Ondansetron", .8)],
    "E87.1": [("Sodium Chloride 0.9%", .7)],
    "N17.9": [("Sodium Chloride 0.9%", .8), ("Lactated Ringer's", .3)],
    "K92.2": [("Pantoprazole", .95), ("Sodium Chloride 0.9%", .7), ("Ondansetron", .3)],
    "K35.80": [("Piperacillin-Tazobactam", .5), ("Ceftriaxone", .45), ("Morphine", .6), ("Ondansetron", .7), ("Lactated Ringer's", .8), ("Ketorolac", .3)],
    "K80.20": [("Ceftriaxone", .4), ("Hydromorphone", .5), ("Ondansetron", .7), ("Lactated Ringer's", .7), ("Ketorolac", .4)],
    "K56.609": [("Hydromorphone", .6), ("Ondansetron", .8), ("Sodium Chloride 0.9%", .9), ("Pantoprazole", .4), ("Enoxaparin", .4)],
    "S72.001A": [("Oxycodone", .8), ("Hydromorphone", .5), ("Acetaminophen", .9), ("Enoxaparin", .9), ("Ondansetron", .5), ("Ceftriaxone", .3)],
    "M17.11": [("Oxycodone", .85), ("Acetaminophen", .9), ("Enoxaparin", .85), ("Ondansetron", .5), ("Ketorolac", .4)],
    "M16.11": [("Oxycodone", .85), ("Acetaminophen", .9), ("Enoxaparin", .85), ("Ondansetron", .5), ("Ketorolac", .4)],
    "O80": [("Oxytocin", .95), ("Acetaminophen", .6), ("Ketorolac", .45)],
    "O82": [("Oxytocin", .95), ("Oxycodone", .7), ("Acetaminophen", .7), ("Ketorolac", .5), ("Enoxaparin", .3)],
    "J21.9": [("Albuterol", .35), ("Sodium Chloride 0.9%", .5), ("Acetaminophen", .7)],
    "H66.90": [("Amoxicillin", .9), ("Acetaminophen", .5)],
    "F32.9": [("Sertraline", .55), ("Olanzapine", .3), ("Lorazepam", .3)],
    "F20.9": [("Olanzapine", .8), ("Haloperidol", .45), ("Lorazepam", .4)],
    "F10.239": [("Lorazepam", .9), ("Thiamine", .95), ("Sodium Chloride 0.9%", .5)],
    "F41.9": [("Lorazepam", .5)],
    "C34.90": [("Morphine", .5), ("Ondansetron", .6), ("Enoxaparin", .4), ("Methylprednisolone", .2)],
    "D70.1": [("Filgrastim", .8), ("Piperacillin-Tazobactam", .85), ("Ondansetron", .5), ("Acetaminophen", .6)],
    "C50.911": [("Ondansetron", .5), ("Acetaminophen", .5), ("Enoxaparin", .4)],
    "G40.909": [("Levetiracetam", .9), ("Lorazepam", .4)],
    "G43.909": [("Ketorolac", .6), ("Ondansetron", .45), ("Sodium Chloride 0.9%", .3), ("Acetaminophen", .3)],
    "M54.50": [("Ketorolac", .5), ("Acetaminophen", .45), ("Oxycodone", .15)],
    "S93.401A": [("Acetaminophen", .4), ("Ketorolac", .3)],
    "S06.0X0A": [("Acetaminophen", .6), ("Ondansetron", .3)],
    "R10.9": [("Ondansetron", .6), ("Morphine", .3), ("Ketorolac", .3), ("Sodium Chloride 0.9%", .5)],
    "R55": [("Sodium Chloride 0.9%", .7)],
    "J06.9": [("Acetaminophen", .4), ("Azithromycin", .08)],
    "I10": [("Lisinopril", .5), ("Metoprolol Tartrate", .3)],
    "Z51.11": [("Carboplatin", .9), ("Paclitaxel", .8), ("Ondansetron", .9)],
    "Z23": [("Influenza Vaccine", 1.0)],
}
DOSES_PER_DAY = {"Q24H": 1, "Daily": 1, "Q12H": 2, "BID": 2, "Q8H": 3, "Q6H": 4, "Q4H": 6, "Q6H PRN": 2,
                 "Q4H PRN": 2.5, "Q3H PRN": 3, "Continuous": 2, "AC & HS": 4, "Once": 0}


def gen_meds(encounters):
    r = R("meds")
    rows = []
    n = 0
    for e in encounters:
        t, dx = e["EncounterType"], e["PrimaryDxCode"]
        opts = MED_MAP.get(dx, [])
        if t == "Outpatient" and dx not in ("Z51.11", "Z23"):
            continue
        if t == "Emergency" and r.random() < 0.3:
            continue
        hours = (e["DischargeDateTime"] - e["AdmitDateTime"]).total_seconds() / 3600
        chosen = [m for m, p in opts if r.random() < (p if t != "Emergency" else p * 0.6)]
        if t == "Inpatient" and e["_age"] >= 18 and "Enoxaparin" not in chosen and "Heparin" not in chosen and r.random() < 0.35:
            chosen.append("Heparin")
        if t in ("Inpatient", "Observation") and r.random() < 0.25 and "Acetaminophen" not in chosen:
            chosen.append("Acetaminophen")
        for m in chosen:
            name, cls, dose, route, freq, cost, ha = MED[m]
            n += 1
            if t == "Outpatient":
                order = e["AdmitDateTime"] + timedelta(minutes=r.randint(5, 40))
            elif t == "Emergency":
                order = e["AdmitDateTime"] + timedelta(minutes=r.randint(15, max(16, int(hours * 60 * 0.8))))
            else:
                order = e["AdmitDateTime"] + timedelta(minutes=r.randint(10, 360))
            per_day = DOSES_PER_DAY[freq]
            if per_day == 0 or t in ("Emergency", "Outpatient"):
                doses = 1 if freq != "Continuous" else max(1, round(hours / 8))
            else:
                remain = max((e["DischargeDateTime"] - order).total_seconds() / 86400, 0.25)
                doses = max(1, round(per_day * remain * r.uniform(0.75, 1.0)))
            rows.append({
                "MedOrderID": f"RX{800000 + n}", "EncounterID": e["EncounterID"], "PatientID": e["PatientID"],
                "OrderDateTime": order.replace(second=0, microsecond=0), "MedicationName": name, "DrugClass": cls,
                "Dose": dose, "Route": route, "Frequency": freq, "DosesDispensed": doses, "UnitCost": cost,
                "HighAlert": ha, "OrderingProviderID": e["AttendingProviderID"],
            })
    return rows


# ---------------------------------------------------------------------------
# Patient satisfaction
# ---------------------------------------------------------------------------
def gen_surveys(encounters, ed_by_enc):
    r = R("surveys")
    rows = []
    n = 0
    fac_eff = {"F01": -0.05, "F02": 0.28, "F03": 0.06}
    for e in encounters:
        if e["EncounterType"] not in ("Inpatient", "Observation") or e["DischargeDisposition"] in ("Expired", "Hospice", "Transfer to Another Hospital"):
            continue
        resp = {"PY01": 0.38, "PY02": 0.36, "PY03": 0.2, "PY07": 0.15}.get(e["PayerID"], 0.3)
        if e["_age"] < 18:
            resp *= 0.6
        if r.random() > resp:
            continue
        received = e["DischargeDateTime"].date() + timedelta(days=r.randint(6, 48))
        if received > AS_OF:
            continue
        los = (e["DischargeDateTime"] - e["AdmitDateTime"]).total_seconds() / 86400
        s = r.gauss(0, 1) + fac_eff.get(e["FacilityID"], 0) - 0.035 * los + (0.12 if e["_age"] >= 65 else 0)
        ed = ed_by_enc.get(e["EncounterID"])
        boarding = None
        if ed and ed["DepartureDateTime"] and ed["DispositionDateTime"]:
            boarding = (ed["DepartureDateTime"] - ed["DispositionDateTime"]).total_seconds() / 3600
            if boarding > 6:
                s -= 0.4
        if e["DeptID"] in C.ICU_DEPT.values():
            s -= 0.1
        if e["DischargeDisposition"] == "Left AMA":
            s -= 0.9

        def item(bias):
            p4 = 1 / (1 + math.exp(-(1.15 + 0.95 * s + bias)))
            if r.random() < p4:
                return 4
            x = r.random() + 0.08 * s
            return 3 if x > 0.3 else 2 if x > 0.08 else 1
        quiet_bias = -0.65 - (0.35 if e["DeptID"] in C.ICU_DEPT.values() else 0)
        rating = int(clamp(round(8.4 + 1.35 * s + r.gauss(0, 0.7)), 0, 10))
        if rating >= 9:
            rec = wchoice(r, ["Definitely Yes", "Probably Yes", "Probably No"], [86, 13, 1])
        elif rating >= 7:
            rec = wchoice(r, ["Definitely Yes", "Probably Yes", "Probably No", "Definitely No"], [30, 58, 10, 2])
        else:
            rec = wchoice(r, ["Probably Yes", "Probably No", "Definitely No"], [22, 45, 33])
        comment = None
        if r.random() < 0.46:
            if s > 0.45:
                comment = r.choice(C.SURVEY_COMMENTS_POS)
            elif s < -0.45:
                comment = r.choice(C.SURVEY_COMMENTS_NEG)
            else:
                comment = r.choice(C.SURVEY_COMMENTS_MIXED + C.SURVEY_COMMENTS_POS[:3])
        n += 1
        rows.append({
            "SurveyID": f"SV{30000 + n}", "EncounterID": e["EncounterID"], "FacilityID": e["FacilityID"], "DeptID": e["DeptID"],
            "DischargeDate": e["DischargeDateTime"].date(), "SurveyReceivedDate": received,
            "NurseCommunication": item(0.55), "DoctorCommunication": item(0.42), "StaffResponsiveness": item(-0.35),
            "Cleanliness": item(0.1), "Quietness": item(quiet_bias), "MedicationExplained": item(-0.25),
            "DischargeInfoProvided": "Yes" if r.random() < clamp(0.86 + 0.05 * s, 0.4, 0.99) else "No",
            "OverallRating": rating, "WouldRecommend": rec, "Comment": comment,
        })
    return rows


# ---------------------------------------------------------------------------
# Daily census
# ---------------------------------------------------------------------------
CENSUS_PARAMS = {
    # dept: (base occupancy, winter bump, ALOS days, weekday pattern strength)
    "D110": (0.87, 0.06, 4.3, 0.02), "D111": (0.85, 0.06, 4.1, 0.02), "D120": (0.81, 0.05, 3.4, 0.03),
    "D130": (0.80, 0.07, 3.6, 0.0), "D140": (0.74, 0.0, 2.8, 0.12), "D150": (0.79, 0.02, 5.0, 0.03),
    "D160": (0.76, 0.03, 4.0, 0.02), "D170": (0.62, 0.0, 2.3, 0.04), "D180": (0.52, 0.22, 2.6, 0.0),
    "D190": (0.91, 0.02, 8.1, 0.0), "D195": (0.66, 0.05, 1.0, 0.05), "D210": (0.83, 0.07, 3.9, 0.02),
    "D230": (0.74, 0.08, 3.4, 0.0), "D270": (0.55, 0.0, 2.2, 0.04), "D310": (0.86, 0.06, 4.0, 0.02),
    "D320": (0.80, 0.05, 3.3, 0.03), "D330": (0.79, 0.07, 3.5, 0.0), "D340": (0.71, 0.0, 2.7, 0.12),
}


def gen_census():
    r = R("census")
    rows = []
    for dept, (base, winter, alos, wk) in CENSUS_PARAMS.items():
        beds = DEPT[dept][5]
        fac = DEPT[dept][2]
        dev = 0.0
        prev = round(base * beds)
        for day in daterange(START, END):
            mu = base + (winter if day.month in (12, 1, 2) else winter / 2 if day.month in (11, 3) else -winter / 3 if day.month in (6, 7, 8) else 0)
            mu += wk * (0.5 if day.weekday() in (1, 2, 3) else -1.0 if day.weekday() in (5, 6) else 0)
            mu += 0.015 if day.year == 2025 else 0
            dev = 0.78 * dev + r.gauss(0, 0.035)
            occ = clamp(mu + dev, 0.2, 1.12)
            census = int(clamp(round(occ * beds), 0, beds + 4))
            lam = max(0.5, census / alos)
            adm = max(0, int(round(r.gauss(lam, math.sqrt(lam)))))
            dis = prev + adm - census
            if dis < 0:
                adm -= dis
                dis = 0
            rows.append({"CensusDate": day, "FacilityID": fac, "DeptID": dept, "StaffedBeds": beds,
                         "Admissions": adm, "Discharges": dis, "MidnightCensus": census})
            prev = census
    rows.sort(key=lambda z: (z["CensusDate"], z["DeptID"]))
    return rows


# ---------------------------------------------------------------------------
# Employees and shifts (Bluestone Memorial nursing & clinical support)
# ---------------------------------------------------------------------------
STAFF_UNITS = ["D100", "D110", "D111", "D120", "D130", "D140", "D150", "D160", "D170", "D180", "D190", "D195"]
CERTS = {
    "D100": "BLS;ACLS;PALS;TNCC", "D130": "BLS;ACLS;CCRN", "D120": "BLS;ACLS", "D170": "BLS;NRP",
    "D180": "BLS;PALS", "D150": "BLS;OCN", "D160": "BLS;ACLS;NIHSS", "D190": "BLS;CPI",
}


def gen_employees():
    r = R("employees")
    rows = []
    n = 0
    used = set()
    for dept in STAFF_UNITS:
        beds = DEPT[dept][5] or 0
        if dept == "D100":
            plan = {"Registered Nurse": 40, "Licensed Practical Nurse": 2, "Patient Care Technician": 10, "Unit Secretary": 5, "Nurse Manager": 1}
        else:
            rn = round(beds * (2.4 if dept == "D130" else 1.25))
            plan = {"Registered Nurse": rn, "Licensed Practical Nurse": round(beds * 0.1),
                    "Certified Nursing Assistant": round(beds * (0.3 if dept == "D130" else 0.45)),
                    "Unit Secretary": 3, "Nurse Manager": 1}
            if dept == "D130":
                plan["Respiratory Therapist"] = 8
            elif dept in ("D110", "D111"):
                plan["Respiratory Therapist"] = 3
            elif dept == "D120":
                plan["Respiratory Therapist"] = 2
        for title, count in plan.items():
            for _ in range(count):
                n += 1
                while True:
                    sex = "F" if r.random() < (0.86 if title in ("Registered Nurse", "Licensed Practical Nurse", "Certified Nursing Assistant") else 0.6) else "M"
                    first = r.choice(C.FEMALE_FIRST if sex == "F" else C.MALE_FIRST)
                    last = r.choice(C.LAST_NAMES)
                    if (first, last) not in used:
                        used.add((first, last))
                        break
                agency = title == "Registered Nurse" and r.random() < 0.06
                hire = date(2025, 12, 31) - timedelta(days=int(r.expovariate(1 / 2400)))
                hire = max(hire, date(1990, 1, 1))
                if agency:
                    hire = date(2025, 1, 1) + timedelta(days=r.randint(0, 300))
                    etype, fte = "Agency", 0.9
                elif title == "Nurse Manager":
                    etype, fte = "Full-Time", 1.0
                else:
                    etype = wchoice(r, ["Full-Time", "Part-Time", "Per Diem"], [72, 20, 8])
                    fte = 1.0 if etype == "Full-Time" and r.random() < 0.45 else 0.9 if etype == "Full-Time" else r.choice([0.5, 0.6, 0.7]) if etype == "Part-Time" else 0.0
                years = (date(2025, 12, 31) - hire).days / 365.25
                base = {"Registered Nurse": (37.5, 0.85, 58), "Licensed Practical Nurse": (25, 0.4, 33),
                        "Certified Nursing Assistant": (17.5, 0.3, 23.5), "Patient Care Technician": (18.5, 0.32, 25),
                        "Respiratory Therapist": (31, 0.6, 43), "Unit Secretary": (17.25, 0.25, 22.5),
                        "Nurse Manager": (52, 0.6, 72)}[title]
                rate = min(base[0] + base[1] * years + r.uniform(-1, 1.5), base[2])
                if agency:
                    rate = r.uniform(88, 115)
                status, term = "Active", None
                if not agency and r.random() < 0.085:
                    term = date(2024, 1, 15) + timedelta(days=r.randint(0, 710))
                    if term <= hire:
                        term = None
                    else:
                        status = "Terminated"
                if status == "Active" and r.random() < 0.02:
                    status = "Leave of Absence"
                pref = "Day" if title in ("Nurse Manager",) else wchoice(r, ["Day", "Night", "Rotating"], [45, 35, 20])
                certs = None
                if title == "Registered Nurse":
                    certs = CERTS.get(dept, "BLS;ACLS" if r.random() < 0.6 else "BLS")
                    if r.random() < 0.12:
                        certs = "BLS"  # lapsed advanced certs -> great for audits
                elif title in ("Licensed Practical Nurse", "Certified Nursing Assistant", "Patient Care Technician", "Respiratory Therapist"):
                    certs = "BLS" if title != "Respiratory Therapist" else "BLS;ACLS;RRT"
                rows.append({
                    "EmployeeID": f"EMP{2000 + n}", "FirstName": first, "LastName": last, "JobTitle": title, "DeptID": dept,
                    "FacilityID": "F01", "EmploymentType": etype, "FTE": fte, "HourlyRate": round(rate, 2),
                    "HireDate": hire, "TermDate": term, "Status": status, "ShiftPreference": pref, "Certifications": certs,
                })
    return rows


def gen_shifts(employees):
    r = R("shifts")
    rows = []
    n = 0
    q_start, q_end = date(2025, 10, 1), date(2025, 12, 31)
    weeks = []
    d = q_start - timedelta(days=q_start.weekday())  # Monday on/before Oct 1
    while d <= q_end:
        weeks.append(d)
        d += timedelta(days=7)
    for emp in employees:
        if emp["Status"] == "Leave of Absence":
            continue
        title = emp["JobTitle"]
        eight = title in ("Unit Secretary", "Nurse Manager")
        for wk in weeks:
            if title == "Nurse Manager":
                days = [wk + timedelta(days=i) for i in range(5)]
            elif eight:
                k = 5 if emp["FTE"] >= 0.9 else 3
                days = sorted(r.sample([wk + timedelta(days=i) for i in range(7)], k))
            else:
                if emp["EmploymentType"] == "Per Diem":
                    k = wchoice(r, [0, 1, 2], [35, 45, 20])
                elif emp["EmploymentType"] == "Agency":
                    k = 3
                else:
                    k = max(1, round(emp["FTE"] * 3.33 + r.uniform(-0.4, 0.4)))
                days = sorted(r.sample([wk + timedelta(days=i) for i in range(7)], k))
            for day in days:
                if day < q_start or day > q_end or day < emp["HireDate"]:
                    continue
                if emp["TermDate"] and day >= emp["TermDate"]:
                    continue
                if eight:
                    stype = "Day 8h" if (title == "Nurse Manager" or r.random() < 0.6) else "Evening 8h"
                    s_start = datetime.combine(day, datetime.min.time()) + timedelta(hours=7 if stype == "Day 8h" else 15)
                    s_end = s_start + timedelta(hours=8.5)
                    sched = 8
                else:
                    pref = emp["ShiftPreference"]
                    night = pref == "Night" or (pref == "Rotating" and r.random() < 0.5)
                    stype = "Night 12h" if night else "Day 12h"
                    s_start = datetime.combine(day, datetime.min.time()) + timedelta(hours=19 if night else 7)
                    s_end = s_start + timedelta(hours=12.5)
                    sched = 12
                n += 1
                x = r.random()
                if x < 0.02:
                    status, cin, cout, worked = "Called Off", None, None, 0.0
                elif x < 0.024:
                    status, cin, cout, worked = "No Show", None, None, 0.0
                else:
                    status = "Worked"
                    cin = s_start + timedelta(minutes=round(clamp(r.gauss(-6, 5), -25, 18)))
                    y = r.random()
                    if y < 0.8:
                        extra = r.gauss(5, 6)
                    elif y < 0.94:
                        extra = r.uniform(15, 60)
                    else:
                        extra = r.uniform(60, 240)
                    cout = s_end + timedelta(minutes=round(clamp(extra, -20, 300)))
                    worked = round((cout - cin).total_seconds() / 3600 - 0.5, 2)
                rows.append({
                    "ShiftID": f"SH{900000 + n}", "EmployeeID": emp["EmployeeID"], "DeptID": emp["DeptID"], "ShiftDate": day,
                    "ShiftType": stype, "ScheduledStart": s_start, "ScheduledEnd": s_end, "ScheduledHours": sched,
                    "ClockIn": cin, "ClockOut": cout, "WorkedHours": worked,
                    "HoursOverSchedule": round(max(0.0, worked - sched), 2) if status == "Worked" else 0.0, "ShiftStatus": status,
                })
    rows.sort(key=lambda z: (z["ShiftDate"], z["ScheduledStart"], z["EmployeeID"]))
    for i, row in enumerate(rows):
        row["ShiftID"] = f"SH{900001 + i}"
    return rows


# ---------------------------------------------------------------------------
# Supply inventory
# ---------------------------------------------------------------------------
LOCATIONS = ["D196", "D100", "D110", "D111", "D120", "D130", "D140", "D150", "D170", "D180", "D200", "D210", "D300", "D310", "D500"]
LOC_AFFINITY = {
    "Orthopedic Implants": ["D196"], "Cardiac": ["D196", "D120", "D130", "D100"], "Surgical": ["D196", "D100", "D200", "D300"],
    "Lab Supplies": ["D500", "D100", "D130"], "Respiratory": ["D130", "D100", "D110", "D111", "D180", "D200", "D300"],
}


def gen_inventory():
    r = R("inventory")
    rows = []
    sku_n = stock_n = 0
    for cat, desc, unit, cost, vidx, perishable, variants in C.SUPPLY_TEMPLATES:
        for v in variants:
            sku_n += 1
            sku = f"SKU-{10000 + sku_n * 7}"
            name = desc.replace("{v}", v).replace("  ", " ").strip().replace(" ,", ",")
            ucost = round(cost * math.exp(r.gauss(0, 0.07)), 2)
            vendor = C.VENDORS[vidx] if r.random() > 0.15 else r.choice(C.VENDORS)
            locs_pool = LOC_AFFINITY.get(cat, LOCATIONS)
            k = min(len(locs_pool), wchoice(r, [1, 2, 3, 4], [20, 40, 28, 12]))
            for loc in sorted(r.sample(locs_pool, k)):
                stock_n += 1
                par = int(clamp(round(lognorm(r, 2400 / max(ucost, 1) ** 0.75, 0.5)), 2, 800))
                rop = max(1, round(par * r.uniform(0.25, 0.4)))
                x = r.random()
                if x < 0.03:
                    qty = 0
                elif x < 0.17:
                    qty = r.randint(0, rop)
                else:
                    qty = r.randint(rop + 1, max(rop + 2, int(par * 1.35)))
                exp_date = None
                if perishable:
                    y = r.random()
                    if y < 0.08:
                        exp_date = AS_OF - timedelta(days=r.randint(1, 120))
                    elif y < 0.18:
                        exp_date = AS_OF + timedelta(days=r.randint(1, 90))
                    else:
                        exp_date = AS_OF + timedelta(days=r.randint(91, 1000))
                rows.append({
                    "StockID": f"STK{4000 + stock_n}", "SKU": sku, "ItemDescription": name, "Category": cat, "Vendor": vendor,
                    "UnitOfMeasure": unit, "UnitCost": ucost, "LocationDeptID": loc, "QtyOnHand": qty, "ParLevel": par,
                    "ReorderPoint": rop, "ReorderQty": par - rop, "LeadTimeDays": wchoice(r, [2, 3, 5, 7, 10, 14], [20, 25, 25, 15, 10, 5]),
                    "LastCountDate": AS_OF - timedelta(days=r.randint(0, 45)), "ExpirationDate": exp_date,
                    "LotNumber": f"L{r.randint(100000, 999999)}" if perishable else None,
                })
    return rows


# ---------------------------------------------------------------------------
# Budget vs actual
# ---------------------------------------------------------------------------
EXP_CATS = ["Salaries & Wages", "Employee Benefits", "Medical Supplies", "Pharmaceuticals", "Purchased Services", "Equipment & Maintenance", "Other Operating"]


def dept_annual_budget(dept):
    _, name, fac, line, utype, beds, _ = DEPT[dept]
    if utype in ("Inpatient", "Critical Care", "Observation"):
        b = beds
        icu = utype == "Critical Care"
        sal = b * (330000 if icu else 205000)
        base = {"Salaries & Wages": sal, "Employee Benefits": sal * 0.28, "Medical Supplies": b * (52000 if icu else 31000),
                "Pharmaceuticals": b * (48000 if icu else 125000 if dept == "D150" else 24000), "Purchased Services": b * 9000,
                "Equipment & Maintenance": b * (11000 if icu else 5500), "Other Operating": b * 4200}
        margin = {"D190": 0.84, "D180": 0.9, "D170": 1.02, "D140": 1.22, "D120": 1.18, "D150": 1.12}.get(dept, 1.06)
    elif utype == "Emergency":
        sal = {"F01": 9_600_000, "F02": 3_800_000, "F03": 5_100_000}[fac]
        base = {"Salaries & Wages": sal, "Employee Benefits": sal * 0.28, "Medical Supplies": sal * 0.14, "Pharmaceuticals": sal * 0.09,
                "Purchased Services": sal * 0.12, "Equipment & Maintenance": sal * 0.05, "Other Operating": sal * 0.03}
        margin = 0.94
    elif utype == "Outpatient":
        sal = {"D400": 3_400_000, "D410": 1_900_000, "D420": 1_700_000, "D430": 2_300_000, "D440": 1_100_000, "D450": 1_600_000}[dept]
        base = {"Salaries & Wages": sal, "Employee Benefits": sal * 0.27, "Medical Supplies": sal * 0.06,
                "Pharmaceuticals": 6_200_000 if dept == "D430" else sal * 0.02, "Purchased Services": sal * 0.08,
                "Equipment & Maintenance": sal * 0.04, "Other Operating": sal * 0.05}
        margin = {"D430": 1.15, "D410": 1.2, "D420": 1.18, "D400": 0.97, "D450": 0.95, "D440": 1.01}[dept]
    else:
        sal = {"D196": 6_200_000, "D500": 5_500_000, "D510": 4_800_000, "D520": 3_100_000}[dept]
        base = {"Salaries & Wages": sal, "Employee Benefits": sal * 0.28, "Medical Supplies": {"D196": 8_500_000, "D500": 3_200_000}.get(dept, sal * 0.05),
                "Pharmaceuticals": 2_000_000 if dept == "D520" else sal * 0.01, "Purchased Services": sal * 0.1,
                "Equipment & Maintenance": 1_800_000 if dept == "D510" else sal * 0.06, "Other Operating": sal * 0.03}
        margin = {"D196": 1.25, "D500": 1.1, "D510": 1.3, "D520": None}[dept]
    return base, margin


def gen_budget():
    r = R("budget")
    rows = []
    for dept_row in C.DEPARTMENTS:
        dept = dept_row[0]
        base, margin = dept_annual_budget(dept)
        bias = {cat: r.gauss(0.01, 0.03) for cat in EXP_CATS}
        rev_bias = r.gauss(0.0, 0.035)
        for year in (2024, 2025):
            growth = 1.0 if year == 2024 else 1.035
            for month in range(1, 13):
                ms = date(year, month, 1)
                dim = ((date(year + (month == 12), month % 12 + 1, 1)) - ms).days
                frac = dim / (366 if year == 2024 else 365)
                winter = month in (12, 1, 2)
                total_exp_budget = 0.0
                for cat in EXP_CATS:
                    bud = round(base[cat] * growth * frac)
                    noise = r.gauss(0, 0.025 if cat in ("Salaries & Wages", "Employee Benefits") else 0.08 if cat != "Pharmaceuticals" else 0.11)
                    seas = (0.045 if winter and cat in ("Salaries & Wages", "Employee Benefits") else 0) + (0.05 if winter and cat in ("Medical Supplies", "Pharmaceuticals") else 0)
                    act = round(bud * (1 + bias[cat] + noise + seas))
                    total_exp_budget += bud
                    rows.append({"MonthStart": ms, "FiscalYear": year, "FiscalMonth": month, "FacilityID": dept_row[2], "DeptID": dept,
                                 "CostCenter": dept_row[6], "LineType": "Expense", "Category": cat, "BudgetAmount": bud, "ActualAmount": act})
                if margin:
                    bud = round(total_exp_budget * margin)
                    act = round(bud * (1 + rev_bias + r.gauss(0, 0.045) + (0.03 if winter else 0)))
                    rows.append({"MonthStart": ms, "FiscalYear": year, "FiscalMonth": month, "FacilityID": dept_row[2], "DeptID": dept,
                                 "CostCenter": dept_row[6], "LineType": "Revenue", "Category": "Net Patient Service Revenue",
                                 "BudgetAmount": bud, "ActualAmount": act})
    rows.sort(key=lambda z: (z["MonthStart"], z["DeptID"], z["LineType"] != "Revenue", z["Category"]))
    return rows


# ---------------------------------------------------------------------------
# Messy registration export (for the data-cleaning lessons)
# ---------------------------------------------------------------------------
def gen_messy_registrations(patients):
    r = R("messy")
    chosen = r.sample(range(len(patients)), 560)
    payer_variants = {
        "PY01": ["Medicare", "MEDICARE", "Medicare ", "medicare", "Medicare Part A"],
        "PY02": ["Silverline Medicare Advantage", "Silverline MA", "SILVERLINE MEDICARE ADV.", "Silverline Medicare Advantage "],
        "PY03": ["State Medicaid", "Medicaid", "STATE MEDICAID", "medicaid"],
        "PY04": ["Keystone Health Partners", "Keystone", "KEYSTONE HEALTH PARTNERS", "Keystone Health Ptnrs"],
        "PY05": ["Evergreen Mutual Insurance", "Evergreen Mutual", "EVERGREEN MUTUAL INS", "evergreen mutual insurance"],
        "PY06": ["Summit Choice PPO", "Summit PPO", "SUMMIT CHOICE PPO", "Summit Choice"],
        "PY07": ["Self-Pay", "Self Pay", "SELF PAY", "self-pay", "Uninsured"],
    }

    def name_fmt(p):
        first, last = p["FirstName"], p["LastName"]
        style = wchoice(r, ["LF", "LF_UP", "FL", "LF_low", "LF_sp"], [45, 20, 15, 10, 10])
        if style == "LF":
            s = f"{last}, {first}"
        elif style == "LF_UP":
            s = f"{last.upper()}, {first.upper()}"
        elif style == "FL":
            s = f"{first} {last}"
        elif style == "LF_low":
            s = f"{last.lower()}, {first.lower()}"
        else:
            s = f"  {last} ,  {first} "
        return s

    def dob_fmt(d):
        style = wchoice(r, ["us", "iso", "dmy_txt", "us_short", "dot"], [45, 25, 12, 10, 8])
        if style == "us":
            return d.strftime("%m/%d/%Y")
        if style == "iso":
            return d.strftime("%Y-%m-%d")
        if style == "dmy_txt":
            return d.strftime("%d-%b-%Y")
        if style == "us_short":
            return f"{d.month}/{d.day}/{d.year}"
        return d.strftime("%m.%d.%Y")

    def sex_fmt(s):
        return r.choice({"F": ["F", "F", "Female", "f", "FEMALE", " F"], "M": ["M", "M", "Male", "m", "MALE", "M "]}[s])

    def phone_fmt(ph):
        digits = "".join(ch for ch in ph if ch.isdigit())
        style = wchoice(r, ["paren", "dash", "dots", "raw", "plus1", "blank"], [40, 22, 12, 14, 7, 5])
        if style == "paren":
            return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
        if style == "dash":
            return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
        if style == "dots":
            return f"{digits[:3]}.{digits[3:6]}.{digits[6:]}"
        if style == "raw":
            return digits
        if style == "plus1":
            return f"+1 {digits[:3]} {digits[3:6]} {digits[6:]}"
        return ""

    def email_fmt(e):
        if not e:
            return r.choice(["", "", "N/A", "none"])
        style = wchoice(r, ["ok", "upper", "space"], [70, 15, 15])
        return e if style == "ok" else e.upper() if style == "upper" else f" {e} "

    def csz(p):
        style = wchoice(r, ["std", "nocomma", "upper", "zip4"], [55, 15, 15, 15])
        if style == "std":
            return f"{p['City']}, {p['State']} {p['ZIP']}"
        if style == "nocomma":
            return f"{p['City']} {p['State']} {p['ZIP']}"
        if style == "upper":
            return f"{p['City'].upper()}, {p['State']} {p['ZIP']}"
        return f"{p['City']}, {p['State']} {p['ZIP']}-{r.randint(1000, 9999)}"

    def mrn_fmt(m):
        style = wchoice(r, ["ok", "strip", "prefix"], [65, 20, 15])
        if style == "ok":
            return m
        if style == "strip":
            return str(int(m))
        return f"MRN-{m}"

    def reg_fmt(d):
        dtv = datetime.combine(d, datetime.min.time()) + timedelta(minutes=r.randint(7 * 60, 19 * 60))
        style = wchoice(r, ["iso", "us"], [60, 40])
        return dtv.strftime("%Y-%m-%d %H:%M:%S") if style == "iso" else dtv.strftime("%m/%d/%Y %I:%M %p")

    raw = []
    for idx in chosen:
        p = patients[idx]
        raw.append((p, {
            "PatientName": name_fmt(p), "DOB": dob_fmt(p["DOB"]), "Sex": sex_fmt(p["Sex"]), "Phone": phone_fmt(p["Phone"]),
            "Email": email_fmt(p["Email"]), "CityStateZip": csz(p), "Insurance": r.choice(payer_variants.get(p["PrimaryPayerID"], ["Medicare"])),
            "MRN": mrn_fmt(p["MRN"]), "RegisteredOn": reg_fmt(p["RegistrationDate"]),
        }))
    # duplicates: 40 exact copies, 50 re-registrations with different formatting
    dups = []
    for i in r.sample(range(len(raw)), 40):
        dups.append((raw[i][0], dict(raw[i][1])))
    for i in r.sample(range(len(raw)), 50):
        p = raw[i][0]
        dups.append((p, {
            "PatientName": name_fmt(p), "DOB": dob_fmt(p["DOB"]), "Sex": sex_fmt(p["Sex"]), "Phone": phone_fmt(p["Phone"]),
            "Email": email_fmt(p["Email"]), "CityStateZip": csz(p), "Insurance": r.choice(payer_variants.get(p["PrimaryPayerID"], ["Medicare"])),
            "MRN": mrn_fmt(p["MRN"]), "RegisteredOn": reg_fmt(p["RegistrationDate"]),
        }))
    allrows = raw + dups
    r.shuffle(allrows)
    out, truth = [], []
    for i, (p, row) in enumerate(allrows):
        rid = f"R{1001 + i}"
        out.append({"RecordID": rid, **row})
        truth.append({"RecordID": rid, "PatientID": p["PatientID"]})
    return out, truth


# ---------------------------------------------------------------------------
# Data dictionary
# ---------------------------------------------------------------------------
DICTIONARY = {
    "facilities": ("Hospitals and outpatient sites in the Bluestone Health System.", {
        "FacilityID": "Primary key (F01–F04).", "FacilityName": "Facility name.", "City": "City.", "State": "State.",
        "FacilityType": "Tertiary, community, acute, or ambulatory.", "LicensedBeds": "State-licensed bed count (0 for ambulatory).",
        "OpenedYear": "Year the facility opened."}),
    "departments": ("Clinical and support departments (cost centers). Inpatient units have staffed beds.", {
        "DeptID": "Primary key.", "DeptName": "Department / unit name (names repeat across facilities, e.g. 'Emergency Department').",
        "FacilityID": "FK → facilities.", "ServiceLine": "Clinical service line.", "UnitType": "Emergency, Inpatient, Critical Care, Observation, Outpatient, Ancillary.",
        "StaffedBeds": "Staffed beds (inpatient units only).", "CostCenter": "Four-digit finance cost center."}),
    "providers": ("Physicians and advanced practice providers.", {
        "ProviderID": "Primary key.", "FirstName": "First name.", "LastName": "Last name.", "Credential": "MD, DO, NP, PA, or CNM.",
        "Specialty": "Clinical specialty.", "PrimaryDeptID": "FK → departments.", "FacilityID": "FK → facilities.",
        "HireDate": "Date hired.", "FTE": "Full-time equivalent (1.0 = full time)."}),
    "payers": ("Insurance payers and contract terms.", {
        "PayerID": "Primary key.", "PayerName": "Payer name.", "PayerType": "Government, Medicare Advantage, Commercial, Self-Pay, Workers' Comp.",
        "AvgAllowedPctOfCharges": "Average allowed amount as a share of billed charges.", "AvgDaysToPay": "Typical days from claim submission to payment.",
        "TimelyFilingDays": "Days after service a claim must be filed (0 = not applicable)."}),
    "diagnoses": ("ICD-10-CM diagnosis lookup table (real public codes; descriptions shortened in places).", {
        "DxCode": "ICD-10-CM code (primary key).", "DxDescription": "Description.", "DxCategory": "Body-system category.",
        "ChronicCondition": "Y if typically chronic.", "ExpectedLOS": "Benchmark inpatient length of stay in days (synthetic, for O/E analysis)."}),
    "patients": ("Registered patients. Ages are realistic; names, contact details, and MRNs are fictional.", {
        "PatientID": "Primary key.", "MRN": "Medical record number — 8 digits stored as TEXT (leading zeros matter!).",
        "FirstName": "First name.", "LastName": "Last name.", "Sex": "F or M.", "DOB": "Date of birth.", "City": "City.", "State": "State.",
        "ZIP": "ZIP code (text).", "Phone": "Fictional phone number (555 area code).", "Email": "Email (example.com) — blank if none on file.",
        "PreferredLanguage": "Preferred language.", "PrimaryPayerID": "FK → payers.", "PCPProviderID": "FK → providers (blank = no PCP).",
        "HeightIn": "Height in inches.", "WeightLb": "Weight in pounds.", "ChronicConditions": "Semicolon-separated chronic condition flags (HTN;DM;HF;COPD;AFIB;CKD;ASTHMA;CANCER;PSYCH;ETOH;OA;SEIZURE).",
        "RegistrationDate": "Date first registered.", "DeceasedDate": "Date of death if the patient died in a Bluestone facility."}),
    "encounters": ("One row per patient encounter (visit or stay), Jan 2024 – Dec 2025. A representative extract, not every visit.", {
        "EncounterID": "Primary key.", "PatientID": "FK → patients.", "EncounterType": "Inpatient, Observation, Emergency, Outpatient.",
        "FacilityID": "FK → facilities.", "DeptID": "FK → departments (unit or clinic).", "AttendingProviderID": "FK → providers.",
        "AdmitSource": "Emergency Department, Direct Admit, Elective, Transfer (blank for ED/outpatient).",
        "AdmitDateTime": "Start of encounter (for ED-admitted patients: time the admit decision was made).",
        "DischargeDateTime": "End of encounter.", "PrimaryDxCode": "FK → diagnoses.", "DischargeDisposition": "Where the patient went.",
        "PayerID": "FK → payers (payer billed for this encounter).", "TotalCharges": "Gross billed charges ($).",
        "Readmit30": "Inpatient encounters only (blank otherwise): Y if the same patient has another Inpatient encounter whose AdmitDateTime is after this DischargeDateTime and whose admit DATE is at most 30 days after this discharge DATE (0–30 days); N otherwise. Stays ending in death (DischargeDisposition = Expired) are always N. Observation stays do not count as readmissions."}),
    "ed_visits": ("Emergency department timeline for every ED arrival (treat-and-release and admitted).", {
        "EDVisitID": "Primary key.", "EncounterID": "FK → encounters (the ED or resulting inpatient/observation encounter).",
        "PatientID": "FK → patients.", "FacilityID": "FK → facilities.", "ArrivalDateTime": "Door time.", "TriageDateTime": "Triage time.",
        "ProviderSeenDateTime": "First provider contact (blank = left without being seen).", "DispositionDateTime": "Decision time (discharge/admit).",
        "DepartureDateTime": "Left the ED.", "ArrivalMode": "Walk-In, Ambulance, Police, Air Transport.", "ESILevel": "Emergency Severity Index 1 (most urgent) – 5.",
        "ChiefComplaint": "Reason for visit.", "EDDisposition": "Discharged, Admitted, Observation, Transferred, LWBS, Left AMA, Expired.",
        "TempF": "Triage temperature °F.", "HeartRate": "Beats/min.", "RespRate": "Breaths/min.", "SystolicBP": "mmHg.", "DiastolicBP": "mmHg.",
        "SpO2": "Oxygen saturation %.", "PainScore": "0–10."}),
    "claims": ("One claim per encounter with payment status as of 12/31/2025.", {
        "ClaimID": "Primary key.", "EncounterID": "FK → encounters.", "PatientID": "FK → patients.", "PayerID": "FK → payers.",
        "ServiceDate": "Discharge/service date.", "SubmitDate": "Date the claim was submitted.", "BilledAmount": "Gross charges billed ($).",
        "AllowedAmount": "Contracted allowed amount (0 if denied).", "PatientResponsibility": "Copay/coinsurance/deductible owed by patient.",
        "PaidAmount": "Amount received to date.", "ClaimStatus": "Paid, Partially Paid, Denied, Appealed, Pending.",
        "DenialReason": "Reason for denial or partial payment.", "PaidDate": "Date payment received (blank if unpaid)."}),
    "lab_results": ("Laboratory results with reference ranges and turnaround timestamps.", {
        "LabResultID": "Primary key.", "EncounterID": "FK → encounters.", "PatientID": "FK → patients.", "TestCode": "Short test code.",
        "TestName": "Test name.", "ResultValue": "Numeric result.", "Units": "Units.", "RefLow": "Reference range low.", "RefHigh": "Reference range high.",
        "AbnormalFlag": "N normal, L/H low/high, LL/HH critical.", "Priority": "STAT or Routine.", "OrderDateTime": "Ordered.",
        "CollectedDateTime": "Specimen collected.", "ResultedDateTime": "Result available."}),
    "medications": ("Pharmacy medication orders and doses dispensed.", {
        "MedOrderID": "Primary key.", "EncounterID": "FK → encounters.", "PatientID": "FK → patients.", "OrderDateTime": "Order time.",
        "MedicationName": "Generic drug name.", "DrugClass": "Therapeutic class.", "Dose": "Dose (text).", "Route": "IV, PO, SubQ, IM, Inhaled.",
        "Frequency": "Dosing frequency.", "DosesDispensed": "Number of doses dispensed.", "UnitCost": "Acquisition cost per dose ($).",
        "HighAlert": "Y for high-alert medications (insulin, opioids, anticoagulants, chemo…).", "OrderingProviderID": "FK → providers."}),
    "patient_satisfaction": ("Post-discharge patient experience surveys (HCAHPS-style, inpatient & observation).", {
        "SurveyID": "Primary key.", "EncounterID": "FK → encounters.", "FacilityID": "FK → facilities.", "DeptID": "FK → departments.",
        "DischargeDate": "Discharge date.", "SurveyReceivedDate": "Date returned.",
        "NurseCommunication": "1 Never, 2 Sometimes, 3 Usually, 4 Always.", "DoctorCommunication": "1–4 scale.", "StaffResponsiveness": "1–4 scale.",
        "Cleanliness": "1–4 scale.", "Quietness": "1–4 scale.", "MedicationExplained": "1–4 scale.", "DischargeInfoProvided": "Yes/No.",
        "OverallRating": "0 (worst) – 10 (best).", "WouldRecommend": "Definitely Yes, Probably Yes, Probably No, Definitely No.",
        "Comment": "Free-text comment (blank if none)."}),
    "daily_census": ("Midnight census, admissions and discharges per inpatient unit per day (unit operations data — covers all patients, not just the encounter extract).", {
        "CensusDate": "Date.", "FacilityID": "FK → facilities.", "DeptID": "FK → departments.", "StaffedBeds": "Staffed beds that day.",
        "Admissions": "Admissions + transfers in.", "Discharges": "Discharges + transfers out.", "MidnightCensus": "Patients in beds at midnight (can exceed staffed beds on surge days)."}),
    "employees": ("Nursing and clinical support staff at Bluestone Memorial Hospital (F01).", {
        "EmployeeID": "Primary key.", "FirstName": "First name.", "LastName": "Last name.", "JobTitle": "Role.", "DeptID": "FK → departments.",
        "FacilityID": "FK → facilities.", "EmploymentType": "Full-Time, Part-Time, Per Diem, Agency.", "FTE": "FTE (0 = per diem).",
        "HourlyRate": "Base hourly rate ($).", "HireDate": "Hire date.", "TermDate": "Termination date (blank if employed).",
        "Status": "Active, Terminated, Leave of Absence.", "ShiftPreference": "Day, Night, Rotating.", "Certifications": "Semicolon-separated certifications."}),
    "shifts": ("Scheduled and worked shifts, Q4 2025 (Oct 1 – Dec 31), Bluestone Memorial staff.", {
        "ShiftID": "Primary key.", "EmployeeID": "FK → employees.", "DeptID": "FK → departments.", "ShiftDate": "Date the shift starts.",
        "ShiftType": "Day 12h (07:00–19:30), Night 12h (19:00–07:30), Day 8h (07:00–15:30), Evening 8h (15:00–23:30).",
        "ScheduledStart": "Scheduled start.", "ScheduledEnd": "Scheduled end (night shifts end the next day).", "ScheduledHours": "Paid hours scheduled (span minus 30-min unpaid meal).",
        "ClockIn": "Actual clock-in (blank if not worked).", "ClockOut": "Actual clock-out.", "WorkedHours": "Paid hours worked = (ClockOut − ClockIn) − 0.5 h meal.",
        "HoursOverSchedule": "WorkedHours − ScheduledHours, floored at 0.", "ShiftStatus": "Worked, Called Off, No Show."}),
    "supply_inventory": ("Supply stock by storage location (snapshot 12/31/2025). One row per SKU per location.", {
        "StockID": "Primary key (SKU + location).", "SKU": "Item number (repeats across locations).", "ItemDescription": "Description.",
        "Category": "Supply category.", "Vendor": "Supplier.", "UnitOfMeasure": "Each, Box, Case, Pack…", "UnitCost": "Cost per unit of measure ($).",
        "LocationDeptID": "FK → departments (storeroom location).", "QtyOnHand": "Units on hand.", "ParLevel": "Target stock level.",
        "ReorderPoint": "Reorder when QtyOnHand ≤ this.", "ReorderQty": "Standard reorder quantity.", "LeadTimeDays": "Vendor lead time.",
        "LastCountDate": "Last physical count.", "ExpirationDate": "Earliest lot expiration (perishables only).", "LotNumber": "Lot number (perishables only)."}),
    "budget": ("Monthly budget vs. actual by department and category, 2024–2025 (fiscal year = calendar year).", {
        "MonthStart": "First day of month.", "FiscalYear": "Year.", "FiscalMonth": "1–12.", "FacilityID": "FK → facilities.", "DeptID": "FK → departments.",
        "CostCenter": "Cost center.", "LineType": "Revenue or Expense.", "Category": "Revenue or expense category.",
        "BudgetAmount": "Budget ($).", "ActualAmount": "Actual ($)."}),
}


def main():
    OUT.mkdir(exist_ok=True)
    providers = gen_providers()
    patients = gen_patients(providers)
    encounters, ed_rows = gen_encounters(patients, providers)
    pats_by_id = {p["PatientID"]: p for p in patients}
    claims = gen_claims(encounters)
    labs = gen_labs(encounters, pats_by_id)
    meds = gen_meds(encounters)
    ed_by_enc = {}
    for row in ed_rows:
        ed_by_enc[row["EncounterID"]] = row
    surveys = gen_surveys(encounters, ed_by_enc)
    census = gen_census()
    employees = gen_employees()
    shifts = gen_shifts(employees)
    inventory = gen_inventory()
    budget = gen_budget()
    messy, truth = gen_messy_registrations(patients)

    counts = {}

    def emit(name, header, rows):
        write_csv(OUT / f"{name}.csv", header, rows)
        counts[name] = len(rows)

    emit("facilities", ["FacilityID", "FacilityName", "City", "State", "FacilityType", "LicensedBeds", "OpenedYear"], [list(f) for f in C.FACILITIES])
    emit("departments", ["DeptID", "DeptName", "FacilityID", "ServiceLine", "UnitType", "StaffedBeds", "CostCenter"], [list(d) for d in C.DEPARTMENTS])
    emit("payers", ["PayerID", "PayerName", "PayerType", "AvgAllowedPctOfCharges", "AvgDaysToPay", "TimelyFilingDays"], [list(p[:6]) for p in C.PAYERS])
    emit("diagnoses", ["DxCode", "DxDescription", "DxCategory", "ChronicCondition", "ExpectedLOS"],
         [[k, v[0], v[1], v[2], v[10] if v[3][0] > 0 else None] for k, v in sorted(C.DX.items())])
    emit("providers", ["ProviderID", "FirstName", "LastName", "Credential", "Specialty", "PrimaryDeptID", "FacilityID", "HireDate", "FTE"],
         [[p["ProviderID"], p["FirstName"], p["LastName"], p["Credential"], p["Specialty"], p["PrimaryDeptID"], p["FacilityID"], fmt_d(p["HireDate"]), p["FTE"]] for p in providers])
    emit("patients", ["PatientID", "MRN", "FirstName", "LastName", "Sex", "DOB", "City", "State", "ZIP", "Phone", "Email", "PreferredLanguage",
                      "PrimaryPayerID", "PCPProviderID", "HeightIn", "WeightLb", "ChronicConditions", "RegistrationDate", "DeceasedDate"],
         [[p["PatientID"], p["MRN"], p["FirstName"], p["LastName"], p["Sex"], fmt_d(p["DOB"]), p["City"], p["State"], p["ZIP"], p["Phone"], p["Email"],
           p["PreferredLanguage"], p["PrimaryPayerID"], p["PCPProviderID"], p["HeightIn"], p["WeightLb"], ";".join(p["Chronic"]),
           fmt_d(p["RegistrationDate"]), fmt_d(p["DeceasedDate"])] for p in patients])
    emit("encounters", ["EncounterID", "PatientID", "EncounterType", "FacilityID", "DeptID", "AttendingProviderID", "AdmitSource", "AdmitDateTime",
                        "DischargeDateTime", "PrimaryDxCode", "DischargeDisposition", "PayerID", "TotalCharges", "Readmit30"],
         [[e["EncounterID"], e["PatientID"], e["EncounterType"], e["FacilityID"], e["DeptID"], e["AttendingProviderID"], e["AdmitSource"],
           fmt_dt(e["AdmitDateTime"]), fmt_dt(e["DischargeDateTime"]), e["PrimaryDxCode"], e["DischargeDisposition"], e["PayerID"],
           money(e["TotalCharges"]), e["Readmit30"]] for e in encounters])
    ed_cols = ["EDVisitID", "EncounterID", "PatientID", "FacilityID", "ArrivalDateTime", "TriageDateTime", "ProviderSeenDateTime", "DispositionDateTime",
               "DepartureDateTime", "ArrivalMode", "ESILevel", "ChiefComplaint", "EDDisposition", "TempF", "HeartRate", "RespRate", "SystolicBP",
               "DiastolicBP", "SpO2", "PainScore"]
    ed_rows.sort(key=lambda z: z["ArrivalDateTime"])
    for i, row in enumerate(ed_rows):
        row["EDVisitID"] = f"ED{200001 + i}"
    emit("ed_visits", ed_cols, [[fmt_dt(row[c]) if isinstance(row[c], datetime) else row[c] for c in ed_cols] for row in ed_rows])
    cl_cols = ["ClaimID", "EncounterID", "PatientID", "PayerID", "ServiceDate", "SubmitDate", "BilledAmount", "AllowedAmount", "PatientResponsibility",
               "PaidAmount", "ClaimStatus", "DenialReason", "PaidDate"]
    emit("claims", cl_cols, [[fmt_d(c[k]) if isinstance(c[k], date) else money(c[k]) if isinstance(c[k], float) else c[k] for k in cl_cols] for c in claims])
    lab_cols = ["LabResultID", "EncounterID", "PatientID", "TestCode", "TestName", "ResultValue", "Units", "RefLow", "RefHigh", "AbnormalFlag", "Priority",
                "OrderDateTime", "CollectedDateTime", "ResultedDateTime"]

    def num(v):
        return int(v) if isinstance(v, float) and v.is_integer() else v
    emit("lab_results", lab_cols, [[fmt_dt(x[c]) if isinstance(x[c], datetime) else num(x[c]) for c in lab_cols] for x in labs])
    med_cols = ["MedOrderID", "EncounterID", "PatientID", "OrderDateTime", "MedicationName", "DrugClass", "Dose", "Route", "Frequency", "DosesDispensed",
                "UnitCost", "HighAlert", "OrderingProviderID"]
    emit("medications", med_cols, [[fmt_dt(x[c]) if isinstance(x[c], datetime) else money(x[c]) if c == "UnitCost" else x[c] for c in med_cols] for x in meds])
    sv_cols = ["SurveyID", "EncounterID", "FacilityID", "DeptID", "DischargeDate", "SurveyReceivedDate", "NurseCommunication", "DoctorCommunication",
               "StaffResponsiveness", "Cleanliness", "Quietness", "MedicationExplained", "DischargeInfoProvided", "OverallRating", "WouldRecommend", "Comment"]
    emit("patient_satisfaction", sv_cols, [[fmt_d(x[c]) if isinstance(x[c], date) else x[c] for c in sv_cols] for x in surveys])
    cen_cols = ["CensusDate", "FacilityID", "DeptID", "StaffedBeds", "Admissions", "Discharges", "MidnightCensus"]
    emit("daily_census", cen_cols, [[fmt_d(x[c]) if isinstance(x[c], date) else x[c] for c in cen_cols] for x in census])
    emp_cols = ["EmployeeID", "FirstName", "LastName", "JobTitle", "DeptID", "FacilityID", "EmploymentType", "FTE", "HourlyRate", "HireDate", "TermDate",
                "Status", "ShiftPreference", "Certifications"]
    emit("employees", emp_cols, [[fmt_d(x[c]) if isinstance(x[c], date) else money(x[c]) if c == "HourlyRate" else x[c] for c in emp_cols] for x in employees])
    sh_cols = ["ShiftID", "EmployeeID", "DeptID", "ShiftDate", "ShiftType", "ScheduledStart", "ScheduledEnd", "ScheduledHours", "ClockIn", "ClockOut",
               "WorkedHours", "HoursOverSchedule", "ShiftStatus"]
    emit("shifts", sh_cols, [[fmt_dt(x[c]) if isinstance(x[c], datetime) else fmt_d(x[c]) if isinstance(x[c], date) else x[c] for c in sh_cols] for x in shifts])
    inv_cols = ["StockID", "SKU", "ItemDescription", "Category", "Vendor", "UnitOfMeasure", "UnitCost", "LocationDeptID", "QtyOnHand", "ParLevel",
                "ReorderPoint", "ReorderQty", "LeadTimeDays", "LastCountDate", "ExpirationDate", "LotNumber"]
    emit("supply_inventory", inv_cols, [[fmt_d(x[c]) if isinstance(x[c], date) else money(x[c]) if c == "UnitCost" else x[c] for c in inv_cols] for x in inventory])
    bud_cols = ["MonthStart", "FiscalYear", "FiscalMonth", "FacilityID", "DeptID", "CostCenter", "LineType", "Category", "BudgetAmount", "ActualAmount"]
    emit("budget", bud_cols, [[fmt_d(x[c]) if isinstance(x[c], date) else x[c] for c in bud_cols] for x in budget])
    write_csv(OUT / "messy" / "patient_registrations_raw.csv",
              ["RecordID", "PatientName", "DOB", "Sex", "Phone", "Email", "CityStateZip", "Insurance", "MRN", "RegisteredOn"],
              [[m[c] for c in ["RecordID", "PatientName", "DOB", "Sex", "Phone", "Email", "CityStateZip", "Insurance", "MRN", "RegisteredOn"]] for m in messy])
    counts["messy/patient_registrations_raw"] = len(messy)
    write_csv(TRUTH / "patient_registrations_truth.csv", ["RecordID", "PatientID"], [[t["RecordID"], t["PatientID"]] for t in truth])
    write_readme(counts)
    for k, v in counts.items():
        print(f"{k:32s} {v:>7,d} rows")


def write_readme(counts):
    lines = [
        "# Bluestone Health System — Course Datasets",
        "",
        "> **Everything here is synthetic.** Bluestone Health System, its hospitals, towns, staff, and patients are fictional.",
        "> Names were randomly combined from common-name lists; phone numbers use the fictional 555 area code; emails use the",
        "> reserved `example.com` domain. ICD-10-CM diagnosis codes are real public codes used only for realism. Nothing here is",
        "> clinical guidance.",
        "",
        "These CSV files are generated by [`tools/generate_data.py`](../tools/generate_data.py) with a fixed random seed, so they are",
        "identical every time they are regenerated. Each lesson's workbook embeds the slice of this data it needs, so you normally",
        "work from the lesson's `.xlsx` file. The full CSVs are here for exploration, Power Query practice, and your own projects.",
        "",
        "**Coverage:** January 1, 2024 – December 31, 2025. **\"As of\" date for age, aging, and status calculations: December 31, 2025.**",
        "",
        "## The health system",
        "",
        "| FacilityID | Facility | Type | Licensed beds |",
        "|---|---|---|---|",
    ]
    for f in C.FACILITIES:
        lines.append(f"| {f[0]} | {f[1]} | {f[4]} | {f[5]} |")
    lines += ["", "## How the tables relate", "", "```mermaid", "erDiagram",
              "    facilities ||--o{ departments : has",
              "    departments ||--o{ providers : employs",
              "    patients ||--o{ encounters : has",
              "    encounters }o--|| diagnoses : \"primary dx\"",
              "    encounters }o--|| providers : attending",
              "    encounters }o--|| departments : \"unit/clinic\"",
              "    encounters }o--|| payers : billed",
              "    encounters ||--o| ed_visits : \"ED timeline\"",
              "    encounters ||--|| claims : billed",
              "    encounters ||--o{ lab_results : has",
              "    encounters ||--o{ medications : has",
              "    encounters ||--o| patient_satisfaction : surveyed",
              "    departments ||--o{ daily_census : reports",
              "    departments ||--o{ employees : staffs",
              "    employees ||--o{ shifts : works",
              "    departments ||--o{ supply_inventory : stores",
              "    departments ||--o{ budget : plans",
              "```", "", "## Files", "", "| File | Rows | Description |", "|---|---:|---|"]
    for name, (desc, _) in DICTIONARY.items():
        lines.append(f"| [`{name}.csv`]({name}.csv) | {counts[name]:,} | {desc} |")
    lines.append(f"| [`messy/patient_registrations_raw.csv`](messy/patient_registrations_raw.csv) | {counts['messy/patient_registrations_raw']:,} | "
                 "A deliberately messy registration export (inconsistent names, dates stored as text, duplicates) for the data-cleaning lessons. |")
    lines += ["", "## Data dictionary", ""]
    for name, (desc, cols) in DICTIONARY.items():
        lines += [f"### `{name}.csv`", "", desc, "", "| Column | Description |", "|---|---|"]
        for c, d in cols.items():
            lines.append(f"| `{c}` | {d} |")
        lines.append("")
    lines += ["## Regenerating", "", "```bash", "python tools/generate_data.py", "```", ""]
    (OUT / "README.md").write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    main()
