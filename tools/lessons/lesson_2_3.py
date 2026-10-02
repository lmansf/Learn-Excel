"""Lesson 2.3 · Dates & Times.

Data story: Ashby Falls Community Hospital (F02). Every inpatient stay admitted in 2025 (with the patient's
DOB and, for ED admissions, the ED arrival time pre-joined), the insurance claim for each of those stays, the
hospital's ED visits in Q4 2025, plus one pay week of worked shifts on 4 West at Bluestone Memorial (the only
facility with time-clock data). A Holidays table and a Settings sheet (ReportDate = 12/31/2025) support the
business-day and "as of" tasks.

Every answer below is computed in Python with small models of the Excel date functions (DATEDIF "Y",
YEARFRAC basis 1, EDATE, EOMONTH, NETWORKDAYS, WORKDAY), so the key never relies on hand-typed numbers.
"""
from __future__ import annotations

import calendar
from collections import defaultdict
from datetime import date, datetime, time, timedelta

from xlcourse import Lesson, Task, data
from xlcourse.data import excel_serial

CODE = "2.3"
REPORT_DATE = date(2025, 12, 31)
FACILITY = "F02"
SHIFT_DEPT = "D110"
PAY_WEEK = (date(2025, 12, 14), date(2025, 12, 20))   # Sunday through Saturday
MEAL = timedelta(minutes=30)
APPEAL_BUSINESS_DAYS = 10
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


# ---------------------------------------------------------------------------
# Python models of the Excel date functions
# ---------------------------------------------------------------------------
def is_leap(y: int) -> bool:
    return calendar.isleap(y)


def datedif_y(start: date, end: date) -> int:
    """DATEDIF(start, end, "Y"): completed years."""
    years = end.year - start.year
    if (end.month, end.day) < (start.month, start.day):
        years -= 1
    return years


def datedif_m(start: date, end: date) -> int:
    """DATEDIF(start, end, "M"): completed months."""
    months = (end.year - start.year) * 12 + end.month - start.month
    if end.day < start.day:
        months -= 1
    return months


def yearfrac_basis1(d1: date, d2: date) -> float:
    """YEARFRAC(d1, d2, 1), the actual/actual algorithm Excel (and LibreOffice) use."""
    if d1 > d2:
        d1, d2 = d2, d1
    days = (d2 - d1).days
    y1, m1, dd1 = d1.year, d1.month, d1.day
    y2, m2, dd2 = d2.year, d2.month, d2.day
    if y1 != y2 and (y2 != y1 + 1 or m1 < m2 or (m1 == m2 and dd1 < dd2)):
        # more than a year apart: average year length over every calendar year touched
        n_years = y2 - y1 + 1
        avg = sum(366 if is_leap(y) else 365 for y in range(y1, y2 + 1)) / n_years
        return days / avg
    if y1 == y2:
        return days / (366 if is_leap(y1) else 365)
    leap = (is_leap(y1) and (m1, dd1) <= (2, 29)) or (is_leap(y2) and (m2, dd2) >= (2, 29))
    return days / (366 if leap else 365)


def yearfrac_basis0(d1: date, d2: date) -> float:
    """YEARFRAC(d1, d2) with the default basis 0 (US 30/360), for the simple dates used in the explanation."""
    a, b = min(d1, d2), max(d1, d2)
    da, db = min(a.day, 30), b.day
    if db == 31 and da == 30:
        db = 30
    return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + (db - da)) / 360


def add_months(d: date, months: int) -> date:
    """EDATE(d, months): same day number, clipped to the end of a shorter month."""
    y, m = divmod(d.month - 1 + months, 12)
    y, m = d.year + y, m + 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def eomonth(d: date, months: int) -> date:
    """EOMONTH(d, months)."""
    first = add_months(date(d.year, d.month, 1), months)
    return date(first.year, first.month, calendar.monthrange(first.year, first.month)[1])


def nth_weekday(y: int, m: int, weekday: int, n: int) -> date:
    d = date(y, m, 1)
    d += timedelta(days=(weekday - d.weekday()) % 7)
    return d + timedelta(weeks=n - 1)


def last_weekday(y: int, m: int, weekday: int) -> date:
    d = date(y, m, calendar.monthrange(y, m)[1])
    return d - timedelta(days=(d.weekday() - weekday) % 7)


def observed(d: date) -> date:
    """Federal rule: a Saturday holiday is observed Friday, a Sunday holiday on Monday."""
    if d.weekday() == 5:
        return d - timedelta(days=1)
    if d.weekday() == 6:
        return d + timedelta(days=1)
    return d


def holiday_calendar() -> list[dict]:
    rows = []
    for y in (2025, 2026):
        thanksgiving = nth_weekday(y, 11, 3, 4)
        items = [
            (observed(date(y, 1, 1)), "New Year's Day"),
            (nth_weekday(y, 1, 0, 3), "Martin Luther King Jr. Day"),
            (last_weekday(y, 5, 0), "Memorial Day"),
            (observed(date(y, 6, 19)), "Juneteenth"),
            (observed(date(y, 7, 4)), "Independence Day"),
            (nth_weekday(y, 9, 0, 1), "Labor Day"),
            (thanksgiving, "Thanksgiving Day"),
            (thanksgiving + timedelta(days=1), "Day after Thanksgiving"),
            (observed(date(y, 12, 25)), "Christmas Day"),
        ]
        for d, name in items:
            if name in ("Independence Day", "Juneteenth", "New Year's Day", "Christmas Day") and d.day != {
                    "Independence Day": 4, "Juneteenth": 19, "New Year's Day": 1, "Christmas Day": 25}[name]:
                name += " (observed)"
            rows.append({"Date": d, "Holiday": name, "Day": d.strftime("%a")})
    return rows


def networkdays(start: date, end: date, holidays: set[date]) -> int:
    """NETWORKDAYS(start, end, holidays) for start <= end: Mon–Fri days, both ends counted, minus holidays."""
    n, d = 0, start
    while d <= end:
        if d.weekday() < 5 and d not in holidays:
            n += 1
        d += timedelta(days=1)
    return n


def workday(start: date, days: int, holidays: set[date]) -> date:
    """WORKDAY(start, days, holidays) for days > 0: the start date itself is never counted."""
    d = start
    while days > 0:
        d += timedelta(days=1)
        if d.weekday() < 5 and d not in holidays:
            days -= 1
    return d


def tod(dt: datetime) -> time:
    return dt.time()


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="02-formulas-functions", slug="03-date-time-functions",
        title="Dates & Times", level="Beginner → Intermediate", minutes=55,
        objectives=[
            "Understand date serial numbers and times as fractions of a day",
            "Build and take apart dates with DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH, and EDATE",
            "Calculate ages, lengths of stay, and turnaround times (DATEDIF, YEARFRAC, NETWORKDAYS, WORKDAY)",
            "Do time math for ED waits and overnight shifts (MOD, [h]:mm)",
        ],
    )

    # ------------------------------------------------------------------ data: inpatient stays (Ashby Falls, 2025)
    patients = data.index(data.load("patients"), "PatientID")
    depts = data.index(data.load("departments"), "DeptID")
    ed_by_enc = {}
    for v in data.load("ed_visits"):
        ed_by_enc.setdefault(v["EncounterID"], v)
    stays = [e for e in data.load("encounters")
             if e["EncounterType"] == "Inpatient" and e["FacilityID"] == FACILITY and e["AdmitDateTime"].year == 2025]
    stays.sort(key=lambda e: (e["AdmitDateTime"], e["EncounterID"]))
    for e in stays:
        e["DOB"] = patients[e["PatientID"]]["DOB"]
        e["Unit"] = depts[e["DeptID"]]["DeptName"]
        v = ed_by_enc.get(e["EncounterID"])
        e["EDArrivalDateTime"] = v["ArrivalDateTime"] if v else None
        assert (v is not None) == (e["AdmitSource"] == "Emergency Department")
    st = L.add_table_sheet(
        "Stays", stays, table="tblStays",
        columns=["EncounterID", "PatientID", "Unit", "AdmitSource", "DOB", "EDArrivalDateTime", "AdmitDateTime",
                 "DischargeDateTime"],
        extra_cols=["AgeAtAdmit", "LOSDays", "Midnights", "BenchMidnights", "AdmitDay"],
        formats={"LOSDays": "0.00", "AgeAtAdmit": "0", "Midnights": "0", "BenchMidnights": "0"},
        widths={"Unit": 24, "AdmitSource": 21, "DOB": 12, "EDArrivalDateTime": 19, "AdmitDateTime": 17,
                "DischargeDateTime": 19, "AgeAtAdmit": 12, "LOSDays": 10, "Midnights": 11, "BenchMidnights": 16,
                "AdmitDay": 12},
    )

    # ------------------------------------------------------------------ data: claims for those stays
    payers = data.index(data.load("payers"), "PayerID")
    stay_ids = {e["EncounterID"] for e in stays}
    claims = [c for c in data.load("claims") if c["EncounterID"] in stay_ids]
    claims.sort(key=lambda c: (c["ServiceDate"], c["ClaimID"]))
    for c in claims:
        c["PayerName"] = payers[c["PayerID"]]["PayerName"]
    cl = L.add_table_sheet(
        "Claims", claims, table="tblClaims",
        columns=["ClaimID", "EncounterID", "PayerName", "ClaimStatus", "ServiceDate", "SubmitDate", "PaidDate"],
        extra_cols=["DaysToPay"], formats={"DaysToPay": "0"},
        widths={"PayerName": 30, "ClaimStatus": 14, "ServiceDate": 12, "SubmitDate": 12, "PaidDate": 12,
                "DaysToPay": 11},
    )

    # ------------------------------------------------------------------ data: ED visits (Ashby Falls, Q4 2025)
    ed = [v for v in data.load("ed_visits")
          if v["FacilityID"] == FACILITY and v["ArrivalDateTime"].year == 2025 and v["ArrivalDateTime"].month >= 10]
    ed.sort(key=lambda v: (v["ArrivalDateTime"], v["EDVisitID"]))
    edd = L.add_table_sheet(
        "ED", ed, table="tblED",
        columns=["EDVisitID", "ESILevel", "ArrivalDateTime", "TriageDateTime", "ProviderSeenDateTime",
                 "DepartureDateTime", "EDDisposition"],
        extra_cols=["DoorToProviderMin"], formats={"DoorToProviderMin": "0"},
        widths={"ArrivalDateTime": 17, "TriageDateTime": 17, "ProviderSeenDateTime": 21, "DepartureDateTime": 19,
                "EDDisposition": 14, "DoorToProviderMin": 19},
    )

    # ------------------------------------------------------------------ data: one pay week of shifts (4 West)
    employees = data.index(data.load("employees"), "EmployeeID")
    shifts = [s for s in data.load("shifts")
              if s["DeptID"] == SHIFT_DEPT and PAY_WEEK[0] <= s["ShiftDate"] <= PAY_WEEK[1] and s["ShiftStatus"] == "Worked"]
    shifts.sort(key=lambda s: (s["ShiftDate"], s["ScheduledStart"], s["ShiftID"]))
    for s in shifts:
        s["JobTitle"] = employees[s["EmployeeID"]]["JobTitle"]
        s["InTime"] = tod(s["ClockIn"])      # the time-clock export keeps only the clock time
        s["OutTime"] = tod(s["ClockOut"])
    sh = L.add_table_sheet(
        "Shifts", shifts, table="tblShifts",
        columns=["ShiftID", "EmployeeID", "JobTitle", "ShiftDate", "ShiftType", ("InTime", "ClockIn"), ("OutTime", "ClockOut")],
        extra_cols=["PaidTime"], formats={"ClockIn": "hh:mm", "ClockOut": "hh:mm"},
        widths={"JobTitle": 28, "ShiftDate": 12, "ShiftType": 12, "ClockIn": 10, "ClockOut": 10, "PaidTime": 11},
    )

    # ------------------------------------------------------------------ data: holidays + settings
    hol_rows = holiday_calendar()
    hol = L.add_table_sheet("Holidays", hol_rows, table="tblHolidays", columns=["Date", "Holiday", "Day"],
                            widths={"Date": 12, "Holiday": 34, "Day": 7})
    holidays = {r["Date"] for r in hol_rows}
    settings_rows = [{"Setting": "ReportDate", "Value": REPORT_DATE,
                      "Notes": "The course's 'today' (the data is as of this date). Use this cell instead of TODAY() "
                               "so your answers never change."}]
    stg = L.add_table_sheet("Settings", settings_rows, columns=["Setting", "Value", "Notes"], as_table=False,
                            formats={"Value": "mm/dd/yyyy"}, widths={"Setting": 14, "Value": 12, "Notes": 90})

    n_stays, n_ed = len(stays), len(ed)
    L.data_note = (f"All {n_stays} inpatient stays admitted at Ashby Falls Community Hospital in 2025 (with each patient's "
                   f"date of birth and, for ED admissions, the ED arrival time) and the insurance claim for each stay; "
                   f"the hospital's {n_ed} emergency department visits in Q4 2025; one pay week "
                   f"({PAY_WEEK[0]:%B} {PAY_WEEK[0].day}–{PAY_WEEK[1].day}, 2025) of worked shifts on Medical-Surgical "
                   f"4 West at Bluestone Memorial Hospital; and the business office's 2025–2026 holiday calendar.")
    L.start_notes = [
        "Settings!B2 holds the report date (12/31/2025). Use it wherever a task says \"as of\" or \"today\". TODAY() "
        "would give a different answer every day you open the file.",
        "The two-midnight rule in the bonus is simplified for teaching date math. It is not billing or clinical guidance.",
    ]

    # ------------------------------------------------------------------ helpers
    sf, sl = st.first_row, st.last_row
    cf, cll = cl.first_row, cl.last_row
    ef, el = edd.first_row, edd.last_row
    hf, hl = sh.first_row, sh.last_row

    def sr(col):   # Stays range as a learner types it
        return f"Stays!{st.col(col)}{sf}:{st.col(col)}{sl}"

    def scol(col):  # first-data-row Stays cell, no sheet name (for fill formulas)
        return f"{st.col(col)}{sf}"

    def cr(col):
        return f"Claims!{cl.col(col)}{cf}:{cl.col(col)}{cll}"

    def ccol(col):
        return f"{cl.col(col)}{cf}"

    def er(col):
        return f"ED!{edd.col(col)}{ef}:{edd.col(col)}{el}"

    def ecol(col):
        return f"{edd.col(col)}{ef}"

    def shr(col):
        return f"Shifts!{sh.col(col)}{hf}:{sh.col(col)}{hl}"

    def shcol(col):
        return f"{sh.col(col)}{hf}"

    report_cell = f"Settings!${stg.col('Value')}${stg.first_row}"
    hol_rng = f"Holidays!${hol.col('Date')}${hol.first_row}:${hol.col('Date')}${hol.last_row}"
    A, B = st.col("AdmitDateTime"), st.col("DOB")

    # ------------------------------------------------------------------ answers (computed in Python)
    # 1. serial number of the first admission
    s1 = stays[0]
    t1_serial = excel_serial(s1["AdmitDateTime"])

    # 2. 65th birthday (Medicare age) of the first patient who was 64 at admission
    i2 = next(i for i, e in enumerate(stays) if datedif_y(e["DOB"], e["AdmitDateTime"].date()) == 64
              and not (e["DOB"].month == 2 and e["DOB"].day == 29))
    s2, row2 = stays[i2], sf + i2
    bday65 = add_months(s2["DOB"], 65 * 12)
    assert bday65 == date(s2["DOB"].year + 65, s2["DOB"].month, s2["DOB"].day) and bday65 > s2["AdmitDateTime"].date()
    leap_days2 = (bday65 - s2["DOB"]).days - 65 * 365

    # 3. age column (DATEDIF "Y") -> average age
    ages = [datedif_y(e["DOB"], e["AdmitDateTime"].date()) for e in stays]
    avg_age = sum(ages) / n_stays
    ages_365 = [((e["AdmitDateTime"].date() - e["DOB"]).days) // 365 for e in stays]
    n_wrong_365 = sum(a != b for a, b in zip(ages, ages_365))
    assert n_wrong_365 > 0
    n_65 = sum(a >= 65 for a in ages)

    # 4. YEARFRAC basis 1 for the youngest patient
    i4 = min(range(n_stays), key=lambda i: stays[i]["AdmitDateTime"].date() - stays[i]["DOB"])
    s4, row4 = stays[i4], sf + i4
    yf1 = yearfrac_basis1(s4["DOB"], s4["AdmitDateTime"].date())
    yf0 = yearfrac_basis0(s4["DOB"], s4["AdmitDateTime"].date())
    months4 = datedif_m(s4["DOB"], s4["AdmitDateTime"].date())
    days4 = (s4["AdmitDateTime"].date() - s4["DOB"]).days
    assert round(yf1, 2) != round(yf0, 2), "the basis must change the 2-decimal answer"

    # 5. weekend admissions
    weekend_n = sum(e["AdmitDateTime"].weekday() >= 5 for e in stays)

    # 6. April admissions, and the <= DATE(2025,4,30) trap
    apr_n = sum(e["AdmitDateTime"].year == 2025 and e["AdmitDateTime"].month == 4 for e in stays)
    apr_last_day = sum(e["AdmitDateTime"].date() == date(2025, 4, 30) for e in stays)
    assert apr_last_day > 0
    apr30_eg = next(e["AdmitDateTime"] for e in stays
                    if e["AdmitDateTime"].date() == date(2025, 4, 30) and e["AdmitDateTime"].hour >= 12)

    # 7. LOS column (decimal days) -> average
    los = [(e["DischargeDateTime"] - e["AdmitDateTime"]).total_seconds() / 86400 for e in stays]
    avg_los = sum(los) / n_stays

    # 8. EOMONTH: internal billing deadline (end of the month after the service month)
    i8 = next(i for i, c in enumerate(claims) if c["ServiceDate"] == date(2025, 1, 30))
    c8, row8 = claims[i8], cf + i8
    deadline8 = eomonth(c8["ServiceDate"], 1)
    assert deadline8 == date(2025, 2, 28)

    # 9. days to pay (paid claims only) -> average
    dtp = [(c["PaidDate"] - c["SubmitDate"]).days for c in claims if c["PaidDate"]]
    avg_dtp = sum(dtp) / len(dtp)
    n_unpaid = sum(1 for c in claims if c["PaidDate"] is None)
    unpaid_serial = int(excel_serial(next(c for c in claims if c["PaidDate"] is None)["SubmitDate"]))

    # 10. NETWORKDAYS: a pending claim's business days waiting as of ReportDate
    i10 = next(i for i, c in enumerate(claims) if c["ClaimStatus"] == "Pending"
               and date(2025, 11, 1) <= c["SubmitDate"] <= date(2025, 11, 25))
    c10, row10 = claims[i10], cf + i10
    nwd10 = networkdays(c10["SubmitDate"], REPORT_DATE, holidays)
    nwd10_nohol = networkdays(c10["SubmitDate"], REPORT_DATE, set())
    assert nwd10_nohol - nwd10 == 3 and {date(2025, 11, 27), date(2025, 11, 28), date(2025, 12, 25)} <= holidays
    assert c10["SubmitDate"] < date(2025, 11, 27)

    # 11. WORKDAY: underpayment appeal deadline, 10 business days after the partial payment posted
    # (an insurer's partial payment with a stated reason, whose 10-business-day window crosses Thanksgiving)
    i11 = next(i for i, c in enumerate(claims) if c["ClaimStatus"] == "Partially Paid" and c["PaidDate"]
               and c["PayerID"] != "PY07" and c["DenialReason"]
               and date(2025, 11, 10) <= c["PaidDate"] <= date(2025, 11, 26)
               and workday(c["PaidDate"], APPEAL_BUSINESS_DAYS, holidays) != workday(c["PaidDate"], APPEAL_BUSINESS_DAYS, set()))
    c11, row11 = claims[i11], cf + i11
    wd11 = workday(c11["PaidDate"], APPEAL_BUSINESS_DAYS, holidays)
    wd11_nohol = workday(c11["PaidDate"], APPEAL_BUSINESS_DAYS, set())
    skipped11 = [r["Holiday"] for r in hol_rows if c11["PaidDate"] < r["Date"] <= wd11 and r["Date"].weekday() < 5]

    # 12. door-to-provider minutes (blank provider time = left without being seen)
    d2p = [round((v["ProviderSeenDateTime"] - v["ArrivalDateTime"]).total_seconds() / 60) for v in ed
           if v["ProviderSeenDateTime"]]
    avg_d2p = sum(d2p) / len(d2p)
    n_lwbs = n_ed - len(d2p)
    late = next(v for v in ed if v["ProviderSeenDateTime"] and v["ProviderSeenDateTime"].date() != v["ArrivalDateTime"].date())
    late_min = round((late["ProviderSeenDateTime"] - late["ArrivalDateTime"]).total_seconds() / 60)

    # 13. paid time with MOD for overnight shifts -> total as a duration
    def paid(s):
        span = (datetime.combine(date.min, s["OutTime"]) - datetime.combine(date.min, s["InTime"]))
        if span < timedelta(0):
            span += timedelta(days=1)
        return span - MEAL
    paid_total = sum((paid(s) for s in shifts), timedelta(0))
    n_overnight = sum(1 for s in shifts if s["OutTime"] < s["InTime"])
    naive_total = sum(((datetime.combine(date.min, s["OutTime"]) - datetime.combine(date.min, s["InTime"])) - MEAL
                       for s in shifts), timedelta(0))
    night = next(s for s in shifts if s["OutTime"] < s["InTime"])
    night_raw = (datetime.combine(date.min, night["OutTime"]) - datetime.combine(date.min, night["InTime"])).total_seconds() / 86400
    night_span = night["ClockOut"] - night["ClockIn"]
    assert all(s["ClockOut"] - s["ClockIn"] == paid(s) + MEAL for s in shifts)

    # ------------------------------------------------------------------ bonus answers
    mids = [(e["DischargeDateTime"].date() - e["AdmitDateTime"].date()).days for e in stays]
    two_mid = sum(m >= 2 for m in mids)
    assert all(e["DischargeDateTime"] - e["AdmitDateTime"] != timedelta(days=2) for e in stays)
    short_two_mid = sum(1 for m, d in zip(mids, los) if m >= 2 and d < 2)
    bench = [(e["DischargeDateTime"].date() - (e["EDArrivalDateTime"] or e["AdmitDateTime"]).date()).days for e in stays]
    bench_n = sum(b >= 2 for b in bench)
    gained = sum(1 for m, b in zip(mids, bench) if b > m)
    tipped = bench_n - two_mid
    assert bench_n >= two_mid
    hours_by_day = defaultdict(list)
    for e, d in zip(stays, los):
        hours_by_day[WEEKDAYS[e["AdmitDateTime"].weekday()]].append(d * 24)
    avg_by_day = {k: sum(v) / len(v) for k, v in hours_by_day.items()}
    fri_hours = avg_by_day["Friday"]
    ranked = sorted(avg_by_day.items(), key=lambda kv: kv[1], reverse=True)
    top_day, top_hours = ranked[0]
    assert ranked[0][1] - ranked[1][1] > 1, "the longest-LOS weekday must be clear-cut"

    # ------------------------------------------------------------------ live (stand-alone) formulas for the key
    age_arr = (f"YEAR({sr('AdmitDateTime')})-YEAR({sr('DOB')})-((MONTH({sr('AdmitDateTime')})*100+DAY({sr('AdmitDateTime')}))"
               f"<(MONTH({sr('DOB')})*100+DAY({sr('DOB')})))")
    mid_arr = f"(INT({sr('DischargeDateTime')})-INT({sr('AdmitDateTime')}))"
    start_day_arr = (f"(INT({sr('AdmitDateTime')})*({sr('EDArrivalDateTime')}=\"\")"
                     f"+INT({sr('EDArrivalDateTime')})*({sr('EDArrivalDateTime')}<>\"\"))")
    bench_arr = f"(INT({sr('DischargeDateTime')})-{start_day_arr})"
    fri_arr = f"(WEEKDAY({sr('AdmitDateTime')},2)=5)"

    L.practice_intro = ("Tasks 1–7 use the Stays sheet, 8–11 the Claims sheet, 12 the ED sheet, and 13 the Shifts sheet. "
                        "Settings!B2 holds the report date (12/31/2025) and the Holidays sheet lists the business office's "
                        "holidays. Several tasks ask you to fill a yellow column on a data sheet: type the formula in the "
                        "first data row and the Table fills the rest (if it doesn't, double-click the fill handle). The gray "
                        "cell on this sheet then summarizes your column.")

    L.tasks = [
        # ---------------- serial numbers
        Task(f"Stays!{A}{sf} shows the first admission of 2025: {s1['AdmitDateTime']:%m/%d/%Y %H:%M}. What number does "
             f"Excel actually store in that cell? Point a formula at the cell (the answer cell is already formatted to show "
             f"4 decimal places).",
             answer=t1_serial, fmt="0.0000", tol=0.0001,
             solution=f"=Stays!{A}{sf}", hint="A date-time is one number: whole days since 1900, plus a fraction of a day",
             explanation=f"The whole part, {int(t1_serial):,}, is the **date serial number**: {s1['AdmitDateTime']:%B} {s1['AdmitDateTime'].day}, {s1['AdmitDateTime'].year} is "
                         f"day {int(t1_serial):,} counting from January 1, 1900. The decimal part, {t1_serial % 1:.4f}, is the "
                         f"time as a fraction of a 24-hour day: {s1['AdmitDateTime']:%H:%M} is "
                         f"{s1['AdmitDateTime'].hour * 60 + s1['AdmitDateTime'].minute:,} minutes ÷ 1,440 minutes per day. You can also "
                         "see the number by giving the cell the General format with Ctrl + Shift + ~ (Mac: Control + Shift + ~). Because dates and times are "
                         "numbers, you can add, subtract, and compare them."),
        # ---------------- EDATE / DATE
        Task(f"The patient in Stays row {row2} (encounter {s2['EncounterID']}) was 64 at admission. On what date does the "
             f"patient turn 65, the usual age of Medicare eligibility? Use the DOB in that row.",
             answer=bday65, fmt="mm/dd/yyyy",
             solution=f"=EDATE(Stays!{B}{row2},65*12)",
             hint="EDATE moves a date by whole months. How many months are in 65 years?",
             explanation="EDATE moves a date forward (or back) by whole months and keeps the day number, so 65 × 12 = 780 months "
                         f"after the DOB is the 65th birthday. `=DATE(YEAR(Stays!{B}{row2})+65,MONTH(Stays!{B}{row2}),DAY(Stays!{B}{row2}))` "
                         "gives the same date. It takes the date apart with YEAR, MONTH, and DAY, adds 65 to the year, and builds "
                         f"it again with DATE. Adding 65 × 365 days would land {leap_days2} days early, because it ignores the "
                         f"{leap_days2} leap days in between."),
        # ---------------- DATEDIF / YEARFRAC
        Task(f"Fill the yellow AgeAtAdmit column on the Stays sheet with each patient's age in completed years on the admit "
             f"date. Start in {st.cell('AgeAtAdmit', 0, sheet=False)}. The gray cell averages your column. "
             f"What was the average age at admission?",
             answer=avg_age, fmt="0.00", tol=0.0001, title="AgeAtAdmit column with DATEDIF (average age)",
             solution=f'=DATEDIF({scol("DOB")},{scol("AdmitDateTime")},"Y")',
             summary=f'=IF(COUNT({st.rng("AgeAtAdmit")})=0,"",AVERAGE({st.rng("AgeAtAdmit")}))',
             fill={"range": f"Stays!{st.col('AgeAtAdmit')}{sf}:{st.col('AgeAtAdmit')}{sl}",
                   "formula": f'=DATEDIF({scol("DOB")},{scol("AdmitDateTime")},"Y")'},
             live=f"=SUMPRODUCT({age_arr})/COUNT({sr('DOB')})",
             hint="DATEDIF(start_date, end_date, \"Y\") counts completed years. Type it in full, because Excel won't suggest it",
             explanation="DATEDIF counts whole birthdays passed, which is how age is stated on a chart. Shortcuts like "
                         f"`=INT(({scol('AdmitDateTime')}-{scol('DOB')})/365)` drift by a day for every leap year lived, and "
                         f"in this data they make {n_wrong_365} patients a year too old. Excel doesn't list DATEDIF in AutoComplete "
                         "or Insert Function, so type it in full. The start date must come first, or DATEDIF returns #NUM!. "
                         f"{n_65} of the {n_stays} patients were 65 or older at admission."),
        Task(f"The youngest patient (Stays row {row4}, encounter {s4['EncounterID']}) was an infant. Calculate the exact age in "
             f"years at admission with YEARFRAC, using basis 1 (actual/actual). Round to 2 decimal places.",
             answer=round(yf1, 2), fmt="0.00",
             solution=f"=ROUND(YEARFRAC(Stays!{B}{row4},Stays!{A}{row4},1),2)",
             hint="YEARFRAC(start_date, end_date, basis). Don't skip the third argument",
             explanation=f"YEARFRAC returns the fraction of a year between two dates: {yf1:.4f} here. Basis 1 counts the real days "
                         f"({days4} of them) and divides by the real length of the year. Without the basis argument YEARFRAC uses basis 0 "
                         f"(the 30/360 banking convention, where every month has 30 days) and returns {yf0:.4f}, which rounds to "
                         f"{yf0:.2f}. Use basis 1 for ages. For infants, clinicians usually state age in months instead: "
                         f"`=DATEDIF(Stays!{B}{row4},Stays!{A}{row4},\"M\")` returns {months4}."),
        # ---------------- WEEKDAY / DATE
        Task("How many stays were admitted on a weekend (Saturday or Sunday)? Use the AdmitDateTime column.",
             answer=weekend_n,
             solution=f"=SUMPRODUCT(--(WEEKDAY({sr('AdmitDateTime')},2)>5))",
             hint="WEEKDAY with return_type 2 numbers Monday as 1 and Sunday as 7",
             explanation="`WEEKDAY(date,2)` returns 1 for Monday through 7 for Sunday, so Saturday and Sunday are exactly the "
                         "values above 5. Comparing the whole column gives a list of TRUE/FALSE values, `--` turns them into 1s and 0s, "
                         "and SUMPRODUCT adds them (Lesson 2.1). With the default return_type 1 (Sunday = 1, Saturday = 7) you'd need "
                         f"two tests: `=SUMPRODUCT((WEEKDAY({sr('AdmitDateTime')})=1)+(WEEKDAY({sr('AdmitDateTime')})=7))`."),
        Task("How many stays were admitted in April 2025? Compare AdmitDateTime with dates you build with DATE.",
             answer=apr_n,
             solution=f"=SUMPRODUCT(({sr('AdmitDateTime')}>=DATE(2025,4,1))*({sr('AdmitDateTime')}<DATE(2025,5,1)))",
             hint="On or after April 1, and before May 1",
             explanation="DATE(2025,4,1) builds the serial number for April 1, and multiplying the two TRUE/FALSE lists keeps only "
                         "rows that pass both tests (AND logic). The upper bound is **before May 1**, not on or before April 30. "
                         f"AdmitDateTime includes a time, so {apr30_eg:%m/%d/%Y %H:%M} is *greater* than DATE(2025,4,30), which means midnight. "
                         f"`<=DATE(2025,4,30)` would miss the {apr_last_day} admissions on April 30 and return {apr_n - apr_last_day}. "
                         f"`=SUMPRODUCT((MONTH({sr('AdmitDateTime')})=4)*(YEAR({sr('AdmitDateTime')})=2025))` also works."),
        # ---------------- LOS
        Task(f"Fill the yellow LOSDays column with each stay's length of stay in days, including the fraction of a day "
             f"(DischargeDateTime minus AdmitDateTime). Start in {st.cell('LOSDays', 0, sheet=False)}. The gray cell averages "
             f"your column. What was the average length of stay?",
             answer=avg_los, fmt="0.00", title="LOSDays column (average length of stay)",
             solution=f"={scol('DischargeDateTime')}-{scol('AdmitDateTime')}",
             summary=f'=IF(COUNT({st.rng("LOSDays")})=0,"",AVERAGE({st.rng("LOSDays")}))',
             fill={"range": f"Stays!{st.col('LOSDays')}{sf}:{st.col('LOSDays')}{sl}",
                   "formula": f"={scol('DischargeDateTime')}-{scol('AdmitDateTime')}"},
             live=f"=(SUM({sr('DischargeDateTime')})-SUM({sr('AdmitDateTime')}))/COUNT({sr('AdmitDateTime')})",
             hint="Later date-time minus earlier date-time gives days",
             explanation="Subtracting two date-times gives the elapsed days, with the hours as a decimal: 2.50 is two and a half "
                         "days. In a cell with the General format, Excel may copy the date format from the cells you referenced and "
                         "show 4.5 days as 01/04/1900 12:00. Change the format to Number when that happens. Multiply by 24 to get hours. "
                         "The live formula in the key uses a shortcut: the sum of the differences equals the difference of the sums."),
        # ---------------- EOMONTH
        Task(f"Ashby Falls' billing standard says a claim must be submitted by the last day of the month after the month of "
             f"service. What is the deadline for claim {c8['ClaimID']} (Claims row {row8}, ServiceDate "
             f"{c8['ServiceDate']:%m/%d/%Y})?",
             answer=deadline8, fmt="mm/dd/yyyy",
             solution=f"=EOMONTH(Claims!{cl.col('ServiceDate')}{row8},1)",
             hint="EOMONTH(start_date, months) returns the last day of a month",
             explanation="`EOMONTH(date,0)` is the last day of the date's own month, and `EOMONTH(date,1)` is the last day of the "
                         "next month. It handles 28-, 29-, 30-, and 31-day months for you. Adding 30 days instead would give "
                         f"{c8['ServiceDate'] + timedelta(days=30):%m/%d/%Y}, which is in the wrong month. `=EOMONTH(date,-1)+1` gives the first day "
                         "of the date's month, another pattern you'll use often."),
        Task(f"Fill the yellow DaysToPay column on the Claims sheet with the calendar days from SubmitDate to PaidDate. "
             f"{n_unpaid} claims have no PaidDate yet, so make those rows return \"\" (empty text). Start in "
             f"{cl.cell('DaysToPay', 0, sheet=False)}. The gray cell averages your column. What is the average days to pay?",
             answer=avg_dtp, fmt="0.0", title="DaysToPay column (average days to pay)",
             solution=f'=IF({ccol("PaidDate")}="","",{ccol("PaidDate")}-{ccol("SubmitDate")})',
             summary=f'=IF(COUNT({cl.rng("DaysToPay")})=0,"",AVERAGE({cl.rng("DaysToPay")}))',
             fill={"range": f"Claims!{cl.col('DaysToPay')}{cf}:{cl.col('DaysToPay')}{cll}",
                   "formula": f'=IF({ccol("PaidDate")}="","",{ccol("PaidDate")}-{ccol("SubmitDate")})'},
             live=f'=SUMPRODUCT(({cr("PaidDate")}<>"")*({cr("PaidDate")}-{cr("SubmitDate")}))/COUNT({cr("PaidDate")})',
             hint="Test for a blank PaidDate with IF before you subtract",
             explanation=f"An empty PaidDate counts as 0 in arithmetic, so a bare `={ccol('PaidDate')}-{ccol('SubmitDate')}` "
                         f"returns minus the submit date's serial number for every unpaid claim (−{unpaid_serial:,} for the first one), which "
                         "wrecks the average. The IF returns empty text for those rows, and AVERAGE ignores text. "
                         f"`=IF({ccol('PaidDate')}=\"\",\"\",DAYS({ccol('PaidDate')},{ccol('SubmitDate')}))` gives the same result. "
                         "Notice that DAYS takes the **end** date first, and that it needs the same blank guard."),
        # ---------------- NETWORKDAYS / WORKDAY
        Task(f"Claim {c10['ClaimID']} (Claims row {row10}) is still Pending. As of the report date in Settings!B2, how many "
             f"business days has it been waiting? Count from its SubmitDate through the report date, both days included, "
             f"skipping weekends and the dates on the Holidays sheet.",
             answer=nwd10,
             solution=f"=NETWORKDAYS(Claims!{cl.col('SubmitDate')}{row10},{report_cell},{hol_rng})",
             hint="NETWORKDAYS(start_date, end_date, holidays)",
             explanation=f"NETWORKDAYS counts Monday-to-Friday dates from the start date through the end date, **including both**, "
                         f"and skips any date in the holidays range. Without the holiday list the answer would be {nwd10_nohol}. "
                         "The three missing days are Thanksgiving, the day after, and Christmas. The `$` signs lock the Settings and "
                         "Holidays references, so you can copy the formula down a column without them moving."),
        Task(f"Claim {c11['ClaimID']} (Claims row {row11}) was only partially paid by {c11['PayerName']}, and the payment "
             f"posted on its PaidDate. Bluestone Health's policy gives the appeals team {APPEAL_BUSINESS_DAYS} business days to send an "
             f"underpayment appeal, so the deadline is the {APPEAL_BUSINESS_DAYS}th business day after the PaidDate (the PaidDate "
             f"itself doesn't count). Weekends and the dates on the Holidays sheet aren't business days. What is the deadline?",
             answer=wd11, fmt="mm/dd/yyyy",
             solution=f"=WORKDAY(Claims!{cl.col('PaidDate')}{row11},{APPEAL_BUSINESS_DAYS},{hol_rng})",
             hint="WORKDAY(start_date, days, holidays) returns a date",
             explanation=f"WORKDAY steps forward the given number of business days, never counting the start date itself, and "
                         f"skips weekends and holidays. This window crosses {len(skipped11)} holidays ({', '.join(skipped11)}). Without the "
                         f"holiday list WORKDAY would return {wd11_nohol:%m/%d/%Y}, {(wd11 - wd11_nohol).days} days too early. WORKDAY "
                         "returns a serial number, so if you see a number like 46000 instead of a date, give the cell a Date format. "
                         "NETWORKDAYS measures a span you already have, and WORKDAY finds the date at the end of a span."),
        # ---------------- time math
        Task(f"Fill the yellow DoorToProviderMin column on the ED sheet with the minutes from ArrivalDateTime to "
             f"ProviderSeenDateTime. {n_lwbs} patients left without being seen and have no ProviderSeenDateTime, so make those "
             f"rows return \"\". Start in {edd.cell('DoorToProviderMin', 0, sheet=False)}. The gray cell averages your column. "
             f"What is the average door-to-provider time in minutes?",
             answer=avg_d2p, fmt="0.0", title="DoorToProviderMin column (average minutes)",
             solution=f'=IF({ecol("ProviderSeenDateTime")}="","",({ecol("ProviderSeenDateTime")}-{ecol("ArrivalDateTime")})*1440)',
             summary=f'=IF(COUNT({edd.rng("DoorToProviderMin")})=0,"",AVERAGE({edd.rng("DoorToProviderMin")}))',
             fill={"range": f"ED!{edd.col('DoorToProviderMin')}{ef}:{edd.col('DoorToProviderMin')}{el}",
                   "formula": f'=IF({ecol("ProviderSeenDateTime")}="","",({ecol("ProviderSeenDateTime")}-{ecol("ArrivalDateTime")})*1440)'},
             live=(f'=SUMPRODUCT(({er("ProviderSeenDateTime")}<>"")*({er("ProviderSeenDateTime")}-{er("ArrivalDateTime")}))'
                   f'*1440/COUNT({er("ProviderSeenDateTime")})'),
             hint="A difference of date-times is in days. A day has 24 × 60 = 1,440 minutes",
             explanation="Subtracting gives a fraction of a day (0.0347 is 50 minutes), and multiplying by 1,440 converts days to "
                         f"minutes. Because these columns hold the full date *and* time, visit {late['EDVisitID']} (arrived "
                         f"{late['ArrivalDateTime']:%H:%M}, seen at {late['ProviderSeenDateTime']:%H:%M} the next morning) still gives "
                         f"a positive {late_min} minutes, with no special handling. The IF keeps LWBS "
                         "visits out of the average instead of counting them as huge negative waits."),
        Task(f"The Shifts sheet's ClockIn and ClockOut columns hold clock times only, with no dates, and {n_overnight} shifts "
             f"end after midnight. Fill the yellow PaidTime column with each shift's paid time: clock-out minus clock-in, "
             f"corrected for shifts that cross midnight, minus a 30-minute unpaid meal break. Keep each result as a time "
             f"value (don't multiply by 24). Start in {sh.cell('PaidTime', 0, sheet=False)}. The gray cell totals your column "
             f"in [h]:mm format. What is the total paid time for the week?",
             answer=paid_total, fmt="[h]:mm", title="PaidTime column with MOD (total paid time for the week)",
             solution=f"=MOD({shcol('ClockOut')}-{shcol('ClockIn')},1)-TIME(0,30,0)",
             summary=f'=IF(COUNT({sh.rng("PaidTime")})=0,"",SUM({sh.rng("PaidTime")}))',
             fill={"range": f"Shifts!{sh.col('PaidTime')}{hf}:{sh.col('PaidTime')}{hl}",
                   "formula": f"=MOD({shcol('ClockOut')}-{shcol('ClockIn')},1)-TIME(0,30,0)"},
             live=f"=SUMPRODUCT(MOD({shr('ClockOut')}-{shr('ClockIn')},1))-COUNT({shr('ClockIn')})*TIME(0,30,0)",
             hint="MOD(…, 1) turns a negative time difference into the right positive one. Half an hour is a time value, not 0.5",
             explanation=f"Shift {night['ShiftID']} clocked in at {night['InTime']:%H:%M} and out at {night['OutTime']:%H:%M}. "
                         f"{night['OutTime']:%H:%M} − {night['InTime']:%H:%M} is negative (−{abs(night_raw):.4f} of a day). MOD(…, 1) adds "
                         "one whole day to a negative result and leaves positive results alone, so the shift becomes "
                         f"{int(night_span.total_seconds() // 3600)}:{int(night_span.total_seconds() % 3600 // 60):02d}. Subtract the meal "
                         "break as a time, `TIME(0,30,0)` or `\"0:30\"` or `30/1440`. Subtracting 0.5 would remove half a *day*. "
                         f"The total is {paid_total.total_seconds() / 3600:,.1f} hours, so it needs **[h]:mm**. Plain h:mm would wrap "
                         "past every 24 hours and show only the leftover hours. Without MOD the week totals "
                         f"{naive_total.total_seconds() / 3600:,.1f} hours, because every overnight shift comes out negative. "
                         "Format your column as [h]:mm (or h:mm) to read each shift."),
    ]

    # ------------------------------------------------------------------ bonus: two-midnight benchmark & weekday effect
    L.bonus_title = "Bonus: The two-midnight check"
    L.bonus_scenario = (
        "Ashby Falls' utilization review nurse is preparing for a Medicare audit. Under the CMS two-midnight benchmark, an "
        "inpatient admission is generally expected to span at least two midnights of hospital care. This bonus uses a "
        "simplified, educational version (not billing guidance): count the midnights between two date-times by comparing "
        "their dates, not the hours between them. The nurse also wants to know whether the day of the week a patient is "
        "admitted affects how long they stay. Fill the yellow Midnights, BenchMidnights, and AdmitDay columns on the Stays "
        "sheet as you go. B2, B4, and B5 also use your LOSDays column from task 7.")
    L.bonus = [
        Task(f"Fill the Midnights column with the number of midnights each stay crossed between AdmitDateTime and "
             f"DischargeDateTime. Start in {st.cell('Midnights', 0, sheet=False)}. The gray cell counts your rows with 2 or more. "
             f"How many stays crossed at least two midnights?",
             answer=two_mid, title="Midnights column (stays with 2 or more midnights)",
             solution=f"=INT({scol('DischargeDateTime')})-INT({scol('AdmitDateTime')})",
             summary=f'=IF(COUNT({st.rng("Midnights")})=0,"",COUNTIF({st.rng("Midnights")},">=2"))',
             fill={"range": f"Stays!{st.col('Midnights')}{sf}:{st.col('Midnights')}{sl}",
                   "formula": f"=INT({scol('DischargeDateTime')})-INT({scol('AdmitDateTime')})"},
             live=f"=SUMPRODUCT(--({mid_arr}>=2))",
             hint="INT strips the time from a date-time. Subtract the two dates",
             explanation="INT(date-time) drops the decimal part, leaving the date at midnight. The difference between two dates is "
                         "the number of midnights crossed: admitted Monday 23:00 and discharged Wednesday 01:00 is 2 midnights in "
                         "only 26 hours. Rounding LOSDays doesn't work: that stay's LOSDays is 1.08, which rounds to 1. "
                         f"{n_stays - two_mid} stays fall short of the benchmark."),
        Task("How many stays crossed two or more midnights even though they lasted LESS than 48 hours?",
             answer=short_two_mid,
             solution=f"=SUMPRODUCT(({sr('Midnights')}>=2)*({sr('LOSDays')}<2))",
             live=f"=SUMPRODUCT(({mid_arr}>=2)*(({sr('DischargeDateTime')}-{sr('AdmitDateTime')})<2))",
             hint="Two conditions on two of your columns: multiply the TRUE/FALSE lists. 48 hours is 2 days",
             explanation="LOSDays measures elapsed time, and 2 days is 48 hours. Midnights measures calendar dates. A patient "
                         "admitted late in the evening crosses a midnight within minutes, so a stay of 30 hours can span 2 midnights. "
                         "That's why a two-midnight review can't use \"LOS ≥ 2 days\" as a shortcut."),
        Task(f"CMS starts the benchmark clock when hospital care begins, which for ED admissions is the ED arrival, not the "
             f"inpatient admit order. Fill the BenchMidnights column: midnights from EDArrivalDateTime to DischargeDateTime, or "
             f"from AdmitDateTime when EDArrivalDateTime is blank. Start in {st.cell('BenchMidnights', 0, sheet=False)}. The gray "
             f"cell counts your rows with 2 or more. How many stays meet the benchmark when ED time counts?",
             answer=bench_n, title="BenchMidnights column (stays meeting the benchmark with ED time)",
             solution=f'=INT({scol("DischargeDateTime")})-INT(IF({scol("EDArrivalDateTime")}="",{scol("AdmitDateTime")},{scol("EDArrivalDateTime")}))',
             summary=f'=IF(COUNT({st.rng("BenchMidnights")})=0,"",COUNTIF({st.rng("BenchMidnights")},">=2"))',
             fill={"range": f"Stays!{st.col('BenchMidnights')}{sf}:{st.col('BenchMidnights')}{sl}",
                   "formula": f'=INT({scol("DischargeDateTime")})-INT(IF({scol("EDArrivalDateTime")}="",{scol("AdmitDateTime")},{scol("EDArrivalDateTime")}))'},
             live=f"=SUMPRODUCT(--({bench_arr}>=2))",
             hint="Use IF to pick the clock start for each row, then count midnights the same way as in B1",
             explanation=f"The IF picks the start of care for each row, and INT turns it into a date. {gained} ED patients arrived "
                         f"before midnight and were admitted after it, so each gains a midnight. Only {tipped} of them moves from "
                         "below the benchmark to meeting it. The rest already had 2 or more. "
                         f"`=INT({scol('DischargeDateTime')})-INT(MIN({scol('EDArrivalDateTime')}:{scol('AdmitDateTime')}))` is a "
                         "shorter trick that works because MIN ignores the blank cell, but the IF says what you mean."),
        Task("Fill the AdmitDay column with the weekday name of each AdmitDateTime (Monday, Tuesday, …). What was the average "
             "length of stay, in HOURS, for patients admitted on a Friday? Round to 1 decimal place.",
             answer=round(fri_hours, 1), fmt="0.0", tol=0.051,
             solution=f'=ROUND(AVERAGEIF({sr("AdmitDay")},"Friday",{sr("LOSDays")})*24,1)',
             live=f"=ROUND(SUMPRODUCT({fri_arr}*({sr('DischargeDateTime')}-{sr('AdmitDateTime')}))*24/SUMPRODUCT(--{fri_arr}),1)",
             hint="TEXT(date, \"dddd\") gives the weekday name. AVERAGEIF (a preview of Lesson 2.5) or AVERAGE(IF(…)) averages the Friday rows",
             explanation=f"Fill AdmitDay with `=TEXT({scol('AdmitDateTime')},\"dddd\")`, which returns full weekday names. AVERAGEIF averages the "
                         "LOSDays values on rows where AdmitDay is \"Friday\", and × 24 converts days to hours. In Microsoft 365 and Excel 2021, "
                         f"`=ROUND(AVERAGE(IF({sr('AdmitDay')}=\"Friday\",{sr('LOSDays')}))*24,1)` also works, as you saw in "
                         "Lesson 2.1. TEXT returns weekday names in your Office language, so a German Excel shows \"Freitag\"."),
        Task("Which weekday of admission has the LONGEST average length of stay? Type the weekday name.",
             answer=top_day, accept=[top_day[:3]], live=False,
             solution=("1. On the Bonus sheet, type the seven weekday names, Monday to Sunday, in G2:G8.\n"
                       f"2. In H2, enter `=AVERAGEIF(Stays!${st.col('AdmitDay')}${sf}:${st.col('AdmitDay')}${sl},G2,"
                       f"Stays!${st.col('LOSDays')}${sf}:${st.col('LOSDays')}${sl})*24` and fill it down "
                       "to H8.\n"
                       "3. Find the largest average by eye, or let Excel find it: `=INDEX(G2:G8,MATCH(MAX(H2:H8),H2:H8,0))` "
                       "(INDEX and MATCH are covered in Lesson 2.6)."),
             hint="Build a small seven-row table of averages, one per weekday",
             explanation=f"{top_day} admissions stay longest, at {top_hours:.1f} hours on average, against {ranked[1][1]:.1f} for "
                         f"{ranked[1][0]}, the runner-up, and {ranked[-1][1]:.1f} for {ranked[-1][0]}. A weekly pattern like this "
                         "often points to weekend gaps in services: a patient admitted late in the week may wait through the weekend "
                         "for a test, a procedure, or a discharge placement."),
    ]

    @L.customize
    def _selftest_admitday(wb, lesson, selftest):
        """B4 and B5 read the learner's AdmitDay column. The library's `fill` would stop the self-test from typing B4's
        own formula, so simulate the AdmitDay column here instead (self-test copy only)."""
        if not selftest:
            return
        ws = wb["Stays"]
        col = st.col("AdmitDay")
        for r in range(sf, sl + 1):
            lesson.set_formula(ws, f"{col}{r}", f'=TEXT({A}{r},"dddd")', dynamic=False)

    return L
