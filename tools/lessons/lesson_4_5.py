"""Lesson 4.5 · Statistics & Forecasting.

Data sheets
  Stays     a seeded random sample of 400 inpatient stays discharged in 2025 (200 Bluestone Memorial, 200 Cedar Ridge)
  Bins      upper limits for the LOS histogram
  EDWaits   every Q4 2025 ED visit at the three hospitals, with door-to-provider minutes (blank = left without being seen)
  EDDaily   daily ED arrivals per hospital, 01/01/2024–12/31/2025 (731 rows, zero-arrival days included)
  Monthly   24 months of system-wide ED visits, inpatient discharges, index stays and 30-day readmissions

Every statistic is computed here in plain Python with the same definitions Excel uses (sample standard deviation,
SKEW's adjusted Fisher-Pearson formula, Student's t distribution via the regularized incomplete beta function), so the
course needs no scipy. FORECAST.ETS is deliberately not graded: Excel's AAA exponential smoothing optimizes its own
parameters and LibreOffice's implementation differs, so the ETS walkthrough lives in the README guide only.
"""
from __future__ import annotations

import math
import random
import statistics as st
from collections import Counter, defaultdict
from datetime import date, timedelta

from xlcourse import Lesson, Task, data

CODE = "4.5"

SAMPLE_SEED = 2025          # random.Random(seed).sample(...) on EncounterID-sorted stays
SAMPLE_PER_HOSPITAL = 200
LOS_BINS = [2, 4, 6, 8, 10, 12, 14]
HIST_BIN = 6                # task 2 asks for the frequency of this bin
PRED_AGE = 80               # task 6
CI_FACILITY = "F02"         # task 9: Ashby Falls door-to-provider confidence interval
D2 = 1.128                  # bias-correction constant for moving ranges of 2 points
FAC_NAMES = {"F01": "Bluestone Memorial", "F02": "Ashby Falls", "F03": "Cedar Ridge"}
DAILY_COLS = {"F01": "Memorial", "F02": "AshbyFalls", "F03": "CedarRidge"}


# ---------------------------------------------------------------------------
# Excel-equivalent statistics
# ---------------------------------------------------------------------------
def _betacf(a: float, b: float, x: float) -> float:
    """Continued fraction for the incomplete beta function (modified Lentz)."""
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    d = 1.0 / (d if abs(d) > tiny else tiny)
    h = d
    for m in range(1, 400):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        d = 1.0 / (d if abs(d) > tiny else tiny)
        c = 1.0 + aa / c
        c = c if abs(c) > tiny else tiny
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        d = 1.0 / (d if abs(d) > tiny else tiny)
        c = 1.0 + aa / c
        c = c if abs(c) > tiny else tiny
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 1e-15:
            break
    return h


def _betai(a: float, b: float, x: float) -> float:
    """Regularized incomplete beta I_x(a, b)."""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    lbt = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x)
    if x < (a + 1) / (a + b + 2):
        return math.exp(lbt) * _betacf(a, b, x) / a
    return 1.0 - math.exp(lbt) * _betacf(b, a, 1 - x) / b


def t_dist_2t(t: float, df: float) -> float:
    """T.DIST.2T(|t|, df): two-tailed probability of Student's t."""
    return _betai(df / 2.0, 0.5, df / (df + t * t))


def t_inv_2t(p: float, df: float) -> float:
    """T.INV.2T(p, df) by bisection on the two-tailed probability."""
    lo, hi = 0.0, 1000.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if t_dist_2t(mid, df) > p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def norm_s_inv(p: float) -> float:
    """NORM.S.INV by bisection on the error function."""
    lo, hi = -40.0, 40.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if 0.5 * (1 + math.erf(mid / math.sqrt(2))) < p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def skew(x: list[float]) -> float:
    """SKEW: n / ((n-1)(n-2)) * sum(((x - mean) / s)^3) with the sample standard deviation."""
    n, m, s = len(x), st.mean(x), st.stdev(x)
    return n / ((n - 1) * (n - 2)) * sum(((v - m) / s) ** 3 for v in x)


def kurt(x: list[float]) -> float:
    """KURT: excess kurtosis with Excel's sample adjustment."""
    n, m, s = len(x), st.mean(x), st.stdev(x)
    s4 = sum(((v - m) / s) ** 4 for v in x)
    return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * s4 - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3))


def regress(x: list[float], y: list[float]) -> dict:
    """Simple least-squares line y = a + b x with the statistics LINEST(..., TRUE, TRUE) reports."""
    n = len(x)
    mx, my = st.mean(x), st.mean(y)
    sxx = sum((v - mx) ** 2 for v in x)
    syy = sum((v - my) ** 2 for v in y)
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    b = sxy / sxx
    a = my - b * mx
    ss_res = sum((yy - (a + b * xx)) ** 2 for xx, yy in zip(x, y))
    ss_reg = syy - ss_res
    df = n - 2
    steyx = math.sqrt(ss_res / df)
    se_b = steyx / math.sqrt(sxx)
    se_a = steyx * math.sqrt(1 / n + mx * mx / sxx)
    r = sxy / math.sqrt(sxx * syy)
    t_b = b / se_b
    t_a = a / se_a
    f = ss_reg / (ss_res / df)
    return dict(n=n, slope=b, intercept=a, r=r, r2=r * r, steyx=steyx, se_slope=se_b, se_intercept=se_a, df=df,
                t_slope=t_b, p_slope=t_dist_2t(abs(t_b), df), t_intercept=t_a, p_intercept=t_dist_2t(abs(t_a), df),
                ss_reg=ss_reg, ss_res=ss_res, f=f, sig_f=t_dist_2t(abs(t_b), df),
                lo_slope=b - t_inv_2t(0.05, df) * se_b, hi_slope=b + t_inv_2t(0.05, df) * se_b,
                lo_int=a - t_inv_2t(0.05, df) * se_a, hi_int=a + t_inv_2t(0.05, df) * se_a)


def welch(a: list[float], b: list[float]) -> dict:
    """T.TEST(a, b, 2, 3): Welch's unequal-variance t-test, fractional degrees of freedom."""
    n1, n2 = len(a), len(b)
    v1, v2 = st.variance(a), st.variance(b)
    se = math.sqrt(v1 / n1 + v2 / n2)
    t = (st.mean(a) - st.mean(b)) / se
    df = (v1 / n1 + v2 / n2) ** 2 / ((v1 / n1) ** 2 / (n1 - 1) + (v2 / n2) ** 2 / (n2 - 1))
    return dict(t=t, df=df, p2=t_dist_2t(abs(t), df), p1=t_dist_2t(abs(t), df) / 2,
                p2_round=t_dist_2t(abs(t), round(df)), m1=st.mean(a), m2=st.mean(b), v1=v1, v2=v2, n1=n1, n2=n2,
                tcrit2=t_inv_2t(0.05, round(df)))


def pooled(a: list[float], b: list[float]) -> dict:
    """T.TEST(a, b, 2, 2): pooled-variance t-test."""
    n1, n2 = len(a), len(b)
    sp2 = ((n1 - 1) * st.variance(a) + (n2 - 1) * st.variance(b)) / (n1 + n2 - 2)
    t = (st.mean(a) - st.mean(b)) / math.sqrt(sp2 * (1 / n1 + 1 / n2))
    df = n1 + n2 - 2
    return dict(t=t, df=df, p2=t_dist_2t(abs(t), df))


def ichart(x: list[float]) -> dict:
    """Individuals chart: centre = mean, sigma-hat = average moving range / d2."""
    mr = [abs(b - a) for a, b in zip(x, x[1:])]
    mrbar = st.mean(mr)
    m = st.mean(x)
    sig = mrbar / D2
    return dict(mean=m, mrbar=mrbar, sigma=sig, ucl=m + 3 * sig, lcl=m - 3 * sig, sd=st.stdev(x), n_mr=len(mr))


def _age(dob: date, on: date) -> int:
    return on.year - dob.year - ((on.month, on.day) < (dob.month, dob.day))


# ---------------------------------------------------------------------------
# Datasets (also imported by scratch scripts that check the README's worked examples)
# ---------------------------------------------------------------------------
def datasets() -> dict:
    enc = data.load("encounters")
    pts = data.index(data.load("patients"), "PatientID")
    dx = data.index(data.load("diagnoses"), "DxCode")

    # ---- Stays: seeded random sample of 2025 inpatient discharges, 200 per hospital
    ip25 = [e for e in enc if e["EncounterType"] == "Inpatient" and e["DischargeDateTime"].year == 2025]
    rng = random.Random(SAMPLE_SEED)
    stays = []
    for fid in ("F01", "F03"):
        pop = sorted((e for e in ip25 if e["FacilityID"] == fid), key=lambda e: e["EncounterID"])
        stays += rng.sample(pop, SAMPLE_PER_HOSPITAL)
    for e in stays:
        e["Facility"] = FAC_NAMES[e["FacilityID"]]
        e["DischargeDate"] = e["DischargeDateTime"].date()
        e["Age"] = _age(pts[e["PatientID"]]["DOB"], e["AdmitDateTime"].date())
        e["DxCategory"] = dx[e["PrimaryDxCode"]]["DxCategory"]
        # LOS days = midnights in hospital, minimum 1 (same definition as Lesson 4.2)
        e["LOSDays"] = max(1, (e["DischargeDateTime"].date() - e["AdmitDateTime"].date()).days)
    stays.sort(key=lambda e: (e["FacilityID"], e["DischargeDate"], e["EncounterID"]))

    # ---- EDWaits: every Q4 2025 ED arrival at the three hospitals
    ed = data.load("ed_visits")
    waits = [r for r in ed if r["ArrivalDateTime"].year == 2025 and r["ArrivalDateTime"].month >= 10]
    for r in waits:
        seen = r["ProviderSeenDateTime"]
        r["Facility"] = FAC_NAMES[r["FacilityID"]]
        r["DoorToProviderMin"] = round((seen - r["ArrivalDateTime"]).total_seconds() / 60) if seen else None
    waits.sort(key=lambda r: (r["FacilityID"], r["ArrivalDateTime"], r["EDVisitID"]))

    # ---- EDDaily: arrivals per hospital per calendar day (zeros included)
    counts = defaultdict(Counter)
    for r in ed:
        counts[r["FacilityID"]][r["ArrivalDateTime"].date()] += 1
    daily = []
    d = date(2024, 1, 1)
    while d <= date(2025, 12, 31):
        row = {"Date": d, "Weekday": d.strftime("%a")}
        for fid, col in DAILY_COLS.items():
            row[col] = counts[fid][d]
        row["AllEDs"] = sum(row[c] for c in DAILY_COLS.values())
        daily.append(row)
        d += timedelta(days=1)

    # ---- Monthly: system-wide volumes and readmissions by discharge month
    months = [date(2024 + i // 12, i % 12 + 1, 1) for i in range(24)]
    mrows = {m: {"MonthStart": m, "MonthNum": i + 1, "EDVisits": 0, "LWBS": 0, "InpatientDischarges": 0, "IndexStays": 0,
                 "Readmits30": 0} for i, m in enumerate(months)}
    for r in ed:
        m = mrows[date(r["ArrivalDateTime"].year, r["ArrivalDateTime"].month, 1)]
        m["EDVisits"] += 1
        m["LWBS"] += r["EDDisposition"] == "LWBS"
    for e in enc:
        if e["EncounterType"] != "Inpatient":
            continue
        m = mrows[date(e["DischargeDateTime"].year, e["DischargeDateTime"].month, 1)]
        m["InpatientDischarges"] += 1
        if e["DischargeDisposition"] == "Expired":
            continue                                  # a patient who died can't be readmitted
        m["IndexStays"] += 1
        m["Readmits30"] += e["Readmit30"] == "Y"
    monthly = [mrows[m] for m in months]
    return dict(stays=stays, waits=waits, daily=daily, monthly=monthly)


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="04-advanced-analysis", slug="05-statistics-forecasting",
        title="Statistics & Forecasting", level="Advanced", minutes=70,
        objectives=[
            "Summarize distributions with the Analysis ToolPak and histograms",
            "Measure relationships with correlation and linear regression",
            "Test differences with t-tests and quantify uncertainty with confidence intervals",
            "Forecast volumes and monitor processes with control charts",
        ],
        data_note="A random sample of 400 inpatient stays discharged in 2025 (200 each from Bluestone Memorial and Cedar "
                  "Ridge), every Q4 2025 emergency department visit at the three hospitals with door-to-provider minutes, "
                  "daily ED arrivals per hospital for 2024–2025, and 24 months of system-wide ED, LWBS, inpatient, and "
                  "30-day readmission counts.",
    )
    L.start_notes.append("Several tasks can be answered with the Analysis ToolPak (Data → Data Analysis). Turn it on first: "
                         "File → Options → Add-ins → Manage Excel Add-ins → Go (Mac: Tools → Excel Add-ins). The guide shows how.")
    L.start_notes.append("Control limits and statistical tests here are for learning process-monitoring methods, not clinical "
                         "or regulatory reporting.")

    D = datasets()
    stays, waits, daily, monthly = D["stays"], D["waits"], D["daily"], D["monthly"]

    # ------------------------------------------------------------------ sheets
    sty = L.add_table_sheet(
        "Stays", stays, table="tblStays",
        columns=["EncounterID", "Facility", "DischargeDate", "Age", "DxCategory", "LOSDays", "TotalCharges"],
        formats={"TotalCharges": "#,##0.00"}, widths={"Facility": 21, "DxCategory": 26, "TotalCharges": 14},
    )
    bins = L.add_table_sheet(
        "Bins", [{"LOSBin": b} for b in LOS_BINS], table="tblBins", columns=["LOSBin"], start_row=3,
        notes=["Bin range for the LOS histogram (Task 2)",
               "Bin upper limits. Use A4:A10 with FREQUENCY, or start the Histogram tool from this sheet (see the guide)."],
        widths={"LOSBin": 10},
    )
    edw = L.add_table_sheet(
        "EDWaits", waits, table="tblEDWaits",
        columns=["EDVisitID", "Facility", "ArrivalDateTime", "ESILevel", "DoorToProviderMin"],
        widths={"Facility": 21, "ArrivalDateTime": 17, "DoorToProviderMin": 19},
    )
    edd = L.add_table_sheet(
        "EDDaily", daily, table="tblEDDaily",
        columns=["Date", "Weekday", "Memorial", "AshbyFalls", "CedarRidge", "AllEDs"],
        widths={"Date": 12, "Weekday": 10},
    )
    mon = L.add_table_sheet(
        "Monthly", monthly, table="tblMonthly",
        columns=["MonthStart", "MonthNum", "EDVisits", "LWBS", "InpatientDischarges", "IndexStays", "Readmits30"],
        formats={"MonthStart": "mmm yyyy"}, widths={"MonthStart": 12, "InpatientDischarges": 20},
    )

    # ------------------------------------------------------------------ ranges as the learner types them
    sf, sl = sty.first_row, sty.last_row
    n_mem = sum(1 for e in stays if e["FacilityID"] == "F01")
    mem_last = sf + n_mem - 1
    ced_first = mem_last + 1

    def S(col, a=sf, b=sl):
        c = sty.col(col)
        return f"Stays!{c}{a}:{c}{b}"

    wf = edw.first_row + next(i for i, r in enumerate(waits) if r["FacilityID"] == CI_FACILITY)
    wl = edw.first_row + max(i for i, r in enumerate(waits) if r["FacilityID"] == CI_FACILITY)
    WCOL = edw.col("DoorToProviderMin")
    W = f"EDWaits!{WCOL}{wf}:{WCOL}{wl}"

    df_, dl_ = edd.first_row, edd.last_row
    row_of = {r["Date"]: df_ + i for i, r in enumerate(daily)}
    MC = edd.col("Memorial")
    r24a, r24b = row_of[date(2024, 1, 1)], row_of[date(2024, 12, 31)]
    r25a, r25b = row_of[date(2025, 1, 1)], row_of[date(2025, 12, 31)]

    mf, ml = mon.first_row, mon.last_row

    def M(col, a=mf, b=ml):
        c = mon.col(col)
        return f"Monthly!{c}{a}:{c}{b}"

    # ------------------------------------------------------------------ answers (computed in Python)
    los = [e["LOSDays"] for e in stays]
    age = [e["Age"] for e in stays]
    chg = [e["TotalCharges"] for e in stays]
    los_mem = [e["LOSDays"] for e in stays if e["FacilityID"] == "F01"]
    los_ced = [e["LOSDays"] for e in stays if e["FacilityID"] == "F03"]
    assert all(e["FacilityID"] == "F01" for e in stays[:n_mem]) and len(los_ced) == SAMPLE_PER_HOSPITAL

    # 1 · skewness (Descriptive Statistics)
    los_skew = skew(los)
    los_kurt = kurt(los)

    # 2 · histogram bin count: values above the previous limit, up to and including this one
    bi = LOS_BINS.index(HIST_BIN)
    lower = LOS_BINS[bi - 1]
    hist_n = sum(1 for v in los if lower < v <= HIST_BIN)
    hist_wrong = sum(1 for v in los if HIST_BIN <= v < HIST_BIN + 2)     # reading "6" as 6–7
    bin_more = sum(1 for v in los if v > LOS_BINS[-1])

    # 3–7 · correlation and regression
    fit_age = regress(age, los)
    fit_chg = regress(los, chg)
    r_age = fit_age["r"]
    pred80 = fit_age["intercept"] + fit_age["slope"] * PRED_AGE
    p_age = fit_age["p_slope"]
    assert 0.001 < p_age < 0.05, "the age slope should be significant but small (teaching point)"

    # 8 · Welch t-test
    tt = welch(los_mem, los_ced)
    tp = pooled(los_mem, los_ced)

    # 9 · CI half-width
    w = [r["DoorToProviderMin"] for r in waits if r["FacilityID"] == CI_FACILITY and r["DoorToProviderMin"] is not None]
    w_all = [r for r in waits if r["FacilityID"] == CI_FACILITY]
    w_n, w_mean, w_sd = len(w), st.mean(w), st.stdev(w)
    w_t = t_inv_2t(0.05, w_n - 1)
    ci_half = w_t * w_sd / math.sqrt(w_n)
    ci_norm = norm_s_inv(0.975) * w_sd / math.sqrt(w_n)
    # if size counted the blank (LWBS) rows too: CONFIDENCE.T(0.05, s, ROWS(...)) uses t with ROWS-1 df
    ci_wrong_n = t_inv_2t(0.05, len(w_all) - 1) * w_sd / math.sqrt(len(w_all))
    assert abs(ci_half - ci_norm) > 0.012, "CONFIDENCE.NORM must fail the check"

    # 10 · trailing 7-day moving average: the busiest week of 2025 at Memorial
    mem = [r["Memorial"] for r in daily]
    ma = {daily[i]["Date"]: st.mean(mem[i - 6:i + 1]) for i in range(6, len(daily))}
    ma_date = max((d for d in ma if d.year == 2025), key=lambda d: (ma[d], d))
    ma_val = ma[ma_date]
    ma_row = row_of[ma_date]
    i_ma = ma_row - df_
    centered = st.mean(mem[i_ma - 3:i_ma + 4])
    assert abs(centered - ma_val) > 0.01

    # 11 · linear trend forecast for month 25 (January 2026)
    mnum = [r["MonthNum"] for r in monthly]
    edv = [r["EDVisits"] for r in monthly]
    fit_ed = regress(mnum, edv)
    fc25 = fit_ed["intercept"] + fit_ed["slope"] * 25
    jan_actuals = [r["EDVisits"] for r in monthly if r["MonthStart"].month == 1]
    dec_actuals = [r["EDVisits"] for r in monthly if r["MonthStart"].month == 12]
    assert min(jan_actuals + dec_actuals) > fc25

    # 12–13 · I-chart, 2024 baseline, 2025 monitored
    base = ichart(mem[r24a - df_:r24b - df_ + 1])
    mon25 = mem[r25a - df_:r25b - df_ + 1]
    above = [(daily[r25a - df_ + i]["Date"], v) for i, v in enumerate(mon25) if v > base["ucl"]]
    assert not any(v < base["lcl"] for v in mon25), "Task 13's explanation says no 2025 day fell below the LCL"
    assert all(abs(v - base["ucl"]) > 0.1 for v in mon25), "no 2025 day may sit on the UCL"
    sd_ucl = base["mean"] + 3 * base["sd"]
    above_sd = sum(1 for v in mon25 if v > sd_ucl)
    assert above_sd != len(above)
    m25 = st.mean(mon25)

    # ------------------------------------------------------------------ practice
    T = "TRUE"
    LOSr, AGEr, CHGr = S("LOSDays"), S("Age"), S("TotalCharges")
    BINr = f"Bins!{bins.col('LOSBin')}{bins.first_row}:{bins.col('LOSBin')}{bins.last_row}"
    linest = f"LINEST({LOSr},{AGEr},{T},{T})"
    linest_chg = f"LINEST({CHGr},{LOSr},{T},{T})"
    t_chg = t_inv_2t(0.05, fit_chg["df"])
    ucl_f = (f"AVERAGE(EDDaily!{MC}{r24a}:{MC}{r24b})+3*SUMPRODUCT(ABS(EDDaily!{MC}{r24a + 1}:{MC}{r24b}"
             f"-EDDaily!{MC}{r24a}:{MC}{r24b - 1}))/COUNT(EDDaily!{MC}{r24a + 1}:{MC}{r24b})/{D2}")

    L.practice_intro = (
        f"Stays (rows {sf}–{sl}) holds the 400 sampled inpatient stays: Bluestone Memorial in rows {sf}–{mem_last} and "
        f"Cedar Ridge in rows {ced_first}–{sl}. Bins (A{bins.first_row}:A{bins.last_row}) holds the LOS histogram bins. "
        f"EDWaits holds Q4 2025 ED visits sorted by hospital. EDDaily holds one row per day from 01/01/2024 (row {df_}) to "
        f"12/31/2025 (row {dl_}), with 2025 starting in row {r25a}. Monthly holds Jan 2024 (row {mf}) to Dec 2025 (row {ml}). "
        "Answer with formulas where you can. Where a task says the ToolPak, a worksheet formula gives the same number. "
        "Each yellow cell holds one number, so wrap FREQUENCY or LINEST in INDEX there. Otherwise they spill into the "
        "cells below."
    )

    # practice row of each task (title row, two intro rows, score row, header row, then tasks)
    def prow(k):
        return 5 + k

    L.tasks = [
        Task("Run Data Analysis → Descriptive Statistics on LOSDays for all 400 stays (tick Summary statistics). What "
             "Skewness does the output report? Enter it to 2 decimal places, or use the worksheet function that "
             "calculates it.",
             answer=los_skew, fmt="0.00",
             solution=f"=SKEW({LOSr})",
             hint="The ToolPak's Skewness row is the SKEW function",
             explanation=f"Skewness measures how lopsided a distribution is. Zero means symmetric, and a positive value means a "
                         f"long tail to the right. LOS has a skewness of {los_skew:.2f}, because most stays are short "
                         f"(median {st.median(los):g} days) while a few run to {max(los)} days. The ToolPak's Descriptive "
                         "Statistics output is a block of typed-in numbers, not formulas, so it won't update if the data "
                         f"changes. `SKEW` will. The output's Kurtosis row ({los_kurt:.2f}) is `KURT`. A positive value means "
                         "heavier tails than a bell curve."),
        Task(f"Build a histogram of LOSDays (all 400 stays) using the bins on the Bins sheet, with the Histogram tool or "
             f"FREQUENCY. How many stays fall in the bin labeled {HIST_BIN}? Enter that one count.",
             answer=hist_n,
             solution=f"=INDEX(FREQUENCY({LOSr},{BINr}),{bi + 1})",
             hint="A bin's number is its upper limit. Which LOS values does that bin collect?",
             explanation=f"Each bin counts the values above the previous limit, up to and including its own limit. Bin "
                         f"{HIST_BIN} therefore counts stays of {lower + 1} or {HIST_BIN} days: more than {lower}, at most "
                         f"{HIST_BIN}. FREQUENCY returns one count per bin plus a final \"More\" count ({bin_more} stays over "
                         f"{LOS_BINS[-1]} days), and INDEX picks count number {bi + 1}. "
                         f"`=COUNTIFS({LOSr},\">{lower}\",{LOSr},\"<={HIST_BIN}\")` gives the same {hist_n}. If you read the "
                         f"label as \"{HIST_BIN} to {HIST_BIN + 1} days\" you'd get {hist_wrong}, which is the most common "
                         "histogram mistake."),
        Task("How strongly is a patient's age related to length of stay? Calculate the correlation coefficient between Age "
             "and LOSDays for all 400 stays, to 3 decimal places.",
             answer=r_age, fmt="0.000", tol=0.0006,
             solution=f"=CORREL({AGEr},{LOSr})",
             hint="CORREL(array1, array2)",
             explanation=f"r = {r_age:.3f} is a weak positive correlation. Older patients stay slightly longer on average, but "
                         f"the scatter plot is a cloud, not a line. Squaring r gives R² = {fit_age['r2']:.3f}, so age accounts "
                         f"for only about {fit_age['r2']:.1%} of the variation in LOS. `PEARSON` returns the same value, and the "
                         "order of the two ranges doesn't matter for correlation."),
        Task("Finance wants to know how much a day of stay adds to the bill. Fit a straight line that predicts TotalCharges "
             "from LOSDays (all 400 stays). What is the slope, in dollars per day? Enter it to 2 decimal places.",
             answer=fit_chg["slope"], fmt="#,##0.00",
             solution=f"=SLOPE({CHGr},{LOSr})",
             hint="SLOPE(known_y's, known_x's). The thing you predict goes first",
             explanation=f"Each extra day of stay goes with about ${fit_chg['slope']:,.2f} more in charges. The line is "
                         f"Charges = ${fit_chg['intercept']:,.2f} + ${fit_chg['slope']:,.2f} × LOSDays, and "
                         f"`=INTERCEPT({CHGr},{LOSr})` gives the intercept. `=RSQ({CHGr},{LOSr})` gives "
                         f"R² = {fit_chg['r2']:.1%}, so LOS explains a little under half of the stay-to-stay differences in "
                         "charges. The rest comes from diagnosis, procedures, ICU days, and so on. Swapping the arguments "
                         f"gives the slope of LOS on charges ({regress(chg, los)['slope']:.6f} days per dollar), which answers "
                         "a different question. SLOPE always takes the y-values (the outcome) first."),
        Task(f"Using the straight line that predicts LOSDays from Age (all 400 stays), what LOS does it predict for an "
             f"{PRED_AGE}-year-old patient? Enter days to 2 decimal places.",
             answer=pred80, fmt="0.00",
             solution=f"=FORECAST.LINEAR({PRED_AGE},{LOSr},{AGEr})",
             hint="FORECAST.LINEAR(x, known_y's, known_x's), or INTERCEPT + SLOPE × 80",
             explanation=f"The line is LOS = {fit_age['intercept']:.3f} + {fit_age['slope']:.4f} × Age, so at age {PRED_AGE} "
                         f"it predicts {pred80:.2f} days. That's the average LOS the line expects for {PRED_AGE}-year-olds, "
                         f"not a forecast for one patient. The typical miss around the line (`STEYX`) is "
                         f"{fit_age['steyx']:.2f} days, far more than the {30 * fit_age['slope']:.2f}-day difference the line "
                         f"predicts between a 50-year-old and an {PRED_AGE}-year-old. "
                         f"`=TREND({LOSr},{AGEr},{PRED_AGE})` and "
                         f"`=INTERCEPT({LOSr},{AGEr})+SLOPE({LOSr},{AGEr})*{PRED_AGE}` give the same result. The older "
                         "`FORECAST` function does too."),
        Task("Is the age effect from Task 5 real, or could it be chance? Run Regression (Data Analysis) with LOSDays as the Y "
             "range and Age as the X range, or build the p-value from LINEST as in guide section 7. What p-value does it "
             "report for the Age coefficient? Enter it to 4 decimal places.",
             answer=p_age, fmt="0.0000", tol=0.00006,
             solution=(f"=LET(fit,{linest},tstat,INDEX(fit,1,1)/INDEX(fit,2,1),"
                       f"T.DIST.2T(ABS(tstat),INDEX(fit,4,2)))"),
             hint="t Stat = coefficient ÷ its standard error. T.DIST.2T turns t into a two-tailed p-value",
             explanation=f"With stats set to TRUE, LINEST returns a 5 × 2 block: row 1 holds the slope ({fit_age['slope']:.4f}) "
                         f"and intercept, row 2 their standard errors ({fit_age['se_slope']:.4f} for the slope), and row 4 "
                         f"column 2 the residual degrees of freedom ({fit_age['df']}). The t Stat is "
                         f"{fit_age['slope']:.4f} ÷ {fit_age['se_slope']:.4f} = {fit_age['t_slope']:.2f}, and `T.DIST.2T` gives "
                         f"p = {p_age:.4f}. That's below 0.05, so the slope is **statistically significant**: a true slope of "
                         f"zero would rarely produce a sample like this one. Yet R² is only {fit_age['r2']:.1%}, so age explains "
                         "almost none of the variation in LOS. Significant means \"probably not zero,\" not \"large\" or "
                         "\"useful.\" Without LET (Excel 2019 and earlier), write "
                         f"`=T.DIST.2T(ABS(INDEX({linest},1,1)/INDEX({linest},2,1)),INDEX({linest},4,2))`."),
        Task("Finance would rather quote a range than a single number. For the TotalCharges-on-LOSDays line from Task 4, what "
             "is the lower end of the 95% confidence interval for the slope? It's the \"Lower 95%\" value the Regression tool "
             "reports for LOSDays. Enter dollars per day to 2 decimal places.",
             answer=fit_chg["lo_slope"], fmt="#,##0.00",
             solution=(f"=LET(fit,{linest_chg},INDEX(fit,1,1)-T.INV.2T(0.05,INDEX(fit,4,2))*INDEX(fit,2,1))"),
             hint="slope − t × (standard error of the slope), where t = T.INV.2T(0.05, residual df)",
             explanation=f"The slope is ${fit_chg['slope']:,.2f} with a standard error of ${fit_chg['se_slope']:,.2f}. With "
                         f"{fit_chg['df']} residual degrees of freedom, `T.INV.2T(0.05,{fit_chg['df']})` = {t_chg:.4f}, so the "
                         f"interval is ${fit_chg['slope']:,.2f} ± {t_chg:.4f} × ${fit_chg['se_slope']:,.2f}, which runs from "
                         f"${fit_chg['lo_slope']:,.2f} to ${fit_chg['hi_slope']:,.2f} per day. Finance can say each extra day "
                         f"goes with roughly ${fit_chg['lo_slope']:,.0f} to ${fit_chg['hi_slope']:,.0f} more in charges. It's "
                         "the same t × standard error recipe CONFIDENCE.T uses for a mean (guide section 9), applied to a "
                         "slope. Without LET, repeat the LINEST call inside each INDEX."),
        Task(f"Do Bluestone Memorial and Cedar Ridge differ in average LOS? Run a two-tailed t-test that does not assume equal "
             f"variances, comparing Memorial's LOSDays (rows {sf}–{mem_last}) with Cedar Ridge's (rows {ced_first}–{sl}). "
             "Enter the p-value to 3 decimal places.",
             answer=tt["p2"], fmt="0.000", tol=0.0006,
             solution=f"=T.TEST({S('LOSDays', sf, mem_last)},{S('LOSDays', ced_first, sl)},2,3)",
             hint="T.TEST(array1, array2, tails, type). Type 3 = two-sample, unequal variance",
             explanation=f"The sample means are {tt['m1']:.3f} days (Memorial) and {tt['m2']:.3f} days (Cedar Ridge), a "
                         f"difference of {tt['m1'] - tt['m2']:.3f} days. T.TEST returns the p-value directly: {tt['p2']:.3f}. "
                         "If the hospitals' true averages were equal, random sampling alone would produce a gap at least "
                         f"this large about {round(tt['p2'] * 100)}% of the time, so you "
                         "**can't conclude** they differ. That's not proof they're the same, only that this sample can't tell "
                         f"them apart. A one-tailed test (tails = 1) gives half the p-value ({tt['p1']:.3f}). Type 2 (equal "
                         f"variances) gives {tp['p2']:.3f} too, because with equal group sizes both tests use the same t "
                         f"statistic ({tt['t']:.3f}) and differ only slightly in degrees of freedom. The ToolPak's "
                         "\"t-Test: Two-Sample Assuming Unequal Variances\" reports the same value as P(T<=t) two-tail."),
        Task(f"Estimate Ashby Falls' true average door-to-provider time with a 95% confidence interval, using its Q4 2025 visits "
             f"(EDWaits rows {wf}–{wl}). What is the interval's half-width (the ± margin), in minutes to 2 decimal places? "
             "Blank cells are patients who left without being seen.",
             answer=ci_half, fmt="0.00",
             solution=f"=CONFIDENCE.T(0.05,STDEV.S({W}),COUNT({W}))",
             hint="CONFIDENCE.T(alpha, standard_dev, size). Alpha for 95% is 0.05",
             explanation=f"{w_n} patients saw a provider. Their mean wait was {w_mean:.2f} minutes with a standard deviation of "
                         f"{w_sd:.2f}. The margin is t × s ÷ √n = {w_t:.4f} × {w_sd:.2f} ÷ √{w_n} = {ci_half:.2f}, so the 95% "
                         f"confidence interval is {w_mean - ci_half:.1f} to {w_mean + ci_half:.1f} minutes. Use COUNT, not "
                         f"ROWS: ROWS also counts the {len(w_all) - w_n} blank rows, which aren't measurements, and gives "
                         f"{ci_wrong_n:.2f}. COUNTA agrees with COUNT here only because the blank cells are truly empty. It "
                         "would also count a text entry such as \"LWBS\" if an export had one. "
                         f"`CONFIDENCE.NORM` (and the old `CONFIDENCE`) uses 1.96 instead of t and gives "
                         f"{ci_norm:.2f}, slightly too narrow. The ToolPak's \"Confidence Level(95.0%)\" row is CONFIDENCE.T."),
        Task(f"A trailing 7-day moving average smooths out the weekday pattern. What was the 7-day moving average of "
             f"Memorial's daily arrivals on {ma_date:%m/%d/%Y} (that day and the 6 days before it)? Enter it to 2 decimal "
             "places.",
             answer=ma_val, fmt="0.00",
             solution=f"=AVERAGE(EDDaily!{MC}{ma_row - 6}:{MC}{ma_row})",
             hint=f"{ma_date:%m/%d/%Y} is EDDaily row {ma_row}. Average 7 rows ending there",
             explanation=f"Rows {ma_row - 6}–{ma_row} cover {ma_date - timedelta(days=6):%m/%d} through {ma_date:%m/%d/%Y}, "
                         f"one of each weekday, so the Monday peak and the weekend dip cancel out. {ma_val:.2f} arrivals a day "
                         f"was the highest 7-day average of 2025, against a 2025 daily mean of {m25:.2f}. A centered "
                         f"average (3 days either side) gives {centered:.2f} for this date. Trailing averages are what you "
                         "can calculate today, and centered ones are for describing the past. The ToolPak's Moving Average "
                         "tool and a chart's Moving Average trendline are both trailing."),
        Task("Fit a linear trend to the 24 monthly EDVisits values, using MonthNum (1–24) as x. What does it forecast for "
             "month 25 (January 2026)? Enter it to 1 decimal place.",
             answer=fc25, fmt="0.0", tol=0.051,
             solution=f"=FORECAST.LINEAR(25,{M('EDVisits')},{M('MonthNum')})",
             hint="FORECAST.LINEAR(x, known_y's, known_x's)",
             explanation=f"The trend line is almost flat (slope {fit_ed['slope']:+.2f} visits a month, R² = "
                         f"{fit_ed['r2']:.4f}), so the forecast is close to the 24-month average of {st.mean(edv):.1f}. But "
                         f"winter is the busy season: the two Januaries had {jan_actuals[0]} and {jan_actuals[1]} visits, and "
                         f"the two Decembers {dec_actuals[0]} and {dec_actuals[1]}. "
                         "A straight line can't see seasons, so it under-forecasts winter and over-forecasts summer. A "
                         "seasonal index or FORECAST.ETS (see the guide) handles that. `=TREND(…,…,25)` gives the same "
                         "number as FORECAST.LINEAR."),
        Task(f"Build an individuals (I) chart for Memorial's daily arrivals with 2024 as the baseline (EDDaily rows "
             f"{r24a}–{r24b}). The moving ranges are the absolute day-to-day changes, σ̂ = average moving range ÷ "
             f"{D2}, and UCL = mean + 3 × σ̂. What is the UCL? Enter it to 2 decimal places.",
             answer=base["ucl"], fmt="0.00",
             solution=f"={ucl_f}",
             hint="Average |today − yesterday| over 2024, divide by 1.128, triple it, add the mean",
             explanation=f"2024 averaged {base['mean']:.3f} arrivals a day. The {base['n_mr']} moving ranges average "
                         f"{base['mrbar']:.3f}, so σ̂ = {base['mrbar']:.3f} ÷ {D2} = {base['sigma']:.3f} and UCL = "
                         f"{base['mean']:.3f} + 3 × {base['sigma']:.3f} = {base['ucl']:.2f}. The LCL is "
                         f"{base['lcl']:.2f}. Many references write this as mean ± 2.66 × MR̄, which is the same thing, because "
                         f"3 ÷ {D2} ≈ 2.66. A helper column of `=ABS(C3-C2)` filled down and averaged works too. The moving "
                         "range measures short-term, day-to-day noise, so a slow seasonal swing doesn't widen the limits the "
                         f"way STDEV.S does (it would put the UCL at {sd_ucl:.2f})."),
        Task(f"Monitor 2025 against the 2024 limits. On how many days in 2025 (EDDaily rows {r25a}–{r25b}) did Memorial's "
             f"arrivals exceed the UCL from Task 12?",
             answer=len(above),
             solution=f'=COUNTIF(EDDaily!{MC}{r25a}:{MC}{r25b},">"&D{prow(12)})',
             live=f'=COUNTIF(EDDaily!{MC}{r25a}:{MC}{r25b},">"&({ucl_f}))',
             hint=f"COUNTIF with \">\"& your Task 12 cell (D{prow(12)})",
             explanation=f"{len(above)} days broke the limit: "
                         + ", ".join(f"{d:%m/%d} ({v})" for d, v in above)
                         + f". Each is a **special-cause** signal worth a look, and in an ED the usual causes are respiratory "
                         f"season, a mass-casualty event, or a neighboring ED on diversion. No day fell below the LCL. "
                         f"Freezing the limits from a baseline year is standard practice: if you recalculate them with the new "
                         f"data, a lasting change gets absorbed into the limits. Limits from STDEV.S would flag only "
                         f"{above_sd} day{'s' if above_sd != 1 else ''}."),
    ]

    # ------------------------------------------------------------------ bonus: p-chart of monthly readmission rates
    n_i = [r["IndexStays"] for r in monthly]
    x_i = [r["Readmits30"] for r in monthly]
    pbar = sum(x_i) / sum(n_i)

    def limits(p, n):
        s = math.sqrt(p * (1 - p) / n)
        return p - 3 * s, p + 3 * s

    k_small = min(range(24), key=lambda i: (n_i[i], i))
    small = monthly[k_small]
    ucl_small = limits(pbar, n_i[k_small])[1]
    outside = []
    for i, r in enumerate(monthly):
        lo, hi = limits(pbar, n_i[i])
        p = x_i[i] / n_i[i]
        assert abs(p - hi) > 0.002 and abs(p - lo) > 0.002
        if p > hi or p < lo:
            outside.append(i)
    assert len(outside) == 1 and monthly[outside[0]]["MonthStart"] == date(2025, 12, 1)
    sig_i = outside[0]
    sig = monthly[sig_i]
    sig_lo = limits(pbar, n_i[sig_i])[0]
    assert all(limits(pbar, n)[0] > 0 for n in n_i), "LCL never clips at 0, so the ABS() one-formula count is valid"
    assert not any(x_i[i] / n_i[i] > limits(pbar, n_i[i])[1] for i in range(24)), "B3's explanation says none is above"
    # corrected centre line without the incomplete month, and the z-score of the highest remaining month
    keep = [i for i in range(24) if i != sig_i]
    pbar2 = sum(x_i[i] for i in keep) / sum(n_i[i] for i in keep)
    k_hi = max(keep, key=lambda i: x_i[i] / n_i[i])
    hi_row = monthly[k_hi]
    z_hi = (x_i[k_hi] / n_i[k_hi] - pbar2) / math.sqrt(pbar2 * (1 - pbar2) / n_i[k_hi])
    z_hi_old = (x_i[k_hi] / n_i[k_hi] - pbar) / math.sqrt(pbar * (1 - pbar) / n_i[k_hi])
    assert sig_i == ml - mf, "the signal month is the last row, so the corrected ranges stop one row earlier"

    small_row = mf + k_small
    hi_mrow = mf + k_hi
    N_, X_, A_ = M("IndexStays"), M("Readmits30"), M("MonthStart")
    pbar_f = f"SUM({X_})/SUM({N_})"
    out_test = f"ABS({X_}/{N_}-D6)>3*SQRT(D6*(1-D6)/{N_})"
    out_test_live = f"ABS({X_}/{N_}-({pbar_f}))>3*SQRT(({pbar_f})*(1-({pbar_f}))/{N_})"
    N2, X2 = M("IndexStays", mf, ml - 1), M("Readmits30", mf, ml - 1)
    nc, xc = mon.col("IndexStays"), mon.col("Readmits30")

    L.bonus_title = "Bonus: Is the readmission rate out of control?"
    L.bonus_scenario = (
        "The quality committee wants a p-chart of Bluestone's system-wide 30-day readmission rate by discharge month, "
        f"January 2024 to December 2025 (Monthly sheet, rows {mf}–{ml}). For each month, n = IndexStays (inpatient discharges "
        "of patients who didn't die) and the rate = Readmits30 ÷ IndexStays. The center line p̄ = total readmissions ÷ total "
        "index stays, and each month gets its own limits: p̄ ± 3 × √(p̄ × (1 − p̄) ÷ n). Your B1 answer lands in cell D6 of "
        "this sheet, so later parts can refer to it. Helper columns for each month's rate, UCL, and LCL make this much "
        "easier. Put them in empty columns to the right of the Monthly table."
    )
    L.bonus = [
        Task("What is the center line p̄ for all 24 months? Enter it as a percentage to 2 decimal places.",
             answer=pbar, fmt="0.00%",
             solution=f"={pbar_f}",
             hint="Total readmissions ÷ total index stays. Don't average the 24 monthly rates",
             explanation=f"{sum(x_i)} readmissions ÷ {sum(n_i):,} index stays = {pbar:.2%}. Averaging the 24 monthly rates "
                         f"gives {st.mean(x / n for x, n in zip(x_i, n_i)):.2%} instead, because it weights a small month the "
                         "same as a big one. A rate for a period is total ÷ total (Lesson 1.4)."),
        Task(f"{small['MonthStart']:%B %Y} had the fewest index stays ({n_i[k_small]}). What is its upper control limit? "
             "Enter it as a percentage to 2 decimal places.",
             answer=ucl_small, fmt="0.00%",
             solution=f"=D6+3*SQRT(D6*(1-D6)/Monthly!{nc}{small_row})",
             live=f"=({pbar_f})+3*SQRT(({pbar_f})*(1-({pbar_f}))/Monthly!{nc}{small_row})",
             hint="p̄ + 3 × SQRT(p̄ × (1 − p̄) / n), with this month's n",
             explanation=f"√({pbar:.4f} × {1 - pbar:.4f} ÷ {n_i[k_small]}) = {math.sqrt(pbar * (1 - pbar) / n_i[k_small]):.4f}, "
                         f"so the UCL is {pbar:.4f} + 3 × that = {ucl_small:.2%}. The biggest month "
                         f"({monthly[max(range(24), key=lambda i: n_i[i])]['MonthStart']:%b %Y}, n = {max(n_i)}) gets a UCL of "
                         f"{limits(pbar, max(n_i))[1]:.2%}. Smaller months get wider limits because a rate from fewer "
                         "patients bounces around more. That's why a p-chart's limit lines look like a staircase."),
        Task("How many of the 24 months fall outside their control limits (above the UCL or below the LCL)?",
             answer=len(outside),
             solution=f"=SUMPRODUCT(--({out_test}))",
             live=f"=SUMPRODUCT(--({out_test_live}))",
             hint="Compare each month's rate with its own limits. Add helper columns for the rate, UCL, LCL, and a "
                  "TRUE/FALSE flag, then COUNTIF the flags. One SUMPRODUCT also works",
             explanation="Only one month signals, and it falls below its LCL. No month is above its UCL. "
                         "`ABS(rate − p̄) > 3σ` catches both directions in one test because "
                         "every LCL here is above zero. When n is small, p̄ − 3σ can go negative and the LCL is set to 0, so "
                         f"test the two sides separately. With helper columns on the Monthly sheet: Rate "
                         f"`={xc}{mf}/{nc}{mf}`, UCL `=Bonus!$D$6+3*SQRT(Bonus!$D$6*(1-Bonus!$D$6)/{nc}{mf})`, LCL the same "
                         "with − instead of +, and a flag column that tests whether the rate is above the UCL or below the "
                         "LCL. Then COUNTIF the flags."),
        Task("Which month is it? Enter its MonthStart date.",
             answer=sig["MonthStart"], fmt="mm/dd/yyyy",
             solution=f"=XLOOKUP(TRUE,{out_test},{A_})",
             live=f"=XLOOKUP(TRUE,{out_test_live},{A_})",
             hint="Read it off your helper columns, or XLOOKUP(TRUE, your test, the MonthStart column)",
             explanation=f"{sig['MonthStart']:%B %Y}: {x_i[sig_i]} readmissions out of {n_i[sig_i]} index stays = "
                         f"{x_i[sig_i] / n_i[sig_i]:.2%}, below its LCL of {sig_lo:.2%}. Before anyone celebrates, ask "
                         "why. The data ends on 12/31/2025, so a patient discharged on 12/20 has only 11 days of follow-up "
                         "in the data, and any readmission in January 2026 hasn't happened yet. December's rate is low "
                         "because the 30-day window isn't complete. That's a **data artifact**, not an improvement. Excluding "
                         "months with incomplete follow-up is standard for readmission reporting."),
        Task(f"Drop December 2025 and recompute p̄ from the other 23 months (rows {mf}–{ml - 1}). The highest remaining month is "
             f"{hi_row['MonthStart']:%B %Y} (row {hi_mrow}). How many standard errors above the new center line is it? "
             "Calculate z = (rate − p̄) ÷ √(p̄ × (1 − p̄) ÷ n) and enter it to 2 decimal places.",
             answer=z_hi, fmt="0.00",
             solution=(f"=LET(p,SUM({X2})/SUM({N2}),n,Monthly!{nc}{hi_mrow},"
                       f"(Monthly!{xc}{hi_mrow}/n-p)/SQRT(p*(1-p)/n))"),
             hint="LET(p, new p̄, n, that month's IndexStays, (rate − p) / SQRT(p*(1−p)/n))",
             explanation=f"The corrected p̄ is {pbar2:.2%}, higher than {pbar:.2%} because the artificially low December no "
                         f"longer drags it down. {hi_row['MonthStart']:%B %Y}'s rate of {x_i[k_hi] / n_i[k_hi]:.2%} is "
                         f"z = {z_hi:.2f} standard errors above it (it was {z_hi_old:.2f} against the old center line). A point "
                         "is a special-cause signal only beyond ±3, so every month is inside the limits: readmissions are "
                         "**stable**, showing common-cause variation only. The chart says the committee shouldn't react to any "
                         "single month. To lower the rate, change the process (discharge planning, follow-up calls) and "
                         "then watch for a shift, such as 8 months in a row below the center line."),
    ]
    return L
