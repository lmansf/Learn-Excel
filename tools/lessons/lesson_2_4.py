"""Lesson 2.4 · Math & Statistical Functions."""
from __future__ import annotations

import math
import statistics as st
from collections import Counter

from xlcourse import Lesson, Task, data

CODE = "2.4"

VANC_DOSE_MG = 1250      # every vancomycin order in the data is 1,250 mg per dose
VANC_VIAL_MG = 500       # teaching assumption: single-dose 500 mg vials only
BOARD_ROUND_MIN = 15     # ED status board rounds length of stay to the nearest 15 minutes
ON_TARGET_DAYS = 1       # "on target" = LOS within 1 day of the expected LOS
TRIM_PCT = 0.10          # TRIMMEAN percent for the bonus


# ---------------------------------------------------------------------------
# Excel-equivalent statistics (so every answer is computed the way Excel does it)
# ---------------------------------------------------------------------------
def pct_inc(values, k):
    """PERCENTILE.INC / QUARTILE.INC: position k*(n-1) in the sorted list, linear interpolation."""
    s = sorted(values)
    h = k * (len(s) - 1)
    lo = math.floor(h)
    if lo + 1 >= len(s):
        return s[-1]
    return s[lo] + (h - lo) * (s[lo + 1] - s[lo])


def pct_exc(values, k):
    """PERCENTILE.EXC / QUARTILE.EXC: position k*(n+1) (1-based), linear interpolation."""
    s = sorted(values)
    h = k * (len(s) + 1)
    lo = math.floor(h)
    return s[lo - 1] + (h - lo) * (s[lo] - s[lo - 1])


def mode_sngl(values):
    """MODE.SNGL: the most frequent value; on a tie, the one that appears first in the range."""
    counts = Counter(values)
    best = max(counts.values())
    return next(v for v in values if counts[v] == best)


def rank_eq(x, values):
    """RANK.EQ with order 0 (largest = 1)."""
    return 1 + sum(1 for v in values if v > x)


def rank_avg(x, values):
    ties = sum(1 for v in values if v == x)
    return rank_eq(x, values) + (ties - 1) / 2


def trimmean(values, pct):
    """TRIMMEAN: drop floor(n*pct) rounded down to an even number, half from each end."""
    s = sorted(values)
    k = math.floor(len(s) * pct)
    k -= k % 2
    k //= 2
    return st.mean(s[k:len(s) - k]), k


def excel_mround(x, m):
    """MROUND for positive x and m (halves round away from zero)."""
    q = x / m
    return m * math.floor(q + 0.5)


def ceiling_math(x, m=1):
    return m * math.ceil(x / m)


def floor_math(x, m=1):
    return m * math.floor(x / m)


def _age(dob, on):
    return on.year - dob.year - ((on.month, on.day) < (dob.month, dob.day))


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="02-formulas-functions", slug="04-math-statistical-functions",
        title="Math & Statistical Functions", level="Beginner → Intermediate", minutes=50,
        objectives=[
            "Round correctly with ROUND, ROUNDUP, ROUNDDOWN, MROUND, CEILING.MATH, and FLOOR.MATH",
            "Use INT, TRUNC, MOD, ABS, and SUMPRODUCT",
            "Describe data with MEDIAN, MODE, STDEV, PERCENTILE, and QUARTILE — and know when the mean misleads",
            "Rank and pick extremes with RANK.EQ, LARGE, and SMALL",
        ],
        data_note="All 573 inpatient stays discharged from Bluestone Memorial Hospital in Q4 2025 (October–December), with "
                  "age, length of stay, expected length of stay, and charges. All 785 emergency department visits at Bluestone "
                  "Memorial in November–December 2025, with wait and length-of-stay minutes. The 388 medication orders written "
                  "for that quarter's Medical-Surgical 5 East stays.",
    )
    L.start_notes.append("The vancomycin vial rule in this lesson is a simplified teaching assumption, not pharmacy guidance.")

    # ------------------------------------------------------------------ data: inpatient stays
    dx = data.index(data.load("diagnoses"), "DxCode")
    pts = data.index(data.load("patients"), "PatientID")
    depts = data.index(data.load("departments"), "DeptID")
    stays = [e for e in data.load("encounters")
             if e["EncounterType"] == "Inpatient" and e["FacilityID"] == "F01"
             and e["DischargeDateTime"].year == 2025 and e["DischargeDateTime"].month >= 10]
    stays.sort(key=lambda e: (e["DischargeDateTime"], e["EncounterID"]))
    for e in stays:
        e["Unit"] = depts[e["DeptID"]]["DeptName"]
        e["AdmitDate"] = e["AdmitDateTime"].date()
        e["DischargeDate"] = e["DischargeDateTime"].date()
        e["Age"] = _age(pts[e["PatientID"]]["DOB"], e["AdmitDate"])
        e["DxDescription"] = dx[e["PrimaryDxCode"]]["DxDescription"]
        e["ExpectedLOS"] = dx[e["PrimaryDxCode"]]["ExpectedLOS"]
        e["LOSDays"] = (e["DischargeDate"] - e["AdmitDate"]).days   # midnights in hospital
    sty = L.add_table_sheet(
        "Stays", stays, table="tblStays",
        columns=["EncounterID", "Unit", "AdmitDate", "DischargeDate", "Age", "PrimaryDxCode", "DxDescription",
                 "ExpectedLOS", "LOSDays", "TotalCharges"],
        extra_cols=["AgeBand"],
        formats={"ExpectedLOS": "0.0", "TotalCharges": "#,##0.00", "AgeBand": "0"},
        widths={"Unit": 26, "DxDescription": 44, "AgeBand": 11, "TotalCharges": 14},
    )

    # ------------------------------------------------------------------ data: ED visits
    ed = [r for r in data.load("ed_visits")
          if r["FacilityID"] == "F01" and r["ArrivalDateTime"].year == 2025 and r["ArrivalDateTime"].month in (11, 12)]
    ed.sort(key=lambda r: (r["ArrivalDateTime"], r["EDVisitID"]))
    for r in ed:
        seen = r["ProviderSeenDateTime"]
        r["DoorToProviderMin"] = round((seen - r["ArrivalDateTime"]).total_seconds() / 60) if seen else None
        r["EDLOSMin"] = round((r["DepartureDateTime"] - r["ArrivalDateTime"]).total_seconds() / 60)
    edd = L.add_table_sheet(
        "ED", ed, table="tblED",
        columns=["EDVisitID", "ArrivalDateTime", "ArrivalMode", "ESILevel", "ChiefComplaint", "EDDisposition",
                 "DoorToProviderMin", "EDLOSMin"],
        widths={"ArrivalDateTime": 17, "ChiefComplaint": 30, "DoorToProviderMin": 19, "EDLOSMin": 11},
    )

    # ------------------------------------------------------------------ data: medication orders (5 East stays)
    east_ids = {e["EncounterID"] for e in stays if e["DeptID"] == "D111"}
    meds = [m for m in data.load("medications") if m["EncounterID"] in east_ids]
    meds.sort(key=lambda m: (m["OrderDateTime"], m["MedOrderID"]))
    med = L.add_table_sheet(
        "Meds", meds, table="tblMeds",
        columns=["MedOrderID", "EncounterID", "OrderDateTime", "MedicationName", "Dose", "Route", "Frequency",
                 "DosesDispensed", "UnitCost"],
        formats={"UnitCost": "#,##0.00"},
        widths={"OrderDateTime": 17, "MedicationName": 26, "DosesDispensed": 15},
    )

    # ------------------------------------------------------------------ helpers
    sf, sl = sty.first_row, sty.last_row
    ef, el = edd.first_row, edd.last_row
    mf, ml = med.first_row, med.last_row

    def sr(col):   # Stays range as a learner types it
        return f"Stays!{sty.col(col)}{sf}:{sty.col(col)}{sl}"

    def er(col):
        return f"ED!{edd.col(col)}{ef}:{edd.col(col)}{el}"

    def mr(col):
        return f"Meds!{med.col(col)}{mf}:{med.col(col)}{ml}"

    # ------------------------------------------------------------------ answers (computed in Python)
    los = [e["LOSDays"] for e in stays]
    charges = [e["TotalCharges"] for e in stays]
    ages = [e["Age"] for e in stays]
    exp_los = [e["ExpectedLOS"] for e in stays]
    esi = [r["ESILevel"] for r in ed]
    d2p = [r["DoorToProviderMin"] for r in ed if r["DoorToProviderMin"] is not None]
    edlos = [r["EDLOSMin"] for r in ed]
    lwbs_n = len(ed) - len(d2p)

    # 1 · mode of ESI
    esi_mode = mode_sngl(esi)
    esi_mode_n = esi.count(esi_mode)
    esi_mean = st.mean(esi)

    # 2 · mean vs median charges
    ch_mean, ch_median = st.mean(charges), st.median(charges)
    below_mean_n = sum(c < ch_mean for c in charges)

    # 3 · standard deviation to the nearest $100
    ch_sd = st.stdev(charges)
    ch_sdp = st.pstdev(charges)
    sd_100 = int(round(ch_sd, -2))
    assert sd_100 != int(round(ch_sdp, -2)), "STDEV.P should round differently, so the task catches it"

    # 4 · 3rd-longest ED stay, reported in whole hours with ROUNDUP (any started hour counts)
    edlos_desc = sorted(edlos, reverse=True)
    third_longest = edlos_desc[2]
    third_longest_hrs = third_longest / 60
    third_longest_h_up = math.ceil(third_longest_hrs)
    assert not third_longest_hrs.is_integer()
    third_longest_h_round = math.floor(third_longest_hrs + 0.5)
    assert third_longest_h_round != third_longest_h_up, "ROUND should give a different answer, so the check catches it"

    # 5 · 3rd-lowest charge
    third_lowest = sorted(charges)[2]

    # 6 · RANK.EQ with ties: the first stay (sheet order) with a 10-day LOS
    rank_los = 10
    r6_i = next(i for i, e in enumerate(stays) if e["LOSDays"] == rank_los)
    r6 = stays[r6_i]
    r6_row = sf + r6_i
    r6_rank = rank_eq(rank_los, los)
    r6_ties = los.count(rank_los)
    r6_avg = rank_avg(rank_los, los)

    # 7 · 90th percentile door-to-provider (blanks = LWBS, ignored)
    d2p_p90 = pct_inc(d2p, 0.9)
    d2p_med = st.median(d2p)
    d2p_over_n = sum(1 for d in d2p if d > pct_inc(d2p, 0.9))
    assert float(d2p_p90).is_integer()
    d2p_p90 = int(d2p_p90)

    # 8 · IQR of ED length of stay
    edlos_q1, edlos_q3 = pct_inc(edlos, 0.25), pct_inc(edlos, 0.75)
    edlos_iqr = edlos_q3 - edlos_q1
    assert float(edlos_iqr).is_integer()
    edlos_iqr = int(edlos_iqr)
    edlos_q1x, edlos_q3x = pct_exc(edlos, 0.25), pct_exc(edlos, 0.75)

    # 9 · SUMPRODUCT total medication cost
    med_cost = sum(m["DosesDispensed"] * m["UnitCost"] for m in meds)
    med_doses = sum(m["DosesDispensed"] for m in meds)
    wrong_sumsum = sum(m["DosesDispensed"] for m in meds) * sum(m["UnitCost"] for m in meds)

    # 10 · age bands with FLOOR.MATH: stays in the 80-89 band
    band = 80
    band_n = sum(1 for a in ages if floor_math(a, 10) == band)
    # What ROUND(age,-1)=80 would count instead: Excel rounds halves away from zero, so ages 75-84.
    band_round_n = sum(1 for a in ages if band - 5 <= a <= band + 4)
    assert band_n != band_round_n
    band_top = mode_sngl([floor_math(a, 10) for a in ages])

    # 11 · ED board: MROUND to 15 minutes, then hours + minutes as text
    def board_ok(r):
        m = r["EDLOSMin"]
        rm = excel_mround(m, BOARD_ROUND_MIN)
        return (r["EDDisposition"] == "Admitted" and m % BOARD_ROUND_MIN >= 8 and rm % 60 >= 10
                and rm // 60 >= 5 and m % 60 >= 10)
    b_i = next(i for i, r in enumerate(ed) if board_ok(r))
    b = ed[b_i]
    b_row = ef + b_i
    b_min = b["EDLOSMin"]
    b_round = excel_mround(b_min, BOARD_ROUND_MIN)
    b_text = f"{b_round // 60} h {b_round % 60} min"
    b_raw_text = f"{b_min // 60} h {b_min % 60} min"
    hh, mm = b_round // 60, b_round % 60
    b_accept = [f"{hh}h {mm}min", f"{hh}h{mm}min", f"{hh} h {mm} m", f"{hh}h {mm}m", f"{hh}h{mm}m", f"{hh} hr {mm} min",
                f"{hh} h {mm} mins", f"{hh} hrs {mm} min"]

    # 12 · on-target stays: |LOS - expected| <= 1 day
    diffs = [l - x for l, x in zip(los, exp_los)]
    on_target = sum(1 for d in diffs if abs(d) <= ON_TARGET_DAYS)
    no_abs = sum(1 for d in diffs if d <= ON_TARGET_DAYS)

    # 13 · vancomycin vials for the longest vancomycin order
    vanc = [(i, m) for i, m in enumerate(meds) if m["MedicationName"] == "Vancomycin"]
    assert all(m["Dose"] == f"{VANC_DOSE_MG} mg" for _, m in vanc)
    v_i, v = max(vanc, key=lambda t: (t[1]["DosesDispensed"], -t[0]))
    assert sum(1 for _, m in vanc if m["DosesDispensed"] == v["DosesDispensed"]) == 1
    v_row = mf + v_i
    v_doses = v["DosesDispensed"]
    vials_per_dose = ceiling_math(VANC_DOSE_MG, VANC_VIAL_MG) // VANC_VIAL_MG
    v_vials = vials_per_dose * v_doses
    v_wrong = ceiling_math(VANC_DOSE_MG * v_doses, VANC_VIAL_MG) // VANC_VIAL_MG
    v_waste = (ceiling_math(VANC_DOSE_MG, VANC_VIAL_MG) - VANC_DOSE_MG) * v_doses

    # ------------------------------------------------------------------ bonus answers
    q1, q3 = pct_inc(charges, 0.25), pct_inc(charges, 0.75)
    iqr = q3 - q1
    fence = q3 + 1.5 * iqr
    low_fence = q1 - 1.5 * iqr
    # COUNTIF(">"&fence) converts the fence to 15 significant digits; make sure no charge sits on that edge.
    assert not any(abs(c - fence) < 0.01 for c in charges)
    high = [c for c in charges if c > fence]
    high_n = len(high)
    high_share = sum(high) / sum(charges)
    sd3 = ch_mean + 3 * ch_sd
    assert not any(abs(c - sd3) < 0.01 for c in charges)
    sd3_n = sum(1 for c in charges if c > sd3)
    tmean, trim_k = trimmean(charges, TRIM_PCT)

    # ------------------------------------------------------------------ practice tasks
    J = sr("TotalCharges")
    L.practice_intro = (f"Tasks use three data sheets. Stays holds {len(stays)} inpatient stays (rows {sf}–{sl}), ED holds "
                        f"{len(ed)} visits (rows {ef}–{el}), and Meds holds {len(meds)} medication orders (rows {mf}–{ml}). "
                        f"Reference a column's data rows, like {J}, rather than a whole column like J:J, because the header "
                        "text in row 1 breaks the row-by-row math in SUMPRODUCT. Structured references such as "
                        "tblStays[TotalCharges] work too. When a task asks you to round, round with a function so the stored "
                        "value matches, not just the display.")

    L.tasks = [
        Task("What is the most common ESI triage level among the ED visits?",
             answer=esi_mode,
             solution=f"=MODE.SNGL({er('ESILevel')})",
             hint="MODE.SNGL",
             explanation=f"MODE.SNGL returns the value that occurs most often: level {esi_mode} appears {esi_mode_n} times out of "
                         f"{len(ed)}. ESI levels are codes on a 1–5 scale, so their mean ({esi_mean:.2f}) isn't a real triage "
                         "level. For codes and categories, the mode is the honest \"typical\" value. The older `MODE` function gives "
                         "the same result."),
        Task("How much higher is the mean (average) TotalCharges per stay than the median TotalCharges? Enter the "
             "difference in dollars, to the cent.",
             answer=round(ch_mean - ch_median, 2), fmt="#,##0.00",
             solution=f"=AVERAGE({J})-MEDIAN({J})",
             hint="AVERAGE(…) − MEDIAN(…)",
             explanation=f"The mean is ${ch_mean:,.2f} but the median is ${ch_median:,.2f}, so half of all stays were charged less "
                         f"than ${ch_median:,.0f}. In fact {below_mean_n} of the {len(stays)} stays ({below_mean_n / len(stays):.0%}) "
                         "fall below the mean. A handful of very long, very expensive stays pull the mean up, while the median "
                         "only cares about the middle value. When data is skewed like this, the median describes a typical stay "
                         "better."),
        Task("What is the sample standard deviation of TotalCharges, rounded to the nearest $100 with ROUND?",
             answer=sd_100, fmt="#,##0",
             solution=f"=ROUND(STDEV.S({J}),-2)",
             hint="STDEV.S, then ROUND with a negative num_digits",
             explanation=f"STDEV.S returns ${ch_sd:,.2f}. A num_digits of -2 rounds to the hundreds place, giving "
                         f"${sd_100:,}. STDEV.S treats the stays as a sample of an ongoing process, which is the usual choice. "
                         f"STDEV.P gives ${ch_sdp:,.2f}, which rounds to ${int(round(ch_sdp, -2)):,}, so the check tells the two "
                         "apart."),
        Task("ED leaders review the three longest ED visits for boarding delays, where an admitted patient waits in the ED "
             "for an inpatient bed. They report each visit's length of stay in whole hours and count any started hour as a "
             "full hour. Find the 3rd-longest EDLOSMin on the ED sheet, convert it to hours, and round it up to a whole "
             "number of hours with ROUNDUP.",
             answer=third_longest_h_up,
             solution=f"=ROUNDUP(LARGE({er('EDLOSMin')},3)/60,0)",
             hint="LARGE(array, k) finds the visit. Divide by 60, then ROUNDUP(…, 0)",
             explanation=f"LARGE returns the k-th largest value, so `LARGE(range,1)` is the same as MAX. The three longest visits "
                         f"lasted {edlos_desc[0]:,}, {edlos_desc[1]:,}, and {edlos_desc[2]:,} minutes. The 3rd is "
                         f"{third_longest:,} ÷ 60 = {third_longest_hrs:.2f} hours, and ROUNDUP turns the started hour into a "
                         f"full one, giving {third_longest_h_up}. `ROUND` and `ROUNDDOWN` both give {third_longest_h_round} "
                         "here, which would under-report the visit because the rule counts every started hour. "
                         f"`=CEILING.MATH(LARGE(…,3)/60)` gives the same {third_longest_h_up}. Change k to 2 or 1 to see "
                         "the other two visits."),
        Task("The charge-capture team audits the cheapest stays, because an unusually low charge often means some charges "
             "were never posted. What is the 3rd-lowest TotalCharges?",
             answer=third_lowest, fmt="#,##0.00",
             solution=f"=SMALL({J},3)",
             hint="SMALL is LARGE's mirror image",
             explanation="SMALL(array, k) returns the k-th smallest value, so `SMALL(range,1)` equals MIN. Like LARGE, it ignores "
                         "blank cells and text."),
        Task(f"Encounter {r6['EncounterID']} (Stays row {r6_row}) stayed {rank_los} days. Rank its LOSDays among all "
             f"{len(stays)} stays with RANK.EQ, where the longest stay is rank 1.",
             answer=r6_rank,
             solution=f"=RANK.EQ(Stays!{sty.col('LOSDays')}{r6_row},{sr('LOSDays')})",
             hint="RANK.EQ(number, ref). Leaving out order ranks the largest value as 1",
             explanation=f"{r6_rank - 1} stays were longer than {rank_los} days, so this one ranks {r6_rank}. The other "
                         f"{r6_ties - 1} stays of exactly {rank_los} days share rank {r6_rank}, and the next rank used is "
                         f"{r6_rank + r6_ties}. `RANK.AVG` returns {r6_avg:g} instead, which is the average of positions "
                         f"{r6_rank}–{r6_rank + r6_ties - 1}. If you copy a RANK formula down a column, lock the ref with $ "
                         f"(`$I${sf}:$I${sl}`) so it doesn't slide."),
        Task(f"Ninety percent of the patients who saw a provider waited at most how many minutes? Calculate the 90th "
             f"percentile of DoorToProviderMin. (The {lwbs_n} blank cells are patients who left without being seen.)",
             answer=d2p_p90,
             solution=f"=PERCENTILE.INC({er('DoorToProviderMin')},0.9)",
             hint="PERCENTILE.INC(array, 0.9)",
             explanation=f"PERCENTILE.INC ignores blank cells, so it uses only the {len(d2p)} visits with a provider time. The "
                         f"median wait was {d2p_med:g} minutes, yet {d2p_over_n} of the {len(d2p)} waits "
                         f"({d2p_over_n / len(d2p):.1%}) were longer than {d2p_p90} minutes. That long tail is why ED leaders "
                         "track the 90th percentile alongside the median. If someone had typed 0 into the blank cells, the "
                         "percentile would drop and the report would look better than reality."),
        Task("What is the interquartile range (IQR = Q3 − Q1) of EDLOSMin, in minutes? Use QUARTILE.INC.",
             answer=edlos_iqr,
             solution=f"=QUARTILE.INC({er('EDLOSMin')},3)-QUARTILE.INC({er('EDLOSMin')},1)",
             hint="QUARTILE.INC(array, 3) − QUARTILE.INC(array, 1)",
             explanation=f"Q1 is {edlos_q1:g} and Q3 is {edlos_q3:g}, so the middle half of ED visits lasted between "
                         f"{int(edlos_q1) // 60} h {int(edlos_q1) % 60} min and {int(edlos_q3) // 60} h {int(edlos_q3) % 60} min. "
                         "The IQR ignores the extreme visits at both ends, so a few boarding patients can't stretch it. "
                         + (f"QUARTILE.EXC interpolates slightly differently, but with {len(edlos)} visits it lands on the same "
                            "Q1 and Q3 here." if (edlos_q1x, edlos_q3x) == (edlos_q1, edlos_q3) else
                            f"QUARTILE.EXC interpolates slightly differently and gives Q1 = {edlos_q1x:g} and Q3 = {edlos_q3x:g}.")),
        Task("What was the total acquisition cost of the medications dispensed for the 5 East stays? Multiply each order's "
             "DosesDispensed by its UnitCost and add up all the orders, in one formula. Enter the total in dollars, to the "
             "cent.",
             answer=round(med_cost, 2), fmt="#,##0.00",
             solution=f"=SUMPRODUCT({mr('DosesDispensed')},{mr('UnitCost')})",
             hint="SUMPRODUCT(array1, array2)",
             explanation=f"SUMPRODUCT multiplies the two columns row by row, then adds the {len(meds)} products. It gives the same "
                         "result as a helper column of `=H2*I2` totaled with SUM. Multiplying two totals, as in "
                         f"`=SUM(H…)*SUM(I…)`, is wrong (it gives ${wrong_sumsum:,.2f}) because it pairs every order's doses with "
                         f"every other order's price. Divide by `SUM({mr('DosesDispensed')})` to get the average cost per dose, "
                         f"${med_cost / med_doses:,.2f}."),
        Task(f"Fill the yellow AgeBand column on the Stays sheet with each stay's 10-year age band using FLOOR.MATH, so "
             f"age 78 becomes 70 and age 80 becomes 80. Type the formula in {sty.cell('AgeBand', 0, sheet=False)}. Stays is "
             "an Excel Table, so Excel fills the formula down the whole column for you. If it doesn't, double-click the "
             f"fill handle. The gray cell counts the stays in the {band}–{band + 9} band.",
             answer=band_n, title=f"AgeBand column with FLOOR.MATH (stays aged {band}–{band + 9})",
             solution=f"=FLOOR.MATH({sty.col('Age')}{sf},10)",
             summary=f'=IF(COUNT(Stays!${sty.col("AgeBand")}${sf}:${sty.col("AgeBand")}${sl})=0,"",'
                     f'COUNTIF(Stays!${sty.col("AgeBand")}${sf}:${sty.col("AgeBand")}${sl},{band}))',
             fill={"range": f"Stays!{sty.col('AgeBand')}{sf}:{sty.col('AgeBand')}{sl}",
                   "formula": f"=FLOOR.MATH({sty.col('Age')}{sf},10)"},
             live=f"=SUMPRODUCT(({sr('Age')}>={band})*({sr('Age')}<{band + 10}))",
             hint="FLOOR.MATH(number, significance) rounds down to a multiple",
             explanation=f"FLOOR.MATH rounds down to the nearest multiple of 10, so every age from {band} to {band + 9} lands in "
                         f"the {band} band. `ROUND(E2,-1)` would be wrong, because it sends ages 75–84 to 80 and gives "
                         f"{band_round_n} instead of {band_n}. `=INT(E2/10)*10`, `=TRUNC(E2,-1)`, and `=ROUNDDOWN(E2,-1)` all work "
                         f"too. In the Table you may see `=FLOOR.MATH([@Age],10)`. The most common band is the {band_top}s."),
        Task(f"Order {v['MedOrderID']} (Meds row {v_row}) is vancomycin {VANC_DOSE_MG:,} mg every 12 hours, with "
             f"{v_doses} doses dispensed. Assume the IV room mixes every dose from {VANC_VIAL_MG} mg single-dose vials and "
             "throws away whatever is left in an opened vial. How many vials did this order use?",
             answer=v_vials,
             solution=f"=CEILING.MATH({VANC_DOSE_MG},{VANC_VIAL_MG})/{VANC_VIAL_MG}*Meds!{med.col('DosesDispensed')}{v_row}",
             hint="Work out vials per dose first: round 1,250 mg UP to a whole number of 500 mg vials with CEILING.MATH "
                  "(or ROUNDUP)",
             explanation=f"`CEILING.MATH({VANC_DOSE_MG},{VANC_VIAL_MG})` rounds up to {ceiling_math(VANC_DOSE_MG, VANC_VIAL_MG):,} mg, "
                         f"which is {vials_per_dose} vials per dose. Two vials would hold only 1,000 mg, so ROUNDDOWN would "
                         f"under-dose. {vials_per_dose} vials × {v_doses} doses = {v_vials} vials. Rounding the order's total "
                         f"instead (`CEILING.MATH({VANC_DOSE_MG}*{v_doses},{VANC_VIAL_MG})/{VANC_VIAL_MG}` = {v_wrong}) assumes "
                         "leftover drug carries over to the next dose, which a single-dose vial doesn't allow. "
                         f"`=ROUNDUP({VANC_DOSE_MG}/{VANC_VIAL_MG},0)*Meds!{med.col('DosesDispensed')}{v_row}` also works. The order wasted {v_waste:,} mg in "
                         "partly used vials."),
        Task(f"Case managers call a stay \"on target\" when its LOSDays is within {ON_TARGET_DAYS} day of its ExpectedLOS in "
             f"either direction (a difference of {ON_TARGET_DAYS}.0 day or less, longer or shorter). How many of the "
             f"{len(stays)} stays were on target? Use one formula.",
             answer=on_target,
             solution=f"=SUMPRODUCT(--(ABS({sr('LOSDays')}-{sr('ExpectedLOS')})<={ON_TARGET_DAYS}))",
             hint="ABS makes −0.9 and +0.9 the same distance. Count TRUE results with SUMPRODUCT(--(…)) from Lesson 2.1",
             explanation="Subtracting the two columns gives each stay's difference from expected, which is negative when the stay "
                         "was shorter than expected. ABS turns every difference into a distance, the comparison turns each "
                         "distance into TRUE or FALSE, `--` turns those into 1s and 0s, and SUMPRODUCT adds them. Without ABS, "
                         "every stay that ended "
                         f"early would count as on target, and you'd get {no_abs}. A helper column of `=ABS(I2-H2)` counted with "
                         f"`COUNTIF(range,\"<=1\")` gives the same {on_target}."),
        Task(f"The ED status board rounds each visit's length of stay to the nearest {BOARD_ROUND_MIN} minutes and shows it as "
             f"hours and minutes, like \"6 h 45 min\". What does it show for visit {b['EDVisitID']} (ED row {b_row}, "
             f"EDLOSMin = {b_min})? Build the text with one formula that refers to the EDLOSMin cell.",
             answer=b_text, accept=b_accept, check="text",
             solution=(f'=INT(MROUND(ED!{edd.col("EDLOSMin")}{b_row},{BOARD_ROUND_MIN})/60)&" h "&'
                       f'MOD(MROUND(ED!{edd.col("EDLOSMin")}{b_row},{BOARD_ROUND_MIN}),60)&" min"'),
             hint="MROUND to 15 first. Then INT(minutes/60) gives hours and MOD(minutes,60) gives the minutes left over. Join with &",
             explanation=f"`MROUND({b_min},{BOARD_ROUND_MIN})` returns {b_round}, because {b_min} is closer to {b_round} than to "
                         f"{b_round - BOARD_ROUND_MIN}. Then `INT({b_round}/60)` is {hh} whole hours and `MOD({b_round},60)` is the "
                         f"{mm} minutes left over. The & operator joins the pieces into \"{b_text}\". Without the rounding the board "
                         f"would show \"{b_raw_text}\". `TRUNC` or `ROUNDDOWN(…,0)` can replace INT here because the minutes are "
                         "positive."),
    ]

    # ------------------------------------------------------------------ bonus
    q3f = f"QUARTILE.INC({J},3)"
    q1f = f"QUARTILE.INC({J},1)"
    fence_f = f"{q3f}+1.5*({q3f}-{q1f})"
    L.bonus_title = "Bonus: Which stays are driving the average charge?"
    L.bonus_scenario = (
        f"The CFO's draft board report says the average inpatient charge in Q4 was ${ch_mean:,.0f}. A board member asks "
        "whether that number is typical and which stays are unusually expensive. Flag the high-charge outliers with two common "
        "rules, measure how much they matter, and find a more representative average. Use the TotalCharges column on the "
        f"Stays sheet ({J}). Your B1 answer lands in cell D6 of this sheet, so later parts can refer to it.")
    L.bonus = [
        Task("Calculate the upper outlier fence with the IQR rule: Q3 + 1.5 × (Q3 − Q1), using QUARTILE.INC. Enter it in "
             "dollars, to the cent.",
             answer=round(fence, 2), fmt="#,##0.00",
             solution=f"={fence_f}",
             hint="Find Q1 and Q3 with QUARTILE.INC, then combine them",
             explanation=f"Q1 is ${q1:,.2f} and Q3 is ${q3:,.2f}, so the IQR is ${iqr:,.2f} and the fence is "
                         f"${q3:,.2f} + 1.5 × ${iqr:,.2f} = ${fence:,.2f}. The lower fence, Q1 − 1.5 × IQR, is "
                         f"{'−' if low_fence < 0 else ''}${abs(low_fence):,.2f}. That's below zero, so the rule can't flag "
                         "any low outliers. Right-skewed data like charges usually has outliers on the high side only."),
        Task("How many stays have TotalCharges above the fence from B1?",
             answer=high_n,
             solution=f'=COUNTIF({J},">"&D6)',
             live=f'=COUNTIF({J},">"&({fence_f}))',
             hint="COUNTIF(range, \">\"&cell) joins the operator to your B1 cell",
             explanation=f"`\">\"&D6` builds the criterion text \">{fence:.2f}\", so COUNTIF counts charges above your fence. "
                         f"That's {high_n} stays, or {high_n / len(stays):.1%} of all stays."),
        Task("What share of all Q4 TotalCharges came from the stays above the fence? Enter it as a percentage with one "
             "decimal place.",
             answer=high_share, fmt="0.0%",
             solution=f"=SUMPRODUCT(({J}>D6)*{J})/SUM({J})",
             live=f"=SUMPRODUCT(({J}>({fence_f}))*{J})/SUM({J})",
             hint="(range>D6) is TRUE or FALSE for each stay. Multiply it by the charges inside SUMPRODUCT, then divide by "
                  "the total",
             explanation=f"`({J}>D6)` is TRUE for an outlier and FALSE otherwise. Multiplying by the charges turns TRUE into 1 "
                         "and FALSE into 0, so only the outlier charges survive, and SUMPRODUCT adds them. "
                         f"Just {high_n / len(stays):.1%} of stays produced {high_share:.1%} of all charges. That concentration is exactly why the mean sits so far above the median. "
                         f"(`=SUMIF({J},\">\"&D6)/SUM({J})` from Lesson 2.5 gives the same answer.)"),
        Task("A second common rule flags any value more than 3 standard deviations above the mean. How many stays have "
             "TotalCharges above AVERAGE + 3 × STDEV.S?",
             answer=sd3_n,
             solution=f'=COUNTIF({J},">"&(AVERAGE({J})+3*STDEV.S({J})))',
             hint="Build the threshold inside the criterion: \">\"&(AVERAGE(…)+3*STDEV.S(…))",
             explanation=f"The threshold is ${ch_mean:,.2f} + 3 × ${ch_sd:,.2f} = ${sd3:,.2f}, and only {sd3_n} stays exceed it, "
                         f"compared with {high_n} under the IQR rule. The outliers themselves inflate the mean and the standard "
                         "deviation, which raises the bar they have to clear. This effect is called **masking**. The 3-SD rule "
                         "assumes roughly symmetric data. Quartiles barely move when a few extreme values change, so the IQR rule "
                         "is the better screen for skewed data like charges and length of stay."),
        Task(f"Calculate a {TRIM_PCT:.0%} trimmed mean of TotalCharges with TRIMMEAN, which drops about 5% of stays from each "
             "end before averaging. Enter it in dollars, to the cent.",
             answer=round(tmean, 2), fmt="#,##0.00",
             solution=f"=TRIMMEAN({J},{TRIM_PCT})",
             hint="TRIMMEAN(array, percent). The percent is the TOTAL share to drop",
             explanation=f"{len(stays)} × {TRIM_PCT:.0%} = {len(stays) * TRIM_PCT:.1f} stays. TRIMMEAN rounds that down to an even "
                         f"number ({2 * trim_k}) and drops {trim_k} from each end. The result, ${tmean:,.2f}, sits between the "
                         f"median (${ch_median:,.2f}) and the mean (${ch_mean:,.2f}). For the board, report the median as the "
                         "typical stay and keep the mean for budgeting, because mean × number of stays = total charges. Then list "
                         "the outlier stays separately so nobody mistakes them for the norm."),
    ]
    return L
