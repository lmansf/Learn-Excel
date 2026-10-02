"""Lesson 1.4 · Your First Formulas & Functions (reference lesson)."""
from __future__ import annotations

import math

from xlcourse import Lesson, Task, data

CODE = "1.4"


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="01-foundations", slug="04-basic-formulas",
        title="Your First Formulas & Functions", level="Beginner", minutes=45,
        objectives=[
            "Write formulas with operators and the correct order of operations",
            "Use SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, and COUNTBLANK",
            "Calculate rates and percentages such as bed occupancy",
            "Copy formulas down a column and read common errors (#DIV/0!, #NAME?, #VALUE!)",
            "Round results with ROUND",
        ],
        data_note="Daily census for Medical-Surgical 4 West at Bluestone Memorial Hospital, Jan–Mar 2025 (90 days), "
                  "plus the unit's supply room stock.",
    )

    # ------------------------------------------------------------------ data
    census = [r for r in data.load("daily_census")
              if r["DeptID"] == "D110" and r["CensusDate"].year == 2025 and r["CensusDate"].month <= 3]
    census.sort(key=lambda r: r["CensusDate"])
    # Charge-nurse notes: a few operational comments (deterministic), blank on most days.
    for i, r in enumerate(census):
        note = None
        if r["MidnightCensus"] > r["StaffedBeds"]:
            note = "Over capacity: surge beds opened"
        elif r["MidnightCensus"] == r["StaffedBeds"]:
            note = "Unit full: diverting new admits to 5 East"
        elif i % 9 == 4:
            note = "1 RN call-off; covered by float pool"
        elif i % 13 == 7:
            note = "Isolation room deep-clean"
        r["Notes"] = note
    cen = L.add_table_sheet(
        "Census", census, table="tblCensus",
        columns=["CensusDate", "StaffedBeds", "Admissions", "Discharges", "MidnightCensus", "Notes"],
        extra_cols=["Occupancy"], formats={"Occupancy": "0.0%"}, widths={"Notes": 38, "Occupancy": 12},
    )
    supplies = [r for r in data.load("supply_inventory") if r["LocationDeptID"] == "D110"]
    supplies.sort(key=lambda r: r["SKU"])
    sup = L.add_table_sheet(
        "Supplies", supplies, table="tblSupplies",
        columns=["SKU", "ItemDescription", "Category", "UnitOfMeasure", "UnitCost", "QtyOnHand"],
        extra_cols=["StockValue"], formats={"StockValue": "#,##0.00", "UnitCost": "#,##0.00"},
        widths={"ItemDescription": 44, "StockValue": 13},
    )

    # ------------------------------------------------------------------ answers (computed in Python)
    adm = [r["Admissions"] for r in census]
    dis = [r["Discharges"] for r in census]
    mc = [r["MidnightCensus"] for r in census]
    beds = [r["StaffedBeds"] for r in census]
    notes = [r["Notes"] for r in census]
    n = len(census)
    first, last = cen.first_row, cen.last_row

    def rng(col):
        return f"Census!{cen.col(col)}{first}:{cen.col(col)}{last}"

    occ_days = sum(1 for c, b in zip(mc, beds) if c / b >= 1)
    feb = [r for r in census if r["CensusDate"].month == 2]
    feb_first = first + census.index(feb[0])
    feb_last = feb_first + len(feb) - 1
    feb_adc = sum(r["MidnightCensus"] for r in feb) / len(feb)
    stock_value = sum(r["UnitCost"] * r["QtyOnHand"] for r in supplies)
    sfirst, slast = sup.first_row, sup.last_row

    L.practice_intro = ("All tasks use the Census sheet (4 West, Q1 2025) unless they say otherwise. "
                        "Use cell ranges like Census!C2:C91 — or click and drag to select them while typing a formula.")
    L.tasks = [
        Task("How many patients were admitted to 4 West during the quarter (total of the Admissions column)?",
             answer=sum(adm), solution=f"=SUM({rng('Admissions')})", hint="SUM",
             explanation="SUM adds every number in the range. Typing =SUM( and then dragging over the column fills in the range for you."),
        Task("What was the average midnight census? (Keep full precision — the check accepts 2 decimal places.)",
             answer=sum(mc) / n, fmt="0.00", solution=f"=AVERAGE({rng('MidnightCensus')})", hint="AVERAGE",
             explanation="AVERAGE = SUM ÷ COUNT of the numbers in the range. This is the unit's *average daily census* (ADC)."),
        Task("What was the highest midnight census on any day?", answer=max(mc), solution=f"=MAX({rng('MidnightCensus')})", hint="MAX"),
        Task("What was the lowest midnight census on any day?", answer=min(mc), solution=f"=MIN({rng('MidnightCensus')})", hint="MIN"),
        Task("How many days of census data are there? Count the dates in the CensusDate column.",
             answer=n, solution=f"=COUNT({rng('CensusDate')})", hint="COUNT counts numbers — and dates are numbers",
             explanation="Excel stores dates as numbers (serial numbers), so COUNT includes them. COUNT ignores text and blanks."),
        Task("On how many days did the charge nurse write a note?", answer=sum(1 for x in notes if x),
             solution=f"=COUNTA({rng('Notes')})", hint="COUNTA counts anything that isn't empty",
             explanation="COUNTA counts every non-empty cell, including text. COUNT would return 0 here because the notes are text."),
        Task("On how many days was the Notes cell left empty?", answer=sum(1 for x in notes if not x),
             solution=f"=COUNTBLANK({rng('Notes')})", hint="COUNTBLANK",
             explanation="COUNTA + COUNTBLANK always equals the number of cells in the range (here 90)."),
        Task("Patient days = the sum of every day's midnight census. How many patient days did 4 West provide in Q1?",
             answer=sum(mc), solution=f"=SUM({rng('MidnightCensus')})", hint="It's a SUM",
             explanation="Each patient in a bed at midnight counts as one patient day. Patient days drive staffing and cost per day."),
        Task("What was the unit's occupancy rate for the whole quarter? Divide total patient days by total staffed-bed days. "
             "Enter it as a percentage.", answer=sum(mc) / sum(beds), fmt="0.0%",
             solution=f"=SUM({rng('MidnightCensus')})/SUM({rng('StaffedBeds')})", hint="SUM(…)/SUM(…), then format as %",
             explanation="A rate for a whole period should be total numerator ÷ total denominator. Format the cell as a percentage (Ctrl+Shift+%) "
                         "to see 91.3% instead of 0.9126…"),
        Task(f"In the Census sheet, fill the yellow Occupancy column with a formula for each day (MidnightCensus ÷ StaffedBeds). "
             f"The gray cell counts the days your column shows 100% or more. (Type it in {cen.cell('Occupancy', 0, sheet=False)}, "
             f"then double-click the fill handle to copy it down.)",
             answer=occ_days, title="Daily occupancy column (days at or over 100%)",
             solution=f"={cen.col('MidnightCensus')}{first}/{cen.col('StaffedBeds')}{first}",
             summary=f'=IF(COUNT({rng("Occupancy")})=0,"",COUNTIF({rng("Occupancy")},">=1"))',
             fill={"range": f"Census!{cen.col('Occupancy')}{first}:{cen.col('Occupancy')}{last}",
                   "formula": f"={cen.col('MidnightCensus')}{first}/{cen.col('StaffedBeds')}{first}"},
             live=f'=SUMPRODUCT(--({rng("MidnightCensus")}>={rng("StaffedBeds")}))',
             hint="Relative references shift down as you copy",
             explanation=f"Type =E{first}/B{first} in the first Occupancy cell and copy it down; each row's formula points to its own row "
                         "(=E3/B3, =E4/B4 …). Because the data is an Excel Table, typing the formula in one cell may fill the whole column "
                         "automatically — and you may see it written as =[@MidnightCensus]/[@StaffedBeds]. Both are correct."),
        Task("Average length of stay (ALOS) = patient days ÷ discharges. Calculate it for the quarter, rounded to 1 decimal place with ROUND.",
             answer=round(sum(mc) / sum(dis), 1), fmt="0.0",
             solution=f"=ROUND(SUM({rng('MidnightCensus')})/SUM({rng('Discharges')}),1)", hint="ROUND(number, 1)",
             tol=0.0001,
             explanation="ROUND changes the stored value, not just how it looks. The check here is strict: an unrounded value won't match."),
        Task("Without typing it into Excel first, what does =(30-6)/4+2^3 return? Then type it to confirm.",
             answer=(30 - 6) / 4 + 2 ** 3, solution="=(30-6)/4+2^3", hint="Parentheses → exponents → × ÷ → + −",
             explanation="(30−6)=24 first (parentheses), then 2^3=8 (exponent), then 24/4=6 (division), then 6+8=14."),
        Task("Switch to the Supplies sheet. Fill the yellow StockValue column with UnitCost × QtyOnHand for every item. "
             "The gray cell totals your column — what is the total value of 4 West's supply room?",
             answer=round(stock_value, 2), fmt="#,##0.00", title="StockValue column (total supply value)",
             solution=f"={sup.col('UnitCost')}{sfirst}*{sup.col('QtyOnHand')}{sfirst}",
             summary=f'=IF(COUNT(Supplies!{sup.col("StockValue")}{sfirst}:{sup.col("StockValue")}{slast})=0,"",'
                     f'SUM(Supplies!{sup.col("StockValue")}{sfirst}:{sup.col("StockValue")}{slast}))',
             fill={"range": f"Supplies!{sup.col('StockValue')}{sfirst}:{sup.col('StockValue')}{slast}",
                   "formula": f"={sup.col('UnitCost')}{sfirst}*{sup.col('QtyOnHand')}{sfirst}"},
             live=f"=SUMPRODUCT(Supplies!{sup.col('UnitCost')}{sfirst}:{sup.col('UnitCost')}{slast},"
                  f"Supplies!{sup.col('QtyOnHand')}{sfirst}:{sup.col('QtyOnHand')}{slast})",
             hint="Multiply with *; copy down",
             explanation="One formula, copied down, gives each item's value; then SUM the column. (The live formula in the key uses SUMPRODUCT, "
                         "which multiplies and adds in one step — you'll meet it in Lesson 2.4.)"),
    ]

    # ------------------------------------------------------------------ bonus
    L.bonus_title = "Bonus: Does 4 West need more beds?"
    L.bonus_scenario = ("The Chief Nursing Officer says February felt 'impossibly full' on 4 West and asks whether the unit needs more staffed "
                        "beds. Hospitals often plan beds so average occupancy is about 85% — enough slack to absorb surges. "
                        f"February is rows {feb_first}–{feb_last} of the Census sheet.")
    L.bonus = [
        Task("What was February's average daily census (ADC)? Use only the February rows.", answer=feb_adc, fmt="0.00",
             solution=f"=AVERAGE(Census!E{feb_first}:E{feb_last})", hint=f"AVERAGE over rows {feb_first}–{feb_last}",
             explanation="Select only the February rows. Selecting the whole column would give the quarterly ADC instead."),
        Task("How many beds would 4 West have needed in February for that ADC to equal 85% occupancy? (ADC ÷ 0.85, keep the decimals.)",
             answer=feb_adc / 0.85, fmt="0.00", solution=f"=AVERAGE(Census!E{feb_first}:E{feb_last})/0.85",
             hint="If ADC is 85% of the beds, beds = ADC ÷ 0.85"),
        Task("You can't open part of a bed. How many EXTRA beds (beyond today's 36) should the unit open? "
             "Round up to a whole bed.",
             answer=math.ceil(feb_adc / 0.85 - census[0]["StaffedBeds"]),
             solution=f"=ROUNDUP(AVERAGE(Census!E{feb_first}:E{feb_last})/0.85-Census!B{feb_first},0)",
             hint="ROUNDUP works like ROUND but always rounds up: ROUNDUP(number, 0)",
             explanation="Beds needed (≈38.6) minus the 36 staffed beds ≈ 2.6, rounded up to 3. ROUND would have given 3 here as well, "
                         "but only ROUNDUP guarantees you never under-build: 2.1 extra beds still means opening 3."),
        Task("Sanity check: on how many February days did the midnight census exceed 85% of 36 beds (i.e. more than 30.6 patients)?",
             answer=sum(1 for r in feb if r["MidnightCensus"] > 0.85 * r["StaffedBeds"]),
             solution=f'=COUNTIF(Census!E{feb_first}:E{feb_last},">"&0.85*36)', hint="COUNTIF(range, \">30.6\") — a preview of Lesson 2.5",
             explanation="COUNTIF counts cells that meet a condition. You'll master it in Lesson 2.5 — here it confirms that almost every "
                         "February day ran above the 85% planning target."),
    ]
    return L
