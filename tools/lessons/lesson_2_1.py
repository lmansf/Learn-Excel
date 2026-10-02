"""Lesson 2.1 · Logical Functions: IF, AND, OR, IFS & More."""
from __future__ import annotations

from xlcourse import Lesson, Task, data

CODE = "2.1"

# Clinical thresholds used throughout the lesson (simplified, educational — not clinical guidance).
FEVER_F = 100.4        # TempF above this = fever
HYPO_F = 96.8          # TempF below this = low temperature
TACHY_HR = 90          # HeartRate above this = tachycardia
TACHYPNEA_RR = 20      # RespRate above this = fast breathing (SIRS)
QSOFA_RR = 22          # RespRate at or above this = 1 qSOFA point
QSOFA_SBP = 100        # SystolicBP at or below this = 1 qSOFA point
FAST_TRACK_EXCLUDE = ("Chest Pain", "Shortness of Breath")
ESI_NAMES = {1: "Resuscitation", 2: "Emergent", 3: "Urgent", 4: "Less Urgent", 5: "Non-Urgent"}


def _pain_level(score: int) -> str:
    if score == 0:
        return "None"
    if score <= 3:
        return "Mild"
    if score <= 6:
        return "Moderate"
    return "Severe"


def _altered(complaint: str) -> bool:
    c = complaint.lower()
    return "altered" in c or "confusion" in c


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="02-formulas-functions", slug="01-logical-functions",
        title="Logical Functions: IF, AND, OR, IFS & More", level="Beginner → Intermediate", minutes=50,
        objectives=[
            "Compare values to produce TRUE/FALSE and use booleans in math",
            "Make decisions with IF, nested IF, and IFS",
            "Combine conditions with AND, OR, NOT (and XOR)",
            "Map codes to labels with SWITCH; trap errors with IFERROR and IFNA",
        ],
        data_note="391 lab results collected December 1–4, 2025 across Bluestone Health, and all 423 emergency department "
                  "visits at Bluestone Memorial Hospital in December 2025 with triage vital signs. Triage blood pressure "
                  "was not recorded for 12 of those visits.",
    )
    L.start_notes.append("The clinical rules in this lesson (lab flags, SIRS, qSOFA, shock index, fast track) are simplified "
                         "for teaching spreadsheet logic. They are not clinical guidance.")

    # ------------------------------------------------------------------ data: lab results
    labs = [r for r in data.load("lab_results")
            if r["CollectedDateTime"].year == 2025 and r["CollectedDateTime"].month == 12 and r["CollectedDateTime"].day <= 4]
    labs.sort(key=lambda r: (r["CollectedDateTime"], r["LabResultID"]))
    lab = L.add_table_sheet(
        "Labs", labs, table="tblLabs",
        columns=["LabResultID", "PatientID", "CollectedDateTime", "TestName", "ResultValue", "Units", "RefLow", "RefHigh",
                 "Priority"],
        extra_cols=["Flag"], formats={"ResultValue": "General", "RefLow": "General", "RefHigh": "General"},
        widths={"TestName": 30, "Flag": 11, "CollectedDateTime": 17},
    )

    # ------------------------------------------------------------------ data: ED visits
    ed = [r for r in data.load("ed_visits")
          if r["FacilityID"] == "F01" and r["ArrivalDateTime"].year == 2025 and r["ArrivalDateTime"].month == 12]
    ed.sort(key=lambda r: (r["ArrivalDateTime"], r["EDVisitID"]))
    # Triage BP "not obtained" on a few visits (deterministic): blank cells drive the IFERROR task and the
    # blank-cell pitfall in the bonus.
    for i, r in enumerate(ed):
        if i % 37 == 11:
            r["SystolicBP"] = None
    edd = L.add_table_sheet(
        "ED", ed, table="tblED",
        columns=["EDVisitID", "ArrivalDateTime", "ArrivalMode", "ESILevel", "ChiefComplaint", "TempF", "HeartRate",
                 "RespRate", "SystolicBP", "SpO2", "PainScore", "EDDisposition"],
        extra_cols=["PainLevel", "FastTrack", "SIRS", "ESIName", "ShockIndex", "qSOFA"],
        formats={"TempF": "0.0", "ShockIndex": "0.00"},
        widths={"ChiefComplaint": 30, "ArrivalDateTime": 17, "PainLevel": 11, "FastTrack": 11, "ShockIndex": 11,
                "ESIName": 13},
    )

    # ------------------------------------------------------------------ helpers
    lf, ll = lab.first_row, lab.last_row
    ef, el = edd.first_row, edd.last_row

    def lr(col):  # Labs range, relative (as a learner types it)
        return f"Labs!{lab.col(col)}{lf}:{lab.col(col)}{ll}"

    def er(col):  # ED range
        return f"ED!{edd.col(col)}{ef}:{edd.col(col)}{el}"

    def lc(col):  # first-row cell on Labs (no sheet name, for fill formulas)
        return f"{lab.col(col)}{lf}"

    def ec(col):
        return f"{edd.col(col)}{ef}"

    def summary_text(rng, value):
        return f'=IF(COUNTA({rng})=0,"",COUNTIF({rng},"{value}"))'

    # ------------------------------------------------------------------ answers (computed in Python)
    # Task 1: a result that sits exactly on its upper limit (equal is not "greater than").
    t1_i = next(i for i, r in enumerate(labs) if r["ResultValue"] == r["RefHigh"] and r["Units"] != "ratio")
    t1 = labs[t1_i]
    t1_row = lf + t1_i
    t1_answer = t1["ResultValue"] > t1["RefHigh"]

    above_high = sum(r["ResultValue"] > r["RefHigh"] for r in labs)

    def lab_flag(r):
        if r["ResultValue"] > r["RefHigh"]:
            return "High"
        if r["ResultValue"] < r["RefLow"]:
            return "Low"
        return "Normal"
    normal_n = sum(lab_flag(r) == "Normal" for r in labs)
    on_limit_n = sum(r["ResultValue"] in (r["RefLow"], r["RefHigh"]) for r in labs)

    moderate_n = sum(_pain_level(r["PainScore"]) == "Moderate" for r in ed)

    fever = [r["TempF"] > FEVER_F for r in ed]
    hypo = [r["TempF"] < HYPO_F for r in ed]
    tachy = [r["HeartRate"] > TACHY_HR for r in ed]
    tachyp = [r["RespRate"] > TACHYPNEA_RR for r in ed]
    temp_abn_n = sum(f or h for f, h in zip(fever, hypo))
    fev_tach_n = sum(f and t for f, t in zip(fever, tachy))
    xor_n = sum(f != t for f, t in zip(fever, tachy))
    assert xor_n == sum(fever) + sum(tachy) - 2 * fev_tach_n

    def fast_track(r):
        return r["ESILevel"] >= 4 and r["ArrivalMode"] == "Walk-In" and r["ChiefComplaint"] not in FAST_TRACK_EXCLUDE
    fast_n = sum(fast_track(r) for r in ed)

    sirs = [int(f or h) + int(t) + int(p) for f, h, t, p in zip(fever, hypo, tachy, tachyp)]
    sirs_total = sum(sirs)
    sirs_pos_n = sum(s >= 2 for s in sirs)

    emergent_n = sum(ESI_NAMES[r["ESILevel"]] == "Emergent" for r in ed)

    blank_bp = [i for i, r in enumerate(ed) if r["SystolicBP"] is None]
    shock_n = sum(1 for r in ed if r["SystolicBP"] is not None and r["HeartRate"] / r["SystolicBP"] >= 1)
    t13_i = blank_bp[0]
    t13 = ed[t13_i]
    t13_row = ef + t13_i

    # ------------------------------------------------------------------ live (stand-alone) formulas for the key
    sirs_live_parts = (f"(({er('TempF')}>{FEVER_F})+({er('TempF')}<{HYPO_F}))+({er('HeartRate')}>{TACHY_HR})"
                       f"+({er('RespRate')}>{TACHYPNEA_RR})")

    L.practice_intro = ("Tasks 1–3 use the Labs sheet. Tasks 4–13 use the ED sheet. Several tasks ask you to fill one of the "
                        "yellow columns on a data sheet: type the formula in the first data row and Excel fills the rest of "
                        "the Table column for you (if it doesn't, double-click the fill handle). The gray cell on this sheet "
                        "then summarizes your column. When a formula on this sheet points at a data sheet, include the sheet "
                        "name (for example Labs!E38), or click the cell on that sheet and Excel adds the name for you.")

    L.tasks = [
        # ---------------- comparisons & booleans
        Task(f"Labs row {t1_row} holds a {t1['TestName']} result of {t1['ResultValue']:g} {t1['Units']}. Write a comparison "
             f"that returns TRUE if that ResultValue is greater than its RefHigh, and FALSE if not.",
             answer=t1_answer,
             solution=f"=Labs!{lab.col('ResultValue')}{t1_row}>Labs!{lab.col('RefHigh')}{t1_row}",
             hint="A comparison needs no function: just =cell>cell. Include the sheet name (Labs!) in both references",
             explanation=f"A comparison always returns TRUE or FALSE. Here the result ({t1['ResultValue']:g}) is exactly equal to the "
                         f"upper limit ({t1['RefHigh']:g}), and *equal* is not *greater than*, so the answer is FALSE. With `>=` it "
                         "would be TRUE. Choosing between `>` and `>=` is a definition decision, so always check the rule you were given."),
        Task("How many of the lab results are above their reference high (ResultValue greater than RefHigh)? "
             "Use one formula that turns every row's comparison into a 1 or 0 and adds them up.",
             answer=above_high,
             solution=f"=SUMPRODUCT(--({lr('ResultValue')}>{lr('RefHigh')}))",
             hint="-- turns TRUE/FALSE into 1/0, and SUMPRODUCT adds them up",
             explanation="Comparing two ranges row by row gives a list of TRUE/FALSE values. The double minus (`--`) converts them to "
                         "1s and 0s, and SUMPRODUCT adds them. In Microsoft 365 and Excel 2021 or later, `=SUM(--(…>…))` works too. Plain `SUM` over a "
                         "column of TRUE/FALSE *cells* returns 0, because SUM ignores logical values stored in cells."),
        Task(f"Fill the yellow Flag column on the Labs sheet with a nested IF: \"High\" if ResultValue is greater than RefHigh, "
             f"\"Low\" if it is less than RefLow, otherwise \"Normal\". A result exactly on a limit is Normal. Start in "
             f"{lab.cell('Flag', 0, sheet=False)}. The gray cell counts your \"Normal\" flags.",
             answer=normal_n, title="Flag column with nested IF (count of Normal)",
             solution=f'=IF({lc("ResultValue")}>{lc("RefHigh")},"High",IF({lc("ResultValue")}<{lc("RefLow")},"Low","Normal"))',
             summary=summary_text(lab.rng("Flag"), "Normal"),
             fill={"range": f"Labs!{lab.col('Flag')}{lf}:{lab.col('Flag')}{ll}",
                   "formula": f'=IF({lc("ResultValue")}>{lc("RefHigh")},"High",IF({lc("ResultValue")}<{lc("RefLow")},"Low","Normal"))'},
             live=f"=SUMPRODUCT(({lr('ResultValue')}>={lr('RefLow')})*({lr('ResultValue')}<={lr('RefHigh')}))",
             hint="IF(test, \"High\", IF(test, \"Low\", \"Normal\"))",
             explanation="The first IF handles High. Its *value_if_false* argument is a second IF that handles Low, and \"Normal\" is "
                         f"whatever is left. Because the tests use `>` and `<`, the {on_limit_n} results sitting exactly on a limit "
                         "fall through to Normal. In the Table you may see it as "
                         "`=IF([@ResultValue]>[@RefHigh],\"High\",IF([@ResultValue]<[@RefLow],\"Low\",\"Normal\"))`, which is the same formula."),
        Task(f"Fill the yellow PainLevel column on the ED sheet with IFS: PainScore 0 is \"None\", 1–3 is \"Mild\", 4–6 is "
             f"\"Moderate\", and 7–10 is \"Severe\". Start in {edd.cell('PainLevel', 0, sheet=False)}. The gray cell counts "
             f"your \"Moderate\" rows.",
             answer=moderate_n, title="PainLevel column with IFS (count of Moderate)",
             solution=f'=IFS({ec("PainScore")}=0,"None",{ec("PainScore")}<=3,"Mild",{ec("PainScore")}<=6,"Moderate",TRUE,"Severe")',
             summary=summary_text(edd.rng("PainLevel"), "Moderate"),
             fill={"range": f"ED!{edd.col('PainLevel')}{ef}:{edd.col('PainLevel')}{el}",
                   "formula": f'=IFS({ec("PainScore")}=0,"None",{ec("PainScore")}<=3,"Mild",{ec("PainScore")}<=6,"Moderate",TRUE,"Severe")'},
             live=f"=SUMPRODUCT(({er('PainScore')}>=4)*({er('PainScore')}<=6))",
             hint="Test from the lowest band up, and end with TRUE as the catch-all",
             explanation="IFS returns the result of the **first** test that is TRUE, so order matters. Testing `<=3` before `<=6` means a "
                         "score of 2 stops at \"Mild\" and never reaches \"Moderate\". The final `TRUE,\"Severe\"` pair is the catch-all for "
                         "everything left (7–10). Without it, any value that fails every test returns #N/A."),
        # ---------------- AND / OR / XOR / NOT
        Task(f"How many ED visits had an abnormal temperature: TempF above {FEVER_F} OR below {HYPO_F}?",
             answer=temp_abn_n,
             solution=f"=SUMPRODUCT(({er('TempF')}>{FEVER_F})+({er('TempF')}<{HYPO_F}))",
             hint="Adding two TRUE/FALSE lists works like OR when both can't be TRUE at once",
             explanation="`+` acts as OR over whole ranges, because each row scores 1 if either test is TRUE. That's safe here because no "
                         "temperature can be both above 100.4 and below 96.8. (When both tests *can* be TRUE, write `--((…)+(…)>0)` "
                         f"so a row counts once. The `--` is needed because `>0` turns the sums back into TRUE/FALSE.) `=COUNTIF({er('TempF')},\">{FEVER_F}\")+COUNTIF({er('TempF')},\"<{HYPO_F}\")` "
                         "gives the same answer."),
        Task(f"How many visits had BOTH a fever (TempF above {FEVER_F}) AND tachycardia (HeartRate above {TACHY_HR})?",
             answer=fev_tach_n,
             solution=f"=SUMPRODUCT(({er('TempF')}>{FEVER_F})*({er('HeartRate')}>{TACHY_HR}))",
             hint="Multiplying two TRUE/FALSE lists works like AND",
             explanation="`*` acts as AND over whole ranges. 1 × 1 = 1 only when both tests are TRUE, and anything × 0 = 0. "
                         "`=SUMPRODUCT(--AND(…))` does **not** work, because AND collapses the whole range into a single TRUE or FALSE "
                         "instead of testing row by row. To use the AND function itself, put `=AND(F2>100.4,G2>90)` in a helper "
                         "column and count the TRUEs with `COUNTIF(range,TRUE)`."),
        Task(f"The sepsis coordinator already reviews visits with both signs. How many visits had exactly ONE of the two: "
             f"a fever (TempF above {FEVER_F}) or tachycardia (HeartRate above {TACHY_HR}), but not both?",
             answer=xor_n,
             solution=f"=SUMPRODUCT(--(({er('TempF')}>{FEVER_F})<>({er('HeartRate')}>{TACHY_HR})))",
             hint="This is XOR logic. Use XOR in a helper column, or compare the two TRUE/FALSE lists with <>",
             explanation="\"Exactly one of two\" is XOR. `=XOR(F2>100.4,G2>90)` answers it for one row, so a helper column plus "
                         "`COUNTIF(range,TRUE)` works. Over whole ranges, `<>` is XOR: TRUE<>FALSE is TRUE, while TRUE<>TRUE and "
                         "FALSE<>FALSE are FALSE. Like AND and OR, the XOR function collapses a range into one value, so "
                         "`SUMPRODUCT(--XOR(range>…, range>…))` gives the wrong answer. Check: fever count + tachycardia count − 2 × both "
                         "= exactly one."),
        Task(f"Fill the yellow FastTrack column: \"Fast Track\" when ESILevel is 4 or 5 AND ArrivalMode is \"Walk-In\" AND the "
             f"ChiefComplaint is NOT \"Chest Pain\" or \"Shortness of Breath\" (it is neither one). Every other visit is \"Main ED\". Start in "
             f"{edd.cell('FastTrack', 0, sheet=False)}. The gray cell counts your \"Fast Track\" rows.",
             answer=fast_n, title="FastTrack column with IF + AND + OR + NOT (count of Fast Track)",
             solution=(f'=IF(AND({ec("ESILevel")}>=4,{ec("ArrivalMode")}="Walk-In",NOT(OR({ec("ChiefComplaint")}="Chest Pain",'
                       f'{ec("ChiefComplaint")}="Shortness of Breath"))),"Fast Track","Main ED")'),
             summary=summary_text(edd.rng("FastTrack"), "Fast Track"),
             fill={"range": f"ED!{edd.col('FastTrack')}{ef}:{edd.col('FastTrack')}{el}",
                   "formula": (f'=IF(AND({ec("ESILevel")}>=4,{ec("ArrivalMode")}="Walk-In",NOT(OR({ec("ChiefComplaint")}="Chest Pain",'
                               f'{ec("ChiefComplaint")}="Shortness of Breath"))),"Fast Track","Main ED")')},
             live=(f'=SUMPRODUCT(({er("ESILevel")}>=4)*({er("ArrivalMode")}="Walk-In")*({er("ChiefComplaint")}<>"Chest Pain")'
                   f'*({er("ChiefComplaint")}<>"Shortness of Breath"))'),
             hint="IF(AND(…, …, NOT(OR(…, …))), \"Fast Track\", \"Main ED\")",
             explanation="AND needs all three parts to be TRUE. The third part, `NOT(OR(complaint=\"Chest Pain\", complaint=\"Shortness of "
                         "Breath\"))`, is TRUE only when the complaint is neither one. `AND(…<>\"Chest Pain\", …<>\"Shortness of Breath\")` "
                         "is the same test written the other way round (De Morgan's law). `ESILevel>=4` covers 4 and 5 because ESI only "
                         "goes up to 5. Text comparisons ignore case, so \"walk-in\" also matches."),
        Task(f"Fill the yellow SIRS column with a vital-sign SIRS score from 0 to 3: one point for an abnormal temperature "
             f"(TempF above {FEVER_F} OR below {HYPO_F}), one for HeartRate above {TACHY_HR}, and one for RespRate above "
             f"{TACHYPNEA_RR}. Start in {edd.cell('SIRS', 0, sheet=False)}. The gray cell adds up your whole column.",
             answer=sirs_total, title="SIRS score column (total points)",
             solution=f"=OR({ec('TempF')}>{FEVER_F},{ec('TempF')}<{HYPO_F})+({ec('HeartRate')}>{TACHY_HR})+({ec('RespRate')}>{TACHYPNEA_RR})",
             summary=f'=IF(COUNTA({edd.rng("SIRS")})=0,"",SUM({edd.rng("SIRS")}))',
             fill={"range": f"ED!{edd.col('SIRS')}{ef}:{edd.col('SIRS')}{el}",
                   "formula": f"=OR({ec('TempF')}>{FEVER_F},{ec('TempF')}<{HYPO_F})+({ec('HeartRate')}>{TACHY_HR})+({ec('RespRate')}>{TACHYPNEA_RR})"},
             live=f"=SUMPRODUCT({sirs_live_parts})",
             hint="TRUE + TRUE = 2. Add three TRUE/FALSE tests together",
             explanation="When you do arithmetic on TRUE/FALSE, Excel treats TRUE as 1 and FALSE as 0, so adding three tests gives a "
                         "score from 0 to 3 with no IF at all. The temperature point uses OR because *either* direction counts, but only "
                         "once. If your column shows TRUE/FALSE instead of numbers, you used AND/OR around everything. The full SIRS "
                         "criteria also include the white-blood-cell count, which isn't on this sheet."),
        Task("How many visits are SIRS-positive, meaning a score of 2 or more in your SIRS column?",
             answer=sirs_pos_n,
             solution=f'=COUNTIF({er("SIRS")},">=2")',
             live=f"=SUMPRODUCT(--(({sirs_live_parts})>=2))",
             hint="COUNTIF with a \">=2\" criterion",
             explanation="Once the score is a number, COUNTIF can count any threshold. The key's live formula does the whole thing in "
                         "one cell: it builds the score for every row as an array, then counts the rows where it is `>=2`."),
        # ---------------- SWITCH / IFERROR / IFNA
        Task(f"Fill the yellow ESIName column with SWITCH: ESILevel 1 is \"Resuscitation\", 2 is \"Emergent\", 3 is \"Urgent\", "
             f"4 is \"Less Urgent\", 5 is \"Non-Urgent\". Start in {edd.cell('ESIName', 0, sheet=False)}. The gray cell counts "
             f"your \"Emergent\" rows.",
             answer=emergent_n, title="ESIName column with SWITCH (count of Emergent)",
             solution=(f'=SWITCH({ec("ESILevel")},1,"Resuscitation",2,"Emergent",3,"Urgent",4,"Less Urgent",5,"Non-Urgent")'),
             summary=summary_text(edd.rng("ESIName"), "Emergent"),
             fill={"range": f"ED!{edd.col('ESIName')}{ef}:{edd.col('ESIName')}{el}",
                   "formula": f'=SWITCH({ec("ESILevel")},1,"Resuscitation",2,"Emergent",3,"Urgent",4,"Less Urgent",5,"Non-Urgent")'},
             live=f"=SUMPRODUCT(--({er('ESILevel')}=2))",
             hint="SWITCH(value, 1, \"…\", 2, \"…\", …)",
             explanation="SWITCH compares one value against a list and returns the result paired with the first exact match. It's "
                         "shorter than five nested IFs, and you write the cell reference only once. Leaving out a default is a deliberate "
                         "choice here. If a bad code like 6 ever appears, SWITCH returns #N/A and the problem stays visible."),
        Task(f"Shock index = HeartRate ÷ SystolicBP, and 1.0 or more is a warning sign. Fill the yellow ShockIndex column, "
             f"wrapping the division in IFERROR so the {len(blank_bp)} visits with no SystolicBP show a blank (\"\") instead of "
             f"#DIV/0!. Start in {edd.cell('ShockIndex', 0, sheet=False)}. The gray cell counts visits with a shock index of "
             f"1.0 or higher (it asks you to fix any errors first).",
             answer=shock_n, title="ShockIndex column with IFERROR (count of 1.0 or higher)",
             solution=f'=IFERROR({ec("HeartRate")}/{ec("SystolicBP")},"")',
             summary=(f'=IF(COUNTA({edd.rng("ShockIndex")})=0,"",IF(SUMPRODUCT(--ISERROR({edd.rng("ShockIndex")}))>0,'
                      f'"Fix the errors first",COUNTIF({edd.rng("ShockIndex")},">=1")))'),
             fill={"range": f"ED!{edd.col('ShockIndex')}{ef}:{edd.col('ShockIndex')}{el}",
                   "formula": f'=IFERROR({ec("HeartRate")}/{ec("SystolicBP")},"")'},
             live=f'=SUMPRODUCT(({er("SystolicBP")}<>"")*({er("HeartRate")}>={er("SystolicBP")}))',
             hint="IFERROR(value, value_if_error)",
             explanation="Dividing by an empty cell divides by zero, so those rows show #DIV/0!. `IFERROR(G2/I2,\"\")` returns the "
                         "division when it works and an empty text string when it fails. COUNTIF then skips the blanks because they "
                         "are text. `=IF(I2=\"\",\"\",G2/I2)` is an even better formula, because it handles the one problem you expect "
                         "(a missing BP) and still lets any *other* mistake show up as an error."),
        Task(f"Visit {t13['EDVisitID']} (ED row {t13_row}) has no SystolicBP. In the yellow cell, calculate its shock index again, "
             f"but wrap the division in IFNA(…, \"\") instead of IFERROR. The check turns green when your cell shows the error "
             f"that IFNA lets through.",
             answer=None, check="custom", custom_check='AND(IFERROR(ERROR.TYPE({cell})=2,FALSE),ISNUMBER(SEARCH("IFNA",FORMULATEXT({cell}))))',
             answer_display="#DIV/0! (IFNA traps only #N/A, so the divide-by-zero error passes through)",
             solution=f'=IFNA(ED!{edd.col("HeartRate")}{t13_row}/ED!{edd.col("SystolicBP")}{t13_row},"")',
             live=False, title="IFNA vs IFERROR experiment",
             hint="IFNA traps only one kind of error",
             explanation="IFERROR catches every error type. IFNA catches only #N/A, the \"value not available\" error that lookups return "
                         "when they can't find a match (Lesson 2.6). A divide-by-zero is a different error, so IFNA hands it straight "
                         "back. That's the point of IFNA. It hides the one error you expect and leaves every other mistake visible."),
    ]

    # ------------------------------------------------------------------ bonus: qSOFA-style screen
    def qsofa(r):
        rr = r["RespRate"] >= QSOFA_RR
        sbp = r["SystolicBP"] is not None and r["SystolicBP"] <= QSOFA_SBP
        return int(rr) + int(sbp) + int(_altered(r["ChiefComplaint"]))
    q = [qsofa(r) for r in ed]
    q_pos = [i for i, s in enumerate(q) if s >= 2]
    q_pos_n = len(q_pos)
    q_hr = sum(ed[i]["HeartRate"] for i in q_pos) / q_pos_n
    q_dc = sum(1 for i in q_pos if ed[i]["EDDisposition"] == "Discharged")
    q_not_sirs = sum(1 for i in q_pos if sirs[i] < 2)
    q_total = sum(q)
    # The blank-BP pitfall (a bare SystolicBP<=100 scores every missing BP) must change the answers, so the checks catch it.
    q_bad = [int(r["RespRate"] >= QSOFA_RR) + int(r["SystolicBP"] is None or r["SystolicBP"] <= QSOFA_SBP)
             + int(_altered(r["ChiefComplaint"])) for r in ed]
    q_bad_total, q_bad_pos = sum(q_bad), sum(1 for s in q_bad if s >= 2)
    assert q_bad_total == q_total + len(blank_bp) and q_bad_pos > q_pos_n and q_dc > 0 and q_not_sirs > 0

    qcol = edd.rng("qSOFA")
    q_first = (f'=({ec("RespRate")}>={QSOFA_RR})+AND({ec("SystolicBP")}<>"",{ec("SystolicBP")}<={QSOFA_SBP})'
               f'+OR(ISNUMBER(SEARCH("altered",{ec("ChiefComplaint")})),ISNUMBER(SEARCH("confusion",{ec("ChiefComplaint")})))')
    q_arr = (f'(({er("RespRate")}>={QSOFA_RR})+({er("SystolicBP")}<>"")*({er("SystolicBP")}<={QSOFA_SBP})'
             f'+ISNUMBER(SEARCH("altered",{er("ChiefComplaint")}))+ISNUMBER(SEARCH("confusion",{er("ChiefComplaint")})))')

    L.bonus_title = "Bonus: A qSOFA-style sepsis screen at triage"
    L.bonus_scenario = (
        "The sepsis committee wants to know what a quick qSOFA-style screen would flag if triage nurses scored every ED patient. "
        "This simplified, educational version (not clinical guidance) gives one point for each of three findings. "
        f"(1) RespRate is {QSOFA_RR} or more. (2) SystolicBP is {QSOFA_SBP} or less. A blank (not recorded) SystolicBP earns "
        "NO point. (3) Altered mentation, meaning the ChiefComplaint contains the word \"Altered\" or \"Confusion\" anywhere. "
        "A score of 2 or more is screen-positive. Build the score in the yellow qSOFA column on the ED sheet, then answer the questions.")
    L.bonus = [
        Task(f"Fill the qSOFA column (0–3) on the ED sheet, starting in {edd.cell('qSOFA', 0, sheet=False)}. The gray cell adds "
             f"up your whole column. What is the total number of qSOFA points?",
             answer=q_total, title="qSOFA column (total points)",
             solution=q_first,
             summary=f'=IF(COUNTA({qcol})=0,"",SUM({qcol}))',
             fill={"range": f"ED!{edd.col('qSOFA')}{ef}:{edd.col('qSOFA')}{el}", "formula": q_first},
             live=f"=SUMPRODUCT({q_arr})",
             hint="Add three tests. Guard the BP test with I2<>\"\". ISNUMBER(SEARCH(\"altered\", E2)) tests 'contains'",
             explanation="Each part is a TRUE/FALSE test, and adding them gives the score. Watch for two traps. First, an empty "
                         "cell counts as 0 in a comparison, so a bare `I2<=100` gives a point to every visit with no BP. "
                         f"`AND(I2<>\"\",I2<=100)` scores only a recorded BP. Without that guard the {len(blank_bp)} visits with no BP "
                         f"each gain a point, so the total is {q_bad_total} instead of {q_total}. "
                         "Second, SEARCH returns the position of the text, or #VALUE! when it isn't there, so `ISNUMBER(SEARCH(…))` turns "
                         "that into TRUE/FALSE. SEARCH ignores case, so \"altered\" finds \"Altered\". Wrapping the two SEARCH tests "
                         "in OR keeps the mentation point at 1 even if a complaint ever mentions both words."),
        Task("How many visits are screen-positive (a qSOFA score of 2 or more)?",
             answer=q_pos_n,
             solution=f'=COUNTIF({er("qSOFA")},">=2")',
             live=f"=SUMPRODUCT(--({q_arr}>=2))",
             hint="COUNTIF on your qSOFA column",
             explanation=f"Once the score is a number, COUNTIF counts any threshold. If your column scored the missing BPs, you'd "
                         f"count {q_bad_pos} here instead of {q_pos_n}, so the blank-cell guard changes a real quality number."),
        Task("What was the average HeartRate of the screen-positive visits? Round to 1 decimal place.",
             answer=round(q_hr, 1), fmt="0.0", tol=0.051,
             solution=f"=ROUND(AVERAGE(IF({er('qSOFA')}>=2,{er('HeartRate')})),1)",
             live=f"=ROUND(AVERAGE(IF({q_arr}>=2,{er('HeartRate')})),1)",
             hint="AVERAGE(IF(test_range>=2, values_range)). AVERAGEIF (Lesson 2.5) also works",
             explanation="`IF(range>=2, HeartRate)` returns the heart rate for screen-positive rows and FALSE for the rest, and AVERAGE "
                         "skips FALSE. Microsoft 365 and Excel 2021 evaluate this array formula automatically. In Excel 2019 and "
                         "earlier, confirm it with Ctrl + Shift + Enter (Mac: ⌘ + Shift + Return). "
                         f"`=AVERAGEIF({er('qSOFA')},\">=2\",{er('HeartRate')})` gives the same result."),
        Task("Safety check: how many screen-positive visits ended with EDDisposition \"Discharged\" (sent home)?",
             answer=q_dc,
             solution=f'=SUMPRODUCT(({er("qSOFA")}>=2)*({er("EDDisposition")}="Discharged"))',
             live=f'=SUMPRODUCT(({q_arr}>=2)*({er("EDDisposition")}="Discharged"))',
             hint="Two conditions over whole columns: multiply the TRUE/FALSE lists",
             explanation="Multiplying the two lists keeps a 1 only where both are TRUE (AND logic), and SUMPRODUCT counts them. "
                         "These are the charts a sepsis committee would pull for review, because the patients screened positive but went home."),
        Task("The two screens don't always agree. How many visits are qSOFA screen-positive (2 or more) but NOT SIRS-positive "
             "(SIRS score below 2 in your SIRS column from task 9)?",
             answer=q_not_sirs,
             solution=f'=SUMPRODUCT(({er("qSOFA")}>=2)*({er("SIRS")}<2))',
             live=f"=SUMPRODUCT(({q_arr}>=2)*(({sirs_live_parts})<2))",
             hint="Same pattern as B4. 'NOT SIRS-positive' is simply SIRS<2",
             explanation="NOT(score>=2) is the same as score<2, so you don't need the NOT function over a range. SIRS looks at "
                         "temperature, heart rate, and breathing. qSOFA looks at breathing, blood pressure, and mental status. Each "
                         "catches patients the other misses, which is why hospitals don't rely on one screen alone."),
    ]
    return L
