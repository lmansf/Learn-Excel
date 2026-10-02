"""Lesson 2.2 · Text Functions."""
from __future__ import annotations

from collections import Counter

from xlcourse import Lesson, Task, data

CODE = "2.2"
NBSP = " "
N_PATIENTS = 400


# ---------------------------------------------------------------------------
# Python models of the Excel text functions (so every answer is computed the way Excel computes it)
# ---------------------------------------------------------------------------
def xl_trim(s: str) -> str:
    """Excel TRIM: removes ASCII spaces (code 32) at both ends and collapses inner runs to one space.
    It does NOT touch non-breaking spaces (code 160) — that's why the lesson needs SUBSTITUTE."""
    return " ".join(part for part in s.split(" ") if part)


def xl_proper(s: str) -> str:
    """Excel PROPER: uppercase a letter that follows a non-letter (or starts the text); lowercase every other letter."""
    out, prev_letter = [], False
    for ch in s:
        if ch.isalpha():
            out.append(ch.lower() if prev_letter else ch.upper())
            prev_letter = True
        else:
            out.append(ch)
            prev_letter = False
    return "".join(out)


def first_part(name: str) -> str:
    """=TRIM(MID(SUBSTITUTE(name,UNICHAR(160)," "),FIND(",",name)+1,LEN(name)))"""
    return xl_trim(name.replace(NBSP, " ")[name.find(",") + 1:])


def last_part(name: str) -> str:
    """=TRIM(LEFT(name,FIND(",",name)-1))"""
    return xl_trim(name[:name.find(",")])


def messy_name(i: int, first: str, last: str) -> str:
    """Deterministic 'LAST, First' registration-system names with mixed case and spacing problems."""
    style = i % 5
    if style == 0:
        lst, fst = last.upper(), first          # ABBOTT, Edward   (typical EHR style)
    elif style == 2:
        lst, fst = last.upper(), first.upper()  # ABBOTT, EDWARD
    elif style == 3:
        lst, fst = last.lower(), first.lower()  # abbott, edward
    else:
        lst, fst = last, first                  # Abbott, Edward
    sep = " "
    if i % 23 == 12:
        sep = NBSP          # pasted from the patient portal
    elif i % 11 == 7:
        sep = ""            # no space after the comma
    elif i % 9 == 4:
        sep = "  "          # two spaces
    lead = " " if i % 13 == 6 else ""
    trail = "  " if i % 8 == 3 else (NBSP if i % 31 == 20 else "")
    return f"{lead}{lst},{sep}{fst}{trail}"


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="02-formulas-functions", slug="02-text-functions",
        title="Text Functions", level="Beginner → Intermediate", minutes=110,
        objectives=[
            "Extract parts of text with LEFT, RIGHT, MID, FIND, and SEARCH",
            "Clean and standardize text with TRIM, CLEAN, UPPER, LOWER, PROPER, and SUBSTITUTE",
            "Join text with &, CONCAT, and TEXTJOIN; format numbers as text with TEXT",
            "Use modern TEXTBEFORE, TEXTAFTER, and TEXTSPLIT (Microsoft 365 and Excel 2024)",
        ],
    )

    # ------------------------------------------------------------------ data: Patients
    patients = sorted(data.load("patients"), key=lambda r: r["PatientID"])[:N_PATIENTS]
    for i, r in enumerate(patients):
        r["MRNNumber"] = int(r["MRN"])  # the export stored MRNs as numbers, so leading zeros are gone
        r["PatientName"] = messy_name(i, r["FirstName"], r["LastName"])
        r["ChronicConditions"] = r["ChronicConditions"] or None
        r["Email"] = r["Email"] or None
    # Hand-picked showcase rows for the guide and the PROPER task (deterministic overrides).
    by_id = {r["PatientID"]: r for r in patients}
    by_id["PT10325"]["PatientName"] = "o'brien,hannah  "      # PROPER handles the apostrophe correctly
    pat = L.add_table_sheet(
        "Patients", patients, table="tblPatients",
        columns=["PatientID", ("MRNNumber", "MRN"), "PatientName", "Phone", "Email", "ChronicConditions"],
        extra_cols=["HasDM", "ContactLine", "FirstName", "ConditionCount"],  # in the order the tasks use them
        formats={"MRN": "General"},
        widths={"PatientName": 26, "Email": 34, "ChronicConditions": 22, "FirstName": 14, "ContactLine": 50,
                "HasDM": 10, "ConditionCount": 16},
    )

    # ------------------------------------------------------------------ data: Encounters (Q4 2025 for the same patients)
    ids = {r["PatientID"] for r in patients}
    dx = data.index(data.load("diagnoses"), "DxCode")
    encounters = [e for e in data.load("encounters")
                  if e["PatientID"] in ids and e["AdmitDateTime"].year == 2025 and e["AdmitDateTime"].month >= 10]
    encounters.sort(key=lambda e: (e["AdmitDateTime"], e["EncounterID"]))
    for e in encounters:
        e["AdmitDate"] = e["AdmitDateTime"].date()
        e["DxDescription"] = dx[e["PrimaryDxCode"]]["DxDescription"]
    enc = L.add_table_sheet(
        "Encounters", encounters, table="tblEncounters",
        columns=["EncounterID", "PatientID", "EncounterType", "AdmitDate", "PrimaryDxCode", "DxDescription"],
        extra_cols=["DxCategory"], widths={"DxDescription": 52, "DxCategory": 13},
    )

    # ------------------------------------------------------------------ data: Providers
    providers = sorted(data.load("providers"), key=lambda r: r["ProviderID"])
    prov = L.add_table_sheet(
        "Providers", providers, table="tblProviders",
        columns=["ProviderID", "FirstName", "LastName", "Credential", "Specialty"],
        widths={"Specialty": 28},
    )

    # ------------------------------------------------------------------ data: Registrations (legacy export, bonus)
    regs = data.load_raw("messy/patient_registrations_raw")
    reg = L.add_table_sheet(
        "Registrations", regs, table="tblRegistrations",
        columns=["RecordID", "PatientName", "CityStateZip"],
        extra_cols=["City", "ZIP5"], widths={"PatientName": 26, "CityStateZip": 32, "City": 18, "ZIP5": 9},
        tab_color="BF9000",
    )

    L.sheet_order = ["Start Here", "Practice", "Patients", "Encounters", "Providers", "Bonus", "Registrations"]

    # ------------------------------------------------------------------ helpers
    pf, pl = pat.first_row, pat.last_row
    ef, el = enc.first_row, enc.last_row
    rf, rl = reg.first_row, reg.last_row

    def prow(pid: str) -> int:
        return pf + next(i for i, r in enumerate(patients) if r["PatientID"] == pid)

    def pcell(col: str, pid: str) -> str:
        return f"Patients!{pat.col(col)}{prow(pid)}"

    def prng(col: str) -> str:
        return f"Patients!{pat.col(col)}{pf}:{pat.col(col)}{pl}"

    def erng(col: str) -> str:
        return f"Encounters!{enc.col(col)}{ef}:{enc.col(col)}{el}"

    def rrng(col: str) -> str:
        return f"Registrations!{reg.col(col)}{rf}:{reg.col(col)}{rl}"

    L.data_note = (f"{len(patients)} patients from Bluestone's registration system (names, MRNs, phones, emails, chronic-condition "
                   f"flags), their {len(encounters)} encounters in Q4 2025, the {len(providers)}-provider roster, and a "
                   f"{len(regs)}-row legacy registration export with messy addresses.")

    def pick(pred, start: int = 10):
        """First patient (index >= start) matching pred(i, row) — deterministic task targets."""
        return next(r for i, r in enumerate(patients) if i >= start and pred(i, r))

    # The README guide quotes these cells by address. If the data ever changes, fail loudly so the guide gets updated.
    def pv(row: int, key: str):
        return patients[row - pf][key]
    guide_refs = {
        ("Patients", 2, "PatientName"): (pv(2, "PatientName"), "ABBOTT, Edward"),
        ("Patients", 2, "Phone"): (pv(2, "Phone"), "(555) 875-0698"),
        ("Patients", 3, "MRN"): (pv(3, "MRNNumber"), 897724),
        ("Patients", 3, "Phone/Email"): ((pv(3, "Phone"), pv(3, "Email")), ("(555) 373-1790", None)),
        ("Patients", 5, "PatientName"): (pv(5, "PatientName"), "hamilton, george  "),
        ("Patients", 6, "PatientName"): (pv(6, "PatientName"), "Romero,  Albert"),
        ("Patients", 8, "PatientName"): (pv(8, "PatientName"), " Chavez, Carol"),
        ("Patients", 9, "PatientName"): (pv(9, "PatientName"), "CLARK,ARTHUR"),
        ("Patients", 12, "ChronicConditions"): (pv(12, "ChronicConditions"), "HTN;CKD;OA"),
        ("Patients", 14, "PatientName"): (pv(14, "PatientName"), f"ANDREWS,{NBSP}BRIAN"),
        ("Patients", 39, "PatientName"): (pv(39, "PatientName"), "O'BRIEN, TERRY"),
        ("Patients", 294, "PatientName"): (pv(294, "PatientName"), " MCDONALD,  VIRGINIA"),
        ("Encounters", 2, "DxDescription"): (encounters[0]["DxDescription"], "Chest pain, unspecified"),
        ("Encounters", 3, "PrimaryDxCode"): (encounters[1]["PrimaryDxCode"], "J44.1"),
        ("Encounters", 5, "PrimaryDxCode"): (encounters[3]["PrimaryDxCode"], "S93.401A"),
        ("Encounters", 128, "AdmitDate"): (encounters[126]["AdmitDate"].isoformat(), "2025-11-14"),
        ("Providers", 2, "Name"): ((providers[0]["FirstName"], providers[0]["LastName"], providers[0]["Credential"]),
                                   ("Roy", "Ferguson", "MD")),
    }
    for ref, (got, want) in guide_refs.items():
        assert got == want, f"README guide example {ref} changed: {got!r} != {want!r}; update the README"

    # ------------------------------------------------------------------ answers (computed in Python)
    # 1 · RIGHT: last 4 digits of a phone (pick one whose last 4 don't start with 0, so it can't be confused with a number)
    t_last4 = pick(lambda i, r: r["Phone"][-4] != "0", start=14)
    last4 = t_last4["Phone"][-4:]

    # 2 · ICD-10 category column → count of E11 (type 2 diabetes, any complication)
    e11 = sum(1 for e in encounters if e["PrimaryDxCode"][:3] == "E11")
    e11_codes = sorted({e["PrimaryDxCode"] for e in encounters if e["PrimaryDxCode"][:3] == "E11"})

    # 3 · last name with FIND + LEFT (an 'ABBOTT, Edward' style row with no spacing problems)
    t_last = pick(lambda i, r: i % 5 == 0 and r["PatientName"] == f"{r['LastName'].upper()}, {r['FirstName']}", start=30)
    last_answer = last_part(t_last["PatientName"])

    # 4 · digits-only phone
    t_phone = pick(lambda i, r: True, start=57)
    phone_digits = "".join(ch for ch in t_phone["Phone"] if ch.isdigit())

    # 5 · MRN as 8-character text. The guide's worked example is PT10002 (row 3), so pick a different patient,
    #     one who lost three leading zeros, so the learner applies the skill instead of copying the example.
    t_mrn = pick(lambda i, r: r["MRN"].startswith("000") and r["MRN"][3] != "0", start=0)
    assert prow(t_mrn["PatientID"]) != 3
    mrn_text = f"{t_mrn['MRNNumber']:08d}"
    assert mrn_text == t_mrn["MRN"]

    # 6 · month label with TEXT. The guide formats row 128 (Nov-2025), so use a December encounter here.
    t_enc = next(e for e in encounters if e["AdmitDate"].month == 12 and e["AdmitDate"].day >= 9)
    e_row = ef + encounters.index(t_enc)
    assert e_row != 128
    month_label = t_enc["AdmitDate"].strftime("%b-%Y")

    # 7 · provider badge label (an MD)
    t_prov = next(p for p in providers if p["Credential"] == "MD" and p["Specialty"] == "Cardiology")
    pr_row = prov.first_row + providers.index(t_prov)
    badge = f"Dr. {t_prov['FirstName']} {t_prov['LastName']}, {t_prov['Credential']}"
    ex_prov = next(p for p in providers if p["Credential"] == "MD" and p is not t_prov)

    # 8 · HasDM column → number of patients with diabetes on their problem list
    dm_count = sum(1 for r in patients if r["ChronicConditions"] and "DM" in r["ChronicConditions"].upper())

    # 9 · portal username = text before the @
    t_mail = pick(lambda i, r: bool(r["Email"]) and NBSP not in r["PatientName"], start=120)
    username = t_mail["Email"].split("@")[0]

    # 10 · number of chronic conditions with TEXTSPLIT
    t_cc = pick(lambda i, r: r["ChronicConditions"] is not None and len(r["ChronicConditions"].split(";")) == 4, start=0)
    cc_count = len(t_cc["ChronicConditions"].split(";"))

    # 11 · TEXTJOIN contact line → rows that contain the " | " separator (= patients with an email on file)
    with_email = sum(1 for r in patients if r["Email"])

    # 12 · first-name column: checksum = total characters of the cleaned first names
    first_chars = sum(len(first_part(r["PatientName"])) for r in patients)
    nbsp_rows = sum(1 for r in patients if NBSP in r["PatientName"])
    no_nbsp_chars = sum(len(xl_trim(r["PatientName"][r["PatientName"].find(",") + 1:])) for r in patients)
    assert no_nbsp_chars != first_chars  # forgetting the non-breaking space must fail the check

    # 13 · display name 'First Last' in Proper Case (case-sensitive check)
    t_disp = by_id["PT10325"]
    display = xl_proper(first_part(t_disp["PatientName"]) + " " + last_part(t_disp["PatientName"]))
    assert display == "Hannah O'Brien", display

    # ------------------------------------------------------------------ practice tasks
    C, Cc = pat.col("PatientName"), pat.col("ChronicConditions")
    L.practice_intro = ("Tasks use the Patients, Encounters, and Providers sheets. When a task names one patient, it also gives "
                        "the row, so you can point your formula at that row's cells (for example Patients!C25). Column tasks "
                        "ask you to fill a yellow column on a data sheet, and a gray cell on this sheet then summarizes your "
                        "column.")
    L.tasks = [
        Task(f"Front-desk staff confirm a caller's identity with the last 4 digits of their phone number. Return the last 4 "
             f"digits of patient {t_last4['PatientID']}'s phone (row {prow(t_last4['PatientID'])} of the Patients sheet).",
             answer=last4, solution=f"=RIGHT({pcell('Phone', t_last4['PatientID'])},4)",
             hint="RIGHT(text, num_chars)",
             explanation="RIGHT returns characters from the end of the text. The phone is stored as text in a fixed "
                         "(555) 123-4567 layout, so its last 4 characters are always the last 4 digits. The result is text: "
                         "a phone ending in 0698 would correctly keep its leading 0."),
        Task(f"On the Encounters sheet, fill the yellow DxCategory column with the ICD-10 category of each PrimaryDxCode "
             f"(its first 3 characters). The gray cell counts encounters in category E11 (type 2 diabetes). "
             f"Type your formula in {enc.cell('DxCategory', 0, sheet=False)}, then copy it down.",
             answer=e11, title="DxCategory column (encounters in category E11)",
             solution=f"=LEFT({enc.col('PrimaryDxCode')}{ef},3)",
             summary=f'=IF(COUNTA({erng("DxCategory")})=0,"",'
                     f'COUNTIF({erng("DxCategory")},"E11"))',
             fill={"range": f"Encounters!{enc.col('DxCategory')}{ef}:{enc.col('DxCategory')}{el}",
                   "formula": f"=LEFT({enc.col('PrimaryDxCode')}{ef},3)"},
             live=f'=SUMPRODUCT(--(LEFT({erng("PrimaryDxCode")},3)="E11"))',
             hint="LEFT(text, num_chars)",
             explanation=f"An ICD-10 code's first three characters are its category, so LEFT(code,3) groups "
                         f"{' and '.join(e11_codes)} together as E11. Codes have different lengths (I10, E11.65, S72.001A), "
                         "but the category is always the first 3 characters, which is why LEFT works with a fixed count here."),
        Task(f"Extract the last name of patient {t_last['PatientID']} (row {prow(t_last['PatientID'])} of the Patients sheet, "
             f"\"{t_last['PatientName']}\"): everything before the comma.",
             answer=last_answer, solution=f'=LEFT({pcell("PatientName", t_last["PatientID"])},'
                                          f'FIND(",",{pcell("PatientName", t_last["PatientID"])})-1)',
             hint="FIND gives the comma's position, and LEFT takes everything before it",
             explanation=f"FIND(\",\",…) returns the comma's position ({len(last_answer) + 1}). The last name is the "
                         f"{len(last_answer)} characters before it, so subtract 1. In Microsoft 365 and Excel 2024, you can also write "
                         f"=TEXTBEFORE({pcell('PatientName', t_last['PatientID'])},\",\")."),
        Task(f"The appointment-reminder system needs phone numbers as 10 digits with no punctuation. Convert patient "
             f"{t_phone['PatientID']}'s phone (row {prow(t_phone['PatientID'])} of the Patients sheet, \"{t_phone['Phone']}\") "
             f"to digits only.",
             answer=phone_digits,
             solution=f'=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({pcell("Phone", t_phone["PatientID"])},'
                      f'"(",""),")","")," ",""),"-","")',
             hint="Nest one SUBSTITUTE per character you want to remove",
             explanation="Each SUBSTITUTE replaces one character with nothing (\"\"). Nesting them removes the "
                         "parentheses, the space, and the dash in one formula. Work from the inside out: the innermost "
                         "SUBSTITUTE runs first and passes its result to the next one. The result is text, which is what you "
                         "want for phone numbers."),
        Task(f"Patient {t_mrn['PatientID']}'s MRN (row {prow(t_mrn['PatientID'])} of the Patients sheet) shows "
             f"{t_mrn['MRNNumber']} because the export stored it as a number and dropped the leading zeros. Return it as the "
             f"8-character text MRN printed on wristbands.",
             answer=mrn_text, solution=f'=TEXT({pcell("MRN", t_mrn["PatientID"])},"00000000")',
             hint="TEXT(value, format_text) with a format of eight 0s",
             explanation="In a format code, each 0 is a required digit, so \"00000000\" pads the number with leading zeros "
                         f"to 8 digits. The result \"{mrn_text}\" is text, so the zeros are part of the value: they survive "
                         "joining with & and match MRNs stored as text in other systems. A custom number format (Lesson 1.3) only "
                         f"changes how the cell looks, so =\"MRN \"&{pcell('MRN', t_mrn['PatientID'])} would still give "
                         f"MRN {t_mrn['MRNNumber']}."),
        Task(f"Monthly quality reports label encounters by month. Return the admit month of encounter {t_enc['EncounterID']} "
             f"(row {e_row} of the Encounters sheet) as text in the form Mon-YYYY, for example Jan-2026.",
             answer=month_label, solution=f'=TEXT(Encounters!{enc.col("AdmitDate")}{e_row},"mmm-yyyy")',
             hint="TEXT with a date format code",
             explanation="In a date format code, mmm is the short month name and yyyy is the 4-digit year. The dash is "
                         "copied as-is. Without TEXT, a date joined to text shows its serial number: "
                         f"=\"Admitted \"&Encounters!{enc.col('AdmitDate')}{e_row} gives Admitted "
                         f"{int(data.excel_serial(t_enc['AdmitDate']))}, because a date is stored as a number."),
        Task(f"Build the ID-badge label for provider {t_prov['ProviderID']} (row {pr_row} of the Providers sheet) in the "
             f"form Dr. First Last, Credential. For example, provider {ex_prov['ProviderID']}'s label is "
             f"Dr. {ex_prov['FirstName']} {ex_prov['LastName']}, {ex_prov['Credential']}.",
             answer=badge,
             solution=f'="Dr. "&Providers!B{pr_row}&" "&Providers!C{pr_row}&", "&Providers!D{pr_row}',
             hint="Join the pieces with &. Spaces and punctuation go inside double quotes",
             explanation="The & operator joins text. Literal text, including every space, the period after Dr, and the "
                         "comma, goes inside double quotes. =CONCAT(\"Dr. \",Providers!B{0},\" \",Providers!C{0},\", \","
                         "Providers!D{0}) gives the same result.".format(pr_row)),
        Task(f"On the Patients sheet, fill the yellow HasDM column with TRUE when the patient's ChronicConditions list "
             f"includes DM (diabetes) and FALSE otherwise. Blank lists should give FALSE. Type your formula in "
             f"{pat.cell('HasDM', 0, sheet=False)}, then copy it down. The gray cell counts the TRUEs.",
             answer=dm_count, title="HasDM column (patients with diabetes)",
             solution=f'=ISNUMBER(SEARCH("DM",{Cc}{pf}))',
             summary=f'=IF(COUNTA({prng("HasDM")})=0,"",COUNTIF({prng("HasDM")},TRUE))',
             fill={"range": f"Patients!{pat.col('HasDM')}{pf}:{pat.col('HasDM')}{pl}",
                   "formula": f'=ISNUMBER(SEARCH("DM",{Cc}{pf}))'},
             live=f'=SUMPRODUCT(--ISNUMBER(SEARCH("DM",{prng("ChronicConditions")})))',
             hint="SEARCH returns a position or #VALUE!, and ISNUMBER turns that into TRUE or FALSE",
             explanation="SEARCH returns the position where DM starts, or #VALUE! when it isn't there (including in blank "
                         "cells). ISNUMBER converts any position to TRUE and the error to FALSE. SEARCH ignores case, so it "
                         "would also find dm. If a code could hide inside a longer code, search for \";DM;\" inside "
                         "\";\"&F2&\";\" instead."),
        Task(f"The patient portal username is the part of the email address before the @. Return the username for patient "
             f"{t_mail['PatientID']} (row {prow(t_mail['PatientID'])} of the Patients sheet).",
             answer=username, solution=f'=TEXTBEFORE({pcell("Email", t_mail["PatientID"])},"@")',
             hint="TEXTBEFORE (Microsoft 365 and Excel 2024), or LEFT + FIND",
             explanation="TEXTBEFORE returns everything before the first @, however long the username is. The classic "
                         "version, which works in every Excel, is "
                         f"=LEFT({pcell('Email', t_mail['PatientID'])},FIND(\"@\",{pcell('Email', t_mail['PatientID'])})-1): "
                         f"FIND returns the position of the @ ({len(username) + 1} here), and -1 stops LEFT just before it, so it "
                         f"keeps the {len(username)} characters of the username."),
        Task(f"How many chronic conditions does patient {t_cc['PatientID']} (row {prow(t_cc['PatientID'])} of the Patients "
             f"sheet) have? Split the semicolon-separated ChronicConditions list with TEXTSPLIT and count the pieces.",
             answer=cc_count, solution=f'=COUNTA(TEXTSPLIT({pcell("ChronicConditions", t_cc["PatientID"])},";"))',
             hint="COUNTA(TEXTSPLIT(…)) returns one number instead of a spill. Without TEXTSPLIT, count the semicolons "
                  "with LEN and SUBSTITUTE (guide section 6)",
             explanation="TEXTSPLIT spills one condition per cell to the right. Wrapping it in COUNTA counts the pieces and "
                         "returns a single number, so nothing spills into the Check column. Without TEXTSPLIT, count the "
                         "semicolons and add 1: =LEN({0})-LEN(SUBSTITUTE({0},\";\",\"\"))+1.".format(
                             pcell("ChronicConditions", t_cc["PatientID"]))),
        Task(f"On the Patients sheet, fill the yellow ContactLine column with the phone and email joined by \" | \" "
             f"(space, vertical bar, space), for example (555) 875-0698 | edward.abbott94@example.com. When there's no "
             f"email, show just the phone, with no dangling delimiter. Type your formula in "
             f"{pat.cell('ContactLine', 0, sheet=False)}, then copy it down. The gray cell counts lines that contain a |.",
             answer=with_email, title="ContactLine column (lines with both phone and email)",
             solution=f'=TEXTJOIN(" | ",TRUE,{pat.col("Phone")}{pf},{pat.col("Email")}{pf})',
             summary=f'=IF(COUNTA({prng("ContactLine")})=0,"",COUNTIF({prng("ContactLine")},"*|*"))',
             fill={"range": f"Patients!{pat.col('ContactLine')}{pf}:{pat.col('ContactLine')}{pl}",
                   "formula": f'=TEXTJOIN(" | ",TRUE,{pat.col("Phone")}{pf},{pat.col("Email")}{pf})'},
             live=f"=COUNTA({prng('Email')})",
             hint="TEXTJOIN(delimiter, ignore_empty, …)",
             explanation="TEXTJOIN puts the delimiter between items. With ignore_empty set to TRUE, it skips the blank email, "
                         "so those rows show only the phone. With FALSE, or with D2&\" | \"&E2, every row gets a delimiter "
                         "and the blank rows end in a dangling \" | \". Without TEXTJOIN (Excel 2016), add the delimiter only "
                         "when there is an email: =D2&IF(E2=\"\",\"\",\" | \"&E2)."),
        Task(f"On the Patients sheet, fill the yellow FirstName column with each patient's first name: everything after "
             f"the comma in PatientName, with no extra spaces. Watch out: some names have doubled, leading, or trailing "
             f"spaces, and {nbsp_rows} were pasted from the patient portal with non-breaking spaces (UNICHAR(160)). Case "
             f"doesn't matter here. Type your formula in {pat.cell('FirstName', 0, sheet=False)}, then copy it down. The "
             f"gray cell adds up the lengths of all your first names, so any leftover space changes the total.",
             answer=first_chars, title="FirstName column (checksum: total characters)",
             solution=f'=TRIM(MID(SUBSTITUTE({C}{pf},UNICHAR(160)," "),FIND(",",{C}{pf})+1,LEN({C}{pf})))',
             summary=f'=IF(COUNTA({prng("FirstName")})=0,"",SUMPRODUCT(LEN({prng("FirstName")})))',
             fill={"range": f"Patients!{pat.col('FirstName')}{pf}:{pat.col('FirstName')}{pl}",
                   "formula": f'=TRIM(MID(SUBSTITUTE({C}{pf},UNICHAR(160)," "),FIND(",",{C}{pf})+1,LEN({C}{pf})))'},
             live=f'=SUMPRODUCT(LEN(TRIM(MID(SUBSTITUTE({prng("PatientName")},UNICHAR(160)," "),'
                  f'FIND(",",{prng("PatientName")})+1,LEN({prng("PatientName")})))))',
             hint="TRIM can't remove a non-breaking space, so SUBSTITUTE it with a normal space first",
             explanation="Work from the inside out. SUBSTITUTE turns each non-breaking space into an ordinary space, "
                         "MID takes everything after the comma (LEN is simply a length that's long enough), and TRIM "
                         f"removes the leading, trailing, and doubled spaces. Without the SUBSTITUTE, the {nbsp_rows} portal rows "
                         f"keep an invisible character and the total comes out {no_nbsp_chars - first_chars} too high. "
                         "On Windows, CHAR(160) is the same character as UNICHAR(160). A version for Microsoft 365 and Excel 2024 is "
                         "=TRIM(SUBSTITUTE(TEXTAFTER(C2,\",\"),UNICHAR(160),\" \"))."),
        Task(f"Patient {t_disp['PatientID']}'s name was typed as \"{t_disp['PatientName']}\" (row {prow(t_disp['PatientID'])} "
             f"of the Patients sheet, with two trailing spaces). Return a clean display name in the form First Last, in "
             f"Proper Case. This check is case-sensitive.",
             answer=display, check="custom", custom_check="EXACT({cell},{key})", live=False,
             solution=f'=PROPER(TRIM(MID({pcell("PatientName", "PT10325")},FIND(",",{pcell("PatientName", "PT10325")})+1,'
                      f'LEN({pcell("PatientName", "PT10325")})))&" "&LEFT({pcell("PatientName", "PT10325")},'
                      f'FIND(",",{pcell("PatientName", "PT10325")})-1))',
             hint="Extract both parts, join them with \" \", then wrap everything in PROPER",
             explanation="MID + FIND takes the first name and TRIM drops the trailing spaces. LEFT + FIND takes the last name. "
                         "Join them with a space and wrap the whole thing in PROPER, which capitalizes the first letter of each "
                         "word and every letter that follows a non-letter. That's why the B after the apostrophe in O'Brien "
                         "comes out right. In Microsoft 365 and Excel 2024: "
                         "=PROPER(TRIM(TEXTAFTER({0},\",\"))&\" \"&TEXTBEFORE({0},\",\")).".format(
                             pcell("PatientName", "PT10325"))),
    ]

    # ------------------------------------------------------------------ bonus: legacy registration export + complexity count
    def city_of(s: str) -> str:
        """=PROPER(TRIM(SUBSTITUTE(TEXTBEFORE(s," ",-2),",","")))"""
        return xl_proper(xl_trim(s.rsplit(" ", 2)[0].replace(",", "")))

    def zip5_of(s: str) -> str:
        """=LEFT(TEXTAFTER(s," ",-1),5)"""
        return s.rsplit(" ", 1)[1][:5]

    def state_of(s: str) -> str:
        return s.rsplit(" ", 2)[1]

    lakeview = sum(1 for r in regs if city_of(r["CityStateZip"]) == "Lakeview Heights")
    lakeview_forms = Counter("UPPER" if r["CityStateZip"].isupper() else "no comma" if "," not in r["CityStateZip"]
                             else "ZIP+4" if "-" in r["CityStateZip"] else "plain"
                             for r in regs if city_of(r["CityStateZip"]) == "Lakeview Heights")
    assert set(lakeview_forms) == {"UPPER", "no comma", "ZIP+4", "plain"}, lakeview_forms
    z45501 = sum(1 for r in regs if zip5_of(r["CityStateZip"]) == "45501")
    z45501_plus4 = sum(1 for r in regs if zip5_of(r["CityStateZip"]) == "45501" and "-" in r["CityStateZip"])
    t_lab = next(r for r in regs if r["CityStateZip"] == "CEDAR RIDGE, OH 45720")
    lab_row = rf + regs.index(t_lab)
    R3 = f"Registrations!{reg.col('CityStateZip')}{lab_row}"
    label = f"{city_of(t_lab['CityStateZip'])}, {state_of(t_lab['CityStateZip'])} {zip5_of(t_lab['CityStateZip'])}"
    assert label == "Cedar Ridge, OH 45720"

    def ncond(s):
        return 0 if not s else len(s.split(";"))
    counts = [ncond(r["ChronicConditions"]) for r in patients]
    complex3 = sum(1 for c in counts if c >= 3)
    avg_cond = sum(counts) / len(counts)
    blanks = sum(1 for r in patients if not r["ChronicConditions"])
    avg_wrong = (sum(counts) + blanks) / len(counts)

    D = reg.col("CityStateZip")
    D2 = f"{D}{rf}"
    rng_city = rrng("City")
    rng_zip = rrng("ZIP5")
    rng_cc = prng("ChronicConditions")
    cc_formula = f'=IF({Cc}{pf}="",0,LEN({Cc}{pf})-LEN(SUBSTITUTE({Cc}{pf},";",""))+1)'
    L.bonus_title = "Bonus: Mailing list for the new diabetes education clinic"
    L.bonus_scenario = (
        "Bluestone is opening a diabetes education clinic in Cedar Ridge. Marketing wants a clean mailing list from the "
        "legacy registration export (the Registrations sheet), and care management wants to know how many current patients "
        "are 'complex' (three or more chronic conditions). The export crams city, state, and ZIP into one CityStateZip "
        "column in several styles: Cedar Ridge, OH 45720 · CEDAR RIDGE, OH 45720 · Cedar Ridge OH 45720 (no comma) · "
        "Cedar Ridge, OH 45720-1280 (ZIP+4). Some city names are two words. (The export also has duplicate records. "
        "Removing those is a Lesson 3.3 job, so leave them in.)")
    L.bonus = [
        Task(f"On the Registrations sheet, fill the yellow City column with just the city name in Proper Case, with no comma, "
             f"state, or ZIP (for example Lakeview Heights). Type your formula in {reg.cell('City', 0, sheet=False)}, then "
             f"copy it down. The gray cell counts rows whose City is exactly Lakeview Heights (case-sensitive). That town "
             f"appears in every messy style, so it's a good test.",
             answer=lakeview, title="City column (rows exactly 'Lakeview Heights')",
             solution=f'=PROPER(TRIM(SUBSTITUTE(TEXTBEFORE({D}{rf}," ",-2),",","")))',
             summary=f'=IF(COUNTA({rng_city})=0,"",SUMPRODUCT(--EXACT({rng_city},"Lakeview Heights")))',
             fill={"range": f"Registrations!{reg.col('City')}{rf}:{reg.col('City')}{rl}",
                   "formula": f'=PROPER(TRIM(SUBSTITUTE(TEXTBEFORE({D}{rf}," ",-2),",","")))'},
             live=f'=COUNTIF({rrng("CityStateZip")},"Lakeview Heights*")',
             hint="The city is everything before the second-to-last space. TEXTBEFORE accepts a negative instance_num.",
             explanation="The commas are unreliable, but the spaces are not: the last two spaces always separate the city, the "
                         f"state, and the ZIP. TEXTBEFORE({D2},\" \",-2) counts spaces from the end, so it returns everything before "
                         "the second-to-last space (\"Cedar Ridge,\" or \"CEDAR RIDGE\"). SUBSTITUTE removes a comma if there is one, "
                         "TRIM tidies up, and PROPER fixes the case. A FIND(\",\") approach fails with #VALUE! on the rows that "
                         "have no comma. Without Microsoft 365 or Excel 2024, the classic trick is "
                         f"=PROPER(SUBSTITUTE(LEFT({D2},LEN({D2})-LEN(TRIM(RIGHT(SUBSTITUTE({D2},\" \",REPT(\" \",100)),100)))-4),\",\",\"\")): "
                         "the TRIM(RIGHT(SUBSTITUTE(…))) part pulls out the last word (the ZIP), and LEFT keeps everything except "
                         "the ZIP and the 4 characters of \" OH \"."),
        Task(f"On the Registrations sheet, fill the yellow ZIP5 column with the 5-digit ZIP code as text. ZIP+4 codes like "
             f"45720-1280 must become 45720. Type your formula in {reg.cell('ZIP5', 0, sheet=False)}, then copy it down. "
             f"The gray cell counts rows in ZIP 45501 (downtown Bluestone).",
             answer=z45501, title="ZIP5 column (rows in ZIP 45501)",
             solution=f'=LEFT(TEXTAFTER({D}{rf}," ",-1),5)',
             summary=f'=IF(COUNTA({rng_zip})=0,"",COUNTIF({rng_zip},"45501"))',
             fill={"range": f"Registrations!{reg.col('ZIP5')}{rf}:{reg.col('ZIP5')}{rl}",
                   "formula": f'=LEFT(TEXTAFTER({D}{rf}," ",-1),5)'},
             live=f'=COUNTIF({rrng("CityStateZip")},"* 45501*")',
             hint="The ZIP is the last word. Keep only its first 5 characters",
             explanation=f"TEXTAFTER({D2},\" \",-1) returns everything after the last space: the whole ZIP or ZIP+4. LEFT(…,5) "
                         f"keeps the first five digits. RIGHT({D2},5) looks tempting but returns \"-1280\" for 45720-1280. "
                         f"{z45501_plus4} of the {z45501} rows in 45501 have a ZIP+4. Keep ZIPs as text: a ZIP like 02134 would "
                         f"lose its leading zero as a number. Classic version: "
                         f"=LEFT(TRIM(RIGHT(SUBSTITUTE({D2},\" \",REPT(\" \",100)),100)),5). SUBSTITUTE swaps every space for "
                         "100 spaces, RIGHT(…,100) then grabs a chunk that holds only the last word plus padding, and TRIM "
                         "strips the padding."),
        Task(f"Write one formula that turns record {t_lab['RecordID']}'s CityStateZip (row {lab_row} of the Registrations sheet, "
             f"\"{t_lab['CityStateZip']}\") into the mailing-label line City, ST 12345: city in Proper Case, a comma and a space, "
             f"the 2-letter state in capitals, a space, and the 5-digit ZIP. This check is case-sensitive.",
             answer=label, check="custom", custom_check="EXACT({cell},{key})", live=False,
             solution=f'=PROPER(TRIM(SUBSTITUTE(TEXTBEFORE({R3}," ",-2),",","")))&", "&'
                      f'UPPER(TEXTBEFORE(TEXTAFTER({R3}," ",-2)," "))&" "&LEFT(TEXTAFTER({R3}," ",-1),5)',
             hint="Build city, state, and ZIP separately, then join them with &. Which piece needs PROPER?",
             explanation="The state is the word between the last two spaces: TEXTAFTER(…,\" \",-2) gives \"OH 45720\", and "
                         "TEXTBEFORE(…,\" \") keeps \"OH\". Apply PROPER to the city only. Wrapping the whole line in PROPER is the "
                         "classic mistake: it turns OH into Oh. If you finished B1 and B2, "
                         f"=Registrations!{reg.col('City')}{lab_row}&\", OH \"&Registrations!{reg.col('ZIP5')}{lab_row} "
                         "also works, but extracting the "
                         "state keeps the formula correct for out-of-state patients."),
        Task(f"Back on the Patients sheet, fill the yellow ConditionCount column with the number of chronic conditions each "
             f"patient has: 0 when ChronicConditions is blank, 1 for HTN, 2 for HTN;DM, and so on. Type your formula in "
             f"{pat.cell('ConditionCount', 0, sheet=False)}, then copy it down. The gray cell counts 'complex' patients "
             f"with 3 or more conditions.",
             answer=complex3, title="ConditionCount column (patients with 3+ conditions)",
             solution=cc_formula,
             summary=f'=IF(COUNTA({prng("ConditionCount")})=0,"",COUNTIF({prng("ConditionCount")},">=3"))',
             fill={"range": f"Patients!{pat.col('ConditionCount')}{pf}:{pat.col('ConditionCount')}{pl}",
                   "formula": cc_formula},
             live=f'=SUMPRODUCT(({rng_cc}<>"")*((LEN({rng_cc})-LEN(SUBSTITUTE({rng_cc},";",""))+1)>=3))',
             hint="Items = semicolons + 1. How many semicolons? Compare LEN before and after removing them",
             explanation="LEN(F2)-LEN(SUBSTITUTE(F2,\";\",\"\")) is the number of semicolons, because removing them shortens "
                         "the text by exactly that many characters. A list always has one more item than delimiters, so add 1. "
                         "The IF handles blank lists: without it, a blank cell has 0 semicolons and would count as 1 condition. "
                         "COUNTA(TEXTSPLIT(F2,\";\")) has the same problem on blank cells, so it needs the same IF."),
        Task(f"What is the average number of chronic conditions per patient across all {N_PATIENTS} patients, including "
             f"those with none? Use your ConditionCount column. Enter it unrounded or rounded to 2 decimal places. The check "
             f"accepts either.",
             answer=avg_cond, fmt="0.00",
             solution=f"=AVERAGE({prng('ConditionCount')})",
             live=f'=SUMPRODUCT(({rng_cc}<>"")*(LEN({rng_cc})-LEN(SUBSTITUTE({rng_cc},";",""))+1))/ROWS({rng_cc})',
             hint="AVERAGE of the column you just filled",
             explanation=f"If your average comes out as {avg_wrong:.2f}, your column counts the {blanks} blank lists as 1 "
                         "condition each. That mistake doesn't change B4's count, because those rows stay below 3, but it "
                         "inflates the average. Testing a formula on the edge cases (blank, one item, many items) catches "
                         "errors like this."),
    ]

    modern = [str(i) for i, t in enumerate(L.tasks, 1)
              if any(f in t.solution for f in ("TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT"))]
    joiners = [str(i) for i, t in enumerate(L.tasks, 1) if "TEXTJOIN" in t.solution]
    assert len(joiners) == 1, joiners
    # Several tasks fill a yellow column on a data sheet, and a gray cell on Practice or Bonus reads it, so Start Here,
    # the Bonus sheet, and the README say so.
    cols = [str(i) for i, t in enumerate(L.tasks, 1) if t.summary]
    bcols = [f"B{i}" for i, t in enumerate(L.bonus, 1) if t.summary]
    assert bcols == ["B1", "B2", "B4"], bcols
    L.practice_how = ("Go to the 'Practice' sheet. Type a formula or value into each yellow cell. For tasks "
                      f"{', '.join(cols[:-1])}, and {cols[-1]}, fill the yellow column on the Encounters or Patients sheet "
                      "instead, and the task's gray cell on Practice reads it.")
    L.bonus_instructions = (
        f"Type a formula or value in each yellow cell. For {', '.join(bcols[:-1])}, and {bcols[-1]}, fill the yellow column on "
        "the Registrations or Patients sheet instead, and the gray cell here reads it. The Check column turns green when your "
        "answer matches. Stuck? Read the hint, then the lesson guide. Answers: right-click a sheet tab → Unhide… → "
        f"'{L.bonus_key_sheet}'.")
    L.bonus_where = ("Fill the yellow columns on the **Registrations** and **Patients** sheets, and type your answers in the "
                     "yellow cells on the **Bonus** sheet.")
    L.start_notes = [
        f"Tasks {' and '.join(modern)} and bonus parts B1–B3 use TEXTBEFORE, TEXTAFTER, or TEXTSPLIT, which need Microsoft 365, "
        "Excel for the web, or Excel 2024. In older versions those functions show #NAME?, so use the classic LEFT, MID, and "
        f"FIND methods from the guide instead. Task {joiners[0]} uses TEXTJOIN, which needs Excel 2019 or later. In Excel 2016, "
        "join the pieces with & and an IF instead.",
        f"Bonus parts {', '.join(bcols[:-1])}, and {bcols[-1]} also fill a yellow column, on the Registrations or Patients "
        "sheet. The data sheets are Excel Tables, so a formula typed in the first row usually fills the whole column "
        "automatically.",
    ]

    @L.customize
    def _start_here(wb, lesson, selftest):
        # Registrations comes after the Bonus tab because only the bonus uses it, so its Start Here line says so.
        ws = wb["Start Here"]
        row = next(r for r in range(1, ws.max_row + 1)
                   if ws.cell(row=r, column=2).value == "Registrations" and ws.cell(row=r, column=3).value)
        ws.cell(row=row, column=3).value += " (used only in the bonus)"
    return L
