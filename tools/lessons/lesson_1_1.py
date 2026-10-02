"""Lesson 1.1 · The Excel Interface & Navigation.

No formulas: learners navigate, select, read the status bar, use Find All, and unhide a column
and a sheet, then type what they find. Live formulas in the hidden key independently confirm
each Python-computed answer (they are never shown as the learner's solution).
"""
from __future__ import annotations

from collections import Counter

from xlcourse import Lesson, Task, data
from xlcourse.data import excel_serial

CODE = "1.1"

COLUMNS = ["PatientID", "MRN", "LastName", "FirstName", "Sex", "DOB", "City", "ZIP", "Phone", "Email",
           "PreferredLanguage", "PrimaryPayerID", "PCPProviderID", "HeightIn", "WeightLb", "ChronicConditions",
           "RegistrationDate"]
HIDDEN_COL = "Phone"


def build() -> Lesson:
    L = Lesson(
        code=CODE, module_dir="01-foundations", slug="01-excel-interface-navigation",
        title="The Excel Interface & Navigation", level="Beginner", minutes=30,
        objectives=[
            "Name the parts of the Excel window: ribbon, Quick Access Toolbar, Name Box, formula bar, grid, sheet tabs, status bar",
            "Understand workbooks, worksheets, cells, ranges, and cell addresses",
            "Move around large datasets fast with keyboard shortcuts and Go To",
            "Read quick statistics from the status bar",
            "Freeze panes, zoom, and hide/unhide rows, columns, and sheets",
        ],
        data_note="500 registered patients from the Bluestone Health System patient index (every 8th record), "
                  "plus a small payer list that starts out hidden.",
    )

    # ------------------------------------------------------------------ data
    # Every 8th patient (starting with the 8th) in PatientID order: deterministic, 500 rows.
    allp = sorted(data.load("patients"), key=lambda r: r["PatientID"])
    pts = allp[7::8]
    assert len(pts) == 500
    pat = L.add_table_sheet(
        "Patients", pts, table="tblPatients", columns=COLUMNS,
        widths={"Email": 32, "ChronicConditions": 20, "PreferredLanguage": 19, "PCPProviderID": 15,
                "PrimaryPayerID": 15, "RegistrationDate": 17},
        freeze=False,                                  # learners freeze panes themselves
        hidden_cols=[HIDDEN_COL],                      # task 9: unhide column I
    )
    payers = sorted(data.load("payers"), key=lambda r: r["PayerID"])
    pay = L.add_table_sheet("Payers", payers, table="tblPayers", columns=["PayerID", "PayerName", "PayerType"],
                            widths={"PayerName": 32, "PayerType": 22},
                            hidden=True)                # task 10: unhide the sheet (Start Here says it's hidden)
    payer_name = {r["PayerID"]: r["PayerName"] for r in payers}

    first, last = pat.first_row, pat.last_row          # 2, 501
    n = len(pts)
    col = pat.col                                      # header -> column letter
    last_col = col(COLUMNS[-1])                        # Q
    hidden_letter = col(HIDDEN_COL)                    # I

    def row_of(i: int) -> int:                         # 0-based patient index -> worksheet row
        return first + i

    def rng(h: str) -> str:
        return f"{col(h)}{first}:{col(h)}{last}"

    def find_all(term: str, whole: bool = False) -> int:
        """Cells on the Patients sheet that Excel's Find All would list (case-insensitive)."""
        t = term.lower()
        hits = 0
        for h in COLUMNS:                              # header row
            v = h.lower()
            hits += (v == t) if whole else (t in v)
        for r in pts:
            for h in COLUMNS:
                v = r[h]
                if v is None:
                    continue
                s = str(v).lower()
                hits += (s == t) if whole else (t in s)
        return hits

    # 1 · Name Box jump
    nb_row = 347
    nb_name = pts[nb_row - first]["LastName"]

    # 2 · bottom-right cell of the data (Ctrl + Down / Ctrl + Right / Ctrl + End)
    br_addr = f"{last_col}{last}"

    # 3 · Ctrl + Right in an empty row lands in the last worksheet column (16,384 = XFD)
    empty_row = 600
    assert empty_row > last + 50
    xfd_addr = f"XFD{empty_row}"

    # 4 · Ctrl + Down from the PCPProviderID header stops at the last filled cell before the first blank
    pcp = [r["PCPProviderID"] for r in pts]
    assert pcp[0] is not None and pcp[1] is not None
    first_blank_i = next(i for i, v in enumerate(pcp) if v is None)
    pcp_stop_row = row_of(first_blank_i - 1)
    no_pcp = sum(1 for v in pcp if v is None)

    # 5 · status bar Count of emails
    emails = sum(1 for r in pts if r["Email"])
    em = [bool(r["Email"]) for r in pts]
    if em[0] and em[1]:                                # Ctrl + Shift + Down from the first data cell (rule 1 or rule 2)
        em_stop_row = row_of(em.index(False) - 1)
    else:
        em_stop_row = row_of(next(i for i in range(1, n) if em[i]))
    assert em_stop_row - first + 1 < 10                # stops after only a few rows

    # 6 · status bar Average of weight
    weights = [r["WeightLb"] for r in pts]
    avg_w = sum(weights) / n

    # 7 · status bar Minimum of DOB
    oldest_dob = min(r["DOB"] for r in pts)

    # 8 · Find All: Millbrook (the term appears only in the City column)
    town = "Millbrook"
    town_count = sum(1 for r in pts if r["City"] == town)
    assert find_all(town) == town_count

    # 9 · hidden Phone column
    ph_i = 296
    ph_pt = pts[ph_i]
    ph_row = row_of(ph_i)
    phone = ph_pt["Phone"]
    digits = "".join(ch for ch in phone if ch.isdigit())
    assert find_all(ph_pt["PatientID"]) == 1

    # 10 · hidden Payers sheet: pick a commercial-payer patient deep in the list
    py_i = next(i for i in range(410, n) if pts[i]["PrimaryPayerID"] == "PY06")
    py_pt = pts[py_i]
    py_row = row_of(py_i)
    py_name = payer_name[py_pt["PrimaryPayerID"]]
    assert find_all(py_pt["PatientID"]) == 1

    # 11 · sheets in the workbook: Start Here + Practice + data sheets (Patients, Payers) + Bonus + Answer Key + Bonus Key
    total_sheets = 2 + len(L.data_sheets) + 3
    assert total_sheets == 7

    # 12 · Freeze Panes: row 1 + columns A:B frozen -> select C2
    freeze_cell = "C2"

    # Bonus
    htn_any = sum(1 for r in pts if r["ChronicConditions"] and "HTN" in r["ChronicConditions"].split(";"))
    assert find_all("HTN") == htn_any                  # no other cell contains the letters "htn"
    htn_only = sum(1 for r in pts if r["ChronicConditions"] == "HTN")
    assert find_all("HTN", whole=True) == htn_only
    max_w = max(weights)
    assert weights.count(max_w) == 1
    heavy = pts[weights.index(max_w)]
    assert find_all(f"{max_w}", whole=True) == 1 and find_all(f"{max_w}") == 1

    L.start_notes = [
        f"Two things are hidden on purpose so you can practice unhiding them: column {hidden_letter} on the Patients sheet "
        "and the whole Payers sheet.",
        "The Patients sheet opens with nothing frozen, so the headers scroll out of view. The guide shows how to freeze them.",
    ]
    L.practice_intro = (f"Every task uses the Patients sheet ({n} patients, rows {first}–{last}) unless it says otherwise. "
                        "No formulas needed: navigate, look, and type what you find. "
                        "Mac users: for Ctrl + arrow shortcuts such as Ctrl + ↓ and Ctrl + Shift + ↓, press ⌘ instead of Ctrl.")
    # The library's generic lines say "type a formula or value". This lesson has no formulas: every answer is read off the screen.
    L.practice_how = ("Go to the 'Practice' sheet. Each answer is something you read off the screen, so type the value itself "
                      "(a name, number, date, or address) into each yellow cell.")
    L.practice_instructions = (
        "Type the value you find (a name, number, date, or address) in each yellow cell. The Check column turns green when "
        "your answer matches. Stuck? Read the hint, then the lesson guide. Answers: right-click a sheet tab → Unhide… → "
        f"'{L.key_sheet}'.")
    L.bonus_instructions = (
        "Type the value you find in each yellow cell. The Check column turns green when your answer matches. "
        f"Stuck? Read the hint, then the lesson guide. Answers: right-click a sheet tab → Unhide… → '{L.bonus_key_sheet}'.")

    L.tasks = [
        Task(f"On the Patients sheet, click in the Name Box (the box at the left end of the formula bar), type C{nb_row}, "
             "and press Enter. What last name is in that cell?",
             answer=nb_name, title=f"Name Box jump to C{nb_row}",
             solution=f"1. Click the Name Box, type C{nb_row}, press Enter.\n2. Read the cell (or the formula bar).",
             live=f"=Patients!C{nb_row}",
             hint="The Name Box jumps to any address you type",
             explanation=f"The Name Box always shows the address of the active cell, and it works in reverse too: type an address, "
                         f"press Enter, and Excel takes you there. That's faster than scrolling {nb_row - first} rows. Column C holds "
                         f"LastName, so C{nb_row} is the last name of the patient in row {nb_row}."),
        Task("What is the address of the bottom-right cell of the patient data? From A1, press Ctrl + ↓ to find the last row, "
             "go back with Ctrl + ↑, then press Ctrl + → to find the last column. Type the address as column letter + row number "
             "(like B12).",
             answer=br_addr, accept=[f"${last_col}${last}"], title="Bottom-right cell of the data",
             solution=f"1. Click A1 and press Ctrl + ↓: the active cell becomes A{last}.\n"
                      f"2. Press Ctrl + ↑ to return to A1, then Ctrl + →: the active cell becomes {last_col}1.\n"
                      f"3. Combine them: {br_addr}. Press Ctrl + End (Mac: Control + End, or Control + Fn + →) to confirm. "
                      "It jumps straight there.",
             live='=ADDRESS(COUNTA(Patients!A:A),COUNTA(Patients!1:1),4)',
             hint="Column letter from Ctrl + →, row number from Ctrl + ↓. Ctrl + End confirms",
             explanation=f"Ctrl + arrow jumps to the edge of the block of filled cells. Column A and row 1 have no gaps, so the jumps land on "
                         f"the true last row ({last}: {n} patients plus the header row) and the true last column ({last_col}). "
                         f"Ctrl + End goes straight to the last used cell, {br_addr}. The hidden column {hidden_letter} doesn't change "
                         "the answer, because hiding a column doesn't remove it."),
        Task(f"Use the Name Box to go to cell A{empty_row}, which is in an empty row below the data, then press Ctrl + →. "
             "Excel races across the empty row to the very last column of the worksheet. "
             "What is the address of the cell you land on?",
             answer=xfd_addr, accept=[f"$XFD${empty_row}"], title="The last column of a worksheet",
             solution=f"1. Type A{empty_row} in the Name Box and press Enter.\n2. Press Ctrl + →. The Name Box shows {xfd_addr}.\n"
                      "3. Press Ctrl + ← to come back to column A.",
             live=f"=ADDRESS({empty_row},COLUMNS(Patients!1:1),4)",
             hint="A worksheet has 16,384 columns",
             explanation="When the row is empty, Ctrl + → has no data to stop at, so it goes to the edge of the sheet: column XFD, the "
                         "16,384th column. Ctrl + ↓ in an empty column goes to row 1,048,576. Every modern worksheet has exactly "
                         "1,048,576 rows × 16,384 columns."),
        Task("Click M1 (the PCPProviderID header) and press Ctrl + ↓ once. On which row number does Excel stop?",
             answer=pcp_stop_row, title="Ctrl + ↓ stops at a gap",
             solution=f"1. Click M1 (or type M1 in the Name Box).\n2. Press Ctrl + ↓. The Name Box shows M{pcp_stop_row}.",
             live=f"=MATCH(TRUE,ISBLANK(Patients!{rng('PCPProviderID')}),0)",
             hint="Ctrl + arrow stops at the edge of a block of filled cells",
             explanation=f"M{pcp_stop_row + 1} is empty because that patient has no primary care provider on file. Ctrl + ↓ stops at the "
                         f"last filled cell before the gap, M{pcp_stop_row}, even though the data continues to row {last}. "
                         "Never assume Ctrl + ↓ found the bottom of a column with blanks. Check the row number, or press Ctrl + ↓ again "
                         "to keep jumping."),
        Task(f"Select the Email data cells {rng('Email')}: type {rng('Email')} in the Name Box and press Enter. "
             "How many patients have an email address on file? Read Count on the status bar.",
             answer=emails, title="Status bar Count (patients with an email)",
             solution=f"1. Type {rng('Email')} in the Name Box and press Enter.\n2. Read Count on the status bar.",
             live=f"=COUNTA(Patients!{rng('Email')})",
             hint="Count = cells that aren't empty",
             explanation=f"The status bar's Count counts every non-empty cell, text included, so it counts the email addresses and skips "
                         f"the blanks. Typing the range in the Name Box selects exactly {n} cells. Ctrl + Shift + ↓ from "
                         f"{col('Email')}{first} would be a trap, because the Email column has gaps: it would select only "
                         f"{col('Email')}{first}:{col('Email')}{em_stop_row}. "
                         "In a Table you can also click any Email cell and press Ctrl + Space (Mac: Control + Space) to select "
                         "just that column's data."),
        Task(f"Select the WeightLb values {rng('WeightLb')} with Go To: press F5 or Ctrl + G (Mac: Control + G), "
             f"type {rng('WeightLb')}, and press Enter. What is the average weight in pounds? Round to 1 decimal place.",
             answer=round(avg_w, 1), fmt="0.0", tol=0.051, title="Status bar Average (weight, lb)",
             solution=f"1. Press F5 (or Ctrl + G), type {rng('WeightLb')} in Reference, press Enter.\n"
                      "2. Read Average on the status bar and round it to 1 decimal place.",
             live=f"=ROUND(AVERAGE(Patients!{rng('WeightLb')}),1)",
             hint="Average is on the status bar by default",
             explanation=f"Go To selects any range you type, however large. The status bar's Average is "
                         f"{avg_w:.4f}, which rounds to {avg_w:.1f}. The average is pulled down by children in the list: the lightest "
                         f"patient weighs {min(weights)} lb, which is a baby and not a typo. Always glance at Minimum and Maximum "
                         "before trusting an average."),
        Task("Turn on Minimum in the status bar: right-click the status bar and tick Minimum. Then click cell F2 (the first "
             "DOB) and press Ctrl + Shift + ↓ to select every date of birth. What is the date of birth of the oldest patient? "
             "Type it as month/day/year, the way the status bar shows it. (If your computer uses day/month dates, type the "
             "month as a word instead, such as 15 Mar 1950.)",
             answer=oldest_dob, fmt="mm/dd/yyyy", title="Status bar Minimum (oldest patient's DOB)",
             solution="1. Right-click the status bar and tick Minimum (and Maximum while you're there).\n"
                      f"2. Click cell F2 and press Ctrl + Shift + ↓ to select {rng('DOB')}.\n3. Read Minimum on the status bar.",
             live=f"=MIN(Patients!{rng('DOB')})",
             hint="The oldest patient has the earliest date, so look at Minimum",
             explanation="Excel stores dates as numbers that grow by 1 each day, so the earliest date is the smallest number and "
                         "Minimum finds it. The status bar shows it as a date because the cells are formatted as dates. (If you "
                         f"see a plain number such as {int(excel_serial(oldest_dob))} instead, that's the date's serial number, and the check accepts it.) "
                         "Ctrl + Shift + ↓ is safe here because the DOB column has no blanks."),
        Task(f"How many patients live in {town}? On the Patients sheet, click a single cell such as A1 so Find searches the "
             f"whole sheet. Then press Ctrl + F (Mac: Control + F), type {town}, click Find All, and read the count at the bottom of "
             "the dialog.",
             answer=town_count, title=f"Find All: patients in {town}",
             solution=f"1. Click A1 (one cell, so Find searches the whole sheet).\n2. Press Ctrl + F, type {town}, and click Find All.\n"
                      "3. Read '… cell(s) found' at the bottom of the dialog.",
             live=f'=COUNTIF(Patients!A1:{last_col}{last},"*{town}*")',
             hint="Find All shows '… cell(s) found'",
             explanation=f"Find All lists every matching cell on the sheet and counts them. That's a count of patients here only because "
                         f"the word {town} appears in the City column and nowhere else. If a search term could also appear inside "
                         "other text, such as an email address, turn on Match entire cell contents under Options. Clicking a single "
                         "cell first matters too. When several cells are selected, Find searches only inside the selection, so with "
                         "the DOB column still selected from the previous task, Find All would report nothing."),
        Task(f"Column {hidden_letter} is hidden (the column letters jump from {col('ZIP')} to {col('Email')}). Unhide it, then "
             f"click cell A1 and use Ctrl + F (Mac: Control + F) to find patient {ph_pt['PatientID']}. "
             "What is that patient's phone number?",
             answer=phone, accept=[digits, f"{digits[:3]}-{digits[3:6]}-{digits[6:]}", f"({digits[:3]}){digits[3:6]}-{digits[6:]}",
                                   f"{digits[:3]} {digits[3:6]} {digits[6:]}"],
             title=f"Unhide column {hidden_letter} (phone number)",
             solution=f"1. Click the {col('ZIP')} column header, Shift + click the {col('Email')} header, then right-click → **Unhide**.\n"
                      f"2. Click A1, press Ctrl + F, type {ph_pt['PatientID']}, click Find Next, then close the dialog.\n"
                      f"3. Read column {hidden_letter} on row {ph_row}.",
             live=f"=Patients!{hidden_letter}{ph_row}",
             hint="Select the columns on both sides of the gap, right-click → **Unhide**",
             explanation=f"You can't click a hidden column, so you select the columns on both sides of it ({col('ZIP')} and "
                         f"{col('Email')}) and choose Unhide. Click A1 before you search, because while those three columns are "
                         f"selected, Find looks only inside them. The patient is on row {ph_row}. Hiding never deletes data: the "
                         "phone numbers were there all along, so a hidden column is not a safe place for confidential data."),
        Task(f"Patient {py_pt['PatientID']} has a PrimaryPayerID in column {col('PrimaryPayerID')}. The payer names are on the "
             "Payers sheet, which is hidden. Unhide it (right-click any sheet tab → **Unhide…**) and type the name of this "
             "patient's payer.",
             answer=py_name, title="Unhide the Payers sheet",
             solution=f"1. On the Patients sheet, click A1, press Ctrl + F, and find {py_pt['PatientID']}: row {py_row}, payer ID "
                      f"{py_pt['PrimaryPayerID']}.\n"
                      "2. Right-click any sheet tab → **Unhide…**, pick Payers, and click OK.\n"
                      f"3. On the Payers sheet, {py_pt['PrimaryPayerID']} is {py_name}.",
             live=f"=INDEX(Payers!{pay.rng('PayerName', sheet=False)},MATCH(Patients!{col('PrimaryPayerID')}{py_row},"
                  f"Payers!{pay.rng('PayerID', sheet=False)},0))",
             hint="Right-click a sheet tab → **Unhide…**",
             explanation="Hidden sheets don't show a tab, so the only clue is the Unhide… command becoming available. This is the "
                         "same move you'll use to open the Answer Key in every lesson. Small lookup lists like this one are often "
                         "hidden to keep a workbook tidy, and in Lesson 2.6 you'll learn to pull names from them automatically "
                         "with a lookup formula."),
        Task("How many worksheets does this workbook contain in total, hidden ones included? Count the tabs you can see, then "
             "right-click a tab → **Unhide…** to see what's still hidden. Look, but don't unhide the answer keys yet!",
             answer=total_sheets, title="Worksheets in this workbook",
             solution="1. Count the visible tabs: Start Here, Practice, Patients, Payers (now unhidden), Bonus.\n"
                      "2. Right-click a tab → **Unhide…**: the list shows Answer Key and Bonus Key. Click Cancel.\n"
                      f"3. 5 + 2 = {total_sheets}.",
             live="=SHEETS()",
             hint="Visible tabs + the names listed in the Unhide dialog",
             explanation=f"A workbook is the file, and each worksheet is one tab inside it. This file has {total_sheets} worksheets, and "
                         "2 of them are still hidden: the Answer Key and the Bonus Key. The total is the same whether or not you "
                         "unhid Payers first, because hiding a sheet doesn't remove it."),
        Task("You want row 1 (the headers) and columns A:B (PatientID and MRN) to stay on screen while you scroll. Which cell "
             "must you select before choosing **View → Freeze Panes → Freeze Panes**? Type its address.",
             answer=freeze_cell, accept=["$C$2"], title="Freeze Panes: which cell to select",
             solution="1. Click C2.\n2. Choose **View → Freeze Panes → Freeze Panes**.\n"
                      "3. Scroll down and right: row 1 and columns A:B stay put. (**View → Freeze Panes → Unfreeze Panes** undoes it.)",
             live=False,
             hint="Excel freezes everything above and to the left of the selected cell",
             explanation="Freeze Panes freezes the rows above the active cell and the columns to its left. To keep 1 row and 2 "
                         "columns, select the cell just below row 1 and just right of column B: C2. Selecting B2 would keep row 1 "
                         "but only column A, and A2 would freeze only row 1, which is the same as Freeze Top Row."),
    ]

    # ------------------------------------------------------------------ bonus
    L.bonus_title = "Bonus: Data-quality scavenger hunt"
    L.bonus_scenario = ("The population-health team is starting a blood-pressure outreach program and will build its mailing list "
                        "from this patient index. Before the letters go out, the data steward asks you to check four facts. Set up "
                        "the Patients sheet first: if you froze panes earlier, choose **View → Freeze Panes → Unfreeze Panes**. Then "
                        "press Ctrl + Home (Mac: Control + Home), click B2, and choose **View → Freeze Panes → Freeze Panes** so the headers "
                        "and the PatientID column stay in view. Also turn on Minimum and Maximum in the status bar. You'll look "
                        "everything up on the Patients sheet and type your answers on the Bonus sheet, without a single formula.")
    L.bonus = [
        Task(f"Each outreach letter is signed by the patient's primary care provider (PCP), and a blank PCPProviderID means no "
             f"PCP is on file. Use Go To (F5 or Ctrl + G, Mac: Control + G) to select {rng('PCPProviderID')}. How many of the {n} "
             "patients have no PCP?",
             answer=no_pcp, title="Patients with no PCP on file",
             solution=f"1. Press F5, type {rng('PCPProviderID')}, press Enter.\n"
                      f"2. The status bar shows Count: {n - no_pcp}.\n3. {n} − {n - no_pcp} = {no_pcp}.",
             live=f"=COUNTBLANK(Patients!{rng('PCPProviderID')})",
             hint="The status bar counts filled cells, but you want the empty ones",
             explanation=f"Count only counts non-empty cells, so it tells you how many patients do have a PCP ({n - no_pcp}). "
                         f"The range {rng('PCPProviderID')} covers rows {first} to {last}, which is {n} cells, so {n} − {n - no_pcp} "
                         f"leaves {no_pcp} blanks. Any of those patients who have hypertension need a PCP assigned before a letter can go out. "
                         f"(If you select by dragging instead, the Name Box shows the size of the selection while you drag, such "
                         f"as {n}R x 1C.)"),
        Task("The program targets every patient with hypertension, coded HTN. Click a single cell such as A1, so Find searches "
             "the whole sheet rather than the column you just selected. Press Ctrl + F (Mac: Control + F), click Options >> and make "
             "sure Match entire cell contents is OFF, then Find All for HTN. How many patients have HTN anywhere in their "
             "ChronicConditions list?",
             answer=htn_any, title="Find All: HTN anywhere in the list",
             solution="1. Click A1, press Ctrl + F, click **Options >>**, and untick Match entire cell contents.\n2. Type HTN and click Find All.\n"
                      "3. Read the count at the bottom of the dialog.",
             live=f'=COUNTIF(Patients!A1:{last_col}{last},"*HTN*")',
             hint="Partial matches count: 'HTN;DM' contains HTN",
             explanation="By default Find matches text anywhere inside a cell, so it finds HTN on its own and also inside lists like "
                         "HTN;DM and HTN;HF. No other column contains the letters HTN, so every hit is a patient with hypertension."),
        Task("Patients whose only condition is hypertension will get a shorter letter. Turn ON Match entire cell contents and "
             "Find All for HTN again. How many patients have hypertension as their ONLY recorded chronic condition?",
             answer=htn_only, title="Find All: HTN as the only condition",
             solution="1. In the Find dialog, tick Match entire cell contents.\n2. Find All for HTN.\n3. Read the count.",
             live=f'=COUNTIF(Patients!A1:{last_col}{last},"HTN")',
             hint="Match entire cell contents finds cells that are exactly HTN",
             explanation=f"With Match entire cell contents on, a cell must equal HTN exactly, so HTN;DM no longer counts. The gap "
                         f"between the two answers ({htn_any} − {htn_only} = {htn_any - htn_only}) is the number of hypertensive "
                         "patients with at least one other condition. When you finish this challenge, turn the option off again, "
                         "because Find keeps your settings until you close Excel."),
        Task(f"Finally, the steward wants to confirm that the largest weight on file is real and not a typo. Which patient is "
             f"the heaviest? Select {rng('WeightLb')}, read Maximum on the status bar, then use Find to locate that weight and "
             "read the PatientID in the frozen column A. Type the PatientID.",
             answer=heavy["PatientID"], title="The heaviest patient",
             solution=f"1. Select {rng('WeightLb')} (Name Box or Go To) and read Maximum: {max_w}.\n"
                      f"2. With the weights still selected, press Ctrl + F, type {max_w} in Find what, and click Find Next. (Match entire cell "
                      f"contents can be on or off here, because no other cell on the sheet contains {max_w}.)\n"
                      f"3. Excel selects {col('WeightLb')}{weights.index(max_w) + first}. The frozen column A shows {heavy['PatientID']}.",
             live=f"=INDEX(Patients!{rng('PatientID')},MATCH(MAX(Patients!{rng('WeightLb')}),Patients!{rng('WeightLb')},0))",
             hint="Status bar Maximum, then Ctrl + F for that number",
             explanation="The status bar tells you what the largest value is but not where it is. Find tells you where, and "
                         "leaving the weights selected helps, because Find then searches only inside the selection. Because "
                         "panes are frozen at B2, column A stays on screen when Find scrolls over to the WeightLb column, so you "
                         "can read the ID without losing your place. Teams use this check to plan bariatric beds and lift "
                         f"equipment, and it's also how you'd spot a typo such as {max_w * 10:.0f} lb."),
    ]

    # sanity: the town and term choices above are the only places these strings occur
    assert Counter(r["City"] for r in pts)[town] == town_count
    return L
