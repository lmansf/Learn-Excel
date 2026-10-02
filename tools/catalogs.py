"""Reference catalogs for the synthetic Bluestone Health System.

Everything here is fictional except the ICD-10-CM diagnosis codes, which are
real public codes used only to make the data feel authentic. Nothing in this
course is clinical guidance.
"""

FACILITIES = [
    # FacilityID, FacilityName, City, State, FacilityType, LicensedBeds, OpenedYear
    ("F01", "Bluestone Memorial Hospital", "Bluestone", "OH", "Tertiary Acute Care", 300, 1952),
    ("F02", "Ashby Falls Community Hospital", "Ashby Falls", "OH", "Community Hospital", 60, 1978),
    ("F03", "Cedar Ridge Medical Center", "Cedar Ridge", "OH", "Acute Care", 90, 1996),
    ("F04", "Bluestone Outpatient Pavilion", "Bluestone", "OH", "Ambulatory Care Center", 0, 2015),
]

# DeptID, DeptName, FacilityID, ServiceLine, UnitType, StaffedBeds, CostCenter
DEPARTMENTS = [
    ("D100", "Emergency Department", "F01", "Emergency", "Emergency", None, 6100),
    ("D110", "Medical-Surgical 4 West", "F01", "Medicine", "Inpatient", 36, 6110),
    ("D111", "Medical-Surgical 5 East", "F01", "Medicine", "Inpatient", 32, 6111),
    ("D120", "Cardiac Step-Down", "F01", "Cardiovascular", "Inpatient", 24, 6120),
    ("D130", "Intensive Care Unit", "F01", "Critical Care", "Critical Care", 20, 6130),
    ("D140", "Orthopedics & Spine", "F01", "Orthopedics", "Inpatient", 22, 6140),
    ("D150", "Oncology", "F01", "Oncology", "Inpatient", 18, 6150),
    ("D160", "Neuroscience & Stroke", "F01", "Neuroscience", "Inpatient", 16, 6160),
    ("D170", "Labor & Delivery", "F01", "Women & Children", "Inpatient", 20, 6170),
    ("D180", "Pediatrics", "F01", "Women & Children", "Inpatient", 14, 6180),
    ("D190", "Behavioral Health", "F01", "Behavioral Health", "Inpatient", 18, 6190),
    ("D195", "Observation Unit", "F01", "Medicine", "Observation", 12, 6195),
    ("D196", "Perioperative Services", "F01", "Surgery", "Ancillary", None, 6196),
    ("D200", "Emergency Department", "F02", "Emergency", "Emergency", None, 6200),
    ("D210", "Medical-Surgical", "F02", "Medicine", "Inpatient", 28, 6210),
    ("D230", "Intensive Care Unit", "F02", "Critical Care", "Critical Care", 8, 6230),
    ("D270", "Labor & Delivery", "F02", "Women & Children", "Inpatient", 10, 6270),
    ("D300", "Emergency Department", "F03", "Emergency", "Emergency", None, 6300),
    ("D310", "Medical-Surgical", "F03", "Medicine", "Inpatient", 30, 6310),
    ("D320", "Cardiac Step-Down", "F03", "Cardiovascular", "Inpatient", 16, 6320),
    ("D330", "Intensive Care Unit", "F03", "Critical Care", "Critical Care", 12, 6330),
    ("D340", "Orthopedics", "F03", "Orthopedics", "Inpatient", 14, 6340),
    ("D400", "Primary Care Clinic", "F04", "Ambulatory", "Outpatient", None, 6400),
    ("D410", "Cardiology Clinic", "F04", "Cardiovascular", "Outpatient", None, 6410),
    ("D420", "Orthopedic Clinic", "F04", "Orthopedics", "Outpatient", None, 6420),
    ("D430", "Infusion & Oncology Clinic", "F04", "Oncology", "Outpatient", None, 6430),
    ("D440", "Endocrinology & Diabetes Clinic", "F04", "Ambulatory", "Outpatient", None, 6440),
    ("D450", "Pediatric Clinic", "F04", "Women & Children", "Outpatient", None, 6450),
    ("D500", "Laboratory", "F01", "Ancillary", "Ancillary", None, 6500),
    ("D510", "Radiology", "F01", "Ancillary", "Ancillary", None, 6510),
    ("D520", "Pharmacy", "F01", "Ancillary", "Ancillary", None, 6520),
]

ED_DEPT = {"F01": "D100", "F02": "D200", "F03": "D300"}
MEDSURG_DEPT = {"F01": ["D110", "D111"], "F02": ["D210"], "F03": ["D310"]}
ICU_DEPT = {"F01": "D130", "F02": "D230", "F03": "D330"}
CARDIAC_DEPT = {"F01": "D120", "F03": "D320"}
ORTHO_DEPT = {"F01": "D140", "F03": "D340"}
OB_DEPT = {"F01": "D170", "F02": "D270"}
OBS_DEPT = {"F01": "D195", "F02": "D210", "F03": "D310"}

PAYERS = [
    # PayerID, PayerName, PayerType, AvgAllowedPctOfCharges, AvgDaysToPay, TimelyFilingDays, DenialRate
    ("PY01", "Medicare", "Government", 0.31, 14, 365, 0.06),
    ("PY02", "Silverline Medicare Advantage", "Medicare Advantage", 0.33, 28, 180, 0.12),
    ("PY03", "State Medicaid", "Government", 0.24, 35, 365, 0.12),
    ("PY04", "Keystone Health Partners", "Commercial", 0.52, 24, 90, 0.08),
    ("PY05", "Evergreen Mutual Insurance", "Commercial", 0.48, 30, 120, 0.10),
    ("PY06", "Summit Choice PPO", "Commercial", 0.55, 21, 90, 0.07),
    ("PY07", "Self-Pay", "Self-Pay", 0.40, 60, 0, 0.0),
    ("PY08", "Workers' Compensation", "Workers' Comp", 0.60, 45, 180, 0.10),
]

DENIAL_REASONS = [
    ("Authorization Required", 0.27),
    ("Medical Necessity", 0.21),
    ("Coding Error", 0.18),
    ("Eligibility / Coverage", 0.14),
    ("Timely Filing", 0.08),
    ("Missing Documentation", 0.08),
    ("Duplicate Claim", 0.04),
]

# Towns served by the system: (City, ZIP, nearest facility, population weight)
TOWNS = [
    ("Bluestone", "45501", "F01", 18),
    ("Bluestone", "45502", "F01", 14),
    ("Bluestone", "45503", "F01", 10),
    ("Lakeview Heights", "45510", "F01", 7),
    ("Millbrook", "45530", "F01", 6),
    ("Riverbend", "45545", "F01", 5),
    ("Ashby Falls", "45610", "F02", 9),
    ("Fairhaven", "45620", "F02", 4),
    ("Pine Hollow", "45640", "F02", 4),
    ("Cedar Ridge", "45720", "F03", 10),
    ("Oak Glen", "45730", "F03", 5),
    ("Granite Springs", "45750", "F03", 4),
]

LANGUAGES = [("English", 88), ("Spanish", 7.5), ("Arabic", 1.2), ("Vietnamese", 1.1), ("Somali", 1.0), ("Mandarin", 0.8), ("Russian", 0.4)]

FEMALE_FIRST = """Mary Patricia Jennifer Linda Elizabeth Barbara Susan Jessica Sarah Karen Lisa Nancy Betty Sandra Margaret
Ashley Kimberly Emily Donna Michelle Carol Amanda Melissa Deborah Stephanie Dorothy Rebecca Sharon Laura Cynthia Amy
Kathleen Angela Shirley Brenda Emma Anna Pamela Nicole Samantha Katherine Christine Helen Debra Rachel Carolyn Janet
Maria Catherine Heather Diane Olivia Julie Joyce Victoria Ruth Virginia Lauren Kelly Christina Joan Evelyn Judith
Andrea Hannah Megan Cheryl Jacqueline Martha Madison Teresa Gloria Sara Janice Ann Kathryn Abigail Sophia Frances
Jean Alice Judy Isabella Julia Grace Amber Denise Danielle Marilyn Beverly Charlotte Natalie Theresa Diana Brittany
Doris Kayla Alexis Lori Marie Ava Mia Harper Aaliyah Camila Lucia Rosa Ana Mei Linh Fatima Amina Leila Priya""".split()

MALE_FIRST = """James Robert John Michael David William Richard Joseph Thomas Christopher Charles Daniel Matthew Anthony
Mark Donald Steven Andrew Paul Joshua Kenneth Kevin Brian George Timothy Ronald Jason Edward Jeffrey Ryan Jacob Gary
Nicholas Eric Jonathan Stephen Larry Justin Scott Brandon Benjamin Samuel Gregory Alexander Patrick Frank Raymond
Jack Dennis Jerry Tyler Aaron Jose Adam Nathan Henry Zachary Douglas Peter Kyle Noah Ethan Jeremy Walter Christian
Keith Roger Terry Austin Sean Gerald Carl Harold Dylan Arthur Lawrence Jordan Jesse Bryan Billy Bruce Gabriel Joe
Logan Alan Juan Albert Willie Elijah Wayne Randy Vincent Mason Roy Ralph Bobby Russell Bradley Philip Eugene Liam
Lucas Mateo Diego Carlos Luis Omar Ahmed Minh Wei Hassan Ivan""".split()

LAST_NAMES = """Smith Johnson Williams Brown Jones Garcia Miller Davis Rodriguez Martinez Hernandez Lopez Gonzalez Wilson
Anderson Thomas Taylor Moore Jackson Martin Lee Perez Thompson White Harris Sanchez Clark Ramirez Lewis Robinson Walker
Young Allen King Wright Scott Torres Nguyen Hill Flores Green Adams Nelson Baker Hall Rivera Campbell Mitchell Carter
Roberts Gomez Phillips Evans Turner Diaz Parker Cruz Edwards Collins Reyes Stewart Morris Morales Murphy Cook Rogers
Gutierrez Ortiz Morgan Cooper Peterson Bailey Reed Kelly Howard Ramos Kim Cox Ward Richardson Watson Brooks Chavez Wood
James Bennett Gray Mendoza Ruiz Hughes Price Alvarez Castillo Sanders Patel Myers Long Ross Foster Jimenez Powell
Jenkins Perry Russell Sullivan Bell Coleman Butler Henderson Barnes Gonzales Fisher Vasquez Simmons Romero Jordan
Patterson Alexander Hamilton Graham Reynolds Griffin Wallace Moreno West Cole Hayes Bryant Herrera Gibson Ellis Tran
Medina Aguilar Stevens Murray Ford Castro Marshall Owens Harrison Fernandez McDonald Woods Washington Kennedy Wells
Vargas Henry Chen Freeman Webb Tucker Guzman Burns Crawford Olson Simpson Porter Hunter Gordon Mendez Silva Shaw Snyder
Mason Dixon Munoz Hunt Hicks Holmes Palmer Wagner Black Robertson Boyd Rose Stone Salazar Fox Warren Mills Meyer Rice
Schmidt Garza Daniels Ferguson Nichols Stephens Soto Weaver Ryan Gardner Payne Grant Dunn Kelley Spencer Hawkins Arnold
Pierce Vazquez Hansen Peters Santos Hart Bradley Knight Elliott Cunningham Duncan Armstrong Hudson Carroll Lane Riley
Andrews Alvarado Ray Delgado Berry Perkins Hoffman Johnston Matthews Pena Richards Contreras Willis Carpenter Lawrence
Sandoval O'Brien Abbott Novak Okafor Haddad Kowalski Lindqvist Nakamura Osei Petrov Yilmaz Abdi""".split()

# ---------------------------------------------------------------------------
# Diagnosis catalog. For each code:
#   desc, category, chronic, weights by encounter type (IP, OBS, ED, OP),
#   age range, sex ("F", "M" or None), winter multiplier, summer multiplier,
#   inpatient route key, mean IP LOS (days), expected LOS (benchmark),
#   base procedure charge, readmission risk, inpatient mortality,
#   chief complaint, ED acuity bias (lower = sicker), outpatient clinic key
# ---------------------------------------------------------------------------
DX = {
    # code: (desc, category, chronic, (ip, obs, ed, op), (amin, amax), sex, winter, summer, route, los, exp_los, proc_charge, readmit, mortality, complaint, acuity, clinic)
    "I21.4": ("Non-ST elevation (NSTEMI) myocardial infarction", "Circulatory", "N", (55, 6, 0, 0), (40, 95), None, 1.15, 0.95, "cardiac", 3.6, 3.4, 26000, 0.15, 0.035, "Chest Pain", 2, None),
    "I50.9": ("Heart failure, unspecified", "Circulatory", "Y", (95, 10, 8, 30), (50, 99), None, 1.25, 0.9, "cardiac", 4.7, 4.1, 0, 0.23, 0.03, "Shortness of Breath", 2.5, "cardio"),
    "I48.91": ("Unspecified atrial fibrillation", "Circulatory", "Y", (35, 12, 18, 40), (45, 99), None, 1.05, 1.0, "cardiac", 3.1, 2.9, 0, 0.13, 0.01, "Palpitations", 2.8, "cardio"),
    "I10": ("Essential (primary) hypertension", "Circulatory", "Y", (0, 3, 14, 140), (25, 99), None, 1.0, 1.0, "medsurg", 2.0, 2.0, 0, 0.05, 0.0, "Headache", 3.6, "primary"),
    "R07.9": ("Chest pain, unspecified", "Symptoms & Signs", "N", (0, 55, 95, 0), (18, 95), None, 1.0, 1.0, "obs", 1.0, 1.0, 0, 0.05, 0.0, "Chest Pain", 2.9, None),
    "I63.9": ("Cerebral infarction, unspecified", "Circulatory", "N", (38, 3, 0, 0), (45, 99), None, 1.05, 1.0, "neuro", 4.6, 4.2, 0, 0.12, 0.06, "Weakness / Numbness", 1.9, None),
    "J18.9": ("Pneumonia, unspecified organism", "Respiratory", "N", (95, 10, 30, 0), (1, 99), None, 2.2, 0.55, "medsurg", 4.3, 3.9, 0, 0.16, 0.03, "Cough / Fever", 2.7, None),
    "J44.1": ("Chronic obstructive pulmonary disease with (acute) exacerbation", "Respiratory", "Y", (75, 12, 25, 18), (45, 99), None, 1.9, 0.6, "medsurg", 3.9, 3.4, 0, 0.20, 0.02, "Shortness of Breath", 2.7, "primary"),
    "J45.909": ("Unspecified asthma, uncomplicated", "Respiratory", "Y", (6, 12, 45, 30), (3, 70), None, 1.4, 0.9, "medsurg", 2.2, 2.1, 0, 0.06, 0.0, "Shortness of Breath", 3.2, "primary"),
    "J06.9": ("Acute upper respiratory infection, unspecified", "Respiratory", "N", (0, 0, 120, 90), (0, 90), None, 2.4, 0.5, "medsurg", 1.0, 1.0, 0, 0.0, 0.0, "Cough / Fever", 4.4, "primary"),
    "U07.1": ("COVID-19", "Respiratory", "N", (22, 6, 40, 0), (1, 99), None, 2.3, 0.8, "medsurg", 4.9, 4.4, 0, 0.14, 0.04, "Cough / Fever", 3.0, None),
    "J96.01": ("Acute respiratory failure with hypoxia", "Respiratory", "N", (30, 0, 0, 0), (30, 99), None, 1.7, 0.8, "icu", 5.6, 4.8, 0, 0.19, 0.09, "Shortness of Breath", 1.6, None),
    "A41.9": ("Sepsis, unspecified organism", "Infectious", "N", (105, 0, 0, 0), (1, 99), None, 1.3, 0.95, "icu", 6.1, 5.0, 0, 0.18, 0.08, "Fever / Altered Mental Status", 1.9, None),
    "N39.0": ("Urinary tract infection, site not specified", "Genitourinary", "N", (55, 12, 85, 20), (2, 99), None, 1.0, 1.05, "medsurg", 3.4, 3.1, 0, 0.12, 0.01, "Painful Urination / Fever", 3.4, "primary"),
    "L03.115": ("Cellulitis of right lower limb", "Skin", "N", (35, 8, 40, 10), (18, 99), None, 0.85, 1.3, "medsurg", 3.8, 3.4, 0, 0.10, 0.0, "Skin Redness / Swelling", 3.5, "primary"),
    "E11.9": ("Type 2 diabetes mellitus without complications", "Endocrine", "Y", (0, 0, 6, 120), (25, 95), None, 1.0, 1.0, "medsurg", 2.5, 2.5, 0, 0.08, 0.0, "Weakness / Fatigue", 3.8, "endo"),
    "E11.65": ("Type 2 diabetes mellitus with hyperglycemia", "Endocrine", "Y", (28, 6, 22, 60), (25, 95), None, 1.05, 1.0, "medsurg", 3.2, 3.0, 0, 0.14, 0.005, "High Blood Sugar", 3.0, "endo"),
    "E86.0": ("Dehydration", "Endocrine", "N", (18, 18, 45, 0), (0, 99), None, 1.1, 1.35, "medsurg", 2.6, 2.4, 0, 0.09, 0.005, "Vomiting / Dizziness", 3.5, None),
    "E87.1": ("Hypo-osmolality and hyponatremia", "Endocrine", "N", (22, 4, 0, 0), (55, 99), None, 1.0, 1.2, "medsurg", 3.5, 3.2, 0, 0.13, 0.01, "Confusion", 2.8, None),
    "N17.9": ("Acute kidney failure, unspecified", "Genitourinary", "N", (40, 4, 0, 0), (40, 99), None, 1.1, 1.15, "medsurg", 4.4, 3.9, 0, 0.18, 0.03, "Weakness / Fatigue", 2.8, None),
    "K92.2": ("Gastrointestinal hemorrhage, unspecified", "Digestive", "N", (34, 4, 0, 0), (35, 99), None, 1.0, 1.0, "medsurg_icu", 3.9, 3.5, 0, 0.14, 0.03, "Bloody Stool / Vomiting Blood", 2.2, None),
    "R10.9": ("Unspecified abdominal pain", "Symptoms & Signs", "N", (0, 28, 110, 0), (2, 95), None, 1.0, 1.0, "obs", 1.0, 1.0, 0, 0.04, 0.0, "Abdominal Pain", 3.3, None),
    "K35.80": ("Unspecified acute appendicitis", "Digestive", "N", (26, 0, 0, 0), (5, 75), None, 1.0, 1.05, "surgery", 2.2, 2.0, 17500, 0.04, 0.001, "Abdominal Pain", 2.7, None),
    "K80.20": ("Calculus of gallbladder without cholecystitis without obstruction", "Digestive", "N", (22, 6, 18, 8), (20, 90), None, 1.0, 1.0, "surgery", 2.1, 1.9, 15800, 0.05, 0.001, "Abdominal Pain", 3.2, "primary"),
    "K56.609": ("Unspecified intestinal obstruction", "Digestive", "N", (20, 0, 0, 0), (30, 99), None, 1.0, 1.0, "surgery", 5.3, 4.6, 9500, 0.14, 0.02, "Abdominal Pain / Vomiting", 2.6, None),
    "S72.001A": ("Fracture of unspecified part of neck of right femur, initial encounter", "Injury", "N", (32, 0, 0, 0), (55, 99), None, 1.2, 0.95, "ortho", 5.2, 4.6, 31500, 0.11, 0.025, "Fall / Hip Pain", 2.5, None),
    "M17.11": ("Unilateral primary osteoarthritis, right knee", "Musculoskeletal", "Y", (34, 0, 0, 45), (45, 90), None, 0.9, 1.0, "ortho_elective", 2.1, 2.0, 38000, 0.04, 0.0, "Knee Pain", 4.5, "ortho"),
    "M16.11": ("Unilateral primary osteoarthritis, right hip", "Musculoskeletal", "Y", (28, 0, 0, 35), (45, 90), None, 0.9, 1.0, "ortho_elective", 2.2, 2.1, 40500, 0.04, 0.0, "Hip Pain", 4.5, "ortho"),
    "S06.0X0A": ("Concussion without loss of consciousness, initial encounter", "Injury", "N", (0, 6, 40, 0), (4, 95), None, 0.9, 1.25, "obs", 1.0, 1.0, 0, 0.02, 0.0, "Head Injury", 3.2, None),
    "S93.401A": ("Sprain of unspecified ligament of right ankle, initial encounter", "Injury", "N", (0, 0, 70, 12), (8, 80), None, 0.85, 1.35, "obs", 1.0, 1.0, 0, 0.0, 0.0, "Ankle Injury", 4.6, "ortho"),
    "M54.50": ("Low back pain, unspecified", "Musculoskeletal", "N", (0, 4, 70, 45), (18, 90), None, 1.0, 1.0, "obs", 1.0, 1.0, 0, 0.0, 0.0, "Back Pain", 4.2, "primary"),
    "O80": ("Encounter for full-term uncomplicated delivery", "Pregnancy", "N", (78, 0, 0, 0), (17, 43), "F", 1.0, 1.0, "ob", 2.0, 2.0, 4200, 0.015, 0.0, "Labor", 3.0, None),
    "O82": ("Encounter for cesarean delivery without indication", "Pregnancy", "N", (36, 0, 0, 0), (18, 45), "F", 1.0, 1.0, "ob", 3.0, 2.9, 9100, 0.02, 0.0, "Labor", 3.0, None),
    "J21.9": ("Acute bronchiolitis, unspecified", "Respiratory", "N", (16, 6, 22, 0), (0, 2), None, 3.0, 0.3, "peds", 2.5, 2.4, 0, 0.05, 0.0, "Cough / Trouble Breathing", 3.0, None),
    "A08.4": ("Viral intestinal infection, unspecified", "Infectious", "N", (5, 6, 45, 8), (0, 80), None, 1.6, 0.8, "medsurg", 2.3, 2.1, 0, 0.03, 0.0, "Vomiting / Diarrhea", 4.0, "primary"),
    "H66.90": ("Otitis media, unspecified, unspecified ear", "Ear", "N", (0, 0, 30, 45), (0, 12), None, 1.7, 0.7, "peds", 1.0, 1.0, 0, 0.0, 0.0, "Ear Pain / Fever", 4.6, "peds"),
    "F32.9": ("Major depressive disorder, single episode, unspecified", "Mental & Behavioral", "Y", (32, 0, 28, 30), (13, 85), None, 1.15, 0.9, "psych", 7.4, 6.8, 0, 0.12, 0.0, "Depression / Suicidal Thoughts", 2.6, "primary"),
    "F20.9": ("Schizophrenia, unspecified", "Mental & Behavioral", "Y", (20, 0, 10, 0), (18, 70), None, 1.0, 1.0, "psych", 9.8, 8.9, 0, 0.18, 0.0, "Psychiatric Evaluation", 2.7, None),
    "F10.239": ("Alcohol dependence with withdrawal, unspecified", "Mental & Behavioral", "Y", (24, 4, 16, 0), (21, 80), None, 1.05, 1.0, "medsurg", 4.1, 3.6, 0, 0.17, 0.005, "Tremors / Alcohol Withdrawal", 2.8, None),
    "F41.9": ("Anxiety disorder, unspecified", "Mental & Behavioral", "Y", (0, 4, 40, 30), (14, 85), None, 1.1, 0.95, "obs", 1.0, 1.0, 0, 0.0, 0.0, "Anxiety / Palpitations", 3.9, "primary"),
    "C34.90": ("Malignant neoplasm of unspecified part of unspecified bronchus or lung", "Neoplasms", "Y", (18, 0, 0, 30), (45, 92), None, 1.0, 1.0, "onc", 5.6, 5.0, 0, 0.21, 0.06, "Shortness of Breath", 2.5, "onc"),
    "C50.911": ("Malignant neoplasm of unspecified site of right female breast", "Neoplasms", "Y", (8, 0, 0, 34), (30, 90), "F", 1.0, 1.0, "onc", 3.2, 3.0, 12000, 0.08, 0.005, "Breast Mass", 3.0, "onc"),
    "Z51.11": ("Encounter for antineoplastic chemotherapy", "Factors Influencing Health", "N", (0, 0, 0, 120), (25, 90), None, 1.0, 1.0, "onc", 1.0, 1.0, 0, 0.0, 0.0, "Scheduled Treatment", 5.0, "onc"),
    "D70.1": ("Agranulocytosis secondary to cancer chemotherapy", "Blood", "N", (16, 0, 0, 0), (25, 90), None, 1.0, 1.0, "onc", 5.0, 4.5, 0, 0.20, 0.03, "Fever", 2.0, None),
    "G40.909": ("Epilepsy, unspecified, not intractable, without status epilepticus", "Nervous System", "Y", (14, 8, 30, 15), (2, 90), None, 1.0, 1.0, "neuro", 2.8, 2.5, 0, 0.09, 0.005, "Seizure", 2.6, "primary"),
    "G43.909": ("Migraine, unspecified, not intractable, without status migrainosus", "Nervous System", "Y", (0, 4, 55, 20), (10, 70), None, 1.0, 1.0, "obs", 1.0, 1.0, 0, 0.0, 0.0, "Headache", 3.7, "primary"),
    "R55": ("Syncope and collapse", "Symptoms & Signs", "N", (8, 30, 35, 0), (12, 99), None, 1.0, 1.1, "obs", 2.1, 2.0, 0, 0.07, 0.0, "Fainting", 3.0, None),
    "Z00.00": ("Encounter for general adult medical examination without abnormal findings", "Factors Influencing Health", "N", (0, 0, 0, 160), (18, 95), None, 0.9, 1.0, "medsurg", 1.0, 1.0, 0, 0.0, 0.0, "Annual Exam", 5.0, "primary"),
    "Z00.129": ("Encounter for routine child health examination without abnormal findings", "Factors Influencing Health", "N", (0, 0, 0, 90), (0, 17), None, 0.9, 1.2, "peds", 1.0, 1.0, 0, 0.0, 0.0, "Well Child Visit", 5.0, "peds"),
    "Z23": ("Encounter for immunization", "Factors Influencing Health", "N", (0, 0, 0, 70), (0, 99), None, 2.6, 0.4, "medsurg", 1.0, 1.0, 0, 0.0, 0.0, "Immunization", 5.0, "primary"),
    "E78.5": ("Hyperlipidemia, unspecified", "Endocrine", "Y", (0, 0, 0, 75), (30, 95), None, 1.0, 1.0, "medsurg", 1.0, 1.0, 0, 0.0, 0.0, "Follow-up Visit", 5.0, "primary"),
}

# Chronic condition flags assigned to patients and the dx codes that "belong" to each.
CHRONIC_LINKS = {
    "HF": ["I50.9"],
    "COPD": ["J44.1"],
    "DM": ["E11.9", "E11.65"],
    "AFIB": ["I48.91"],
    "HTN": ["I10"],
    "CKD": ["N17.9"],
    "ASTHMA": ["J45.909"],
    "CANCER": ["C34.90", "C50.911", "Z51.11", "D70.1"],
    "PSYCH": ["F32.9", "F20.9", "F41.9"],
    "ETOH": ["F10.239"],
    "OA": ["M17.11", "M16.11"],
    "SEIZURE": ["G40.909"],
}

# Unit daily room & board charge by inpatient route
ROOM_RATE = {
    "medsurg": 3200, "cardiac": 4800, "icu": 8900, "ortho": 3600, "onc": 4200, "neuro": 4500,
    "ob": 3500, "peds": 3600, "psych": 2400, "obs": 2900,
}

LAB_TESTS = [
    # code, name, units, ref_low, ref_high, normal_mean, normal_sd, decimals, crit_low, crit_high
    ("GLU", "Glucose", "mg/dL", 70, 99, 96, 14, 0, 40, 500),
    ("NA", "Sodium", "mmol/L", 136, 145, 139.5, 2.4, 0, 120, 160),
    ("K", "Potassium", "mmol/L", 3.5, 5.1, 4.2, 0.35, 1, 2.8, 6.2),
    ("CREAT", "Creatinine", "mg/dL", 0.6, 1.3, 0.95, 0.2, 2, None, None),
    ("BUN", "Blood Urea Nitrogen", "mg/dL", 7, 20, 14, 4, 0, None, None),
    ("WBC", "White Blood Cell Count", "K/uL", 4.5, 11.0, 7.4, 1.7, 1, 1.0, 30.0),
    ("HGB", "Hemoglobin", "g/dL", 12.0, 17.5, 13.9, 1.3, 1, 7.0, 20.0),
    ("PLT", "Platelet Count", "K/uL", 150, 400, 255, 55, 0, 20, 1000),
    ("TROP", "Troponin I, High Sensitivity", "ng/L", 0, 34, 6, 4, 0, None, None),
    ("BNP", "B-Type Natriuretic Peptide", "pg/mL", 0, 100, 42, 22, 0, None, None),
    ("LACT", "Lactate", "mmol/L", 0.5, 2.0, 1.2, 0.3, 1, None, 4.0),
    ("A1C", "Hemoglobin A1c", "%", 4.0, 5.6, 5.3, 0.25, 1, None, None),
    ("LDL", "LDL Cholesterol", "mg/dL", 0, 99, 112, 28, 0, None, None),
    ("TSH", "Thyroid Stimulating Hormone", "mIU/L", 0.4, 4.0, 1.9, 0.8, 2, None, None),
    ("INR", "Prothrombin Time INR", "ratio", 0.8, 1.2, 1.0, 0.08, 1, None, 5.0),
]

MEDICATIONS = [
    # name, class, dose, route, frequency, unit_cost, high_alert, uses (dx categories or 'any')
    ("Ceftriaxone", "Antibiotic", "1 g", "IV", "Q24H", 3.20, "N"),
    ("Piperacillin-Tazobactam", "Antibiotic", "4.5 g", "IV", "Q6H", 14.50, "N"),
    ("Vancomycin", "Antibiotic", "1250 mg", "IV", "Q12H", 9.80, "N"),
    ("Azithromycin", "Antibiotic", "500 mg", "PO", "Daily", 1.10, "N"),
    ("Cephalexin", "Antibiotic", "500 mg", "PO", "Q6H", 0.45, "N"),
    ("Amoxicillin", "Antibiotic", "400 mg/5 mL", "PO", "BID", 0.70, "N"),
    ("Heparin", "Anticoagulant", "5000 units", "SubQ", "Q8H", 1.60, "Y"),
    ("Enoxaparin", "Anticoagulant", "40 mg", "SubQ", "Daily", 6.75, "Y"),
    ("Apixaban", "Anticoagulant", "5 mg", "PO", "BID", 8.90, "Y"),
    ("Warfarin", "Anticoagulant", "5 mg", "PO", "Daily", 0.25, "Y"),
    ("Insulin Lispro", "Insulin", "Sliding scale", "SubQ", "AC & HS", 4.10, "Y"),
    ("Insulin Glargine", "Insulin", "20 units", "SubQ", "Daily", 11.40, "Y"),
    ("Morphine", "Opioid Analgesic", "2 mg", "IV", "Q4H PRN", 1.95, "Y"),
    ("Hydromorphone", "Opioid Analgesic", "0.5 mg", "IV", "Q3H PRN", 2.40, "Y"),
    ("Oxycodone", "Opioid Analgesic", "5 mg", "PO", "Q4H PRN", 0.60, "Y"),
    ("Acetaminophen", "Analgesic", "650 mg", "PO", "Q6H PRN", 0.08, "N"),
    ("Ketorolac", "NSAID", "15 mg", "IV", "Q6H", 1.30, "N"),
    ("Ondansetron", "Antiemetic", "4 mg", "IV", "Q6H PRN", 0.85, "N"),
    ("Furosemide", "Diuretic", "40 mg", "IV", "BID", 1.20, "N"),
    ("Metoprolol Tartrate", "Beta Blocker", "25 mg", "PO", "BID", 0.12, "N"),
    ("Diltiazem", "Calcium Channel Blocker", "10 mg", "IV", "Once", 2.75, "N"),
    ("Lisinopril", "ACE Inhibitor", "10 mg", "PO", "Daily", 0.09, "N"),
    ("Atorvastatin", "Statin", "40 mg", "PO", "Daily", 0.18, "N"),
    ("Aspirin", "Antiplatelet", "81 mg", "PO", "Daily", 0.03, "N"),
    ("Pantoprazole", "Proton Pump Inhibitor", "40 mg", "IV", "Daily", 2.10, "N"),
    ("Albuterol", "Bronchodilator", "2.5 mg", "Inhaled", "Q4H", 0.55, "N"),
    ("Ipratropium-Albuterol", "Bronchodilator", "3 mL", "Inhaled", "Q6H", 1.05, "N"),
    ("Methylprednisolone", "Corticosteroid", "40 mg", "IV", "Q8H", 3.90, "N"),
    ("Prednisone", "Corticosteroid", "40 mg", "PO", "Daily", 0.20, "N"),
    ("Sodium Chloride 0.9%", "IV Fluid", "1000 mL", "IV", "Continuous", 1.75, "N"),
    ("Lactated Ringer's", "IV Fluid", "1000 mL", "IV", "Continuous", 1.90, "N"),
    ("Haloperidol", "Antipsychotic", "5 mg", "IM", "Once", 1.15, "N"),
    ("Olanzapine", "Antipsychotic", "10 mg", "PO", "Daily", 0.65, "N"),
    ("Lorazepam", "Benzodiazepine", "1 mg", "IV", "Q4H PRN", 1.40, "Y"),
    ("Sertraline", "Antidepressant", "50 mg", "PO", "Daily", 0.10, "N"),
    ("Levetiracetam", "Anticonvulsant", "500 mg", "IV", "BID", 2.95, "N"),
    ("Oxytocin", "Uterotonic", "30 units", "IV", "Continuous", 3.10, "Y"),
    ("Carboplatin", "Chemotherapy", "AUC 5", "IV", "Once", 185.00, "Y"),
    ("Paclitaxel", "Chemotherapy", "175 mg/m2", "IV", "Once", 240.00, "Y"),
    ("Filgrastim", "Growth Factor", "480 mcg", "SubQ", "Daily", 310.00, "N"),
    ("Influenza Vaccine", "Vaccine", "0.5 mL", "IM", "Once", 24.00, "N"),
    ("Thiamine", "Vitamin", "100 mg", "IV", "Daily", 0.95, "N"),
    ("Potassium Chloride", "Electrolyte", "20 mEq", "PO", "Once", 0.35, "Y"),
]

CHIEF_COMPLAINT_EXTRA = ["Fall", "Laceration", "Dizziness", "Rash", "Nausea"]

SURVEY_COMMENTS_POS = [
    "The nurses were wonderful and very attentive.",
    "Dr. explained everything clearly. Excellent care.",
    "Staff were kind and respectful the whole stay.",
    "Room was clean and the food was better than expected.",
    "Excellent communication from the care team.",
    "Discharge instructions were clear and easy to follow.",
    "Thank you to the night shift nurses for the great care.",
    "Everyone was professional and caring.",
]
SURVEY_COMMENTS_NEG = [
    "Waited too long for the call light to be answered at night.",
    "Very noisy hallway, hard to sleep.",
    "Long wait in the emergency room before getting a bed.",
    "Nobody explained my new medications before discharge.",
    "Bathroom was not cleaned during my stay.",
    "Felt rushed by the doctor, questions not answered.",
    "Parking was expensive and hard to find.",
    "Food was cold and arrived late.",
    "Billing office was rude on the phone.",
]
SURVEY_COMMENTS_MIXED = [
    "Good nurses but the wait for discharge was very long.",
    "Care was fine, room was noisy at night.",
    "Doctors were great; food needs improvement.",
    "Clean room, but the call light took too long.",
]

VENDORS = ["Northstar Medical Supply", "Apex Surgical Co.", "CareLine Distributors", "Pinnacle Biomedical", "Harbor Health Products", "Summit Orthopedic Systems", "Keystone Lab Solutions"]

# Supply catalog templates: (category, description, unit, base_cost, vendor_idx, perishable, variants)
SUPPLY_TEMPLATES = [
    ("PPE", "Nitrile Exam Gloves {v} (Box/100)", "Box", 8.40, 0, False, ["Small", "Medium", "Large", "X-Large"]),
    ("PPE", "Isolation Gown, Level 2 {v}", "Case", 54.00, 2, False, ["Regular", "X-Large"]),
    ("PPE", "N95 Respirator {v} (Box/20)", "Box", 31.50, 0, True, ["Small", "Regular"]),
    ("PPE", "Procedure Mask, Ear Loop (Box/50)", "Box", 6.25, 2, False, [""]),
    ("PPE", "Face Shield, Full Length (Box/24)", "Box", 18.90, 4, False, [""]),
    ("PPE", "Shoe Covers, Non-Skid (Case/300)", "Case", 41.00, 4, False, [""]),
    ("IV Supplies", "IV Catheter {v} x 1.16 in", "Box", 72.00, 3, True, ["14G", "16G", "18G", "20G", "22G", "24G"]),
    ("IV Supplies", "Primary IV Tubing Set, {v}", "Case", 96.00, 3, True, ["60 Drop", "15 Drop", "Pump Compatible"]),
    ("IV Supplies", "Saline Flush Syringe {v} (Box/30)", "Box", 9.80, 0, True, ["3 mL", "5 mL", "10 mL"]),
    ("IV Supplies", "IV Start Kit with Chlorhexidine", "Each", 2.15, 2, True, [""]),
    ("IV Supplies", "IV Extension Set {v}", "Each", 1.85, 3, True, ["7 in", "30 in"]),
    ("Wound Care", "Gauze Sponge {v} Sterile (Pack/2)", "Pack", 0.42, 0, True, ["2x2", "4x4"]),
    ("Wound Care", "Foam Dressing {v}", "Each", 4.95, 4, True, ["3x3", "4x4", "6x6"]),
    ("Wound Care", "Transparent Film Dressing {v}", "Each", 0.88, 4, True, ["2.4x2.8", "4x4.75"]),
    ("Wound Care", "Abdominal Pad 8x10 Sterile", "Each", 0.61, 0, True, [""]),
    ("Wound Care", "Silver Alginate Dressing {v}", "Each", 12.40, 3, True, ["2x2", "4x5"]),
    ("Wound Care", "Elastic Bandage {v}", "Each", 1.35, 2, False, ["2 in", "4 in", "6 in"]),
    ("Respiratory", "Nasal Cannula, {v}", "Each", 0.95, 0, False, ["Adult", "Pediatric", "Infant"]),
    ("Respiratory", "Non-Rebreather Mask, {v}", "Each", 2.20, 0, False, ["Adult", "Pediatric"]),
    ("Respiratory", "Nebulizer Kit with Mouthpiece", "Each", 2.85, 2, False, [""]),
    ("Respiratory", "Suction Canister {v}", "Each", 3.10, 2, False, ["1200 mL", "1500 mL", "3000 mL"]),
    ("Respiratory", "Yankauer Suction Tip", "Each", 0.74, 2, False, [""]),
    ("Respiratory", "Endotracheal Tube {v} Cuffed", "Each", 6.80, 3, True, ["6.5 mm", "7.0 mm", "7.5 mm", "8.0 mm"]),
    ("Lab Supplies", "Blood Culture Bottle, {v}", "Each", 7.95, 6, True, ["Aerobic", "Anaerobic", "Pediatric"]),
    ("Lab Supplies", "Vacutainer Tube {v} (Box/100)", "Box", 24.50, 6, True, ["Lavender 4 mL", "Light Blue 2.7 mL", "Gold SST 5 mL", "Green 4 mL", "Gray 4 mL"]),
    ("Lab Supplies", "Safety Lancet {v} (Box/200)", "Box", 21.00, 6, True, ["21G", "23G"]),
    ("Lab Supplies", "Specimen Container 4 oz Sterile", "Each", 0.38, 6, False, [""]),
    ("Lab Supplies", "Glucose Test Strips (Box/50)", "Box", 18.75, 6, True, [""]),
    ("Lab Supplies", "Urinalysis Reagent Strips (Bottle/100)", "Bottle", 29.00, 6, True, [""]),
    ("Surgical", "Absorbable Suture {v}", "Box", 118.00, 1, True, ["2-0", "3-0", "4-0"]),
    ("Surgical", "Nylon Suture {v}", "Box", 64.00, 1, True, ["3-0", "4-0", "5-0"]),
    ("Surgical", "Surgical Blade #{v} (Box/50)", "Box", 26.50, 1, False, ["10", "11", "15"]),
    ("Surgical", "Skin Stapler 35W", "Each", 14.25, 1, False, [""]),
    ("Surgical", "Laparoscopic Trocar {v}", "Each", 88.00, 1, False, ["5 mm", "12 mm"]),
    ("Surgical", "Universal Surgical Drape Pack", "Each", 39.00, 1, False, [""]),
    ("Surgical", "Surgical Gown, Sterile {v}", "Each", 4.60, 1, False, ["Large", "X-Large"]),
    ("Orthopedic Implants", "Hip Femoral Stem, Size {v}", "Each", 2650.00, 5, False, ["10", "11", "12", "13", "14"]),
    ("Orthopedic Implants", "Knee Tibial Tray, Size {v}", "Each", 1980.00, 5, False, ["3", "4", "5", "6"]),
    ("Orthopedic Implants", "Cortical Screw 3.5 mm x {v}", "Each", 46.00, 5, False, ["20 mm", "30 mm", "40 mm"]),
    ("Orthopedic Implants", "Locking Plate {v}", "Each", 610.00, 5, False, ["6-Hole", "8-Hole", "10-Hole"]),
    ("Cardiac", "Drug-Eluting Stent {v}", "Each", 1450.00, 3, True, ["2.5 x 18 mm", "3.0 x 18 mm", "3.5 x 23 mm"]),
    ("Cardiac", "Guide Wire 0.035 in x {v}", "Each", 38.00, 3, True, ["150 cm", "260 cm"]),
    ("Cardiac", "ECG Electrodes (Pack/30)", "Pack", 5.10, 0, True, [""]),
    ("Cardiac", "Defibrillator Pads, {v}", "Pair", 42.00, 3, True, ["Adult", "Pediatric"]),
    ("Linens", "Patient Gown, {v}", "Dozen", 36.00, 4, False, ["Standard", "Bariatric", "Pediatric"]),
    ("Linens", "Bath Blanket", "Dozen", 48.00, 4, False, [""]),
    ("Linens", "Fitted Bed Sheet", "Dozen", 52.00, 4, False, [""]),
    ("Patient Care", "Foley Catheter Tray {v}", "Each", 11.80, 0, True, ["14 Fr", "16 Fr", "18 Fr"]),
    ("Patient Care", "Incontinence Brief, {v} (Pack/18)", "Pack", 12.60, 2, False, ["Medium", "Large", "X-Large"]),
    ("Patient Care", "Thermometer Probe Covers (Box/200)", "Box", 9.40, 0, False, [""]),
    ("Patient Care", "Disposable Washcloths (Pack/8)", "Pack", 1.90, 2, False, [""]),
    ("Patient Care", "Sequential Compression Sleeve, {v}", "Pair", 33.00, 3, False, ["Knee Length", "Thigh Length"]),
    ("Patient Care", "Hand Sanitizer Foam 1000 mL", "Each", 11.20, 4, True, [""]),
    ("Pharmacy Supplies", "Oral Syringe {v} (Box/100)", "Box", 14.00, 0, False, ["1 mL", "5 mL", "10 mL"]),
    ("Pharmacy Supplies", "Pill Crusher Pouches (Box/1000)", "Box", 22.00, 2, False, [""]),
    ("Pharmacy Supplies", "IV Bag Label Roll", "Roll", 8.75, 2, False, [""]),
]
