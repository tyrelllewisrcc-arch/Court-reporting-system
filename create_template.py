#!/usr/bin/env python3
"""
CourtCIMS Blank Template Generator — CourtCIMSV4
Recreates the Court Case Information Management System as a blank operational
file with all 11 sheets, structured Excel tables, XLOOKUP/FILTER formulas,
data validations, and the full ChargeMap / Helper reference data.

Run:  python3 create_template.py
Output: CourtCIMS_Blank_Template.xlsx
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ──────────────────────────────────────────────────────────────────────────────
# STYLE HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

FILL_DARK_BLUE  = fill("1F4E79")
FILL_MED_BLUE   = fill("2E75B6")
FILL_LIGHT_BLUE = fill("DEEAF1")
FILL_TEAL       = fill("00B0C4")
FILL_DARK_TEAL  = fill("1F6B75")
FILL_NAVY       = fill("003366")
FILL_GREEN      = fill("375623")
FILL_YELLOW     = fill("FFE699")
FILL_ORANGE     = fill("C65911")

FONT_WHITE_BOLD = Font(name="Calibri", color="FFFFFF", bold=True, size=10)
FONT_WHITE      = Font(name="Calibri", color="FFFFFF", bold=False, size=10)
FONT_DARK_BOLD  = Font(name="Calibri", color="1F1F1F", bold=True, size=10)
FONT_DARK       = Font(name="Calibri", color="1F1F1F", bold=False, size=10)
FONT_TITLE      = Font(name="Calibri", color="1F4E79", bold=True, size=14)
FONT_SMALL_BOLD = Font(name="Calibri", color="FFFFFF", bold=True, size=9)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center")

_thin = Side(style="thin", color="9DC3E6")
BORDER_THIN = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def styled_cell(ws, row, col, value=None, fill_=None, font=None, align=None, border=None, num_fmt=None):
    cell = ws.cell(row=row, column=col, value=value)
    if fill_:   cell.fill      = fill_
    if font:    cell.font      = font
    if align:   cell.alignment = align
    if border:  cell.border    = border
    if num_fmt: cell.number_format = num_fmt
    return cell


def header_row(ws, row, columns, fill_=None, font=None, start_col=1):
    """Write a list of header labels across a row."""
    f = fill_ or FILL_DARK_BLUE
    fn = font  or FONT_WHITE_BOLD
    for ci, label in enumerate(columns, start=start_col):
        cell = ws.cell(row=row, column=ci, value=label)
        cell.fill      = f
        cell.font      = fn
        cell.alignment = ALIGN_CENTER
        cell.border    = BORDER_THIN


def make_table(ws, ref, name, style="TableStyleMedium2"):
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name=style,
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True,  showColumnStripes=False,
    )
    ws.add_table(tab)
    return tab


def col_widths(ws, width_map):
    """width_map: {col_number: width}"""
    for col, w in width_map.items():
        ws.column_dimensions[get_column_letter(col)].width = w


def add_dv(ws, dv_type, formula1, sqref):
    dv = DataValidation(type=dv_type, formula1=formula1,
                        allow_blank=True, sqref=sqref, showErrorMessage=True)
    dv.error      = "Please select from the list."
    dv.errorTitle = "Invalid Input"
    ws.add_data_validation(dv)

# ──────────────────────────────────────────────────────────────────────────────
# REFERENCE DATA
# ──────────────────────────────────────────────────────────────────────────────

# Columns shared by most case sheets (20 cols = A:T)
CASE_COLS_20 = [
    "CourtBookNo", "DateOfOccurence", "ArraignmentDate", "ArrestingOfficer",
    "Station", "ComplainantVictim", "Defendant", "Sex", "Age", "DOB",
    "Address", "CaseFileReceived", "Charge", "NextHearingDate", "Remark",
    "Further Remark", "Outcome", "ConcludedDate", "EntryDate", "DateUpdated",
]

# 15-column subset used by Monthly_Returns / Daily_Exports data tables
CASE_COLS_15 = [
    "CourtBookNo", "DateOfOccurence", "ArraignmentDate", "ArrestingOfficer",
    "Station", "ComplainantVictim", "Defendant", "Sex", "Age", "DOB",
    "Address", "Charge", "Further Remark", "Outcome", "ConcludedDate",
]

OUTCOME_LIST = [
    "P.I Conducted", "Struck Out", "Bond Over", "Withdrawn", "Dismissed",
    "Conditional Discharged", "Convicted", "Acquitted", "Adjourned",
    "Set for trial", "Set for P.I.", "Bench Warrant",
]

OUTCOME_STATUS_MAP = [
    ("P.I Conducted",        "P.I Conducted"),
    ("Struck Out",           "Struck Out"),
    ("Bond Over",            "Bond Over"),
    ("Withdrawn",            "Withdrawn"),
    ("Dismissed",            "Dismissed"),
    ("Conditional Discharged","Conditional Discharged"),
    ("Convicted",            "Convicted"),
    ("Acquitted",            "Acquitted"),
    ("Adjourned",            "Active"),
    ("Set for trial",        "set for trial"),
    ("Set for P.I.",         "Active"),
    ("Bench Warrant",        "Active"),
]

STAT_PATTERNS = [
    ("Drugs",                    "Dangerous Drugs"),
    ("Pipe",                     "Dangerous Drugs"),
    ("Utensil",                  "Dangerous Drugs"),
    ("Cannabis",                 "Dangerous Drugs"),
    ("Firearm",                  "Against Firearms Act"),
    ("Ammunition",               "Against Firearms Act"),
    ("Gun License",              "Against Firearms Act"),
    ("Police",                   "Against Police Act"),
    ("Harm",                     "Summary Jurisdiction Offences"),
    ("Drove Motor Vehicle",      "Other Offences"),
    ("Used Motor Vehicle",       "Other Offences"),
    ("Unregistered Motor Vehicle","Other Offences"),
    ("Unlicensed Motor Vehicle", "Other Offences"),
    ("Driver's License",         "Other Offences"),
    ("third party risk insurance","Other Offences"),
    ("Without Due Care",         "Other Offences"),
    ("Reckless Driving",         "Other Offences"),
    ("Driving Whilst Unfit",     "Other Offences"),
    ("Fail to Provide Specimen", "Other Offences"),
    ("Assault",                  "Summary Jurisdiction Offences"),
    ("Murder",                   "Other Offences"),
    ("Robbery",                  "Summary Jurisdiction Offences"),
    ("Wounding",                 "Summary Jurisdiction Offences"),
    ("Damage to",                "Summary Jurisdiction Offences"),
    ("Obtaining Property",       "Summary Jurisdiction Offences"),
]

ARRESTING_OFFICERS = [
    "Abi Taca DC 2393", "Adan Uh CPL 1672", "Adolfo Zetina PC 2659",
    "Adrian Chavez PC 326", "Aisha Hall PC #800", "Alba Sho PC 2621",
    "Albert Lawrence PC 1204", "Alexie Muschamp PC 2520",
    "Alfonso Chuc CPL 176", "Allan Domingo PC 2029", "Allan Woods SGT 1037",
    "Avian Crawford PC 1730", "Charles Garcia PC 1047",
    "Claude Pitts SGT 920", "Delson Arzu DC 135", "Desmond Nunez PC 1929",
    "Efren Nasario PC 949", "Erin Pate PC 2272", "Eyon Valerio CPL 404",
    "Gabriel Cruz PC 471", "Hector Esquivel PC 1597",
    "Hilberbrant Tillett PC 2638", "Hubert Bermudez DC 1440",
    "Ian Ferrera PC 2743", "John Bruhier PC 253", "Justin Mcfoy PC 2861",
    "Keeshawn Gallego PC 1636", "Kersha Lacayo PC 2501",
    "Kevin Rodriguez PC 2447", "Leonel Perez PC 1786",
    "Leticia Moguel SGT 824", "Nigel Castillo PC #593",
    "Oswald Young PC 2656", "Roy Flores PC 950", "Shamlit Chan PC 2183",
    "Yann S. Chin DC 1454",
]

# (ChargeName, PriorityRank, AnnexureNature, AnnexureSubCategory)
CHARGE_MAP_DATA = [
    ("Aggrated Assault",                    75,  "AGAINST THE PERSON",     "Aggrated Assault"),
    ("Aggravated Assault",                  75,  "AGAINST THE PERSON",     "Aggravated Assault"),
    ("Asault a Police Officer",             50,  "AGAINST THE PERSON",     "Aggravated Assault"),
    ("Assault of A Child Under the age of 16 by Penetration", 90, "AGAINST PUBLIC MORALITY", "Sexual Assault"),
    ("Assaulting A Police Officer",         60,  "AGAINST LAWFUL AUTHORITY","Other"),
    ("Attempt Murder",                      98,  "AGAINST THE PERSON",     "Attempted Murder"),
    ("Attempt Robbery",                     85,  "AGAINST THE PERSON",     "Other"),
    ("Being Member of A Gang",              85,  "OTHERS",                 "Other"),
    ("Being the Member of a Gang",          85,  "OTHERS",                 "Other"),
    ("Breach of Court Order",               50,  "AGAINST THE PERSON",     "Other"),
    ("Breach of Protection Order",          50,  "AGAINST THE PERSON",     "Other"),
    ("Burglary",                            72,  "AGAINST PROPERTY",       "Burglary"),
    ("Burglary Escaping Lawful Custody",    72,  "AGAINST PROPERTY",       "Burglary"),
    ("Causing Deaath By Careless Conduct",  50,  "AGAINST THE PERSON",     "Manslaughter"),
    ("Child Neglect",                       50,  "AGAINST PUBLIC MORALITY","Other"),
    ("Common Assault",                      60,  "AGAINST THE PERSON",     "Common Assault"),
    ("Comtempt of Judicial Order (7 counts)",50, "AGAINST LAWFUL AUTHORITY","Other"),
    ("Damage to Property",                  55,  "AGAINST PROPERTY",       "Damage to Property"),
    ("Dangerous Harm",                      75,  "AGAINST THE PERSON",     "Grievous Harm"),
    ("Discharging Firearm in Public",       98,  "OTHERS",                 "Other"),
    ("Disorderly Conduct",                  45,  "OTHERS",                 "Other"),
    ("Driving Whilst Unfit",                65,  "OTHERS",                 "Other"),
    ("Drove Motor Vehicle not covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle Whilst Being Disqualified From Holding or Obtaining a Belize Driver's License", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle whilst not Being Covered by Third Party Risk Insurance", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle Wihtout a Valid License", 65, "OTHERS",          "Other"),
    ("Drove Motor Vehicle Wihtuot Due Care and Attention", 65, "AGAINST THE PERSON", "Manslaughter"),
    ("Drove Motor Vehicle Without a Belize Driver's License", 65, "OTHERS","Other"),
    ("DroveMotor Vehicle without a valid driver's license", 65, "OTHERS",  "Other"),
    ("Drove Motor Vehicle without a Protective Helmet.", 65, "OTHERS",     "Other"),
    ("Drove Motor Vehicle Without a Valid Driver's License", 65, "OTHERS", "Other"),
    ("Drove Motor vehicle Without a Valid Driver's License", 65, "OTHERS", "Other"),
    ("Drove motor vehicle without a valid driver's license", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle Without a Valid Driving License", 65, "OTHERS",  "Other"),
    ("Drove Motor Vehicle Without a Valid License", 65, "OTHERS",          "Other"),
    ("Drove Motor Vehicle without a valid License", 65, "OTHERS",          "Other"),
    ("Drove Motor vehicle without a Valid License", 65, "OTHERS",          "Other"),
    ("Drove Motor Vehicle Without a Valid License in Respect to Class", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle Without a Valid Without A Valid Driver's License", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle Without being the holder of a Valid Driver's License", 65, "OTHERS", "Other"),
    ("Drove motor vehicle without being the holder of a Valid Driver's License", 65, "OTHERS", "Other"),
    ("Drove Motor Vehicle Without Due Care and Attention", 65, "OTHERS",   "Other"),
    ("Drove Motor Vehicle without due care and attention", 65, "OTHERS",   "Other"),
    ("Drove Unlicensed Motor Vehicle",      65,  "OTHERS",                 "Other"),
    ("Engaging in Sexual Activity in the Presence of a Child", 50, "AGAINST PUBLIC MORALITY", "Other"),
    ("Escape From Lawful Custody",          50,  "AGAINST LAWFUL AUTHORITY","Escape"),
    ("Escaping Lawful Custody",             50,  "OTHERS",                 "Other"),
    ("Exposing Person in Public",           45,  "OTHERS",                 "Other"),
    ("Exposing Person in Public (2x)",      45,  "OTHERS",                 "Other"),
    ("Fail to give way when changing",      65,  "OTHERS",                 "Other"),
    ("Fail to Provide Specimen",            65,  "OTHERS",                 "Other"),
    ("Fail to provide Specimen",            65,  "OTHERS",                 "Other"),
    ("Failded to display Insurance Disk",   65,  "OTHERS",                 "Other"),
    ("Failed to Provide Specimen",          65,  "OTHERS",                 "Other"),
    ("Failed to Signal When Changing Direction", 65, "OTHERS",             "Other"),
    ("Failed to Signal when changing direction", 65, "OTHERS",             "Other"),
    ("Failure to provide Urine Specimen",   65,  "OTHERS",                 "Other"),
    ("Following Passenger About",           50,  "AGAINST THE PERSON",     "Other"),
    ("Found Drunk",                         45,  "OTHERS",                 "Other"),
    ("Going Equipped",                      80,  "OTHERS",                 "Other"),
    ("Greivous Harm",                       50,  "AGAINST THE PERSON",     "Other"),
    ("Grevious Harm",                       50,  "AGAINST THE PERSON",     "Grievous Harm"),
    ("Grievous Harm",                       78,  "AGAINST THE PERSON",     "Other"),
    ("Handling Stolen Goods",               50,  "AGAINST THE PERSON",     "Handling Stolen Goods"),
    ("Harm",                                50,  "AGAINST THE PERSON",     "Harm"),
    ("harm",                                50,  "AGAINST THE PERSON",     "Harm"),
    ("Kept Ammunition Without A Gun License",95, "OTHERS",                 "Other"),
    ("Kept Ammunition Without a Gun License",95, "OTHERS",                 "Other"),
    ("Kept Ammunition without a Gun License",95, "OTHERS",                 "Other"),
    ("Kept ammunition without a Gun License",95, "OTHERS",                 "Other"),
    ("Kept Ammunition Withut a Gun License", 95, "OTHERS",                 "Other"),
    ("Kept Firearm Without a Gun a License", 95, "OTHERS",                 "Other"),
    ("Kept Firearm Without A Gun License",   95, "OTHERS",                 "Other"),
    ("Kept Firearm Without a Gun License",   95, "OTHERS",                 "Other"),
    ("Kept Firearm without a Gun license",   95, "OTHERS",                 "Other"),
    ("Kept Prohibited Ammunition",           90, "OTHERS",                 "Other"),
    ("Kept Prohibited Firearm",             100, "OTHERS",                 "Other"),
    ("Loitering",                            45, "OTHERS",                 "Other"),
    ("Manslaughter By Negligence",           95, "AGAINST THE PERSON",     "Manslaughter"),
    ("Manslaugther by Negligence",           50, "AGAINST THE PERSON",     "Manslaughter"),
    ("Michievous Act",                       50, "OTHERS",                 "Other"),
    ("Murder",                              100, "AGAINST THE PERSON",     "Murder"),
    ("Obstruction",                          50, "AGAINST LAWFUL AUTHORITY","Other"),
    ("Obtaining Property By Deception",      80, "OTHERS",                 "Obtaining Property By Deception"),
    ("Obtaining Property by Deception (2 counts)", 80, "OTHERS",           "Obtaining Property By Deception"),
    ("Operating as a Tour Operator Withut first Obtaining a License", 50, "OTHERS", "Other"),
    ("Possession of Article with Blade",     60, "OTHERS",                 "Other"),
    ("Possession of Article With Blade",     60, "OTHERS",                 "Possession of Controlled Drugs"),
    ("Possession of Controled Drugs",        50, "OTHERS",                 "Possession of Controled Drugs"),
    ("Possession of Controlled Drugs",       75, "OTHERS",                 "Possession of Controlled Drugs"),
    ("Possession of Controlled Drugs Intent to Supply to another", 85, "OTHERS", "Possession of Controlled Drugs With Intent to Suplly to Another"),
    ("Possession of Controlled Drugs Wiht intent to Supply to Another", 85, "OTHERS", "Possession of Controlled Drugs With Intent to Suplly to Another"),
    ("Possession of Controlled Drugs With Intent to Suplly to Another", 75, "OTHERS", "Possession of Controlled Drugs With Intent to Suplly to Another"),
    ("Possession of Controlled Drugs With Intent to Supply to another", 85, "OTHERS", "Other"),
    ("Possession of Controlled Drugs With Intent to Supply to Another", 85, "OTHERS", "Possession of Controlled Drugs with Intent to Supply to Another"),
    ("Possession of Controlled Drugs with Intent to Supply to Another", 85, "OTHERS", "Possession of Controlled Drugs with Intent to Supply to Another"),
    ("Possession of Controlled Drugs With Intent to Supply to Another (18g Crack Cocaine)", 85, "OTHERS", "Other"),
    ("Possession Of Controlled Drugs With Intent to Supply to Another (20grams Cocaine)", 85, "OTHERS", "Possession of Controlled Drugs with Intent to Supply to Another"),
    ("Possession of Controlled Drugs With Intent to Supply to Another (324g Cocaine)", 85, "OTHERS", "Other"),
    ("Possession of Controlled Drugs With Intent to Supply to Another (3x)", 85, "OTHERS", "Possession of Controlled Drugs with Intent to Supply to Another"),
    ("Possession Of Controlled Drugs With Intent to Supply to Another (446grams of Cannabis)", 85, "OTHERS", "Possession of Controlled Drugs with Intent to Supply to Another"),
    ("Possession of Cotrolled Drugs",        50, "OTHERS",                 "Possession of Controlled Drugs"),
    ("Possession of Firearm With Intent to Cause unlawful Violence", 95, "AGAINST THE PERSON", "Harm"),
    ("Possession of Firearm with serial number removed", 96, "OTHERS",     "Other"),
    ("Possession of Pipe",                   62, "OTHERS",                 "Possession of Pipe"),
    ("Possession of Utensil for Smoking Controlled Drugs", 75, "OTHERS",   "Possession of Pipe"),
    ("Producing a Video Recording to Promote a Gang Related Activity", 85, "OTHERS", "Other"),
    ("Rape",                                 97, "AGAINST PUBLIC MORALITY","Rape"),
    ("Rape of a Child",                      97, "AGAINST PUBLIC MORALITY","Rape"),
    ("Rape of a Child (2 counts)",           97, "AGAINST PUBLIC MORALITY","Rape"),
    ("Reckless Driving",                     65, "AGAINST THE PERSON",     "Manslaughter"),
    ("Resist Arrest",                        55, "AGAINST LAWFUL AUTHORITY","Other"),
    ("Resist Lawful Arrest",                 50, "AGAINST LAWFUL AUTHORITY","Other"),
    ("Robbery",                              85, "AGAINST THE PERSON",     "Robbery"),
    ("Sexual Assault",                       90, "AGAINST PUBLIC MORALITY","Sexual Assault"),
    ("Sexual Assault (3x)",                  90, "AGAINST PUBLIC MORALITY","Sexual Assault"),
    ("Sexual Assault (4x)",                  90, "AGAINST PUBLIC MORALITY","Sexual Assault"),
    ("Smoking Cannabis",                     50, "OTHERS",                 "Other"),
    ("Smoking Cannabis in Public",           50, "OTHERS",                 "Other"),
    ("Taking Conveyance",                    80, "AGAINST PROPERTY",       "Theft"),
    ("Taking Conveyance Without Authority",  80, "OTHERS",                 "Taking Vehicle/Conveyance"),
    ("Taking Motor Vehicle Without Authority",80, "OTHERS",                "Taking Vehicle/Conveyance"),
    ("Theft",                                60, "AGAINST PROPERTY",       "Theft"),
    ("Theft by Taking Possession of Card",   60, "AGAINST PROPERTY",       "Theft"),
    ("Threat of Death",                      50, "AGAINST THE PERSON",     "Other"),
    ("Threats of Death",                     50, "AGAINST THE PERSON",     "Other"),
    ("Throwing Missiles",                    50, "AGAINST THE PERSON",     "Other"),
    ("Unlawful Sexual Intercourse",          50, "AGAINST PUBLIC MORALITY","Unlawful Sexual Intercourse"),
    ("Unlawful Sexual Intercourse (11 Counts)",50,"AGAINST PUBLIC MORALITY","Unlawful Sexual Intercourse"),
    ("Unlawful Sexual Intercourse (2 Counts)",50, "AGAINST PUBLIC MORALITY","Unlawful Sexual Intercourse"),
    ("Unlawful Sexual Intercourse (2x)",     50, "AGAINST PUBLIC MORALITY","Unlawful Sexual Intercourse"),
    ("Unlawful Sexual Intercourse (6x)",     50, "AGAINST PUBLIC MORALITY","Unlawful Sexual Intercourse"),
    ("Unregistered Motor Vehicle",           65, "OTHERS",                 "Other"),
    ("Use of Deadly means of Haarm",         82, "AGAINST THE PERSON",     "Grievous Harm"),
    ("Use of Deadly Means of Harm",          82, "AGAINST THE PERSON",     "Attempted Murder"),
    ("Use of Deadly means of Harm",          82, "AGAINST THE PERSON",     "Wounding"),
    ("Use of Threatening Word",              50, "AGAINST LAWFUL AUTHORITY","Other"),
    ("Used a Computer System to Publish Data that is Obscene and Vulgar", 50, "AGAINST THE PERSON", "Other"),
    ("Used A Computer System to Transmit Computer data That Threatens the Other Person With Violence (4 counts)", 50, "AGAINST THE PERSON", "Other"),
    ("Used a Computer System to Transmit Computer Data thatis Profane (2 counts)", 50, "AGAINST THE PERSON", "Other"),
    ("Used an Unlicensed Motor Vehicle",     65, "OTHERS",                 "Other"),
    ("Used Motor Vehicle Not Covered by third Party Risk Insurance", 65, "OTHERS", "Other"),
    ("Used Motor Vehicle not covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used motor Vehicle not Covered by Third party risk insurance", 65, "OTHERS", "Other"),
    ("used motor Vehicle not covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used motor vehicle whilst being covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used Motor Vehicle Whilst not Being covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used motor vehicle Whilst not being covered by third Party Risk Insurance", 65, "OTHERS", "Other"),
    ("Used Motor Vehicle Whilst not beingcovered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used Motor Vehicle whilst not covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used Motor Whilst not being covered by third party risk insurance", 65, "OTHERS", "Other"),
    ("Used Unlicensed Motor Vehicle",        65, "OTHERS",                 "Other"),
    ("Using A Computer System to Transmit Computer Data That is Lewd", 50, "AGAINST THE PERSON", "Other"),
    ("Using Abusive Language",               50, "AGAINST THE PERSON",     "Other"),
    ("Using Abusive Languagge",              50, "AGAINST LAWFUL AUTHORITY","Escape"),
    ("Using Insulting Words",                50, "AGAINST THE PERSON",     "Other"),
    ("Using Obscene Language",               50, "AGAINST THE PERSON",     "Other"),
    ("Using Threatening Word",               50, "AGAINST THE PERSON",     "Aggravated Assault"),
    ("Wanfering Abroad",                     50, "AGAINST THE PERSON",     "Other"),
    ("Wondering Abroad",                     45, "OTHERS",                 "Other"),
    ("Wounding",                             80, "AGAINST THE PERSON",     "Wounding"),
]


def get_statutory_nature(charge_name: str) -> str:
    """Derive StatutoryNature from charge name using keyword matching."""
    c = charge_name.upper()
    if any(p in c for p in ["DRUG", "CANNABIS", "PIPE", "UTENSIL", "SMOKING CANNABIS"]):
        return "Dangerous Drugs"
    if any(p in c for p in ["FIREARM", "AMMUNITION", "GUN LICENSE", "GUN LICENCE",
                              "PROHIBITED FIREARM", "PROHIBITED AMMO"]):
        return "Against Firearms Act"
    if any(p in c for p in ["ASSAULT A POLICE", "ASSAULTING A POLICE", "ASSAULTING THE POLICE"]):
        return "Against Police Act"
    if any(p in c for p in ["DROVE MOTOR", "DROVE UNLICENSED", "DROVEMOTOR",
                              "DRIVING WHILST UNFIT", "RECKLESS DRIVING",
                              "USED MOTOR", "USED AN UNLICENSED", "USED UNLICENSED",
                              "UNREGISTERED MOTOR", "DRIVER'S LICENSE", "DRIVER'S LICENCE",
                              "THIRD PARTY RISK", "WITHOUT DUE CARE",
                              "FAIL TO PROVIDE SPECIMEN", "FAILURE TO PROVIDE",
                              "FAILED TO PROVIDE SPECIMEN", "FAIL TO GIVE WAY",
                              "FAILED TO SIGNAL", "FAIL TO SIGNAL",
                              "FAILDED TO DISPLAY", "WITHOUT A PROTECTIVE HELMET",
                              "DISQUALIFIED FROM HOLDING"]):
        return "Other Offences"
    return ""


# ──────────────────────────────────────────────────────────────────────────────
# COLUMN-WIDTH MAPS
# ──────────────────────────────────────────────────────────────────────────────

CASE_COL_WIDTHS = {
    1: 16, 2: 13, 3: 13, 4: 24, 5: 13, 6: 24, 7: 24,
    8:  5, 9:  6, 10:13, 11:28, 12:13, 13:40, 14:13,
    15:30, 16:35, 17:20, 18:13, 19:13, 20:13,
}

# ──────────────────────────────────────────────────────────────────────────────
# BUILD WORKBOOK
# ──────────────────────────────────────────────────────────────────────────────

wb = Workbook()
wb.remove(wb.active)   # drop default "Sheet"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1: Active_Case
# ══════════════════════════════════════════════════════════════════════════════
ws_ac = wb.create_sheet("Active_Case")
ws_ac.sheet_view.showGridLines = True

ACTIVE_COLS = CASE_COLS_20[:18]   # A:R  (no EntryDate / DateUpdated)
header_row(ws_ac, 1, ACTIVE_COLS)
ws_ac.row_dimensions[1].height = 30

# Blank data row so the table is valid (row 2 is intentionally empty)
make_table(ws_ac, f"A1:R2", "tblActive", "TableStyleMedium2")

col_widths(ws_ac, {k: v for k, v in CASE_COL_WIDTHS.items() if k <= 18})
ws_ac.freeze_panes = "B2"

# Data validations (applied to rows 2 onwards)
add_dv(ws_ac, "list", '"San Pedro,Caye Caulker"',  "E2:E10000")
add_dv(ws_ac, "list", '"M,F"',                     "H2:H10000")
add_dv(ws_ac, "list", '"Yes,No"',                  "L2:L10000")
outcome_str = ",".join(OUTCOME_LIST)
add_dv(ws_ac, "list", f'"{outcome_str}"',           "Q2:Q10000")

# Instructional note in a visible cell
ws_ac["T1"] = "← Enter new active cases directly into this table"
ws_ac["T1"].font = Font(color="808080", italic=True, size=9)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2: Operations
# ══════════════════════════════════════════════════════════════════════════════
ws_op = wb.create_sheet("Operations")
ws_op.sheet_view.showGridLines = True
ws_op.row_dimensions[8].height = 32

header_row(ws_op, 8, CASE_COLS_20, fill_=FILL_DARK_TEAL)

# Template data row (row 9) — XLOOKUP formulas auto-populate from DataBase
# Column A (CourtBookNo) is the only manually-entered field per row.
# Columns B-M pull from tblDataBase via XLOOKUP.
# Column I (Age)     is derived from DOB.
# Column S (EntryDate) defaults to ArraignmentDate.

def ops_xlookup(field: str) -> str:
    return (
        f'=IF([@CourtBookNo]="","",XLOOKUP([@CourtBookNo],'
        f'tblDataBase[CourtBookNo],tblDataBase[{field}],""))'
    )

ops_row_formulas = {
    2:  ops_xlookup("DateOfOccurence"),
    3:  ops_xlookup("ArraignmentDate"),
    4:  ops_xlookup("ArrestingOfficer"),
    5:  ops_xlookup("Station"),
    6:  ops_xlookup("ComplainantVictim"),
    7:  ops_xlookup("Defendant"),
    8:  ops_xlookup("Sex"),
    9:  '=IF([@CourtBookNo]="","",IFERROR(DATEDIF([@DOB],TODAY(),"Y"),""))',
    10: ops_xlookup("DOB"),
    11: ops_xlookup("Address"),
    12: ops_xlookup("CaseFileReceived"),
    13: ops_xlookup("Charge"),
    19: "=[@ArraignmentDate]",
    20: "=TODAY()",
}

for col, formula in ops_row_formulas.items():
    ws_op.cell(row=9, column=col, value=formula)

make_table(ws_op, "A8:T9", "tblOps", "TableStyleMedium9")

col_widths(ws_op, CASE_COL_WIDTHS)
ws_op.freeze_panes = "A9"

# Instructions above the header
ws_op["A1"] = "OPERATIONS REGISTER"
ws_op["A1"].font = FONT_TITLE
ws_op["A3"] = "Enter the CourtBookNo in column A — all case details auto-fill from DataBase via XLOOKUP."
ws_op["A3"].font = Font(italic=True, color="555555", size=9)
ws_op["A4"] = "Columns N–R (NextHearingDate, Remark, Further Remark, Outcome, ConcludedDate) are entered manually."
ws_op["A4"].font = Font(italic=True, color="555555", size=9)


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3: Sheet1  (configuration / parameter sheet)
# ══════════════════════════════════════════════════════════════════════════════
ws_s1 = wb.create_sheet("Sheet1")

ws_s1.merge_cells("A1:D1")
ws_s1["A1"] = "CourtCIMS System Configuration"
ws_s1["A1"].font = FONT_TITLE
ws_s1["A1"].fill = FILL_LIGHT_BLUE
ws_s1["A1"].alignment = ALIGN_CENTER

params = [
    (3, "System Version",  "CourtCIMSV4"),
    (4, "Court Name",      "San Pedro Magistrate Court"),
    (5, "Jurisdiction",    "Belize"),
    (6, "Prosecutor",      "Claude Pitts SGT 920"),
    (7, "Magistrate",      "Ms. T. Brown"),
    (8, "Supervisor",      "Supt. Egbert Castellanos"),
    (9, "Last Updated",    "=TODAY()"),
]
for row, label, value in params:
    ws_s1.cell(row, 1, label).font  = FONT_DARK_BOLD
    ws_s1.cell(row, 2, value).alignment = ALIGN_LEFT
    if row == 9:
        ws_s1.cell(row, 2).number_format = "DD-MMM-YYYY"

ws_s1.column_dimensions["A"].width = 20
ws_s1.column_dimensions["B"].width = 40


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4: DataBase  (master case repository — all cases ever entered)
# ══════════════════════════════════════════════════════════════════════════════
ws_db = wb.create_sheet("DataBase")
ws_db.sheet_view.showGridLines = True
ws_db.row_dimensions[3].height = 32

header_row(ws_db, 3, CASE_COLS_20, fill_=FILL_DARK_BLUE)

# Blank data row to satisfy table minimum
make_table(ws_db, "A3:T4", "tblDataBase", "TableStyleMedium2")

col_widths(ws_db, CASE_COL_WIDTHS)
ws_db.freeze_panes = "B4"

# Data validations
add_dv(ws_db, "list", '"San Pedro,Caye Caulker"', "E5:E50000")
add_dv(ws_db, "list", '"M,F"',                    "H5:H50000")
add_dv(ws_db, "list", '"Yes,No"',                 "L5:L50000")
add_dv(ws_db, "list", f'"{outcome_str}"',          "Q5:Q50000")

# Instructions
ws_db["A1"] = "MASTER DATABASE  — All cases (active + disposed). Do NOT delete rows; archive by year."
ws_db["A1"].font = Font(italic=True, color="555555", size=9)
ws_db.row_dimensions[1].height = 14


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5: Daily_Exports  (formatted daily sitrep export)
# ══════════════════════════════════════════════════════════════════════════════
ws_de = wb.create_sheet("Daily_Exports")

# ── Police letterhead ───────────────────────────────────────────────────────
ws_de.merge_cells("A1:R1")
ws_de["A1"] = "Belize Police Department"
ws_de["A1"].fill      = FILL_NAVY
ws_de["A1"].font      = Font(name="Calibri", color="FFFFFF", bold=True, size=14)
ws_de["A1"].alignment = ALIGN_CENTER
ws_de.row_dimensions[1].height = 28

info_de = [
    (2,  "From:",    "Claude Pitts SGT 920"),
    (3,  "To:",      "OC Prosecution Branch, I/C CRO Office, Compol HNCIB"),
    (4,  "Thru:",    "Supt. Egbert Castellanos ( OC San Pedro Sub Formation)"),
    (5,  "Subject:", '="San Pedro Magistrate Court Daily Sitrep for "&TEXT(TODAY(),"dd-mmm-yyyy")'),
    (6,  "Dated:",   "=TODAY()"),
]
for r, lbl, val in info_de:
    ws_de.cell(r, 1, lbl).font = FONT_DARK_BOLD
    cell = ws_de.cell(r, 2, val)
    if r == 6:
        cell.number_format = "DD-MMM-YYYY"

# Divider
ws_de.row_dimensions[7].height = 8
ws_de.row_dimensions[8].height = 8

# ── Data table ─────────────────────────────────────────────────────────────
header_row(ws_de, 9, CASE_COLS_15, fill_=FILL_TEAL)
ws_de.row_dimensions[9].height = 30

make_table(ws_de, "A9:O10", "tblDailyExports", "TableStyleMedium9")

col_widths(ws_de, {
    1: 16, 2: 13, 3: 13, 4: 24, 5: 13, 6: 24, 7: 24,
    8:  5, 9:  6, 10:13, 11:28, 12:40, 13:30, 14:18, 15:13,
})
ws_de.freeze_panes = "A10"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6: Monthly_Returns  (disposed-case monthly statistics)
# ══════════════════════════════════════════════════════════════════════════════
ws_mr = wb.create_sheet("Monthly_Returns")

# ── Metadata header ─────────────────────────────────────────────────────────
meta = [
    (2, "Reporting Period", ""),         # user fills: "March 1st 2026 to March 31st 2026"
    (3, "Court No.",        "San Pedro Magistrate Court"),
    (4, "Magistrate:",      "Ms. T. Brown"),
    (5, "Prosecutor",       "Claude Pitts SGT 920"),
]
for r, lbl, val in meta:
    ws_mr.cell(r, 1, lbl).font = FONT_DARK_BOLD
    ws_mr.cell(r, 2, val)

# ── Summary count table ─────────────────────────────────────────────────────
# tblReturnCount  E2:N5
# Headers: Station | P.I Conducted | Struck Out | Bond Over | Withdrawn |
#          Dismissed | Conditional Discharged | Convicted | Acquitted | Total

summary_outcome_hdrs = [
    "P.I Conducted","Struck Out","Bond Over","Withdrawn",
    "Dismissed","Conditional Discharged","Convicted","Acquitted","Total",
]

styled_cell(ws_mr, 2, 5, "Station",
            fill_=FILL_DARK_BLUE, font=FONT_WHITE_BOLD, align=ALIGN_CENTER)

for ci, lbl in enumerate(summary_outcome_hdrs, start=6):
    styled_cell(ws_mr, 2, ci, lbl,
                fill_=FILL_DARK_BLUE, font=FONT_WHITE_BOLD, align=ALIGN_CENTER)
    ws_mr.column_dimensions[get_column_letter(ci)].width = 14

ws_mr.column_dimensions["E"].width = 16
ws_mr.column_dimensions["N"].width = 8

outcome_keys = summary_outcome_hdrs[:-1]   # without "Total"

stations_rows = [("San Pedro", 3), ("Caye Caulker", 4), ("TOTAL", 5)]
for stn, row in stations_rows:
    styled_cell(ws_mr, row, 5, stn, font=FONT_DARK_BOLD)
    if stn != "TOTAL":
        for ci, ok in enumerate(outcome_keys, start=6):
            ws_mr.cell(row, ci,
                f'=COUNTIFS(tblReturns[Station],"{stn}",tblReturns[Outcome],"{ok}")')
        ws_mr.cell(row, 14, f"=SUM(F{row}:M{row})")
    else:
        for ci in range(6, 15):
            col_letter = get_column_letter(ci)
            ws_mr.cell(row, ci, f"=SUM({col_letter}3:{col_letter}4)")

make_table(ws_mr, "E2:N5", "tblReturnCount", "TableStyleMedium2")

# ── Data table (disposed cases for the period) ──────────────────────────────
header_row(ws_mr, 9, CASE_COLS_15, fill_=FILL_DARK_BLUE)
ws_mr.row_dimensions[9].height = 30

make_table(ws_mr, "A9:O10", "tblReturns", "TableStyleMedium9")

col_widths(ws_mr, {
    1: 16, 2: 13, 3: 13, 4: 24, 5: 13, 6: 24, 7: 24,
    8:  5, 9:  6, 10:13, 11:28, 12:40, 13:30, 14:18, 15:13,
})
ws_mr.freeze_panes = "A10"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7: ChargeMap  (all known charge categories)
# ══════════════════════════════════════════════════════════════════════════════
ws_cm = wb.create_sheet("ChargeMap")

# Title
ws_cm.merge_cells("A1:G1")
ws_cm["A1"] = "CHARGE CLASSIFICATION MAP"
ws_cm["A1"].font      = FONT_TITLE
ws_cm["A1"].fill      = FILL_LIGHT_BLUE
ws_cm["A1"].alignment = ALIGN_CENTER
ws_cm.row_dimensions[1].height = 25

ws_cm["A3"] = ("tblChargeMap is used by Operations and other sheets to classify charges "
               "into Annexure categories and statutory types.")
ws_cm["A3"].font = Font(italic=True, color="555555", size=9)

CM_HEADERS = ["ChargeName","PriorityRank","AnnexureNature",
              "AnnexureSubCategory","IsStatutory","StatutoryNature","StatOverride"]
header_row(ws_cm, 6, CM_HEADERS, fill_=FILL_GREEN)
ws_cm.row_dimensions[6].height = 30

for ri, (charge_name, priority, ann_nature, ann_sub) in enumerate(CHARGE_MAP_DATA, start=7):
    ws_cm.cell(ri, 1, charge_name)
    ws_cm.cell(ri, 2, priority)
    ws_cm.cell(ri, 3, ann_nature)
    ws_cm.cell(ri, 4, ann_sub)
    ws_cm.cell(ri, 5, '=[@StatutoryNature]<>""')
    ws_cm.cell(ri, 6, get_statutory_nature(charge_name))
    ws_cm.cell(ri, 7, "")   # StatOverride — user may manually override

last_cm = 6 + len(CHARGE_MAP_DATA)
make_table(ws_cm, f"A6:G{last_cm}", "tblChargeMap", "TableStyleLight9")

col_widths(ws_cm, {1: 70, 2: 13, 3: 26, 4: 45, 5: 13, 6: 30, 7: 28})


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8: Helper  (all reference / lookup data)
# ══════════════════════════════════════════════════════════════════════════════
ws_h = wb.create_sheet("Helper")

ws_h.merge_cells("A1:R1")
ws_h["A1"] = "HELPER / REFERENCE DATA — Do not rename or delete tables or columns."
ws_h["A1"].font      = Font(color="FFFFFF", bold=True, size=10)
ws_h["A1"].fill      = FILL_DARK_BLUE
ws_h["A1"].alignment = ALIGN_CENTER

# ── Section A: Outcome → CaseStatusBucket  (cols A:B) ───────────────────────
header_row(ws_h, 3, ["OutcomeText", "CaseStatusBucket"], start_col=1)
for ri, (outcome, status) in enumerate(OUTCOME_STATUS_MAP, start=4):
    ws_h.cell(ri, 1, outcome)
    ws_h.cell(ri, 2, status)
last_os = 3 + len(OUTCOME_STATUS_MAP)
make_table(ws_h, f"A3:B{last_os}", "tblOutcomes", "TableStyleLight9")

# ── Section B: StatutoryNature patterns  (cols E:F) ─────────────────────────
header_row(ws_h, 3, ["Pattern", "StatutoryNature"], start_col=5)
for ri, (pattern, stat_nat) in enumerate(STAT_PATTERNS, start=4):
    ws_h.cell(ri, 5, pattern)
    ws_h.cell(ri, 6, stat_nat)
last_sp = 3 + len(STAT_PATTERNS)
make_table(ws_h, f"E3:F{last_sp}", "tblStatPatterns", "TableStyleLight9")

# ── Section C: StatutoryNature labels (col H) ───────────────────────────────
stat_nature_labels = [
    "Dangerous Drugs",
    "Against Firearms Act",
    "Against Liqour Act",
    "Against Police Act",
    "Gambling",
    "Summary Jurisdiction Offences",
    "Other Offences",
]
header_row(ws_h, 3, ["StatutoryNature"], start_col=8)
for ri, lbl in enumerate(stat_nature_labels, start=4):
    ws_h.cell(ri, 8, lbl)
make_table(ws_h, f"H3:H{3+len(stat_nature_labels)}", "tblStatNature", "TableStyleLight9")

# ── Section D: ArrestingOfficer list  (col J) ───────────────────────────────
header_row(ws_h, 3, ["ArrestingOfficer"], start_col=10)
for ri, officer in enumerate(ARRESTING_OFFICERS, start=4):
    ws_h.cell(ri, 10, officer)
last_ao = 3 + len(ARRESTING_OFFICERS)
make_table(ws_h, f"J3:J{last_ao}", "tblOfficers", "TableStyleLight9")

# ── Section E: Station list  (col L) ────────────────────────────────────────
header_row(ws_h, 3, ["Station"], start_col=12)
stations = ["San Pedro", "Caye Caulker"]
for ri, stn in enumerate(stations, start=4):
    ws_h.cell(ri, 12, stn)
make_table(ws_h, f"L3:L{3+len(stations)}", "tblStations", "TableStyleLight9")

# Section labels
for col, label in [(1,"Outcome Mapping"),(5,"Stat Patterns"),(8,"Stat Labels"),(10,"Officers"),(12,"Stations")]:
    ws_h.cell(2, col, label).font = Font(bold=True, color="1F4E79", size=9)

col_widths(ws_h, {1:22, 2:22, 5:30, 6:30, 8:30, 10:28, 12:16})


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 9: View  (filtered read-only view of DataBase)
# ══════════════════════════════════════════════════════════════════════════════
ws_v = wb.create_sheet("View")

ws_v.merge_cells("A1:T1")
ws_v["A1"] = "CASE VIEW — Filtered live view of DataBase (read-only). Refresh by pressing F9."
ws_v["A1"].font      = Font(color="FFFFFF", bold=True, size=10)
ws_v["A1"].fill      = FILL_DARK_TEAL
ws_v["A1"].alignment = ALIGN_CENTER

ws_v["A3"] = "Filter by Outcome:"
ws_v["A3"].font = FONT_DARK_BOLD
ws_v["B3"] = ""   # user can enter filter value here

# View headers (row 6)
VIEW_HEADERS = ["RowID"] + CASE_COLS_20[:19]
header_row(ws_v, 6, VIEW_HEADERS, fill_=FILL_DARK_TEAL)
ws_v.row_dimensions[6].height = 30

# FILTER formula — shows all cases where ConcludedDate is blank (active cases)
# Falls back to all cases if no active cases, or shows the outcome filter from B3.
# Uses Excel 365 FILTER + LET + SEQUENCE functions.
filter_formula = (
    '=IFERROR('
    'LET('
    'db,tblDataBase,'
    'filt,IF(B3="",db[ConcludedDate]="",db[Outcome]=B3),'
    'result,FILTER(db,filt,"No matching cases"),'
    'HSTACK(SEQUENCE(ROWS(result)),result)'
    '),'
    '"DataBase is empty — add cases to the DataBase sheet first."'
    ')'
)
ws_v.cell(7, 1, filter_formula)

col_widths(ws_v, {1: 8, **{k+1: v for k, v in CASE_COL_WIDTHS.items() if k <= 19}})
ws_v.freeze_panes = "A7"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 10: Statistics
# ══════════════════════════════════════════════════════════════════════════════
ws_st = wb.create_sheet("Statistics")

ws_st.merge_cells("A1:N1")
ws_st["A1"] = "STATISTICS DASHBOARD"
ws_st["A1"].font      = Font(name="Calibri", color="FFFFFF", bold=True, size=16)
ws_st["A1"].fill      = FILL_DARK_BLUE
ws_st["A1"].alignment = ALIGN_CENTER
ws_st.row_dimensions[1].height = 36

ws_st.merge_cells("A3:N3")
ws_st["A3"] = "All figures below pull live from tblDataBase / tblReturns — no manual entry required."
ws_st["A3"].font      = Font(italic=True, color="555555", size=10)
ws_st["A3"].alignment = ALIGN_CENTER

# Summary KPIs
kpi_headers = ["Metric","Value"]
header_row(ws_st, 5, kpi_headers, fill_=FILL_MED_BLUE)
ws_st.row_dimensions[5].height = 24

kpis = [
    ("Total Cases in DataBase",      "=ROWS(tblDataBase)-1"),
    ("Active Cases (not concluded)", '=COUNTIF(tblDataBase[ConcludedDate],"")'),
    ("Convicted (all time)",         '=COUNTIF(tblDataBase[Outcome],"Convicted")'),
    ("Dismissed (all time)",         '=COUNTIF(tblDataBase[Outcome],"Dismissed")'),
    ("Struck Out (all time)",        '=COUNTIF(tblDataBase[Outcome],"Struck Out")'),
    ("Withdrawn (all time)",         '=COUNTIF(tblDataBase[Outcome],"Withdrawn")'),
    ("San Pedro cases",              '=COUNTIF(tblDataBase[Station],"San Pedro")'),
    ("Caye Caulker cases",           '=COUNTIF(tblDataBase[Station],"Caye Caulker")'),
]

for ri, (label, formula) in enumerate(kpis, start=6):
    ws_st.cell(ri, 1, label).font  = FONT_DARK_BOLD
    cell = ws_st.cell(ri, 2, formula)
    cell.font = Font(name="Calibri", bold=True, size=12, color="1F4E79")
    cell.alignment = ALIGN_RIGHT

# Outcome breakdown table
header_row(ws_st, 16, ["Outcome","Count","% of Total"], fill_=FILL_DARK_BLUE, start_col=1)
ws_st.row_dimensions[16].height = 24

for ri, outcome in enumerate(OUTCOME_LIST, start=17):
    ws_st.cell(ri, 1, outcome)
    ws_st.cell(ri, 2, f'=COUNTIF(tblDataBase[Outcome],"{outcome}")')
    ws_st.cell(ri, 3, f'=IFERROR(B{ri}/SUM(B17:B28),0)')
    ws_st.cell(ri, 3).number_format = "0.0%"

col_widths(ws_st, {1: 35, 2: 14, 3: 14})


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 11: 2026Returns  (full-year disposed cases for 2026 reporting)
# ══════════════════════════════════════════════════════════════════════════════
ws_26 = wb.create_sheet("2026Returns")

ws_26.merge_cells("A1:T1")
ws_26["A1"] = "2026 ANNUAL RETURNS — Disposed/Concluded Cases"
ws_26["A1"].font      = Font(name="Calibri", color="FFFFFF", bold=True, size=12)
ws_26["A1"].fill      = FILL_ORANGE
ws_26["A1"].alignment = ALIGN_CENTER
ws_26.row_dimensions[1].height = 28

ws_26["A3"] = "Paste or import concluded cases here for 2026 annual statistical reporting."
ws_26["A3"].font = Font(italic=True, color="555555", size=9)

header_row(ws_26, 5, CASE_COLS_20, fill_=FILL_ORANGE)
ws_26.row_dimensions[5].height = 30

make_table(ws_26, "A5:T6", "tbl2026Returns", "TableStyleMedium2")

col_widths(ws_26, CASE_COL_WIDTHS)
ws_26.freeze_panes = "B6"

add_dv(ws_26, "list", '"San Pedro,Caye Caulker"', "E7:E50000")
add_dv(ws_26, "list", '"M,F"',                    "H7:H50000")
add_dv(ws_26, "list", f'"{outcome_str}"',          "Q7:Q50000")


# ──────────────────────────────────────────────────────────────────────────────
# WORKBOOK PROPERTIES
# ──────────────────────────────────────────────────────────────────────────────
wb.properties.title   = "CourtCIMS Blank Template"
wb.properties.subject = "Court Case Information Management System"
wb.properties.creator = "CourtCIMSV4 Template Generator"

# ──────────────────────────────────────────────────────────────────────────────
# SAVE
# ──────────────────────────────────────────────────────────────────────────────
import os
output_path = os.path.join(os.path.dirname(__file__), "CourtCIMS_Blank_Template.xlsx")
wb.save(output_path)
print(f"✓  Saved: {output_path}")
print(f"   Sheets ({len(wb.sheetnames)}): {', '.join(wb.sheetnames)}")
