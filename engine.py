"""
Marksheet & Transcript engine (Python port of Code.gs).

Pure logic + openpyxl template filling. No web/Google dependency here, so the
data source (live Google Sheet, uploaded xlsx, or a DB) can be swapped without
touching this file. Rendering to PDF is done with LibreOffice headless by the
caller (see render.py).

Faithful port of the Apps Script: same PROGRAMS config, grade scales, alias
handling, credential rules and cell mappings.
"""

import re
import datetime

# ============================== SETTINGS ==============================

FIXED_HEADERS = {
    "id": "Student ID",
    "first": "First Name",
    "last": "Last Name",
    "start": "Start Date",
    "finish": "Finish Date",
}
HEADER_ROW = 1
SUBHEADER_ROW = 2
FIRST_DATA_ROW = 3

GRADES_STANDARD = [(90, "A+"), (85, "A"), (75, "B+"), (70, "B"), (65, "C+"), (60, "C")]
GRADES_HIGH = [(95, "A+"), (90, "A"), (85, "B+"), (80, "B"), (75, "C+"), (70, "C")]

GRADE_ROUNDING = "round"
HONOURS_MARK = 89.5
CREDENTIAL_TEXT = {"honours": "Diploma with Honours", "pass": "Diploma", "fail": "Incomplete"}

# portrait + fit-to-page + letter + 0.25" margins, to mirror the Apps Script export
PDF_SETTINGS = {"portrait": True, "margin": 0.25}

TRANSCRIPT_STANDARD = {
    "range": "A1:K37",
    "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
    "firstRow": 6, "lastRow": 29, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11,
    "averageCell": "J31", "averageGradeCell": "K31",
    "credentialCell": "I4", "passMark": 60,
}

PROGRAMS = [
    {
        "key": "IBA",
        "label": "International Business Administration",
        "dataSheet": "IBA",
        "marksheetSheet": "IBA Marksheet",
        "transcriptSheet": "IBA Transcript",
        "marksheet": {
            "range": "A1:Z9",
            "idCell": "C2", "nameCell": "A3", "startCell": "P8", "endCell": "W8", "finalMarkCell": "E8",
            "moduleNameRow": 4, "firstCol": 3,
            "rows": [
                {"type": "participation", "row": 5},
                {"type": "final", "row": 6},
            ],
            "averageRow": 7,
        },
        "transcript": TRANSCRIPT_STANDARD,
        "grades": GRADES_STANDARD,
    },
    {
        "key": "HT",
        "label": "Hospitality and Tourism",
        "dataSheet": "HT",
        "marksheetSheet": "HT Marksheet",
        "transcriptSheet": "HT Transcript",
        "marksheet": {
            "range": "A1:Z9",
            "idCell": "C2", "nameCell": "A3", "startCell": "P8", "endCell": "W8", "finalMarkCell": "E8",
            "moduleNameRow": 4, "firstCol": 3,
            "rows": [
                {"type": "participation", "row": 5},
                {"type": "final", "row": 6},
            ],
            "averageRow": 7,
        },
        "transcript": TRANSCRIPT_STANDARD,
        "grades": GRADES_STANDARD,
    },
    {
        "key": "CSW",
        "label": "Community Service Worker",
        "dataSheet": "CSW",
        "marksheetSheet": "CSW Marksheet",
        "transcriptSheet": "CSW Transcript",
        "marksheet": {
            "range": "A1:U11",
            "idCell": "C2", "nameCell": "A3", "startCell": "K10", "endCell": "R10", "finalMarkCell": "D10",
            "moduleNameRow": 5, "firstCol": 3,
            "rows": [
                {"type": "assignment", "row": 6},
                {"type": "participation", "row": 7},
                {"type": "final", "row": 8},
            ],
            "averageRow": 9,
        },
        "transcript": {
            "range": "A1:K32",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 24, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11,
            "averageCell": "J26", "averageGradeCell": "K26",
            "credentialCell": "I4", "passMark": 70,
        },
        "grades": GRADES_HIGH,
    },
    {
        "key": "MOA",
        "label": "Medical Office Administration",
        "dataSheet": "MOA",
        "marksheetSheet": "MOA Marksheet",
        "transcriptSheet": "MOA Transcript",
        "aliases": {"Buisness Communication": "Business Communication"},
        "marksheet": {
            "range": "A1:L13",
            "idCell": "C3", "nameCell": "A4", "startCell": "G12", "endCell": "K12", "finalMarkCell": "C12",
            "moduleNameRow": 6, "firstCol": 3,
            "rows": [
                {"type": "assignment", "row": 7},
                {"type": "participation", "row": 8},
                {"type": "midterm", "row": 9},
                {"type": "final", "row": 10},
            ],
            "averageRow": 11,
        },
        "transcript": {
            "range": "A1:K23",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 15, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11,
            "averageCell": "J17", "averageGradeCell": "K17",
            "credentialCell": "I4", "passMark": 60,
        },
        "grades": GRADES_STANDARD,
    },
    {
        "key": "AC_NEW",
        "label": "Accounting, Payroll, Business & Tax (New)",
        "dataSheet": "AC2 (New)",
        "marksheetSheet": "AC2 (New) Marksheet",
        "transcriptSheet": "AC2 (New) Transcript",
        "aliases": {
            "Microsoft Word Level 2": "Microsoft Word 2",
            "Microsoft Excel Level 1 & Excel 2": "Microsoft Excel 1 and Excel 2",
            "Payroll Fundamentals Level 1 & 2": "Payroll Fundamentals 1 & 2",
        },
        "marksheet": {
            "range": "A1:N13",
            "idCell": "C2", "nameCell": "A3", "startCell": "G12", "endCell": "L12", "finalMarkCell": "C12",
            "moduleNameRow": 5, "firstCol": 3,
            "rows": [
                {"type": "assignment", "row": 6, "nth": 0},
                {"type": "assignment", "row": 7, "nth": 1, "optional": True},
                {"type": "quiz", "row": 8},
                {"type": "participation", "row": 9},
                {"type": "final", "row": 10},
            ],
            "averageRow": 11,
        },
        "transcript": {
            "range": "A1:K26",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 17, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11,
            "averageCell": "J18", "averageGradeCell": "K18",
            "credentialCell": "I4", "passMark": 70,
        },
        "grades": GRADES_STANDARD,
        "averageGrades": GRADES_HIGH,
    },
    {
        "key": "AC_OLD",
        "label": "Accounting, Payroll, Business & Tax (Old)",
        "dataSheet": "AC2 (Old)",
        "marksheetSheet": "AC2 (Old) Marksheet",
        "transcriptSheet": "AC2 (Old) Transcript",
        "aliases": {
            "Microsoft Word Level 2": "Microsoft Word 2",
            "Microsoft Excel Level 1 & Excel 2": "Microsoft Excel 1 and Excel 2",
            "Payroll Fundamentals Level 1": "Payroll Fundamentals 1",
            "Payroll Fundamental Level 2": "Payroll Fundamental 2",
        },
        "labelTypes": {"assignment 2": "quiz"},
        "marksheet": {
            "range": "A1:O12",
            "idCell": "C2", "nameCell": "A3", "startCell": "H11", "endCell": "M11", "finalMarkCell": "C11",
            "moduleNameRow": 5, "firstCol": 3,
            "rows": [
                {"type": "assignment", "row": 6},
                {"type": "quiz", "row": 7},
                {"type": "participation", "row": 8},
                {"type": "final", "row": 9},
            ],
            "averageRow": 10,
        },
        "transcript": {
            "range": "A1:K27",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 18, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11,
            "averageCell": "J19", "averageGradeCell": "K19",
            "credentialCell": "I4", "passMark": 70,
        },
        "grades": GRADES_STANDARD,
        "averageGrades": GRADES_HIGH,
    },
    {
        "key": "PSW",
        "label": "Personal Support Worker (NACC)",
        "dataSheet": "PSW",
        "marksheetSheet": "PSW Marksheet",
        "transcriptSheet": "PSW Transcript",
        "transcriptPassSheet": "PSW Transcript Pass",
        "naccModule": "NACC Exam",
        "aliases": {
            "PSW Foundations": "PSW Foundation",
            "Care Planning/Restorative/Documentation/Working in the Community":
                "Care Planning/Restorative Care/Electronic Documentation/ Working in the Community",
        },
        "labelTypes": {"pratical": "practical"},
        "marksheet": {
            "range": "A1:Q17",
            "idCell": "C2", "nameCell": "A3", "startCell": "E16", "endCell": "N16",
            "moduleNameRow": 5, "firstCol": 3,
            "rows": [
                {"type": "assignment", "row": 6, "nth": 0, "minCount": 2},
                {"type": "assignment", "row": 7, "nth": 1, "minCount": 2},
                {"type": "assignment", "row": 8},
                {"type": "participation", "row": 9},
                {"type": "practical", "row": 10},
                {"type": "theory", "row": 11, "nth": 0},
                {"type": "theory", "row": 12, "nth": 1},
                {"type": "theory", "row": 13, "nth": 2},
                {"type": "theory", "row": 14},
                {"type": "status", "row": 6},
            ],
            "averageRow": 15,
        },
        "transcript": {
            "range": "A1:K28",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 20, "marksLastRow": 17, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11, "statusCol": 8,
            "averageCell": "J22", "averageGradeCell": "K22",
            "credentialCell": "I4", "passMark": 70,
        },
        "transcriptPass": {
            "range": "A1:K29",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 21, "marksLastRow": 17, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11, "statusCol": 8,
            "averageCell": "J23", "averageGradeCell": "K23",
            "credentialCell": "I4", "passMark": 70,
        },
        "grades": GRADES_HIGH,
    },
    {
        "key": "ECA",
        "label": "Early Childhood Assistant (NACC)",
        "dataSheet": "ECA",
        "marksheetSheet": "ECA Marksheet",
        "transcriptSheet": "ECA Transcript",
        "transcriptPassSheet": "ECA Transcript Pass",
        "naccModule": "NACC Exam",
        "marksheet": {
            "range": "A1:Q15",
            "idCell": "C2", "nameCell": "A3", "startCell": "F14", "endCell": "N14",
            "moduleNameRow": 5, "firstCol": 3,
            "rows": [
                {"type": "assignment", "row": 6, "nth": 0},
                {"type": "assignment", "row": 7, "nth": 1},
                {"type": "presentation", "row": 8},
                {"type": ["assignment", "presentation"], "row": 9},
                {"type": "participation", "row": 10},
                {"type": "midterm", "row": 11},
                {"type": "final", "row": 12},
                {"type": "status", "row": 6},
            ],
            "averageRow": 13,
        },
        "transcript": {
            "range": "A1:K28",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 20, "marksLastRow": 19, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11, "statusCol": 8,
            "averageCell": "J22", "averageGradeCell": "K22",
            "credentialCell": "I4", "passMark": 70,
        },
        "transcriptPass": {
            "range": "A1:K29",
            "nameCell": "C2", "idCell": "C3", "programCell": "C4", "startCell": "I2", "endCell": "I3",
            "firstRow": 6, "lastRow": 23, "marksLastRow": 19, "moduleCol": 2, "earnedCol": 10, "gradeCol": 11, "statusCol": 8,
            "statusCols": {"nacc exam": 7},
            "averageCell": "J22", "averageGradeCell": "K22",
            "credentialCell": "I4", "passMark": 70,
        },
        "grades": GRADES_HIGH,
    },
]

PROGRAMS_BY_KEY = {p["key"]: p for p in PROGRAMS}


# ============================== HELPERS ==============================

def norm_(v):
    if v is None:
        v = ""
    s = str(v).replace("‘", "'").replace("’", "'")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def module_key_(v):
    if v is None:
        v = ""
    s = re.sub(r"\(\s*M\d+\s*\)\s*$", "", str(v), flags=re.I)
    return norm_(s)


def id_str_(v):
    """Stringify an ID the way Sheets/JS does: 10663.0 -> '10663', not '10663.0'."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str("" if v is None else v).strip()


def to_num_(v):
    if v == "" or v is None:
        return None
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    if n != n:  # NaN
        return None
    return n


def pct_(points):
    """Marks out of 100 -> fraction for the % formatted template cells; blank stays blank."""
    return "" if points is None else points / 100.0


def component_type_(label):
    s = str(label)
    if s.startswith("score") or s.startswith("status"):
        return "status"
    for t in ("assignment", "quiz", "participation", "midterm", "final",
              "theory", "practical", "presentation"):
        if s.startswith(t):
            return t
    return None


def grade_for_(points, grades, exact=False):
    if GRADE_ROUNDING == "round" and not exact:
        p = __import__("math").floor(points + 0.5 + 1e-9)
    else:
        p = points + 1e-9
    for mn, g in grades:
        if p >= mn:
            return g
    return "F"


def alias_map_(program):
    out = {}
    for k, v in (program.get("aliases") or {}).items():
        out[module_key_(k)] = module_key_(v)
    return out


def credential_formula_(marks_range, pass_mark):
    avg = "ROUND(AVERAGE(%s),6)" % marks_range
    return ("=IF(COUNT(%s)=0,\"\","
            "IF(%s<%s,\"%s\","
            "IF(%s>=%s,\"%s\",\"%s\")))") % (
        marks_range,
        avg, pass_mark / 100.0, CREDENTIAL_TEXT["fail"],
        avg, HONOURS_MARK / 100.0, CREDENTIAL_TEXT["honours"], CREDENTIAL_TEXT["pass"],
    )


# ============================== DATA READING ==============================

def read_student(values, sid, program):
    """
    values: 2D list of the data sheet (row-major, as returned by the data source).
    Returns a student dict or None.
    """
    header = values[HEADER_ROW - 1] if len(values) >= HEADER_ROW else []
    sub = values[SUBHEADER_ROW - 1] if len(values) >= SUBHEADER_ROW else []
    F = FIXED_HEADERS
    label_types = (program.get("labelTypes") or {})

    def col_of(title):
        for i, h in enumerate(header):
            if norm_(h) == norm_(title):
                return i
        raise ValueError('Column "%s" was not found in row %d of a data sheet.' % (title, HEADER_ROW))

    c_id, c_first, c_last = col_of(F["id"]), col_of(F["first"]), col_of(F["last"])
    c_start, c_finish = col_of(F["start"]), col_of(F["finish"])
    skip = {c_id, c_first, c_last, c_start, c_finish}
    for i, h in enumerate(header):
        if norm_(h) == "status":
            skip.add(i)

    target = id_str_(sid)
    row = None
    for r in range(FIRST_DATA_ROW - 1, len(values)):
        cell = values[r][c_id] if c_id < len(values[r]) else None
        if id_str_(cell) == target:
            row = values[r]
            break
    if row is None:
        return None

    def cell(c):
        return row[c] if c < len(row) else None

    titled = [c for c in range(len(header))
              if c not in skip and header[c] is not None and str(header[c]).strip()]

    # The Sheets API trims empty cells at the end of each row, and the last
    # module's merged title leaves row 1 short, so the last module runs to the
    # widest of the three rows.
    width = max(len(header), len(sub), len(row))
    modules = {}
    for i, start in enumerate(titled):
        end = titled[i + 1] if i + 1 < len(titled) else width
        name = str(header[start]).strip()
        parts = {k: [] for k in ("assignment", "quiz", "participation", "midterm",
                                 "final", "theory", "practical", "presentation")}
        total = 0.0
        any_value = False
        status = None
        for c in range(start, end):
            label = norm_(sub[c]) if c < len(sub) else ""
            if not label:
                continue
            t = component_type_(label_types.get(label, label))
            if not t:
                raise ValueError('Unknown mark column "%s" under "%s" in a data sheet.'
                                 % (sub[c], name))
            if t == "status":
                raw = cell(c)
                status = str("" if raw is None else raw).strip()
                continue
            v = to_num_(cell(c))
            parts[t].append(v)
            if v is not None:
                total += v
                any_value = True
        modules[norm_(name)] = {
            "name": name, "parts": parts, "status": status,
            "average": (total if any_value else None),
        }

    first = str(cell(c_first) or "").strip()
    last = str(cell(c_last) or "").strip()
    return {
        "id": id_str_(cell(c_id)),
        "fullName": (first + " " + last).strip(),
        "start": cell(c_start),
        "finish": cell(c_finish),
        "modules": modules,
    }


def find_module(student, template_name, amap, strict=True):
    key = module_key_(template_name)
    if key in amap:
        key = amap[key]
    m = student["modules"].get(key)
    if not m and strict:
        raise ValueError(
            'Template module "%s" was not found in the data sheet. Check its spelling, '
            'or add it to "aliases" for this program.' % str(template_name).strip())
    return m


def choose_transcript(program, student):
    if program.get("transcriptPassSheet") and program.get("naccModule"):
        nacc = student["modules"].get(module_key_(program["naccModule"]))
        if nacc and str(nacc.get("status") or "").strip().upper() == "PASS":
            return program["transcriptPassSheet"], program["transcriptPass"]
    return program["transcriptSheet"], program["transcript"]


# ============================== PURE BUILDERS ==============================

def build_marksheet_rows(student, template_names, program):
    M = program["marksheet"]
    amap = alias_map_(program)
    mods = [find_module(student, n, amap) for n in template_names]

    def is_status(m):
        return m is not None and m.get("status") is not None

    rows = []
    for spec in M["rows"]:
        values = []
        for m in mods:
            if spec["type"] == "status":
                values.append(m["status"] if is_status(m) else None)
                continue
            if is_status(m):
                values.append(None)
                continue
            types = spec["type"] if isinstance(spec["type"], list) else [spec["type"]]
            lst = []
            for t in types:
                lst.extend(m["parts"].get(t, []))
            miss = None if spec.get("optional") else "N/A"
            if len(lst) < spec.get("minCount", 0):
                values.append(miss)
                continue
            if "nth" in spec:
                if len(lst) <= spec["nth"]:
                    values.append(miss)
                else:
                    values.append(pct_(lst[spec["nth"]]))
                continue
            if not lst:
                values.append(miss)
                continue
            nums = [v for v in lst if v is not None]
            values.append(pct_(sum(nums)) if nums else "")
        rows.append({"row": spec["row"], "values": values})

    rows.append({
        "row": M["averageRow"],
        "values": [None if is_status(m) else pct_(m["average"]) for m in mods],
    })

    nums = [m["average"] for m in mods if not is_status(m) and m["average"] is not None]
    final_mark = (sum(nums) / len(nums) / 100.0) if nums else ""
    return {"rows": rows, "finalMark": final_mark}


def build_transcript_block(student, template_names, program):
    amap = alias_map_(program)
    avg_scale = program.get("averageGrades") or program["grades"]
    rows = []
    scores = []
    for n in template_names:
        if not str("" if n is None else n).strip():
            rows.append({"kind": "skip"})
            continue
        m = find_module(student, n, amap)
        if m.get("status") is not None:
            mk = module_key_(n)
            key = amap[mk] if mk in amap else mk
            rows.append({"kind": "status", "key": key, "text": m["status"]})
        elif m["average"] is None:
            rows.append({"kind": "empty", "earned": "", "grade": ""})
        else:
            rows.append({"kind": "mark", "earned": m["average"] / 100.0,
                         "grade": grade_for_(m["average"], program["grades"])})
            scores.append(m["average"])
    avg_points = (sum(scores) / len(scores)) if scores else None
    return {
        "rows": rows,
        "average": "" if avg_points is None else avg_points / 100.0,
        "averageGrade": "" if avg_points is None else grade_for_(avg_points, avg_scale, True),
    }


# ============================== LOOKUP ==============================

def lookup(source, raw_id, program_key=None):
    """
    Mirrors the search part of getStudentDocuments in Code.gs. source has
    .values(sheet_name) -> 2D list. Returns {"ambiguous": [...]} when the ID is
    in several programs, else {"program": program, "student": student}.
    """
    sid = str("" if raw_id is None else raw_id).strip()
    if not sid:
        raise ValueError("Please enter a Student ID.")

    programs = [p for p in PROGRAMS if not program_key or p["key"] == program_key]
    if hasattr(source, "rows_for"):        # live Sheet: fetch just the header rows + the student's row
        tables = source.rows_for(sid, [p["dataSheet"] for p in programs])
    else:
        tables = {p["dataSheet"]: source.values(p["dataSheet"]) for p in programs}

    matches = []
    for p in programs:
        student = read_student(tables[p["dataSheet"]], sid, p)
        if student:
            matches.append((p, student))

    if not matches:
        raise ValueError('No student found with ID "%s".' % sid)
    if len(matches) > 1:
        return {"ambiguous": [{"key": p["key"], "label": p["label"],
                               "name": s["fullName"]} for p, s in matches]}
    return {"program": matches[0][0], "student": matches[0][1]}
