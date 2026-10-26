/**
 * Marksheet & Transcript Viewer  (IBA / HT / CSW / MOA / AC New / AC Old / PSW / ECA)
 * ----------------------------------------------------------------------------------
 * Web app bound to the spreadsheet. Enter a Student ID -> the script finds the
 * student in the program data sheets, writes the marks into that program's
 * Marksheet and Transcript templates, exports both to PDF and returns them to
 * the browser (preview + one combined PDF download).
 *
 * To add another program later: add a data sheet + two template sheets, then
 * add one entry to PROGRAMS below (copy the closest existing one).
 */

/* ============================== SETTINGS ============================== */

// Header text of the fixed columns in every program's data sheet
const FIXED_HEADERS = {
  id: 'Student ID',
  first: 'First Name',
  last: 'Last Name',
  start: 'Start Date',
  finish: 'Finish Date'
};
const HEADER_ROW = 1;       // module names (merged over their columns)
const SUBHEADER_ROW = 2;    // Assignment / Quiz / Participation / Midterm / Final
const FIRST_DATA_ROW = 3;

// Grade scales (minimum % -> grade). Anything below the last entry is F.
const GRADES_STANDARD = [[90, 'A+'], [85, 'A'], [75, 'B+'], [70, 'B'], [65, 'C+'], [60, 'C']];
const GRADES_HIGH     = [[95, 'A+'], [90, 'A'], [85, 'B+'], [80, 'B'], [75, 'C+'], [70, 'C']];

// How MODULE marks become grades: 'round' grades the whole-percent mark shown on the documents
// (84.5 -> 85% -> A, 79.5 -> 80% -> B); 'exact' grades the unrounded mark (84.5 -> B+).
// The overall AVERAGE's grade always uses the exact average (e.g. 79.89% on the CSW scale is C+).
const GRADE_ROUNDING = 'round';

// Credential rules: average >= HONOURS_MARK (unrounded) -> Honours; >= the program's
// passMark -> Diploma; otherwise Incomplete.
const HONOURS_MARK = 89.5;
const CREDENTIAL_TEXT = { honours: 'Diploma with Honours', pass: 'Diploma', fail: 'Incomplete' };

// PDF page setup (matches the sample PDFs: US Letter, portrait)
const PDF_SETTINGS = { size: 'letter', portrait: true, margin: 0.25 };

/*
 * Marksheet "rows": each entry fills one template row with one kind of mark.
 *   type      assignment | quiz | participation | midterm | final | theory | practical |
 *             presentation  (or an array of types, which are added together)
 *             'status' writes the text result (PASS, Passed GPA 1&2) for pass/fail modules
 *   row       template row number
 *   nth       (optional) use only the Nth (0 = first) mark of that type, e.g. Assignment 1 / 2.
 *             Without nth, all marks of that type are added together.
 *   minCount  (optional) show N/A unless the module has at least this many marks of that type
 *   optional  (optional) leave the cell alone when the module has no such mark
 *             (used for a 2nd assignment row that is merged for most modules)
 * A module without that kind of mark gets "N/A".
 */
const TRANSCRIPT_STANDARD = {
  range: 'A1:K37',
  nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
  firstRow: 6, lastRow: 29, moduleCol: 2, earnedCol: 10, gradeCol: 11,
  averageCell: 'J31', averageGradeCell: 'K31',
  credentialCell: 'I4', passMark: 60
};

const PROGRAMS = [
  {
    key: 'IBA',
    label: 'International Business Administration',
    dataSheet: 'IBA',
    marksheetSheet: 'IBA Marksheet',
    transcriptSheet: 'IBA Transcript',
    marksheet: {
      range: 'A1:Z9',
      idCell: 'C2', nameCell: 'A3', startCell: 'P8', endCell: 'W8', finalMarkCell: 'E8',
      moduleNameRow: 4, firstCol: 3,
      rows: [
        { type: 'participation', row: 5 },
        { type: 'final', row: 6 }
      ],
      averageRow: 7
    },
    transcript: TRANSCRIPT_STANDARD,
    grades: GRADES_STANDARD
  },
  {
    key: 'HT',
    label: 'Hospitality and Tourism',
    dataSheet: 'HT',
    marksheetSheet: 'HT Marksheet',
    transcriptSheet: 'HT Transcript',
    marksheet: {
      range: 'A1:Z9',
      idCell: 'C2', nameCell: 'A3', startCell: 'P8', endCell: 'W8', finalMarkCell: 'E8',
      moduleNameRow: 4, firstCol: 3,
      rows: [
        { type: 'participation', row: 5 },
        { type: 'final', row: 6 }
      ],
      averageRow: 7
    },
    transcript: TRANSCRIPT_STANDARD,
    grades: GRADES_STANDARD
  },
  {
    key: 'CSW',
    label: 'Community Service Worker',
    dataSheet: 'CSW',
    marksheetSheet: 'CSW Marksheet',
    transcriptSheet: 'CSW Transcript',
    marksheet: {
      range: 'A1:U11',
      idCell: 'C2', nameCell: 'A3', startCell: 'K10', endCell: 'R10', finalMarkCell: 'D10',
      moduleNameRow: 5, firstCol: 3,
      rows: [
        { type: 'assignment', row: 6 },
        { type: 'participation', row: 7 },
        { type: 'final', row: 8 }
      ],
      averageRow: 9
    },
    transcript: {
      range: 'A1:K32',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 24, moduleCol: 2, earnedCol: 10, gradeCol: 11,
      averageCell: 'J26', averageGradeCell: 'K26',
      credentialCell: 'I4', passMark: 70
    },
    grades: GRADES_HIGH
  },
  {
    key: 'MOA',
    label: 'Medical Office Administration',
    dataSheet: 'MOA',
    marksheetSheet: 'MOA Marksheet',
    transcriptSheet: 'MOA Transcript',
    // Template spelling -> data sheet spelling
    aliases: { 'Buisness Communication': 'Business Communication' },
    marksheet: {
      range: 'A1:L13',
      idCell: 'C3', nameCell: 'A4', startCell: 'G12', endCell: 'K12', finalMarkCell: 'C12',
      moduleNameRow: 6, firstCol: 3,
      rows: [
        { type: 'assignment', row: 7 },       // all assignments added together
        { type: 'participation', row: 8 },    // "Discretionary"
        { type: 'midterm', row: 9 },
        { type: 'final', row: 10 }
      ],
      averageRow: 11
    },
    transcript: {
      range: 'A1:K23',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 15, moduleCol: 2, earnedCol: 10, gradeCol: 11,
      averageCell: 'J17', averageGradeCell: 'K17',
      credentialCell: 'I4', passMark: 60
    },
    grades: GRADES_STANDARD
  },
  {
    key: 'AC_NEW',
    label: 'Accounting, Payroll, Business & Tax (New)',
    dataSheet: 'AC2 (New)',
    marksheetSheet: 'AC2 (New) Marksheet',
    transcriptSheet: 'AC2 (New) Transcript',
    aliases: {
      'Microsoft Word Level 2': 'Microsoft Word 2',
      'Microsoft Excel Level 1 & Excel 2': 'Microsoft Excel 1 and Excel 2',
      'Payroll Fundamentals Level 1 & 2': 'Payroll Fundamentals 1 & 2'
    },
    marksheet: {
      range: 'A1:N13',
      idCell: 'C2', nameCell: 'A3', startCell: 'G12', endCell: 'L12', finalMarkCell: 'C12',
      moduleNameRow: 5, firstCol: 3,
      rows: [
        { type: 'assignment', row: 6, nth: 0 },
        { type: 'assignment', row: 7, nth: 1, optional: true },   // only Sage has two assignments
        { type: 'quiz', row: 8 },
        { type: 'participation', row: 9 },
        { type: 'final', row: 10 }
      ],
      averageRow: 11
    },
    transcript: {
      range: 'A1:K26',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 17, moduleCol: 2, earnedCol: 10, gradeCol: 11,
      averageCell: 'J18', averageGradeCell: 'K18',
      credentialCell: 'I4', passMark: 70      // 60% per module, 70% overall average
    },
    grades: GRADES_STANDARD,                  // module grades
    averageGrades: GRADES_HIGH                // "Final Grading System" for the overall average
  },
  {
    key: 'AC_OLD',
    label: 'Accounting, Payroll, Business & Tax (Old)',
    dataSheet: 'AC2 (Old)',
    marksheetSheet: 'AC2 (Old) Marksheet',
    transcriptSheet: 'AC2 (Old) Transcript',
    aliases: {
      'Microsoft Word Level 2': 'Microsoft Word 2',
      'Microsoft Excel Level 1 & Excel 2': 'Microsoft Excel 1 and Excel 2',
      'Payroll Fundamentals Level 1': 'Payroll Fundamentals 1',
      'Payroll Fundamental Level 2': 'Payroll Fundamental 2'
    },
    // In this data sheet the Sage module's "Assignment 2" column is a quiz mark
    labelTypes: { 'assignment 2': 'quiz' },
    marksheet: {
      range: 'A1:O12',
      idCell: 'C2', nameCell: 'A3', startCell: 'H11', endCell: 'M11', finalMarkCell: 'C11',
      moduleNameRow: 5, firstCol: 3,
      rows: [
        { type: 'assignment', row: 6 },
        { type: 'quiz', row: 7 },
        { type: 'participation', row: 8 },
        { type: 'final', row: 9 }
      ],
      averageRow: 10
    },
    transcript: {
      range: 'A1:K27',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 18, moduleCol: 2, earnedCol: 10, gradeCol: 11,
      averageCell: 'J19', averageGradeCell: 'K19',
      credentialCell: 'I4', passMark: 70
    },
    grades: GRADES_STANDARD,
    averageGrades: GRADES_HIGH
  },
  {
    key: 'PSW',
    label: 'Personal Support Worker (NACC)',
    dataSheet: 'PSW',
    marksheetSheet: 'PSW Marksheet',
    transcriptSheet: 'PSW Transcript',
    transcriptPassSheet: 'PSW Transcript Pass',   // used only when the NACC Exam is PASS
    naccModule: 'NACC Exam',
    aliases: {
      'PSW Foundations': 'PSW Foundation',
      'Care Planning/Restorative/Documentation/Working in the Community':
        'Care Planning/Restorative Care/Electronic Documentation/ Working in the Community'
    },
    labelTypes: { 'pratical': 'practical' },      // typo in the data sheet header
    marksheet: {
      range: 'A1:Q17',
      idCell: 'C2', nameCell: 'A3', startCell: 'E16', endCell: 'N16',
      moduleNameRow: 5, firstCol: 3,
      rows: [
        { type: 'assignment', row: 6, nth: 0, minCount: 2 },   // Assignment 1 (only when there are two)
        { type: 'assignment', row: 7, nth: 1, minCount: 2 },   // Assignment 2
        { type: 'assignment', row: 8 },                        // total assignments
        { type: 'participation', row: 9 },                     // Discretionary (none in PSW)
        { type: 'practical', row: 10 },
        { type: 'theory', row: 11, nth: 0 },
        { type: 'theory', row: 12, nth: 1 },
        { type: 'theory', row: 13, nth: 2 },
        { type: 'theory', row: 14 },                           // total theory
        { type: 'status', row: 6 }                             // GPA / clinical placements
      ],
      averageRow: 15
    },
    transcript: {
      range: 'A1:K28',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 20, marksLastRow: 17, moduleCol: 2, earnedCol: 10, gradeCol: 11, statusCol: 8,
      averageCell: 'J22', averageGradeCell: 'K22',
      credentialCell: 'I4', passMark: 70
    },
    transcriptPass: {
      range: 'A1:K29',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 21, marksLastRow: 17, moduleCol: 2, earnedCol: 10, gradeCol: 11, statusCol: 8,
      averageCell: 'J23', averageGradeCell: 'K23',
      credentialCell: 'I4', passMark: 70
    },
    grades: GRADES_HIGH
  },
  {
    key: 'ECA',
    label: 'Early Childhood Assistant (NACC)',
    dataSheet: 'ECA',
    marksheetSheet: 'ECA Marksheet',
    transcriptSheet: 'ECA Transcript',
    transcriptPassSheet: 'ECA Transcript Pass',   // used only when the NACC Exam is PASS
    naccModule: 'NACC Exam',
    marksheet: {
      range: 'A1:Q15',
      idCell: 'C2', nameCell: 'A3', startCell: 'F14', endCell: 'N14',
      moduleNameRow: 5, firstCol: 3,
      rows: [
        { type: 'assignment', row: 6, nth: 0 },
        { type: 'assignment', row: 7, nth: 1 },
        { type: 'presentation', row: 8 },
        { type: ['assignment', 'presentation'], row: 9 },      // total
        { type: 'participation', row: 10 },                    // Discretionary (none in ECA)
        { type: 'midterm', row: 11 },
        { type: 'final', row: 12 },
        { type: 'status', row: 6 }                             // Practicum I / II
      ],
      averageRow: 13
    },
    transcript: {
      range: 'A1:K28',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 20, marksLastRow: 19, moduleCol: 2, earnedCol: 10, gradeCol: 11, statusCol: 8,
      averageCell: 'J22', averageGradeCell: 'K22',
      credentialCell: 'I4', passMark: 70
    },
    transcriptPass: {
      range: 'A1:K29',
      nameCell: 'C2', idCell: 'C3', programCell: 'C4', startCell: 'I2', endCell: 'I3',
      firstRow: 6, lastRow: 23, marksLastRow: 19, moduleCol: 2, earnedCol: 10, gradeCol: 11, statusCol: 8,
      statusCols: { 'nacc exam': 7 },                          // NACC result sits in column G on this sheet
      averageCell: 'J22', averageGradeCell: 'K22',
      credentialCell: 'I4', passMark: 70
    },
    grades: GRADES_HIGH
  }
];


/* ============================== WEB APP ============================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Marksheet & Transcript')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Called from the browser.
 * - Searches every program for the ID (or only `programKey` if given).
 * - If the ID exists in more than one program, returns { ambiguous: [...] } so
 *   the page can ask which one to open.
 * - Otherwise returns both PDFs (base64) plus student details.
 */
function getStudentDocuments(rawId, programKey) {
  const id = String(rawId == null ? '' : rawId).trim();
  if (!id) throw new Error('Please enter a Student ID.');

  // The templates are shared sheets, so only one request may use them at a time.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    throw new Error('The system is busy. Please try again in a moment.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const matches = [];
    PROGRAMS.forEach(function (p) {
      if (programKey && p.key !== programKey) return;
      const sheet = mustGetSheet_(ss, p.dataSheet);
      const student = readStudent_(sheet.getDataRange().getValues(), id, p);
      if (student) matches.push({ program: p, student: student });
    });

    if (!matches.length) throw new Error('No student found with ID "' + id + '".');
    if (matches.length > 1) {
      return {
        ambiguous: matches.map(function (m) {
          return { key: m.program.key, label: m.program.label, name: m.student.fullName };
        })
      };
    }

    const program = matches[0].program;
    const student = matches[0].student;
    const pick = chooseTranscript_(program, student);
    const msSheet = mustGetSheet_(ss, program.marksheetSheet);
    const trSheet = mustGetSheet_(ss, pick.sheetName);
    const T = pick.config;

    fillMarksheet_(msSheet, program, student);
    fillTranscript_(trSheet, program, student, T);
    SpreadsheetApp.flush();

    const safeName = student.fullName.replace(/[\\/:*?"<>|]+/g, '').trim();
    return {
      studentId: student.id,
      studentName: student.fullName,
      program: String(trSheet.getRange(T.programCell).getDisplayValue()).trim() || program.label,
      filename: student.id + ' - ' + safeName + ' - Marksheet & Transcript.pdf',
      marksheetBase64: exportPdf_(ss, msSheet, program.marksheet.range),
      transcriptBase64: exportPdf_(ss, trSheet, T.range)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Programs with a "Pass" transcript (PSW, ECA) use it only when the student's
 * NACC Exam result is PASS; FAIL (or blank) uses the regular transcript.
 */
function chooseTranscript_(program, student) {
  if (program.transcriptPassSheet && program.naccModule) {
    const nacc = student.modules[moduleKey_(program.naccModule)];
    if (nacc && String(nacc.status).trim().toUpperCase() === 'PASS') {
      return { sheetName: program.transcriptPassSheet, config: program.transcriptPass };
    }
  }
  return { sheetName: program.transcriptSheet, config: program.transcript };
}

/* ============================ DATA READING ============================ */

/**
 * Pure function: finds the student in a program data sheet (2D array of values).
 * A module is a titled column in row 1 plus the untitled (merged) columns after it;
 * each column's sub-header says which kind of mark it holds.
 * Returns { id, fullName, start, finish, modules: { key: { name, parts, average } } }
 * or null when the ID is not in this sheet. parts = { assignment: [..], quiz: [..],
 * participation: [..], midterm: [..], final: [..] } (a type is empty when the module lacks it).
 */
function readStudent_(values, id, program) {
  const header = values[HEADER_ROW - 1] || [];
  const sub = values[SUBHEADER_ROW - 1] || [];
  const F = FIXED_HEADERS;
  const labelTypes = (program && program.labelTypes) || {};

  const colOf = function (title) {
    const i = header.findIndex(function (h) { return norm_(h) === norm_(title); });
    if (i < 0) throw new Error('Column "' + title + '" was not found in row ' + HEADER_ROW + ' of a data sheet.');
    return i;
  };
  const cId = colOf(F.id), cFirst = colOf(F.first), cLast = colOf(F.last),
        cStart = colOf(F.start), cFinish = colOf(F.finish);
  const skip = [cId, cFirst, cLast, cStart, cFinish];
  header.forEach(function (h, i) { if (norm_(h) === 'status') skip.push(i); });

  // Find the row
  let row = null;
  for (let r = FIRST_DATA_ROW - 1; r < values.length; r++) {
    if (String(values[r][cId]).trim() === id) { row = values[r]; break; }
  }
  if (!row) return null;

  // Titled columns start a module; it runs until the next titled column
  const titled = [];
  for (let c = 0; c < header.length; c++) {
    if (skip.indexOf(c) < 0 && String(header[c]).trim()) titled.push(c);
  }

  const modules = {};
  titled.forEach(function (start, i) {
    const end = i + 1 < titled.length ? titled[i + 1] : header.length;
    const name = String(header[start]).trim();
    const parts = { assignment: [], quiz: [], participation: [], midterm: [], final: [],
                    theory: [], practical: [], presentation: [] };
    let total = 0, anyValue = false, status;

    for (let c = start; c < end; c++) {
      const label = norm_(sub[c]);
      if (!label) continue;
      const type = componentType_(labelTypes[label] || label);
      if (!type) throw new Error('Unknown mark column "' + sub[c] + '" under "' + name + '" in a data sheet.');
      if (type === 'status') {            // PASS / FAIL / "Passed GPA 1&2": text, not part of the average
        status = String(row[c] == null ? '' : row[c]).trim();
        continue;
      }
      const v = toNum_(row[c]);
      parts[type].push(v);
      if (v !== null) { total += v; anyValue = true; }
    }
    modules[norm_(name)] = { name: name, parts: parts, status: status, average: anyValue ? total : null };
  });

  const first = String(row[cFirst]).trim();
  const last = String(row[cLast]).trim();
  return {
    id: String(row[cId]).trim(),
    fullName: (first + ' ' + last).trim(),
    start: row[cStart],
    finish: row[cFinish],
    modules: modules
  };
}

/** "assignment 2" -> "assignment", "theory 1" -> "theory", "score"/"status" -> "status"; null if unrecognised. */
function componentType_(label) {
  if (String(label).indexOf('score') === 0 || String(label).indexOf('status') === 0) return 'status';
  const types = ['assignment', 'quiz', 'participation', 'midterm', 'final', 'theory', 'practical', 'presentation'];
  for (let i = 0; i < types.length; i++) {
    if (String(label).indexOf(types[i]) === 0) return types[i];
  }
  return null;
}

/* ============================ TEMPLATE FILLING ============================ */

function fillMarksheet_(sheet, program, student) {
  const M = program.marksheet;
  const n = sheet.getRange(M.range).getLastColumn() - M.firstCol + 1;
  const names = sheet.getRange(M.moduleNameRow, M.firstCol, 1, n).getValues()[0];

  const b = buildMarksheetRows_(student, names, program);

  b.rows.forEach(function (r) {
    const hasSkips = r.values.some(function (v) { return v === null; });
    if (!hasSkips) {
      sheet.getRange(r.row, M.firstCol, 1, n).setNumberFormat('0%').setValues([r.values]);
    } else {
      r.values.forEach(function (v, i) {
        if (v !== null) sheet.getRange(r.row, M.firstCol + i).setNumberFormat('0%').setValue(v);
      });
    }
  });

  sheet.getRange(M.idCell).setValue(isNaN(Number(student.id)) ? student.id : Number(student.id));
  sheet.getRange(M.nameCell).setValue(student.fullName);
  sheet.getRange(M.startCell).setValue(student.start);
  sheet.getRange(M.endCell).setValue(student.finish);
  if (M.finalMarkCell) setUnlessFormula_(sheet.getRange(M.finalMarkCell), b.finalMark);
}

/**
 * Pure: { rows: [{row, values}], finalMark } where each values entry is a fraction,
 * 'N/A', '' (blank mark), text (pass/fail result) or null (leave the cell alone),
 * in template column order.
 */
function buildMarksheetRows_(student, templateNames, program) {
  const M = program.marksheet;
  const aliasMap = aliasMap_(program);
  const mods = templateNames.map(function (n) { return findModule_(student, n, aliasMap); });
  const isStatus = function (m) { return m.status !== undefined; };

  const rows = M.rows.map(function (spec) {
    return {
      row: spec.row,
      values: mods.map(function (m) {
        // Pass/fail modules (placements, GPA) only use the 'status' row
        if (spec.type === 'status') return isStatus(m) ? m.status : null;
        if (isStatus(m)) return null;

        const types = [].concat(spec.type);
        const list = [];
        types.forEach(function (t) { (m.parts[t] || []).forEach(function (v) { list.push(v); }); });
        const miss = spec.optional ? null : 'N/A';

        if (list.length < (spec.minCount || 0)) return miss;
        if (spec.nth !== undefined) {
          if (list.length <= spec.nth) return miss;
          return pct_(list[spec.nth]);
        }
        if (!list.length) return miss;
        const nums = list.filter(function (v) { return v !== null; });
        return nums.length ? pct_(nums.reduce(function (a, v) { return a + v; }, 0)) : '';
      })
    };
  });

  rows.push({ row: M.averageRow, values: mods.map(function (m) { return isStatus(m) ? null : pct_(m.average); }) });

  const nums = mods.filter(function (m) { return !isStatus(m) && m.average !== null; })
                   .map(function (m) { return m.average; });
  return {
    rows: rows,
    finalMark: nums.length ? nums.reduce(function (a, v) { return a + v; }, 0) / nums.length / 100 : ''
  };
}

function fillTranscript_(sheet, program, student, T) {
  const count = T.lastRow - T.firstRow + 1;
  const names = sheet.getRange(T.firstRow, T.moduleCol, count, 1).getValues().map(function (r) { return r[0]; });

  const t = buildTranscriptBlock_(student, names, program);

  // Write marks in runs of consecutive module rows; text results go in their own cell
  let i = 0;
  while (i < count) {
    const e = t.rows[i];
    if (e.kind === 'mark' || e.kind === 'empty') {
      const block = [];
      let j = i;
      while (j < count && (t.rows[j].kind === 'mark' || t.rows[j].kind === 'empty')) {
        block.push([t.rows[j].earned, t.rows[j].grade]);
        j++;
      }
      sheet.getRange(T.firstRow + i, T.earnedCol, block.length, 2).setValues(block);
      i = j;
    } else {
      if (e.kind === 'status') {
        const col = (T.statusCols && T.statusCols[e.key]) || T.statusCol || 8;
        sheet.getRange(T.firstRow + i, col).setValue(e.text);
      }
      i++;
    }
  }

  setUnlessFormula_(sheet.getRange(T.averageCell), t.average);
  sheet.getRange(T.averageGradeCell).setValue(t.averageGrade);
  sheet.getRange(T.nameCell).setValue(student.fullName);
  sheet.getRange(T.idCell).setNumberFormat('@').setValue(String(student.id));
  sheet.getRange(T.startCell).setValue(student.start);
  sheet.getRange(T.endCell).setValue(student.finish);

  if (T.credentialCell) {
    // A live formula so the cell always follows the earned marks (module rows only).
    const lastMarkRow = T.marksLastRow || T.lastRow;
    const marks = sheet.getRange(T.firstRow, T.earnedCol, lastMarkRow - T.firstRow + 1, 1).getA1Notation();
    sheet.getRange(T.credentialCell).setFormula(credentialFormula_(marks, T.passMark));
  }
  // "Date Issued" is a formula in the templates and updates on its own.
}

/** Pure: "Credential Awarded" formula for a range of earned marks. */
function credentialFormula_(marksRange, passMark) {
  const avg = 'ROUND(AVERAGE(' + marksRange + '),6)';   // ROUND avoids 89.4999999 floating-point misses
  return '=IF(COUNT(' + marksRange + ')=0,"",' +
         'IF(' + avg + '<' + (passMark / 100) + ',"' + CREDENTIAL_TEXT.fail + '",' +
         'IF(' + avg + '>=' + (HONOURS_MARK / 100) + ',"' + CREDENTIAL_TEXT.honours + '","' + CREDENTIAL_TEXT.pass + '")))';
}

/**
 * Pure: one entry per template row, in order:
 *   { kind:'skip' }                       blank name cell (labels, spacing rows)
 *   { kind:'mark', earned, grade }        graded module
 *   { kind:'empty', earned:'', grade:'' } module without marks
 *   { kind:'status', key, text }          pass/fail module (placements, GPA, NACC)
 * plus the overall average and its grade.
 */
function buildTranscriptBlock_(student, templateNames, program) {
  const aliasMap = aliasMap_(program);
  const avgScale = program.averageGrades || program.grades;
  const rows = [];
  const scores = [];
  templateNames.forEach(function (n) {
    if (!String(n == null ? '' : n).trim()) { rows.push({ kind: 'skip' }); return; }
    const m = findModule_(student, n, aliasMap);
    if (m.status !== undefined) {
      rows.push({ kind: 'status', key: moduleKey_(n) in aliasMap ? aliasMap[moduleKey_(n)] : moduleKey_(n), text: m.status });
    } else if (m.average === null) {
      rows.push({ kind: 'empty', earned: '', grade: '' });
    } else {
      rows.push({ kind: 'mark', earned: m.average / 100, grade: gradeFor_(m.average, program.grades) });
      scores.push(m.average);
    }
  });
  const avgPoints = scores.length ? scores.reduce(function (a, b) { return a + b; }, 0) / scores.length : null;
  return {
    rows: rows,
    average: avgPoints === null ? '' : avgPoints / 100,
    averageGrade: avgPoints === null ? '' : gradeFor_(avgPoints, avgScale, true)
  };
}

/** Grade for a mark (out of 100) on the given scale. `exact` skips GRADE_ROUNDING (used for the average). */
function gradeFor_(points, grades, exact) {
  // The tiny offset only absorbs floating-point noise (e.g. 84.49999999999999)
  const p = (GRADE_ROUNDING === 'round' && !exact) ? Math.floor(points + 0.5 + 1e-9) : points + 1e-9;
  for (let i = 0; i < grades.length; i++) {
    if (p >= grades[i][0]) return grades[i][1];
  }
  return 'F';
}

/** Template module name -> student's module (handles aliases, "(M3)" tags, stray spaces). */
function findModule_(student, templateName, aliasMap) {
  let key = moduleKey_(templateName);
  if (aliasMap[key]) key = aliasMap[key];
  const m = student.modules[key];
  if (!m) {
    throw new Error('Template module "' + String(templateName).trim() +
      '" was not found in the data sheet. Check its spelling, or add it to "aliases" for this program in Code.gs.');
  }
  return m;
}

function aliasMap_(program) {
  const map = {};
  Object.keys(program.aliases || {}).forEach(function (k) {
    map[moduleKey_(k)] = moduleKey_(program.aliases[k]);
  });
  return map;
}

/** Writes a value unless the cell holds a formula (keeps the template's own formulas). */
function setUnlessFormula_(range, value) {
  if (!range.getFormula()) range.setValue(value);
}

/* ============================== PDF EXPORT ============================== */

function exportPdf_(ss, sheet, a1Range) {
  const P = PDF_SETTINGS;
  const params = {
    format: 'pdf',
    gid: sheet.getSheetId(),
    range: a1Range,
    size: P.size,
    portrait: P.portrait,
    scale: 4,                         // fit to page
    top_margin: P.margin,
    bottom_margin: P.margin,
    left_margin: P.margin,
    right_margin: P.margin,
    horizontal_alignment: 'CENTER',
    vertical_alignment: 'TOP',
    gridlines: false,
    printtitle: false,
    sheetnames: false,
    pagenumbers: false,
    fzr: false,
    fzc: false
  };
  const query = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + query;

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('PDF export failed (HTTP ' + res.getResponseCode() + '). Re-authorise the script and try again.');
  }
  return Utilities.base64Encode(res.getContent());
}

/* ============================== HELPERS ============================== */

function mustGetSheet_(ss, name) {
  const s = ss.getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" was not found.');
  return s;
}

function norm_(v) {
  return String(v == null ? '' : v).replace(/[\u2018\u2019]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Module name used for matching: ignores case/spacing and a trailing "(M12)" tag. */
function moduleKey_(v) {
  return norm_(String(v == null ? '' : v).replace(/\(\s*M\d+\s*\)\s*$/i, ''));
}

function toNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

/** Marks (out of 100) -> fraction for the % formatted template cells; blank stays blank. */
function pct_(points) {
  return points === null ? '' : points / 100;
}