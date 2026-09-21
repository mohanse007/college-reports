/**
 * Course Allocation Report Engine (Client-side / Browser)
 * Uses ExcelJS to parse raw SIS/ERP Excel files and generate college reports.
 */

function getOrdinalSuffix(n) {
  const num = parseInt(n, 10);
  if (isNaN(num)) return `${n}th`;
  if (num % 100 >= 11 && num % 100 <= 13) return `${num}th`;
  const last = num % 10;
  if (last === 1) return `${num}st`;
  if (last === 2) return `${num}nd`;
  if (last === 3) return `${num}rd`;
  return `${num}th`;
}

function cleanText(val) {
  if (val === null || val === undefined) return '';
  let s = String(val);
  s = s.replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '"');
  s = s.replace(/[–—]/g, '-').replace(/[�]/g, '');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function formatCourseName(name) {
  let cleaned = cleanText(name);
  cleaned = cleaned.replace(/[\s\-]+[\(\[]?\s*[Tt]\s*[\)\]]?$/, ' (T)');
  cleaned = cleaned.replace(/[\s\-]+[\(\[]?\s*[Pp]\s*[\)\]]?$/, ' (P)');
  return cleaned;
}

const DEPARTMENT_ORDER = [
  'Languages',
  'Agriculture',
  'BBA',
  'Chemistry',
  'B COM',
  'Computer Science',
  'Maths',
  'Microbiology',
  'Spl Eng',
  'Zoology',
  'Physics',
  'Bio-Technology',
  'History',
  'BCA'
];

const BATCH_TO_DEPT_MAP = {
  'bsc mathematics': 'Maths',
  'bsc physics': 'Physics',
  'bsc computer science': 'Computer Science',
  'bsc microbiology': 'Microbiology',
  'bsc chemistry': 'Chemistry',
  'bsc agriculture': 'Agriculture',
  'bsc bio-technology': 'Bio-Technology',
  'bsc biotechnology': 'Bio-Technology',
  'bsc zoology': 'Zoology',
  'ba special english': 'Spl Eng',
  'ba history': 'History',
  'b.com general': 'B COM',
  'b.com computers': 'B COM',
  'b com general': 'B COM',
  'b com computers': 'B COM',
  'bcom': 'B COM',
  'bba': 'BBA',
  'bc': 'B COM',
  'bcm': 'B COM',
  'bca': 'BCA',
  'bscs': 'Computer Science',
  'bsm': 'Maths',
  'bsmb': 'Microbiology',
  'ba': 'Spl Eng',
  'bsz': 'Zoology',
  'bsp': 'Physics',
  'bsbt': 'Bio-Technology',
  'bah': 'History'
};

class CourseReportEngine {
  constructor(facultyList = []) {
    this.setFacultyList(facultyList);
  }

  setFacultyList(facultyList) {
    this.facultyList = facultyList || [];
    this.facultyMap = new Map();
    this.phoneMap = new Map();

    for (const f of this.facultyList) {
      const origName = cleanText(f.name);
      const phone = cleanText(f.phone);
      if (!origName) continue;

      const norm = this.normalizeName(origName);
      this.facultyMap.set(norm, { original: origName, phone, dept: f.department || '' });

      const simple = origName.toLowerCase();
      if (!this.phoneMap.has(simple) && phone) {
        this.phoneMap.set(simple, phone);
      }
    }
  }

  normalizeName(name) {
    if (!name) return '';
    let cleaned = String(name).toLowerCase();
    cleaned = cleaned.replace(/[\.\-_]/g, ' ');
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    return cleaned;
  }

  lookupPhone(facultyName) {
    if (!facultyName) return '';
    const nameStr = cleanText(facultyName);
    if (!nameStr) return '';

    if (nameStr.includes(',')) {
      const parts = nameStr.split(',').map(p => p.trim()).filter(Boolean);
      const phones = parts.map(p => this.lookupSinglePhone(p)).filter(Boolean);
      return phones.length > 0 ? phones.join(', ') : '';
    }

    return this.lookupSinglePhone(nameStr);
  }

  lookupSinglePhone(name) {
    const raw = cleanText(name);
    if (!raw) return '';

    const simple = raw.toLowerCase();
    if (this.phoneMap.has(simple)) {
      return this.phoneMap.get(simple);
    }

    const norm = this.normalizeName(raw);
    if (this.facultyMap.has(norm)) {
      return this.facultyMap.get(norm).phone;
    }

    for (const [key, info] of this.facultyMap.entries()) {
      if (norm.length >= 4 && (key.includes(norm) || norm.includes(key))) {
        if (info.phone) return info.phone;
      }
    }

    const normTokens = new Set(norm.split(' ').filter(t => t.length > 1));
    if (normTokens.size > 0) {
      let bestMatch = null;
      let maxOverlap = 0;

      for (const [key, info] of this.facultyMap.entries()) {
        const keyTokens = new Set(key.split(' ').filter(t => t.length > 1));
        let overlap = 0;
        for (const t of normTokens) {
          if (keyTokens.has(t)) overlap++;
        }

        if (overlap >= 2 && overlap > maxOverlap) {
          maxOverlap = overlap;
          bestMatch = info;
        } else if (overlap === 1 && normTokens.size === 1 && keyTokens.size === 1 && overlap > maxOverlap) {
          maxOverlap = overlap;
          bestMatch = info;
        }
      }

      if (bestMatch && bestMatch.phone) {
        return bestMatch.phone;
      }
    }

    return '';
  }

  mapBatchToDept(batchName) {
    if (!batchName) return '';
    let cleaned = String(batchName).replace(/\s*\b\d{4}\b.*$/g, '').trim().toLowerCase();
    if (BATCH_TO_DEPT_MAP[cleaned]) {
      return BATCH_TO_DEPT_MAP[cleaned];
    }
    let fb = cleaned.replace(/^(bsc|ba|bcom|bba|bca)\s+/i, '').trim();
    return fb ? (fb.charAt(0).toUpperCase() + fb.slice(1)) : cleanText(batchName);
  }

  extractMetadata(worksheet) {
    let collegeName = "St.Ann's College for Women (A)";
    let semesterTitle = "4th Semester Course Code, Course name with faculty Details";
    let admittedBatch = "Admitted Batch 2024-2025";
    let termDetected = "";
    let yearDetected = "";

    try {
      let metaStr = "";
      const maxRows = Math.min(10, worksheet.rowCount || 10);
      for (let r = 1; r <= maxRows; r++) {
        const row = worksheet.getRow(r);
        const cellVal = cleanText(row.getCell(1).value);
        if (cellVal.includes("Term:") || cellVal.includes("Batch Start Year:")) {
          metaStr = cellVal;
          break;
        }
      }

      if (metaStr) {
        const termMatch = metaStr.match(/Term:\s*S?(\d+)/i);
        if (termMatch) {
          const semNum = parseInt(termMatch[1], 10);
          termDetected = `S${semNum}`;
          const ord = getOrdinalSuffix(semNum);
          semesterTitle = `${ord} Semester Course Code, Course name with faculty Details`;
        }

        const yearMatch = metaStr.match(/Batch Start Year:\s*(\d{4})/i);
        if (yearMatch) {
          const startYear = parseInt(yearMatch[1], 10);
          yearDetected = `${startYear}`;
          admittedBatch = `Admitted Batch ${startYear}-${startYear + 1}`;
        }
      }
    } catch (err) {
      console.warn("Warning extracting metadata:", err);
    }

    return {
      college_name: collegeName,
      semester_title: semesterTitle,
      admitted_batch: admittedBatch,
      term: termDetected,
      year: yearDetected
    };
  }

  parseRawData(worksheet) {
    let headerRow = 5;
    const maxScan = Math.min(15, worksheet.rowCount || 15);
    for (let r = 1; r <= maxScan; r++) {
      const val = cleanText(worksheet.getRow(r).getCell(2).value).toLowerCase();
      if (val.includes("course code")) {
        headerRow = r;
        break;
      }
    }

    const records = [];
    let lastCourseCode = "";
    let lastCourseName = "";
    let lastFaculty = "";
    let lastCourseComm = "";

    const totalRows = worksheet.rowCount || 0;
    for (let r = headerRow + 1; r <= totalRows; r++) {
      const row = worksheet.getRow(r);
      let cCode = cleanText(row.getCell(2).value);
      let cName = cleanText(row.getCell(3).value);
      let cComm = cleanText(row.getCell(4).value);
      let faculty = cleanText(row.getCell(5).value);
      let batch = cleanText(row.getCell(6).value);

      if (!batch) continue;

      if (cCode) {
        lastCourseCode = cCode;
        lastCourseName = cName;
        lastFaculty = faculty;
        lastCourseComm = cComm;
      } else {
        cCode = lastCourseCode;
        if (!cName) cName = lastCourseName;
        if (!faculty) faculty = lastFaculty;
        if (!cComm) cComm = lastCourseComm;
      }

      if (!cCode && cComm) {
        const parts = cComm.split('-');
        if (parts.length >= 2) {
          const possibleCode = parts[0].trim().toUpperCase();
          if (["CVAC", "SDS", "SD", "MD", "ENG", "TEL", "HIN", "SAN", "SKT"].some(p => possibleCode.startsWith(p))) {
            cCode = possibleCode;
            cName = parts.slice(1).join('-').trim();
          }
        }
      }

      const codeUpper = cCode.trim().toUpperCase();
      if (["GAMES", "LIBRARY", "MI"].includes(codeUpper) || ["GAMES", "LIBRARY"].includes(cName.trim().toUpperCase())) {
        continue;
      }

      const dept = this.mapBatchToDept(batch);
      const isLanguage = ["ENG", "TEL", "HIN", "SAN", "SKT"].some(p => codeUpper.startsWith(p));

      records.push({
        course_code: cCode,
        course_name: formatCourseName(cName),
        course_comm: cComm,
        faculty_name: faculty,
        batch_name: batch,
        dept: isLanguage ? "Languages" : dept,
        is_language: isLanguage
      });
    }

    return records;
  }

  sortCoursesInDept(records) {
    const theoryList = [];
    const practicalMap = new Map();
    const seen = new Set();

    for (const rec of records) {
      const key = `${rec.course_code.toUpperCase()}|||${rec.course_name.toUpperCase()}|||${rec.faculty_name.toUpperCase()}`;
      if (seen.has(key)) continue;
      seen.add(key);

      const code = rec.course_code.trim();
      if (code.toUpperCase().startsWith("P ")) {
        const base = code.slice(2).trim().toUpperCase();
        practicalMap.set(base, rec);
      } else {
        theoryList.push(rec);
      }
    }

    function getSortKey(rec) {
      const codeU = rec.course_code.trim().toUpperCase();
      let cat = 1; // Major
      if (["SDS", "SD ", "SD", "MD", "CVAC", "SEC"].some(p => codeU.startsWith(p))) {
        cat = 3; // Skill / Multi-Disciplinary
      } else if (codeU.endsWith("M")) {
        cat = 2; // Minor
      }

      const numMatch = codeU.match(/\d+/);
      const num = numMatch ? parseInt(numMatch[0], 10) : 999;
      const prefixMatch = codeU.match(/^[A-Z]+/);
      const prefix = prefixMatch ? prefixMatch[0] : "";

      return { cat, num, prefix, codeU };
    }

    theoryList.sort((a, b) => {
      const ka = getSortKey(a);
      const kb = getSortKey(b);
      if (ka.cat !== kb.cat) return ka.cat - kb.cat;
      if (ka.num !== kb.num) return ka.num - kb.num;
      if (ka.prefix !== kb.prefix) return ka.prefix.localeCompare(kb.prefix);
      return ka.codeU.localeCompare(kb.codeU);
    });

    const ordered = [];
    for (const tRec of theoryList) {
      ordered.push(tRec);
      const tBase = tRec.course_code.trim().toUpperCase();
      if (practicalMap.has(tBase)) {
        ordered.push(practicalMap.get(tBase));
        practicalMap.delete(tBase);
      }
    }

    for (const pRec of practicalMap.values()) {
      ordered.push(pRec);
    }

    return ordered;
  }

  async generateWorkbook(rawArrayBuffer, overrides = {}) {
    const inWb = new ExcelJS.Workbook();
    await inWb.xlsx.load(rawArrayBuffer);
    const inWs = inWb.worksheets[0];

    const meta = this.extractMetadata(inWs);
    const collegeName = overrides.college_name || meta.college_name;
    const semesterTitle = overrides.semester_title || meta.semester_title;
    const admittedBatch = overrides.admitted_batch || meta.admitted_batch;

    const records = this.parseRawData(inWs);

    const deptRecordsMap = new Map();
    for (const r of records) {
      const d = r.dept;
      if (!deptRecordsMap.has(d)) deptRecordsMap.set(d, []);
      deptRecordsMap.get(d).push(r);
    }

    const outWb = new ExcelJS.Workbook();
    outWb.creator = "St. Ann's College Report Generator";
    outWb.lastModifiedBy = "St. Ann's College Report Generator";
    outWb.created = new Date();
    outWb.modified = new Date();

    // Sheet 1: Department-wise Course List
    const ws1 = outWb.addWorksheet('Sheet1');
    ws1.views = [{ showGridLines: true }];

    const borderThin = {
      top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
    };

    ws1.mergeCells('A1:D1');
    const r1 = ws1.getCell('A1');
    r1.value = collegeName;
    r1.font = { name: 'Calibri', size: 16, bold: true };
    r1.alignment = { horizontal: 'center', vertical: 'center' };

    ws1.mergeCells('A2:D2');
    const r2 = ws1.getCell('A2');
    r2.value = semesterTitle;
    r2.font = { name: 'Calibri', size: 16, bold: true };
    r2.alignment = { horizontal: 'center', vertical: 'center' };

    ws1.mergeCells('A3:D3');
    const r3 = ws1.getCell('A3');
    r3.value = admittedBatch;
    r3.font = { name: 'Calibri', size: 16, bold: true };
    r3.alignment = { horizontal: 'center', vertical: 'center' };

    let currentRow = 4;

    const existingDepts = Array.from(deptRecordsMap.keys());
    const orderedDepts = [];
    for (const d of DEPARTMENT_ORDER) {
      if (existingDepts.includes(d)) orderedDepts.push(d);
    }
    for (const d of existingDepts.sort()) {
      if (!orderedDepts.includes(d)) orderedDepts.push(d);
    }

    const tableHeaders = ['Course Code', 'Course Name', 'Name of Staff', 'Phone Number'];

    for (const dept of orderedDepts) {
      const deptRecs = deptRecordsMap.get(dept);
      const sortedRecs = this.sortCoursesInDept(deptRecs);
      if (sortedRecs.length === 0) continue;

      ws1.mergeCells(`A${currentRow}:D${currentRow}`);
      for (let col = 1; col <= 4; col++) {
        const cell = ws1.getRow(currentRow).getCell(col);
        cell.border = borderThin;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD9D9D9' }
        };
      }
      const deptCell = ws1.getCell(`A${currentRow}`);
      deptCell.value = dept;
      deptCell.font = { name: 'Calibri', size: 14, bold: true };
      deptCell.alignment = { horizontal: 'center', vertical: 'center' };
      currentRow++;

      const thRow = ws1.getRow(currentRow);
      for (let col = 1; col <= 4; col++) {
        const cell = thRow.getCell(col);
        cell.value = tableHeaders[col - 1];
        cell.font = { name: 'Calibri', size: 11, bold: true };
        cell.border = borderThin;
        cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF2F2F2' }
        };
      }
      currentRow++;

      for (const rec of sortedRecs) {
        const dRow = ws1.getRow(currentRow);
        const phone = this.lookupPhone(rec.faculty_name);

        const rowValues = [
          rec.course_code,
          rec.course_name,
          rec.faculty_name,
          phone
        ];

        for (let col = 1; col <= 4; col++) {
          const cell = dRow.getCell(col);
          cell.value = rowValues[col - 1];
          cell.font = { name: 'Calibri', size: 11 };
          cell.border = borderThin;
          const align = (col === 1 || col === 4) ? 'center' : 'left';
          cell.alignment = { horizontal: align, vertical: 'center', wrapText: true };
        }
        currentRow++;
      }

      currentRow++;
    }

    ws1.getColumn(1).width = 16;
    ws1.getColumn(2).width = 50;
    ws1.getColumn(3).width = 32;
    ws1.getColumn(4).width = 20;

    // Sheet 2: Summary List
    const ws2 = outWb.addWorksheet('Sheet2');
    ws2.views = [{ showGridLines: true }];

    const s2Headers = ['Sl.No', 'Subject', 'Subject Code', 'Title of the paper'];
    const s2ThRow = ws2.getRow(1);
    for (let col = 1; col <= 4; col++) {
      const cell = s2ThRow.getCell(col);
      cell.value = s2Headers[col - 1];
      cell.font = { name: 'Calibri', size: 11, bold: true };
      cell.border = borderThin;
      cell.alignment = { horizontal: 'center', vertical: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2F2F2' }
      };
    }

    const categoryCourses = {
      'Languages': [],
      'Skill Course': [],
      'Multi Disciplinary': [],
      'Major': []
    };

    const seenS2 = new Set();
    for (const r of records) {
      const codeU = r.course_code.trim().toUpperCase();
      const name = r.course_name.trim();
      const s2Key = `${codeU}|||${name.toUpperCase()}`;
      if (seenS2.has(s2Key)) continue;
      seenS2.add(s2Key);

      let cat = 'Major';
      if (['ENG', 'TEL', 'HIN', 'SAN', 'SKT'].some(p => codeU.startsWith(p))) {
        cat = 'Languages';
      } else if (['SDS', 'SD ', 'SD', 'SEC'].some(p => codeU.startsWith(p))) {
        cat = 'Skill Course';
      } else if (['MD', 'CVAC'].some(p => codeU.startsWith(p))) {
        cat = 'Multi Disciplinary';
      }

      categoryCourses[cat].push({
        subject: r.dept,
        code: r.course_code,
        title: name
      });
    }

    let s2Row = 2;
    let serialNo = 1;

    for (const [catName, cList] of Object.entries(categoryCourses)) {
      if (cList.length === 0) continue;

      ws2.mergeCells(`A${s2Row}:D${s2Row}`);
      for (let col = 1; col <= 4; col++) {
        const cell = ws2.getRow(s2Row).getCell(col);
        cell.border = borderThin;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFE8E8E8' }
        };
      }
      const catCell = ws2.getCell(`A${s2Row}`);
      catCell.value = catName;
      catCell.font = { name: 'Calibri', size: 11, bold: true };
      catCell.alignment = { horizontal: 'left', vertical: 'center' };
      s2Row++;

      cList.sort((a, b) => a.code.localeCompare(b.code));

      let isFirstInCat = true;
      for (const item of cList) {
        const row = ws2.getRow(s2Row);
        const rowVals = [
          String(serialNo++),
          isFirstInCat ? (catName === 'Major' ? item.subject : catName) : '',
          item.code,
          item.title
        ];

        for (let col = 1; col <= 4; col++) {
          const cell = row.getCell(col);
          cell.value = rowVals[col - 1];
          cell.font = { name: 'Calibri', size: 11 };
          cell.border = borderThin;
          const align = (col === 1 || col === 3) ? 'center' : 'left';
          cell.alignment = { horizontal: align, vertical: 'center' };
        }
        isFirstInCat = false;
        s2Row++;
      }
    }

    ws2.getColumn(1).width = 10;
    ws2.getColumn(2).width = 24;
    ws2.getColumn(3).width = 18;
    ws2.getColumn(4).width = 54;

    return {
      workbook: outWb,
      meta: {
        college_name: collegeName,
        semester_title: semesterTitle,
        admitted_batch: admittedBatch,
        term: meta.term,
        year: meta.year,
        total_departments: orderedDepts.length,
        total_courses: records.length
      },
      departments: orderedDepts,
      deptRecordsMap
    };
  }
}

if (typeof window !== 'undefined') {
  window.CourseReportEngine = CourseReportEngine;
}

if (typeof module !== 'undefined' && module.exports) { module.exports = { CourseReportEngine }; }
