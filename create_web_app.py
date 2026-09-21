# Script to assemble web app files
import os

ENGINE_JS = """/**
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
  s = s.replace(/[\u2013\u2014]/g, '-').replace(/[\ufffd]/g, '');
  s = s.replace(/\\s+/g, ' ').trim();
  return s;
}

function formatCourseName(name) {
  let cleaned = cleanText(name);
  cleaned = cleaned.replace(/[\\s\\-]+[\\(\\[]?\\s*[Tt]\\s*[\\)\\]]?$/, ' (T)');
  cleaned = cleaned.replace(/[\\s\\-]+[\\(\\[]?\\s*[Pp]\\s*[\\)\\]]?$/, ' (P)');
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
    cleaned = cleaned.replace(/[\\.\\-_]/g, ' ');
    cleaned = cleaned.replace(/\\s+/g, ' ').trim();
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
    let cleaned = String(batchName).replace(/\\s*\\b\\d{4}\\b.*$/g, '').trim().toLowerCase();
    if (BATCH_TO_DEPT_MAP[cleaned]) {
      return BATCH_TO_DEPT_MAP[cleaned];
    }
    let fb = cleaned.replace(/^(bsc|ba|bcom|bba|bca)\\s+/i, '').trim();
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
        const termMatch = metaStr.match(/Term:\\s*S?(\\d+)/i);
        if (termMatch) {
          const semNum = parseInt(termMatch[1], 10);
          termDetected = `S${semNum}`;
          const ord = getOrdinalSuffix(semNum);
          semesterTitle = `${ord} Semester Course Code, Course name with faculty Details`;
        }

        const yearMatch = metaStr.match(/Batch Start Year:\\s*(\\d{4})/i);
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

      const numMatch = codeU.match(/\\d+/);
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
"""

with open('course_report_engine.js', 'w', encoding='utf-8') as f:
    f.write(ENGINE_JS + "\nif (typeof module !== 'undefined' && module.exports) { module.exports = { CourseReportEngine }; }\n")

print('course_report_engine.js written successfully.')

APP_JS = """/**
 * Main Web Application Logic for College Report Portal
 */

let facultyList = [];
let currentRawFile = null;
let currentRawBuffer = null;
let detectedMetadata = null;
let reportEngine = null;

// Initialize application on DOM ready
document.addEventListener('DOMContentLoaded', () => {
  initFacultyData();
  initTabs();
  initCourseAllocationUI();
  initFacultyDirectoryUI();
  lucide.createIcons();
});

// -------------------------------------------------------------
// 1. Faculty Data Management & LocalStorage
// -------------------------------------------------------------
function initFacultyData() {
  const stored = localStorage.getItem('college_faculty_directory');
  if (stored) {
    try {
      facultyList = JSON.parse(stored);
    } catch (e) {
      console.warn('Could not parse stored faculty list, using default', e);
      facultyList = [...DEFAULT_FACULTY_DATA];
    }
  } else {
    facultyList = [...DEFAULT_FACULTY_DATA];
  }

  reportEngine = new CourseReportEngine(facultyList);
  updateFacultyStats();
}

function saveFacultyData() {
  localStorage.setItem('college_faculty_directory', JSON.stringify(facultyList));
  reportEngine.setFacultyList(facultyList);
  updateFacultyStats();
}

function updateFacultyStats() {
  const total = facultyList.length;
  const withPhone = facultyList.filter(f => f.phone && f.phone.trim().length > 0).length;
  const badge = document.getElementById('facultyCountBadge');
  if (badge) {
    badge.textContent = `${total} Faculty Members (${withPhone} with phones)`;
  }
}

// -------------------------------------------------------------
// 2. Tab Navigation
// -------------------------------------------------------------
function initTabs() {
  const navBtns = document.querySelectorAll('.nav-tab-btn');
  const panels = document.querySelectorAll('.tab-panel');

  navBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-tab');

      navBtns.forEach(b => {
        b.classList.remove('active', 'border-blue-600', 'text-blue-600', 'bg-blue-50');
        b.classList.add('border-transparent', 'text-slate-600', 'hover:text-slate-900', 'hover:bg-slate-50');
      });

      btn.classList.add('active', 'border-blue-600', 'text-blue-600', 'bg-blue-50');
      btn.classList.remove('border-transparent', 'text-slate-600');

      panels.forEach(p => {
        if (p.id === targetId) {
          p.classList.remove('hidden');
        } else {
          p.classList.add('hidden');
        }
      });

      if (targetId === 'tab-faculty') {
        renderFacultyTable();
      }

      lucide.createIcons();
    });
  });
}

// -------------------------------------------------------------
// 3. Course Allocation Report UI
// -------------------------------------------------------------
function initCourseAllocationUI() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('rawFileInput');
  const generateBtn = document.getElementById('generateReportBtn');
  const previewBtn = document.getElementById('previewReportBtn');

  // Drag & drop events
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.add('border-blue-500', 'bg-blue-50/60');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.remove('border-blue-500', 'bg-blue-50/60');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileSelected(files[0]);
    }
  });

  dropzone.addEventListener('click', () => {
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  generateBtn.addEventListener('click', handleGenerateReport);
  if (previewBtn) {
    previewBtn.addEventListener('click', handleTogglePreview);
  }
}

async function handleFileSelected(file) {
  if (!file.name.match(/\\.(xlsx|xls)$/i)) {
    alert('Please select a valid Excel file (.xlsx or .xls)');
    return;
  }

  currentRawFile = file;
  document.getElementById('fileNameLabel').textContent = file.name;
  document.getElementById('fileSizeLabel').textContent = (file.size / 1024).toFixed(1) + ' KB';
  document.getElementById('fileStatusCard').classList.remove('hidden');
  document.getElementById('metaDetectionCard').classList.remove('hidden');

  try {
    const arrayBuffer = await file.arrayBuffer();
    currentRawBuffer = arrayBuffer;

    // Fast parse metadata
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    const ws = wb.worksheets[0];

    const meta = reportEngine.extractMetadata(ws);
    detectedMetadata = meta;

    document.getElementById('inputCollegeName').value = meta.college_name;
    document.getElementById('inputSemesterTitle').value = meta.semester_title;
    document.getElementById('inputAdmittedBatch').value = meta.admitted_batch;

    document.getElementById('badgeTerm').textContent = meta.term || 'Detected';
    document.getElementById('badgeYear').textContent = meta.year || 'Detected';
    document.getElementById('badgeStatus').textContent = 'Ready to Generate';

    document.getElementById('generateReportBtn').disabled = false;
    document.getElementById('generateReportBtn').classList.remove('opacity-50', 'cursor-not-allowed');

    const previewBtn = document.getElementById('previewReportBtn');
    if (previewBtn) {
      previewBtn.disabled = false;
      previewBtn.classList.remove('opacity-50', 'cursor-not-allowed');
    }

  } catch (err) {
    console.error('Error reading raw file:', err);
    alert('Failed to read Excel file. Please ensure it is a valid format.');
  }
}

async function handleGenerateReport() {
  if (!currentRawBuffer) {
    alert('Please upload a raw Course Allocation Excel file first.');
    return;
  }

  const generateBtn = document.getElementById('generateReportBtn');
  const originalText = generateBtn.innerHTML;
  generateBtn.disabled = true;
  generateBtn.innerHTML = `<svg class="animate-spin -ml-1 mr-2 h-5 w-5 text-white inline" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path></svg> Generating Excel...`;

  try {
    const overrides = {
      college_name: document.getElementById('inputCollegeName').value.trim(),
      semester_title: document.getElementById('inputSemesterTitle').value.trim(),
      admitted_batch: document.getElementById('inputAdmittedBatch').value.trim()
    };

    const res = await reportEngine.generateWorkbook(currentRawBuffer, overrides);
    const buffer = await res.workbook.xlsx.writeBuffer();

    // Trigger browser download
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    // Sanitize filename based on semester title
    let outName = (overrides.semester_title || 'Course_Allocation_Report')
      .replace(/[,]/g, '')
      .replace(/\\s+/g, '_')
      .replace(/[^a-zA-Z0-9_\\-]/g, '') + '.xlsx';

    a.download = outName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    // Show success banner
    showNotification(`Success! Report generated with ${res.meta.total_departments} departments and ${res.meta.total_courses} courses. Download started!`, 'success');

  } catch (err) {
    console.error('Error generating report:', err);
    showNotification('Error generating report: ' + err.message, 'error');
  } finally {
    generateBtn.disabled = false;
    generateBtn.innerHTML = originalText;
    lucide.createIcons();
  }
}

async function handleTogglePreview() {
  const previewContainer = document.getElementById('previewContainer');
  if (!previewContainer.classList.contains('hidden')) {
    previewContainer.classList.add('hidden');
    document.getElementById('previewBtnText').textContent = 'Live Preview Report';
    return;
  }

  if (!currentRawBuffer) {
    alert('Please upload a file first.');
    return;
  }

  document.getElementById('previewBtnText').textContent = 'Loading Preview...';

  try {
    const inWb = new ExcelJS.Workbook();
    await inWb.xlsx.load(currentRawBuffer);
    const inWs = inWb.worksheets[0];

    const records = reportEngine.parseRawData(inWs);
    const deptRecordsMap = new Map();
    for (const r of records) {
      const d = r.dept;
      if (!deptRecordsMap.has(d)) deptRecordsMap.set(d, []);
      deptRecordsMap.get(d).push(r);
    }

    const orderedDepts = [];
    for (const d of DEPARTMENT_ORDER) {
      if (deptRecordsMap.has(d)) orderedDepts.push(d);
    }
    for (const d of Array.from(deptRecordsMap.keys()).sort()) {
      if (!orderedDepts.includes(d)) orderedDepts.push(d);
    }

    let html = '';
    for (const dept of orderedDepts) {
      const deptRecs = deptRecordsMap.get(dept);
      const sortedRecs = reportEngine.sortCoursesInDept(deptRecs);
      if (sortedRecs.length === 0) continue;

      html += `
        <div class="mb-6 rounded-lg border border-slate-200 overflow-hidden shadow-sm">
          <div class="bg-slate-200 px-4 py-2.5 font-bold text-slate-800 text-center tracking-wide text-base">
            ${dept} (${sortedRecs.length} courses)
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left text-xs text-slate-700">
              <thead class="bg-slate-100 uppercase text-slate-600 font-semibold border-b border-slate-200">
                <tr>
                  <th class="px-3 py-2 w-32 text-center">Course Code</th>
                  <th class="px-3 py-2">Course Name</th>
                  <th class="px-3 py-2 w-64">Name of Staff</th>
                  <th class="px-3 py-2 w-40 text-center">Phone Number</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-100">
      `;

      for (const rec of sortedRecs) {
        const phone = reportEngine.lookupPhone(rec.faculty_name);
        const phoneClass = phone ? 'text-emerald-700 font-medium' : 'text-slate-400 italic';

        html += `
          <tr class="hover:bg-slate-50/80 transition-colors">
            <td class="px-3 py-2 font-mono font-semibold text-slate-900 text-center">${rec.course_code}</td>
            <td class="px-3 py-2">${rec.course_name}</td>
            <td class="px-3 py-2 font-medium">${rec.faculty_name || '-'}</td>
            <td class="px-3 py-2 text-center ${phoneClass}">${phone || 'Not found'}</td>
          </tr>
        `;
      }

      html += `
              </tbody>
            </table>
          </div>
        </div>
      `;
    }

    document.getElementById('previewTableContent').innerHTML = html;
    previewContainer.classList.remove('hidden');
    document.getElementById('previewBtnText').textContent = 'Hide Preview';

  } catch (err) {
    console.error('Error generating preview:', err);
    alert('Failed to generate preview: ' + err.message);
    document.getElementById('previewBtnText').textContent = 'Live Preview Report';
  }
}

// -------------------------------------------------------------
// 4. Faculty Directory Management UI
// -------------------------------------------------------------
function initFacultyDirectoryUI() {
  const searchInput = document.getElementById('facultySearchInput');
  const importInput = document.getElementById('staffListImportInput');
  const importBtn = document.getElementById('importStaffBtn');
  const resetBtn = document.getElementById('resetFacultyBtn');
  const exportBtn = document.getElementById('exportFacultyBtn');

  searchInput.addEventListener('input', () => {
    renderFacultyTable(searchInput.value.trim().toLowerCase());
  });

  importBtn.addEventListener('click', () => {
    importInput.click();
  });

  importInput.addEventListener('change', async (e) => {
    if (e.target.files.length === 0) return;
    const file = e.target.files[0];
    await handleStaffListImport(file);
    e.target.value = '';
  });

  resetBtn.addEventListener('click', () => {
    if (confirm('Are you sure you want to reset the faculty directory to the default 53 staff members? Any custom edits will be reverted.')) {
      facultyList = [...DEFAULT_FACULTY_DATA];
      saveFacultyData();
      renderFacultyTable();
      showNotification('Faculty directory reset to default.', 'info');
    }
  });

  exportBtn.addEventListener('click', handleExportFaculty);
}

function renderFacultyTable(searchTerm = '') {
  const tbody = document.getElementById('facultyTableBody');
  if (!tbody) return;

  const filtered = facultyList.filter(f => {
    if (!searchTerm) return true;
    return (f.name && f.name.toLowerCase().includes(searchTerm)) ||
           (f.phone && f.phone.includes(searchTerm)) ||
           (f.department && f.department.toLowerCase().includes(searchTerm));
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="px-4 py-8 text-center text-slate-500">No matching faculty members found.</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map((f, idx) => `
    <tr class="hover:bg-slate-50/80 transition-colors border-b border-slate-100">
      <td class="px-4 py-3 text-slate-500 text-center text-xs">${f.sl_no || (idx + 1)}</td>
      <td class="px-4 py-3 font-medium text-slate-900">${f.name}</td>
      <td class="px-4 py-3">
        <span class="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-mono font-medium ${f.phone ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-slate-100 text-slate-500'}">
          ${f.phone || 'None'}
        </span>
      </td>
      <td class="px-4 py-3 text-right">
        <button onclick="editFacultyPhone('${f.name}')" class="text-blue-600 hover:text-blue-800 text-xs font-medium px-2.5 py-1 rounded bg-blue-50 hover:bg-blue-100 transition-colors">
          Edit Phone
        </button>
      </td>
    </tr>
  `).join('');
}

window.editFacultyPhone = function(name) {
  const staff = facultyList.find(f => f.name === name);
  if (!staff) return;

  const newPhone = prompt(`Update phone number for: ${staff.name}`, staff.phone || '');
  if (newPhone !== null) {
    staff.phone = newPhone.trim();
    saveFacultyData();
    renderFacultyTable(document.getElementById('facultySearchInput').value.trim().toLowerCase());
    showNotification(`Phone updated for ${staff.name}`, 'success');
  }
};

async function handleStaffListImport(file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    const ws = wb.worksheets[0];

    let importedCount = 0;
    const existingMap = new Map();
    for (const f of facultyList) {
      existingMap.set(reportEngine.normalizeName(f.name), f);
    }

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const slNo = cleanText(row.getCell(1).value);
      const name = cleanText(row.getCell(2).value);
      const phone = cleanText(row.getCell(3).value);

      if (!name) continue;

      const norm = reportEngine.normalizeName(name);
      if (existingMap.has(norm)) {
        const existing = existingMap.get(norm);
        if (phone && phone !== existing.phone) {
          existing.phone = phone;
          importedCount++;
        }
      } else {
        const newStaff = { sl_no: slNo || String(facultyList.length + 1), name, phone };
        facultyList.push(newStaff);
        existingMap.set(norm, newStaff);
        importedCount++;
      }
    }

    saveFacultyData();
    renderFacultyTable();
    showNotification(`Successfully updated/imported ${importedCount} faculty contacts!`, 'success');

  } catch (err) {
    console.error('Error importing staff list:', err);
    showNotification('Failed to parse Staff List Excel: ' + err.message, 'error');
  }
}

async function handleExportFaculty() {
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Faculty Directory');

    ws.columns = [
      { header: 'Sl No', key: 'sl_no', width: 10 },
      { header: 'Staff Name', key: 'name', width: 35 },
      { header: 'Phone Number', key: 'phone', width: 22 }
    ];

    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E7FF' }
    };

    facultyList.forEach((f, idx) => {
      ws.addRow({
        sl_no: f.sl_no || (idx + 1),
        name: f.name,
        phone: f.phone
      });
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Faculty_Directory_St_Anns.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
    showNotification('Exported faculty directory to Excel.', 'success');
  } catch (err) {
    console.error('Export failed:', err);
    showNotification('Export failed: ' + err.message, 'error');
  }
}

// -------------------------------------------------------------
// 5. Toast Notifications
// -------------------------------------------------------------
function showNotification(msg, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  const bgClass = type === 'success' ? 'bg-emerald-600 text-white' :
                  type === 'error' ? 'bg-rose-600 text-white' :
                  'bg-slate-800 text-white';

  toast.className = `px-4 py-3 rounded-lg shadow-lg flex items-center space-x-2 text-sm transition-all transform duration-300 translate-y-2 opacity-0 ${bgClass}`;
  toast.innerHTML = `<span>${msg}</span>`;

  container.appendChild(toast);
  requestAnimationFrame(() => {
    toast.classList.remove('translate-y-2', 'opacity-0');
  });

  setTimeout(() => {
    toast.classList.add('opacity-0', 'translate-y-2');
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}
"""

with open('app.js', 'w', encoding='utf-8') as f:
    f.write(APP_JS)
print('app.js written successfully.')

INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>St. Ann's College - Report Generator Hub</title>
  <!-- Tailwind CSS via CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Lucide Icons -->
  <script src="./lucide.min.js"></script>
  <!-- ExcelJS -->
  <script src="./exceljs.min.js"></script>
  <!-- Embedded Faculty Dataset (Default 53 entries) -->
  <script src="./faculty_data.js"></script>
  <!-- Report Engine & App Script -->
  <script src="./course_report_engine.js"></script>
  <script src="./app.js"></script>
  <style>
    body { font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; }
  </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen flex flex-col">

  <!-- Top Navigation Header -->
  <header class="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-xs">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex justify-between items-center h-16">
        <div class="flex items-center space-x-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-700 to-indigo-600 flex items-center justify-center text-white font-bold shadow-md shadow-blue-500/20">
            <i data-lucide="graduation-cap" class="w-6 h-6"></i>
          </div>
          <div>
            <h1 class="text-lg font-bold text-slate-900 leading-tight">St. Ann's College for Women (A)</h1>
            <p class="text-xs text-slate-500 font-medium">Autonomous College Report Generator Hub</p>
          </div>
        </div>
        <div class="flex items-center space-x-2">
          <span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
            <span class="w-2 h-2 rounded-full bg-emerald-500 mr-1.5 animate-pulse"></span>
            Cloudflare Pages Ready
          </span>
          <span class="text-xs text-slate-400 hidden sm:inline">• 100% Client-Side & Private</span>
        </div>
      </div>
    </div>
  </header>

  <!-- Report Selection Sub-navigation -->
  <div class="bg-white border-b border-slate-200">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <nav class="flex space-x-2 sm:space-x-4 py-2 overflow-x-auto" aria-label="Tabs">
        <button data-tab="tab-course-allocation" class="nav-tab-btn active px-3.5 py-2 text-sm font-semibold rounded-lg border-b-2 border-blue-600 text-blue-600 bg-blue-50 flex items-center space-x-2 whitespace-nowrap transition-all">
          <i data-lucide="file-spreadsheet" class="w-4 h-4"></i>
          <span>Course Allocation Report</span>
          <span class="ml-1.5 px-1.5 py-0.5 text-[10px] font-bold bg-blue-600 text-white rounded-full">ACTIVE</span>
        </button>
        <button data-tab="tab-faculty" class="nav-tab-btn px-3.5 py-2 text-sm font-medium rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-50 flex items-center space-x-2 whitespace-nowrap transition-all">
          <i data-lucide="users" class="w-4 h-4"></i>
          <span>Faculty Directory</span>
          <span id="facultyCountBadge" class="ml-1.5 px-1.5 py-0.5 text-[10px] font-semibold bg-slate-100 text-slate-600 rounded-full">53 Loaded</span>
        </button>
        <button data-tab="tab-workload" class="nav-tab-btn px-3.5 py-2 text-sm font-medium rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-50 flex items-center space-x-2 whitespace-nowrap transition-all">
          <i data-lucide="bar-chart-3" class="w-4 h-4"></i>
          <span>Faculty Workload</span>
          <span class="ml-1.5 px-1.5 py-0.5 text-[10px] font-medium bg-slate-100 text-slate-500 rounded">Upcoming</span>
        </button>
        <button data-tab="tab-timetable" class="nav-tab-btn px-3.5 py-2 text-sm font-medium rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-50 flex items-center space-x-2 whitespace-nowrap transition-all">
          <i data-lucide="calendar" class="w-4 h-4"></i>
          <span>Timetable Matrix</span>
          <span class="ml-1.5 px-1.5 py-0.5 text-[10px] font-medium bg-slate-100 text-slate-500 rounded">Upcoming</span>
        </button>
      </nav>
    </div>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 flex-1 w-full">

    <!-- ============================================================= -->
    <!-- TAB 1: Course Allocation Report (Active) -->
    <!-- ============================================================= -->
    <div id="tab-course-allocation" class="tab-panel space-y-6">
      <!-- Welcome & Instructions Card -->
      <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white rounded-2xl p-6 shadow-xl relative overflow-hidden">
        <div class="absolute -right-8 -bottom-8 opacity-10 pointer-events-none">
          <i data-lucide="table" class="w-64 h-64"></i>
        </div>
        <div class="relative z-10 max-w-3xl">
          <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/30 text-blue-200 border border-blue-400/30 mb-2">
            Report 1 • Course List with Faculty Details
          </span>
          <h2 class="text-2xl font-bold text-white tracking-tight">Convert Raw SIS Allocation Excel into College Standard Report</h2>
          <p class="mt-2 text-slate-300 text-sm leading-relaxed">
            Drop your raw course allocation spreadsheet below. The engine automatically extracts the Semester and Admitted Batch, segregates General Languages into their own top section, applies the 1-to-10 course code hierarchy with Theory/Practical pairing, and fetches official faculty contact numbers.
          </p>
        </div>
      </div>

      <!-- File Upload & Configuration Grid -->
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">

        <!-- Upload Drag & Drop Zone (7 cols) -->
        <div class="lg:col-span-7 bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between">
          <div>
            <h3 class="text-base font-semibold text-slate-900 mb-1 flex items-center">
              <i data-lucide="upload-cloud" class="w-5 h-5 text-blue-600 mr-2"></i>
              Step 1: Upload Raw Course Allocation Export
            </h3>
            <p class="text-xs text-slate-500 mb-4">Export from SIS/ERP containing columns: Course Code, Course Name, Community, Faculty, Batch/Section Name</p>

            <div id="dropzone" class="border-2 border-dashed border-slate-300 hover:border-blue-500 bg-slate-50/70 hover:bg-blue-50/40 rounded-xl p-8 text-center cursor-pointer transition-all flex flex-col items-center justify-center">
              <input type="file" id="rawFileInput" accept=".xlsx, .xls" class="hidden">
              <div class="w-14 h-14 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center mb-3 shadow-inner">
                <i data-lucide="file-up" class="w-7 h-7"></i>
              </div>
              <p class="text-sm font-semibold text-slate-800">Click to browse or drag and drop raw Excel file here</p>
              <p class="text-xs text-slate-500 mt-1">Supports .xlsx or .xls (Course_Allocation_Report_*.xlsx)</p>
            </div>
          </div>

          <!-- File Uploaded Info Card -->
          <div id="fileStatusCard" class="hidden mt-4 bg-slate-50 border border-slate-200 rounded-xl p-3.5 flex items-center justify-between">
            <div class="flex items-center space-x-3">
              <div class="w-9 h-9 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center">
                <i data-lucide="file-check-2" class="w-5 h-5"></i>
              </div>
              <div>
                <p id="fileNameLabel" class="text-sm font-bold text-slate-800 truncate max-w-xs">filename.xlsx</p>
                <p id="fileSizeLabel" class="text-xs text-slate-500">0 KB</p>
              </div>
            </div>
            <span id="badgeStatus" class="px-2.5 py-1 text-xs font-semibold rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">
              Loaded
            </span>
          </div>
        </div>

        <!-- Detected Headers & Overrides (5 cols) -->
        <div class="lg:col-span-5 bg-white rounded-2xl p-6 border border-slate-200 shadow-sm flex flex-col justify-between">
          <div>
            <h3 class="text-base font-semibold text-slate-900 mb-1 flex items-center">
              <i data-lucide="sliders" class="w-5 h-5 text-indigo-600 mr-2"></i>
              Step 2: Auto-Detected Report Headers
            </h3>
            <p class="text-xs text-slate-500 mb-4">Values are automatically parsed from row 1 metadata. You can customize them before generation if needed.</p>

            <div class="space-y-3.5">
              <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">College Header Title</label>
                <input type="text" id="inputCollegeName" value="St.Ann's College for Women (A)" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none">
              </div>

              <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Semester Title</label>
                <input type="text" id="inputSemesterTitle" placeholder="e.g. 3rd Semester Course Code, Course name with faculty Details" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none">
              </div>

              <div>
                <label class="block text-xs font-semibold text-slate-700 mb-1">Admitted Batch Text</label>
                <input type="text" id="inputAdmittedBatch" placeholder="e.g. Admitted Batch 2025-2026" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none">
              </div>

              <div id="metaDetectionCard" class="hidden pt-2 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600">
                <span>Detected Term: <strong id="badgeTerm" class="text-blue-700 font-mono">-</strong></span>
                <span>Batch Year: <strong id="badgeYear" class="text-blue-700 font-mono">-</strong></span>
              </div>
            </div>
          </div>

          <!-- Action Buttons -->
          <div class="pt-6 mt-6 border-t border-slate-100 flex flex-col sm:flex-row gap-3">
            <button id="generateReportBtn" disabled class="flex-1 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-semibold text-sm rounded-xl shadow-md shadow-blue-500/20 transition-all flex items-center justify-center space-x-2 opacity-50 cursor-not-allowed">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span>Generate & Download Excel</span>
            </button>
            <button id="previewReportBtn" disabled class="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium text-sm rounded-xl transition-all flex items-center justify-center space-x-1.5 opacity-50 cursor-not-allowed">
              <i data-lucide="eye" class="w-4 h-4"></i>
              <span id="previewBtnText">Live Preview</span>
            </button>
          </div>
        </div>

      </div>

      <!-- Live Preview Section (Toggled by user) -->
      <div id="previewContainer" class="hidden bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
        <div class="flex items-center justify-between mb-4 pb-3 border-b border-slate-200">
          <div>
            <h3 class="text-lg font-bold text-slate-900">Report Preview: Sheet 1 (Department-wise Course List)</h3>
            <p class="text-xs text-slate-500">Live preview of departments, courses, theory/practical pairing, and phone numbers.</p>
          </div>
          <span class="text-xs bg-blue-50 text-blue-700 px-3 py-1 rounded-full border border-blue-200 font-medium">
            Sheet 1 & Sheet 2 included in download
          </span>
        </div>
        <div id="previewTableContent" class="space-y-4 max-h-[600px] overflow-y-auto pr-1"></div>
      </div>

    </div>

    <!-- ============================================================= -->
    <!-- TAB 2: Faculty Directory Management -->
    <!-- ============================================================= -->
    <div id="tab-faculty" class="tab-panel hidden space-y-6">
      <div class="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-6 border-b border-slate-200">
          <div>
            <h2 class="text-xl font-bold text-slate-900 flex items-center">
              <i data-lucide="contact" class="w-6 h-6 text-blue-600 mr-2"></i>
              Faculty Contact Directory
            </h2>
            <p class="text-xs text-slate-500 mt-1">Official phone numbers used to automatically populate reports. Custom phone edits are saved directly in your browser.</p>
          </div>
          <div class="flex flex-wrap items-center gap-2.5">
            <input type="file" id="staffListImportInput" accept=".xlsx, .xls" class="hidden">
            <button id="importStaffBtn" class="px-3.5 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 font-semibold text-xs rounded-xl transition-all flex items-center space-x-1.5">
              <i data-lucide="upload" class="w-4 h-4"></i>
              <span>Import Staff List Report (.xlsx)</span>
            </button>
            <button id="exportFacultyBtn" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium text-xs rounded-xl transition-all flex items-center space-x-1.5">
              <i data-lucide="download" class="w-4 h-4"></i>
              <span>Export to Excel</span>
            </button>
            <button id="resetFacultyBtn" class="px-3 py-2 text-rose-600 hover:text-rose-700 hover:bg-rose-50 text-xs font-medium rounded-xl transition-all">
              Reset to Defaults
            </button>
          </div>
        </div>

        <!-- Search Bar -->
        <div class="py-4">
          <div class="relative max-w-md">
            <i data-lucide="search" class="w-4 h-4 absolute left-3.5 top-3 text-slate-400"></i>
            <input type="text" id="facultySearchInput" placeholder="Search staff name or phone number..." class="w-full pl-10 pr-4 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
          </div>
        </div>

        <!-- Faculty Table -->
        <div class="overflow-x-auto rounded-xl border border-slate-200">
          <table class="w-full text-left text-sm text-slate-700">
            <thead class="bg-slate-100 text-xs uppercase text-slate-600 font-semibold border-b border-slate-200">
              <tr>
                <th class="px-4 py-3 w-16 text-center">#</th>
                <th class="px-4 py-3">Faculty Name</th>
                <th class="px-4 py-3 w-48">Phone Number</th>
                <th class="px-4 py-3 w-32 text-right">Action</th>
              </tr>
            </thead>
            <tbody id="facultyTableBody" class="divide-y divide-slate-100">
              <!-- Dynamically populated -->
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ============================================================= -->
    <!-- TAB 3: Faculty Workload (Modular Slot for Future Report) -->
    <!-- ============================================================= -->
    <div id="tab-workload" class="tab-panel hidden space-y-6">
      <div class="bg-white rounded-2xl p-8 border border-slate-200 shadow-sm text-center max-w-2xl mx-auto">
        <div class="w-16 h-16 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center mx-auto mb-4">
          <i data-lucide="bar-chart-2" class="w-8 h-8"></i>
        </div>
        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200 mb-2">
          Modular Extension Slot
        </span>
        <h2 class="text-2xl font-bold text-slate-900">Faculty Workload Analysis Report</h2>
        <p class="text-sm text-slate-500 mt-2">
          This tab is reserved for your upcoming <strong>Faculty Workload & Hours Distribution Report</strong>.
          The modular architecture of this web portal allows adding this report seamlessly once the raw workload format or requirements are provided.
        </p>
        <div class="mt-6 pt-6 border-t border-slate-100 flex justify-center gap-4 text-xs text-slate-400">
          <span>• Theory / Practical Hour Summary</span>
          <span>• Departmental Load Metrics</span>
          <span>• Individual Staff Allotments</span>
        </div>
      </div>
    </div>

    <!-- ============================================================= -->
    <!-- TAB 4: Timetable Matrix (Modular Slot for Future Report) -->
    <!-- ============================================================= -->
    <div id="tab-timetable" class="tab-panel hidden space-y-6">
      <div class="bg-white rounded-2xl p-8 border border-slate-200 shadow-sm text-center max-w-2xl mx-auto">
        <div class="w-16 h-16 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center mx-auto mb-4">
          <i data-lucide="calendar-range" class="w-8 h-8"></i>
        </div>
        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200 mb-2">
          Modular Extension Slot
        </span>
        <h2 class="text-2xl font-bold text-slate-900">Batch-wise Timetable Matrix Report</h2>
        <p class="text-sm text-slate-500 mt-2">
          This tab is reserved for your upcoming <strong>Section & Batch Timetable Matrix</strong> report.
          Ready to plug in when timetable source formats are defined.
        </p>
      </div>
    </div>

  </main>

  <!-- Footer -->
  <footer class="bg-white border-t border-slate-200 py-4 mt-auto">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row justify-between items-center text-xs text-slate-500 gap-2">
      <p>© St. Ann's College for Women (A) • Autonomous College Reports</p>
      <p>Deployable on Cloudflare Pages • Vercel • GitHub Pages</p>
    </div>
  </footer>

  <!-- Toast Notification Container -->
  <div id="toastContainer" class="fixed bottom-5 right-5 z-50 flex flex-col space-y-2 pointer-events-none"></div>

</body>
</html>
"""

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(INDEX_HTML)
print('index.html written successfully.')

# Wrangler config for Cloudflare Pages
WRANGLER_TOML = """name = "st-anns-college-reports"
compatibility_date = "2024-01-01"
pages_build_output_dir = "."
"""
with open('wrangler.toml', 'w', encoding='utf-8') as f:
    f.write(WRANGLER_TOML)

# Cloudflare Pages Headers for caching
HEADERS = """/*
  Cache-Control: public, max-age=3600
/exceljs.min.js
  Cache-Control: public, max-age=31536000, immutable
/lucide.min.js
  Cache-Control: public, max-age=31536000, immutable
"""
with open('_headers', 'w', encoding='utf-8') as f:
    f.write(HEADERS)

# Run_Web_App.bat launcher
BAT_CONTENT = """@echo off
echo ========================================================
echo  St. Ann's College for Women (A) - Report Generator Web
echo ========================================================
echo Starting local web portal at http://localhost:8000 ...
start "" "http://localhost:8000"
python -m http.server 8000
pause
"""
with open('Run_Web_App.bat', 'w', encoding='utf-8') as f:
    f.write(BAT_CONTENT)
README_CONTENT = """# Deploying St. Ann's College Report Generator to Cloudflare Pages

This web application is **100% client-side** (built with pure HTML5, Tailwind CSS, and ExcelJS).
- Zero server maintenance, zero compute fees (100% free forever on Cloudflare Pages).
- Instant report generation (<0.5 seconds).
- 100% private: Student batches and faculty contacts never leave the user's browser.

---

## Method 1: 1-Click Drag-and-Drop (Easiest - No Git Required)

1. Log into your [Cloudflare Dashboard](https://dash.cloudflare.com/).
2. On the left sidebar, click **Workers & Pages**.
3. Click **Create application** -> select the **Pages** tab.
4. Click **Upload assets**.
5. Give your project a name (e.g. `st-anns-reports`).
6. Select or drag-and-drop the files from this folder:
   - `index.html`
   - `course_report_engine.js`
   - `app.js`
   - `faculty_data.js`
   - `exceljs.min.js`
   - `lucide.min.js`
   - `wrangler.toml`
   - `_headers`
7. Click **Deploy site**.
8. In 5 seconds, your app will be live at:
   `https://st-anns-reports.pages.dev`

---

## Method 2: Connect via GitHub (Automatic Deployments on Commit)

1. Push this folder to a GitHub repository.
2. In Cloudflare Dashboard -> **Workers & Pages** -> **Create application** -> **Pages**.
3. Click **Connect to Git** and select your repository.
4. Set the build settings:
   - **Framework preset**: None
   - **Build command**: (Leave blank)
   - **Build output directory**: `.`
5. Click **Save and Deploy**.

---

## Running Locally on Windows Anytime

Just double-click **`Run_Web_App.bat`** in this folder. It opens your browser immediately at `http://localhost:8000`.
"""
with open('README_CLOUDFLARE.md', 'w', encoding='utf-8') as f:
    f.write(README_CONTENT)

print('All web portal files created successfully.')


