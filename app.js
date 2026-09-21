/**
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
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
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
      .replace(/\s+/g, '_')
      .replace(/[^a-zA-Z0-9_\-]/g, '') + '.xlsx';

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
