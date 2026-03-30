const state = { excelFiles: [], datFiles: [], shiftMap: {}, records: [], filtered: [], sortCol: 'date', sortDir: 1, page: 1, perPage: 50 };

// INITIALIZE ON LOAD
window.onload = () => {
    const isDark = localStorage.getItem('theme') === 'dark';
    if (isDark) {
        document.documentElement.setAttribute('data-theme', 'dark');
        document.getElementById('theme-btn').textContent = '☀️';
    }

    // AUTO-LOAD PERSISTED DATA
    const savedData = localStorage.getItem('hr_report_records');
    if (savedData) {
        try {
            state.records = JSON.parse(savedData);
            document.getElementById('steps-container').style.display = 'none'; // Hide upload
            document.getElementById('btn-gen').style.display = 'none'; // Hide generate button

            finishUIBuild();
        } catch (e) {
            console.error("Failed to load saved data", e);
        }
    }
};

function toggleTheme() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    document.documentElement.setAttribute('data-theme', isDark ? 'light' : 'dark');
    document.getElementById('theme-btn').textContent = isDark ? '🌙' : '☀️';
    localStorage.setItem('theme', isDark ? 'light' : 'dark');
}

function clearData() {
    if (confirm("Are you sure you want to clear the data? You will need to re-upload your files.")) {
        localStorage.removeItem('hr_report_records');
        location.reload();
    }
}

function handleFiles(files, type) {
    const arr = Array.from(files);
    if (type === 'excel') state.excelFiles.push(...arr);
    else state.datFiles.push(...arr);
    renderPills(type); checkReady();
}
function dzDrag(e, id) { e.preventDefault(); document.getElementById(id).classList.add('drag-over'); }
function dzLeave(id) { document.getElementById(id).classList.remove('drag-over'); }
function dzDrop(e, type) {
    e.preventDefault(); document.getElementById(type === 'excel' ? 'dz-excel' : 'dz-dat').classList.remove('drag-over');
    handleFiles(e.dataTransfer.files, type);
}
function renderPills(type) {
    const arr = type === 'excel' ? state.excelFiles : state.datFiles;
    document.getElementById('fl-' + type).innerHTML = arr.map((f, i) => `<div class="file-pill">📄 ${f.name} <button class="remove" onclick="removeFile('${type}',${i})">×</button></div>`).join('');
}
function removeFile(type, idx) { type === 'excel' ? state.excelFiles.splice(idx, 1) : state.datFiles.splice(idx, 1); renderPills(type); checkReady(); }
function checkReady() { document.getElementById('btn-gen').disabled = !(state.excelFiles.length && state.datFiles.length); }
function setProgress(pct, msg) { document.getElementById('progress-area').style.display = 'block'; document.getElementById('progress-bar').style.width = pct + '%'; document.getElementById('progress-text').textContent = msg; }
function showError(msg) { const el = document.getElementById('error-box'); el.textContent = '⚠ ' + msg; el.style.display = 'block'; setTimeout(() => el.style.display = 'none', 6000); }

async function parseShiftMaster(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const wb = XLSX.read(e.target.result, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
                const map = {}; let headerRow = 0;

                for (let i = 0; i < Math.min(5, rows.length); i++) {
                    const r = rows[i].map(c => String(c || '').toLowerCase());
                    if (r.some(c => c.includes('userid') || c.includes('user id') || c.includes('emp code'))) { headerRow = i; break; }
                }

                const hdr = rows[headerRow].map(c => String(c || '').toLowerCase().trim());
                const col = k => hdr.findIndex(h => h.includes(k));

                const idCol = col('userid') !== -1 ? col('userid') : (col('user') !== -1 ? col('user') : col('emp code'));
                const nameCol = col('particular') !== -1 ? col('particular') : col('name');
                const branchCol = col('branch');
                const deptCol = col('department') !== -1 ? col('department') : col('dept');
                const startCol = col('shift start') !== -1 ? col('shift start') : col('start');
                const endCol = col('shift end') !== -1 ? col('shift end') : col('end');

                for (let i = headerRow + 1; i < rows.length; i++) {
                    const r = rows[i]; const uid = r[idCol];
                    if (!uid && !r[nameCol]) continue;

                    const toTime = v => {
                        if (!v) return null;
                        if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
                        if (typeof v === 'number') return Math.round(v * 24 * 60);
                        const m = String(v).match(/(\d+):(\d+)/);
                        return m ? parseInt(m[1]) * 60 + parseInt(m[2]) : null;
                    };

                    if (uid) {
                        map[String(uid).trim()] = {
                            name: r[nameCol] ? String(r[nameCol]).trim() : 'Unknown',
                            branch: r[branchCol] ? String(r[branchCol]).trim() : '',
                            department: r[deptCol] ? String(r[deptCol]).trim() : '',
                            shiftStart: toTime(r[startCol]), shiftEnd: toTime(r[endCol])
                        };
                    }
                }
                resolve(map);
            } catch (err) { reject(err); }
        };
        reader.readAsArrayBuffer(file);
    });
}

async function parseDatFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            const lines = e.target.result.split(/\r?\n/); const punches = [];
            for (const line of lines) {
                if (!line.trim()) continue;
                const parts = line.trim().split(/\s+/); if (parts.length < 2) continue;
                const dateStr = (parts[1] && parts[2] && !parts[1].includes(' ')) ? parts[1] + ' ' + parts[2] : parts[1];
                const d = new Date(dateStr);
                if (!isNaN(d.getTime())) punches.push({ uid: String(parts[0]).trim(), datetime: d });
            }
            resolve(punches);
        };
        reader.readAsText(file);
    });
}

async function generateReport() {
    document.getElementById('error-box').style.display = 'none';
    document.getElementById('btn-gen').disabled = true;

    try {
        setProgress(10, 'Reading shift master...');
        state.shiftMap = {};
        for (const f of state.excelFiles) Object.assign(state.shiftMap, await parseShiftMaster(f));

        setProgress(30, 'Parsing attendance logs...');
        let allPunches = [];
        for (const f of state.datFiles) allPunches.push(...await parseDatFile(f));
        if (!allPunches.length) throw new Error('No valid punch records found in .dat files.');

        setProgress(55, 'Processing data...');
        const grouped = {};
        for (const p of allPunches) {
            const dKey = p.datetime.toISOString().split('T')[0];
            const key = p.uid + '|' + dKey;
            if (!grouped[key]) grouped[key] = { uid: p.uid, date: dKey, punches: [] };
            grouped[key].punches.push(p.datetime);
        }

        state.records = [];
        const fmtHHMM = mins => {
            if (mins === null || isNaN(mins)) return '';
            return Math.floor(mins / 60).toString().padStart(2, '0') + ':' + (mins % 60).toString().padStart(2, '0');
        };

        for (const key of Object.keys(grouped)) {
            const g = grouped[key];
            const info = state.shiftMap[g.uid] || { name: 'User ' + g.uid, branch: '', department: '', shiftStart: null, shiftEnd: null };
            const punches = g.punches.sort((a, b) => a - b);
            const firstIn = punches[0], lastOut = punches[punches.length - 1];
            const hoursWorked = ((lastOut - firstIn) / 3600000);

            const inMins = firstIn.getHours() * 60 + firstIn.getMinutes();
            const outMins = lastOut.getHours() * 60 + lastOut.getMinutes();

            let lateBy = 0, earlyBy = 0, status = 'Present';
            if (info.shiftStart !== null) {
                lateBy = Math.max(0, inMins - info.shiftStart);
                if (lateBy > 15) status = 'Late';
            }
            if (info.shiftEnd !== null) earlyBy = Math.max(0, info.shiftEnd - outMins);
            if (hoursWorked < 4.5) status = 'Half Day';
            if (hoursWorked < 1) status = 'Absent';

            const fmtMins = m => (m <= 0) ? '—' : Math.floor(m / 60) + 'h ' + (m % 60) + 'm';
            const shiftDisplay = (info.shiftStart !== null && info.shiftEnd !== null) ? `${fmtHHMM(info.shiftStart)} - ${fmtHHMM(info.shiftEnd)}` : '—';

            // GET DAY OF WEEK
            const dayName = new Date(g.date).toLocaleDateString('en-US', { weekday: 'short' });

            state.records.push({
                uid: g.uid, name: info.name, branch: info.branch, department: info.department,
                date: g.date, day: dayName, shiftDisplay, firstIn: fmtHHMM(inMins), lastOut: fmtHHMM(outMins),
                hoursWorked: Math.round(hoursWorked * 100) / 100, status,
                lateBy: lateBy > 0 ? fmtMins(lateBy) : '—', earlyBy: earlyBy > 0 ? fmtMins(earlyBy) : '—'
            });
        }

        setProgress(90, 'Saving Session...');

        // SAVE TO PERSISTENCE (LOCAL STORAGE)
        try {
            localStorage.setItem('hr_report_records', JSON.stringify(state.records));
        } catch (e) {
            console.warn("Storage quota exceeded, session not saved.");
        }

        setProgress(100, `Done! ${state.records.length} records processed.`);

        // HIDE FILE UPLOADS ON SUCCESS
        document.getElementById('steps-container').style.display = 'none';
        document.getElementById('btn-gen').style.display = 'none';
        document.getElementById('progress-area').style.display = 'none';

        finishUIBuild();

    } catch (err) { showError(err.message || 'Failed to process files.'); console.error(err); document.getElementById('btn-gen').disabled = false; }
}

// Shared function for both Generate and Auto-Load
function finishUIBuild() {
    const branches = [...new Set(state.records.map(r => r.branch).filter(Boolean))].sort();
    const depts = [...new Set(state.records.map(r => r.department).filter(Boolean))].sort();
    document.getElementById('f-branch').innerHTML = '<option value="">All Branches</option>' + branches.map(v => `<option value="${v}">${v}</option>`).join('');
    document.getElementById('f-dept').innerHTML = '<option value="">All Departments</option>' + depts.map(v => `<option value="${v}">${v}</option>`).join('');

    state.filtered = [...state.records];
    state.page = 1;

    const statsRow = document.getElementById('stats-row');
    const tableWrap = document.getElementById('table-wrap');

    statsRow.style.display = 'grid';
    tableWrap.style.display = 'block';
    document.getElementById('btn-clear').style.display = 'flex';
    document.getElementById('dl-wrap').style.display = 'flex';

    // Trigger Animations
    statsRow.classList.add('fade-in');
    tableWrap.classList.add('fade-in', 'stagger-1');

    applyFilters();
}

function applyFilters() {
    const search = document.getElementById('search').value.toLowerCase();
    const branch = document.getElementById('f-branch').value;
    const dept = document.getElementById('f-dept').value;
    const status = document.getElementById('f-status').value;

    // STRICT DATE PARSING FOR UI & EXPORT
    const fromVal = document.getElementById('date-from').value;
    const toVal = document.getElementById('date-to').value;
    let fromTime = -Infinity, toTime = Infinity;

    if (fromVal) { const d = new Date(fromVal); if (!isNaN(d.getTime())) fromTime = d.setHours(0, 0, 0, 0); }
    if (toVal) { const d = new Date(toVal); if (!isNaN(d.getTime())) toTime = d.setHours(23, 59, 59, 999); }

    // GATING THE DOWNLOAD OPTION
    const dlBtn = document.getElementById('btn-dl');
    const dlMsg = document.getElementById('dl-msg');
    if (fromVal && toVal) {
        dlBtn.style.display = 'flex';
        dlMsg.style.display = 'none';
    } else {
        dlBtn.style.display = 'none';
        dlMsg.style.display = 'block';
    }

    state.filtered = state.records.filter(r => {
        if (search && !(r.name.toLowerCase().includes(search) || r.uid.includes(search))) return false;
        if (branch && r.branch !== branch) return false;
        if (dept && r.department !== dept) return false;
        if (status && r.status !== status) return false;
        const rTime = new Date(r.date).getTime();
        if (fromVal && rTime < fromTime) return false;
        if (toVal && rTime > toTime) return false;
        return true;
    });
    state.page = 1; renderTable(); renderStats(state.filtered);
}

function sortBy(col) {
    state.sortDir = (state.sortCol === col) ? state.sortDir * -1 : 1;
    state.sortCol = col;
    state.filtered.sort((a, b) => {
        let av = a[col] || '', bv = b[col] || '';
        if (col === 'hoursWorked') return (av - bv) * state.sortDir;
        return String(av).localeCompare(String(bv)) * state.sortDir;
    });
    renderTable();
}

function renderTable() {
    const total = state.filtered.length, start = (state.page - 1) * state.perPage;
    const rows = state.filtered.slice(start, start + state.perPage);
    const statusBadge = s => {
        const cls = { Present: 'present', Late: 'late', Absent: 'absent', 'Half Day': 'half' }[s] || 'present';
        return `<span class="badge badge-${cls}">${s}</span>`;
    };

    document.getElementById('table-body').innerHTML = rows.map(r => `
<tr>
    <td data-label="Employee">
        <div><strong>${r.name}</strong><br><span style="font-size:11px;color:var(--text3)">ID: ${r.uid}</span></div>
        <div class="mobile-status" style="display:none;">${statusBadge(r.status)}</div>
    </td>
    <td data-label="Branch">${r.branch || '—'}</td>
    <td data-label="Department">${r.department || '—'}</td>
    <td data-label="Date">${r.date}</td>
    <td data-label="Day">${r.day}</td>
    <td data-label="Shift"><span class="shift-txt">${r.shiftDisplay}</span></td>
    <td data-label="First In">${r.firstIn}</td>
    <td data-label="Last Out">${r.lastOut}</td>
    <td data-label="Hours" class="${r.hoursWorked >= 8 ? 'hours-ok' : r.hoursWorked >= 4 ? 'hours-low' : 'hours-zero'}">${r.hoursWorked}h</td>
    <td data-label="Status">${statusBadge(r.status)}</td>
    <td data-label="Late By" style="color:var(--amber);font-weight:500;">${r.lateBy}</td>
    <td data-label="Early Out" style="color:var(--red)">${r.earlyBy}</td>
</tr>
`).join('');

    if (window.innerWidth <= 768) { document.querySelectorAll('.mobile-status').forEach(el => el.style.display = 'block'); }

    document.getElementById('page-info').textContent = total === 0 ? 'No records' : `Showing ${start + 1}–${Math.min(start + state.perPage, total)} of ${total} records`;

    const pages = Math.ceil(total / state.perPage); let html = '';
    if (pages > 1) {
        html += `<button class="page-btn" onclick="goPage(${state.page - 1})" ${state.page === 1 ? 'disabled' : ''}>‹</button>`;
        for (let i = 1; i <= pages; i++) {
            if (i === 1 || i === pages || Math.abs(i - state.page) <= 1) html += `<button class="page-btn ${i === state.page ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
            else if (i === state.page - 2 || i === state.page + 2) html += `<span class="page-btn" style="cursor:default">…</span>`;
        }
        html += `<button class="page-btn" onclick="goPage(${state.page + 1})" ${state.page === pages ? 'disabled' : ''}>›</button>`;
    }
    document.getElementById('page-btns').innerHTML = html;
}

function goPage(p) { if (p >= 1 && p <= Math.ceil(state.filtered.length / state.perPage)) { state.page = p; renderTable(); } }

function renderStats(records) {
    const present = records.filter(r => r.status !== 'Absent').length;
    document.getElementById('st-emp').textContent = new Set(records.map(r => r.uid)).size;
    document.getElementById('st-days').textContent = new Set(records.map(r => r.date)).size;
    document.getElementById('st-att').textContent = records.length ? Math.round(present / records.length * 100) + '%' : '0%';
    document.getElementById('st-abs').textContent = records.length - present;
}

function downloadExcel() {
    const dataToExport = state.filtered;
    if (dataToExport.length === 0) return showError("No data available to download.");

    const wb = XLSX.utils.book_new();

    // 1. SUMMARY SHEET
    const summaryData = {};
    dataToExport.forEach(r => {
        if (!summaryData[r.uid]) summaryData[r.uid] = { 'ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Dept': r.department, 'Present': 0, 'Late': 0, 'Half Day': 0, 'Absent': 0, 'Days': 0, '_hrs': 0 };
        const s = summaryData[r.uid];
        s['Days']++; s['_hrs'] += r.hoursWorked;
        if (r.status === 'Present') s['Present']++; else if (r.status === 'Late') { s['Present']++; s['Late']++; } else if (r.status === 'Half Day') s['Half Day']++; else s['Absent']++;
    });
    const summaryRows = Object.values(summaryData).map(s => { s['Avg Hrs'] = Math.round(s['_hrs'] / s['Days'] * 100) / 100; s['Att %'] = Math.round(s['Present'] / s['Days'] * 100) + '%'; delete s['_hrs']; return s; });
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    wsSummary['!cols'] = [8, 20, 15, 15, 8, 8, 9, 8, 8, 8, 8].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    // 2. DAILY DETAIL SHEET (ADDED DAY COLUMN)
    const detailRows = dataToExport.map(r => ({ 'ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Dept': r.department, 'Date': r.date, 'Day': r.day, 'Shift': r.shiftDisplay, 'First In': r.firstIn, 'Last Out': r.lastOut, 'Hours': r.hoursWorked, 'Status': r.status, 'Late By': r.lateBy, 'Early Out': r.earlyBy }));
    const wsDetail = XLSX.utils.json_to_sheet(detailRows);
    wsDetail['!cols'] = [8, 20, 15, 15, 12, 6, 13, 10, 10, 8, 10, 10, 10].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsDetail, 'Daily Detail');

    // 3. LATE REPORT (ADDED DAY COLUMN)
    const lateData = dataToExport.filter(r => r.status === 'Late' || r.lateBy !== '—');
    if (lateData.length) {
        const wsLate = XLSX.utils.json_to_sheet(lateData.map(r => ({ 'ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Date': r.date, 'Day': r.day, 'Shift': r.shiftDisplay, 'First In': r.firstIn, 'Late By': r.lateBy })));
        wsLate['!cols'] = [8, 20, 15, 12, 6, 13, 10, 10].map(w => ({ wch: w }));
        XLSX.utils.book_append_sheet(wb, wsLate, 'Late Report');
    }

    // 4. ABSENT REPORT (ADDED DAY COLUMN)
    const absentData = dataToExport.filter(r => r.status === 'Absent');
    if (absentData.length) {
        const wsAbsent = XLSX.utils.json_to_sheet(absentData.map(r => ({ 'ID': r.uid, 'Name': r.name, 'Branch': r.branch, 'Date': r.date, 'Day': r.day, 'Status': 'Absent' })));
        wsAbsent['!cols'] = [8, 20, 15, 12, 6, 10].map(w => ({ wch: w }));
        XLSX.utils.book_append_sheet(wb, wsAbsent, 'Absent Report');
    }

    const from = document.getElementById('date-from').value;
    const to = document.getElementById('date-to').value;
    XLSX.writeFile(wb, `HR_Report_${from}_to_${to}.xlsx`);
}
