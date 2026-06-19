// â”€â”€ Global state â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    let allData = [];
    let sheetCols = [];
    let workbook = null;
    let currentFiltered = [];
    let viewMode = 'detail';

    const requiredHeaders = ["Date", "Name", "Process", "UOM", "TAT / Hour", "Target", "Count", "Working Hours", "Notes", "Leave Report"];
    const D_IDX = 0, N_IDX = 1, P_IDX = 2, TAT_IDX = 4, CNT_IDX = 6, WH_IDX = 7;

    // âœ… Summary headers exactly like your 2nd image
    const summaryHeaders = ["Name", "Date", "Working Hours", "Hours", "Productivity %", "Remarks", "Leave", "Traget Status", "Concatenate"];

    // â”€â”€ Helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const norm = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
    const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const pNum = v => { const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? null : n; };

    const monthShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const fmtDate = d => {
      if (!d) return '';
      const day = d.getDate(); // no leading zero
      const mon = monthShort[d.getMonth()];
      const yr = d.getFullYear();
      return `${day}-${mon}-${yr}`; // 2-Feb-2026
    };

    const isoDate = d => d ? d.toISOString().split('T')[0] : '';

    // Round like Excel screenshot
    const f1 = v => (v === null || v === undefined || isNaN(Number(v))) ? '--' : (Math.round(Number(v) * 10) / 10).toFixed(1);

    // âœ… Round half up (1.4->1, 1.5->2)
    const f0 = v => {
      const n = Number(v);
      if (v === null || v === undefined || isNaN(n)) return '--';
      return String(Math.floor(n + 0.5));
    };

    function parseDate(v) {
      if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
      if (typeof v === 'number') {
        const d = new Date(Math.round((v - 25569) * 86400 * 1000));
        return isNaN(d.getTime()) ? null : d;
      }
      if (!v) return null;
      const s = String(v).trim();
      let d = new Date(s);
      if (!isNaN(d.getTime()) && s.length > 5) return d;
      const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (m) {
        d = new Date(`${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`);
        if (!isNaN(d.getTime())) return d;
      }
      return null;
    }

    function matchCol(headers, aliases) {
      for (const alias of aliases) {
        const idx = headers.findIndex(h => h === alias || h.replace(/\s+/g, '') === alias.replace(/\s+/g, ''));
        if (idx >= 0) return idx;
      }
      return -1;
    }

    function uniqJoin(arr) {
      const out = [];
      const seen = new Set();
      for (const x of arr) {
        const v = String(x || '').trim();
        if (!v) continue;
        if (seen.has(v)) continue;
        seen.add(v);
        out.push(v);
      }
      return out.join(', ');
    }

    // âœ… NEW: Determine divisor for Productivity % based on Remarks / Leave keywords
    function getProductivityDivisor(remarksJoined, leaveJoined) {
      const txt = `${String(remarksJoined || '')} ${String(leaveJoined || '')}`.toLowerCase();

      if (txt.includes('half day leave') || txt.includes('session_01 leave') || txt.includes('session_02 leave')) return 4;
      if (txt.includes('2 hr per morn') || txt.includes('2 hr per eve')) return 6;
      if (txt.includes('1 hr per morn') || txt.includes('1 hr per eve')) return 7;

      return 8;
    }

    // â”€â”€ View Toggle â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function toggleSummaryView() {
      if (!allData.length) return;

      viewMode = (viewMode === 'detail') ? 'summary' : 'detail';

      const btn = document.getElementById('btnSummaryNav');
      const wrapDetail = document.getElementById('tableWrapper');
      const wrapSummary = document.getElementById('summaryWrapper');

      const btnDlTop = document.getElementById('btnSummaryDownload');
      const btnDlFilters = document.getElementById('btnExportSummary2');
      const btnFinal = document.getElementById('btnFinalDownload');

      if (viewMode === 'summary') {
        btn.classList.add('active');
        btn.textContent = 'Back To Detail View';
        wrapDetail.style.display = 'none';
        wrapSummary.style.display = 'block';
        btnDlTop.style.display = 'inline-flex';
        btnDlFilters.style.display = 'inline-flex';
        btnFinal.style.display = 'inline-flex';

        const summaryRows = buildNameSummaryRows(currentFiltered);
        renderNameSummaryTable(summaryRows);
        wrapSummary.scrollIntoView({ behavior: 'smooth', block: 'start' });
      } else {
        btn.classList.remove('active');
        btn.textContent = 'Open Summary View';
        wrapDetail.style.display = 'block';
        wrapSummary.style.display = 'none';
        btnDlTop.style.display = 'none';
        btnDlFilters.style.display = 'none';
        btnFinal.style.display = 'none';
        wrapDetail.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }

    function handleSelectedFile(file) {
      if (!file) return;
      document.getElementById('fileNameDisplay').textContent = 'Selected file: ' + file.name;
      hideError();

      const reader = new FileReader();
      reader.onload = e => {
        try { processWorkbook(e.target.result); }
        catch (err) { showError(err.message); }
      };
      reader.onerror = () => showError('Failed to read file.');
      reader.readAsArrayBuffer(file);
    }

    function initDropZone() {
      const dropZone = document.getElementById('dropZone');
      const fileInput = document.getElementById('fileInput');
      if (!dropZone || !fileInput) return;

      ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, e => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.add('dragover');
        });
      });

      ['dragleave', 'dragend', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, e => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.remove('dragover');
        });
      });

      dropZone.addEventListener('drop', e => {
        const file = e.dataTransfer && e.dataTransfer.files ? e.dataTransfer.files[0] : null;
        if (!file) return;
        handleSelectedFile(file);
      });
    }

    // â”€â”€ File Load â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function loadFromFile(input) {
      const file = input.files[0];
      if (!file) return;
      handleSelectedFile(file);
    }

    function processWorkbook(buf) {
      workbook = XLSX.read(buf, { type: 'array', cellDates: true });
      const sheets = workbook.SheetNames;
      if (!sheets.length) { showError('No sheets found.'); return; }

      const sel = document.getElementById('sheetSelect');
      sel.innerHTML = sheets.map((s, i) => `<option value="${i}">${esc(s)}</option>`).join('');

      let best = 0;
      sheets.forEach((s, i) => { if (/report|daily|data|march|employee/i.test(s)) best = i; });
      sel.value = best;
      document.getElementById('sheetPicker').style.display = 'flex';
      loadSheet();
    }

    function loadSheet() {
      const idx = parseInt(document.getElementById('sheetSelect').value) || 0;
      const ws = workbook.Sheets[workbook.SheetNames[idx]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
      parseAndRender(raw);
    }

    // â”€â”€ Parse â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function parseAndRender(rows) {
      if (!rows || rows.length === 0) return showError('Sheet is empty');

      let hIdx = -1;
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const r = rows[i].map(norm);
        if (r.some(h => h === 'date') && r.some(h => h === 'name')) { hIdx = i; break; }
      }
      if (hIdx < 0) return showError('Cannot find header row with "Date" and "Name".');

      const rawCols = rows[hIdx];
      const normCols = rawCols.map(norm);

      const mapOutToIn = [];
      const requiredNorms = requiredHeaders.map(norm);

      for (let i = 0; i < requiredNorms.length; i++) {
        const rNorm = requiredNorms[i];
        let foundIdx = -1;
        if (rNorm === 'tat / hour') foundIdx = matchCol(normCols, ['tat / hour', 'tat/hour', 'tat per hour', 'tat']);
        else if (rNorm === 'working hours') foundIdx = matchCol(normCols, ['working hours', 'duration', 'workinghours', 'work hours']);
        else foundIdx = matchCol(normCols, [rNorm]);
        mapOutToIn[i] = foundIdx;
      }

      sheetCols = requiredHeaders;
      allData = [];

      for (let i = hIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r.length || r.every(c => !String(c).trim())) continue;

        const raw = {};
        for (let c = 0; c < requiredHeaders.length; c++) {
          const inIdx = mapOutToIn[c];
          raw[c] = (inIdx >= 0 && r[inIdx] !== undefined) ? r[inIdx] : '';
        }

        const dateObj = parseDate(raw[D_IDX]);
        const dateSort = isoDate(dateObj) || String(raw[D_IDX] || '');

        const tatVal = pNum(raw[TAT_IDX]);
        const cntVal = pNum(raw[CNT_IDX]);
        const durVal = pNum(raw[WH_IDX]);

        const countIsEmpty = (cntVal === null || cntVal === 0);
        const tatIsEmpty = (tatVal === null || tatVal === 0);

        let hours = null;
        if (countIsEmpty || tatIsEmpty) hours = durVal;
        else hours = cntVal / tatVal;

        let minHours = null;
        if (hours !== null && durVal !== null) minHours = Math.min(hours, durVal);
        else if (hours !== null) minHours = hours;
        else if (durVal !== null) minHours = durVal;

        if (!String(raw[N_IDX] || '').trim() && !dateSort) continue;

        allData.push({ raw, dateObj, dateSort, __hours: hours, __minHours: minHours });
      }

      if (!allData.length) return showError('No valid data rows found after header.');

      document.getElementById('btnSummaryNav').disabled = false;

      populateFilters();
      applyFilters();

      document.getElementById('filtersBar').style.display = 'flex';
      document.getElementById('statsRow').style.display = 'grid';

      viewMode = 'detail';
      document.getElementById('btnSummaryNav').classList.remove('active');
      document.getElementById('btnSummaryNav').textContent = 'Open Summary View';
      document.getElementById('tableWrapper').style.display = 'block';
      document.getElementById('summaryWrapper').style.display = 'none';
      document.getElementById('btnSummaryDownload').style.display = 'none';
      document.getElementById('btnExportSummary2').style.display = 'none';
      document.getElementById('btnFinalDownload').style.display = 'none';
    }

    // â”€â”€ Filters â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function populateFilters() {
      const names = [...new Set(allData.map(d => String(d.raw[N_IDX] || '').trim()))].filter(Boolean).sort();
      const procs = [...new Set(allData.map(d => String(d.raw[P_IDX] || '').trim()))].filter(Boolean).sort();
      const dates = allData.map(d => d.dateSort).filter(Boolean).sort();

      document.getElementById('fName').innerHTML = '<option value="">All</option>' + names.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
      document.getElementById('fProcess').innerHTML = '<option value="">All</option>' + procs.map(p => `<option value="${esc(p)}">${esc(p)}</option>`).join('');

      if (dates.length) {
        document.getElementById('fFrom').value = dates[0];
        document.getElementById('fTo').value = dates[dates.length - 1];
      }
    }

    function applyFilters() {
      const nameFilter = document.getElementById('fName').value;
      const from = document.getElementById('fFrom').value;
      const to = document.getElementById('fTo').value;
      const processFilter = document.getElementById('fProcess').value;

      const fd = allData.filter(d => {
        if (nameFilter && String(d.raw[N_IDX] || '').trim() !== nameFilter) return false;
        if (processFilter && String(d.raw[P_IDX] || '').trim() !== processFilter) return false;
        if (from && d.dateSort < from) return false;
        if (to && d.dateSort > to) return false;
        return true;
      }).sort((a, b) => a.dateSort.localeCompare(b.dateSort));

      currentFiltered = fd;
      updateStats(fd);

      if (viewMode === 'summary') {
        const summaryRows = buildNameSummaryRows(fd);
        renderNameSummaryTable(summaryRows);
      } else {
        renderTable(fd);
      }
    }

    function clearFilters() {
      document.getElementById('fName').value = '';
      document.getElementById('fProcess').value = '';
      const dates = allData.map(d => d.dateSort).filter(Boolean).sort();
      if (dates.length) {
        document.getElementById('fFrom').value = dates[0];
        document.getElementById('fTo').value = dates[dates.length - 1];
      }
      applyFilters();
    }

    // â”€â”€ Stats (simple) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function updateStats(data) {
      document.getElementById('sTotal').textContent = data.length.toLocaleString();
      document.getElementById('sEmp').textContent = new Set(data.map(d => d.raw[N_IDX])).size;

      const totalCount = data.reduce((s, d) => s + (pNum(d.raw[CNT_IDX]) || 0), 0);
      document.getElementById('sCount').textContent = totalCount.toLocaleString();

      const hArr = data.map(d => d.__hours).filter(v => v !== null && v !== undefined);
      const mArr = data.map(d => d.__minHours).filter(v => v !== null && v !== undefined);

      const avgH = hArr.length ? (hArr.reduce((a, b) => a + b, 0) / hArr.length) : null;
      const avgM = mArr.length ? (mArr.reduce((a, b) => a + b, 0) / mArr.length) : null;

      document.getElementById('sAvgH').textContent = avgH === null ? '--' : f1(avgH);
      document.getElementById('sAvgMin').textContent = avgM === null ? '--' : f1(avgM);
    }

    // â”€â”€ Detail Table â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function renderTable(data) {
      const w = document.getElementById('tableWrapper');

      if (!data.length) {
        w.innerHTML = `<div class="state-box"><div class="state-icon">NONE</div><div class="state-title">No records</div><div class="state-sub">Adjust the filters to see matching rows.</div></div>`;
        return;
      }

      const grpRow = `<tr class="grp-row">
        <th class="grp-sheet" colspan="${sheetCols.length}">Sheet Data</th>
        <th class="grp-calc" colspan="2">Calculated</th>
      </tr>`;

      const colRow = `<tr class="col-row">
        ${sheetCols.map(h => `<th>${esc(h)}</th>`).join('')}
        <th>Hours</th>
        <th>Min Hours Cal</th>
      </tr>`;

      let body = '';
      let lastDate = '';

      data.forEach(row => {
        if (row.dateSort !== lastDate) {
          lastDate = row.dateSort;
          const label = row.dateObj ? fmtDate(row.dateObj) : row.dateSort;
          body += `<tr class="date-sep"><td colspan="${sheetCols.length + 2}">${esc(label)}</td></tr>`;
        }

        const cells = sheetCols.map((_, ci) => {
          if (ci === D_IDX) return `<td>${esc(row.dateObj ? fmtDate(row.dateObj) : String(row.raw[ci] || ''))}</td>`;
          if (ci === P_IDX) return `<td class="process-col">${esc(String(row.raw[ci] || '').trim()) || '--'}</td>`;
          return `<td>${esc(String(row.raw[ci] || '').trim()) || '--'}</td>`;
        }).join('');

        const hVal = row.__hours === null || row.__hours === undefined ? '--' : f1(row.__hours);
        const mVal = row.__minHours === null || row.__minHours === undefined ? '--' : f1(row.__minHours);

        body += `<tr>${cells}<td class="calc">${hVal}</td><td class="calc">${mVal}</td></tr>`;
      });

      w.innerHTML = `
        <div class="tbl-top">
          <div class="tbl-title">Report Data <span class="rec-count">${data.length} records</span></div>
          <div class="formula-note">
            <span>Hours</span> = Count / TAT per Hour (or Working Hours) -
            <span>Min Hours Cal</span> = Min(Hours, Working Hours)
          </div>
        </div>
        <div class="tbl-scroll">
          <table>
            <thead>${grpRow}${colRow}</thead>
            <tbody>${body}</tbody>
          </table>
        </div>`;
    }

    // â”€â”€ âœ… Summary rows per (Name + Date) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function buildNameSummaryRows(filteredRows) {
      const byName = new Map();

      filteredRows.forEach(d => {
        const name = String(d.raw[N_IDX] || '').trim();
        if (!name) return;

        const dateKey = d.dateSort || '';
        if (!dateKey) return;

        if (!byName.has(name)) byName.set(name, new Map());
        const byDate = byName.get(name);

        if (!byDate.has(dateKey)) {
          byDate.set(dateKey, {
            name,
            dateSort: dateKey,
            dateObj: d.dateObj,
            whSum: 0,
            hoursSum: 0,
            remarks: [],
            leave: []
          });
        }

        const g = byDate.get(dateKey);

        g.whSum += (pNum(d.raw[WH_IDX]) || 0);

        const mh = (d.__minHours === null || d.__minHours === undefined) ? 0 : (Number(d.__minHours) || 0);
        g.hoursSum += mh;

        const note = String(d.raw[8] || '').trim();
        if (note) g.remarks.push(note);

        const lv = String(d.raw[9] || '').trim();
        if (lv) g.leave.push(lv);
      });

      const out = [];
      const names = Array.from(byName.keys()).sort((a, b) => a.localeCompare(b));

      names.forEach(name => {
        const byDate = byName.get(name);
        const dates = Array.from(byDate.keys()).sort();

        dates.forEach(dateKey => {
          const g = byDate.get(dateKey);

          const remarksJoined = uniqJoin(g.remarks);
          const leaveJoined = uniqJoin(g.leave);

          const divisor = getProductivityDivisor(remarksJoined, leaveJoined);

          const productivity = g.hoursSum ? (g.hoursSum * 100 / divisor) : 0;

          let targetStatus = '';
          if (productivity !== 0 && productivity < 100) targetStatus = 'Target not achieved';

          const concatenate = [remarksJoined, leaveJoined, targetStatus].filter(Boolean).join(', ');

          out.push({
            name: g.name,
            dateSort: g.dateSort,
            dateDisplay: (g.dateObj ? fmtDate(g.dateObj) : g.dateSort),
            whSum: g.whSum,
            hoursSum: g.hoursSum,
            productivity,
            remarks: remarksJoined,
            leave: leaveJoined,
            targetStatus,
            concatenate,
            __divisor: divisor
          });
        });
      });

      return out;
    }

    // â”€â”€ âœ… Summary table render â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function renderNameSummaryTable(summaryRows) {
      const w = document.getElementById('summaryWrapper');

      if (!summaryRows || !summaryRows.length) {
        w.innerHTML = `<div class="state-box">
          <div class="state-icon">SUM</div>
          <div class="state-title">No summary records</div>
          <div class="state-sub">Adjust filters or check your sheet data.</div>
        </div>`;
        return;
      }

      const grpRow = `<tr class="grp-row">
        <th class="grp-sheet" colspan="${summaryHeaders.length}">Summary Report</th>
      </tr>`;

      const colRow = `<tr class="col-row">` +
        summaryHeaders.map(h => `<th style="white-space:nowrap">${esc(h)}</th>`).join('') +
        `</tr>`;

      let body = '';
      summaryRows.forEach(r => {
        body += `<tr>
          <td>${esc(r.name)}</td>
          <td>${esc(r.dateDisplay) || '--'}</td>
          <td class="num">${f1(r.whSum)}</td>
          <td class="calc">${f1(r.hoursSum)}</td>
          <td class="num">${f0(r.productivity)}</td>
          <td class="process-col">${esc(r.remarks) || ''}</td>
          <td class="process-col">${esc(r.leave) || ''}</td>
          <td>${esc(r.targetStatus) || ''}</td>
          <td class="process-col">${esc(r.concatenate) || ''}</td>
        </tr>`;
      });

      w.innerHTML = `
        <div class="tbl-top">
          <div class="tbl-title">Summary Report <span class="rec-count">${summaryRows.length} records</span></div>
          <div class="formula-note">
            <span>Working Hours</span> = Sum of working hours per day -
            <span>Hours</span> = Sum of Min Hours Cal per day -
            <span>Productivity %</span> = Hours x 100 / divisor (8, 7, 6, or 4 based on Remarks and Leave)
          </div>
        </div>
        <div class="tbl-scroll">
          <table>
            <thead>${grpRow}${colRow}</thead>
            <tbody>${body}</tbody>
          </table>
        </div>`;
    }

    // â”€â”€ âœ… Export Summary CSV â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function exportSummaryCSV(summaryRows) {
      if (!summaryRows || !summaryRows.length) {
        showError('No summary rows to export (check filters).');
        return;
      }

      const header = summaryHeaders;
      const rows = summaryRows.map(r => {
        const vals = [
          r.name || '',
          r.dateDisplay || '',
          (r.whSum == null ? '' : (Math.round(Number(r.whSum) * 10) / 10).toFixed(1)),
          (r.hoursSum == null ? '' : (Math.round(Number(r.hoursSum) * 10) / 10).toFixed(1)),
          (r.productivity == null ? '' : String(Math.floor(Number(r.productivity) + 0.5))),
          r.remarks || '',
          r.leave || '',
          r.targetStatus || '',
          r.concatenate || ''
        ].map(v => String(v).replace(/"/g, '""'));
        return vals.map(v => `"${v}"`).join(',');
      });

      const csv = [header.map(v => `"${v}"`).join(','), ...rows].join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = `summary_report_${new Date().toISOString().split('T')[0]}.csv`;
      a.click();
    }

    // â”€â”€ âœ… FINAL REPORT (NEW) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    const finalHeaders = [
      "Date", "Name", "Working Hours", "Productivity %", "Average of Productivity",
      "Remarks", "Quality(100%)", "Working days", "Late Login", "Leave", "Permission",
      "No. of Process Worked", "Name of process worked", "Notes", "Errors %", "Floor Ethics %",
      "Login", "Leave"
    ];

    function cleanFinalRemarks(text) {
      return String(text || '')
        .split(',')
        .map(v => v.trim())
        .filter(v => v && v.toLowerCase() !== 'present')
        .join(', ');
    }

    function buildFinalReportRows(filteredRows) {
      const summaryRows = buildNameSummaryRows(filteredRows);

      // Name -> productivity array (exclude 0 for average)
      const prodByName = new Map();
      // Name -> working days count (productivity > 0)
      const workingDaysByName = new Map();

      summaryRows.forEach(r => {
        const name = String(r.name || '').trim();
        if (!name) return;

        const prod = Number(r.productivity) || 0;

        if (!prodByName.has(name)) prodByName.set(name, []);
        if (prod > 0) prodByName.get(name).push(prod);

        if (!workingDaysByName.has(name)) workingDaysByName.set(name, new Set());
        if (prod > 0 && r.dateSort) workingDaysByName.get(name).add(r.dateSort);
      });

      // Name -> Set(process) across the whole filtered month/date range
      const procMap = new Map();
      filteredRows.forEach(d => {
        const name = String(d.raw[N_IDX] || '').trim();
        const proc = String(d.raw[P_IDX] || '').trim();
        if (!name || !proc) return;

        if (!procMap.has(name)) procMap.set(name, new Set());
        procMap.get(name).add(proc);
      });

      const processRowIndexByName = new Map();

      return summaryRows.map(r => {
        const name = String(r.name || '').trim();
        const prodList = prodByName.get(name) || [];
        const avgProd = prodList.length ? (prodList.reduce((a, b) => a + b, 0) / prodList.length) : 0;

        const workingDays = workingDaysByName.get(name) ? workingDaysByName.get(name).size : 0;

        const procSet = procMap.get(name) || new Set();
        const procList = Array.from(procSet).sort();
        const rowIndex = processRowIndexByName.get(name) || 0;
        const processForThisRow = procList[rowIndex] || "";
        processRowIndexByName.set(name, rowIndex + 1);

        return {
          Date: r.dateDisplay || '',
          Name: name,
          WorkingHours: r.whSum ?? '',
          Productivity: r.productivity ?? '',
          AvgProductivity: avgProd,
          Remarks: cleanFinalRemarks(r.concatenate),
          Quality: "100%",
          WorkingDays: workingDays,
          LateLogin: "",
          Leave: r.leave || "",
          Permission: "",
          NoProcessWorked: procList.length,
          NameProcessWorked: processForThisRow,
          Notes: "",
          Errors: "",
          FloorEthics: "",
          Login: "",
          Leave2: ""
        };
      });
    }

    function exportFinalCSV(rows) {
      if (!rows || !rows.length) {
        showError('No final report rows to export (check filters).');
        return;
      }

      const csvRows = [];

      rows.forEach((r, index) => {
        const wh = (r.WorkingHours === '' ? '' : (Math.round(Number(r.WorkingHours) * 10) / 10).toFixed(1));
        const prod = (r.Productivity === '' ? '' : String(Math.floor(Number(r.Productivity) + 0.5)));
        const avg = (r.AvgProductivity === '' ? '' : String(Math.floor(Number(r.AvgProductivity) + 0.5)));

        const vals = [
          r.Date || "",
          r.Name || "",
          wh,
          prod,
          avg,
          r.Remarks || "",
          r.Quality || "",
          String(r.WorkingDays ?? ""),
          r.LateLogin || "",
          r.Leave || "",
          r.Permission || "",
          String(r.NoProcessWorked ?? ""),
          r.NameProcessWorked || "",
          r.Notes || "",
          r.Errors || "",
          r.FloorEthics || "",
          r.Login || "",
          r.Leave2 || ""
        ].map(v => String(v).replace(/"/g, '""'));

        csvRows.push(vals.map(v => `"${v}"`).join(','));

        const nextRow = rows[index + 1];
        if (nextRow && String(nextRow.Name || '').trim() !== String(r.Name || '').trim()) {
          csvRows.push(new Array(finalHeaders.length).fill('""').join(','));
        }
      });

      const csv = [finalHeaders.map(v => `"${v}"`).join(','), ...csvRows].join('\n');

      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = `final_report_${new Date().toISOString().split('T')[0]}.csv`;
      a.click();
    }

    // â”€â”€ Export CSV (Detail) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function exportCSV() {
      const h = [...sheetCols, 'Hours', 'Min Hours Cal'];
      const rows = allData.map(d => {
        const vals = sheetCols.map((_, ci) => String(d.raw[ci] || '').replace(/"/g, '""'));
        vals.push(d.__hours !== null && d.__hours !== undefined ? d.__hours.toFixed(4) : '');
        vals.push(d.__minHours !== null && d.__minHours !== undefined ? d.__minHours.toFixed(4) : '');
        return vals.map(v => `"${v}"`).join(',');
      });
      const csv = [h.map(v => `"${v}"`).join(','), ...rows].join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = `employee_report_${new Date().toISOString().split('T')[0]}.csv`;
      a.click();
    }

    // â”€â”€ UI helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    function showError(msg) {
      const b = document.getElementById('errorBox');
      document.getElementById('errorMsg').textContent = msg;
      b.style.display = 'flex';
    }

    function hideError() {
      document.getElementById('errorBox').style.display = 'none';
    }

    initDropZone();
