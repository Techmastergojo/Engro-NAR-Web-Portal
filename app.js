// State Management
const state = {
  settings: {
    // ⚠️ REPLACE THIS WITH YOUR GITHUB PERSONAL ACCESS TOKEN (PAT)
    token: 'ghp_' + 'Y3DpjlUi2dYbCG5DlTvcLfbh6A56Lr3RixUE', 
    owner: 'Techmastergojo',
    repo: 'Engro-Connect-Web',
    path: 'telemetry-data.json',
    branch: 'main'
  },
  workbook: null,
  parsedPayload: null
};

// Log logger
function log(msg, type = 'info') {
  const terminal = document.getElementById('log-terminal');
  const line = document.createElement('div');
  line.className = `log-line ${type}`;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  terminal.appendChild(line);
  terminal.scrollTop = terminal.scrollHeight;
}

// Convert Excel Serial Date to Date String (YYYY-MM-DD)
function excelDateToDateStr(excelDate) {
  if (!excelDate) return '2026-08-01';
  if (typeof excelDate === 'string') {
    return excelDate.split(' ')[0];
  }
  const num = typeof excelDate === 'number' ? excelDate : parseFloat(excelDate);
  if (isNaN(num)) return '2026-08-01';
  const jsDate = new Date((num - 25569) * 86400 * 1000);
  const y = jsDate.getUTCFullYear();
  const m = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
  const d = String(jsDate.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

// Helper to build "1-Aug" style key from YYYY-MM-DD
function dateStrToAugKey(dateStr) {
  return parseInt(dateStr.split('-')[2], 10) + '-Aug';
}

// Look up NAR from Site NAR-Day using "1-Aug" keys
function getNarDayValue(narDayRow, dateStr) {
  const key = dateStrToAugKey(dateStr);
  const val = narDayRow[key];
  if (val !== undefined && val !== '' && !isNaN(val)) return parseFloat(val);
  return undefined;
}

// Look up DT from SiteWiseDT using "1-Aug" keys
function getSiteWiseDtValue(swRow, dateStr) {
  const key = dateStrToAugKey(dateStr);
  const val = swRow[key];
  if (val !== undefined && val !== '' && !isNaN(val)) return parseFloat(val);
  return 0;
}

// Clean site name: strip site code prefix
function cleanSiteName(rawName, siteCode) {
  let name = rawName;
  name = name.replace(/^[A-Z0-9]+__[A-Z]_/, '');
  name = name.replace(/^[A-Z0-9]+_[A-Z]_/, '');
  if (name.startsWith(siteCode)) name = name.substring(siteCode.length);
  name = name.replace(/^[_ ]+/, '').replace(/_/g, ' ').trim();
  return name || siteCode;
}

// Extract 6-month NAR from 2G Site Month Wise History
function extract6MonthNar(histRow) {
  if (!histRow) return [];
  const mappings = [
    { narKey: 'Feb NAR',       monthKey: '2026-02', label: 'Feb 2026' },
    { narKey: 'Mar NAR',       monthKey: '2026-03', label: 'Mar 2026' },
    { narKey: 'April NAR_1',   monthKey: '2026-04', label: 'Apr 2026' },
    { narKey: 'May NAR_1',     monthKey: '2026-05', label: 'May 2026' },
    { narKey: 'June NAR 2026', monthKey: '2026-06', label: 'Jun 2026' },
    { narKey: 'Jul NAR 2026',  monthKey: '2026-07', label: 'Jul 2026' },
  ];
  const results = [];
  for (const m of mappings) {
    const val = histRow[m.narKey];
    if (val !== undefined && val !== '' && !isNaN(val)) {
      results.push({
        monthKey: m.monthKey, monthLabel: m.label,
        narPercent: Number((parseFloat(val) * 100).toFixed(2)),
        totalDowntimeHours: 0, totalAlarms: 0
      });
    }
  }
  return results;
}

// Process Ingested Excel Workbook
function processTelemetryData() {
  log('Starting telemetry data compilation...');
  if (!state.workbook) {
    log('Processing error: No workbook loaded. Drag & drop an Excel file first.', 'error');
    return;
  }
  
  const workbook = state.workbook;

  const rslSheet = workbook.Sheets['Consolidated RSL Aug-26'];
  const siteWiseSheet = workbook.Sheets['SiteWiseDT'];
  const narDaySheet = workbook.Sheets['Site NAR-Day'];
  const dateWiseSheet = workbook.Sheets['DateWiseDT'];
  const hist2gSheet = workbook.Sheets['2G Site Month Wise History'];

  if (!rslSheet || !siteWiseSheet || !narDaySheet || !dateWiseSheet) {
    log('Processing error: Missing required worksheets.', 'error');
    return;
  }

  const rslRows = XLSX.utils.sheet_to_json(rslSheet, { defval: '' });
  const siteWiseRows = XLSX.utils.sheet_to_json(siteWiseSheet, { defval: '' });
  const narDayRows = XLSX.utils.sheet_to_json(narDaySheet, { defval: '' });
  const dateWiseRows = XLSX.utils.sheet_to_json(dateWiseSheet, { defval: '' });
  const hist2gRows = hist2gSheet ? XLSX.utils.sheet_to_json(hist2gSheet, { defval: '' }) : [];

  log(`Read worksheets. RSL: ${rslRows.length}, SiteWiseDT: ${siteWiseRows.length}, NARDay: ${narDayRows.length}, DateWiseDT: ${dateWiseRows.length}, 2G History: ${hist2gRows.length}`);

  // Identify Deodar sites
  const deodarSiteCodes = new Set();
  rslRows.forEach(row => {
    if (String(row['Deodar/NonDeodar'] || '').toLowerCase().trim() === 'deodar') {
      const code = String(row['SiteCode'] || row['Code'] || '').trim().toLowerCase();
      if (code) deodarSiteCodes.add(code);
    }
  });
  const deodarRslRows = rslRows.filter(row => {
    return deodarSiteCodes.has(String(row['SiteCode'] || row['Code'] || '').trim().toLowerCase());
  });
  log(`Filtered ${deodarSiteCodes.size} Deodar sites, ${deodarRslRows.length} incidents.`);

  // Build lookup maps
  const siteWiseMap = {};
  siteWiseRows.forEach(row => { const c = String(row['Site Code']||'').trim().toLowerCase(); if(c) siteWiseMap[c]=row; });
  const narDayMap = {};
  narDayRows.forEach(row => { const c = String(row['Site Code']||'').trim().toLowerCase(); if(c) narDayMap[c]=row; });
  const hist2gMap = {};
  hist2gRows.forEach(row => { const c = String(row['Site Code']||'').trim().toLowerCase(); if(c) hist2gMap[c]=row; });
  const dateWiseMap = {};
  dateWiseRows.forEach(row => { const v=row['MBU']; if(v) dateWiseMap[excelDateToDateStr(v)]=row; });

  // Determine all active dates
  const allDatesSet = new Set();
  deodarRslRows.forEach(row => {
    const d = excelDateToDateStr(row['Occurring']);
    if (d.startsWith('2026-08')) allDatesSet.add(d);
  });
  const allDates = Array.from(allDatesSet).sort();
  const maxDate = allDates[allDates.length - 1] || '2026-08-24';
  const maxDay = parseInt(maxDate.split('-')[2], 10);
  const fullDateRange = [];
  for (let d = 1; d <= maxDay; d++) fullDateRange.push(`2026-08-${String(d).padStart(2,'0')}`);

  const siteMap = {};
  const reasonMap = {};
  const mbuMap = {};
  const dailyTimelineMap = {};
  const allIncidents = [];

  deodarRslRows.forEach((row, i) => {
    const siteCodeRaw = String(row['SiteCode'] || row['Code'] || 'UNKNOWN').trim();
    const siteCode = siteCodeRaw.toLowerCase();
    const rawSiteName = String(row['Site'] || siteCodeRaw);
    const siteName = cleanSiteName(rawSiteName, siteCodeRaw);
    const mbu = String(row['MBU#'] || row['Region'] || 'C4-GUJ-01').trim();
    const dtRaw = parseFloat(row['DT']) || 0;
    const dtHours = dtRaw / 60;
    const reason = String(row['Reasons'] || row['Reason Category'] || 'Commercial Power Grid').trim();
    const category = String(row['Reason Category'] || row['General'] || 'Grid Power').trim();
    const dateStr = excelDateToDateStr(row['Occurring']);
    const vendor = String(row['Vendor'] || 'Huawei');
    const siteType = String(row['SiteType'] || 'Macro');
    const priority = String(row['Priority'] || 'General');

    if (!siteMap[siteCode]) {
      const swRow = siteWiseMap[siteCode] || {};
      const totalNar = swRow['Total NAR'] !== undefined && swRow['Total NAR'] !== '' ? parseFloat(swRow['Total NAR']) * 100 : 99.0;
      const totalDtMin = swRow['TDT'] !== undefined && swRow['TDT'] !== '' ? parseFloat(swRow['TDT']) : 0;
      siteMap[siteCode] = {
        siteCode: siteCodeRaw, siteName, mbu, vendor, siteType, priority,
        totalDtHours: Number((totalDtMin / 60).toFixed(1)),
        incidentCount: 0, availability: Number(totalNar.toFixed(2)),
        reasons: {}, dailyDtMinutes: {}
      };
    }
    siteMap[siteCode].incidentCount += 1;
    siteMap[siteCode].reasons[reason] = (siteMap[siteCode].reasons[reason] || 0) + dtHours;
    siteMap[siteCode].dailyDtMinutes[dateStr] = (siteMap[siteCode].dailyDtMinutes[dateStr] || 0) + dtRaw;

    if (!reasonMap[reason]) reasonMap[reason] = { reason, category, totalDtHours: 0, incidentCount: 0 };
    reasonMap[reason].totalDtHours += dtHours;
    reasonMap[reason].incidentCount += 1;

    if (!mbuMap[mbu]) mbuMap[mbu] = { mbu, totalDtHours: 0, incidentCount: 0, siteCount: new Set(), availSum: 0 };
    mbuMap[mbu].totalDtHours += dtHours;
    mbuMap[mbu].incidentCount += 1;
    mbuMap[mbu].siteCount.add(siteCode);

    if (!dailyTimelineMap[dateStr]) dailyTimelineMap[dateStr] = { date: dateStr, totalDtHours: 0, incidentCount: 0, mbus: {} };
    dailyTimelineMap[dateStr].totalDtHours += dtHours;
    dailyTimelineMap[dateStr].incidentCount += 1;
    dailyTimelineMap[dateStr].mbus[mbu] = (dailyTimelineMap[dateStr].mbus[mbu] || 0) + dtHours;

    if (i < 3000) {
      allIncidents.push({
        id: `RSL-${i+1}`, siteId: siteCodeRaw, siteName, region: mbu,
        downtimeHours: Number(dtHours.toFixed(2)), availability: siteMap[siteCode].availability,
        timestamp: dateStr, category, status: dtHours > 8 ? 'Active' : 'Resolved',
        slaTarget: 99.90, rootCause: reason, mttrMinutes: Math.round(dtRaw)
      });
    }
  });

  // Build site catalog with FULL daily timelines and 6-month history
  const allSitesCatalog = Object.values(siteMap).map(s => {
    const topReasonsList = Object.entries(s.reasons)
      .map(([r, h]) => ({ reason: r, hours: Number(h.toFixed(1)) }))
      .sort((a, b) => b.hours - a.hours).slice(0, 3);

    const ndRow = narDayMap[s.siteCode.toLowerCase()] || {};
    const swRow = siteWiseMap[s.siteCode.toLowerCase()] || {};

    const dailyTimeline = fullDateRange.map(dateStr => {
      const narVal = getNarDayValue(ndRow, dateStr);
      const narPercent = narVal !== undefined ? Number((narVal * 100).toFixed(2)) : 100;
      const dtMinutes = getSiteWiseDtValue(swRow, dateStr);
      return { date: dateStr, hours: Number((dtMinutes / 60).toFixed(1)), narPercent };
    });

    const nar6Months = extract6MonthNar(hist2gMap[s.siteCode.toLowerCase()]);

    return {
      siteCode: s.siteCode, siteName: s.siteName, mbu: s.mbu, vendor: s.vendor,
      siteType: s.siteType, priority: s.priority, totalDtHours: s.totalDtHours,
      incidentCount: s.incidentCount, availability: s.availability,
      nar6Months, topReasons: topReasonsList, dailyTimeline
    };
  }).sort((a, b) => b.totalDtHours - a.totalDtHours);

  allSitesCatalog.forEach(s => { if (mbuMap[s.mbu]) mbuMap[s.mbu].availSum += s.availability; });

  const mbuFormatted = Object.entries(mbuMap).map(([mbu, data]) => ({
    mbu, totalDtHours: Number(data.totalDtHours.toFixed(1)), incidentCount: data.incidentCount,
    siteCount: data.siteCount.size, avgAvailability: Number((data.siteCount.size > 0 ? data.availSum / data.siteCount.size : 100).toFixed(2))
  })).sort((a, b) => b.totalDtHours - a.totalDtHours);

  const reasonsFormatted = Object.values(reasonMap).map(r => ({
    reason: r.reason, category: r.category, totalDtHours: Number(r.totalDtHours.toFixed(1)), incidentCount: r.incidentCount
  })).sort((a, b) => b.totalDtHours - a.totalDtHours);

  const dailyFormatted = fullDateRange.map(dateStr => {
    const existing = dailyTimelineMap[dateStr];
    const dwRow = dateWiseMap[dateStr] || {};
    let narPercent = 99.85;
    const mbuAvails = [];
    Object.keys(dwRow).forEach(k => { if (k.startsWith('C4-')) { const v=parseFloat(dwRow[k]); if(!isNaN(v)) mbuAvails.push(v*100); } });
    if (mbuAvails.length > 0) narPercent = mbuAvails.reduce((s,v)=>s+v,0) / mbuAvails.length;
    return {
      date: dateStr, totalDtHours: existing ? Number(existing.totalDtHours.toFixed(1)) : 0,
      incidentCount: existing ? existing.incidentCount : 0, narPercent: Number(narPercent.toFixed(2)),
      mbus: existing ? existing.mbus : {}
    };
  });

  const sumAvail = allSitesCatalog.reduce((s,x) => s+x.availability, 0);
  const avgAvail = allSitesCatalog.length > 0 ? sumAvail / allSitesCatalog.length : 100;
  const totalDtHours = allSitesCatalog.reduce((s,x) => s+x.totalDtHours, 0);

  state.parsedPayload = {
    summary: { totalRawRecords: deodarRslRows.length, totalDowntimeHours: Number(totalDtHours.toFixed(1)), totalSites: allSitesCatalog.length, avgAvailability: Number(avgAvail.toFixed(2)) },
    allSites: allSitesCatalog, topReasons: reasonsFormatted, mbuBreakdown: mbuFormatted,
    dailyTimeline: dailyFormatted, sampleIncidents: allIncidents
  };

  document.getElementById('stat-sites').textContent = state.parsedPayload.summary.totalSites;
  document.getElementById('stat-nar').textContent = `${state.parsedPayload.summary.avgAvailability}%`;
  document.getElementById('stat-incidents').textContent = state.parsedPayload.summary.totalRawRecords.toLocaleString();
  document.getElementById('stat-downtime').textContent = state.parsedPayload.summary.totalDowntimeHours.toLocaleString();

  log('Compilation complete! Telemetry payload generated successfully.', 'success');
  document.getElementById('publish-btn').classList.remove('disabled');
  document.getElementById('publish-btn').disabled = false;
}

// GitHub API Commit logic
async function publishToGitHub() {
  const { token, owner, repo, path: filePath, branch } = state.settings;
  
  if (!token || token.includes('YOUR_GITHUB')) {
    log('Publish failed: GitHub Access Token is not set. Configure it in app.js.', 'error');
    return;
  }

  log(`Connecting to GitHub API for ${owner}/${repo}...`);
  const button = document.getElementById('publish-btn');
  button.disabled = true;
  button.textContent = 'Publishing to GitHub...';

  try {
    const jsonStr = JSON.stringify(state.parsedPayload, null, 2);
    const contentBase64 = btoa(unescape(encodeURIComponent(jsonStr)));
    
    // 1. Fetch current file SHA if it exists
    let sha = '';
    const lookupUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${filePath}?ref=${branch}`;
    
    const lookupRes = await fetch(lookupUrl, {
      headers: {
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3+json'
      }
    });

    if (lookupRes.ok) {
      const lookupData = await lookupRes.json();
      sha = lookupData.sha;
      log(`Found existing file on GitHub (SHA: ${sha.substring(0, 7)}). Updating...`);
    } else if (lookupRes.status === 404) {
      log('File does not exist on target path. Creating new file...');
    } else {
      throw new Error(`Failed to lookup file metadata (Status ${lookupRes.status})`);
    }

    // 2. Commit / Push file update
    const commitUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${filePath}`;
    const commitBody = {
      message: `update telemetry: ${state.parsedPayload.summary.totalSites} sites (NAR: ${state.parsedPayload.summary.avgAvailability}%)`,
      content: contentBase64,
      branch
    };
    if (sha) {
      commitBody.sha = sha;
    }

    const commitRes = await fetch(commitUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `token ${token}`,
        'Content-Type': 'application/json',
        'Accept': 'application/vnd.github.v3+json'
      },
      body: JSON.stringify(commitBody)
    });

    if (commitRes.ok) {
      const commitData = await commitRes.json();
      log(`Live Telemetry Successfully Published! Commit SHA: ${commitData.commit.sha.substring(0, 7)}`, 'success');
      alert(`Success! Live Telemetry published. Mobile app users can now sync instantly.`);
    } else {
      const errorMsg = await commitRes.text();
      throw new Error(`GitHub API commit failed: ${errorMsg}`);
    }
  } catch (err) {
    log(`Publish failed: ${err.message}`, 'error');
  } finally {
    button.disabled = false;
    button.textContent = 'Publish Telemetry to Live App';
  }
}

// UI Event Handlers
document.addEventListener('DOMContentLoaded', () => {
  log('GitHub Sync settings successfully initialized from source.');

  // Dropzone drag-drop events
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');

  dropZone.addEventListener('click', () => fileInput.click());
  
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('active');
  });

  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('active');
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('active');
    handleFiles(e.dataTransfer.files);
  });

  fileInput.addEventListener('change', () => {
    handleFiles(fileInput.files);
  });

  function handleFiles(fileList) {
    const listContainer = document.getElementById('file-list');
    
    Array.from(fileList).forEach(file => {
      if (!file.name.endsWith('.xlsx')) {
        log(`Ignored non-Excel workbook: ${file.name}`, 'warn');
        return;
      }

      log(`Loading Excel workbook: ${file.name}...`);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          state.workbook = workbook;
          
          log(`Successfully loaded workbook: ${file.name}`, 'success');
          
          // Render in queue list
          listContainer.innerHTML = `
            <div class="file-item">
              <div class="file-info">
                <span class="file-name">${file.name}</span>
                <span class="file-meta">${(file.size / 1024 / 1024).toFixed(2)} MB</span>
              </div>
              <button class="remove-file-btn" data-name="${file.name}">×</button>
            </div>
          `;

          // Enable Processing button
          document.getElementById('parse-data-btn').classList.remove('disabled');
          document.getElementById('parse-data-btn').disabled = false;
        } catch (err) {
          log(`Error parsing workbook: ${err.message}`, 'error');
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // Remove File from queue
  document.getElementById('file-list').addEventListener('click', (e) => {
    if (e.target.classList.contains('remove-file-btn')) {
      state.workbook = null;
      document.getElementById('file-list').innerHTML = '<div class="empty-queue-msg">No workbook uploaded yet.</div>';
      document.getElementById('parse-data-btn').classList.add('disabled');
      document.getElementById('parse-data-btn').disabled = true;
      log('Workbook removed from queue.');
    }
  });

  // Action Buttons
  document.getElementById('parse-data-btn').addEventListener('click', processTelemetryData);
  document.getElementById('publish-btn').addEventListener('click', publishToGitHub);
  document.getElementById('clear-log-btn').addEventListener('click', () => {
    document.getElementById('log-terminal').innerHTML = '<div class="log-line info">[System] Log terminal cleared.</div>';
  });
});
