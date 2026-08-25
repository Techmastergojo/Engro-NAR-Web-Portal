// State Management
const state = {
  settings: {
    // ⚠️ REPLACE THIS WITH YOUR GITHUB PERSONAL ACCESS TOKEN (PAT)
    token: 'YOUR_GITHUB_PERSONAL_ACCESS_TOKEN_HERE', 
    owner: 'Techmastergojo',
    repo: 'Engro-Connect-Web',
    path: 'telemetry-data.json',
    branch: 'main'
  },
  files: {}, // Maps fileName -> content
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

// Custom CSV Parser
function parseCSV(content) {
  const rows = [];
  let currentField = '';
  let inQuotes = false;
  let currentRow = [];
  
  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const nextChar = content[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentField += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      currentRow.push(currentField.trim());
      currentField = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
      currentRow.push(currentField.trim());
      rows.push(currentRow);
      currentRow = [];
      currentField = '';
    } else {
      currentField += char;
    }
  }
  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField.trim());
    rows.push(currentRow);
  }
  
  if (rows.length === 0) return [];
  
  const rawHeaders = rows[0];
  const headers = rawHeaders.map(h => h.replace(/^\ufeff/, '').trim());
  
  const records = [];
  for (let i = 1; i < rows.length; i++) {
    const row = {};
    headers.forEach((header, idx) => {
      row[header] = rows[i][idx] !== undefined ? rows[i][idx] : '';
    });
    records.push(row);
  }
  return records;
}

function normalizeDateStr(dateStr) {
  if (!dateStr) return '2026-08-01';
  return dateStr.split(' ')[0];
}

function getDailyValue(row, dateStr) {
  const key1 = `${dateStr} 00:00:00`;
  const key2 = dateStr;
  if (row[key1] !== undefined && row[key1] !== '') return parseFloat(row[key1]);
  if (row[key2] !== undefined && row[key2] !== '') return parseFloat(row[key2]);
  return undefined;
}

// Process Ingested CSV Files
function processTelemetryData() {
  log('Starting telemetry data compilation...');
  
  // Identify uploaded sheets
  let rslRows = [];
  let siteWiseRows = [];
  let narDayRows = [];
  let dateWiseRows = [];

  Object.entries(state.files).forEach(([name, content]) => {
    const records = parseCSV(content);
    if (name.includes('Consolidated RSL')) rslRows = records;
    else if (name.includes('SiteWiseDT')) siteWiseRows = records;
    else if (name.includes('Site NAR-Day')) narDayRows = records;
    else if (name.includes('DateWiseDT')) dateWiseRows = records;
  });

  if (rslRows.length === 0 || siteWiseRows.length === 0 || narDayRows.length === 0 || dateWiseRows.length === 0) {
    log('Processing error: Missing required files in the queue. Make sure you uploaded all 4 sheets.', 'error');
    return;
  }

  log(`Loaded ${rslRows.length} outage logs, ${siteWiseRows.length} site list, ${narDayRows.length} timeline matrix.`);

  // 1. Identify Deodar Sites
  const deodarSiteCodes = new Set();
  rslRows.forEach(row => {
    const isDeodar = String(row['Deodar/NonDeodar'] || '').toLowerCase().trim() === 'deodar';
    if (isDeodar) {
      const siteCode = String(row['SiteCode'] || row['Code'] || '').trim().toLowerCase();
      if (siteCode) deodarSiteCodes.add(siteCode);
    }
  });

  const deodarRslRows = rslRows.filter(row => {
    const siteCode = String(row['SiteCode'] || row['Code'] || '').trim().toLowerCase();
    return deodarSiteCodes.has(siteCode);
  });

  log(`Filtered ${deodarSiteCodes.size} unique Deodar sites representing ${deodarRslRows.length} incidents.`);

  // Mapping structures
  const siteWiseMap = {};
  siteWiseRows.forEach(row => {
    const code = String(row['Site Code'] || '').trim().toLowerCase();
    if (code) siteWiseMap[code] = row;
  });

  const narDayMap = {};
  narDayRows.forEach(row => {
    const code = String(row['Site Code'] || '').trim().toLowerCase();
    if (code) narDayMap[code] = row;
  });

  const dateWiseMap = {};
  dateWiseRows.forEach(row => {
    const mbuVal = row['MBU'];
    if (mbuVal) {
      const dateStr = normalizeDateStr(mbuVal);
      dateWiseMap[dateStr] = row;
    }
  });

  const siteMap = {};
  const reasonMap = {};
  const mbuMap = {};
  const dailyTimelineMap = {};
  const allIncidents = [];

  deodarRslRows.forEach((row, i) => {
    const siteCodeRaw = String(row['SiteCode'] || row['Code'] || 'UNKNOWN').trim();
    const siteCode = siteCodeRaw.toLowerCase();
    const rawSiteName = String(row['Site'] || siteCodeRaw);
    const siteName = rawSiteName.replace(/^[A-Z0-9]+__S_/, '').replace(/^[A-Z0-9]+_H_/, '').replace(/_/g, ' ');
    const mbu = String(row['MBU#'] || row['Region'] || 'C4-GUJ-01').trim();
    
    // DT to Hours
    const dtRaw = parseFloat(row['DT']) || 0;
    const dtHours = dtRaw / 60;
    
    const reason = String(row['Reasons'] || row['Reason Category'] || 'Commercial Power Grid').trim();
    const category = String(row['Reason Category'] || row['General'] || 'Grid Power').trim();
    const dateStr = normalizeDateStr(row['Occurring']);
    const vendor = String(row['Vendor'] || 'Huawei');
    const siteType = String(row['SiteType'] || 'Macro');
    const priority = String(row['Priority'] || 'General');

    if (!siteMap[siteCode]) {
      const swRow = siteWiseMap[siteCode] || {};
      const excelAvail = swRow['Total NAR'] !== undefined && swRow['Total NAR'] !== '' ? parseFloat(swRow['Total NAR']) * 100 : 99.0;
      const excelDtHours = swRow['TDT'] !== undefined && swRow['TDT'] !== '' ? parseFloat(swRow['TDT']) / 60 : 0;

      siteMap[siteCode] = {
        siteCode: siteCodeRaw,
        siteName,
        mbu,
        vendor,
        siteType,
        priority,
        totalDtHours: excelDtHours,
        incidentCount: 0,
        availability: Number(excelAvail.toFixed(2)),
        reasons: {},
        dailyDt: {}
      };
    }
    siteMap[siteCode].incidentCount += 1;
    siteMap[siteCode].reasons[reason] = (siteMap[siteCode].reasons[reason] || 0) + dtHours;
    siteMap[siteCode].dailyDt[dateStr] = (siteMap[siteCode].dailyDt[dateStr] || 0) + dtHours;

    if (!reasonMap[reason]) {
      reasonMap[reason] = { reason, category, totalDtHours: 0, incidentCount: 0 };
    }
    reasonMap[reason].totalDtHours += dtHours;
    reasonMap[reason].incidentCount += 1;

    if (!mbuMap[mbu]) {
      mbuMap[mbu] = { mbu, totalDtHours: 0, incidentCount: 0, siteCount: new Set(), availSum: 0 };
    }
    mbuMap[mbu].totalDtHours += dtHours;
    mbuMap[mbu].incidentCount += 1;
    mbuMap[mbu].siteCount.add(siteCode);

    if (!dailyTimelineMap[dateStr]) {
      dailyTimelineMap[dateStr] = { date: dateStr, totalDtHours: 0, incidentCount: 0, mbus: {} };
    }
    dailyTimelineMap[dateStr].totalDtHours += dtHours;
    dailyTimelineMap[dateStr].incidentCount += 1;
    dailyTimelineMap[dateStr].mbus[mbu] = (dailyTimelineMap[dateStr].mbus[mbu] || 0) + dtHours;

    if (i < 3000) {
      const swRow = siteWiseMap[siteCode] || {};
      const excelAvail = swRow['Total NAR'] !== undefined && swRow['Total NAR'] !== '' ? parseFloat(swRow['Total NAR']) * 100 : 99.0;
      allIncidents.push({
        id: `RSL-${i + 1}`,
        siteId: siteCodeRaw,
        siteName: siteName,
        region: mbu,
        downtimeHours: Number(dtHours.toFixed(2)),
        availability: Number(excelAvail.toFixed(2)),
        timestamp: dateStr,
        category: category,
        status: dtHours > 8 ? 'Active' : 'Resolved',
        slaTarget: 99.90,
        rootCause: reason,
        mttrMinutes: Math.round(dtRaw)
      });
    }
  });

  const allSitesCatalog = Object.values(siteMap).map(s => {
    const topReasonsList = Object.entries(s.reasons)
      .map(([r, h]) => ({ reason: r, hours: Number(h.toFixed(1)) }))
      .sort((a, b) => b.hours - a.hours)
      .slice(0, 3);
    
    const ndRow = narDayMap[s.siteCode.toLowerCase()] || {};
    const dailyTimeline = Object.entries(s.dailyDt).map(([d, h]) => {
      const excelDailyNarVal = getDailyValue(ndRow, d);
      const excelDailyNar = excelDailyNarVal !== undefined ? excelDailyNarVal * 100 : 100;
      return {
        date: d,
        hours: Number(h.toFixed(1)),
        narPercent: Number(excelDailyNar.toFixed(2))
      };
    });

    return {
      siteCode: s.siteCode,
      siteName: s.siteName,
      mbu: s.mbu,
      vendor: s.vendor,
      siteType: s.siteType,
      priority: s.priority,
      totalDtHours: Number(s.totalDtHours.toFixed(1)),
      incidentCount: s.incidentCount,
      availability: s.availability,
      topReasons: topReasonsList,
      dailyTimeline: dailyTimeline.sort((a, b) => a.date.localeCompare(b.date))
    };
  }).sort((a, b) => b.totalDtHours - a.totalDtHours);

  allSitesCatalog.forEach(s => {
    if (mbuMap[s.mbu]) mbuMap[s.mbu].availSum += s.availability;
  });

  const mbuFormatted = Object.entries(mbuMap).map(([mbu, data]) => {
    const avgAvail = data.siteCount.size > 0 ? (data.availSum / data.siteCount.size) : 100;
    return {
      mbu,
      totalDtHours: Number((data.totalDtHours).toFixed(1)),
      incidentCount: data.incidentCount,
      siteCount: data.siteCount.size,
      avgAvailability: Number(avgAvail.toFixed(2))
    };
  }).sort((a, b) => b.totalDtHours - a.totalDtHours);

  const reasonsFormatted = Object.values(reasonMap).map(r => ({
    reason: r.reason,
    category: r.category,
    totalDtHours: Number(r.totalDtHours.toFixed(1)),
    incidentCount: r.incidentCount
  })).sort((a, b) => b.totalDtHours - a.totalDtHours);

  const dailyFormatted = Object.values(dailyTimelineMap).map(d => {
    const dwRow = dateWiseMap[d.date] || {};
    let narPercent = 99.85;
    const mbuAvails = [];
    Object.keys(dwRow).forEach(k => {
      if (k.startsWith('C4-')) {
        const val = parseFloat(dwRow[k]);
        if (!isNaN(val)) mbuAvails.push(val * 100);
      }
    });
    if (mbuAvails.length > 0) {
      narPercent = mbuAvails.reduce((sum, v) => sum + v, 0) / mbuAvails.length;
    }

    return {
      date: d.date,
      totalDtHours: Number(d.totalDtHours.toFixed(1)),
      incidentCount: d.incidentCount,
      narPercent: Number(narPercent.toFixed(2)),
      mbus: d.mbus
    };
  }).sort((a, b) => a.date.localeCompare(b.date));

  const sumAvail = allSitesCatalog.reduce((sum, s) => sum + s.availability, 0);
  const avgAvail = allSitesCatalog.length > 0 ? (sumAvail / allSitesCatalog.length) : 100;
  const totalSitesDowntimeHours = allSitesCatalog.reduce((sum, s) => sum + s.totalDtHours, 0);

  state.parsedPayload = {
    summary: {
      totalRawRecords: deodarRslRows.length,
      totalDowntimeHours: Number(totalSitesDowntimeHours.toFixed(1)),
      totalSites: allSitesCatalog.length,
      avgAvailability: Number(avgAvail.toFixed(2))
    },
    allSites: allSitesCatalog,
    topReasons: reasonsFormatted,
    mbuBreakdown: mbuFormatted,
    dailyTimeline: dailyFormatted,
    sampleIncidents: allIncidents
  };

  // Render metrics to UI
  document.getElementById('stat-sites').textContent = state.parsedPayload.summary.totalSites;
  document.getElementById('stat-nar').textContent = `${state.parsedPayload.summary.avgAvailability}%`;
  document.getElementById('stat-incidents').textContent = state.parsedPayload.summary.totalRawRecords.toLocaleString();
  document.getElementById('stat-downtime').textContent = state.parsedPayload.summary.totalDowntimeHours.toLocaleString();

  log('Compilation complete! Telemetry payload generated successfully.', 'success');
  
  // Enable Publish button
  document.getElementById('publish-btn').classList.remove('disabled');
  document.getElementById('publish-btn').disabled = false;
}

// GitHub API Commit logic
async function publishToGitHub() {
  const { token, owner, repo, path: filePath, branch } = state.settings;
  
  if (!token) {
    log('Publish failed: GitHub Access Token is not set. Save it in Step 1.', 'error');
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
  // Settings load from local storage is disabled since parameters are hardcoded in source.
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
    const emptyMsg = listContainer.querySelector('.empty-queue-msg');
    if (emptyMsg) emptyMsg.remove();

    Array.from(fileList).forEach(file => {
      if (!file.name.endsWith('.csv')) {
        log(`Ignored non-CSV file: ${file.name}`, 'warn');
        return;
      }

      const reader = new FileReader();
      reader.onload = (e) => {
        state.files[file.name] = e.target.result;
        log(`Uploaded and cached file: ${file.name}`);
        
        // Render in queue list
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
          <div class="file-info">
            <span class="file-name">${file.name}</span>
            <span class="file-meta">${(file.size / 1024).toFixed(1)} KB</span>
          </div>
          <button class="remove-file-btn" data-name="${file.name}">×</button>
        `;
        listContainer.appendChild(item);

        // Check if all files loaded to enable Processing
        const filesCount = Object.keys(state.files).length;
        if (filesCount >= 4) {
          document.getElementById('parse-data-btn').classList.remove('disabled');
          document.getElementById('parse-data-btn').disabled = false;
        }
      };
      reader.readAsText(file);
    });
  }

  // Remove File from queue
  document.getElementById('file-list').addEventListener('click', (e) => {
    if (e.target.classList.contains('remove-file-btn')) {
      const fileName = e.target.dataset.name;
      delete state.files[fileName];
      e.target.parentElement.remove();
      log(`Removed file from queue: ${fileName}`);
      
      const filesCount = Object.keys(state.files).length;
      if (filesCount < 4) {
        document.getElementById('parse-data-btn').classList.add('disabled');
        document.getElementById('parse-data-btn').disabled = true;
      }
      if (filesCount === 0) {
        document.getElementById('file-list').innerHTML = '<div class="empty-queue-msg">No files uploaded yet.</div>';
      }
    }
  });

  // Action Buttons
  document.getElementById('parse-data-btn').addEventListener('click', processTelemetryData);
  document.getElementById('publish-btn').addEventListener('click', publishToGitHub);
  document.getElementById('clear-log-btn').addEventListener('click', () => {
    document.getElementById('log-terminal').innerHTML = '<div class="log-line info">[System] Log terminal cleared.</div>';
  });
});
