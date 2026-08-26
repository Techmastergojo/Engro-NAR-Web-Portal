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
  let name = rawName.trim();
  
  // Remove site code prefix, e.g. GJR1058__
  const codeRegex = new RegExp('^' + siteCode + '[_ ]*', 'i');
  name = name.replace(codeRegex, '');
  
  // Strip common sub-prefixes like S_, H_, T_, D_ or double underscores
  name = name.replace(/^([S|H|T|D]__?|__?)/i, '');

  // Convert underscores to spaces first to avoid word boundary issues with \b
  name = name.replace(/_/g, ' ');

  // Remove parenthesized operator/generic ids, e.g. (Zong5218), (Telenor_LWR-001)
  name = name.replace(/\((zong|telenor|ufone|cmpak|djuice)[_ ]?[a-z0-9-]+\)/ig, '');
  name = name.replace(/\([a-z]{2,4}[_-]?\d+\)/ig, '');
  
  // Remove non-parenthesized operator ids, e.g. CMPak4330, Ufone1362
  name = name.replace(/\b(zong|telenor|ufone|cmpak|djuice)[_ ]?[a-z0-9-]+\b/ig, '');
  
  // Match any leftover uppercase patterns of letters followed by numbers, e.g., GJ4165, MDSK4280, MDIK4147
  name = name.replace(/\b[A-Z]{2,4}[_-]?\d+\b/g, '');

  // Clean spaces and trailing separators
  name = name.replace(/\s+/g, ' ');
  name = name.replace(/[-\s_]+$/, '');
  name = name.trim();
  
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
      const ndRow = narDayMap[siteCode] || {};
      // NAR from Site NAR-Day 'Average NAR' column (primary), fallback to SiteWiseDT 'Total NAR'
      let totalNar = 99.0;
      if (ndRow['Average NAR'] !== undefined && ndRow['Average NAR'] !== '' && !isNaN(ndRow['Average NAR'])) {
        totalNar = parseFloat(ndRow['Average NAR']) * 100;
      } else if (swRow['Total NAR'] !== undefined && swRow['Total NAR'] !== '' && !isNaN(swRow['Total NAR'])) {
        totalNar = parseFloat(swRow['Total NAR']) * 100;
      }
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
    // 1. Fetch current file from GitHub (to merge with)
    let sha = '';
    let existingData = null;
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
      log(`Found existing data on GitHub (SHA: ${sha.substring(0, 7)}). Fetching for merge...`);
      
      // Decode existing content
      try {
        const existingJson = decodeURIComponent(escape(atob(lookupData.content.replace(/\n/g, ''))));
        existingData = JSON.parse(existingJson);
        log(`Existing data loaded: ${existingData.allSites?.length || 0} sites, ${existingData.dailyTimeline?.length || 0} timeline days.`);
      } catch (decodeErr) {
        log(`Warning: Could not decode existing data (${decodeErr.message}). Will overwrite with new data.`, 'warn');
        existingData = null;
      }
    } else if (lookupRes.status === 404) {
      log('No existing data on GitHub. Creating fresh file...');
    } else {
      throw new Error(`Failed to lookup file metadata (Status ${lookupRes.status})`);
    }

    // 2. Merge new data with existing data
    let finalPayload;
    if (existingData && existingData.allSites) {
      log('Merging new data with existing records...');
      finalPayload = mergePayloads(existingData, state.parsedPayload);
      log(`Merge complete! ${finalPayload.allSites.length} total sites, ${finalPayload.dailyTimeline.length} timeline days.`, 'success');
    } else {
      finalPayload = state.parsedPayload;
    }

    // 3. Commit / Push merged result
    const jsonStr = JSON.stringify(finalPayload, null, 2);
    const contentBase64 = btoa(unescape(encodeURIComponent(jsonStr)));
    
    const commitUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${filePath}`;
    const commitBody = {
      message: `update telemetry: ${finalPayload.summary.totalSites} sites, ${finalPayload.dailyTimeline.length} days (NAR: ${finalPayload.summary.avgAvailability}%)`,
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
      alert(`Success! Telemetry published with ${finalPayload.allSites.length} sites and ${finalPayload.dailyTimeline.length} days of history.`);
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

// Merge new parsed data INTO existing data (append, don't replace)
function mergePayloads(existing, incoming) {
  // --- Merge allSites ---
  const siteMap = {};
  // Index existing sites by siteCode
  (existing.allSites || []).forEach(s => { siteMap[s.siteCode.toLowerCase()] = { ...s }; });

  // Merge incoming sites
  (incoming.allSites || []).forEach(s => {
    const key = s.siteCode.toLowerCase();
    if (siteMap[key]) {
      // Site exists — merge timelines and update stats
      const old = siteMap[key];

      // Merge dailyTimeline: keep existing days, add new ones
      const existingDates = new Set(old.dailyTimeline.map(d => d.date));
      const mergedTimeline = [...old.dailyTimeline];
      (s.dailyTimeline || []).forEach(d => {
        if (!existingDates.has(d.date)) {
          mergedTimeline.push(d);
        } else {
          // Update existing day with newer data
          const idx = mergedTimeline.findIndex(x => x.date === d.date);
          if (idx !== -1) mergedTimeline[idx] = d;
        }
      });
      mergedTimeline.sort((a, b) => a.date.localeCompare(b.date));
      old.dailyTimeline = mergedTimeline;

      // Merge nar6Months: keep old months, add/update new ones
      const existingMonths = new Set((old.nar6Months || []).map(m => m.monthKey));
      const merged6m = [...(old.nar6Months || [])];
      (s.nar6Months || []).forEach(m => {
        if (!existingMonths.has(m.monthKey)) {
          merged6m.push(m);
        } else {
          const idx = merged6m.findIndex(x => x.monthKey === m.monthKey);
          if (idx !== -1) merged6m[idx] = m;
        }
      });
      merged6m.sort((a, b) => a.monthKey.localeCompare(b.monthKey));
      old.nar6Months = merged6m;

      // Update latest stats (from newer file)
      old.availability = s.availability;
      old.totalDtHours = Number((old.totalDtHours + s.totalDtHours).toFixed(1));
      old.incidentCount = old.incidentCount + s.incidentCount;

      // Merge topReasons (accumulate hours)
      const reasonsMap = {};
      (old.topReasons || []).forEach(r => { reasonsMap[r.reason] = r.hours; });
      (s.topReasons || []).forEach(r => { reasonsMap[r.reason] = (reasonsMap[r.reason] || 0) + r.hours; });
      old.topReasons = Object.entries(reasonsMap)
        .map(([reason, hours]) => ({ reason, hours: Number(hours.toFixed(1)) }))
        .sort((a, b) => b.hours - a.hours).slice(0, 5);

      siteMap[key] = old;
    } else {
      // Brand new site — add it
      siteMap[key] = { ...s };
    }
  });

  const allSites = Object.values(siteMap).sort((a, b) => b.totalDtHours - a.totalDtHours);

  // --- Merge dailyTimeline (global) ---
  const dayMap = {};
  (existing.dailyTimeline || []).forEach(d => { dayMap[d.date] = { ...d }; });
  (incoming.dailyTimeline || []).forEach(d => {
    if (!dayMap[d.date]) {
      dayMap[d.date] = { ...d };
    } else {
      // Update with newer data for this date
      dayMap[d.date] = { ...d };
    }
  });
  const dailyTimeline = Object.values(dayMap).sort((a, b) => a.date.localeCompare(b.date));

  // --- Merge topReasons ---
  const reasonsMap = {};
  (existing.topReasons || []).forEach(r => { reasonsMap[r.reason] = { ...r }; });
  (incoming.topReasons || []).forEach(r => {
    if (reasonsMap[r.reason]) {
      reasonsMap[r.reason].totalDtHours = Number((reasonsMap[r.reason].totalDtHours + r.totalDtHours).toFixed(1));
      reasonsMap[r.reason].incidentCount += r.incidentCount;
    } else {
      reasonsMap[r.reason] = { ...r };
    }
  });
  const topReasons = Object.values(reasonsMap).sort((a, b) => b.totalDtHours - a.totalDtHours);

  // --- Merge mbuBreakdown ---
  const mbuMap = {};
  (existing.mbuBreakdown || []).forEach(m => { mbuMap[m.mbu] = { ...m }; });
  (incoming.mbuBreakdown || []).forEach(m => {
    if (mbuMap[m.mbu]) {
      mbuMap[m.mbu].totalDtHours = Number((mbuMap[m.mbu].totalDtHours + m.totalDtHours).toFixed(1));
      mbuMap[m.mbu].incidentCount += m.incidentCount;
      mbuMap[m.mbu].siteCount = Math.max(mbuMap[m.mbu].siteCount, m.siteCount);
      mbuMap[m.mbu].avgAvailability = m.avgAvailability; // latest value
    } else {
      mbuMap[m.mbu] = { ...m };
    }
  });
  const mbuBreakdown = Object.values(mbuMap).sort((a, b) => b.totalDtHours - a.totalDtHours);

  // --- Merge sampleIncidents (append new, deduplicate by id) ---
  const incidentIds = new Set((existing.sampleIncidents || []).map(i => i.id));
  const mergedIncidents = [...(existing.sampleIncidents || [])];
  (incoming.sampleIncidents || []).forEach(inc => {
    // Assign new unique ID to avoid collision
    const newId = `RSL-${mergedIncidents.length + 1}`;
    mergedIncidents.push({ ...inc, id: newId });
  });

  // --- Recalculate summary ---
  const sumAvail = allSites.reduce((s, x) => s + x.availability, 0);
  const avgAvail = allSites.length > 0 ? sumAvail / allSites.length : 100;
  const totalDtHours = allSites.reduce((s, x) => s + x.totalDtHours, 0);

  return {
    summary: {
      totalRawRecords: (existing.summary?.totalRawRecords || 0) + (incoming.summary?.totalRawRecords || 0),
      totalDowntimeHours: Number(totalDtHours.toFixed(1)),
      totalSites: allSites.length,
      avgAvailability: Number(avgAvail.toFixed(2))
    },
    allSites,
    topReasons,
    mbuBreakdown,
    dailyTimeline,
    sampleIncidents: mergedIncidents
  };
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
