const statusEl    = document.getElementById('status');
const runBtn      = document.getElementById('run');
const runBtnText  = runBtn.querySelector('span');
const formatEl    = document.getElementById('format');
const sessionType = document.getElementById('sessionType');
const searchEl    = document.getElementById('search');
const resultBar   = document.getElementById('resultBar');

/* ---------- helpers ---------- */

function log(msg) {
  const t = new Date().toLocaleTimeString();
  statusEl.textContent = `[${t}] ${msg}\n` + statusEl.textContent;
}

function timestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function download(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ---------- UI wiring ---------- */

formatEl.addEventListener('change', () => {
  if (!runBtn.disabled) runBtnText.textContent = `Export ${formatEl.value.toUpperCase()}`;
});

/* ---------- Report API fields (what we request) ---------- */

const FIELDS = [
  'SessionID', 'Name', 'SessionType',
  'CustomProperty1', 'CustomProperty2', 'CustomProperty3', 'CustomProperty4',
  'CustomProperty5', 'CustomProperty6', 'CustomProperty7', 'CustomProperty8',
  'GuestMachineName', 'GuestMachineDomain', 'GuestMachineDescription',
  'GuestMachineManufacturerName', 'GuestMachineModel',
  'GuestMachineProductNumber', 'GuestMachineSerialNumber',
  'GuestOperatingSystemName', 'GuestOperatingSystemVersion',
  'GuestOperatingSystemManufacturerName', 'GuestOperatingSystemLanguage',
  'GuestProcessorName', 'GuestProcessorVirtualCount', 'GuestProcessorArchitecture',
  'GuestSystemMemoryTotalMegabytes',
  'GuestPrivateNetworkAddress', 'GuestHardwareNetworkAddress',
  'GuestLoggedOnUserName', 'GuestLoggedOnUserDomain', 'GuestIsLocalAdminPresent',
  'GuestLastActivityTime', 'GuestLastBootTime', 'GuestInfoUpdateTime',
  'GuestTimeZoneName', 'GuestTimeZoneOffsetHours',
  'ConnectionCount',
];

/* ---------- CSV parser ---------- */

function parseCSV(text) {
  const rows = [];
  let i = 0;
  const len = text.length;
  // Detect delimiter from first line
  const firstNL = text.indexOf('\n');
  const sample = firstNL > 0 ? text.substring(0, firstNL) : text;
  const tabs = (sample.match(/\t/g) || []).length;
  const commas = (sample.match(/,/g) || []).length;
  const delim = tabs > commas ? '\t' : ',';

  while (i < len) {
    const row = [];
    while (true) {
      let field = '';
      if (i < len && text[i] === '"') {
        i++;
        while (i < len) {
          if (text[i] === '"') {
            if (i + 1 < len && text[i + 1] === '"') { field += '"'; i += 2; }
            else { i++; break; }
          } else { field += text[i++]; }
        }
      } else {
        while (i < len && text[i] !== delim && text[i] !== '\r' && text[i] !== '\n') {
          field += text[i++];
        }
      }
      row.push(field);
      if (i >= len || text[i] === '\r' || text[i] === '\n') break;
      if (text[i] === delim) i++;
    }
    if (i < len && text[i] === '\r') i++;
    if (i < len && text[i] === '\n') i++;
    if (row.length > 1 || row[0] !== '') rows.push(row);
  }
  return rows;
}

function serializeCSV(headers, rows) {
  const esc = (v) => {
    if (v == null) return '';
    const s = String(v);
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.map(esc).join(',')];
  for (const r of rows) lines.push(headers.map(h => esc(r[h] ?? '')).join(','));
  return lines.join('\n');
}

/* ---------- Transformation: State split ---------- */

function splitState(val) {
  if (!val || !val.trim()) return { country: '', state: '' };
  const s = val.trim();
  // "AU - NSW", "AU  - VIC", "NSW" (no dash)
  const m = s.match(/^([A-Z]{2})\s*-\s*(.+)$/i);
  if (m) return { country: m[1].trim(), state: m[2].trim() };
  // Bare state like "NSW"
  return { country: '', state: s };
}

/* ---------- Transformation: OS End of Life ---------- */

function getOSEndOfLife(osName) {
  if (!osName) return '';
  const n = osName.toLowerCase();
  if (n.includes('windows xp'))         return 'Apr 2014 (EOL)';
  if (n.includes('windows vista'))       return 'Apr 2017 (EOL)';
  if (n.includes('windows 7'))           return 'Jan 2020 (EOL)';
  if (n.includes('windows 8.1'))         return 'Jan 2023 (EOL)';
  if (n.includes('windows 8'))           return 'Jan 2016 (EOL)';
  if (n.includes('windows 10'))          return 'Oct 2025';
  if (n.includes('windows 11'))          return 'Supported';
  if (n.includes('server 2008'))         return 'Jan 2020 (EOL)';
  if (n.includes('server 2012'))         return 'Oct 2023 (EOL)';
  if (n.includes('server 2016'))         return 'Jan 2027';
  if (n.includes('server 2019'))         return 'Jan 2029';
  if (n.includes('server 2022'))         return 'Oct 2031';
  if (n.includes('server 2025'))         return 'Oct 2034';
  return '';
}

/* ---------- Transformation: Processor age ---------- */

function getProcessorAge(procName) {
  if (!procName) return '';
  const now = new Date().getFullYear();

  // Intel Core iX-NNNNN (e.g. i7-4770, i7-10700, i5-12400)
  const coreMatch = procName.match(/i[3579]-(\d{3,5})/i);
  if (coreMatch) {
    const model = coreMatch[1];
    let gen;
    if (model.length <= 3) gen = 1;
    else if (model.length === 4) gen = parseInt(model[0]);
    else gen = parseInt(model.substring(0, model.length - 3));
    const years = { 1:2008, 2:2011, 3:2012, 4:2013, 5:2015, 6:2015, 7:2017, 8:2017, 9:2018, 10:2020, 11:2021, 12:2021, 13:2022, 14:2023, 15:2025 };
    const y = years[gen];
    if (y) { const age = now - y; return `Gen ${gen} (~${y}, ${age}yr${age !== 1 ? 's' : ''} old)`; }
  }

  // Intel Celeron J/N series (e.g. J1900, N4000, N5095)
  const celMatch = procName.match(/celeron.*?([JN])(\d)(\d{2,3})/i);
  if (celMatch) {
    const series = celMatch[1].toUpperCase() + celMatch[2];
    const celYears = { J1:2013, J3:2016, J4:2017, N2:2013, N3:2014, N4:2017, N5:2021, N6:2021 };
    const y = celYears[series];
    if (y) { const age = now - y; return `${series}xxx (~${y}, ${age}yr${age !== 1 ? 's' : ''} old)`; }
  }

  // Intel Celeron G series (e.g. G1820, G4900, G5900)
  const celGMatch = procName.match(/celeron.*?G(\d)\d{3}/i);
  if (celGMatch) {
    const s = parseInt(celGMatch[1]);
    const gYears = { 1:2013, 3:2015, 4:2018, 5:2020, 6:2021 };
    const y = gYears[s];
    if (y) { const age = now - y; return `G${s}xxx (~${y}, ${age}yr${age !== 1 ? 's' : ''} old)`; }
  }

  // AMD Ryzen (e.g. Ryzen 5 3600, Ryzen 7 7700)
  const ryzenMatch = procName.match(/ryzen\s+\d\s+(\d)\d{3}/i);
  if (ryzenMatch) {
    const s = parseInt(ryzenMatch[1]);
    const rYears = { 1:2017, 2:2018, 3:2019, 4:2020, 5:2020, 7:2022, 8:2024, 9:2024 };
    const y = rYears[s];
    if (y) { const age = now - y; return `Ryzen ${s}xxx (~${y}, ${age}yr${age !== 1 ? 's' : ''} old)`; }
  }

  return '';
}

/* ---------- Transformation: RAM to GB ---------- */

function mbToGB(val) {
  const mb = parseInt(val);
  if (isNaN(mb) || mb <= 0) return '';
  return `${Math.round(mb / 1024)} GB`;
}

/* ---------- Process CSV into transformed rows ---------- */

const OUTPUT_COLUMNS = [
  'SessionID', 'Name', 'SessionType',
  'CustomProperty1', 'Country', 'State',
  'CustomProperty3', 'CustomProperty4', 'CustomProperty5',
  'Bepoz Version', 'CustomProperty7', 'SystemID',
  'GuestMachineName', 'GuestMachineDomain', 'GuestMachineDescription',
  'GuestMachineManufacturerName', 'GuestMachineModel',
  'GuestMachineProductNumber', 'GuestMachineSerialNumber',
  'GuestOperatingSystemName', 'OS End of Life', 'GuestOperatingSystemVersion',
  'GuestOperatingSystemManufacturerName', 'GuestOperatingSystemLanguage',
  'GuestProcessorName', 'Processor Age', 'GuestProcessorVirtualCount', 'GuestProcessorArchitecture',
  'RAM (GB)',
  'GuestLoggedOnUserName', 'GuestLoggedOnUserDomain', 'GuestIsLocalAdminPresent',
  'GuestLastActivityTime', 'GuestLastBootTime', 'GuestInfoUpdateTime',
  'GuestPrivateNetworkAddress', 'GuestHardwareNetworkAddress',
  'GuestTimeZoneName', 'GuestTimeZoneOffsetHours',
  'ConnectionCount',
];

function processData(csvText) {
  const parsed = parseCSV(csvText);
  if (parsed.length < 2) return { headers: [], rows: [], count: 0 };

  const headers = parsed[0];
  const col = (name) => headers.indexOf(name);

  const rows = [];
  for (let i = 1; i < parsed.length; i++) {
    const raw = parsed[i];
    const get = (name) => { const idx = col(name); return idx >= 0 ? (raw[idx] || '') : ''; };

    const { country, state } = splitState(get('CustomProperty2'));
    const osName = get('GuestOperatingSystemName');
    const proc = get('GuestProcessorName');
    const ramMB = get('GuestSystemMemoryTotalMegabytes');

    const row = {};
    // Pass-through fields
    for (const f of FIELDS) {
      if (f === 'CustomProperty2' || f === 'CustomProperty6' || f === 'CustomProperty8' ||
          f === 'GuestSystemMemoryTotalMegabytes') continue;
      row[f] = get(f);
    }
    // Transformed fields
    row['Country'] = country;
    row['State'] = state;
    row['Bepoz Version'] = get('CustomProperty6');
    row['SystemID'] = get('CustomProperty8');
    row['OS End of Life'] = getOSEndOfLife(osName);
    row['Processor Age'] = getProcessorAge(proc);
    row['RAM (GB)'] = mbToGB(ramMB);

    rows.push(row);
  }

  return { headers: OUTPUT_COLUMNS, rows, count: rows.length };
}

/* ---------- export ---------- */

runBtn.addEventListener('click', async () => {
  const format = formatEl.value;
  const typeFilter = sessionType.value;
  const nameSearch = searchEl.value.trim();

  /* 1. Validate tab */
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url) { log('No active tab found.'); return; }

  const origin = new URL(tab.url).origin;

  /* 2. Build Report API URL — always fetch CSV for processing */
  const params = new URLSearchParams();
  params.set('ReportType', 'Session');
  FIELDS.forEach(f => params.append('SelectFields', f));

  const filters = [];
  if (typeFilter) filters.push(`SessionType='${typeFilter}'`);
  if (nameSearch) {
    const safe = nameSearch.replace(/'/g, "''");
    filters.push(`Name LIKE '%${safe}%'`);
  }
  if (filters.length) params.set('Filter', filters.join(' AND '));
  const reportUrl = `${origin}/Report.csv?${params.toString()}`;

  /* 3. Loading state */
  const desc = [typeFilter, nameSearch ? `"${nameSearch}"` : ''].filter(Boolean).join(' ');
  log(`Fetching${desc ? ` ${desc}` : ''} devices via Report API...`);
  runBtn.disabled = true;
  runBtn.querySelector('svg').outerHTML =
    '<svg class="spinner" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 2a10 10 0 0 1 10 10"/></svg>';
  runBtnText.textContent = 'Fetching...';

  try {
    /* 4. Inject fetch into the page for auth cookies */
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: async (url) => {
        try {
          const r = await fetch(url, { credentials: 'include' });
          if (!r.ok) return { ok: false, error: `HTTP ${r.status}: ${r.statusText}` };
          const ct = r.headers.get('content-type') || '';
          return { ok: true, data: await r.text(), contentType: ct };
        } catch (e) {
          return { ok: false, error: e.message };
        }
      },
      args: [reportUrl],
      world: 'MAIN',
    });

    /* 5. Handle errors */
    if (!result?.ok) {
      const err = result?.error || 'Unknown error';
      if (/404|Not Found/i.test(err)) {
        log('Report API not found. Is the Report Manager extension installed in ScreenConnect?');
      } else if (/401|403|Unauthorized|Forbidden/i.test(err)) {
        log('Not authenticated. Please log into ScreenConnect first.');
      } else {
        log(`Error: ${err}`);
      }
      return;
    }

    if (result.contentType?.includes('text/html')) {
      log('Received HTML instead of data. You may need to log in, or Report Manager may not be installed.');
      return;
    }

    /* 6. If filter returned 0 rows, retry without it */
    let csvText = result.data;
    if (filters.length) {
      const lines = csvText.trim().split('\n');
      if (lines.length <= 1) {
        log('Filter returned no results. Retrying without filter...');
        params.delete('Filter');
        const retryUrl = `${origin}/Report.csv?${params.toString()}`;
        const [{ result: r2 }] = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: async (url) => {
            try {
              const r = await fetch(url, { credentials: 'include' });
              if (!r.ok) return { ok: false };
              return { ok: true, data: await r.text(), contentType: r.headers.get('content-type') || '' };
            } catch (e) { return { ok: false }; }
          },
          args: [retryUrl],
          world: 'MAIN',
        });
        if (r2?.ok && !r2.contentType?.includes('text/html')) {
          csvText = r2.data;
          log('Filter not supported — exporting all sessions.');
        }
      }
    }

    /* 7. Process and transform */
    const { headers, rows, count } = processData(csvText);

    resultBar.textContent = `${count} device${count === 1 ? '' : 's'} exported`;
    resultBar.style.display = 'block';

    /* 8. Download in chosen format */
    if (format === 'json') {
      const json = JSON.stringify(rows, null, 2);
      download(json, `screenconnect-devices-${timestamp()}.json`, 'application/json');
    } else {
      const csv = serializeCSV(headers, rows);
      download(csv, `screenconnect-devices-${timestamp()}.csv`, 'text/csv;charset=utf-8');
    }
    log(`Exported ${count} device(s) as ${format.toUpperCase()}.`);

  } catch (e) {
    log('Error: ' + (e?.message || e));
  } finally {
    runBtn.disabled = false;
    runBtn.querySelector('svg').outerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
    runBtnText.textContent = `Export ${formatEl.value.toUpperCase()}`;
  }
});
