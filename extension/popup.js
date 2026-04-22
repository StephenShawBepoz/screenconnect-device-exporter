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
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
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
  'GuestOperatingSystemName',
  'GuestProcessorName', 'GuestProcessorVirtualCount',
  'GuestSystemMemoryTotalMegabytes',
  'GuestPrivateNetworkAddress', 'GuestHardwareNetworkAddress',
  'GuestLoggedOnUserName', 'GuestLoggedOnUserDomain',
  'GuestLastActivityTime', 'GuestLastBootTime', 'GuestInfoUpdateTime',
  'GuestTimeZoneName',
  'ConnectionCount',
];

/* ---------- CSV parser ---------- */

function parseCSV(text) {
  const rows = [];
  let i = 0;
  const len = text.length;
  const firstNL = text.indexOf('\n');
  const sample = firstNL > 0 ? text.substring(0, firstNL) : text;
  const delim = (sample.match(/\t/g) || []).length > (sample.match(/,/g) || []).length ? '\t' : ',';

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

/* ---------- Transformations ---------- */

function splitState(val) {
  if (!val || !val.trim()) return { country: '', state: '' };
  const s = val.trim();
  const m = s.match(/^([A-Z]{2})\s*-\s*(.+)$/i);
  if (m) return { country: m[1].trim(), state: m[2].trim() };
  return { country: '', state: s };
}

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

function getProcessorAge(procName) {
  if (!procName) return '';
  const now = new Date().getFullYear();
  const fmt = (label, year) => { const age = now - year; return `${label} (~${year}, ${age}yr${age !== 1 ? 's' : ''} old)`; };
  const genYears = { 1:2008, 2:2011, 3:2012, 4:2013, 5:2015, 6:2015, 7:2017, 8:2017, 9:2018, 10:2020, 11:2021, 12:2021, 13:2022, 14:2023, 15:2025 };

  // 1. Try explicit "Nth Gen" text (e.g. "12th Gen Intel(R) Core(TM) i7-1260P")
  const genText = procName.match(/(\d{1,2})(?:st|nd|rd|th)\s+Gen/i);
  if (genText) {
    const gen = parseInt(genText[1]);
    if (genYears[gen]) return fmt(`Gen ${gen}`, genYears[gen]);
  }

  // 2. Intel Core iX-NNNNN from model number
  const coreMatch = procName.match(/i[3579]-(\d{3,5})/i);
  if (coreMatch) {
    const model = coreMatch[1];
    let gen;
    if (model.length <= 3) gen = 1;
    else if (model.length === 4) gen = parseInt(model[0]);
    else gen = parseInt(model.substring(0, model.length - 3));
    if (genYears[gen]) return fmt(`Gen ${gen}`, genYears[gen]);
  }

  // 3. Intel Xeon E-2xxx (e.g. E-2224G, E-2314)
  const xeonMatch = procName.match(/xeon.*?e-(\d)(\d)/i);
  if (xeonMatch) {
    const sub = parseInt(xeonMatch[2]);
    const xeonYears = { 1: 2018, 2: 2019, 3: 2021, 4: 2023 };
    if (xeonYears[sub]) return fmt(`Xeon E-${xeonMatch[1]}${sub}xx`, xeonYears[sub]);
  }

  // 4. Intel Celeron J/N series (e.g. J1900, N4000, N5095)
  const celMatch = procName.match(/celeron.*?([JN])(\d)/i);
  if (celMatch) {
    const key = celMatch[1].toUpperCase() + celMatch[2];
    const celYears = { J1:2013, J3:2016, J4:2017, N2:2013, N3:2014, N4:2017, N5:2021, N6:2021 };
    if (celYears[key]) return fmt(`Celeron ${key}xxx`, celYears[key]);
  }

  // 5. Intel Celeron/Pentium G series (e.g. G4900, G5400, G7400)
  const gMatch = procName.match(/(?:celeron|pentium).*?G(\d)\d{3}/i);
  if (gMatch) {
    const s = parseInt(gMatch[1]);
    const gYears = { 1:2013, 3:2015, 4:2018, 5:2020, 6:2021, 7:2022 };
    if (gYears[s]) return fmt(`G${s}xxx`, gYears[s]);
  }

  // 6. AMD Ryzen (e.g. Ryzen 5 3600, Ryzen 7 7700X)
  const ryzenMatch = procName.match(/ryzen\s+\d\s+(\d)\d{3}/i);
  if (ryzenMatch) {
    const s = parseInt(ryzenMatch[1]);
    const rYears = { 1:2017, 2:2018, 3:2019, 4:2020, 5:2020, 7:2022, 8:2024, 9:2024 };
    if (rYears[s]) return fmt(`Ryzen ${s}xxx`, rYears[s]);
  }

  return '';
}

function mbToGB(val) {
  const mb = parseInt(val);
  if (isNaN(mb) || mb <= 0) return '';
  return `${Math.round(mb / 1024)} GB`;
}

/* ---------- Output columns: clean names in BDM-friendly order ---------- */

const OUTPUT_COLUMNS = [
  'Device Name',
  'Client',
  'Country',
  'State',
  'Region',
  'Device Type',
  'Bepoz Version',
  'System ID',
  'Machine Name',
  'Domain',
  'Manufacturer',
  'Model',
  'Serial Number',
  'Product Number',
  'Operating System',
  'OS End of Life',
  'Processor',
  'Processor Age',
  'CPU Cores',
  'RAM',
  'Logged-in User',
  'User Domain',
  'Last Activity',
  'Last Boot',
  'Last Updated',
  'Private IP',
  'MAC Address',
  'Timezone',
  'Connections',
  'Session ID',
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

    rows.push({
      'Device Name':      get('Name'),
      'Client':           get('CustomProperty1'),
      'Country':          country,
      'State':            state,
      'Region':           get('CustomProperty3'),
      'Device Type':      get('CustomProperty4'),
      'Bepoz Version':    get('CustomProperty6'),
      'System ID':        get('CustomProperty8'),
      'Machine Name':     get('GuestMachineName'),
      'Domain':           get('GuestMachineDomain'),
      'Manufacturer':     get('GuestMachineManufacturerName'),
      'Model':            get('GuestMachineModel'),
      'Serial Number':    get('GuestMachineSerialNumber'),
      'Product Number':   get('GuestMachineProductNumber'),
      'Operating System': get('GuestOperatingSystemName'),
      'OS End of Life':   getOSEndOfLife(get('GuestOperatingSystemName')),
      'Processor':        get('GuestProcessorName'),
      'Processor Age':    getProcessorAge(get('GuestProcessorName')),
      'CPU Cores':        get('GuestProcessorVirtualCount'),
      'RAM':              mbToGB(get('GuestSystemMemoryTotalMegabytes')),
      'Logged-in User':   get('GuestLoggedOnUserName'),
      'User Domain':      get('GuestLoggedOnUserDomain'),
      'Last Activity':    get('GuestLastActivityTime'),
      'Last Boot':        get('GuestLastBootTime'),
      'Last Updated':     get('GuestInfoUpdateTime'),
      'Private IP':       get('GuestPrivateNetworkAddress'),
      'MAC Address':      get('GuestHardwareNetworkAddress'),
      'Timezone':         get('GuestTimeZoneName'),
      'Connections':      get('ConnectionCount'),
      'Session ID':       get('SessionID'),
    });
  }

  return { headers: OUTPUT_COLUMNS, rows, count: rows.length };
}

/* ---------- export ---------- */

runBtn.addEventListener('click', async () => {
  const format = formatEl.value;
  const typeFilter = sessionType.value;
  const nameSearch = searchEl.value.trim();

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url) { log('No active tab found.'); return; }

  const origin = new URL(tab.url).origin;

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

  const desc = [typeFilter, nameSearch ? `"${nameSearch}"` : ''].filter(Boolean).join(' ');
  log(`Fetching${desc ? ` ${desc}` : ''} devices...`);
  runBtn.disabled = true;
  runBtn.querySelector('svg').outerHTML =
    '<svg class="spinner" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 2a10 10 0 0 1 10 10"/></svg>';
  runBtnText.textContent = 'Fetching...';

  try {
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: async (url) => {
        try {
          const r = await fetch(url, { credentials: 'include' });
          if (!r.ok) return { ok: false, error: `HTTP ${r.status}: ${r.statusText}` };
          const ct = r.headers.get('content-type') || '';
          return { ok: true, data: await r.text(), contentType: ct };
        } catch (e) { return { ok: false, error: e.message }; }
      },
      args: [reportUrl],
      world: 'MAIN',
    });

    if (!result?.ok) {
      const err = result?.error || 'Unknown error';
      if (/404|Not Found/i.test(err)) log('Report API not found. Is Report Manager installed?');
      else if (/401|403|Unauthorized|Forbidden/i.test(err)) log('Not authenticated. Please log in first.');
      else log(`Error: ${err}`);
      return;
    }

    if (result.contentType?.includes('text/html')) {
      log('Received HTML instead of data. Log in or check Report Manager is installed.');
      return;
    }

    let csvText = result.data;
    if (filters.length && csvText.trim().split('\n').length <= 1) {
      log('Filter returned no results. Retrying without filter...');
      params.delete('Filter');
      const [{ result: r2 }] = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: async (url) => {
          try { const r = await fetch(url, { credentials: 'include' }); if (!r.ok) return { ok: false }; return { ok: true, data: await r.text(), contentType: r.headers.get('content-type') || '' }; }
          catch (e) { return { ok: false }; }
        },
        args: [`${origin}/Report.csv?${params.toString()}`],
        world: 'MAIN',
      });
      if (r2?.ok && !r2.contentType?.includes('text/html')) {
        csvText = r2.data;
        log('Filter not supported — exporting all sessions.');
      }
    }

    const { headers, rows, count } = processData(csvText);

    resultBar.textContent = `${count} device${count === 1 ? '' : 's'} exported`;
    resultBar.style.display = 'block';

    if (format === 'json') {
      download(JSON.stringify(rows, null, 2), `screenconnect-devices-${timestamp()}.json`, 'application/json');
    } else {
      download(serializeCSV(headers, rows), `screenconnect-devices-${timestamp()}.csv`, 'text/csv;charset=utf-8');
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
