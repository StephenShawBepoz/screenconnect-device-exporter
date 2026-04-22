const statusEl    = document.getElementById('status');
const runBtn      = document.getElementById('run');
const runBtnText  = runBtn.querySelector('span');
const formatEl    = document.getElementById('format');
const sessionType = document.getElementById('sessionType');
const searchEl    = document.getElementById('search');

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

/* ---------- Report API fields ---------- */

const FIELDS = [
  // Identity
  'SessionID', 'Name', 'SessionType',
  // Custom properties (1-8)
  'CustomProperty1', 'CustomProperty2', 'CustomProperty3', 'CustomProperty4',
  'CustomProperty5', 'CustomProperty6', 'CustomProperty7', 'CustomProperty8',
  // Machine
  'GuestMachineName', 'GuestMachineDomain', 'GuestMachineDescription',
  'GuestMachineManufacturerName', 'GuestMachineModel',
  'GuestMachineProductNumber', 'GuestMachineSerialNumber',
  // Operating system
  'GuestOperatingSystemName', 'GuestOperatingSystemVersion',
  'GuestOperatingSystemManufacturerName', 'GuestOperatingSystemLanguage',
  'GuestOperatingSystemInstallationTime',
  // Processor & memory
  'GuestProcessorName', 'GuestProcessorVirtualCount', 'GuestProcessorArchitecture',
  'GuestSystemMemoryTotalMegabytes', 'GuestSystemMemoryAvailableMegabytes',
  // Network
  'GuestPrivateNetworkAddress', 'GuestHardwareNetworkAddress',
  // User
  'GuestLoggedOnUserName', 'GuestLoggedOnUserDomain', 'GuestIsLocalAdminPresent',
  // Timestamps
  'GuestLastActivityTime', 'GuestLastBootTime', 'GuestInfoUpdateTime',
  // Timezone
  'GuestTimeZoneName', 'GuestTimeZoneOffsetHours',
  // Connection
  'ConnectionCount',
];

/* ---------- export ---------- */

runBtn.addEventListener('click', async () => {
  const format = formatEl.value;
  const typeFilter = sessionType.value;
  const nameSearch = searchEl.value.trim();

  /* 1. Validate tab */
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url) { log('No active tab found.'); return; }

  const origin = new URL(tab.url).origin;

  /* 2. Build Report API URL with filters */
  const endpoint = format === 'json' ? 'Report.json' : 'Report.csv';
  const params = new URLSearchParams();
  params.set('ReportType', 'Session');
  FIELDS.forEach(f => params.append('SelectFields', f));

  const filters = [];
  if (typeFilter) filters.push(`SessionType='${typeFilter}'`);
  if (nameSearch) {
    // Escape single quotes in the search term
    const safe = nameSearch.replace(/'/g, "''");
    filters.push(`Name LIKE '%${safe}%'`);
  }
  if (filters.length) params.set('Filter', filters.join(' AND '));
  const reportUrl = `${origin}/${endpoint}?${params.toString()}`;

  /* 3. Loading state */
  const desc = [typeFilter, nameSearch ? `"${nameSearch}"` : ''].filter(Boolean).join(' ');
  log(`Fetching${desc ? ` ${desc}` : ''} devices via Report API...`);
  runBtn.disabled = true;
  runBtn.querySelector('svg').outerHTML =
    '<svg class="spinner" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M12 2a10 10 0 0 1 10 10"/></svg>';
  runBtnText.textContent = 'Fetching...';

  try {
    /* 4. Inject fetch into the page so it uses the existing auth cookies */
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
      log('Received an HTML page instead of data. You may need to log in, or the Report Manager extension may not be installed.');
      return;
    }

    /* 6. If the session-type filter failed (returned 0 rows), retry without it */
    if (typeFilter && format === 'csv') {
      const lines = result.data.trim().split('\n');
      if (lines.length <= 1) {
        log(`No results for SessionType='${typeFilter}'. Retrying without filter...`);
        params.delete('Filter');
        const retryUrl = `${origin}/${endpoint}?${params.toString()}`;
        const [{ result: r2 }] = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: async (url) => {
            try {
              const r = await fetch(url, { credentials: 'include' });
              if (!r.ok) return { ok: false, error: `HTTP ${r.status}` };
              return { ok: true, data: await r.text(), contentType: r.headers.get('content-type') || '' };
            } catch (e) { return { ok: false, error: e.message }; }
          },
          args: [retryUrl],
          world: 'MAIN',
        });
        if (r2?.ok && !r2.contentType?.includes('text/html')) {
          result.data = r2.data;
          log('Filter not supported on this instance — exporting all sessions.');
        }
      }
    }

    /* 7. Count records */
    let count = '?';
    if (format === 'csv') {
      count = Math.max(0, result.data.trim().split('\n').length - 1);
    } else {
      try { const p = JSON.parse(result.data); count = Array.isArray(p) ? p.length : '?'; } catch {}
    }

    /* 8. Download */
    const ext  = format === 'json' ? 'json' : 'csv';
    const mime = format === 'json' ? 'application/json' : 'text/csv;charset=utf-8';
    download(result.data, `screenconnect-devices-${timestamp()}.${ext}`, mime);
    log(`Exported ${count} device(s) as ${format.toUpperCase()}.`);

  } catch (e) {
    log('Error: ' + (e?.message || e));
  } finally {
    /* 9. Restore button */
    runBtn.disabled = false;
    runBtn.querySelector('svg').outerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
    runBtnText.textContent = `Export ${formatEl.value.toUpperCase()}`;
  }
});
