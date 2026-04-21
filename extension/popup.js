const statusEl = document.getElementById('status');
const runBtn   = document.getElementById('run');
const stopBtn  = document.getElementById('stop');
const progressContainer = document.getElementById('progressContainer');
const progressFill  = document.getElementById('progressFill');
const progressCount = document.getElementById('progressCount');
const progressLabel = document.querySelector('.progress-label');

function log(msg) {
  const t = new Date().toLocaleTimeString();
  statusEl.textContent = `[${t}] ${msg}\n` + statusEl.textContent;
}

function showProgress(current, total) {
  progressContainer.classList.add('active');
  progressCount.textContent = `${current} / ${total}`;
  progressFill.style.width = total > 0 ? `${Math.round((current / total) * 100)}%` : '0%';
}

function hideProgress() {
  progressContainer.classList.remove('active');
  progressFill.style.width = '0%';
}

let pollTimer = null;

async function pollProgress(tabId) {
  try {
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId },
      func: () => window.__scProgress,
      world: 'MAIN'
    });
    if (result) {
      showProgress(result.current, result.total);
      if (result.name) {
        log(`${result.current}/${result.total}: ${result.name}`);
      }
    }
  } catch (e) { /* ignore poll errors */ }
}

runBtn.addEventListener('click', async () => {
  const delay = parseInt(document.getElementById('delay').value, 10) || 600;
  const format = document.getElementById('format').value;

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !/screenconnect\.com/.test(tab.url || '')) {
    log('Not on a ScreenConnect tab.');
    return;
  }

  log('Scanning devices...');
  if (progressLabel) progressLabel.textContent = 'Scanning devices...';
  showProgress(0, 0);

  // Poll for progress every 800ms
  pollTimer = setInterval(() => pollProgress(tab.id), 800);

  try {
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: collectDeep,
      args: [delay],
      world: 'MAIN'
    });
    clearInterval(pollTimer);
    hideProgress();
    if (!result || !result.rows || !result.rows.length) {
      log('No rows collected.');
      return;
    }
    if (format === 'json') {
      const json = JSON.stringify(result.rows, null, 2);
      download(json, `screenconnect-devices-${Date.now()}.json`, 'application/json');
    } else {
      const csv = toCSV(result.rows, result.columns);
      download(csv, `screenconnect-devices-${Date.now()}.csv`, 'text/csv;charset=utf-8');
    }
    log(`Done! ${result.rows.length} device(s) exported as ${format.toUpperCase()}.`);
  } catch (e) {
    clearInterval(pollTimer);
    hideProgress();
    log('Error: ' + (e?.message || e));
  }
});

stopBtn.addEventListener('click', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: () => { window.__scStop = true; },
    world: 'MAIN'
  });
  log('Stop signal sent.');
});

function toCSV(rows, columns) {
  const escape = (v) => {
    if (v == null) return '';
    const s = String(v).replace(/\r?\n/g, ' ').trim();
    return /[",]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const header = columns.map(escape).join(',');
  const body = rows.map(r => columns.map(c => escape(r[c])).join(',')).join('\n');
  return header + '\n' + body;
}

function download(text, filename, mime = 'text/csv;charset=utf-8') {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ====== Injected into MAIN world — fully self-contained ======
async function collectDeep(delay) {
  window.__scStop = false;
  window.__scProgress = { current: 0, total: 0, name: '' };
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const getRows = () => Array.from(document.querySelectorAll('table.DetailTable tbody tr'));

  // --- Inline row parser ---
  const parseRowFields = (tr) => {
    const name = tr.querySelector('h3.SessionTitle')?.textContent.trim() || '';
    const pTags = tr.querySelectorAll('.SessionInfoPanel p');
    const raw = {};
    pTags.forEach(p => {
      const title = (p.getAttribute('title') || '').trim();
      if (title) raw[title] = title;
    });
    const getField = (prefix) => {
      for (const key of Object.keys(raw)) {
        if (key.startsWith(prefix)) {
          let val = key.substring(prefix.length).trim();
          if (val.startsWith(':')) val = val.substring(1).trim();
          return val;
        }
      }
      return '';
    };
    const guest = tr.querySelector('.StatusDiagramPanel .Guest');
    const online = guest ? (guest.classList.contains('Connected') ? 'Online' : 'Offline') : '';
    const dbRaw = getField('Database Details');
    const dbParts = dbRaw.split('|').map(s => s.trim());
    const userRaw = getField('User');
    let user = userRaw, idle = '';
    const idleMatch = userRaw.match(/^(.+?)\s*\(Idle\s+(.+?)\)$/);
    if (idleMatch) { user = idleMatch[1].trim(); idle = idleMatch[2].trim(); }
    return {
      'Name': name, 'Online/Offline': online,
      'GROUP': getField('GROUP'), 'STATE': getField('STATE'),
      'Business Unit': getField('Business Unit'), 'Device Type': getField('Device Type'),
      'Manufacturer': dbParts[0] || '', 'Model': dbParts[1] || '',
      'OS': dbParts[2] || '', '.NET Version': dbParts[3] || '',
      'POS Version': getField('POS Version'),
      'Access Restriction Level': getField('Access Restriction Level'),
      'SystemID': getField('SystemID'), 'User': user, 'Idle Time': idle
    };
  };

  // --- Detail panel helpers ---
  const getPanelFingerprint = () => {
    const dds = document.querySelectorAll('.DetailTabContent .CollapsiblePanel .Content dl dd');
    return Array.from(dds).map(dd => dd.textContent.trim()).join('|');
  };

  const getPanelSessionName = () => {
    const dls = document.querySelectorAll('.DetailTabContent .CollapsiblePanel .Content dl');
    for (const dl of dls) {
      const dts = dl.querySelectorAll('dt');
      const dds = dl.querySelectorAll('dd');
      for (let i = 0; i < dts.length; i++) {
        if (dts[i].textContent.trim().replace(/:$/, '') === 'Name')
          return (dds[i]?.textContent || '').trim();
      }
    }
    return '';
  };

  const waitForPanel = async (expectedName, prevFP, timeoutMs = 10000) => {
    const start = Date.now();
    // Phase 1: wait for panel Name to match the device we clicked
    while (Date.now() - start < timeoutMs) {
      const loading = document.querySelector('.QueuedGuestInfoActivityIndicator')?.offsetHeight > 0;
      if (!loading) {
        const name = getPanelSessionName();
        if (name === expectedName) {
          // Name matches — give a brief moment for remaining fields to populate
          await sleep(300);
          return true;
        }
      }
      await sleep(150);
    }
    // Phase 2: if name never matched (truncated name?), accept fingerprint change
    const fp = getPanelFingerprint();
    if (prevFP && fp !== prevFP) {
      await sleep(500);
      return true;
    }
    return false;
  };

  const parsePanel = () => {
    const data = {};
    document.querySelectorAll('.DetailTabContent .CollapsiblePanel').forEach(p => {
      const section = p.querySelector('.Header')?.textContent.trim() || 'Other';
      const dl = p.querySelector('.Content dl');
      if (!dl) return;
      const dts = dl.querySelectorAll('dt');
      const dds = dl.querySelectorAll('dd');
      for (let i = 0; i < dts.length; i++) {
        const k = dts[i].textContent.trim().replace(/:$/, '');
        data[`${section}.${k}`] = (dds[i]?.textContent || '').trim();
      }
    });
    return data;
  };

  const parseMemoryToGB = (s) => {
    if (!s) return '';
    const m = s.match(/(\d+(?:\.\d+)?)\s*MB\s*\/\s*(\d+(?:\.\d+)?)\s*MB/i);
    if (m) return Math.ceil(parseFloat(m[2]) / 1024) + ' GB';
    const gbMatch = s.match(/(\d+(?:\.\d+)?)\s*GB/i);
    if (gbMatch) return Math.ceil(parseFloat(gbMatch[1])) + ' GB';
    const mbMatch = s.match(/(\d+(?:\.\d+)?)\s*MB/i);
    if (mbMatch) return Math.ceil(parseFloat(mbMatch[1]) / 1024) + ' GB';
    return s;
  };

  // --- Main loop ---
  const rows = getRows();
  const total = rows.length;
  window.__scProgress = { current: 0, total, name: 'Starting...' };
  const results = [];
  const seen = new Set();

  for (let i = 0; i < total; i++) {
    if (window.__scStop) break;
    const tr = getRows()[i];
    if (!tr) continue;

    const rec = parseRowFields(tr);
    if (!rec['Name'] || seen.has(rec['Name'])) continue;
    seen.add(rec['Name']);

    // Select row via full mouse event sequence on the <tr>
    const prevFP = getPanelFingerprint();
    tr.scrollIntoView({ block: 'center' });
    const evtOpts = { bubbles: true, cancelable: true, view: window };
    tr.dispatchEvent(new MouseEvent('mousedown', evtOpts));
    tr.dispatchEvent(new MouseEvent('mouseup', evtOpts));
    tr.click();
    await sleep(300);
    const loaded = await waitForPanel(rec['Name'], prevFP);
    await sleep(delay);

    if (loaded) {
      const panel = parsePanel();
      rec['Machine'] = panel['Device.Machine'] || '';
      rec['Processor'] = panel['Device.Processor(s)'] || '';
      rec['RAM'] = parseMemoryToGB(panel['Device.Available Memory'] || '');
      rec['RAM (raw)'] = panel['Device.Available Memory'] || '';
      rec['Manufacturer & Model'] = panel['Device.Manufacturer & Model'] || '';
      rec['OS (full)'] = panel['Device.Operating System'] || '';
      rec['OS Installed'] = panel['Device.Operating System Installation'] || '';
      rec['Machine Product/Serial'] = panel['Device.Machine Product/Serial'] || '';
    }

    results.push(rec);
    window.__scProgress = { current: results.length, total, name: rec['Name'] };
    console.log(`[SC Exporter] ${results.length}/${total}: ${rec['Name']}`);
  }

  const columns = [
    'Name', 'Online/Offline', 'GROUP', 'STATE', 'Business Unit',
    'Device Type', 'Machine', 'Manufacturer & Model', 'OS (full)', 'Processor', 'RAM',
    'Manufacturer', 'Model', 'OS', '.NET Version',
    'POS Version', 'Access Restriction Level', 'SystemID', 'User', 'Idle Time',
    'OS Installed', 'RAM (raw)', 'Machine Product/Serial'
  ];
  return { rows: results, columns };
}
