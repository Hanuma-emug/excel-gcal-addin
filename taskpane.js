/* ─────────────────────────────────────────────
   GCal Reminders — Excel Online Add-in
   taskpane.js
───────────────────────────────────────────── */

'use strict';

/* ── State ── */
let accessToken   = null;
let userEmail     = '';
let sheetHeaders  = [];
let sheetRows     = [];       // array of objects keyed by header
let previewEvents = [];       // built events ready to push

const NONE = '__none__';

/* ── Office ready ── */
Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    document.body.innerHTML = '<p style="padding:16px;color:#a4262c">This add-in only works in Excel.</p>';
    return;
  }
  bindButtons();
});

/* ─────────────────────────────────────────────
   STEP 1 — Connect Google
───────────────────────────────────────────── */
function bindButtons () {
  document.getElementById('btn-connect').addEventListener('click', connectGoogle);
  document.getElementById('btn-read').addEventListener('click', readSheet);
  document.getElementById('btn-to-map').addEventListener('click', goToMap);
  document.getElementById('btn-to-preview').addEventListener('click', goToPreview);
  document.getElementById('btn-push').addEventListener('click', pushEvents);
  document.getElementById('btn-back').addEventListener('click', function () { showPanel(3); });
}

function connectGoogle () {
  const clientId = document.getElementById('client-id-input').value.trim();
  if (!clientId) {
    setStatus('connect-status', 'Please enter your Google Client ID.', 'error');
    return;
  }
  localStorage.setItem('gcal_client_id', clientId);

  /* Open Office dialog for Google OAuth — avoids iframe popup restrictions */
  const dialogUrl = window.location.origin + '/excel-gcal-addin/auth-dialog.html'
    + '?client_id=' + encodeURIComponent(clientId);

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 60, width: 40, promptBeforeOpen: false },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        setStatus('connect-status', 'Could not open the sign-in window. Check your Client ID.', 'error');
        return;
      }
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        try {
          const msg = JSON.parse(arg.message);
          if (msg.type === 'token') {
            accessToken = msg.token;
            userEmail   = msg.email || 'your Google account';
            dialog.close();
            document.getElementById('connected-email').textContent = 'Connected: ' + userEmail;
            markDone(1);
            showPanel(2);
          } else if (msg.type === 'error') {
            dialog.close();
            setStatus('connect-status', 'Sign-in failed: ' + msg.message, 'error');
          }
        } catch (e) {
          dialog.close();
          setStatus('connect-status', 'Unexpected error during sign-in.', 'error');
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
        if (arg.error === 12006) {
          setStatus('connect-status', 'Sign-in window was closed. Please try again.', 'error');
        }
      });
    }
  );
}

/* ─────────────────────────────────────────────
   STEP 2 — Read active sheet
───────────────────────────────────────────── */
function readSheet () {
  setStatus('read-status', 'Reading sheet...', '');
  document.getElementById('data-preview').classList.add('hidden');

  Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load(['values', 'rowCount', 'columnCount']);

    return context.sync().then(function () {
      const values = range.values;
      if (!values || values.length < 2) {
        setStatus('read-status', 'Sheet appears empty or has no data rows.', 'error');
        return;
      }

      /* First row = headers */
      sheetHeaders = values[0].map(function (h) { return String(h).trim(); });

      /* Remaining rows = data, convert to objects */
      sheetRows = [];
      for (let r = 1; r < values.length; r++) {
        const obj = {};
        sheetHeaders.forEach(function (h, i) {
          obj[h] = values[r][i] !== undefined && values[r][i] !== null ? String(values[r][i]).trim() : '';
        });
        sheetRows.push(obj);
      }

      /* Update preview */
      document.getElementById('row-count').textContent = sheetRows.length;

      const colsEl = document.getElementById('preview-cols');
      colsEl.innerHTML = '';
      sheetHeaders.forEach(function (h) {
        const chip = document.createElement('span');
        chip.className = 'col-chip';
        chip.textContent = h;
        colsEl.appendChild(chip);
      });

      document.getElementById('data-preview').classList.remove('hidden');
      setStatus('read-status', '', '');
    });
  }).catch(function (err) {
    setStatus('read-status', 'Error reading sheet: ' + err.message, 'error');
  });
}

/* ─────────────────────────────────────────────
   STEP 3 — Column mapping
───────────────────────────────────────────── */
function goToMap () {
  const selectors = ['map-date', 'map-company', 'map-contact', 'map-phone', 'map-notes'];
  const labels    = ['-- Select date column --', '-- None --', '-- None --', '-- None --', '-- None --'];

  selectors.forEach(function (id, idx) {
    const sel = document.getElementById(id);
    sel.innerHTML = '';

    /* Add a blank/none option for optional fields */
    if (idx > 0) {
      const opt = document.createElement('option');
      opt.value = NONE;
      opt.textContent = labels[idx];
      sel.appendChild(opt);
    }

    sheetHeaders.forEach(function (h) {
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = h;
      sel.appendChild(opt);
    });

    /* Auto-detect common column names */
    const autoMap = {
      'map-date':    ['next follow', 'follow up', 'follow-up', 'next connection', 'date', 'due'],
      'map-company': ['company', 'organization', 'org', 'firm', 'business'],
      'map-contact': ['contact', 'name', 'person', 'client', 'full name'],
      'map-phone':   ['phone', 'telephone', 'mobile', 'tel', 'cell', 'number'],
      'map-notes':   ['notes', 'note', 'status', 'remarks', 'comment']
    };

    const candidates = autoMap[id] || [];
    for (const cand of candidates) {
      const match = sheetHeaders.find(function (h) {
        return h.toLowerCase().replace(/[\s\-_]+/g, ' ').includes(cand);
      });
      if (match) { sel.value = match; break; }
    }
  });

  /* Pre-fill title template using detected columns */
  const contactCol = document.getElementById('map-contact').value;
  const companyCol = document.getElementById('map-company').value;
  let titleVal = 'Follow up';
  if (contactCol && contactCol !== NONE) titleVal += ': {' + contactCol + '}';
  if (companyCol && companyCol !== NONE) titleVal += ' @ {' + companyCol + '}';
  document.getElementById('title-template').value = titleVal;

  markDone(2);
  showPanel(3);
}

/* ─────────────────────────────────────────────
   STEP 4 — Preview events
───────────────────────────────────────────── */
function goToPreview () {
  const dateCol     = document.getElementById('map-date').value;
  const companyCol  = document.getElementById('map-company').value;
  const contactCol  = document.getElementById('map-contact').value;
  const phoneCol    = document.getElementById('map-phone').value;
  const notesCol    = document.getElementById('map-notes').value;
  const template    = document.getElementById('title-template').value.trim() || 'Follow-up reminder';
  const skipPast    = document.getElementById('skip-past').checked;
  const reminderMin = parseInt(document.getElementById('reminder-select').value, 10);

  if (!dateCol || dateCol === NONE) {
    alert('Please select the follow-up date column.');
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  previewEvents = [];

  sheetRows.forEach(function (row, idx) {
    const rawDate = row[dateCol] || '';
    const parsed  = parseDate(rawDate);

    /* Build the row object with column values */
    const fields = {
      Company: companyCol  !== NONE ? (row[companyCol]  || '') : '',
      Contact: contactCol  !== NONE ? (row[contactCol]  || '') : '',
      Phone:   phoneCol    !== NONE ? (row[phoneCol]    || '') : '',
      Notes:   notesCol    !== NONE ? (row[notesCol]    || '') : ''
    };
    /* Also allow {exact column name} in templates */
    Object.keys(row).forEach(function (k) { fields[k] = row[k]; });

    const title = buildTitle(template, fields);
    const desc  = buildDescription(fields);

    let status = 'ok';
    if (!parsed) {
      status = 'skip-nodate';
    } else if (skipPast && parsed < today) {
      status = 'skip-past';
    }

    previewEvents.push({
      idx,
      title,
      dateISO:    parsed ? toISO(parsed) : null,
      dateLabel:  parsed ? parsed.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' }) : rawDate,
      description: desc,
      reminderMin,
      status,
      selected:   status === 'ok'
    });
  });

  renderPreview();
  markDone(3);
  showPanel(4);
}

function renderPreview () {
  const okCount   = previewEvents.filter(function (e) { return e.status === 'ok'; }).length;
  const skipCount = previewEvents.filter(function (e) { return e.status !== 'ok'; }).length;

  document.getElementById('summary-bar').innerHTML =
    '<strong>' + okCount + '</strong> events ready to create' +
    (skipCount > 0 ? ' &nbsp;·&nbsp; <span style="color:#605e5c">' + skipCount + ' skipped</span>' : '');

  const list = document.getElementById('event-list');
  list.innerHTML = '';

  previewEvents.forEach(function (ev, i) {
    const row = document.createElement('div');
    row.className = 'event-row';

    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.checked = ev.selected;
    chk.disabled = ev.status !== 'ok';
    chk.addEventListener('change', function () {
      previewEvents[i].selected = chk.checked;
    });

    const info = document.createElement('div');
    info.className = 'event-info';
    info.innerHTML =
      '<div class="event-title">' + escHtml(ev.title) + '</div>' +
      '<div class="event-date">' + escHtml(ev.dateLabel) +
        (ev.status === 'skip-past'   ? ' <span style="color:#a4262c">(past date — skipped)</span>' : '') +
        (ev.status === 'skip-nodate' ? ' <span style="color:#a4262c">(no valid date)</span>'       : '') +
      '</div>';

    row.appendChild(chk);
    row.appendChild(info);
    list.appendChild(row);
  });

  /* Reset progress & results */
  document.getElementById('progress-section').classList.add('hidden');
  document.getElementById('result-list').innerHTML = '';
  document.getElementById('btn-push').disabled = false;
  document.getElementById('btn-push').textContent = 'Push to Google Calendar';
}

/* ─────────────────────────────────────────────
   STEP 5 — Push to Google Calendar
───────────────────────────────────────────── */
async function pushEvents () {
  const toPush = previewEvents.filter(function (e) { return e.selected && e.status === 'ok'; });
  if (toPush.length === 0) {
    alert('No events selected to push.');
    return;
  }

  document.getElementById('btn-push').disabled = true;
  document.getElementById('btn-push').textContent = 'Pushing...';
  document.getElementById('btn-back').disabled = true;

  const progressSection = document.getElementById('progress-section');
  const progressFill    = document.getElementById('progress-fill');
  const progressLabel   = document.getElementById('progress-label');
  const resultList      = document.getElementById('result-list');

  progressSection.classList.remove('hidden');
  resultList.innerHTML = '';

  let done = 0;

  for (const ev of toPush) {
    addResultRow(resultList, ev.title, 'pending', 'Creating...');

    try {
      /* Check for existing event to avoid duplicates */
      const existing = await checkDuplicate(ev.title, ev.dateISO);
      if (existing) {
        updateLastResultRow(resultList, 'skip', 'Already exists');
      } else {
        await createCalendarEvent(ev);
        updateLastResultRow(resultList, 'ok', 'Created');
      }
    } catch (err) {
      updateLastResultRow(resultList, 'error', err.message || 'Failed');
    }

    done++;
    const pct = Math.round((done / toPush.length) * 100);
    progressFill.style.width  = pct + '%';
    progressLabel.textContent = done + ' / ' + toPush.length + ' done';
  }

  document.getElementById('btn-push').textContent = 'Done!';
  document.getElementById('btn-back').disabled    = false;
}

async function checkDuplicate (title, dateISO) {
  const timeMin = dateISO + 'T00:00:00Z';
  const timeMax = dateISO + 'T23:59:59Z';
  const url = 'https://www.googleapis.com/calendar/v3/calendars/primary/events'
    + '?q=' + encodeURIComponent(title)
    + '&timeMin=' + encodeURIComponent(timeMin)
    + '&timeMax=' + encodeURIComponent(timeMax)
    + '&singleEvents=true';

  const res  = await apiFetch(url, 'GET');
  const data = await res.json();
  return data.items && data.items.length > 0;
}

async function createCalendarEvent (ev) {
  const event = {
    summary:     ev.title,
    description: ev.description,
    start:       { date: ev.dateISO },
    end:         { date: ev.dateISO },
    reminders: {
      useDefault: false,
      overrides:  ev.reminderMin > 0
        ? [{ method: 'popup', minutes: ev.reminderMin }]
        : []
    }
  };

  const res = await apiFetch(
    'https://www.googleapis.com/calendar/v3/calendars/primary/events',
    'POST',
    event
  );

  if (!res.ok) {
    const err = await res.json();
    throw new Error((err.error && err.error.message) || ('HTTP ' + res.status));
  }
  return res.json();
}

function apiFetch (url, method, body) {
  const opts = {
    method,
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Content-Type':  'application/json'
    }
  };
  if (body) opts.body = JSON.stringify(body);
  return fetch(url, opts);
}

/* ─────────────────────────────────────────────
   Helpers
───────────────────────────────────────────── */
function buildTitle (template, fields) {
  return template.replace(/\{([^}]+)\}/g, function (_, key) {
    return fields[key] !== undefined ? fields[key] : ('{' + key + '}');
  });
}

function buildDescription (fields) {
  return Object.entries(fields)
    .filter(function (kv) { return kv[1]; })
    .map(function (kv) { return kv[0] + ': ' + kv[1]; })
    .join('\n');
}

function parseDate (raw) {
  if (!raw) return null;
  /* Excel serial number (number stored as string) */
  if (/^\d{4,5}(\.\d+)?$/.test(raw)) {
    const serial = parseFloat(raw);
    const msOffset = (serial - 25569) * 86400 * 1000;
    const d = new Date(msOffset);
    if (!isNaN(d.getTime())) return d;
  }
  const d = new Date(raw);
  return isNaN(d.getTime()) ? null : d;
}

function toISO (date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function escHtml (str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function setStatus (id, msg, type) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = msg;
  el.className = 'status-msg' + (type ? ' ' + type : '');
}

function showPanel (n) {
  [1, 2, 3, 4].forEach(function (i) {
    document.getElementById('panel-' + i).classList.toggle('hidden', i !== n);
  });
  /* Activate current step dot */
  [1, 2, 3, 4].forEach(function (i) {
    document.getElementById('dot-' + i).classList.toggle('active', i === n);
  });
}

function markDone (n) {
  const dot = document.getElementById('dot-' + n);
  if (dot) {
    dot.classList.remove('active');
    dot.classList.add('done');
    dot.textContent = '✓';
  }
}

function addResultRow (list, name, badgeClass, badgeLabel) {
  const row = document.createElement('div');
  row.className = 'result-row';
  row.innerHTML =
    '<span class="result-name">' + escHtml(name) + '</span>' +
    '<span class="result-badge badge-' + badgeClass + '">' + badgeLabel + '</span>';
  list.appendChild(row);
}

function updateLastResultRow (list, badgeClass, badgeLabel) {
  const rows  = list.querySelectorAll('.result-row');
  const last  = rows[rows.length - 1];
  if (!last) return;
  const badge = last.querySelector('.result-badge');
  badge.className  = 'result-badge badge-' + badgeClass;
  badge.textContent = badgeLabel;
}
