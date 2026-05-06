'use strict';

/* ─── DOM refs ─────────────────────────────────────────── */
const uploadScreen    = document.getElementById('upload-screen');
const dashScreen      = document.getElementById('dashboard-screen');
const dropZone        = document.getElementById('drop-zone');
const fileInput       = document.getElementById('file-input');
const errorBanner     = document.getElementById('error-banner');
const errorText       = document.getElementById('error-text');
const btnReset        = document.getElementById('btn-reset');
const fileNameLabel   = document.getElementById('file-name-label');

/* ─── Navigation ────────────────────────────────────────── */
function showScreen(name) {
  uploadScreen.classList.toggle('active', name === 'upload');
  dashScreen.classList.toggle('active', name === 'dashboard');
}

/* ─── Drag & drop ───────────────────────────────────────── */
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) processFile(fileInput.files[0]);
});

btnReset.addEventListener('click', () => {
  fileInput.value = '';
  errorBanner.classList.add('hidden');
  showScreen('upload');
});

/* ─── Show/hide error ───────────────────────────────────── */
function showError(msg) {
  errorText.textContent = msg;
  errorBanner.classList.remove('hidden');
}

/* ─── Excel parsing ─────────────────────────────────────── */
function processFile(file) {
  errorBanner.classList.add('hidden');

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const data = extractData(wb);
      renderDashboard(data, file.name);
      showScreen('dashboard');
    } catch (err) {
      showError('Impossible de lire le fichier : ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ─── Data extraction ───────────────────────────────────── */
function extractData(wb) {
  const sheetNames = wb.SheetNames;

  // ── Find sheets by partial name match (case-insensitive) ──
  const find = keyword =>
    wb.Sheets[sheetNames.find(n => n.toLowerCase().includes(keyword.toLowerCase()))] || null;

  const wsMain   = find('suivi forfait') || find('suivi');
  const ws2026   = find('2026') && !find('conges-2026') ? find('2026') : find('conges-2026') || find('2026');
  const ws2025   = find('conges-2025') || find('2025');
  const wsData   = find('data');

  if (!wsMain) throw new Error('Onglet "Suivi forfait" introuvable. Vérifie le nom des onglets.');

  const cell = (ws, ref) => ws && ws[ref] ? ws[ref].v : null;
  const num  = (ws, ref) => { const v = cell(ws, ref); return (v !== null && !isNaN(+v)) ? +v : null; };

  // ── Main sheet data ──
  const cpPlafond  = num(wsMain, 'H6') ?? 25;
  const rttPlafond = num(wsMain, 'H7') ?? 9;
  const cpAcquis   = num(wsMain, 'K11');
  const dateRaw    = cell(wsMain, 'K10');

  // ── 2026 congés ──
  const ws26 = find('conges-2026') || find('conges 2026');
  let cp2026  = null;
  let rtt2026 = null;
  let rttMonths2026 = [];

  if (ws26) {
    cp2026  = num(ws26, 'B14'); // SUM(B2:B13)
    rtt2026 = num(ws26, 'C14');

    const monthCodes = ['J','F','M','A','M','J','J','A','S','O','N','D'];
    const monthNames = ['Janvier','Février','Mars','Avril','Mai','Juin',
                        'Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
    for (let i = 0; i < 12; i++) {
      const rttVal = num(ws26, `C${i + 2}`);
      if (rttVal) {
        rttMonths2026.push({ name: monthNames[i], val: rttVal });
      }
    }
  } else {
    // fallback: read from main sheet columns D/E
    let cpSum = 0, rttSum = 0;
    for (let r = 5; r <= 16; r++) {
      cpSum  += num(wsMain, `D${r}`) ?? 0;
      rttSum += num(wsMain, `E${r}`) ?? 0;
    }
    cp2026  = cpSum;
    rtt2026 = rttSum;
  }

  // ── 2025 congés ──
  let cp2025  = null;
  let rtt2025 = null;

  if (ws2025) {
    cp2025  = num(ws2025, 'B11');
    rtt2025 = num(ws2025, 'C11');
  } else if (wsData) {
    // Try reading TOTAL row from Data sheet
    for (let r = 5; r <= 30; r++) {
      const a = cell(wsData, `A${r}`);
      if (a && String(a).toLowerCase() === 'total') {
        cp2025  = num(wsData, `B${r}`);
        rtt2025 = num(wsData, `C${r}`);
        break;
      }
    }
  }

  // ── Dates / year labels ──
  let arrivalDate = null;
  if (dateRaw instanceof Date) arrivalDate = dateRaw;
  else if (typeof dateRaw === 'number') {
    arrivalDate = XLSX.SSF.parse_date_code(dateRaw);
    if (arrivalDate) arrivalDate = new Date(arrivalDate.y, arrivalDate.m - 1, arrivalDate.d);
  }

  const arrivalStr = arrivalDate
    ? arrivalDate.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
    : null;

  const taux = num(wsMain, 'B2') ?? 1;

  return {
    cp: {
      past:     cp2025 ?? 0,
      current:  cp2026 ?? 0,
      plafond:  cpPlafond,
      acquis:   cpAcquis,
    },
    rtt: {
      current:  rtt2026 ?? 0,
      past:     rtt2025,
      plafond:  rttPlafond,
      months:   rttMonths2026,
    },
    arrivalStr,
    taux,
  };
}

/* ─── Render dashboard ──────────────────────────────────── */
function renderDashboard(d, fileName) {
  fileNameLabel.textContent = fileName;

  // ── Pills ──
  document.getElementById('pill-arrival-text').textContent =
    d.arrivalStr ? `Arrivée le ${d.arrivalStr}` : 'Date d\'arrivée inconnue';

  const tauxPct = Math.round(d.taux * 100);
  document.getElementById('pill-taux-text').textContent =
    tauxPct === 100 ? 'Temps plein' : `Temps partiel ${tauxPct}%`;

  document.getElementById('pill-acquis-text').textContent =
    d.cp.acquis !== null ? `CP acquis : ${fmt(d.cp.acquis)} j` : 'CP acquis : —';

  // ── CP card ──
  const cpTotal  = round2(d.cp.past + d.cp.current);
  const cpRef    = d.cp.acquis ?? d.cp.plafond;
  const cpReste  = round2(cpRef - cpTotal);
  const cpPct    = Math.min((cpTotal / cpRef) * 100, 100);
  const cpOver   = cpTotal > cpRef;
  const cpBarPct = cpOver ? 100 : cpPct;
  const overPct  = cpOver ? round2(((cpTotal - cpRef) / cpRef) * 15) : 0; // visual overflow capped at 15%

  setText('cp-big-num', cpOver ? `+${fmt(Math.abs(cpReste))} j` : `${fmt(cpReste)} j`);
  setText('cp-status-label', cpOver ? 'Dépassement de' : 'Reste à poser');
  setNumStyle('cp-big-num', cpOver ? 'red' : cpReste <= 3 ? 'orange' : 'green');

  setBadge('cp-badge',
    cpOver       ? 'Trop pris'             : null,
    cpReste <= 0 ? 'danger'
    : cpReste <= 3 ? 'warn'
    : 'ok',
    cpOver       ? null : cpReste <= 3 ? 'Attention'  : 'OK'
  );

  setBar('cp-bar', cpBarPct, cpOver ? 'var(--red)' : 'var(--text-primary)');

  const cpOvEl = document.getElementById('cp-overflow-bar');
  if (cpOver) {
    cpOvEl.style.display = 'block';
    cpOvEl.style.width   = Math.min(overPct, 25) + '%';
  } else {
    cpOvEl.style.display = 'none';
  }

  setTicks('cp-ticks', `0`, `${fmt(cpRef)} acquis`, `${fmt(d.cp.plafond)} plafond`);

  setText('cp-past',          `${fmt(d.cp.past)} j`);
  setText('cp-current',       `${fmt(d.cp.current)} j`);
  setText('cp-total-taken',   `${fmt(cpTotal)} j`);
  setText('cp-acquis-detail', d.cp.acquis !== null ? `${fmt(d.cp.acquis)} j` : '—');
  setText('cp-plafond-detail',`${fmt(d.cp.plafond)} j`);

  const totalEl = document.getElementById('cp-total-taken');
  totalEl.style.color = cpOver ? 'var(--red)' : 'inherit';

  // ── RTT card ──
  const rttTotal = d.rtt.current;
  const rttReste = round2(d.rtt.plafond - rttTotal);
  const rttOver  = rttTotal > d.rtt.plafond;
  const rttPct   = Math.min((rttTotal / d.rtt.plafond) * 100, 100);

  setText('rtt-big-num', rttOver ? `+${fmt(Math.abs(rttReste))} j` : `${fmt(rttReste)} j`);
  setText('rtt-status-label', rttOver ? 'Dépassement de' : 'Reste à poser');
  setNumStyle('rtt-big-num', rttOver ? 'red' : rttReste <= 2 ? 'orange' : 'green');

  setBadge('rtt-badge',
    rttOver       ? 'Trop pris' : null,
    rttOver       ? 'danger' : rttReste <= 2 ? 'warn' : 'ok',
    rttOver       ? null : rttReste <= 2 ? 'Attention' : 'OK'
  );

  setBar('rtt-bar', rttPct, rttOver ? 'var(--red)' : 'var(--text-primary)');
  setTicks('rtt-ticks', '0', `${fmt(rttTotal)} pris`, `${fmt(d.rtt.plafond)} plafond`);

  const rttDetail = document.getElementById('rtt-months-detail');
  rttDetail.innerHTML = '';
  if (d.rtt.months && d.rtt.months.length > 0) {
    d.rtt.months.forEach(m => {
      const row = document.createElement('div');
      row.className = 'rtt-month-row';
      row.innerHTML = `<span class="rtt-month-name">${m.name}</span><span class="rtt-month-val">${fmt(m.val)} j</span>`;
      rttDetail.appendChild(row);
    });
  } else {
    rttDetail.innerHTML = '<span style="font-size:13px;color:var(--text-muted)">Aucun RTT posé en 2026</span>';
  }

  setText('rtt-total-taken',  `${fmt(rttTotal)} j`);
  setText('rtt-plafond-detail', `${fmt(d.rtt.plafond)} j`);

  // ── Alert zone ──
  const alerts = document.getElementById('alert-zone');
  alerts.innerHTML = '';

  if (cpOver) {
    const diff = fmt(Math.abs(cpReste));
    addAlert(alerts, 'danger',
      `Tu as pris <strong>${diff} jours de CP en trop</strong> par rapport aux ${fmt(cpRef)} j acquis.
       Rapproche-toi de ton service RH pour régulariser.`);
  } else if (cpReste <= 3 && cpReste >= 0) {
    addAlert(alerts, 'warn',
      `Plus que <strong>${fmt(cpReste)} j de CP à poser</strong> avant la fin du cycle (mai).
       Pense à planifier tes derniers congés.`);
  } else {
    addAlert(alerts, 'ok',
      `Il te reste <strong>${fmt(cpReste)} j de CP à poser</strong> d'ici la fin du cycle mai ${new Date().getFullYear()}.`);
  }

  if (rttOver) {
    addAlert(alerts, 'danger',
      `Tu as pris <strong>${fmt(Math.abs(rttReste))} RTT en trop</strong>. Contacte ton service RH.`);
  } else if (rttReste <= 2 && rttReste >= 0) {
    addAlert(alerts, 'warn',
      `Plus que <strong>${fmt(rttReste)} RTT à poser</strong> avant fin décembre. Les RTT non pris sont perdus !`);
  } else {
    addAlert(alerts, 'ok',
      `Il te reste <strong>${fmt(rttReste)} RTT à poser</strong> avant fin décembre. Les RTT non pris sont perdus.`);
  }

  // ── Previous year RTT note ──
  const prevNote = document.getElementById('rtt-prev-note');
  if (d.rtt.past !== null && d.rtt.past > 0) {
    prevNote.style.display = 'flex';
    document.getElementById('rtt-prev-note-text').textContent =
      `RTT 2025 (cycle précédent) : ${fmt(d.rtt.past)} j posés — ce cycle est maintenant terminé.`;
  } else {
    prevNote.style.display = 'none';
  }
}

/* ─── Helpers ───────────────────────────────────────────── */
function round2(n)      { return Math.round(n * 100) / 100; }
function fmt(n)         { return n === null ? '—' : (Number.isInteger(n * 10) ? n.toFixed(1 * (n % 1 !== 0)) : n.toFixed(1)); }
function setText(id, t) { const el = document.getElementById(id); if (el) el.textContent = t; }

function setNumStyle(id, state) {
  const el = document.getElementById(id);
  if (!el) return;
  el.style.color = state === 'red' ? 'var(--red)' : state === 'orange' ? 'var(--orange)' : 'var(--green)';
}

function setBar(id, pct, color) {
  const el = document.getElementById(id);
  if (!el) return;
  el.style.width    = Math.min(Math.max(pct, 0), 100) + '%';
  el.style.background = color;
}

function setTicks(id, left, mid, right) {
  const el = document.getElementById(id);
  if (!el) return;
  el.innerHTML = `<span>${left}</span><span style="color:var(--text-primary);font-weight:500">${mid}</span><span>${right}</span>`;
}

function setBadge(id, dangerLabel, state, okLabel) {
  const el = document.getElementById(id);
  if (!el) return;
  el.className = 'card-badge badge-' + state;
  el.textContent = state === 'danger' ? (dangerLabel || 'Dépassé') : state === 'warn' ? (okLabel || 'Attention') : (okLabel || 'OK');
}

function addAlert(container, type, html) {
  const icons = {
    danger: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`,
    warn:   `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`,
    ok:     `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`,
  };
  const div = document.createElement('div');
  div.className = `alert alert-${type}`;
  div.innerHTML = icons[type] + `<span>${html}</span>`;
  container.appendChild(div);
}
