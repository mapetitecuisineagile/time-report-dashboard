'use strict';
document.addEventListener('DOMContentLoaded', () => {

/* ─── STATE ─── */
const MONTHS_2026 = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
const MONTHS_2025 = ['Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
const SHORT = ['J','F','M','A','M','J','J','A','S','O','N','D'];
const SHORT25 = ['A','M','J','J','A','S','O','N','D'];

let state = {
  params: { forfait: 218, cpAcquis: 24.8, cpPlafond: 25, rttPlafond: 9 },
  data2026: MONTHS_2026.map(m => ({ mois: m, prod: null, int: null, cp: null, rtt: null })),
  data2025: MONTHS_2025.map(m => ({ mois: m, cp: null, rtt: null })),
};

/* ─── HELPERS ─── */
const $  = id => document.getElementById(id);
const r2 = n  => Math.round((n || 0) * 100) / 100;
const fmt = n => n == null ? '—' : (n % 1 === 0 ? n.toString() : n.toFixed(1));
const sum = (arr, key) => r2(arr.reduce((a, row) => a + (row[key] || 0), 0));

/* ─── PAGES ─── */
function show(pageId) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  $(pageId).classList.add('active');
}

/* ─── UPLOAD ─── */
const dropzone = $('dropzone');
const fileIn   = $('file-in');

dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('over'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('over'));
dropzone.addEventListener('drop', e => { e.preventDefault(); dropzone.classList.remove('over'); handleFile(e.dataTransfer.files[0]); });
fileIn.addEventListener('change', () => handleFile(fileIn.files[0]));
dropzone.addEventListener('click', e => { if (e.target.tagName !== 'LABEL') fileIn.click(); });

$('btn-back').addEventListener('click', () => show('page-upload'));

function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      parseWorkbook(wb);
      $('header-file').textContent = file.name;
      launch();
    } catch(err) {
      const el = $('upload-err');
      el.textContent = 'Impossible de lire le fichier : ' + err.message;
      el.classList.remove('hidden');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ─── PARSE EXCEL ─── */
function parseWorkbook(wb) {
  const names = wb.SheetNames;
  const findSheet = kw => wb.Sheets[names.find(n => n.toLowerCase().includes(kw.toLowerCase()))] || null;

  const wsMain  = findSheet('suivi forfait') || findSheet('suivi');
  const ws2026  = findSheet('conges-2026') || findSheet('conges 2026');
  const ws2025  = findSheet('conges-2025') || findSheet('conges 2025');
  const wsData  = findSheet('data');

  const val = (ws, ref) => ws && ws[ref] ? ws[ref].v : null;
  const num = (ws, ref) => { const v = val(ws, ref); return v !== null && !isNaN(+v) ? +v : null; };

  // Params from main sheet
  if (wsMain) {
    state.params.forfait    = num(wsMain, 'H3') || 218;
    state.params.cpPlafond  = num(wsMain, 'H6') || 25;
    state.params.rttPlafond = num(wsMain, 'H7') || 9;
    state.params.cpAcquis   = num(wsMain, 'K11') || 24.8;

    // 2026 monthly data: rows 5-16 = J,F,M,A,M,J,J,A,S,O,N,D
    for (let i = 0; i < 12; i++) {
      const r = i + 5;
      state.data2026[i].prod = num(wsMain, `B${r}`);
      state.data2026[i].int  = num(wsMain, `C${r}`);
      state.data2026[i].cp   = num(wsMain, `D${r}`);
      state.data2026[i].rtt  = num(wsMain, `E${r}`);
    }
  }

  // 2025 data from Conges-2025 or Data sheet
  if (ws2025) {
    for (let i = 0; i < 9; i++) {
      const r = i + 2;
      state.data2025[i].cp  = num(ws2025, `B${r}`);
      state.data2025[i].rtt = num(ws2025, `C${r}`);
    }
  } else if (wsData) {
    const months25 = ['AVRIL','MAI','JUIN','JUILLET','AOUT','SEPTEMBRE','OCTOBRE','NOVEMBRE','DECEMBRE'];
    for (let r = 1; r <= 30; r++) {
      const label = val(wsData, `A${r}`);
      if (!label) continue;
      const idx = months25.findIndex(m => String(label).toUpperCase().includes(m));
      if (idx >= 0) {
        state.data2025[idx].cp  = num(wsData, `B${r}`);
        state.data2025[idx].rtt = num(wsData, `C${r}`);
      }
    }
  }

  syncParamInputs();
}

/* ─── INIT ─── */
function launch() {
  buildTable2026();
  buildTable2025();
  syncParamInputs();
  recalc();
  show('page-app');
}

/* ─── BUILD TABLES ─── */
function buildTable2026() {
  const tbody = $('table-body');
  tbody.innerHTML = '';
  state.data2026.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.dataset.idx = i;
    tr.innerHTML = `
      <td class="col-mois">${SHORT[i]}</td>
      <td class="editable" data-key="prod">${fmt(row.prod)}</td>
      <td class="editable" data-key="int">${fmt(row.int)}</td>
      <td class="editable" data-key="cp">${fmt(row.cp)}</td>
      <td class="editable" data-key="rtt">${fmt(row.rtt)}</td>
      <td class="col-total">${fmt(rowTotal26(row))}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.addEventListener('click', onCellClick26);
}

function buildTable2025() {
  const tbody = $('table-body-2025');
  tbody.innerHTML = '';
  state.data2025.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.dataset.idx = i;
    tr.innerHTML = `
      <td class="col-mois">${SHORT25[i]}</td>
      <td class="editable" data-key="cp">${fmt(row.cp)}</td>
      <td class="editable" data-key="rtt">${fmt(row.rtt)}</td>
    `;
    tbody.appendChild(tr);
  });
  tbody.addEventListener('click', onCellClick25);
}

function rowTotal26(row) {
  const t = (row.prod||0) + (row.int||0) + (row.cp||0) + (row.rtt||0);
  return t || null;
}

/* ─── INLINE EDITING ─── */
function onCellClick26(e) {
  const td = e.target.closest('td.editable');
  if (!td || td.querySelector('input')) return;
  const tr  = td.closest('tr');
  const idx = +tr.dataset.idx;
  const key = td.dataset.key;
  startEdit(td, state.data2026[idx][key], val => {
    state.data2026[idx][key] = val;
    td.textContent = fmt(val);
    tr.cells[5].textContent = fmt(rowTotal26(state.data2026[idx]));
    updateFooter26();
    recalc();
  });
}

function onCellClick25(e) {
  const td = e.target.closest('td.editable');
  if (!td || td.querySelector('input')) return;
  const tr  = td.closest('tr');
  const idx = +tr.dataset.idx;
  const key = td.dataset.key;
  startEdit(td, state.data2025[idx][key], val => {
    state.data2025[idx][key] = val;
    td.textContent = fmt(val);
    updateFooter25();
    recalc();
  });
}

function startEdit(td, current, onDone) {
  const input = document.createElement('input');
  input.type  = 'number';
  input.value = current != null ? current : '';
  input.min   = 0; input.step = 0.5;
  td.textContent = '';
  td.appendChild(input);
  input.focus();
  input.select();

  function commit() {
    const v = input.value.trim() === '' ? null : parseFloat(input.value);
    td.removeChild(input);
    onDone(v);
  }
  input.addEventListener('blur', commit);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); commit(); }
    if (e.key === 'Escape') { td.textContent = fmt(current); td.removeChild(input); }
  });
}

/* ─── FOOTERS ─── */
function updateFooter26() {
  const prod = sum(state.data2026, 'prod');
  const int_ = sum(state.data2026, 'int');
  const cp   = sum(state.data2026, 'cp');
  const rtt  = sum(state.data2026, 'rtt');
  $('tot-prod').textContent = fmt(prod);
  $('tot-int').textContent  = fmt(int_);
  $('tot-cp').textContent   = fmt(cp);
  $('tot-rtt').textContent  = fmt(rtt);
  $('tot-all').textContent  = fmt(r2(prod + int_ + cp + rtt));
}

function updateFooter25() {
  $('tot25-cp').textContent  = fmt(sum(state.data2025, 'cp'));
  $('tot25-rtt').textContent = fmt(sum(state.data2025, 'rtt'));
}

/* ─── RECALC & RENDER ─── */
function recalc() {
  updateFooter26();
  updateFooter25();

  const p = state.params;

  // Forfait
  const travaille = r2(sum(state.data2026, 'prod') + sum(state.data2026, 'int'));
  const resteF    = r2(p.forfait - travaille);

  // CP cycle mai→mai
  const cp2025  = sum(state.data2025, 'cp');
  const cp2026  = sum(state.data2026, 'cp');
  const cpTotal = r2(cp2025 + cp2026);
  const cpReste = r2(p.cpAcquis - cpTotal);

  // RTT cycle jan→déc 2026
  const rttTotal = sum(state.data2026, 'rtt');
  const rttReste = r2(p.rttPlafond - rttTotal);

  /* ── Forfait counter ── */
  const forfaitPct = Math.min((travaille / p.forfait) * 100, 100);
  $('val-forfait').textContent = fmt(resteF);
  $('hint-forfait').textContent = `${fmt(travaille)} jh travaillés sur ${p.forfait}`;
  setBar('bar-forfait', forfaitPct);
  setBadge('bdg-forfait',
    resteF <= 0   ? ['Objectif atteint', 'ok'] :
    resteF <= 20  ? ['Bientôt !', 'warn'] :
                    [fmt(Math.round(forfaitPct)) + '%', 'ok']
  );
  $('cnt-forfait').style.color = 'inherit';

  /* ── CP counter ── */
  const cpPct = Math.min((cpTotal / p.cpAcquis) * 100, 100);
  $('val-cp').textContent = fmt(Math.abs(cpReste));
  $('val-cp').style.color = cpReste < 0 ? 'var(--red)' : 'inherit';
  $('hint-cp').textContent = `${fmt(cpTotal)} j pris sur ${p.cpAcquis} acquis`;
  setBar('bar-cp', cpPct, cpReste < 0 ? 'var(--red)' : null);
  setBadge('bdg-cp',
    cpReste < 0   ? ['Dépassé de ' + fmt(Math.abs(cpReste)) + 'j', 'danger'] :
    cpReste <= 3  ? ['Plus que ' + fmt(cpReste) + 'j', 'warn'] :
                    [fmt(cpReste) + 'j restants', 'ok']
  );

  /* ── RTT counter ── */
  const rttPct = Math.min((rttTotal / p.rttPlafond) * 100, 100);
  $('val-rtt').textContent = fmt(Math.abs(rttReste));
  $('val-rtt').style.color = rttReste < 0 ? 'var(--red)' : 'inherit';
  $('hint-rtt').textContent = `${fmt(rttTotal)} j pris sur ${p.rttPlafond} de plafond`;
  setBar('bar-rtt', rttPct, rttReste < 0 ? 'var(--red)' : null);
  setBadge('bdg-rtt',
    rttReste < 0  ? ['Dépassé de ' + fmt(Math.abs(rttReste)) + 'j', 'danger'] :
    rttReste <= 2 ? ['Plus que ' + fmt(rttReste) + 'j', 'warn'] :
                    [fmt(rttReste) + 'j restants', 'ok']
  );
}

function setBar(id, pct, color) {
  const el = $(id);
  el.style.width = Math.max(0, pct) + '%';
  if (color) el.style.background = color;
  else el.style.background = '';
}

function setBadge(id, [label, type]) {
  const el = $(id);
  el.textContent = label;
  el.className = 'cnt-badge badge-' + type;
}

/* ─── PARAMS ─── */
function syncParamInputs() {
  $('p-forfait').value    = state.params.forfait;
  $('p-cp-acquis').value  = state.params.cpAcquis;
  $('p-cp-plafond').value = state.params.cpPlafond;
  $('p-rtt-plafond').value = state.params.rttPlafond;
}

['p-forfait','p-cp-acquis','p-cp-plafond','p-rtt-plafond'].forEach(id => {
  $(id).addEventListener('input', () => {
    state.params.forfait    = +$('p-forfait').value    || 218;
    state.params.cpAcquis   = +$('p-cp-acquis').value  || 24.8;
    state.params.cpPlafond  = +$('p-cp-plafond').value || 25;
    state.params.rttPlafond = +$('p-rtt-plafond').value || 9;
    recalc();
  });
});

/* ─── START WITHOUT FILE ─── */
$('header-file').textContent = 'nouveau fichier';
launch();
show('page-upload'); // reset to upload on load

}); // fin DOMContentLoaded
