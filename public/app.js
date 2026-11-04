/* global Office, Excel */
import { CONFIG, CORRECTION_HEADERS, FIELD_LABELS, LOG_HEADERS } from './config.js';
import { createCombo } from './combobox.js';
import {
  checkClassification, computeMoveTarget, detectLayout, diffFields, filterPaths, finalRowAfterMove, findLevel2, findPath, formatTime,
  keyOf, level1List, level2Entries, masterOrder, parseMaster, pathProblem, planCorrectionColumns, safeText, suggest, tidy, valueOptions
} from './core.js';

const $ = id => document.getElementById(id);
const COPY_FIELDS = ['level1', 'kode2', 'level2', 'level3', 'nama', 'spesifikasi'];
const state = { master: null, order: null, layout: null, sheetId: null, sheetName: '', lastRow: 0, rowIndex: -1, row: null, multi: false, index: null, indexSheetId: null, orig: null, draft: null, rowKey: '', busy: false };
let queue = Promise.resolve(), selectionTimer, noteTimer, pendingNote = null;

function tell(text, error = false) { $('message').textContent = text; $('message').classList.toggle('error', error); }
function failure(error) {
  const code = error?.code || '';
  if (['AccessDenied', 'ProtectedRange', 'InvalidOperationInCellEditMode'].includes(code)) return 'Excel menolak perubahan. Pastikan sel dapat diedit (sheet tidak diproteksi) dan Anda tidak sedang mengetik di sel.';
  return error?.message || 'Operasi Excel gagal. Coba lagi.';
}
function enqueue(work) {
  queue = queue.then(async () => {
    state.busy = true; render();
    try { await work(); } catch (error) { tell(failure(error), true); }
    finally { state.busy = false; render(); }
  });
  return queue;
}

// ───────────── Pembacaan data Excel ─────────────

async function readMaster(context) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.masterSheet);
  sheet.load('isNullObject'); await context.sync();
  if (sheet.isNullObject) throw new Error(`Sheet ${CONFIG.masterSheet} tidak ditemukan pada workbook ini.`);
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load('isNullObject,rowIndex,columnIndex,rowCount,columnCount'); await context.sync();
  if (used.isNullObject) throw new Error('Master_Level masih kosong.');
  if (used.rowCount > CONFIG.maxMasterRows) throw new Error(`Master melebihi ${CONFIG.maxMasterRows} baris.`);
  const data = sheet.getRangeByIndexes(used.rowIndex, used.columnIndex, used.rowCount, Math.min(used.columnCount, CONFIG.maxMasterColumns));
  data.load('values'); await context.sync();
  return parseMaster(data.values, CONFIG.headerSearchRows);
}

async function loadLayout(context, sheet) {
  const head = sheet.getRangeByIndexes(0, 0, CONFIG.headerSearchRows, CONFIG.maxHeaderColumns);
  head.load('values'); await context.sync();
  return detectLayout(head.values, CONFIG.headerSearchRows);
}

/** Seleksi aktif + tata letak sheet + isi baris terpilih. */
async function readActiveRow(context) {
  const range = context.workbook.getSelectedRange();
  const sheet = range.worksheet;
  range.load('rowIndex,rowCount');
  sheet.load('id,name');
  const used = sheet.getUsedRange(true);
  used.load('rowIndex,rowCount');
  await context.sync();
  const layout = await loadLayout(context, sheet);
  const info = { sheetId: sheet.id, sheetName: sheet.name, layout, lastRow: used.rowIndex + used.rowCount - 1, rowIndex: range.rowIndex, multi: range.rowCount > 1, row: null };
  if (!layout || range.rowIndex <= layout.headerRow) return info;
  const width = Math.max(layout.lastUsed, ...Object.values(layout.correction), 0) + 1;
  const rowRange = sheet.getRangeByIndexes(range.rowIndex, 0, 1, width);
  rowRange.load('values'); await context.sync();
  info.row = rowRange.values[0];
  return info;
}

/** Membaca kolom-kolom yang dikoreksi untuk seluruh baris data. */
async function readIndex(context, sheet, layout, lastRow) {
  const first = layout.headerRow + 1, count = lastRow - layout.headerRow;
  if (count <= 0) return [];
  const names = COPY_FIELDS.filter(f => f in layout.cols);
  const ranges = Object.fromEntries(names.map(f => [f, sheet.getRangeByIndexes(first, layout.cols[f], count, 1)]));
  for (const range of Object.values(ranges)) range.load('values');
  await context.sync();
  const rows = [];
  for (let i = 0; i < count; i++) {
    const item = { row: first + i };
    let any = false;
    for (const f of COPY_FIELDS) { item[f] = f in ranges ? tidy(ranges[f].values[i][0]) : ''; if (item[f]) any = true; }
    if (any) rows.push(item);
  }
  return rows;
}

const fieldsOfRow = (row, layout) => Object.fromEntries(COPY_FIELDS.map(f => [f, f in layout.cols ? tidy(row[layout.cols[f]]) : '']));
const cell = name => (state.row && state.layout && name in state.layout.cols ? tidy(state.row[state.layout.cols[name]]) : '');
const noteOf = (row, layout) => (row && layout && 'note' in layout.correction ? tidy(row[layout.correction.note]) : '');

async function refresh() {
  const info = await Excel.run(async context => {
    const base = await readActiveRow(context);
    let index = null;
    if (base.layout && base.row && (state.indexSheetId !== base.sheetId || !state.index)) {
      index = await readIndex(context, context.workbook.worksheets.getItem(base.sheetId), base.layout, base.lastRow);
    }
    return { ...base, index };
  });
  Object.assign(state, { layout: info.layout, sheetId: info.sheetId, sheetName: info.sheetName, lastRow: info.lastRow, rowIndex: info.rowIndex, row: info.row, multi: info.multi });
  if (info.index) { state.index = info.index; state.indexSheetId = info.sheetId; }
  const key = `${info.sheetId}:${info.rowIndex}`;
  if (key !== state.rowKey) {
    state.rowKey = key;
    if (info.row) initDraft(); else { state.orig = null; state.draft = null; }
  }
}

function initDraft() {
  state.orig = fieldsOfRow(state.row, state.layout);
  state.draft = { ...state.orig };
  $('note').value = noteOf(state.row, state.layout);
  syncInputs();
}

// ───────────── Formulir koreksi ─────────────

function syncInputs() {
  const d = state.draft;
  if (!d) return;
  fill1();
  $('level1').value = d.level1;
  level2Combo.setText(d.level2); level3Combo.setText(d.level3); namaCombo.setText(d.nama); spekCombo.setText(d.spesifikasi);
}
function fill1() {
  const select = $('level1');
  const list = state.master ? level1List(state.master) : [];
  const current = state.draft?.level1;
  if (current && !list.some(v => keyOf(v) === keyOf(current))) list.push(current);
  select.replaceChildren(new Option('(pilih Level 1)', ''));
  for (const value of list) select.add(new Option(value, value));
}

function pickLevel2(entry) {
  const d = state.draft;
  d.level1 = entry.level1; d.level2 = entry.level2; d.kode2 = entry.kode2;
  if (!findPath(state.master, d.level1, d.level2, d.level3)) d.level3 = '';
  syncInputs(); renderEdit();
}
const level2Combo = createCombo({
  input: $('level2-input'), toggle: $('level2-toggle'), list: $('level2-list'),
  getOptions(text, browse) {
    if (!state.master) return [];
    const all = level2Entries(state.master), q = keyOf(text), d = state.draft;
    let list = all;
    if (browse || !q) list = all.filter(e => !d.level1 || keyOf(e.level1) === keyOf(d.level1));
    else list = all.filter(e => keyOf(`${e.level2} ${e.kode2} ${e.level1}`).includes(q))
      .sort((a, b) => Number(keyOf(b.level1) === keyOf(d.level1)) - Number(keyOf(a.level1) === keyOf(d.level1)));
    return list.slice(0, CONFIG.optionLimit).map(e => ({ label: e.level2, sub: `${e.kode2 || '-'} · ${e.level1}`, data: e }));
  },
  onPick: pickLevel2,
  onCommit(text) {
    const d = state.draft; if (!d || !state.master) return;
    const q = keyOf(text);
    if (q === keyOf(d.level2)) { level2Combo.setText(d.level2); return; }
    const matches = level2Entries(state.master).filter(e => keyOf(e.level2) === q);
    const pick = matches.find(e => keyOf(e.level1) === keyOf(d.level1)) || (matches.length === 1 ? matches[0] : null);
    if (pick) pickLevel2(pick);
    level2Combo.setText(state.draft.level2);
  }
});

function pickLevel3(path) {
  const d = state.draft;
  d.level1 = path.level1; d.level2 = path.level2; d.kode2 = path.kode2; d.level3 = path.level3;
  syncInputs(); renderEdit();
}
const level3Combo = createCombo({
  input: $('level3-input'), toggle: $('level3-toggle'), list: $('level3-list'),
  getOptions(text, browse) {
    if (!state.master) return [];
    const d = state.draft, q = keyOf(text), all = state.master.paths;
    let list;
    if (browse || !q) list = all.filter(p => (!d.level1 || keyOf(p.level1) === keyOf(d.level1)) && (!d.level2 || keyOf(p.level2) === keyOf(d.level2)));
    else list = filterPaths(state.master, text, 500).sort((a, b) => Number(keyOf(b.level2) === keyOf(d.level2)) - Number(keyOf(a.level2) === keyOf(d.level2)));
    return list.slice(0, CONFIG.optionLimit).map(p => ({ label: p.level3, sub: `${p.level1} › ${p.level2}`, data: p }));
  },
  onPick: pickLevel3,
  onCommit(text) {
    const d = state.draft; if (!d || !state.master) return;
    const q = keyOf(text);
    if (q === keyOf(d.level3)) { level3Combo.setText(d.level3); return; }
    const matches = state.master.byLevel3.get(q) || [];
    const pick = matches.find(p => keyOf(p.level1) === keyOf(d.level1) && keyOf(p.level2) === keyOf(d.level2)) || (matches.length === 1 ? matches[0] : null);
    if (pick) pickLevel3(pick); else level3Combo.setText(d.level3);
  }
});

function freeCombo(field, prefix) {
  return createCombo({
    input: $(`${prefix}-input`), toggle: $(`${prefix}-toggle`), list: $(`${prefix}-list`),
    getOptions: (text, browse) => (state.draft ? valueOptions(state.index || [], field, state.draft, text, browse, CONFIG.optionLimit) : []),
    onPick(value) { state.draft[field] = value; syncInputs(); renderEdit(); },
    onCommit() {}
  });
}
const namaCombo = freeCombo('nama', 'nama');
const spekCombo = freeCombo('spesifikasi', 'spek');
$('nama-input').addEventListener('input', () => { if (state.draft) { state.draft.nama = $('nama-input').value; renderEdit(); } });
$('spek-input').addEventListener('input', () => { if (state.draft) { state.draft.spesifikasi = $('spek-input').value; renderEdit(); } });

$('level1').addEventListener('change', () => {
  const d = state.draft; if (!d || !state.master) return;
  d.level1 = $('level1').value;
  const l2 = d.level1 && d.level2 ? findLevel2(state.master, d.level1, d.level2) : null;
  if (l2) d.kode2 = l2.kode2; else { d.level2 = ''; d.kode2 = ''; }
  if (!findPath(state.master, d.level1, d.level2, d.level3)) d.level3 = '';
  syncInputs(); renderEdit();
});

function currentProblem() {
  if (!state.draft || !state.master) return '';
  return pathProblem(state.draft, state.master);
}

function renderEdit() {
  const d = state.draft;
  if (!d) return;
  $('kode2').textContent = d.kode2 ? `(Kode ${d.kode2}, otomatis)` : '';
  const changes = diffFields(state.orig, d);
  const box = $('changes'); box.replaceChildren(); box.hidden = !changes.length;
  for (const c of changes) {
    const line = document.createElement('div');
    line.textContent = `${FIELD_LABELS[c.field]}: ${c.old || '(kosong)'} → ${c.new || '(kosong)'}`;
    box.append(line);
  }
  const problem = changes.length ? currentProblem() : '';
  $('problem').textContent = problem; $('problem').hidden = !problem;
  $('save').disabled = state.busy || !changes.length || Boolean(problem);
  $('reset').disabled = state.busy || !changes.length;
}

function render() {
  const ready = Boolean(state.master);
  $('count').textContent = ready ? `${state.master.paths.length} baris master` : '';
  $('reload').disabled = state.busy;
  const layout = state.layout, rowOk = Boolean(layout && state.row && state.draft);
  $('setup').hidden = Boolean(layout);
  if (!layout) $('setup-text').textContent = `Sheet "${state.sheetName || 'aktif'}" tidak dikenali sebagai tabel katalog. Pastikan baris judul memuat Level 1, Level 2, Level 3, dan Nama Produk / Pekerjaan, lalu pilih sebuah sel di sheet tersebut.`;
  $('row-card').hidden = !rowOk;
  if (!rowOk) { if (layout && state.multi) tell('Pilih satu baris data saja.', true); return; }

  $('row-address').textContent = `${state.sheetName} · baris ${state.rowIndex + 1}`;
  $('r-nama').textContent = cell('nama') || '(kosong)';
  $('r-spek').textContent = cell('spesifikasi') || '-';
  $('r-misc').textContent = [cell('satuan') || '-', cell('merek') || '-'].join(' · ');
  const cur = state.orig;
  $('c-l1').textContent = `Level 1: ${cur.level1 || '(kosong)'}`;
  $('c-l2').textContent = `Level 2: ${cur.level2 || '(kosong)'}${cur.kode2 ? ` (${cur.kode2})` : ''}`;
  $('c-l3').textContent = `Level 3: ${cur.level3 || '(kosong)'}`;
  let invalid = false;
  if (ready) {
    const check = checkClassification(cur, state.master);
    invalid = !check.ok;
    $('check').textContent = check.ok ? `✔ ${check.text}` : `✖ ${check.text}`;
    $('check').className = `check ${check.ok ? 'ok' : 'bad'}`;
  } else { $('check').textContent = 'Master belum dimuat.'; $('check').className = 'check'; }

  // Saran Level 3: hanya berguna bila klasifikasi tidak sesuai master.
  const box = $('suggestions'); box.replaceChildren();
  const suggestions = ready && invalid ? suggest(cell('nama'), cell('spesifikasi'), state.master, CONFIG.suggestionLimit) : [];
  $('suggest-card').hidden = !suggestions.length;
  for (const s of suggestions) for (const path of s.entries.slice(0, 3)) {
    const b = document.createElement('button'); b.type = 'button'; b.className = 'opt'; b.textContent = path.level3;
    const small = document.createElement('small'); small.textContent = `${path.level1} › ${path.level2}`; b.append(small);
    b.addEventListener('click', () => pickLevel3(path));
    box.append(b);
  }
  for (const id of ['valid', 'prev', 'next', 'next-invalid']) $(id).disabled = state.busy;
  renderEdit();
}

// ───────────── Menyimpan ─────────────

async function ensureNoteColumn(context, sheet, layout) {
  if ('note' in layout.correction) return layout.correction.note;
  const col = planCorrectionColumns(layout).plan.note;
  const head = sheet.getCell(layout.headerRow, col);
  head.values = [[CORRECTION_HEADERS.note]];
  head.format.font.bold = true; head.format.fill.color = '#FFF2CC';
  head.format.columnWidth = 220;
  await context.sync();
  return col;
}

async function appendLog(context, entries) {
  if (!entries.length) return;
  let sheet = context.workbook.worksheets.getItemOrNullObject(CONFIG.logSheet);
  sheet.load('isNullObject'); await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(CONFIG.logSheet);
    const head = sheet.getRangeByIndexes(0, 0, 1, LOG_HEADERS.length);
    head.values = [LOG_HEADERS]; head.format.font.bold = true; head.format.fill.color = '#D9EAD3';
    sheet.freezePanes.freezeRows(1);
    await context.sync();
  }
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load('isNullObject,rowIndex,rowCount'); await context.sync();
  const next = used.isNullObject ? 1 : used.rowIndex + used.rowCount;
  const rows = entries.map(e => [e.time, e.sheet, e.rowFrom, e.rowTo, e.action, e.column, safeText(e.old), safeText(e.new), safeText(e.l1), safeText(e.l2), safeText(e.l3), safeText(e.nama), safeText(e.note)]);
  sheet.getRangeByIndexes(next, 0, rows.length, LOG_HEADERS.length).values = rows;
  await context.sync();
}

const logBase = (sheetName, fields, note) => ({ time: formatTime(), sheet: sheetName, l1: fields.level1, l2: fields.level2, l3: fields.level3, nama: fields.nama, note: tidy(note) });

async function saveCorrection() {
  clearTimeout(noteTimer); pendingNote = null;
  const orig = { ...state.orig }, draft = { ...state.draft };
  const changes = diffFields(orig, draft);
  if (!changes.length) throw new Error('Tidak ada perubahan untuk disimpan.');
  const problem = pathProblem(draft, state.master);
  if (problem) throw new Error(problem);
  const move = $('move-row').checked, noteText = $('note').value;
  const target = { sheetId: state.sheetId, rowIndex: state.rowIndex };
  const order = state.order;

  const result = await Excel.run(async context => {
    const active = await readActiveRow(context);
    if (!active.layout || active.sheetId !== target.sheetId || active.rowIndex !== target.rowIndex) throw new Error('Baris aktif berubah. Pilih kembali baris yang ingin disimpan.');
    const now = fieldsOfRow(active.row, active.layout);
    if (COPY_FIELDS.some(f => f in active.layout.cols && tidy(now[f]) !== tidy(orig[f]))) throw new Error('Isi baris berubah sejak dibuka di panel. Pilih ulang baris ini untuk memuat nilai terbaru.');
    const layout = active.layout, sheet = context.workbook.worksheets.getItem(target.sheetId), row = target.rowIndex;
    const noteCol = await ensureNoteColumn(context, sheet, layout);
    const storedNote = noteOf(active.row, layout);

    for (const c of changes) {
      if (!(c.field in layout.cols)) continue;
      const range = sheet.getCell(row, layout.cols[c.field]);
      range.values = [[safeText(c.new)]];
      range.format.fill.color = CONFIG.changedFill;
    }
    if (tidy(noteText) !== storedNote) sheet.getCell(row, noteCol).values = [[safeText(noteText)]];
    await context.sync();

    let finalRow = row, moved = false;
    if (move) {
      const rows = await readIndex(context, sheet, layout, active.lastRow);
      const self = { row };
      const others = rows.filter(r => r.row !== row);
      const insertBefore = computeMoveTarget(others, self, orig, draft, order);
      if (insertBefore !== null) {
        const width = Math.max(layout.lastUsed, noteCol) + 1;
        sheet.getRangeByIndexes(insertBefore, 0, 1, 1).getEntireRow().insert(Excel.InsertShiftDirection.down);
        await context.sync();
        const srcIndex = row >= insertBefore ? row + 1 : row;
        sheet.getRangeByIndexes(insertBefore, 0, 1, width).copyFrom(sheet.getRangeByIndexes(srcIndex, 0, 1, width), Excel.RangeCopyType.all);
        await context.sync();
        sheet.getRangeByIndexes(srcIndex, 0, 1, 1).getEntireRow().delete(Excel.DeleteShiftDirection.up);
        await context.sync();
        finalRow = finalRowAfterMove(row, insertBefore); moved = true;
      }
    }

    const base = { ...logBase(active.sheetName, draft, noteText), rowFrom: row + 1, rowTo: finalRow + 1 };
    const entries = changes.map(c => ({ ...base, action: 'Koreksi', column: FIELD_LABELS[c.field], old: c.old, new: c.new }));
    if (moved) entries.push({ ...base, action: 'Pindah baris', column: '-', old: `Baris ${row + 1}`, new: `Baris ${finalRow + 1}` });
    await appendLog(context, entries);

    sheet.getCell(finalRow, layout.cols.nama).select();
    await context.sync();
    return { finalRow, moved };
  });
  state.index = null; state.rowKey = '';
  await refresh();
  tell(result.moved ? `Tersimpan. Baris dipindah ke baris ${result.finalRow + 1} (kelompok yang benar).` : `Tersimpan di baris ${result.finalRow + 1}.`);
  if ($('auto-next').checked) await jumpToInvalid();
}

async function markValid() {
  clearTimeout(noteTimer); pendingNote = null;
  const noteText = $('note').value, target = { sheetId: state.sheetId, rowIndex: state.rowIndex }, fields = { ...state.orig };
  await Excel.run(async context => {
    const active = await readActiveRow(context);
    if (!active.layout || active.sheetId !== target.sheetId || active.rowIndex !== target.rowIndex) throw new Error('Baris aktif berubah. Pilih kembali baris yang ingin ditandai.');
    const sheet = context.workbook.worksheets.getItem(target.sheetId);
    if (tidy(noteText) !== noteOf(active.row, active.layout)) {
      const col = await ensureNoteColumn(context, sheet, active.layout);
      sheet.getCell(target.rowIndex, col).values = [[safeText(noteText)]];
      await context.sync();
    }
    await appendLog(context, [{ ...logBase(active.sheetName, fields, noteText), rowFrom: target.rowIndex + 1, rowTo: target.rowIndex + 1, action: 'Valid', column: '-', old: '', new: '' }]);
  });
  tell(`Baris ${target.rowIndex + 1} ditandai valid (tercatat di sheet Log).`);
  if ($('auto-next').checked) await jumpToInvalid(); else await go(Math.min(target.rowIndex + 1, state.lastRow));
}

async function writeNote(p) {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(p.sheetId);
    sheet.load('name'); await context.sync();
    const layout = await loadLayout(context, sheet);
    if (!layout) return;
    const width = Math.max(layout.lastUsed, ...Object.values(layout.correction), 0) + 1;
    const rowRange = sheet.getRangeByIndexes(p.rowIndex, 0, 1, width);
    rowRange.load('values'); await context.sync();
    const stored = noteOf(rowRange.values[0], layout);
    if (tidy(p.text) === stored) return;
    const col = await ensureNoteColumn(context, sheet, layout);
    sheet.getCell(p.rowIndex, col).values = [[safeText(p.text)]];
    await context.sync();
    await appendLog(context, [{ ...logBase(sheet.name, fieldsOfRow(rowRange.values[0], layout), p.text), rowFrom: p.rowIndex + 1, rowTo: p.rowIndex + 1, action: 'Catatan', column: CORRECTION_HEADERS.note, old: stored, new: tidy(p.text) }]);
  });
  tell(`Catatan baris ${p.rowIndex + 1} tersimpan.`);
}
function flushNote() {
  clearTimeout(noteTimer);
  const p = pendingNote; pendingNote = null;
  if (p) enqueue(() => writeNote(p));
}

// ───────────── Navigasi ─────────────

async function go(rowIndex) {
  if (rowIndex < 0 || !state.layout) return;
  await Excel.run(async context => {
    context.workbook.worksheets.getItem(state.sheetId).getCell(rowIndex, state.layout.cols.nama).select();
    await context.sync();
  });
  await refresh();
}

async function jumpToInvalid() {
  if (!state.layout || !state.master) return;
  const from = state.rowIndex, layout = state.layout, sheetId = state.sheetId;
  const found = await Excel.run(async context => {
    const rows = await readIndex(context, context.workbook.worksheets.getItem(sheetId), layout, state.lastRow);
    return rows.find(r => r.row > from && !checkClassification(r, state.master).ok)?.row ?? -1;
  });
  if (found < 0) { tell('Tidak ada lagi baris yang tidak sesuai master di bawah baris ini.'); return; }
  await go(found);
}

async function reload() {
  state.master = await Excel.run(readMaster);
  state.order = masterOrder(state.master);
  state.index = null; state.rowKey = '';
  await refresh();
  tell(`Master_Level dimuat (${state.master.paths.length} baris).`);
}

// ───────────── Event ─────────────

$('save').addEventListener('click', () => enqueue(saveCorrection));
$('valid').addEventListener('click', () => enqueue(markValid));
$('reset').addEventListener('click', () => { if (state.orig) { state.draft = { ...state.orig }; syncInputs(); renderEdit(); } });
$('reload').addEventListener('click', () => enqueue(reload));
$('prev').addEventListener('click', () => enqueue(() => go(Math.max(state.rowIndex - 1, (state.layout?.headerRow ?? 0) + 1))));
$('next').addEventListener('click', () => enqueue(() => go(Math.min(state.rowIndex + 1, state.lastRow))));
$('next-invalid').addEventListener('click', () => enqueue(jumpToInvalid));
$('note').addEventListener('input', () => {
  if (!state.row) return;
  pendingNote = { sheetId: state.sheetId, rowIndex: state.rowIndex, text: $('note').value };
  clearTimeout(noteTimer); noteTimer = setTimeout(flushNote, 1200);
});
$('note').addEventListener('change', flushNote);

if (typeof Office === 'undefined') {
  tell('Office.js belum termuat. Periksa koneksi ke layanan Microsoft dan buka melalui Add-ins di Excel.', true);
} else {
  Office.onReady(info => {
    if (info.host !== Office.HostType.Excel) { tell('Panel harus dibuka sebagai add-in di Excel.', true); return; }
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.13')) { tell('Versi Excel ini belum mendukung ExcelApi 1.13. Gunakan Excel web atau perbarui Excel.', true); return; }
    $('connection').textContent = 'Terhubung ke Excel · Master_Level dibaca dari workbook aktif.';
    enqueue(async () => {
      await Excel.run(async context => {
        context.workbook.onSelectionChanged.add(() => {
          flushNote();
          clearTimeout(selectionTimer);
          selectionTimer = setTimeout(() => enqueue(refresh), 150);
        });
        await context.sync();
      });
      await reload().catch(async error => { tell(failure(error), true); await refresh(); });
    });
  }).catch(error => tell(failure(error), true));
}
