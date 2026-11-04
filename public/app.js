/* global Office, Excel */
import { CONFIG, CORRECTION_HEADERS, LEVEL_TARGETS, STATUS } from './config.js';
import { buildCorrection, checkClassification, detectLayout, filterPaths, keyOf, nextPending, parseMaster, planCorrectionColumns, safeText, suggest, tidy } from './core.js';

const $ = id => document.getElementById(id);
const state = { master: null, layout: null, sheetId: null, sheetName: '', lastRow: 0, rowIndex: -1, row: null, picked: null, rowKey: '', busy: false };
let queue = Promise.resolve(), selectionTimer;

function tell(text, error = false) { $('message').textContent = text; $('message').classList.toggle('error', error); }
function failure(error) {
  const code = error?.code || '';
  if (['AccessDenied', 'ProtectedRange', 'InvalidOperationInCellEditMode'].includes(code)) return 'Excel menolak perubahan. Pastikan sel dapat diedit dan Anda tidak sedang dalam mode mengetik di sel.';
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

/** Membaca seleksi aktif, tata letak sheet, dan isi baris yang dipilih. */
async function readActiveRow(context) {
  const range = context.workbook.getSelectedRange();
  const sheet = range.worksheet;
  range.load('rowIndex,rowCount');
  sheet.load('id,name');
  const used = sheet.getUsedRange(true);
  used.load('rowIndex,rowCount');
  const head = sheet.getRangeByIndexes(0, 0, CONFIG.headerSearchRows, CONFIG.maxHeaderColumns);
  head.load('values');
  await context.sync();
  const layout = detectLayout(head.values, CONFIG.headerSearchRows);
  const info = { sheetId: sheet.id, sheetName: sheet.name, layout, lastRow: used.rowIndex + used.rowCount - 1, rowIndex: range.rowIndex, multi: range.rowCount > 1, row: null };
  if (!layout || range.rowIndex <= layout.headerRow) return info;
  const width = Math.max(layout.lastUsed, ...Object.values(layout.correction), 0) + 1;
  const rowRange = sheet.getRangeByIndexes(range.rowIndex, 0, 1, width);
  rowRange.load('values'); await context.sync();
  info.row = rowRange.values[0];
  return info;
}

async function refresh() {
  const info = await Excel.run(readActiveRow);
  const key = `${info.sheetId}:${info.rowIndex}`;
  Object.assign(state, { layout: info.layout, sheetId: info.sheetId, sheetName: info.sheetName, lastRow: info.lastRow, rowIndex: info.rowIndex, row: info.row, multi: info.multi });
  if (key !== state.rowKey) { // baris berbeda: reset formulir
    state.rowKey = key; state.picked = null;
    $('level3-search').value = ''; $('note').value = '';
    $('level-target').value = LEVEL_TARGETS[0];
    const existing = state.row && state.layout?.correction;
    if (existing) {
      const get = name => (name in state.layout.correction ? tidy(state.row[state.layout.correction[name]]) : '');
      if (get('levelTarget')) $('level-target').value = get('levelTarget');
      $('note').value = get('note');
    }
  }
}

const cell = (name) => (state.row && state.layout && name in state.layout.cols ? tidy(state.row[state.layout.cols[name]]) : '');
const corrCell = (name) => (state.row && state.layout && name in state.layout.correction ? tidy(state.row[state.layout.correction[name]]) : '');

function optionButton(path, onPick, extra) {
  const b = document.createElement('button');
  b.type = 'button'; b.className = 'opt';
  b.textContent = path.level3;
  const small = document.createElement('small');
  small.textContent = `${path.level1} › ${path.level2}${extra ? ` · ${extra}` : ''}`;
  b.append(small);
  b.addEventListener('click', () => onPick(path));
  return b;
}

function pick(path) { state.picked = path; $('level-target').value = state.picked ? 'Level 3' : LEVEL_TARGETS[0]; render(); }

function render() {
  const ready = Boolean(state.master);
  $('count').textContent = ready ? `${state.master.paths.length} baris master` : '';
  $('reload').disabled = state.busy;
  const layout = state.layout;
  const rowOk = Boolean(layout && state.row);
  const needsSetup = Boolean(layout && !layout.hasAllCorrection);
  if (!layout) {
    $('setup').hidden = false; $('prepare').hidden = true;
    $('setup-text').textContent = `Sheet "${state.sheetName || 'aktif'}" tidak dikenali sebagai tabel katalog. Pastikan baris judul memuat Level 1, Level 2, Level 3, dan Nama Produk / Pekerjaan, lalu pilih sebuah sel di sheet tersebut.`;
  } else if (needsSetup) {
    $('setup').hidden = false; $('prepare').hidden = false;
    $('setup-text').textContent = `Kolom koreksi belum lengkap di sheet "${state.sheetName}". Tombol di bawah menambahkan judul yang kurang di sisi kanan tabel (data asli tidak diubah).`;
  } else $('setup').hidden = true;

  $('row-card').hidden = !rowOk;
  if (!rowOk) {
    if (layout && state.multi) tell('Pilih satu baris data saja.', true);
    return;
  }
  const rowNumber = state.rowIndex + 1;
  $('row-address').textContent = `${state.sheetName} · baris ${rowNumber}`;
  $('r-nama').textContent = cell('nama') || '(kosong)';
  $('r-spek').textContent = cell('spesifikasi') || '-';
  $('r-misc').textContent = [cell('satuan') || '-', cell('merek') || '-'].join(' · ');
  const cur = { level1: cell('level1'), level2: cell('level2'), level3: cell('level3') };
  $('c-l1').textContent = `Level 1: ${cur.level1 || '(kosong)'}`;
  $('c-l2').textContent = `Level 2: ${cur.level2 || '(kosong)'}`;
  $('c-l3').textContent = `Level 3: ${cur.level3 || '(kosong)'}`;
  if (ready) {
    const check = checkClassification(cur, state.master);
    $('check').textContent = check.ok ? `✔ ${check.text}` : `✖ ${check.text}`;
    $('check').className = `check ${check.ok ? 'ok' : 'bad'}`;
  } else { $('check').textContent = 'Master belum dimuat.'; $('check').className = 'check'; }

  const status = corrCell('status');
  $('row-status').textContent = status || 'Belum ditinjau';
  $('row-status').className = `badge ${keyOf(status).replace(/\s+/g, '') === 'valid' ? 'valid' : status === STATUS.corrected ? 'dikoreksi' : status === STATUS.review ? 'review' : ''}`;
  const savedPath = [corrCell('level1'), corrCell('level2'), corrCell('level3')].filter(Boolean).join(' › ');
  $('existing').textContent = savedPath ? `Koreksi tersimpan: ${savedPath}` : '';

  // Saran
  const box = $('suggestions'); box.replaceChildren();
  const suggestions = ready ? suggest(cell('nama'), cell('spesifikasi'), state.master, CONFIG.suggestionLimit) : [];
  $('suggest-card').hidden = !suggestions.length;
  for (const s of suggestions) for (const path of s.entries.slice(0, 3)) box.append(optionButton(path, pick));

  // Hasil pencarian
  const results = $('results'); results.replaceChildren();
  const term = $('level3-search').value;
  if (ready && tidy(term)) for (const path of filterPaths(state.master, term)) results.append(optionButton(path, pick));

  const picked = $('picked'); picked.replaceChildren(); picked.hidden = !state.picked;
  if (state.picked) {
    const text = document.createElement('span');
    text.textContent = `Level 3 baru: ${state.picked.level3} (${state.picked.level1} › ${state.picked.level2})`;
    const clear = document.createElement('button'); clear.type = 'button'; clear.textContent = 'Batal';
    clear.addEventListener('click', () => { state.picked = null; render(); });
    picked.append(text, clear);
  }
  const lock = state.busy || !ready || needsSetup;
  for (const id of ['save-corrected', 'save-valid', 'save-review', 'prev', 'next', 'next-pending']) $(id).disabled = lock && id.startsWith('save') ? true : state.busy;
}

async function prepareColumns() {
  await Excel.run(async context => {
    const info = await readActiveRow(context);
    if (!info.layout) throw new Error('Tata letak katalog tidak dikenali pada sheet aktif.');
    const { plan, created } = planCorrectionColumns(info.layout);
    const sheet = context.workbook.worksheets.getItem(info.sheetId);
    for (const name of created) {
      const head = sheet.getCell(info.layout.headerRow, plan[name]);
      head.values = [[CORRECTION_HEADERS[name]]];
      head.format.font.bold = true; head.format.fill.color = '#FFF2CC';
      sheet.getRangeByIndexes(info.layout.headerRow, plan[name], 1, 1).format.columnWidth = name === 'note' ? 220 : 130;
    }
    const first = info.layout.headerRow + 1, count = Math.max(info.lastRow - info.layout.headerRow, 1);
    const lists = { status: Object.values(STATUS).join(','), levelTarget: LEVEL_TARGETS.join(',') };
    for (const [name, source] of Object.entries(lists)) {
      const col = sheet.getRangeByIndexes(first, plan[name], count, 1);
      col.dataValidation.clear();
      col.dataValidation.rule = { list: { inCellDropDown: true, source } };
    }
    await context.sync();
  });
  state.rowKey = '';
  await refresh();
  tell('Kolom koreksi siap di sisi kanan tabel.');
}

async function writeRow(statusLabel, requireContent) {
  const layout = state.layout;
  if (!layout?.hasAllCorrection) throw new Error('Siapkan kolom koreksi terlebih dahulu.');
  const isValid = statusLabel === STATUS.valid;
  const levelTarget = isValid ? LEVEL_TARGETS[0] : $('level-target').value;
  const note = $('note').value;
  if (requireContent) {
    const meaningful = state.picked || (levelTarget && levelTarget !== LEVEL_TARGETS[0]) || tidy(note);
    if (!meaningful) throw new Error('Pilih Level 3 yang benar, tentukan level seharusnya, atau isi catatan sebelum menyimpan koreksi.');
  }
  const values = buildCorrection({ status: statusLabel, levelTarget, path: isValid ? null : state.picked, note });
  const target = { sheetId: state.sheetId, rowIndex: state.rowIndex };
  await Excel.run(async context => {
    const active = await readActiveRow(context);
    if (active.sheetId !== target.sheetId || active.rowIndex !== target.rowIndex) throw new Error('Baris aktif berubah. Pilih kembali baris yang ingin disimpan.');
    const sheet = context.workbook.worksheets.getItem(target.sheetId);
    for (const [name, col] of Object.entries(layout.correction)) {
      sheet.getCell(target.rowIndex, col).values = [[safeText(values[name])]];
    }
    await context.sync();
  });
  tell(`Baris ${target.rowIndex + 1}: ${statusLabel}.`);
}

async function go(rowIndex) {
  if (rowIndex < 0) return;
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(state.sheetId);
    const colIndex = state.layout.cols.nama;
    sheet.getCell(rowIndex, colIndex).select();
    await context.sync();
  });
  await refresh();
}

async function findPending() {
  const layout = state.layout;
  if (!layout?.hasAllCorrection) throw new Error('Siapkan kolom koreksi terlebih dahulu.');
  const startRow = Math.max(state.rowIndex, layout.headerRow);
  const count = state.lastRow - startRow;
  if (count <= 0) return -1;
  const values = await Excel.run(async context => {
    const range = context.workbook.worksheets.getItem(state.sheetId).getRangeByIndexes(startRow + 1, layout.correction.status, count, 1);
    range.load('values'); await context.sync();
    return range.values.map(r => r[0]);
  });
  const idx = nextPending(values, -1);
  return idx < 0 ? -1 : startRow + 1 + idx;
}

async function afterSave() {
  if ($('auto-next').checked) {
    const row = await findPending();
    if (row >= 0) { await go(row); return; }
    tell('Selesai: tidak ada lagi baris yang belum ditinjau di bawah baris ini.');
  }
  await refresh();
}

async function save(label, requireContent) { await writeRow(label, requireContent); await afterSave(); }

async function reload() {
  const master = await Excel.run(readMaster);
  state.master = master;
  await refresh();
  tell(`Master_Level dimuat (${master.paths.length} baris).`);
}

for (const option of LEVEL_TARGETS) $('level-target').add(new Option(option, option));
$('level3-search').addEventListener('input', render);
$('save-corrected').addEventListener('click', () => enqueue(() => save(STATUS.corrected, true)));
$('save-valid').addEventListener('click', () => enqueue(() => save(STATUS.valid, false)));
$('save-review').addEventListener('click', () => enqueue(() => save(STATUS.review, false)));
$('prepare').addEventListener('click', () => enqueue(prepareColumns));
$('reload').addEventListener('click', () => enqueue(reload));
$('prev').addEventListener('click', () => enqueue(() => go(Math.max(state.rowIndex - 1, (state.layout?.headerRow ?? 0) + 1))));
$('next').addEventListener('click', () => enqueue(() => go(Math.min(state.rowIndex + 1, state.lastRow))));
$('next-pending').addEventListener('click', () => enqueue(async () => {
  const row = await findPending();
  if (row < 0) tell('Tidak ada baris yang belum ditinjau di bawah baris ini.'); else await go(row);
}));

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
          clearTimeout(selectionTimer);
          selectionTimer = setTimeout(() => enqueue(refresh), 150);
        });
        await context.sync();
      });
      await reload().catch(async error => { tell(failure(error), true); await refresh(); });
    });
  }).catch(error => tell(failure(error), true));
}
