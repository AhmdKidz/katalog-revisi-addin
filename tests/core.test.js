import test from 'node:test';
import assert from 'node:assert/strict';
import { buildCorrection, checkClassification, detectLayout, nextPending, parseMaster, planCorrectionColumns, safeText, suggest } from '../public/core.js';

const masterRows = [
  ['Level 1', 'Kode Level 2', 'Level 2', 'Level 3', 'Inisial'],
  ['Supplier', 'SUP-01', 'Material Semen, Beton & Grout', 'Semen', ''],
  ['', '', '', 'Ready Mix', ''],
  ['Supplier', 'SUP-02', 'Baja & Besi', 'Besi Beton', '']
];

test('parseMaster mengisi induk yang kosong', () => {
  const m = parseMaster(masterRows);
  assert.equal(m.paths.length, 3);
  assert.equal(m.byLevel3.get('ready mix')[0].level2, 'Material Semen, Beton & Grout');
  assert.equal(m.byLevel3.get('ready mix')[0].kode2, 'SUP-01');
});

test('checkClassification: sesuai, induk salah, tidak ada, kosong', () => {
  const m = parseMaster(masterRows);
  assert.equal(checkClassification({ level1: 'Supplier', level2: 'Material Semen, Beton & Grout', level3: 'Semen' }, m).ok, true);
  assert.equal(checkClassification({ level1: 'Supplier', level2: 'Baja & Besi', level3: 'Semen' }, m).code, 'induk-salah');
  assert.equal(checkClassification({ level1: 'Supplier', level2: 'x', level3: 'Pasir' }, m).code, 'l3-tidak-ada');
  assert.equal(checkClassification({}, m).code, 'kosong');
});

test('detectLayout pada format DG (judul baris 1)', () => {
  const rows = [['Tanggal', 'Level 1', 'Kode Level 2', 'Level 2', 'Level 3', 'Nama Produk / Pekerjaan', 'Spesifikasi', 'Satuan', 'Merek', 'Tahun Sumber']];
  const l = detectLayout(rows);
  assert.equal(l.headerRow, 0);
  assert.equal(l.cols.nama, 5);
  assert.equal(l.cols.level3, 4);
  assert.equal(l.lastUsed, 9);
  assert.equal(l.hasAllCorrection, false);
  const { plan, created } = planCorrectionColumns(l);
  assert.equal(plan.status, 10);
  assert.equal(created.length, 7);
});

test('detectLayout pada format DSU rapih (judul baris 3, banyak kolom)', () => {
  const rows = [[], [], ['Nama Vendor', 'Tanggal', 'Level 1', 'Kode Level 2', 'Level 2', 'Level 3', 'Nama produk / pekerjaan', 'Level 4: Spesifikasi teknis', 'Level 5: Satuan', 'Level 6: Merek 1']];
  const l = detectLayout(rows);
  assert.equal(l.headerRow, 2);
  assert.equal(l.cols.spesifikasi, 7);
  assert.equal(l.cols.satuan, 8);
  assert.equal(l.cols.merek, 9);
});

test('detectLayout mengenali kolom koreksi yang sudah ada', () => {
  const headers = ['Level 1', 'Level 2', 'Level 3', 'Nama Produk', 'Koreksi: Status', 'Koreksi: Catatan'];
  const l = detectLayout([headers]);
  assert.equal(l.correction.status, 4);
  assert.equal(l.lastCol, 3);
  const { plan, created } = planCorrectionColumns(l);
  assert.equal(plan.status, 4);
  assert.equal(plan.levelTarget, 6);
  assert.equal(created.includes('status'), false);
});

test('detectLayout null bila judul wajib tidak ada', () => {
  assert.equal(detectLayout([['A', 'B']]), null);
});

test('suggest memberi Level 3 yang cocok dengan nama produk', () => {
  const m = parseMaster(masterRows);
  const s = suggest('Semen Portland Tipe 1', 'Kemasan 40 kg', m);
  assert.equal(s[0].label, 'Semen');
});

test('safeText, buildCorrection, nextPending', () => {
  assert.equal(safeText('=1+1'), "'=1+1");
  assert.equal(safeText('aman'), 'aman');
  const c = buildCorrection({ status: 'Dikoreksi', levelTarget: 'Level 3', path: { level1: 'A', kode2: 'K', level2: 'B', level3: 'C' }, note: '  hai   ' });
  assert.equal(c.note, 'hai');
  assert.equal(c.level3, 'C');
  assert.equal(nextPending(['Valid', 'Dikoreksi', '', 'Valid'], -1), 2);
  assert.equal(nextPending(['Valid'], -1), -1);
});
