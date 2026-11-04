import test from 'node:test';
import assert from 'node:assert/strict';
import {
  checkClassification, computeMoveTarget, detectLayout, diffFields, finalRowAfterMove, findLevel2, level1List, level2Entries,
  masterOrder, parseMaster, pathProblem, planCorrectionColumns, safeText, suggest, valueOptions
} from '../public/core.js';

const masterRows = [
  ['Level 1', 'Kode Level 2', 'Level 2', 'Level 3', 'Inisial'],
  ['Supplier', 'SUP-01', 'Material Semen, Beton & Grout', 'Semen', ''],
  ['', '', '', 'Ready Mix', ''],
  ['Supplier', 'SUP-02', 'Baja & Besi', 'Besi Beton', ''],
  ['Supplier', 'SUP-05', 'Produk Beton Pracetak & Prategang', 'Tiang Pancang', '']
];
const master = parseMaster(masterRows);

test('parseMaster mengisi induk yang kosong', () => {
  assert.equal(master.paths.length, 4);
  assert.equal(master.byLevel3.get('ready mix')[0].kode2, 'SUP-01');
  assert.deepEqual(level1List(master), ['Supplier']);
  assert.equal(level2Entries(master).length, 3);
  assert.equal(findLevel2(master, 'Supplier', 'Baja & Besi').kode2, 'SUP-02');
});

test('checkClassification dan pathProblem', () => {
  assert.equal(checkClassification({ level1: 'Supplier', level2: 'Material Semen, Beton & Grout', level3: 'Semen' }, master).ok, true);
  assert.equal(checkClassification({ level1: 'Supplier', level2: 'Baja & Besi', level3: 'Semen' }, master).code, 'induk-salah');
  assert.equal(pathProblem({ level1: 'Supplier', level2: 'Baja & Besi', level3: '' }, master).includes('Level 3'), true);
  assert.equal(pathProblem({ level1: 'Supplier', level2: 'Baja & Besi', level3: 'Semen' }, master).includes('tidak ada'), true);
  assert.equal(pathProblem({ level1: 'Supplier', level2: 'Baja & Besi', level3: 'Besi Beton' }, master), '');
});

test('detectLayout DG, DSU, dan kolom catatan', () => {
  const dg = detectLayout([['Tanggal', 'Level 1', 'Kode Level 2', 'Level 2', 'Level 3', 'Nama Produk / Pekerjaan', 'Spesifikasi', 'Satuan', 'Merek', 'Tahun Sumber']]);
  assert.equal(dg.headerRow, 0); assert.equal(dg.cols.kode2, 2); assert.equal(dg.lastUsed, 9);
  const { plan, created } = planCorrectionColumns(dg);
  assert.deepEqual(created, ['note']); assert.equal(plan.note, 10);
  const dsu = detectLayout([[], [], ['Nama Vendor', 'Tanggal', 'Level 1', 'Kode Level 2', 'Level 2', 'Level 3', 'Nama produk / pekerjaan', 'Level 4: Spesifikasi teknis', 'Level 5: Satuan', 'Level 6: Merek 1', 'Catatan Koreksi']]);
  assert.equal(dsu.headerRow, 2); assert.equal(dsu.cols.spesifikasi, 7); assert.equal(dsu.correction.note, 10); assert.equal(dsu.hasAllCorrection, true);
  assert.equal(dsu.lastCol, 9);
  assert.equal(detectLayout([['A', 'B']]), null);
});

test('diffFields hanya melaporkan kolom yang berubah', () => {
  const o = { level1: 'Supplier', kode2: 'SUP-05', level2: 'Produk Beton', level3: 'Tiang', nama: 'A', spesifikasi: 'x' };
  const d = { ...o, level2: 'Semen', kode2: 'SUP-01' };
  assert.deepEqual(diffFields(o, d).map(c => c.field), ['kode2', 'level2']);
  assert.deepEqual(diffFields(o, o), []);
});

// Baris data mulai indeks 1 (judul di indeks 0).
const R = (row, level2, level3, nama, spesifikasi = '') => ({ row, level1: 'Supplier', level2, level3, nama, spesifikasi });
const S = 'Material Semen, Beton & Grout', B = 'Baja & Besi', P = 'Produk Beton Pracetak & Prategang';
const order = masterOrder(master);

test('computeMoveTarget: perubahan spesifikasi pindah ke bawah baris terakhir spesifikasi tersebut', () => {
  const rows = [
    R(2, S, 'Semen', 'Semen', 'Merah Putih'),
    R(3, S, 'Semen', 'Semen', '1; Tipe 1'),
    R(5, S, 'Semen', 'Semen', 'OPC'),
    R(6, S, 'Semen', 'Semen', 'PCC; Curah'),
    R(7, S, 'Semen', 'Semen', 'PCC; Curah'),
    R(8, S, 'Semen', 'Semen', 'PCC; Curah'),
    R(9, S, 'Semen', 'Semen', 'PCC; Curah'),
    R(10, S, 'Semen', 'Semen', 'PCC; Jumbo Bag')
  ];
  const orig = { level1: 'Supplier', level2: S, level3: 'Semen', nama: 'Semen', spesifikasi: 'Merah Putih' };
  const upd = { level1: 'Supplier', level2: S, level3: 'Semen', nama: 'Semen', spesifikasi: 'PCC; Curah' };
  const target = computeMoveTarget(rows, { row: 4 }, orig, upd, order);
  assert.equal(target, 10); // disisipkan sebelum baris 10
  assert.equal(finalRowAfterMove(4, target), 9); // posisi akhir setelah pindah adalah baris 9 (bawah baris 9 lama)
});

test('computeMoveTarget: Level 3 baru di Level 2 yang sama mengikuti urutan master', () => {
  const rows = [R(1, S, 'Semen', 'Semen'), R(2, S, 'Semen', 'Semen'), R(3, B, 'Besi Beton', 'Besi'), R(5, P, 'Tiang Pancang', 'Tiang')];
  const orig = { level1: 'Supplier', level2: P, level3: 'Tiang Pancang', nama: 'Beton' };
  const upd = { level1: 'Supplier', level2: S, level3: 'Ready Mix', nama: 'Beton' };
  assert.equal(computeMoveTarget(rows, { row: 4 }, orig, upd, order), 3); // setelah baris Semen terakhir (2)
});

test('computeMoveTarget: tidak pindah bila sudah di posisi paling bawah kelompoknya', () => {
  const rows = [R(1, S, 'Semen', 'Semen', 'OPC')];
  const o = { level1: 'Supplier', level2: S, level3: 'Semen', nama: 'Semen', spesifikasi: 'OPC' };
  assert.equal(computeMoveTarget(rows, { row: 2 }, o, { ...o }, order), null);
});

test('valueOptions memprioritaskan kelompok yang sama', () => {
  const rows = [
    { level3: 'Semen', nama: 'Semen', spesifikasi: 'OPC' }, { level3: 'Semen', nama: 'Semen', spesifikasi: 'OPC' },
    { level3: 'Ready Mix', nama: 'Beton', spesifikasi: 'fc 30' }
  ];
  const draft = { level3: 'Semen', nama: 'Semen' };
  assert.deepEqual(valueOptions(rows, 'nama', draft, '', true).map(o => o.label), ['Semen']);
  assert.deepEqual(valueOptions(rows, 'nama', draft, 'bet', false).map(o => o.label), ['Beton']);
  assert.equal(valueOptions(rows, 'spesifikasi', draft, '', true)[0].label, 'OPC');
});

test('suggest dan safeText', () => {
  assert.equal(suggest('Semen Portland Tipe 1', 'Kemasan 40 kg', master)[0].label, 'Semen');
  assert.equal(safeText('=1+1'), "'=1+1");
});
