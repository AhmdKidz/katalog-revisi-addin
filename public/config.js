export const CONFIG = Object.freeze({
  masterSheet: 'Master_Level',
  headerSearchRows: 20,   // baris awal yang dipindai untuk mencari judul kolom
  maxHeaderColumns: 80,
  maxMasterRows: 20000,
  maxMasterColumns: 12,
  suggestionLimit: 6
});

// Judul kolom koreksi yang ditambahkan di sisi kanan tabel.
export const CORRECTION_HEADERS = Object.freeze({
  status: 'Koreksi: Status',
  levelTarget: 'Koreksi: Seharusnya di Level',
  level1: 'Koreksi: Level 1',
  kode2: 'Koreksi: Kode Level 2',
  level2: 'Koreksi: Level 2',
  level3: 'Koreksi: Level 3',
  note: 'Koreksi: Catatan'
});

export const STATUS = Object.freeze({
  valid: 'Valid',
  corrected: 'Dikoreksi',
  review: 'Perlu Review'
});

// Pilihan "nama produk/pekerjaan ini seharusnya berada di level mana".
export const LEVEL_TARGETS = Object.freeze([
  'Sudah sesuai',
  'Level 1',
  'Level 2',
  'Level 3',
  'Level 4 (Spesifikasi)',
  'Level 5 (Satuan)',
  'Level 6 (Merek)',
  'Bukan produk/pekerjaan (hapus)',
  'Belum ada di master (usulan baru)'
]);
