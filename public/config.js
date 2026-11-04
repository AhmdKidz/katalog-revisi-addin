export const CONFIG = Object.freeze({
  masterSheet: 'Master_Level',
  logSheet: 'Log',
  headerSearchRows: 20,   // baris awal yang dipindai untuk mencari judul kolom
  maxHeaderColumns: 80,
  maxMasterRows: 20000,
  maxMasterColumns: 12,
  suggestionLimit: 6,
  optionLimit: 60,
  changedFill: '#FFFF00'   // warna sel yang sudah dikoreksi
});

// Satu-satunya kolom baru di sisi kanan tabel.
export const CORRECTION_HEADERS = Object.freeze({
  note: 'Catatan Koreksi'
});

// Kolom katalog yang boleh dikoreksi (kunci = nama kolom pada tata letak).
export const FIELD_LABELS = Object.freeze({
  level1: 'Level 1',
  kode2: 'Kode Level 2',
  level2: 'Level 2',
  level3: 'Level 3',
  nama: 'Nama Produk / Pekerjaan',
  spesifikasi: 'Spesifikasi'
});

export const LOG_HEADERS = Object.freeze([
  'Waktu', 'Sheet', 'Baris Awal', 'Baris Akhir', 'Aksi', 'Kolom', 'Nilai Lama', 'Nilai Baru',
  'Level 1 (akhir)', 'Level 2 (akhir)', 'Level 3 (akhir)', 'Nama Produk (akhir)', 'Catatan'
]);
