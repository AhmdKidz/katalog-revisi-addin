import { CORRECTION_HEADERS } from './config.js';

export const tidy = value => String(value ?? '').trim().replace(/\s+/g, ' ');
export const keyOf = value => tidy(value).normalize('NFKC').toLocaleLowerCase('id-ID');
export const headerKey = value => keyOf(value).replace(/[^a-z0-9]/g, '');

const MATCHERS = {
  level1: k => k === 'level1',
  kode2: k => k === 'kodelevel2',
  level2: k => k === 'level2',
  level3: k => k === 'level3',
  nama: k => k.startsWith('namaproduk') || k.startsWith('namapekerjaan'),
  spesifikasi: k => k === 'spesifikasi' || k.startsWith('level4spesifikasi'),
  satuan: k => k === 'satuan' || k.startsWith('level5satuan'),
  merek: k => k === 'merek' || k.startsWith('level6merek') || k === 'merek1'
};
const REQUIRED = ['level1', 'level2', 'level3', 'nama'];

/**
 * Mencari baris judul dan posisi kolom pada sheet katalog.
 * `rows` = nilai sel mulai dari baris pertama sheet (indeks 0 = baris 1).
 * Mengembalikan null bila judul wajib tidak ditemukan.
 */
export function detectLayout(rows, headerSearchRows = 20) {
  const correctionKeys = Object.fromEntries(Object.entries(CORRECTION_HEADERS).map(([name, label]) => [name, headerKey(label)]));
  for (let r = 0; r < Math.min(rows.length, headerSearchRows); r++) {
    const keys = rows[r].map(headerKey);
    const cols = {};
    for (const [name, test] of Object.entries(MATCHERS)) {
      const index = keys.findIndex(k => k && test(k));
      if (index >= 0) cols[name] = index;
    }
    if (!REQUIRED.every(name => name in cols)) continue;
    const correction = {};
    for (const [name, key] of Object.entries(correctionKeys)) {
      const index = keys.indexOf(key);
      if (index >= 0) correction[name] = index;
    }
    let lastCol = -1;
    keys.forEach((k, i) => { if (k && !Object.values(correction).includes(i)) lastCol = Math.max(lastCol, i); });
    const corrIdx = Object.values(correction);
    // Kolom koreksi baru dibuat tepat di kanan kolom terakhir (atau kolom koreksi yang sudah ada).
    const lastUsed = Math.max(lastCol, ...corrIdx, -1);
    return { headerRow: r, cols, correction, lastCol, lastUsed, hasAllCorrection: Object.keys(CORRECTION_HEADERS).every(n => n in correction) };
  }
  return null;
}

/** Posisi kolom koreksi yang akan dipakai (yang sudah ada dipertahankan, sisanya ditambahkan di kanan). */
export function planCorrectionColumns(layout) {
  const plan = { ...layout.correction };
  let next = layout.lastUsed + 1;
  const created = [];
  for (const name of Object.keys(CORRECTION_HEADERS)) {
    if (!(name in plan)) { plan[name] = next++; created.push(name); }
  }
  return { plan, created };
}

/**
 * Membaca Master_Level. Kolom induk yang kosong diisi dari baris di atasnya
 * (format sel digabung / diturunkan). Mengembalikan { paths, byLevel3 }.
 */
export function parseMaster(values, headerSearchRows = 20) {
  let headerRow = -1, idx;
  for (let r = 0; r < Math.min(values.length, headerSearchRows); r++) {
    const keys = values[r].map(headerKey);
    const found = { level1: keys.indexOf('level1'), kode2: keys.indexOf('kodelevel2'), level2: keys.indexOf('level2'), level3: keys.indexOf('level3') };
    if (found.level1 >= 0 && found.level2 >= 0 && found.level3 >= 0) { headerRow = r; idx = found; break; }
  }
  if (headerRow < 0) throw new Error('Header Level 1, Level 2, dan Level 3 tidak ditemukan pada Master_Level.');
  const paths = [], byLevel3 = new Map();
  let p1 = '', pk = '', p2 = '';
  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    const one = tidy(row[idx.level1]), kode = idx.kode2 >= 0 ? tidy(row[idx.kode2]) : '', two = tidy(row[idx.level2]), three = tidy(row[idx.level3]);
    if (one && keyOf(one) !== keyOf(p1)) { p1 = one; pk = ''; p2 = ''; }
    if (two) { p2 = two; if (kode) pk = kode; }
    else if (kode) pk = kode;
    if (!three || three.startsWith('#')) continue;
    const entry = { level1: p1, kode2: pk, level2: p2, level3: three, id: keyOf(three) };
    paths.push(entry);
    if (!byLevel3.has(entry.id)) byLevel3.set(entry.id, []);
    byLevel3.get(entry.id).push(entry);
  }
  if (!paths.length) throw new Error('Master_Level belum memiliki baris Level 3.');
  return { paths, byLevel3 };
}

/** Memeriksa klasifikasi yang sudah ada pada sebuah baris terhadap master. */
export function checkClassification({ level1, level2, level3 }, master) {
  const l1 = tidy(level1), l2 = tidy(level2), l3 = tidy(level3);
  if (!l1 && !l2 && !l3) return { ok: false, code: 'kosong', text: 'Klasifikasi masih kosong.' };
  if (!l3) return { ok: false, code: 'l3-kosong', text: 'Level 3 kosong.' };
  const candidates = master.byLevel3.get(keyOf(l3));
  if (!candidates) return { ok: false, code: 'l3-tidak-ada', text: `Level 3 "${l3}" tidak ada di Master_Level.` };
  const match = candidates.some(c => keyOf(c.level1) === keyOf(l1) && keyOf(c.level2) === keyOf(l2));
  if (!match) return { ok: false, code: 'induk-salah', text: `Level 3 "${l3}" ada di master, tetapi induk Level 1/2 tidak cocok.`, expected: candidates };
  return { ok: true, code: 'sesuai', text: 'Sesuai dengan Master_Level.' };
}

const tokens = text => keyOf(text).split(/[^a-z0-9]+/).filter(t => t.length > 2);

/** Saran Level 3 berdasarkan kemiripan kata dengan nama produk dan spesifikasi. */
export function suggest(nama, spesifikasi, master, limit = 6) {
  const nameKey = keyOf(nama), all = `${nameKey} ${keyOf(spesifikasi)}`;
  const nameTokens = new Set(tokens(nama)), allTokens = new Set(tokens(all));
  const scored = [];
  for (const [id, entries] of master.byLevel3) {
    const label = entries[0].level3, labelTokens = tokens(label);
    if (!labelTokens.length && !id) continue;
    let score = 0;
    if (nameKey && id && nameKey.includes(id)) score += 3 + id.length / 20;
    if (nameKey && id && nameKey === id) score += 4;
    if (labelTokens.length) {
      const inName = labelTokens.filter(t => nameTokens.has(t)).length;
      const inAll = labelTokens.filter(t => allTokens.has(t)).length;
      score += (inName / labelTokens.length) * 2 + (inAll / labelTokens.length);
    }
    if (score > 0.9) scored.push({ id, label, entries, score });
  }
  return scored.sort((a, b) => b.score - a.score || a.label.localeCompare(b.label, 'id')).slice(0, limit);
}

/** Nilai baris yang akan ditulis ke kolom koreksi. */
export function buildCorrection({ status, levelTarget, path, note }) {
  return {
    status,
    levelTarget: levelTarget || '',
    level1: path?.level1 || '',
    kode2: path?.kode2 || '',
    level2: path?.level2 || '',
    level3: path?.level3 || '',
    note: tidy(note)
  };
}

/** Menolak teks yang akan ditafsirkan Excel sebagai rumus. */
export function safeText(value) {
  const text = String(value ?? '');
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

/** Mencari indeks baris berikutnya (setelah `from`) yang kolom statusnya masih kosong. */
export function nextPending(statusValues, from) {
  for (let i = from + 1; i < statusValues.length; i++) if (!tidy(statusValues[i])) return i;
  return -1;
}

export function filterPaths(master, term, limit = 50) {
  const k = keyOf(term);
  const out = [];
  for (const p of master.paths) {
    if (!k || keyOf(`${p.level3} ${p.level2} ${p.level1}`).includes(k)) out.push(p);
    if (out.length >= limit) break;
  }
  return out;
}
