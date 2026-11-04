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

// ───────────────────────── v2: koreksi langsung pada baris ─────────────────────────

const EDIT_ORDER = ['level1', 'kode2', 'level2', 'level3', 'nama', 'spesifikasi'];

export function level1List(master) {
  const seen = new Map();
  for (const p of master.paths) if (p.level1 && !seen.has(keyOf(p.level1))) seen.set(keyOf(p.level1), p.level1);
  return [...seen.values()];
}

/** Level 2 unik (beserta kode dan induk Level 1-nya). */
export function level2Entries(master) {
  const seen = new Map();
  for (const p of master.paths) {
    const id = `${keyOf(p.level1)}|${keyOf(p.level2)}`;
    if (p.level2 && !seen.has(id)) seen.set(id, { level1: p.level1, kode2: p.kode2, level2: p.level2 });
  }
  return [...seen.values()];
}

export function findLevel2(master, level1, level2) {
  return level2Entries(master).find(e => keyOf(e.level1) === keyOf(level1) && keyOf(e.level2) === keyOf(level2));
}

export function findPath(master, level1, level2, level3) {
  return (master.byLevel3.get(keyOf(level3)) || []).find(p => keyOf(p.level1) === keyOf(level1) && keyOf(p.level2) === keyOf(level2));
}

/** Daftar perubahan antara nilai asli dan draf. */
export function diffFields(orig, draft) {
  const out = [];
  for (const field of EDIT_ORDER) {
    const before = tidy(orig[field]), after = tidy(draft[field]);
    if (before !== after) out.push({ field, old: before, new: after });
  }
  return out;
}

/** Pesan masalah bila kombinasi Level 1/2/3 pada draf tidak ada di master. */
export function pathProblem(draft, master) {
  if (!tidy(draft.level1)) return 'Pilih Level 1.';
  if (!tidy(draft.level2)) return 'Pilih Level 2.';
  if (!tidy(draft.level3)) return 'Pilih Level 3 yang sesuai dengan Level 2.';
  if (!findPath(master, draft.level1, draft.level2, draft.level3)) return 'Kombinasi Level 1, Level 2, dan Level 3 tidak ada di Master_Level.';
  return '';
}

const pathKey = (l1, l2, l3) => `${keyOf(l1)}|${keyOf(l2)}|${keyOf(l3)}`;

/** Urutan tiap klasifikasi di Master_Level (dipakai untuk menentukan kelompok tujuan). */
export function masterOrder(master) {
  const order = new Map();
  master.paths.forEach((p, i) => { const k = pathKey(p.level1, p.level2, p.level3); if (!order.has(k)) order.set(k, i); });
  return order;
}

/**
 * Menentukan posisi baru sebuah baris setelah dikoreksi.
 * `rows` = semua baris data KECUALI baris yang dikoreksi: { row, level1, level2, level3, nama }.
 * Hasil: indeks baris (absolut) tempat baris disisipkan SEBELUM-nya, atau null bila tetap di tempat.
 */
export function computeMoveTarget(rows, self, orig, updated, order) {
  const sameGroup = keyOf(orig.level1) === keyOf(updated.level1) && keyOf(orig.level2) === keyOf(updated.level2) && keyOf(orig.level3) === keyOf(updated.level3);
  if (sameGroup && keyOf(orig.nama) === keyOf(updated.nama)) return null;
  const me = pathKey(updated.level1, updated.level2, updated.level3);
  const group = rows.filter(r => pathKey(r.level1, r.level2, r.level3) === me);
  let insertBefore = null;
  if (group.length) {
    const byName = group.filter(r => keyOf(r.nama) === keyOf(updated.nama));
    const anchor = byName.length ? byName : group;
    insertBefore = Math.max(...anchor.map(r => r.row)) + 1;
  } else {
    const mine = order.get(me);
    if (mine === undefined) return null;
    const known = rows.map(r => ({ r, key: order.get(pathKey(r.level1, r.level2, r.level3)) })).filter(k => k.key !== undefined);
    if (!known.length) return null;
    const before = known.filter(k => k.key <= mine);
    if (before.length) {
      const best = Math.max(...before.map(k => k.key));
      insertBefore = Math.max(...before.filter(k => k.key === best).map(k => k.r.row)) + 1;
    } else {
      const first = Math.min(...known.map(k => k.key));
      insertBefore = Math.min(...known.filter(k => k.key === first).map(k => k.r.row));
    }
  }
  if (insertBefore === self.row || insertBefore === self.row + 1) return null;
  return insertBefore;
}

/** Nomor baris akhir setelah baris dipindah dari `from` ke posisi sebelum `insertBefore`. */
export const finalRowAfterMove = (from, insertBefore) => (from < insertBefore ? insertBefore - 1 : insertBefore);

/** Pilihan isian bebas (Nama / Spesifikasi) dari nilai yang sudah ada di sheet. */
export function valueOptions(rows, field, draft, text, browse, limit = 60) {
  const l3 = keyOf(draft.level3), nm = keyOf(draft.nama);
  const counts = new Map();
  for (const r of rows) {
    const value = r[field];
    if (!value) continue;
    const k = keyOf(value);
    let e = counts.get(k);
    if (!e) { e = { label: value, n: 0, rank: 2 }; counts.set(k, e); }
    e.n++;
    const inL3 = keyOf(r.level3) === l3;
    const rank = field === 'nama' ? (inL3 ? 0 : 2) : (inL3 && keyOf(r.nama) === nm ? 0 : inL3 ? 1 : 2);
    if (rank < e.rank) e.rank = rank;
  }
  let list = [...counts.values()];
  const q = keyOf(text);
  if (browse || !q) {
    const best = list.filter(e => e.rank === 0);
    list = best.length ? best : list.filter(e => e.rank < 2).length ? list.filter(e => e.rank < 2) : list;
  } else list = list.filter(e => keyOf(e.label).includes(q));
  list.sort((a, b) => a.rank - b.rank || b.n - a.n || a.label.localeCompare(b.label, 'id'));
  return list.slice(0, limit).map(e => ({ label: e.label, sub: `${e.n} baris${e.rank === 0 ? ' · kelompok sama' : ''}`, data: e.label }));
}

export function formatTime(date = new Date()) {
  const p = n => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${p(date.getMonth() + 1)}-${p(date.getDate())} ${p(date.getHours())}:${p(date.getMinutes())}:${p(date.getSeconds())}`;
}
