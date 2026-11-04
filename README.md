# Koreksi Klasifikasi Katalog (Excel Add-in) — v1.0

Add-in Excel (Office.js, task pane) untuk meninjau dan mengoreksi klasifikasi Level 1–3 pada katalog sumber daya (`Katalog_DG_V2_Rapih`, `Katalog_DSU_V2`, dan sheet sejenis). Berjalan di Excel Online maupun Excel Desktop dari satu hosting.

## Cara kerja

1. Add-in mendeteksi otomatis baris judul (Level 1, Level 2, Level 3, Nama Produk / Pekerjaan, dst.) pada sheet aktif, jadi tidak terkunci pada satu sheet.
2. Tombol **Siapkan kolom koreksi** menambahkan kolom di kanan tabel (data asli tidak diubah):

| Kolom | Isi |
| --- | --- |
| Koreksi: Status | Valid / Dikoreksi / Perlu Review (kosong = belum ditinjau) |
| Koreksi: Seharusnya di Level | Level tempat nama produk/pekerjaan seharusnya berada |
| Koreksi: Level 1 / Kode Level 2 / Level 2 / Level 3 | Klasifikasi yang benar (diisi otomatis dari Master_Level) |
| Koreksi: Catatan | Catatan bebas |

3. Klik sel pada baris mana pun: panel menampilkan data baris, klasifikasi saat ini, hasil validasi terhadap `Master_Level`, dan saran Level 3 dari nama produk.
4. Cari Level 3 yang benar → Level 1/Kode/Level 2 terisi otomatis dari master → **Simpan koreksi**. Opsi lain: **Tandai valid**, **Perlu review**.
5. Navigasi: baris sebelumnya/berikutnya, atau lompat ke baris berikutnya yang statusnya masih kosong.

Master dibaca dari sheet `Master_Level` di workbook aktif (header: Level 1, Kode Level 2, Level 2, Level 3). Tombol **Muat ulang master** memuat perubahan terbaru.

## Hosting (Cloudflare Pages)

| Pengaturan | Nilai |
| --- | --- |
| Production branch | `main` |
| Framework preset | `None` |
| Build command | kosong |
| Build output directory | `public` |
| Nama project | `katalog-revisi-addin` |

Jika nama project/URL berbeda, ganti `https://katalog-revisi-addin.pages.dev/` di `manifest.xml`.

## Memasang di Excel

- **Excel Online**: Insert → Add-ins → Manage My Add-ins → Upload My Add-in → pilih `manifest.xml`.
- **Excel Desktop**: Insert → My Add-ins → Upload My Add-in (atau Shared Folder Catalog). Keduanya memakai manifest yang sama.

## Pengembangan

```
npm test
```

Jangan commit workbook/data katalog (sudah ada di `.gitignore`).
