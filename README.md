# Koreksi Klasifikasi Katalog (Excel Add-in) — v2

Add-in Excel (Office.js, task pane) untuk meninjau dan mengoreksi klasifikasi katalog sumber daya (`Katalog_DG_V2_Rapih`, `Katalog_DSU_V2`, dan sheet sejenis). Berjalan di Excel Online maupun Excel Desktop.

## Cara kerja (v2: koreksi langsung)

1. Add-in mendeteksi otomatis baris judul pada sheet aktif.
2. Klik sel pada baris mana pun: panel menampilkan data baris, klasifikasi saat ini, dan hasil pengecekan terhadap `Master_Level`.
3. Bagian **Koreksi**: Level 1 (dropdown), Level 2, Level 3, Nama Produk / Pekerjaan, dan Spesifikasi (ketik + saran, atau klik panah untuk daftar). Kode Level 2 mengikuti Level 2 otomatis.
4. **Simpan koreksi** menulis perubahan langsung ke baris tersebut; sel yang berubah diwarnai kuning.
5. Baris otomatis pindah ke kelompok yang benar (opsi dapat dimatikan).
6. Satu-satunya kolom baru: **Catatan Koreksi** (kanan tabel). Isi catatan tersimpan otomatis.
7. Semua perubahan (nilai lama/baru, pindah baris, catatan, tandai valid) dicatat di sheet **Log**.

## Hosting (Cloudflare Pages)

| Pengaturan | Nilai |
| --- | --- |
| Production branch | `main` |
| Framework preset | `None` |
| Build command | kosong |
| Build output directory | `public` |

Jika URL berbeda dari `https://katalog-revisi-addin.pages.dev/`, ubah `manifest.xml`.

## Uji lokal

```
npx office-addin-dev-certs install
node server.mjs
```

Lalu pasang `manifest-localhost.xml` (Excel Online: Insert → Add-ins → Upload My Add-in; Excel Desktop: Shared Folder Catalog).

## Pengembangan

```
npm test
```

Jangan commit workbook/data katalog (sudah ada di `.gitignore`).
