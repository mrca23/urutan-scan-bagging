# Urutan Scan Bagging

Web untuk mencari No. Waybill dari export JMS **Pencarian Scan** (contoh: scan Unpack): masuk **bagging ke berapa** dan **waybill ke berapa**, berdasarkan waktu scan.

**Pakai online:** https://mrca23.github.io/urutan-scan-bagging/

## Cara tarik data
JMS > Pencarian Scan > pilih jenis scan (mis. Unpack), station & tanggal > export (.xlsx)

## Cara pakai
Ada video tutorial di bagian bawah halaman web.

1. Buka link di atas, klik **Pilih File** (atau tarik file ke kotak upload).
2. Ketik No. Waybill (bisa banyak sekaligus, pisahkan spasi / baris baru), klik **Cari**.
3. Hasil: `Masuk bagging ke-X dan waybill ke-Y`, plus No. Bagging, urutan total, waktu scan. Klik **Lihat di daftar** untuk loncat ke barisnya.

Daftar urutan bisa dilihat **Per Bagging** (lipat/buka per bagging) atau **Semua AWB**. Kalau file berisi lebih dari satu jenis scan, tiap jenis scan punya tab sendiri dan urutannya dihitung terpisah.

## Definisi
| Istilah | Arti |
|---|---|
| Bagging ke-X | Urutan bagging berdasarkan waktu scan **pertama** di bagging itu |
| Waybill ke-Y | Urutan AWB di dalam baggingnya berdasarkan waktu scan |
| Urutan total | Urutan AWB dari semua scan (jenis scan yang sama) berdasarkan waktu scan |

Waktu scan sama persis -> urutan mengikuti urutan baris di file.

## Deteksi kolom (tidak bergantung nama/urutan kolom)
Kolom dikenali dari **isi data**, nama kolom hanya bonus kecil:
- **No. Waybill**: pola kode huruf+angka (mis. `JY1762155487`), hampir semua unik.
- **No. Bagging**: pola kode huruf+angka (mis. `LY000019928878`), berulang, jumlah nilai beda paling banyak.
- **Waktu Scan**: kolom tanggal+jam paling unik dan paling awal (Waktu Upload selalu sama/lebih lambat).
- **Jenis Scan**: teks dengan sedikit variasi yang berisi kata jenis scan (Unpack, Bagging, Incoming, ...).
- Discan oleh / Station Scan: opsional, dari nama kolom (hanya info tambahan).

Hasil deteksi bisa dicek dan diubah manual di bagian "Kolom terdeteksi".

File diolah langsung di browser (pakai [SheetJS](https://sheetjs.com)). Tidak ada data yang dikirim ke server.

## Catatan pengembangan
- Repo publik: jangan commit data (`.xlsx/.csv`, folder `test/`).
- Video tutorial: `tutorial.mp4` (kompres dari rekaman asli: `ffmpeg -i <asli>.mp4 -vf scale=1280:-2 -crf 27 -preset slow -b:a 96k -movflags +faststart tutorial.mp4`), poster `tutorial.jpg`.
- Setiap `app.js` diubah, naikkan `?v=` di `index.html`.
