# Data Analisa TMS Fiqri

Kumpulan script Python dan Google Colab untuk membantu pengolahan serta analisis data Transport Management System (TMS).

## Daftar fungsi

| No. | Fungsi | File | Status |
|---:|---|---|---|
| 1 | Master Baru & Generate Kode Zona Baru | [`scripts/master_baru_generate_kode_zona.py`](scripts/master_baru_generate_kode_zona.py) | Tersedia |
| 2 | Penambahan Toko & Generate Kode Zona | [`scripts/penambahan_toko_generate_kode_zona.py`](scripts/penambahan_toko_generate_kode_zona.py) | Tersedia |
| 3 | Update UJP Master | [`scripts/update_ujp_master.py`](scripts/update_ujp_master.py) | Tersedia |
| 4 | Penambahan Master Toko | [`scripts/penambahan_master_toko.py`](scripts/penambahan_master_toko.py) | Tersedia |

## 1. Master Baru & Generate Kode Zona Baru

Script ini membaca sheet **DATA MASTER** dari file Excel dan membuat kolom **P2** sebagai kode zona. Kode dibentuk berdasarkan kombinasi:

- OP;
- kecamatan; dan
- pola nilai pada seluruh kolom UJP yang tersedia.

Kombinasi OP, kecamatan, dan UJP yang sama akan mendapatkan kode P2 yang sama. Pola UJP yang berbeda akan mendapatkan suffix berbeda. Panjang kode P2 dibatasi maksimal 15 karakter.

### Kolom wajib

- `OP`
- `KECAMATAN`
- minimal satu kolom UJP yang didukung

Daftar kolom UJP yang didukung dapat dilihat pada konstanta `UJP_COLUMNS` di dalam script.

### Cara menjalankan di Google Colab

1. Buka [Google Colab](https://colab.research.google.com/).
2. Buat notebook baru.
3. Jalankan perintah berikut untuk mengunduh repository:

   ```python
   !git clone https://github.com/fiqriadi25/data-analisa-tms-fiqri.git
   %cd data-analisa-tms-fiqri
   !pip install -r requirements.txt
   ```

4. Jalankan script:

   ```python
   %run scripts/master_baru_generate_kode_zona.py
   ```

5. Upload file Excel ketika diminta. Hasil akan dibuat dengan nama:

   ```text
   data_master_generate_zona.xlsx
   ```

## 2. Penambahan Toko & Generate Kode Zona

Script ini digunakan untuk menentukan kode zona toko baru berdasarkan referensi pada sheet **MASTER TOKO**. Input wajib memiliki sheet **DATA MASTER** dan **MASTER TOKO**.

Status hasil yang tersedia:

- `SUDAH ADA`: kode zona sebelumnya dipertahankan;
- `PAKAI ZONA MASTER (UJP SAMA)`: menggunakan zona master karena profil UJP sama;
- `SUFFIX BARU (UJP BERBEDA)`: kecamatan sudah ada, tetapi satu atau lebih nilai UJP berbeda;
- `KODE ZONA BARU`: kecamatan belum mempunyai referensi zona;
- `KODE BARU (BATAS 15 KARAKTER)`: prefix lama tidak cukup untuk suffix baru;
- `PERLU REVIEW`: data tidak memenuhi validasi dan memerlukan pemeriksaan manual.

Jalankan di Google Colab:

```python
%run scripts/penambahan_toko_generate_kode_zona.py
```

Jalankan secara lokal:

```bash
python scripts/penambahan_toko_generate_kode_zona.py data_master_olah.xlsx -o hasil_zona.xlsx
```

File output berisi tiga sheet: `DATA MASTER`, `MASTER TOKO`, dan `AUDIT REFERENSI`.

## 3. Update UJP Master

Script ini menganalisis kolom update UJP untuk menentukan apakah toko tetap menggunakan kode zona lama, memerlukan update UJP, dipindahkan ke zona lain, atau mendapatkan suffix zona baru.

Nilai pada kolom UJP asli **tidak diubah**. Hasil hanya:

- menyesuaikan `Kode Zona` jika diperlukan; dan
- menambahkan kolom `Keterangan`.

Klasifikasi zona:

- **Tipe A:** tidak ada perubahan UJP;
- **Tipe B1:** semua toko dalam zona berubah ke signature UJP yang sama;
- **Tipe B2:** perubahan campuran menghasilkan signature akhir yang sama sehingga update dibatalkan;
- **Tipe C:** terdapat beberapa signature akhir sehingga kelompok minoritas memperoleh suffix zona baru.

Jalankan di Google Colab:

```python
%run scripts/update_ujp_master.py
```

Jalankan secara lokal:

```bash
python scripts/update_ujp_master.py data_master.xlsx -o hasil_update_ujp.xlsx
```

## 4. Penambahan Master Toko

Script ini melengkapi sheet **DATA MASTER** dengan `KODE ZONA` dan `NAMA ZONA`
berdasarkan kombinasi OP, kecamatan, dan pola seluruh kolom UJP yang terdeteksi.
Kode dibuat secara deterministik, unik per OP/kecamatan/pola UJP, dan dibatasi
maksimal 15 karakter.

Hasil tetap mempertahankan sheet dan format workbook sumber, lalu menambahkan:

- `MAPPING_ZONA`, berisi hubungan kecamatan, pola UJP, suffix, dan kode zona;
- `AUDIT_ZONA`, berisi hasil validasi panjang, format, keunikan, dan konsistensi kode.

Jalankan di Google Colab:

```python
%run scripts/penambahan_master_toko.py
```

Jalankan secara lokal:

```bash
python scripts/penambahan_master_toko.py data_master.xlsx -o hasil_master_toko.xlsx
```

Kolom wajib adalah `OP`, `KECAMATAN`, dan minimal satu kolom dengan awalan
`UJP `. Kolom UJP tambahan di luar daftar standar akan terdeteksi otomatis.

## Menambah script baru

Simpan setiap fungsi baru sebagai file Python terpisah di folder `scripts/`. Gunakan format nama file huruf kecil dan underscore, misalnya:

```text
scripts/
├── master_baru_generate_kode_zona.py
├── penambahan_toko_generate_kode_zona.py
├── update_ujp_master.py
├── penambahan_master_toko.py
├── validasi_master_toko.py
└── analisa_ujp.py
```

Setelah menambahkan script, perbarui tabel **Daftar fungsi** pada README ini.

## Struktur repository

```text
data-analisa-tms-fiqri/
├── scripts/
│   ├── master_baru_generate_kode_zona.py
│   ├── penambahan_toko_generate_kode_zona.py
│   ├── update_ujp_master.py
│   └── penambahan_master_toko.py
├── tests/
│   ├── test_master_baru_generate_kode_zona.py
│   ├── test_penambahan_toko_generate_kode_zona.py
│   ├── test_update_ujp_master.py
│   └── test_penambahan_master_toko.py
├── .gitignore
├── LICENSE
├── README.md
└── requirements.txt
```

## Catatan keamanan data

File sumber dan file hasil Excel tidak disimpan ke GitHub karena sudah dikecualikan melalui `.gitignore`. Hindari melakukan commit data operasional, kredensial, atau informasi sensitif perusahaan.

## Menjalankan pengujian

```bash
pip install -r requirements.txt
pytest
```

## Lisensi

Project ini menggunakan lisensi MIT. Lihat file [`LICENSE`](LICENSE).
