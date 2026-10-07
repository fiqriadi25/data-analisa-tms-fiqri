# Data Analisa TMS Fiqri

Kumpulan script Python dan Google Colab untuk membantu pengolahan serta analisis data Transport Management System (TMS).

## Daftar fungsi

| No. | Fungsi | File | Status |
|---:|---|---|---|
| 1 | Master Baru & Generate Kode Zona Baru | [`scripts/master_baru_generate_kode_zona.py`](scripts/master_baru_generate_kode_zona.py) | Tersedia |
| 2 | Penambahan Toko & Generate Kode Zona | [`scripts/penambahan_toko_generate_kode_zona.py`](scripts/penambahan_toko_generate_kode_zona.py) | Tersedia |

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

## Menambah script baru

Simpan setiap fungsi baru sebagai file Python terpisah di folder `scripts/`. Gunakan format nama file huruf kecil dan underscore, misalnya:

```text
scripts/
├── master_baru_generate_kode_zona.py
├── penambahan_toko_generate_kode_zona.py
├── validasi_master_toko.py
└── analisa_ujp.py
```

Setelah menambahkan script, perbarui tabel **Daftar fungsi** pada README ini.

## Struktur repository

```text
data-analisa-tms-fiqri/
├── scripts/
│   ├── master_baru_generate_kode_zona.py
│   └── penambahan_toko_generate_kode_zona.py
├── tests/
│   ├── test_master_baru_generate_kode_zona.py
│   └── test_penambahan_toko_generate_kode_zona.py
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
