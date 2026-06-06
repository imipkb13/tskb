# Aplikasi Olah Data Excel Impor Barang

Aplikasi desktop berbasis Python untuk mengolah data excel impor barang, menyalin data dari file invoice dan packing list ke file rekap.

## Fitur

- Memproses file packing list (_pl) dan invoice (_iv) secara otomatis
- Ekstraksi data berdasarkan kata kunci (Ocean Vessel, INVOICE NO., B/L NO., Date:, Net Weight)
- Mendukung file dengan 2 sheet (CI dan attachment)
- Menyimpan hasil ke file rekap dengan 2 sheet (Rekap BL dan Detail BL)
- Menghapus tanda ":" dan "." sesuai ketentuan
- Progress tracking dengan logging detail

## Kebutuhan Sistem

- Python 3.8 atau lebih tinggi
- pandas >= 1.5.0
- openpyxl >= 3.0.0

## Instalasi

1. Install dependencies:
```bash
pip install -r requirements.txt
```

## Cara Menggunakan

### Versi GUI (Desktop Application)

Jalankan aplikasi dengan interface grafis:
```bash
python excel_processor_app.py
```

Langkah-langkah:
1. Klik "Browse" untuk memilih folder sumber yang berisi file _pl dan _iv
2. Klik "Browse" untuk menentukan lokasi file rekap output
3. Klik "Proses Data" untuk memulai pemrosesan
4. Lihat progress di log window

### Versi Command Line

Jalankan aplikasi melalui command line:
```bash
python excel_processor_cli.py
```

Ikuti petunjuk untuk memasukkan:
- Path folder sumber
- Path file rekap output

## Format File

### Packing List (_pl)
- 1 sheet
- Row ke-2: Nama supplier (akan disalin ke kolom F)
- Kata kunci yang dicari:
  - "Ocean Vessel" → nama kapal (kolom E)
  - "INVOICE NO." → nomor invoice (kolom I)
  - "B/L NO." → nomor BL (kolom H)
  - "Date:" → tanggal invoice (kolom J)
  - "Net Weight" → berat total (kolom M)

### Invoice (_iv)
- 2 sheet: "CI" dan "attachment"

#### Sheet CI
- "B/L NO." → nomor BL (kolom I sheet 2)
- "INVOICE NO." → nomor invoice (kolom J sheet 2)
- "Date:" → tanggal invoice (kolom K sheet 2)

#### Sheet Attachment
- Dimulai dari row dengan angka 1 di kolom A
- Kolom A (nomor urut) → kolom M sheet 2
- Kolom B (nomor kontrak) → kolom Y sheet 2
- Kolom C, D, E → kolom O, P, Q sheet 2
- Kolom F, G, H, I → kolom S, T, U, V sheet 2
- Kolom L → kolom X sheet 2
- Kolom N → kolom AA sheet 2 (semua tanda "." dihapus)

## Struktur File Rekap

### Sheet 1: Rekap BL
Kolom yang diisi: A-J, L-N
- C: Nomor kapal (4 karakter pertama dari nama file)
- E: Nama kapal (dari keyword "Ocean Vessel")
- F: Supplier (dari row 2 packing list)
- H: B/L NO.
- I: INVOICE NO.
- J: Date
- M: Net Weight

### Sheet 2: Detail BL
Berisi detail barang dari sheet attachment dengan mapping kolom sesuai spesifikasi.

## Contoh Nama File

- `ABCD_001_pl.xlsx` - Packing list untuk kapal ABCD
- `ABCD_001_iv.xlsx` - Invoice untuk kapal ABCD

4 karakter pertama sebelum "_" adalah nomor kapal.

## Catatan Penting

- Semua file _pl dan _iv untuk satu kapal harus berada dalam folder yang sama
- Tanda ":" akan dihapus otomatis saat menyalin data
- Untuk kolom N di attachment, semua tanda "." akan dihapus sebelum disalin ke kolom AA
- Proses akan berhenti jika tidak ada lagi nomor urut berurutan di kolom A sheet attachment

## Troubleshooting

Jika terjadi error:
1. Pastikan format file Excel sesuai spesifikasi
2. Periksa apakah keyword yang dicari ada dalam file
3. Pastikan sheet "CI" dan "attachment" ada di file invoice
4. Lihat log untuk detail error

## License

Free to use for personal and commercial projects.