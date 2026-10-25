# Tokopedia Transaction PDF Parser

Parse export PDF **Riwayat Transaksi Tokopedia** (Laporan Data Pengguna) menjadi spreadsheet Excel yang sudah dikategorikan.

Satu barang = satu baris. Order multi-item dipecah. Status **Tidak Berhasil** ikut dicantumkan.

## Output columns

| No | Tanggal | Kategori | Nama Barang | Nilai (Rp) | No. Invoice | Penjual | Status |
|----|---------|----------|-------------|------------|-------------|---------|--------|

Sheet tambahan: `Ringkasan`, `Catatan`.

## Setup

```bash
cd tokopedia-trx-parser
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

1. Unduh riwayat transaksi dari Tokopedia (PDF Laporan Data Pengguna).
2. Jalankan:

```bash
python parse_tokopedia_trx.py /path/to/TrxTokopedia.pdf
# atau
python parse_tokopedia_trx.py /path/to/TrxTokopedia.pdf -o output/hasil.xlsx
```

## Notes

- **Jangan commit PDF/XLSX** yang berisi data pribadi. Sudah di-ignore di `.gitignore`.
- Satu barang dengan `@Rp`: `Nilai` = `Jumlah Total` bila lebih kecil (diskon). Bila lebih besar, tetap `@Rp`.
- Satu barang tanpa `@Rp`: `Nilai` = `Jumlah Total`.
- Multi-item: tiap barang `@Rp × qty`. Bila `Jumlah Total` lebih kecil, baris berikutnya adalah `Potongan harga (tidak teralokasi)` (kategori Penyesuaian). Selisih positif (ongkir / biaya layanan) tidak dicantumkan.
- No. Invoice = baris invoice tepat **sebelum** `Status` order.
- Kategori otomatis berbasis kata kunci di `CAT_RULES` — silakan edit / koreksi manual di Excel.

## Open in Cursor IDE

```bash
cursor .
# atau di Cursor IDE: File → Open Folder → pilih folder tokopedia-trx-parser
```

## Push ke GitHub

```bash
cd tokopedia-trx-parser
gh repo create tokopedia-trx-parser --public --source=. --remote=origin --push
# atau buat repo kosong di GitHub, lalu:
# git remote add origin git@github.com:<user>/tokopedia-trx-parser.git
# git push -u origin main
```

## License

MIT
