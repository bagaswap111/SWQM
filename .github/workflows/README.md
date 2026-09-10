# GitHub Action - Build Proposal DOCX on Release

Workflow ini akan otomatis membuat file DOCX proposal dan menguploadnya ke GitHub Release setiap kali release baru dipublish.

## 📁 Struktur File

```
.github/
  workflows/
    build-docx-release.yml    # GitHub Actions workflow
scripts/
  build_docx.py               # Script Python untuk generate DOCX
output/
  PROPOSAL_SWQMS.docx         # File DOCX yang di-generate
```

## 🚀 Cara Kerja

1. **Trigger**: Workflow dijalankan saat release baru di-publish (`published` atau `prereleased`)
2. **Build Environment**: Setup Python 3.11 dengan dependencies `python-docx` dan `markdown`
3. **Generate DOCX**: Menjalankan script `build_docx.py` untuk membuat file proposal
4. **Upload to Release**: File DOCX diupload ke GitHub Release menggunakan `softprops/action-gh-release`
5. **Upload Artifact**: File juga disimpan sebagai artifact selama 30 hari

## 📦 Dependencies

- Python 3.11+
- python-docx
- markdown

## 🔧 Kustomisasi

### Mengubah Nama File Output
Edit baris berikut di `scripts/build_docx.py`:
```python
output_path = 'output/NAMA_FILE_BARU.docx'
```

### Menambahkan Section Baru
Tambahkan kode di fungsi `build_proposal()` dalam `scripts/build_docx.py`.

### Mengubah Trigger Event
Edit bagian `on:` di `.github/workflows/build-docx-release.yml`:
```yaml
on:
  release:
    types: [published, prereleased]  # Tambahkan 'created' jika perlu
```

### Menambahkan Platform Lain
Untuk upload ke Google Drive atau storage lain, tambahkan step baru setelah upload ke release.

## 🧪 Testing Lokal

Jalankan script secara lokal sebelum commit:

```bash
# Install dependencies
pip install python-docx markdown

# Jalankan script
python scripts/build_docx.py

# Cek output
ls -lh output/PROPOSAL_SWQMS.docx
```

## 📄 Konten Proposal

Proposal yang di-generate mencakup:
- Halaman Judul
- Daftar Isi
- Ringkasan Eksekutif
- Latar Belakang
- Tujuan Proyek
- Manfaat & Value Proposition
- Spesifikasi Teknis Sistem
- Rencana Anggaran Biaya (RAB)
- Deliverables
- Garansi & Dukungan
- Penutup & Kontak

## 🔐 Secrets yang Diperlukan

Workflow ini menggunakan `GITHUB_TOKEN` yang sudah tersedia secara default. Tidak perlu setup secrets tambahan.

## 📊 Output

Setiap release akan memiliki:
1. **Attachment**: `PROPOSAL_SWQMS.docx` terlampir di release
2. **Artifact**: File tersedia untuk download selama 30 hari

## 🛠️ Troubleshooting

### Build Gagal
- Cek log GitHub Actions untuk error detail
- Pastikan semua dependencies terinstall
- Verifikasi path file benar

### File Tidak Terupload
- Pastikan tag release valid
- Cek permission `GITHUB_TOKEN`
- Verifikasi file ada di path `output/PROPOSAL_SWQMS.docx`

## 📝 License

MIT License
