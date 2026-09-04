# DOCX Cross-OS Structure Check

Alat diagnostik buat jawab pertanyaan: apakah file .docx yang sama, dibaca
di Windows vs macOS, punya struktur (margin, font, alignment, mirror margin)
yang identik? Hasilnya dipakai buat memutuskan apakah checker sebaiknya
menerima input .doc/.docx alih-alih PDF.

## Struktur folder

Taruh folder ini **sejajar dengan root repo**, dan pastikan `.github/`
ada persis di root repo (bukan di dalam `docx-cross-os-check/`), karena
GitHub Actions cuma baca workflow dari `.github/workflows/` di root.

```
<root-repo>/
├── .github/
│   └── workflows/
│       └── docx-cross-os.yml
├── docx-cross-os-check/
│   ├── extract_structure.py
│   ├── compare_structure.py
│   ├── test-assets/
│   │   └── sample.docx   ← kamu tambahkan manual
│   └── README.md
├── codes/
│   └── aggregator/
│       └── ...
```

Jangan taruh ini di dalam `codes/aggregator/pipelines/visual/` — itu tempat
modul checker production, sedangkan ini tool sekali-pakai buat eksperimen.

## Cara pakai

### 1. Siapkan file uji
Copy dokumen .docx yang sudah kamu pakai buat test PDF checker (yang ada
kesalahan format sengaja) ke `docx-cross-os-check/test-assets/sample.docx`.

### 2. Push ke GitHub
```bash
git add .github/workflows/docx-cross-os.yml docx-cross-os-check/
git commit -m "Add cross-OS docx structure check"
git push
```

### 3. Jalankan workflow
- Buka repo di GitHub → tab **Actions**.
- Kalau ini run pertama, GitHub mungkin minta kamu klik "I understand my
  workflows, go ahead and enable them" — klik itu dulu.
- Pilih workflow **DOCX Cross-OS Structure Check** di sidebar kiri.
- Klik **Run workflow** (dropdown di kanan) → pilih branch → **Run workflow**.
- Tunggu sampai dua job (`windows-latest` dan `macos-latest`) selesai
  (biasanya 1-2 menit).

### 4. Ambil hasil
- Buka run yang baru selesai → scroll ke bagian **Artifacts** di bawah.
- Download `structure-windows-latest` dan `structure-macos-latest`.
- Extract keduanya — masing-masing berisi satu file
  `structure_Windows.json` / `structure_Darwin.json`.

### 5. Bandingkan
```bash
python compare_structure.py structure_Windows.json structure_Darwin.json
```

Output-nya langsung bilang bagian mana yang beda (margin, font, alignment,
line spacing, dst), atau bilang "struktur identik" kalau enggak ada
perbedaan.

### 6. Laporkan ke Pak Ali
Bawa output `compare_structure.py` apa adanya — itu bukti konkret, bukan
cuma "kayaknya sama", buat mendukung rekomendasi format input mana yang
dipilih (doc/docx, RTF, atau tetap PDF).

## Catatan
- Kalau mau uji dokumen lain, ganti `test-assets/sample.docx` lalu push lagi
  (workflow otomatis jalan karena ada `push.paths` trigger), atau trigger
  manual lewat `workflow_dispatch`.
- Script ini fokus ke margin/font/alignment/mirror-margin karena itu yang
  paling relevan sama masalah `pdf_extractor.py` kamu sekarang. Kalau nanti
  butuh cek hal lain (misalnya numbering, section breaks), tinggal tambah
  field di `extract_structure.py`.
