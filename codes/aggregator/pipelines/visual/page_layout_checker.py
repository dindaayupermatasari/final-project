"""
page_layout_checker.py
=======================
SUB-TAHAP: RENDER & LOCATE
---------------------------
python-docx tidak menyimpan nomor halaman -- itu hasil kalkulasi
rendering (font, margin, page break, line spacing), bukan properti
statis di XML. Supaya bisa memeriksa dua aturan berikut, dokumen harus
di-render dulu ke PDF, baru discan per halaman:

  1. Setiap Bab wajib dimulai di halaman baru bernomor GANJIL.
  2. Kalau suatu Bab berakhir di halaman GANJIL, halaman GENAP
     berikutnya dibiarkan kosong dengan tulisan "Halaman ini sengaja
     dikosongkan" di tengah halaman.

Alur modul ini:
    1. render_to_pdf()            -> convert docx ke pdf (LibreOffice headless)
    2. extract_page_texts()       -> ambil teks tiap halaman (PyMuPDF/fitz)
    3. extract_footer_numbers()   -> baca ANGKA TERCETAK di footer LANGSUNG
       dari koordinat fisiknya di tiap halaman PDF (ground truth paling
       akurat -- lihat docstring fungsinya untuk alasan lengkap), lalu
       disaring lewat _validate_footer_sequence() untuk membuang angka
       yang menyimpang dari urutan wajar (lihat CHANGELOG terbaru).
    4. find_chapter_boundaries()  -> cari halaman mulai tiap heading setingkat
       bab (BAB N atau FRONT_HEADINGS). Regex & set heading di-IMPORT dari
       extractor.py -- single source of truth definisi "judul bab", supaya
       tidak ada dua definisi yang bisa berbeda persepsi di pipeline yang
       sama.
    5. check_page_layout()        -> entry point. Return (findings, checks)
       dengan BENTUK YANG SAMA seperti comparator.compare_all, supaya
       main.py tinggal menggabungkan ke findings/checks utama.

Kategori semua finding di sini: "Format" (soal layout fisik halaman,
bukan sistematika penulisan).

CHANGELOG (revisi terbaru -- validasi urutan footer):
- Sebelumnya nomor halaman TERCETAK dihitung MURNI dari rumus
  (w:pgNumType/start tiap section + jumlah halaman fisik hasil render).
  Masalahnya: rumus itu asumsi jumlah halaman fisik hasil render SAMA
  PERSIS dengan yang dilihat mahasiswa di Word. Kalau font "Times New
  Roman" tidak benar-benar terpasang di environment yang menjalankan
  LibreOffice (server sering tidak punya font Microsoft ini), LO
  men-substitusi font lain yang metrik lebarnya beda tipis -> wrapping
  baris beda -> jumlah halaman fisik bisa geser -> SEMUA section
  "continuing" (tidak restart) di belakangnya ikut salah hitung,
  walau rumusnya sendiri benar.
- Lalu ditambah extract_footer_numbers(): baca angka yang BENAR-BENAR
  tercetak di footer tiap halaman, langsung dari PDF (ground truth,
  bukan hasil kalkulasi). check_page_layout() PRIORITASKAN angka ini;
  rumus w:pgNumType lama cuma jadi FALLBACK untuk halaman yang
  footernya gagal terbaca (mis. sampul, atau kasus ambigu).
- BUG BARU DITEMUKAN: ground truth footer itu sendiri bisa "kotor" di
  satu halaman spesifik -- field PAGE otomatis Word/LO bisa salah
  hitung ulang saat render (mis. dipicu tabel/gambar besar di dekat
  footer yang mengacaukan urutan render field), sehingga angka yang
  TERCETAK di halaman itu memang salah, bukan salah baca koordinat.
  Kasus nyata: field PAGE di suatu halaman nyasar, membuat BAB 1
  dilaporkan berakhir di halaman ganjil (7) padahal aslinya genap (8).
  FIX: ditambah _validate_footer_sequence() yang membandingkan tiap
  angka footer ke angka TERPERCAYA TERAKHIR sebelumnya (bukan cuma
  tetangga fisik langsung). Nomor tercetak yang valid hanya boleh naik
  konsisten mengikuti jarak halaman fisik, ATAU restart ke 1 (section
  baru). Angka yang menyimpang dari kedua pola itu di-drop, sehingga
  check_page_layout() fallback ke rumus w:pgNumType HANYA untuk
  halaman itu saja -- halaman lain yang footernya konsisten tetap
  pakai ground truth asli.
- BUG BARU DITEMUKAN: baseline (titik pertama yang berhasil terbaca)
  dulu dipercaya TANPA pembanding sama sekali. Kalau titik pertama itu
  sendiri kena glitch field PAGE (skenario yang sama seperti di atas,
  tapi apes kena di halaman paling awal yang berhasil terbaca), OFFSET
  KONSTAN dari baseline yang salah itu ikut terbawa ke SELURUH rantai
  footer di belakangnya tanpa pernah tertangkap (validasi lama cuma cek
  konsistensi relatif, bukan kebenaran absolut). Kasus nyata: BAB 1
  yang aslinya berakhir di halaman genap (8) malah dilaporkan berakhir
  di halaman ganjil (9). FIX: baseline sekarang wajib dikonfirmasi
  dulu oleh minimal satu titik berikutnya yang konsisten dengannya
  (lihat docstring _validate_footer_sequence() untuk detail lengkap).

Dependency TAMBAHAN di luar python-docx:
    - LibreOffice terpasang & bisa dipanggil lewat command `soffice`
    - PyMuPDF  (pip install pymupdf --break-system-packages), import
      sebagai `fitz`

CHANGELOG (revisi terbaru -- confidence flag utk hasil berbasis font):
- BUG DITEMUKAN (kasus nyata: Proposal_PA_Bintang_Punta_Dzaky): di
  environment yang TIDAK punya font "Times New Roman" asli (cuma ada
  pengganti metric-compatible seperti Liberation Serif),
  _verify_font_available() kalau found is None (tooling fc-list/registry
  sendiri tidak tersedia) ATAU SKIP_FONT_CHECK=1 tetap MELANJUTKAN
  render tanpa membawa sinyal apa pun ke hasil akhir -- finding
  ganjil/genap & notice halaman kosong tetap disajikan seolah 100%
  pasti benar, padahal fondasinya (asumsi wrapping baris identik dengan
  Word) tidak pernah benar-benar terverifikasi. Kasus nyata: dokumen
  yang di Word BAB 1-nya berakhir di halaman genap (8) malah dilaporkan
  berakhir di halaman ganjil (9) -- SATU baris terakhir tabel jadwal
  yang mepet ke batas halaman cukup tergeser gara-gara font substitusi,
  padahal urutan angka footer yang terbaca sendiri konsisten sempurna
  (jadi tidak tertangkap _validate_footer_sequence(), yang memang
  didesain untuk glitch SATU halaman, bukan pergeseran sistemik akibat
  font).
  FIX: _verify_font_available() sekarang RETURN status tri-state
  ("verified"/"skipped"/"unconfirmed") alih-alih cuma print warning
  untuk kasus "unconfirmed". render_to_pdf() meneruskan status ini ke
  caller (return jadi tuple (pdf_path, font_status) -- BREAKING CHANGE,
  lihat docstring render_to_pdf(), main.py WAJIB di-update untuk
  meneruskan font_status ini ke check_page_layout()). check_page_layout()
  sekarang menerima param font_status: kalau bukan "verified", SETIAP
  finding yang dihasilkan modul ini (semuanya bergantung pada asumsi
  wrapping halaman identik dengan Word) ditandai confidence="rendah" +
  catatan eksplisit di description, BUKAN dihapus/disembunyikan --
  supaya pembimbing tetap lihat potensi masalahnya, tapi tahu untuk
  verifikasi manual dulu sebelum menindaklanjuti, bukan menelan mentah-
  mentah sebagai temuan pasti.
"""

import os
import pathlib
import re
import subprocess
import sys
import tempfile
from collections import defaultdict

import fitz  # PyMuPDF

from aggregator.pipelines.visual.extractor import CHAPTER_TITLE_RE, FRONT_HEADINGS

BLANK_PAGE_NOTICE = "HALAMAN INI SENGAJA DIKOSONGKAN"

# Path ke soffice.exe LibreOffice Portable (versi lama/stabil, dipakai
# khusus supaya rendering PDF konsisten & tidak ikut berubah kalau
# LibreOffice sistem di-update ke versi lebih baru). Override lewat env
# var LIBREOFFICE_PATH kalau lokasinya beda di mesin lain / saat deploy
# (mis. di Linux/server nanti cuma "soffice" karena sudah ada di PATH).
LIBREOFFICE_PATH = os.environ.get(
    "LIBREOFFICE_PATH",
    r"C:\Users\Dicky Zulfikar\Downloads\LibreOfficePortablePrevious\App\libreoffice\program\soffice.exe",
)

# Folder induk untuk SEMUA file/folder sementara konversi LibreOffice
# (input.docx, profil sementara, output pdf). WAJIB path TANPA spasi di
# Windows -- default tempfile.gettempdir() ada di
# C:\Users\<nama user>\AppData\Local\Temp, dan kalau nama user
# mengandung spasi (mis. "Dicky Zulfikar") soffice bisa crash dengan
# exit code 0xE06D7363 walau file sudah dinamai "input.docx".
# Override lewat env var LO_TEMP_ROOT. Kalau tidak di-set: di Windows
# dipakai C:\temp kalau temp bawaan berspasi; selain itu temp bawaan OS.
def _resolve_lo_temp_root() -> str | None:
    env = os.environ.get("LO_TEMP_ROOT")
    if env:
        os.makedirs(env, exist_ok=True)
        return env
    default_tmp = tempfile.gettempdir()
    if os.name == "nt" and " " in default_tmp:
        safe = r"C:\temp"
        try:
            os.makedirs(safe, exist_ok=True)
            return safe
        except OSError:
            return None
    return None


LO_TEMP_ROOT = _resolve_lo_temp_root()

# wdFormatPDF = 17 (konstanta MS Word, tidak berubah antar versi Word).
_WD_FORMAT_PDF = 17

# RENDER_ENGINE mengontrol engine mana yang dipakai convert docx -> pdf:
#   "libreoffice" (DEFAULT) -- SATU-SATUNYA engine yang dipakai baik
#                  pas kamu tes lokal MAUPUN pas sudah di-deploy ke
#                  server nanti. SENGAJA dibuat default & satu-satunya
#                  jalur normal -- supaya tes lokal BENERAN mewakili apa
#                  yang bakal jalan di production, bukan cuma "kelihatan
#                  akurat pas ditest" gara-gara diam-diam pakai engine
#                  lain. Akurasinya dijamin lewat _verify_font_available()
#                  (lihat di bawah), BUKAN lewat ganti-ganti engine.
#   "word"        -- Word COM, HANYA untuk tes/debug MANUAL sekali-sekali
#                  di laptop Windows-mu (mis. mau bandingkan "apa kata
#                  Word sendiri soal halaman ini" saat curiga ada bug).
#                  JANGAN dipakai sebagai default harian -- kalau dipakai
#                  terus buat tes rutin, kamu balik lagi ke masalah yang
#                  sama: tes lokal jadi tidak mewakili production.
RENDER_ENGINE = os.environ.get("RENDER_ENGINE", "libreoffice")

# Kalau True, _verify_font_available() cuma WARNING bukan RuntimeError.
# HANYA untuk kondisi darurat/debugging (mis. sedang investigasi bug lain
# yang tidak berhubungan dengan font, dan mau skip cek ini sementara).
# JANGAN pernah di-set true di server produksi.
_SKIP_FONT_CHECK = os.environ.get("SKIP_FONT_CHECK") == "1"


def _font_available_fc_list(font_name: str) -> bool | None:
    """
    Cek font terpasang lewat `fc-list` (fontconfig) -- tersedia by
    default di HAMPIR SEMUA server Linux (termasuk yang otomatis
    diinstal bareng LibreOffice), jadi ini jalur deteksi utama untuk
    lingkungan deployment. Return None (bukan False) kalau `fc-list`
    sendiri tidak ada -- artinya TIDAK BISA DIPASTIKAN, beda dari
    "dicek dan memang tidak ketemu".
    """
    try:
        result = subprocess.run(["fc-list"], capture_output=True, text=True, timeout=10)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return None
    return font_name.lower() in result.stdout.lower()


def _font_available_windows_registry(font_name: str) -> bool | None:
    """
    Cek font terpasang lewat Windows Registry (daftar font resmi OS,
    bukan sekadar file .ttf nyasar di folder sembarang). Return None
    kalau winreg tidak bisa diakses -- TIDAK BISA DIPASTIKAN.
    """
    try:
        import winreg
    except ImportError:
        return None
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
        )
    except OSError:
        return None
    try:
        i = 0
        while True:
            try:
                name, _, _ = winreg.EnumValue(key, i)
            except OSError:
                break
            if font_name.lower() in name.lower():
                return True
            i += 1
        return False
    finally:
        winreg.CloseKey(key)


def _verify_font_available(font_name: str = "Times New Roman") -> str:
    """
    Cek APAKAH font ini beneran terpasang di mesin ini SEBELUM mulai
    render -- kalau PASTI TIDAK ADA (found is False), GAGAL KERAS dengan
    pesan jelas, BUKAN diam-diam lanjut render pakai font substitusi
    LibreOffice yang lebar hurufnya beda tipis (root cause SELURUH kelas
    bug "wrapping baris geser -> jumlah halaman salah -> ganjil/genap
    salah" yang sudah didokumentasikan panjang lebar di CHANGELOG modul
    ini).

    INI DIPANGGIL DI SETIAP render_to_pdf(), baik pas dites lokal
    MAUPUN pas sudah di server produksi -- justru supaya KALAU ada
    satu lingkungan yang lupa dipasangi font-nya, langsung ketauan
    lewat error yang jelas SAAT ITU JUGA, bukan diam-diam menghasilkan
    angka halaman yang salah dan baru ketauan setelah mahasiswa
    komplain hasil check-nya aneh. Ini pengganti dari ide "auto-switch
    engine" sebelumnya -- gagal cepat & jelas lebih baik daripada
    silently correct di satu environment tapi silently wrong di
    environment lain.

    RETURN (bukan None lagi -- lihat CHANGELOG terbaru modul ini):
      "verified"    -- font dikonfirmasi ADA. Nomor halaman hasil render
                       bisa dipercaya penuh.
      "skipped"     -- SKIP_FONT_CHECK=1, verifikasi sengaja dilewati
                       (mode darurat/debug). Font BISA SAJA tidak ada;
                       caller wajib anggap nomor halaman TIDAK terjamin.
      "unconfirmed" -- tooling deteksi font (fc-list/registry) sendiri
                       tidak tersedia di environment ini -- TIDAK BISA
                       DIPASTIKAN font-nya ada atau tidak. Sebelumnya
                       kasus ini cuma di-print sebagai warning ke stderr
                       (gampang tidak kebaca / tidak sampai ke pengguna
                       akhir) lalu tetap lanjut render seolah tidak ada
                       masalah -- finding yang dihasilkan disajikan
                       seolah 100% pasti padahal fondasinya tidak pernah
                       benar-benar diverifikasi. Sekarang status ini
                       WAJIB diteruskan sampai ke check_page_layout()
                       supaya finding berbasis nomor halaman ikut
                       ditandai confidence="rendah", bukan cuma warning
                       yang gampang terlewat di log server.

    Caller (render_to_pdf) meneruskan return value ini apa adanya ke
    check_page_layout() lewat parameter font_status -- lihat docstring
    keduanya untuk alur lengkapnya.
    """
    if _SKIP_FONT_CHECK:
        return "skipped"

    if sys.platform == "win32":
        found = _font_available_windows_registry(font_name)
    else:
        found = _font_available_fc_list(font_name)

    if found is False:
        raise RuntimeError(
            f"Font '{font_name}' TIDAK ditemukan terpasang di sistem ini. "
            "Kalau dipaksa lanjut, LibreOffice akan diam-diam mengganti "
            "font ini dengan font lain yang lebar hurufnya beda -> "
            "wrapping baris & jumlah halaman ikut geser -> hasil cek "
            "ganjil/genap jadi TIDAK BISA DIPERCAYA. Pasang dulu font "
            "aslinya:\n"
            "  Linux/server : sudo apt install ttf-mscorefonts-installer\n"
            "                 (paket ini berisi font MS ASLI, bukan cuma\n"
            "                 pengganti metric-compatible -- supaya hasil\n"
            "                 render di server IDENTIK dengan di Windows,\n"
            "                 bukan cuma 'mirip')\n"
            "  Windows      : pastikan Microsoft Office/Times New Roman\n"
            "                 memang terpasang (biasanya sudah default)."
        )
    if found is None:
        print(
            f"[peringatan] Tidak bisa memastikan font '{font_name}' "
            "terpasang di sistem ini (tooling deteksi tidak tersedia di "
            "environment ini). Finding berbasis nomor halaman akan "
            "ditandai confidence rendah -- cek manual dulu sebelum "
            "menindaklanjuti.",
            file=sys.stderr,
        )
        return "unconfirmed"
    return "verified"


def _render_via_libreoffice(docx_path: str, output_dir: str) -> str:
    """
    Convert docx -> pdf pakai LibreOffice headless. PORTABLE -- jalan di
    Windows maupun Linux (asal LibreOffice terpasang & LIBREOFFICE_PATH
    benar), jadi ini engine SATU-SATUNYA yang dipakai baik pas tes lokal
    maupun deployment (lihat penjelasan RENDER_ENGINE di atas soal
    kenapa cuma satu engine, bukan auto-switch).

    Akurasi font-nya dijamin oleh _verify_font_available() yang dipanggil
    LEBIH DULU di render_to_pdf() sebelum fungsi ini jalan -- kalau
    fungsi ini sampai dipanggil, font sudah dipastikan (atau di-skip
    sengaja lewat SKIP_FONT_CHECK).

    CHANGELOG: file disalin dulu ke nama sederhana ("input.docx") dan
    soffice dijalankan dengan profil pengguna SEMENTARA yang bersih
    (-env:UserInstallation). Alasannya: di Windows, soffice (terutama
    build Portable) bisa crash dengan exit code 0xE06D7363 (unhandled C++
    exception) kalau path file mengandung spasi/tanda kutip, atau profil
    bawaannya korup/terkunci oleh instance lain. Hasil PDF SELALU bernama
    "input.pdf" di output_dir -- pemanggil memakai path yang dikembalikan
    fungsi ini, bukan menebak nama dari nama docx.
    """
    import shutil

    work_dir = tempfile.mkdtemp(prefix="lo_work_", dir=LO_TEMP_ROOT)
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_", dir=LO_TEMP_ROOT)
    src = os.path.join(work_dir, "input.docx")
    shutil.copyfile(docx_path, src)
    # as_uri() meng-encode spasi/karakter khusus (%20), beda dengan
    # "file:///" + path mentah yang rusak kalau path berspasi.
    profile_url = pathlib.Path(profile_dir).resolve().as_uri()
    try:
        subprocess.run(
            [
                LIBREOFFICE_PATH,
                f"-env:UserInstallation={profile_url}",
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                output_dir,
                src,
            ],
            check=True,
            capture_output=True,
            timeout=180,
        )
    except FileNotFoundError:
        raise RuntimeError(
            f"LibreOffice ('{LIBREOFFICE_PATH}') tidak ditemukan. "
            "Cek lagi apakah path LIBREOFFICE_PATH sudah sesuai lokasi "
            "soffice di mesin ini (atau set env var LIBREOFFICE_PATH)."
        )
    except subprocess.CalledProcessError as e:
        detail = (e.stderr or b"").decode(errors="ignore").strip() or (
            e.stdout or b""
        ).decode(errors="ignore").strip()
        raise RuntimeError(
            f"soffice keluar dengan kode {e.returncode & 0xFFFFFFFF:#x}. "
            f"{detail or '(tanpa pesan)'} "
            f"[soffice='{LIBREOFFICE_PATH}', temp_root='{LO_TEMP_ROOT or tempfile.gettempdir()}'] "
            "Kode 0xe06d7363 = soffice crash; coba LibreOffice yang "
            "terinstal normal (bukan Portable) lewat env var LIBREOFFICE_PATH."
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("Konversi docx ke pdf melebihi batas waktu (180s).")
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
        shutil.rmtree(profile_dir, ignore_errors=True)

    pdf_path = os.path.join(output_dir, "input.pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(
            "Konversi PDF (LibreOffice) selesai tapi file output tidak ditemukan."
        )
    return pdf_path


def _render_via_word(docx_path: str, output_dir: str) -> str:
    """
    Convert docx -> pdf pakai MS WORD (COM automation). HANYA bisa jalan
    di Windows dengan MS Word terpasang -- TIDAK PORTABLE ke server
    Linux, jadi JANGAN dipakai untuk versi yang di-deploy.

    KENAPA LEBIH AKURAT (untuk tes lokal): PDF dihasilkan oleh program
    YANG SAMA PERSIS dipakai menulis dokumennya, jadi font Times New
    Roman pasti tersedia (tidak ada substitusi) dan wrapping baris
    dijamin identik dengan yang dilihat mahasiswa sendiri di Word.

    CATATAN: import pywin32 SENGAJA di dalam fungsi ini (bukan di top
    of file) supaya modul ini tetap bisa di-import & jalan normal di
    server Linux yang tidak punya pywin32 terpasang sama sekali -- kalau
    RENDER_ENGINE="libreoffice" (mode deployment), fungsi ini tidak
    pernah dipanggil, jadi import pywin32 yang gagal di sini tidak boleh
    bikin seluruh modul gagal di-import duluan.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation cuma bisa jalan di Windows.")

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError(
            "pywin32 belum terpasang -- jalankan `pip install pywin32`, "
            "atau set RENDER_ENGINE=libreoffice untuk pakai LibreOffice saja."
        ) from e

    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, f"{base}.pdf")

    # CoInitialize WAJIB dipanggil di tiap thread yang mau pakai COM --
    # penting kalau fungsi ini dipanggil dari thread request web server,
    # bukan main thread.
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False  # cegah popup "convert format?" nge-hang
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=_WD_FORMAT_PDF)
    except Exception as e:
        raise RuntimeError(f"Gagal convert docx ke pdf (Word COM): {e}")
    finally:
        # try/finally WAJIB: kalau Word.Quit() tidak pernah dipanggil
        # (mis. exception di tengah), proses WINWORD.EXE nyangkut jadi
        # zombie process -- makin numpuk kalau dipanggil berkali-kali.
        if doc is not None:
            doc.Close(SaveChanges=False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()

    if not os.path.exists(pdf_path):
        raise RuntimeError(
            "Konversi PDF (Word COM) selesai tapi file output tidak ditemukan."
        )
    return pdf_path


def render_to_pdf(docx_path: str, output_dir: str | None = None) -> tuple[str, str]:
    """
    Convert docx -> pdf. Entry point tunggal yang dipanggil modul lain
    (check_page_layout dst) -- mereka TIDAK PERLU tahu engine mana yang
    aktif, cuma dapat path pdf jadi.

    DEFAULT & JALUR NORMAL: LibreOffice, SAMA baik dites lokal maupun
    di server produksi -- lihat penjelasan RENDER_ENGINE di atas kenapa
    ini sengaja TIDAK auto-switch ke Word lagi (supaya tes lokal beneran
    mewakili apa yang jalan di production, bukan cuma kelihatan akurat
    gara-gara diam-diam pakai engine lain).

    Font diverifikasi WAJIB duluan (_verify_font_available) sebelum
    render apa pun jalan -- gagal cepat & jelas kalau font yang dipakai
    dokumen PASTI tidak terpasang (RuntimeError, lihat fungsi itu).

    RETURN (BREAKING CHANGE -- lihat CHANGELOG modul ini): sekarang
    tuple (pdf_path, font_status), BUKAN cuma pdf_path polos lagi.
    font_status salah satu dari "verified"/"skipped"/"unconfirmed" (lihat
    _verify_font_available()). CALLER (main.py) WAJIB di-update untuk
    unpack tuple ini DAN meneruskan font_status ke check_page_layout()
    lewat parameter font_status-nya -- kalau tidak diteruskan, seluruh
    manfaat perubahan ini (finding halaman ditandai confidence rendah
    saat font tidak terjamin) hilang percuma, balik seperti sebelumnya
    (finding selalu disajikan seolah pasti benar).

    Jalur Word (RENDER_ENGINE="word") selalu "verified" -- Word memakai
    font aslinya sendiri lewat OS, tidak ada substitusi metric-compatible
    sama sekali, jadi tidak perlu _verify_font_available().
    """
    output_dir = output_dir or tempfile.mkdtemp(
        prefix="page_layout_", dir=LO_TEMP_ROOT
    )

    if RENDER_ENGINE == "word":
        return _render_via_word(docx_path, output_dir), "verified"

    font_status = _verify_font_available()  # gagal cepat kalau font PASTI tidak ada
    pdf_path = _render_via_libreoffice(docx_path, output_dir)
    return pdf_path, font_status


def extract_page_texts(pdf_path: str) -> list[str]:
    """Ambil teks tiap halaman PDF. Index list = halaman fisik - 1."""
    texts = []
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            texts.append(page.get_text())
    return texts


# Tinggi zona footer yang discan, dalam point (1cm ~= 28.35pt).
#
# BUG DIPERBAIKI: nilai lama (60pt) TERLALU KECIL. Diverifikasi langsung
# dari hasil render PDF dokumen asli (Proposal_PA_Bintang_Punta_Dzaky):
# angka footer BENERAN tercetak konsisten di ~64.2pt dari tepi bawah
# halaman (bukan di bawah 60pt), jadi ZONA LAMA MELEWATKAN HAMPIR SEMUA
# ANGKA FOOTER DI DOKUMEN (0 dari 61 halaman ke-detect). Akibatnya
# extract_footer_numbers() balik dict kosong terus, dan check_page_layout
# selalu jatuh ke fallback rumus w:pgNumType (yang sudah ditandai kurang
# akurat di docstring modul ini) -- inilah sumber salah hitung "BAB 1
# berakhir di halaman ganjil (9)" padahal harusnya genap, karena ground
# truth (footer asli) tidak pernah benar-benar dipakai.
#
# 70pt dipilih semula karena ada celah aman terverifikasi: footer asli di
# ~64.2pt, sedangkan baris body teks terdekat ke footer di dokumen yang
# sama ada di ~74.8pt ke atas. 70pt menangkap footer TANPA ikut
# menangkap baris body teks (yang bisa bikin candidates ambigu dan
# malah gagal ke-detect juga).
#
# BUG DITEMUKAN & DIPERBAIKI (2026-09-12, laporan user): 70pt ternyata
# TIDAK CUKUP untuk halaman PERTAMA setiap section baru -- footer di
# halaman itu bisa duduk ~6-7pt lebih TINGGI (lebih jauh dari tepi
# bawah) dibanding halaman lain di section yang sama, kemungkinan
# quirk render LibreOffice/Word soal jarak footer khusus di halaman
# pertama section (mis. dari transisi "Different First Page" atau
# footer-distance yang sedikit bergeser tepat di batas section).
# Kasus nyata (dokumen "Buku_PA_Awal_Raya.pdf"): footer "1" di halaman
# fisik 13 (halaman PERTAMA BAB 1, section baru) duduk di y0=770.72pt
# dari atas (page height 841.92pt) -- cuma 71.2pt dari tepi bawah, TIPIS
# di luar batas 70pt lama (perlu >= page_height-70 = 771.92pt), MELESET
# cuma ~1.2pt. Akibatnya footer halaman itu SATU-SATUNYA yang gagal
# terbaca, dan karena itu PERSIS halaman tempat restart ke "1" terjadi,
# _validate_footer_sequence() di build_page_number_sections() kehilangan
# titik restart itu SAMA SEKALI -- bukan cuma 1 halaman yang salah, tapi
# SELURUH 116 halaman ber-format decimal setelahnya (BAB 1 s.d. akhir
# dokumen) ikut gagal divalidasi sebagai satu rangkaian yang konsisten
# (loncatan roman->decimal di titik itu terlihat seperti urutan rusak
# tanpa ada halaman 13 sebagai jembatan), sehingga compare_page_numbering()
# salah menyimpulkan section BAB 1 masih mewarisi format section
# sebelumnya (lowerRoman) -- padahal dokumen aslinya SUDAH benar (footer
# BAB 1 memang "1" format Arab).
#
# Dinaikkan ke 90pt: masih ada celah aman ke body teks terdekat (~740-
# 747pt dari atas pada dokumen uji, alias ~95-102pt dari tepi bawah --
# di ATAS 90pt, jadi tidak ikut tertangkap), sekaligus cukup longgar
# untuk variasi ~7pt di halaman pertama section yang ditemukan di atas.
FOOTER_ZONE_PT = 100

_ROMAN_VALUES = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
_ROMAN_RE = re.compile(r"^[ivxlcdm]+$", re.IGNORECASE)
_DECIMAL_RE = re.compile(r"^\d+$")


def _roman_to_int(s: str) -> int | None:
    """Konversi angka romawi (huruf kecil) ke int. None kalau tidak valid."""
    s = s.lower()
    total = 0
    prev = 0
    for ch in reversed(s):
        val = _ROMAN_VALUES.get(ch)
        if val is None:
            return None
        if val < prev:
            total -= val
        else:
            total += val
            prev = val
    return total if total > 0 else None


def _footer_pair_agrees(prev_page: int, prev_val: int, page: int, val: int) -> bool:
    """
    Cek apakah satu pasang titik footer (page, val) konsisten terhadap
    titik sebelumnya, mengikuti salah satu dari dua pola valid:
      (a) naik konsisten mengikuti jarak halaman fisik (lanjut dari
          section sebelumnya, atau di tengah satu section w:pgNumType
          yang sama -- w:fmt "continue")
      (b) restart ke 1 (section baru dengan w:pgNumType start=1, mis.
          BAB 1 selalu mengulang penomoran dari halaman "1")
    """
    expected_continue = prev_val + (page - prev_page)
    is_continue = val == expected_continue
    is_restart = val == 1 and page > prev_page
    return is_continue or is_restart


def _is_confirmed_format_transition(
    footer_numbers: dict[int, int],
    footer_fmts: dict[int, str] | None,
    pages_sorted: list[int],
    last_trusted_page: int,
    page: int,
) -> bool:
    """
    True kalau `page` memulai section penomoran baru dengan FORMAT beda
    dari titik terpercaya terakhir (mis. lowerRoman -> decimal) dan titik
    sesudahnya mengonfirmasi (lanjut naik mengikuti jarak halaman, atau
    restart ke 1) pada format yang sama. Nilai awal section baru bebas
    (tidak harus 1) -- apakah harus 1 itu urusan comparator, bukan
    validasi urutan.
    """
    if not footer_fmts:
        return False
    if footer_fmts.get(page) == footer_fmts.get(last_trusted_page):
        return False
    idx = pages_sorted.index(page)
    if idx + 1 >= len(pages_sorted):
        return False
    nxt = pages_sorted[idx + 1]
    if footer_fmts.get(nxt) != footer_fmts.get(page):
        return False
    return _footer_pair_agrees(page, footer_numbers[page], nxt, footer_numbers[nxt])


def _validate_footer_sequence(
    footer_numbers: dict[int, int],
    footer_fmts: dict[int, str] | None = None,
) -> dict[int, int]:
    """
    Footer yang terbaca "meyakinkan" oleh extract_footer_numbers() (tepat
    satu kandidat konsisten di zona footer) TETAP bisa salah kalau field
    PAGE Word sendiri salah hitung ulang saat render -- kasus nyata yang
    ditemukan: field PAGE di suatu halaman nyasar ke angka yang melenceng
    gara-gara ada tabel/gambar besar di dekat footer yang mengacaukan
    urutan render field. Ini bukan salah baca koordinat (extract_footer_
    numbers sudah benar membaca ANGKA YANG TERCETAK), tapi angka yang
    tercetak itu sendiri yang keliru di sumbernya -- ground truth "kotor"
    di satu halaman itu.

    Satu-satunya cara mendeteksi ini tanpa akses ke isi Word asli:
    bandingkan tiap angka ke angka TERPERCAYA TERAKHIR sebelumnya (bukan
    cuma tetangga fisik langsung, karena halaman di antaranya bisa saja
    sudah didrop lebih dulu karena ambigu di extract_footer_numbers).
    Kalau suatu angka TIDAK cocok pola manapun (lihat _footer_pair_agrees)
    relatif terhadap angka terpercaya terakhir, angka itu DI-DROP dari
    hasil -- caller (check_page_layout -> to_displayed()) lalu fallback
    ke rumus w:pgNumType KHUSUS untuk halaman itu saja, sementara seluruh
    halaman lain yang footernya konsisten tetap pakai ground truth asli.

    BUG DITEMUKAN (revisi ini -- baseline yang tidak terkonfirmasi):
    versi sebelumnya mempercayai TITIK PERTAMA yang berhasil terbaca
    di seluruh dokumen sebagai baseline TANPA pembanding apa pun (wajar
    di awal, karena memang belum ada titik sebelumnya). Masalahnya:
    kalau titik pertama itu SENDIRI adalah halaman yang field PAGE-nya
    nyasar (skenario "tabel/gambar besar dekat footer" yang sama seperti
    di atas, tapi apesnya kena persis di halaman paling awal yang
    berhasil terbaca), maka baseline yang salah itu tidak pernah
    tertangkap -- karena semua validasi sesudahnya cuma cek KONSISTENSI
    RELATIF terhadap baseline, bukan kebenaran absolut. Setiap titik
    setelahnya akan tetap "konsisten naik mengikuti jarak halaman" *relatif
    terhadap baseline yang salah*, sehingga OFFSET KONSTAN dari baseline
    itu ikut terbawa ke SELURUH rantai footer di belakangnya, termasuk ke
    halaman-halaman yang dipakai buat cek "bab berakhir di halaman
    ganjil/genap". Kasus nyata: BAB 1 yang aslinya berakhir di halaman
    genap (8) malah dilaporkan berakhir di halaman ganjil (9) -- offset
    +1 konstan yang sumbernya bukan di sekitar BAB 1 sama sekali,
    melainkan dari baseline di awal dokumen yang meleset.

    FIX: baseline sekarang WAJIB dikonfirmasi lebih dulu oleh MINIMAL
    SATU titik berikutnya yang konsisten dengannya (pola yang sama persis
    dipakai untuk validasi titik-titik selanjutnya). Kalau titik pertama
    tidak terkonfirmasi oleh titik kedua, titik pertama di-drop (diperlakukan
    sebagai glitch satu halaman, sama seperti glitch yang sudah ditangani
    di tengah dokumen), lalu titik kedua dicoba jadi baseline baru
    dibanding titik ketiga, dst -- sampai ketemu dua titik berturut-turut
    yang saling konsisten (baseline terkonfirmasi), atau kandidat habis.
    Kalau sampai habis tanpa ada dua titik yang saling konfirmasi sama
    sekali, TIDAK ADA satu pun rantai footer yang cukup bisa dipercaya --
    fungsi ini return dict kosong, dan caller full-fallback ke rumus
    w:pgNumType untuk seluruh dokumen (lebih aman daripada memercayai
    ground truth yang tidak terverifikasi sama sekali).

    CATATAN: ini murni statistical/sequential check terhadap titik
    terpercaya terakhir yang VALID -- kalau DUA kesalahan berurutan
    kebetulan "konsisten" satu sama lain (naik sesuai jarak di antara
    keduanya, padahal dua-duanya sama-sama salah), itu tetap tidak akan
    tertangkap. Kasus ini dianggap cukup jarang untuk tidak ditangani di
    sini; kalau muncul lagi, perlu heuristik tambahan (mis. cross-check
    ke rumus w:pgNumType sebagai sinyal sekunder).
    """
    if not footer_numbers:
        return {}

    pages_sorted = sorted(footer_numbers)

    # --- Cari baseline yang TERKONFIRMASI oleh titik berikutnya ---
    baseline_idx = None
    for i in range(len(pages_sorted) - 1):
        p1, p2 = pages_sorted[i], pages_sorted[i + 1]
        if _footer_pair_agrees(p1, footer_numbers[p1], p2, footer_numbers[p2]):
            baseline_idx = i
            break

    if baseline_idx is None:
        # Tidak ada dua titik berurutan yang saling konfirmasi sama
        # sekali (mis. cuma satu footer berhasil terbaca di seluruh
        # dokumen, atau semua kandidat yang terbaca saling
        # bertentangan) -- tidak ada dasar yang cukup kuat buat
        # dipercaya sebagai ground truth.
        return {}

    validated: dict[int, int] = {}
    last_trusted_page = pages_sorted[baseline_idx]
    last_trusted_val = footer_numbers[last_trusted_page]
    validated[last_trusted_page] = last_trusted_val

    for phys_page in pages_sorted[baseline_idx + 1 :]:
        val = footer_numbers[phys_page]
        if _footer_pair_agrees(last_trusted_page, last_trusted_val, phys_page, val):
            validated[phys_page] = val
            last_trusted_page, last_trusted_val = phys_page, val
        elif _is_confirmed_format_transition(
            footer_numbers, footer_fmts, pages_sorted, last_trusted_page, phys_page
        ):
            # FIX: transisi format (romawi -> arab) dengan nilai awal
            # BEBAS (mis. BAB 1 mulai dari 4, bukan 1) -- bukan glitch
            # field PAGE, karena halaman berikutnya konsisten dengannya.
            # Tanpa ini seluruh halaman arab ikut ke-drop dan BAB 1
            # dianggap masih romawi (False finding "Romawi kecil").
            validated[phys_page] = val
            last_trusted_page, last_trusted_val = phys_page, val
        # else: menyimpang dari kedua pola -> field PAGE di halaman ini
        # tidak bisa dipercaya, DROP. last_trusted_page/val SENGAJA tidak
        # diupdate supaya halaman berikutnya tetap dibandingkan ke titik
        # terpercaya terakhir yang valid, bukan ke angka yang baru saja
        # kita curigai salah.

    return validated


_PT_TO_CM = 2.54 / 72.0


def extract_footer_details(pdf_path: str) -> dict[int, dict]:
    """
    Versi LEBIH LENGKAP dari extract_footer_numbers(): selain nilai int,
    ikut menyimpan FORMAT token asli ("lowerRoman"/"decimal") dan apakah
    posisinya rata tengah horizontal -- dua info yang sebelumnya SELALU
    DIBUANG (extract_footer_numbers() versi lama langsung convert semua
    ke int polos begitu ketemu, format & posisi tidak pernah disimpan).

    KENAPA INI DIBUTUHKAN (root cause skor "Umum (Penomoran Halaman)"
    selalu 0% di jalur PDF, terlepas dokumennya benar atau salah):
    compare_page_numbering() di comparator.py butuh 3 hal per section --
    pgnum_fmt (romawi/arab), pgnum_start (restart ke berapa), dan
    footer_has_page_field + footer_alignment (field nomor ada & rata
    tengah) -- tapi extract_section_params_pdf() di pdf_extractor.py
    SENGAJA isi keempatnya None (didelegasikan ke modul ini), padahal
    modul ini SEBELUMNYA tidak pernah benar-benar menghitung salah satu
    dari empat itu -- extract_footer_numbers() cuma menjawab "nomor
    tercetaknya berapa", bukan "formatnya apa" atau "posisinya di
    mana". Akibatnya keempat field itu None SELAMANYA di jalur PDF, dan
    cek 0/1/2 di compare_page_numbering() PASTI gagal apapun isi
    dokumennya. Fungsi ini (dipakai oleh build_page_number_sections() di
    bawah) mengisi kekosongan itu.

    Kandidat & validasi urutan memakai kriteria PERSIS SAMA seperti
    extract_footer_numbers() (satu kandidat meyakinkan per halaman +
    _validate_footer_sequence()) -- lihat docstring extract_footer_numbers()
    untuk alasan lengkap kenapa dibaca dari koordinat, bukan reading
    order teks.

    centered: titik-tengah horizontal TOKEN angka (bukan seluruh
    footer_zone) dibanding titik-tengah lebar halaman, toleransi 0.5cm
    -- lebih longgar dari _infer_alignment() paragraf (pdf_extractor.py)
    karena token angka jauh lebih sempit dari kolom teks, jadi noise
    render (kerning, font substitusi LibreOffice) proporsinya lebih
    besar relatif terhadap lebar elemen kecil ini.

    has_page_field: TIDAK disimpan di sini secara eksplisit (PDF tidak
    bisa membedakan field PAGE beneran vs angka yang diketik manual) --
    caller (build_page_number_sections()) menyimpulkannya sebagai True
    kalau halaman itu ADA di dict return ini sama sekali (ada kandidat
    meyakinkan yang lolos validasi urutan). Best-effort, bukan jaminan.

    Return {halaman_fisik: {"value": int, "fmt": "lowerRoman"|"decimal",
    "centered": bool}}. Halaman yang gagal terbaca / ambigu / didrop
    validasi urutan TIDAK punya entri sama sekali (caller anggap
    "nomor tidak meyakinkan di halaman ini", bukan "pasti tidak ada").
    """
    tol_center_pt = 0.5 / _PT_TO_CM

    raw: dict[int, dict] = {}
    with fitz.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf, start=1):
            page_height = page.rect.height
            page_center_x = page.rect.width / 2
            # (x0, y0, x1, y1, text, block_no, line_no, word_no)
            words = page.get_text("words")
            footer_zone = [w for w in words if w[1] >= page_height - FOOTER_ZONE_PT]

            candidates = []
            for w in footer_zone:
                text = w[4].strip()
                x0, x1 = w[0], w[2]
                if _DECIMAL_RE.match(text):
                    candidates.append((int(text), "decimal", x0, x1))
                elif _ROMAN_RE.match(text):
                    val = _roman_to_int(text)
                    if val is not None:
                        candidates.append((val, "lowerRoman", x0, x1))

            unique_vals = {c[0] for c in candidates}
            if len(unique_vals) == 1:
                value, fmt, x0, x1 = candidates[0]
                token_center = (x0 + x1) / 2
                centered = abs(token_center - page_center_x) <= tol_center_pt
                raw[page_num] = {"value": value, "fmt": fmt, "centered": centered}
            # len == 0 -> tidak ada kandidat (lewati)
            # len > 1  -> ambigu (lewati)

    # Validasi urutan pakai FUNGSI YANG SAMA dengan extract_footer_numbers()
    # (beroperasi di {page: value}) -- lalu halaman yang di-drop di situ
    # ikut dibuang dari detail juga, supaya dua fungsi ini tidak pernah
    # tidak-sinkron satu sama lain (satu bilang halaman X terpercaya,
    # yang lain bilang tidak).
    value_only = {p: d["value"] for p, d in raw.items()}
    fmt_only = {p: d["fmt"] for p, d in raw.items()}
    validated_values = _validate_footer_sequence(value_only, fmt_only)
    return {p: raw[p] for p in validated_values}


def extract_footer_numbers(pdf_path: str) -> dict[int, int]:
    """
    Baca ANGKA TERCETAK di footer LANGSUNG dari posisi fisiknya di tiap
    halaman PDF -- ground truth paling akurat untuk "nomor halaman yang
    benar-benar dilihat pembaca", karena tidak bergantung pada asumsi
    apa pun soal berapa banyak halaman fisik yang seharusnya dihasilkan
    render (asumsi itu bisa meleset kalau environment render beda font
    dengan environment penulisan dokumen -- lihat CHANGELOG di atas).

    KENAPA PAKAI KOORDINAT (posisi fisik), BUKAN "ambil baris terakhir
    dari page.get_text()": kalau ada tabel/gambar besar dekat footer,
    PyMuPDF bisa mengembalikan teks tabel SETELAH footer dalam urutan
    linier (reading order) hasil ekstraksinya -- "baris terakhir" jadi
    salah tangkap isi sel tabel, bukan footer. Filter berbasis
    koordinat y (posisi fisik di halaman, bukan urutan teks hasil
    ekstraksi) kebal terhadap masalah susunan-linier ini.

    Return {halaman_fisik: nilai_int}. Nilai romawi (i, ii, iii, ...)
    dikonversi ke int (1, 2, 3, ...) juga, supaya caller bisa langsung
    bandingkan ganjil/genap tanpa peduli formatnya. Halaman yang TIDAK
    ketemu kandidat footer yang meyakinkan (tidak ada angka sama sekali
    di zona footer -- mis. halaman sampul -- ATAU ketemu lebih dari
    satu angka BERBEDA di zona footer -- kasus ambigu, mis. ada data
    tabel yang kebetulan mepet ke footer) TIDAK dimasukkan ke dict.

    Sebelum dikembalikan, hasil ini disaring lagi lewat
    _validate_footer_sequence() untuk menangkap kasus field PAGE yang
    salah hitung ulang di satu halaman spesifik (mis. gara-gara tabel
    besar dekat footer) -- lihat docstring fungsi itu untuk detail pola
    validasinya. Jadi caller WAJIB tetap fallback ke nilai hasil rumus
    w:pgNumType untuk halaman yang tidak ada di sini -- baik karena
    tidak terbaca sama sekali, ambigu, MAUPUN karena didrop validasi
    urutan -- ini bukan best-effort mengira-ngira, tapi "hanya percaya
    angka yang ekstraksinya tidak ambigu DAN konsisten urutannya".

    IMPLEMENTASI (diubah 2026-08-20): sekarang cuma delegate ke
    extract_footer_details() dan ambil field "value"-nya saja -- dulu
    dua fungsi ini py logic scan kandidat SENDIRI-SENDIRI (duplikat),
    resikonya kalau salah satu diubah/dipatch, yang lain bisa diam-diam
    ketinggalan dan dua fungsi ini jadi tidak sinkron. Sekarang cuma ada
    SATU tempat yang scan kandidat mentah dari PDF.
    """
    return {p: d["value"] for p, d in extract_footer_details(pdf_path).items()}


def build_page_number_sections(pdf_path: str, margin_section: dict) -> list[dict]:
    """
    Bangun ULANG `sections` versi PDF -- pengganti section tunggal
    ber-None (pgnum_fmt/pgnum_start/footer_has_page_field/
    footer_alignment) dari extract_section_params_pdf() di
    pdf_extractor.py, yang memang cuma fokus ke margin geometri dan
    SENGAJA mendelegasikan 4 field itu ke modul ini (lihat komentar di
    situ) -- tapi sebelum fungsi ini ada, tidak ada satu pun kode yang
    benar-benar mengisinya, jadi compare_page_numbering() di
    comparator.py SELALU gagal di cek 0/1/2, apapun isi dokumennya.

    Section dipecah di titik TRANSISI: format berubah (romawi<->arab)
    ATAU nilai footer RESTART ke 1 -- dua-duanya meniru semantik SATU-
    SATUNYA alasan sah w:pgNumType section break berganti di docx asli.
    Section pertama otomatis mulai dari halaman fisik PERTAMA yang
    footernya berhasil dibaca meyakinkan (lihat extract_footer_details).

    margin_section: dict TUNGGAL hasil extract_section_params_pdf(pdf_path)[0]
    -- field geometri (left_cm/left_twips/top_cm/dst, width/height) di-
    COPY apa adanya ke SETIAP section baru di sini, tidak dihitung ulang.
    Ini SENGAJA: margin adalah SATU pengaturan untuk seluruh dokumen
    (lihat catatan di compare_sections(), comparator.py -- section break
    biasa dipakai mahasiswa untuk kebutuhan LAIN, bukan sengaja ganti
    margin), sementara section di fungsi ini murni soal penomoran
    halaman -- dua concern yang independen, makanya digabung lewat copy
    geometri, bukan dihitung ulang dari nol.

    CARA PAKAI: JANGAN panggil fungsi ini sendirian di main.py -- pakai
    finalize_page_numbering() di bawah, yang juga membereskan wiring
    "section" tiap paragraf (assign_paragraph_sections()). Kalau cuma
    build_page_number_sections() saja yang dipanggil tanpa itu,
    paragraphs["section"] tetap 0 semua (warisan lama) dan
    bab1_section_idx di compare_page_numbering() jadi salah lagi.

    Fallback: kalau TIDAK ADA satu pun footer yang bisa dipercaya di
    seluruh dokumen (mis. dokumen memang tidak punya nomor halaman),
    balik 1 section (copy margin_section) dengan pgnum_fmt/pgnum_start
    =None & footer_has_page_field=False -- ini SECARA JUJUR melaporkan
    "footer memang tidak ada", beda maknanya dari None versi lama (yang
    artinya "belum sempat dicek sama sekali").
    """

    def _base(index: int, start_physical_page: int) -> dict:
        base = dict(margin_section)
        base["index"] = index
        base["start_physical_page"] = start_physical_page
        return base

    footer_details = extract_footer_details(pdf_path)
    if not footer_details:
        sec = _base(0, 1)
        sec.update(
            pgnum_fmt=None,
            pgnum_start=None,
            footer_has_page_field=False,
            footer_alignment=None,
        )
        return [sec]

    pages_sorted = sorted(footer_details)
    sections: list[dict] = []
    cur: dict | None = None
    cur_centered_votes: list[bool] = []
    cur_last_value: int | None = None

    def _flush_alignment():
        if cur is not None and cur_centered_votes:
            cur["footer_alignment"] = (
                "center"
                if sum(cur_centered_votes) / len(cur_centered_votes) >= 0.5
                else "left"
            )

    for p in pages_sorted:
        d = footer_details[p]
        is_restart = cur is not None and d["value"] == 1 and cur_last_value != 1
        is_new_section = cur is None or d["fmt"] != cur["pgnum_fmt"] or is_restart
        if is_new_section:
            _flush_alignment()
            cur = _base(len(sections), p)
            cur.update(
                pgnum_fmt=d["fmt"],
                pgnum_start=d["value"],
                footer_has_page_field=True,
                # Nomor TERCETAK tiap halaman fisik yang terbaca di section
                # ini. Dipakai compare_page_numbering() untuk membaca nomor
                # di halaman tempat BAB 1 BENAR-BENAR mulai, bukan nomor
                # halaman pertama section (yang bisa lebih awal dari BAB 1,
                # mis. halaman kosong/pra-bab yang sudah berformat Arab).
                page_values={},
            )
            sections.append(cur)
            cur_centered_votes = []
        cur_last_value = d["value"]
        cur["page_values"][p] = d["value"]
        cur_centered_votes.append(d["centered"])
    _flush_alignment()

    return sections


def assign_paragraph_sections(paragraphs: list[dict], pn_sections: list[dict]) -> None:
    """
    Timpa elem["section"] TIAP paragraf supaya match section penomoran
    dari build_page_number_sections(), bukan 0 hardcoded (nilai bawaan
    dari extract_paragraph_params_pdf() di pdf_extractor.py -- lihat
    "section": 0 di situ, yang memang benar untuk kebutuhan margin
    (compare_sections tidak peduli per-section), tapi SALAH untuk
    kebutuhan penomoran halaman (compare_page_numbering() butuh tahu
    paragraf "BAB 1" ada di section MANA, buat cari pre_bab1_sections).

    Dipetakan dari elem["page_index"] (0-indexed, dari pdf_extractor.py)
    -> halaman fisik (1-indexed) -> section yang "menaungi" halaman itu
    (section dengan start_physical_page terbesar yang masih <= halaman
    tsb) -- logic sama seperti _section_governing() di check_page_layout()
    bawah, sengaja diduplikasi kecil di sini karena beda struktur input
    (pn_sections punya start_physical_page langsung, bukan section_pages
    dict terpisah).

    In-place (memodifikasi dict paragraf yang ADA, tidak bikin list
    baru) -- WAJIB dipanggil SEBELUM comparator.compare_all(), dan
    idealnya lewat finalize_page_numbering() di bawah supaya urutan
    panggilannya tidak salah.
    """
    starts = sorted((sec["start_physical_page"], sec["index"]) for sec in pn_sections)
    if not starts:
        return
    for elem in paragraphs:
        physical_page = elem["page_index"] + 1
        governing_index = starts[0][1]
        for start_page, idx in starts:
            if start_page <= physical_page:
                governing_index = idx
            else:
                break
        elem["section"] = governing_index


def finalize_page_numbering(extracted: dict, pdf_path: str) -> dict:
    """
    Convenience wrapper -- gabungkan 3 langkah wiring yang WAJIB
    dijalankan di main.py SETELAH pdf_extractor.extract_all_pdf() dan
    SEBELUM comparator.compare_all(), atau skor "Umum (Penomoran
    Halaman)" akan TETAP 0% seperti sebelumnya (lihat root-cause
    lengkap di docstring build_page_number_sections()):

      1. build_page_number_sections() -- bangun section penomoran
         beneran (bukan 1 section ber-None) dari footer PDF asli.
      2. assign_paragraph_sections()  -- perbaiki elem["section"] tiap
         paragraf supaya cocok dengan section barusan (bukan 0 semua).
      3. Ganti extracted["sections"] dengan section baru itu -- dipakai
         BERSAMA oleh compare_sections() (margin) dan
         compare_page_numbering() (nomor halaman) di comparator.py,
         lihat compare_all().

    Memodifikasi `extracted` IN-PLACE (paragraphs & sections di
    dalamnya) dan mengembalikan objek yang sama, murni untuk
    kenyamanan chaining -- tidak membuat dict baru.

    Contoh pemakaian di main.py:
        extracted = pdf_extractor.extract_all_pdf(pdf_path)
        extracted = page_layout_checker.finalize_page_numbering(extracted, pdf_path)
        findings, checks = comparator.compare_all(extracted, section_pages)
    """
    margin_section = extracted["sections"][0]
    pn_sections = build_page_number_sections(pdf_path, margin_section)
    assign_paragraph_sections(extracted["paragraphs"], pn_sections)
    extracted["sections"] = pn_sections
    return extracted


# Baris ToC (Daftar Isi/Gambar/Tabel) buat "BAB N ..." punya dot-leader
# ("....") atau spasi tab-stop panjang sebelum nomor halaman di ujung
# baris -- pola ini TIDAK PERNAH muncul di heading "BAB N" ASLI (yang
# cuma diisi "BAB N" polos atau "BAB N NAMA BAB", tanpa dot-leader/
# nomor halaman menempel di baris yang sama).
#
# BUG DITEMUKAN: tanpa guard ini, find_chapter_boundaries() bisa nangkep
# baris Daftar Isi sebagai boundary "BAB N" -- CHAPTER_TITLE_RE (grup
# opsional "(\s+\S.*)?") menerima APAPUN setelah "BAB N", termasuk
# dot-leader+nomor halaman ToC (mis. "BAB 1 PENDAHULUAN.......... 1"),
# dan baris itu tetap ALL CAPS (dot & angka tidak berubah oleh
# .upper()) sehingga lolos syarat "stripped == upper" juga. Beda dari
# guard POST_BAB_ONLY_HEADINGS (DAFTAR PUSTAKA/LAMPIRAN) yang bisa
# pakai state seen_bab, "BAB N" sendiri tidak bisa disaring dengan cara
# yang sama -- makanya butuh guard TERPISAH berbasis pola ToC ini.
# Efek bug ini: start_page BAB N (dan seterusnya, berantai ke boundary
# berikutnya) ke-ambil dari halaman Daftar Isi (jauh lebih awal),
# membuat SEMUA cek ganjil/genap & halaman-notice yang bergantung pada
# boundary ini salah -- inilah sumber ketidakkonsistenan yang dilaporkan.
TOC_TAIL_RE = re.compile(r"\.{2,}|\s{3,}[ivxlcdm\d]+\s*$", re.IGNORECASE)


def _is_chapter_boundary_line(line: str) -> str | None:
    """
    Cek satu baris teks: apakah ini heading setingkat bab? Pakai regex
    & FRONT_HEADINGS yang SAMA dengan extractor.py. Return label kalau
    cocok, None kalau tidak.

    Baris dengan penanda ToC (dot-leader / tab-stop + nomor halaman di
    ujung) ditolak LEBIH DULU, sebelum dicocokkan ke CHAPTER_TITLE_RE
    -- lihat TOC_TAIL_RE di atas untuk alasan lengkap.
    """
    stripped = line.strip()
    if not stripped:
        return None
    if TOC_TAIL_RE.search(stripped):
        return None
    upper = stripped.upper()

    if CHAPTER_TITLE_RE.match(upper) and stripped == upper:
        return upper
    if upper in FRONT_HEADINGS and stripped == upper:
        return upper
    return None


# Heading yang secara struktur SELALU muncul setelah minimal satu
# "BAB N" (bukan sebelum BAB 1). Dipakai buat nyaring false positive
# di find_chapter_boundaries(): entri Daftar Isi utk "DAFTAR PUSTAKA"
# kadang ke-render sebagai baris berdiri sendiri tanpa dot-leader/
# nomor halaman (tabstop-nya gak ketangkep di extractor teks PDF),
# jadi persis sama dengan heading asli secara tekstual. Kalau label
# begini ketemu SEBELUM ada BAB N sama sekali, hampir pasti itu entri
# Daftar Isi yang salah tangkap, BUKAN heading section yang sebenarnya
# -- jadi diabaikan di sini, biar boundary asalinya (yang benar-benar
# di halaman Daftar Pustaka) yang kepakai.
POST_BAB_ONLY_HEADINGS = {"DAFTAR PUSTAKA", "DAFTAR LAMPIRAN"}


def find_chapter_boundaries(page_texts: list[str]) -> list[dict]:
    """
    Return list terurut berdasarkan urutan kemunculan di dokumen:
    [{"label": "BAB 1", "start_page": 5}, ...]. start_page 1-indexed
    (nomor halaman fisik cetak), bukan index array.
    """
    boundaries = []
    seen_labels = set()
    seen_bab = False
    for page_num, text in enumerate(page_texts, start=1):
        for line in text.splitlines():
            label = _is_chapter_boundary_line(line)
            if not label or label in seen_labels:
                continue
            if label in POST_BAB_ONLY_HEADINGS and not seen_bab:
                # Kemunculan "duluan" sebelum BAB N mana pun -- abaikan,
                # ini entri Daftar Isi yang salah tangkap (lihat
                # docstring POST_BAB_ONLY_HEADINGS), JANGAN masukin ke
                # seen_labels supaya kemunculan aslinya nanti tetap
                # bisa ke-capture sebagai boundary.
                continue
            if CHAPTER_TITLE_RE.match(label):
                seen_bab = True
            boundaries.append({"label": label, "start_page": page_num})
            seen_labels.add(label)
            break  # satu heading per halaman cukup buat jadi boundary
    return boundaries


def find_section_start_pages(
    page_texts: list[str], paragraphs: list[dict]
) -> dict[int, int]:
    """
    Generalisasi dari find_chapter_boundaries(): kalau itu cuma nyari
    heading setingkat bab, ini nyari SEMBARANG section index dengan
    mencocokkan cuplikan teks paragraf PERTAMA di tiap section terhadap
    teks tiap halaman render, berurutan sesuai urutan dokumen (section
    N+1 dicari MULAI dari halaman tempat section N ketemu, bukan dari
    awal lagi -- selain lebih cepat, ini juga mencegah cuplikan pendek
    nyasar ke halaman lain yang kebetulan mirip di bagian dokumen yang
    jauh berbeda).

    Return {section_index: halaman_fisik_1_indexed}. Section yang
    cuplikan teksnya tidak ketemu di render (kasus jarang -- mis. beda
    hasil wrapping bikin substring pecah) tidak masuk dict, bukan error,
    supaya section lain yang berhasil match tetap dilaporkan.
    """
    section_first_text = {}
    for p in paragraphs:
        sec = p["section"]
        if sec not in section_first_text and p["full_text"].strip():
            snippet = " ".join(p["full_text"].split())[:30]
            section_first_text[sec] = snippet

    result = {}
    search_from = 0
    for sec_idx in sorted(section_first_text):
        snippet = section_first_text[sec_idx]
        for page_num in range(search_from, len(page_texts)):
            normalized = " ".join(page_texts[page_num].split())
            if snippet in normalized:
                result[sec_idx] = page_num + 1
                search_from = page_num
                break
    return result


def build_displayed_page_map(
    sections: list[dict], section_pages: dict[int, int], total_pages: int
) -> dict[int, int]:
    """
    Bangun peta {halaman_fisik: nomor_halaman_TERCETAK}.

    KOREKSI PENTING (bug sebelumnya): cek "bab harus mulai di halaman
    ganjil" itu soal nomor yang TERTULIS di footer halaman -- angka
    yang dipakai pembimbing buat nyari "halaman N" -- BUKAN urutan
    fisik lembar dari awal dokumen (posisi ke berapa hasil render PDF).
    Begitu suatu section restart penomoran (w:pgNumType start=N, kasus
    umum: BAB 1 selalu restart ke 1), halaman fisik ke-8 di render bisa
    aja punya nomor tercetak "1" -- itu GANJIL, walau posisi fisiknya
    genap. Versi lama pakai posisi fisik mentah buat cek ganjil/genap,
    jadi false-positive tiap kali ada bab yang restart start=1 tapi
    posisi fisiknya genap (mis. gara-gara Pra-Bab yang romawi makan
    banyak halaman).

    section_pages: {section_index: halaman_fisik_mulai}, dari
    find_section_start_pages(). Section yang snippet-nya gak ketemu di
    render (kasus jarang) dilewati di sini -- caller fallback ke nomor
    fisik apa adanya untuk halaman yang gak ke-cover peta ini
    (best-effort, bukan error).
    """
    sections_by_index = {sec["index"]: sec for sec in sections}
    sorted_starts = sorted(
        (
            (sec_idx, phys_start)
            for sec_idx, phys_start in section_pages.items()
            if sec_idx in sections_by_index
        ),
        key=lambda pair: pair[1],
    )

    displayed_map = {}
    running_last_displayed = 0  # nomor tercetak terakhir yang sudah "dipakai"

    for i, (sec_idx, phys_start) in enumerate(sorted_starts):
        phys_end = (
            sorted_starts[i + 1][1] - 1 if i + 1 < len(sorted_starts) else total_pages
        )
        sec = sections_by_index[sec_idx]
        # pgnum_start None berarti LANJUT dari section sebelumnya (bukan
        # restart) -- perilaku asli OOXML, simetris dengan w:fmt.
        disp_start = (
            sec["pgnum_start"]
            if sec["pgnum_start"] is not None
            else running_last_displayed + 1
        )
        for offset, phys in enumerate(range(phys_start, phys_end + 1)):
            displayed_map[phys] = disp_start + offset
        running_last_displayed = disp_start + (phys_end - phys_start)

    return displayed_map


def _new_counter():
    return {"checked": 0, "failed": 0}


def _is_notice_page(text: str) -> bool:
    """
    Cek apakah satu halaman adalah halaman kosong-sengaja: isinya HANYA
    (atau nyaris hanya) tulisan "Halaman ini sengaja dikosongkan", bukan
    halaman konten yang KEBETULAN menyebut frasa itu di suatu tempat.
    """
    normalized = " ".join(text.split()).upper()
    return (
        BLANK_PAGE_NOTICE in normalized
        and len(normalized) <= len(BLANK_PAGE_NOTICE) + 20
    )


_LOW_CONFIDENCE_FONT_STATUSES = {"skipped", "unconfirmed"}

_LOW_CONFIDENCE_NOTE = (
    " (CATATAN: font 'Times New Roman' tidak terverifikasi terpasang di "
    "server yang merender dokumen ini, jadi jumlah halaman fisik hasil "
    "render berpotensi bergeser dari yang mahasiswa lihat sendiri di "
    "Word -- mohon cek manual halaman terkait di Word sebelum "
    "menindaklanjuti temuan ini.)"
)


def check_page_layout(
    page_texts: list[str],
    sections: list[dict],
    section_pages: dict[int, int],
    footer_numbers: dict[int, int] | None = None,
    font_status: str = "unconfirmed",
) -> tuple[list[dict], dict]:
    """
    Entry point tahap Render & Locate. Return (findings, checks) dengan
    bentuk sama seperti comparator.compare_all.

    font_status: status dari render_to_pdf()/_verify_font_available()
    ("verified"/"skipped"/"unconfirmed") untuk render PDF yang dipakai
    menghasilkan page_texts/section_pages/footer_numbers di atas. Kalau
    BUKAN "verified", SETIAP finding yang dihasilkan fungsi ini (semua
    bergantung pada asumsi wrapping halaman render == wrapping di Word)
    ditandai "confidence": "rendah" + catatan eksplisit ditambahkan ke
    description-nya -- BUKAN dihapus atau disembunyikan (pembimbing
    tetap perlu tahu potensi masalahnya), cuma diberi sinyal supaya
    diverifikasi manual dulu sebelum ditindaklanjuti sebagai temuan
    pasti. Default sengaja "unconfirmed" (bukan "verified") supaya
    caller yang LUPA meneruskan font_status dari render_to_pdf() gagal
    ke arah AMAN (terlalu banyak memberi catatan kehati-hatian) alih-
    alih diam-diam menganggap semuanya pasti benar seperti sebelum
    perbaikan ini ada.

    Lihat CHANGELOG modul ini (kasus Proposal_PA_Bintang_Punta_Dzaky)
    untuk kronologi lengkap kenapa flag ini perlu ada.

    CHANGELOG: dulu fungsi ini nerima docx_path dan ngerender PDF-nya
    SENDIRI di dalam. Sekarang nerima page_texts yang SUDAH dirender --
    render-nya dipindah jadi tanggung jawab main.py, supaya PDF yang
    sama bisa dipakai bareng compare_page_numbering() (butuh nomor
    halaman fisik juga) tanpa render 2x dokumen yang sama.

    KOREKSI (bug sebelumnya #1): cek ganjil/genap sekarang pakai NOMOR
    HALAMAN TERCETAK (hasil build_displayed_page_map(), yang menghitung
    ulang berdasarkan w:pgNumType tiap section), BUKAN posisi fisik
    lembar di render PDF. Dua hal ini beda kalau ada section yang
    restart penomoran (mis. BAB 1 selalu restart ke halaman "1") --
    posisi fisik lembar ke-8 bisa aja punya nomor tercetak "1", dan
    "1" itu yang relevan buat aturan "bab wajib mulai di halaman
    ganjil" (aturan ini soal nomor yang dibaca orang di footer, bukan
    urutan fisik lembar). sections & section_pages dibutuhkan buat
    bikin peta konversi fisik->tercetak ini.

    KOREKSI (bug sebelumnya #2): peta di atas MENGHITUNG nomor tercetak
    dari rumus (asumsi jumlah halaman fisik render == jumlah halaman
    yang mahasiswa lihat di Word). Asumsi ini bisa meleset kalau
    environment render (LibreOffice) tidak punya font yang sama persis
    dengan yang dipakai menulis dokumen (mis. "Times New Roman" tidak
    terpasang di server -> LO substitusi font lain -> wrapping baris
    beda -> jumlah halaman fisik ikut beda). footer_numbers (dari
    extract_footer_numbers(), BACA LANGSUNG angka di footer per
    koordinat fisik halaman, bukan hasil kalkulasi) dipakai sebagai
    GROUND TRUTH prioritas utama; rumus w:pgNumType di atas cuma jadi
    fallback untuk halaman yang footernya gagal terbaca dengan yakin
    (mis. halaman sampul, atau kasus ambigu).

    KOREKSI (bug sebelumnya #3): ground truth footer itu sendiri bisa
    kotor di satu halaman spesifik (field PAGE Word salah hitung ulang
    saat render, mis. gara-gara tabel besar dekat footer). Ini sudah
    ditangani DI DALAM extract_footer_numbers() lewat
    _validate_footer_sequence() -- jadi footer_numbers yang diterima
    fungsi ini SUDAH bersih dari angka yang menyimpang urutan; fungsi
    ini sendiri tidak perlu tahu detail itu, cukup pakai footer_numbers
    apa adanya sebagai prioritas utama.

    Existence check (apakah suatu halaman fisik ADA / dalam batas
    dokumen, dan apakah suatu halaman fisik ADALAH halaman notice)
    tetap dilakukan di ruang FISIK (karena page_texts & notice_pages
    inherently fisik) -- cuma parity ganjil/genap dan string lokasi ke
    user yang dikonversi ke nomor tercetak.
    """
    findings = []
    checks = defaultdict(_new_counter)
    page_number_reliable = font_status not in _LOW_CONFIDENCE_FONT_STATUSES

    def _add_finding(chapter_key: str, location: str, description: str) -> None:
        findings.append(
            {
                "category": "Format",
                "chapter": chapter_key,
                "location": location,
                "description": (
                    description
                    if page_number_reliable
                    else description + _LOW_CONFIDENCE_NOTE
                ),
                "confidence": "tinggi" if page_number_reliable else "rendah",
            }
        )

    total_pages = len(page_texts)
    boundaries = find_chapter_boundaries(page_texts)
    notice_pages_physical = {
        p for p, text in enumerate(page_texts, start=1) if _is_notice_page(text)
    }

    displayed_map = build_displayed_page_map(sections, section_pages, total_pages)
    footer_numbers = footer_numbers or {}

    # --- sumber kebenaran TAMBAHAN (prioritas TERTINGGI): restart nomor
    #     halaman EKSPLISIT (w:pgNumType/@start) di section BERIKUTNYA
    #     setelah suatu halaman.
    #
    # KENAPA PERLU: to_displayed() di bawah memprioritaskan footer_numbers
    # (hasil BACA render PDF) di atas rumus w:pgNumType. Tapi footer hasil
    # render itu sendiri BISA BERBEDA dari yang mahasiswa lihat di Word
    # kalau LibreOffice mem-paginasi konten SEDIKIT beda dari Word --
    # BUKAN cuma soal font substitusi (sudah ditangani _verify_font_
    # available), tapi juga bisa dari perbedaan algoritma line-breaking,
    # tinggi baris tabel, dst antara dua engine yang berbeda. Untuk
    # paragraf/tabel yang KEBETULAN pas di batas halaman, beda 1 baris
    # kecil saja sudah cukup menggeser seluruh halaman berikutnya.
    #
    # SEBALIKNYA, w:pgNumType/@start di section BERIKUTNYA adalah properti
    # INTRINSIK dokumen -- ditulis mahasiswa lewat Word (Format Halaman >
    # Page Number Format > Start at), TIDAK bergantung sama sekali pada
    # engine render mana pun yang dipakai. Kalau section berikutnya
    # memang restart eksplisit ke N, dan TIDAK ADA halaman notice
    # menyisip di antaranya, maka "halaman terakhir section SEBELUM itu"
    # PASTI bernomor tercetak (N - 1) -- ini fakta struktural, bukan
    # hasil kalkulasi yang bisa meleset gara-gara render.
    #
    # Dipakai HANYA untuk menimpa hasil to_displayed() pada titik transisi
    # section (physical_page + 1 masuk section BARU yang py restart
    # eksplisit) -- untuk halaman lain (bukan tepat di batas section),
    # to_displayed() biasa (footer ground truth -> fallback rumus) tetap
    # dipakai apa adanya, karena override ini cuma valid tepat di titik
    # itu.
    sections_by_index = {sec["index"]: sec for sec in sections}
    sorted_section_starts = sorted(
        (
            (sec_idx, phys_start)
            for sec_idx, phys_start in section_pages.items()
            if sec_idx in sections_by_index
        ),
        key=lambda pair: pair[1],
    )

    def _section_governing(physical_page: int):
        """Section yang menaungi suatu halaman fisik tertentu, ditelusuri
        dari section_pages (peta section_index -> halaman fisik AWAL
        section itu). None kalau physical_page di luar jangkauan section
        manapun yang berhasil dipetakan (mis. snippet section itu gagal
        ketemu di render)."""
        governing = None
        for sec_idx, phys in sorted_section_starts:
            if phys <= physical_page:
                governing = sections_by_index[sec_idx]
            else:
                break
        return governing

    def _authoritative_content_end_display(content_end_physical: int) -> int | None:
        """Override content_end_display dari restart eksplisit section
        berikutnya, KALAU physical_page+1 memang masuk section baru yang
        py restart. None kalau tidak berlaku (bukan titik transisi
        section, atau section berikutnya tidak restart eksplisit) --
        caller fallback ke to_displayed() biasa."""
        cur_sec = _section_governing(content_end_physical)
        next_sec = _section_governing(content_end_physical + 1)
        if next_sec is None or next_sec.get("pgnum_start") is None:
            return None
        if cur_sec is not None and cur_sec["index"] == next_sec["index"]:
            return None  # bukan titik transisi section, override gak relevan
        return next_sec["pgnum_start"] - 1

    def to_displayed(physical: int) -> int:
        # Prioritas 1: angka footer yang benar2 kebaca di render (ground
        # truth, sudah divalidasi urutannya di extract_footer_numbers).
        # Prioritas 2 (fallback): hasil rumus w:pgNumType.
        # Prioritas 3 (fallback terakhir): posisi fisik apa adanya,
        # supaya tidak crash / kehilangan finding lain kalau dua-duanya
        # gagal (mis. halaman yang section-nya juga gak ke-cover peta).
        if physical in footer_numbers:
            return footer_numbers[physical]
        return displayed_map.get(physical, physical)

    # seen_bab: sama persis konvensi grouping di extractor.py --
    # FRONT_HEADINGS SEBELUM BAB 1 pertama digabung ke "Pra-Bab" (biar
    # tidak numpuk jadi grup sendiri-sendiri kecil di summary), sedang
    # FRONT_HEADINGS SETELAH BAB N (mis. DAFTAR PUSTAKA, DAFTAR
    # LAMPIRAN) dapat grup sendiri sesuai labelnya -- itu memang sudah
    # jadi konvensi grup existing dari extractor.py, jadi disamakan di
    # sini supaya findings dari kedua modul kekelompokkan konsisten,
    # bukan pecah jadi grup yang beda konvensi.
    seen_bab = False
    # first_checked_boundary: menggantikan peran "i == 0" lama (lihat cek
    # 1 di bawah) sekarang bahwa boundary Pra-Bab dilewati sepenuhnya --
    # tanpa ini, cek "mulai di halaman ganjil" tidak akan pernah jalan
    # sama sekali untuk BAB 1 kalau dokumen punya Pra-Bab (karena BAB 1
    # bukan lagi i == 0 di `boundaries`, tapi DIA-lah boundary pertama
    # yang sungguh-sungguh dicek/di-skor sekarang).
    first_checked_boundary = True

    for i, b in enumerate(boundaries):
        label, start_page_physical = b["label"], b["start_page"]

        if CHAPTER_TITLE_RE.match(label):
            seen_bab = True
            chapter_key = label
        elif not seen_bab:
            chapter_key = "Pra-Bab"
        else:
            chapter_key = label

        # DIHAPUS (atas permintaan user eksplisit): aturan ganjil/genap
        # DAN aturan halaman "sengaja dikosongkan" sengaja dikosongkan
        # untuk Pra-Bab (dari setelah cover sampai sebelum BAB 1) --
        # hitungannya sekarang mulai dari BAB 1 saja, berlaku sama untuk
        # docx maupun pdf (fungsi ini dipakai bersama oleh keduanya lewat
        # render LibreOffice). Boundary Pra-Bab dilewati total: tidak
        # checked, tidak failed, tidak ada finding sama sekali.
        if chapter_key == "Pra-Bab":
            continue

        # DIKECUALIKAN TOTAL (atas permintaan user, 2026-09-17): DAFTAR
        # LAMPIRAN dan BIODATA PENULIS -- sama seperti Pra-Bab di atas,
        # tidak checked/failed/finding sama sekali, berlaku sama untuk
        # docx maupun pdf (fungsi ini dipakai bersama oleh keduanya).
        # LAMPIRAN <huruf> (mis. "LAMPIRAN A") sendiri MEMANG tidak
        # pernah jadi boundary di sini sejak awal -- modul ini cuma
        # mengenali CHAPTER_TITLE_RE ("BAB N") & FRONT_HEADINGS
        # (lihat _is_chapter_boundary_line()), sedang LAMPIRAN_TITLE_RE
        # sengaja tidak diimpor ke modul ini -- jadi entri "DAFTAR
        # LAMPIRAN" (heading daftar isi lampiran di bagian akhir) dan
        # "BIODATA PENULIS" ini SATU-SATUNYA titik terkait
        # lampiran/biodata yang perlu dikecualikan eksplisit di modul
        # page-layout ini.
        if chapter_key in ("DAFTAR LAMPIRAN", "BIODATA PENULIS"):
            continue

        start_page_display = to_displayed(start_page_physical)

        # --- cek 1: boundary INI mulai di halaman (TERCETAK) ganjil ---
        # KOREKSI (bug dobel-hitung sebelumnya): cek ini HANYA dijalankan
        # untuk boundary CHECKED PERTAMA di dokumen (dulu "i == 0", kini
        # first_checked_boundary karena Pra-Bab di-skip di atas). Untuk
        # boundary ke-2 dst, "mulai di halaman genap" itu SELALU 100%
        # ekuivalen dengan "boundary SEBELUMNYA berakhir ganjil tanpa
        # notice" -- dua pertanyaan yang menguji FAKTA FISIK YANG SAMA
        # (ada/tidaknya celah halaman kosong di antara keduanya), cuma
        # dilihat dari dua sisi. Kalau tetap dicek dua-duanya, satu
        # halaman kosong yang hilang dilaporkan sebagai DUA finding
        # terpisah (skor "dihukum dua kali" untuk satu kesalahan yang
        # sama) -- itu sudah tercakup lengkap oleh cek 2 (bagian bawah)
        # milik boundary sebelumnya, jadi di sini tidak perlu diulang.
        if first_checked_boundary:
            checks[chapter_key]["checked"] += 1
            if start_page_display % 2 == 0:
                checks[chapter_key]["failed"] += 1
                _add_finding(
                    chapter_key,
                    f"Halaman {start_page_display}",
                    (
                        f"{label} dimulai di halaman {start_page_display} (genap), "
                        f"seharusnya di halaman ganjil."
                    ),
                )
        first_checked_boundary = False

        # --- cek 2: di mana konten ASLI bab ini berakhir? ---
        # next_start = halaman FISIK mulai bab/boundary berikutnya (atau
        # akhir dokumen kalau ini bab terakhir). Existence check tetap
        # di ruang fisik.
        next_start_physical = (
            boundaries[i + 1]["start_page"]
            if i + 1 < len(boundaries)
            else total_pages + 1
        )
        boundary_end_physical = next_start_physical - 1

        # Kalau halaman tepat sebelum boundary berikutnya SUDAH berupa
        # halaman notice, halaman itu BUKAN konten asli bab ini -- jadi
        # konten aslinya berakhir satu halaman sebelum itu.
        content_end_physical = (
            boundary_end_physical - 1
            if boundary_end_physical in notice_pages_physical
            else boundary_end_physical
        )
        content_end_display = _authoritative_content_end_display(content_end_physical)
        if content_end_display is None:
            content_end_display = to_displayed(content_end_physical)

        checks[chapter_key]["checked"] += 1
        if content_end_display % 2 == 1:
            # Bab berakhir ganjil (tercetak) -> WAJIB ada halaman notice
            # tepat setelahnya.
            expected_notice_physical = content_end_physical + 1
            if (
                expected_notice_physical > total_pages
                or expected_notice_physical not in notice_pages_physical
            ):
                checks[chapter_key]["failed"] += 1
                expected_notice_display = content_end_display + 1
                _add_finding(
                    chapter_key,
                    f"Halaman {expected_notice_display}",
                    (
                        f"{label} berakhir di halaman ganjil ({content_end_display}), "
                        f"halaman {expected_notice_display} seharusnya dikosongkan "
                        f"dengan tulisan 'Halaman ini sengaja dikosongkan' di "
                        f"tengah halaman."
                    ),
                )
        else:
            # Bab berakhir genap (tercetak) -> TIDAK boleh ada halaman
            # notice yang disisipkan (karena memang tidak diperlukan).
            if boundary_end_physical in notice_pages_physical:
                checks[chapter_key]["failed"] += 1
                boundary_end_display = to_displayed(boundary_end_physical)
                _add_finding(
                    chapter_key,
                    f"Halaman {boundary_end_display}",
                    (
                        f"{label} sudah berakhir di halaman genap ({content_end_display}), "
                        f"halaman kosong bertulisan 'Halaman ini sengaja "
                        f"dikosongkan' di halaman {boundary_end_display} tidak diperlukan."
                    ),
                )

    return findings, dict(checks)
