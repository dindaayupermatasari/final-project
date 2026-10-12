"""
scorer.py
=========
TAHAP 4: CALCULATE SCORE
-------------------------
Skor dihitung PER BAB, berbasis rasio parameter yang lulus terhadap
total parameter yang diperiksa pada bab tersebut:

    skor_bab = (checked - failed) / checked * 100

Skor keseluruhan dokumen = rasio agregat (weighted), BUKAN rata-rata
sederhana dari skor semua bab:

    overall_score = (total_checked - total_failed) / total_checked * 100

di mana total_checked dan total_failed dijumlahkan dari SEMUA bab
terlebih dahulu, baru dihitung rasionya sekali di akhir. Ini membuat
bab dengan parameter lebih banyak (mis. BAB 3 dengan 300+ parameter)
berpengaruh lebih besar terhadap skor akhir dibanding bab kecil
(mis. "Umum (Margin Halaman)" dengan hanya 3-6 parameter).

CHANGELOG (revisi):
- Formula lama "100 - penalty(MAJOR*5 + MINOR*2)" dibuang. Formula itu
  gampang mentok ke 0 begitu jumlah temuan banyak, karena penalti tidak
  dibatasi relatif terhadap jumlah parameter yang diperiksa.
- severity (MAJOR/MINOR) dihapus dari struktur, skor sekarang murni
  dari rasio checked vs failed.
- overall_score diubah dari rata-rata sederhana skor per bab (Option A)
  menjadi rasio agregat/weighted di atas total checked-failed semua
  bab. Rata-rata sederhana membuat bab kecil (sample size rendah)
  punya bobot yang sama besar dengan bab besar, sehingga dokumen
  dengan total temuan lebih SEDIKIT bisa mendapat overall_score lebih
  RENDAH hanya karena kebetulan gagal di satu bab kecil. skor per bab
  (per_chapter) tetap dihitung dengan formula lama dan tetap
  ditampilkan apa adanya untuk keperluan diagnostik.
"""


def calculate_score(checks: dict) -> dict:
    per_chapter = {}
    for chapter, c in checks.items():
        checked = c["checked"]
        failed = c["failed"]
        score = round((checked - failed) / checked * 100, 2) if checked else 0.0
        per_chapter[chapter] = {
            "checked": checked,
            "failed": failed,
            "score": score,
        }

    total_checked = sum(v["checked"] for v in per_chapter.values())
    total_failed = sum(v["failed"] for v in per_chapter.values())
    overall_score = (
        round((total_checked - total_failed) / total_checked * 100, 2)
        if total_checked
        else 0.0
    )

    return {
        "total_findings": total_failed,
        "overall_score": overall_score,
        "per_chapter": per_chapter,
    }
