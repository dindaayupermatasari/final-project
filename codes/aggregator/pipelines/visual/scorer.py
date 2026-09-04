"""
scorer.py
=========
TAHAP 4: CALCULATE SCORE
-------------------------
Skor dihitung PER BAB, berbasis rasio parameter yang lulus terhadap
total parameter yang diperiksa pada bab tersebut:

    skor_bab = (checked - failed) / checked * 100

Skor keseluruhan dokumen = rata-rata skor semua bab (simple average),
konsisten dengan skema "Compute Overall Score" (Option A: skor per
pipeline/bagian dihitung dulu, baru dirata-rata).

CHANGELOG (revisi):
- Formula lama "100 - penalty(MAJOR*5 + MINOR*2)" dibuang. Formula itu
  gampang mentok ke 0 begitu jumlah temuan banyak, karena penalti tidak
  dibatasi relatif terhadap jumlah parameter yang diperiksa.
- severity (MAJOR/MINOR) dihapus dari struktur, skor sekarang murni
  dari rasio checked vs failed.
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

    valid_scores = [v["score"] for v in per_chapter.values() if v["checked"] > 0]
    overall_score = (
        round(sum(valid_scores) / len(valid_scores), 2) if valid_scores else 0.0
    )

    return {
        "total_findings": sum(v["failed"] for v in per_chapter.values()),
        "overall_score": overall_score,
        "per_chapter": per_chapter,
    }
