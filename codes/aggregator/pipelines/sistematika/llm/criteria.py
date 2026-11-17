"""Checklist kriteria per subbab Bab 1, diturunkan dari panduan (production_snapshot.json)
dan slide Pak Ali. Teks kriteria boleh memakai placeholder {problem_domain}, {problem},
{uniqueness} yang diisi dari hasil title_decomposer.

PENTING: kunci memakai JUDUL subbab, bukan nomor. Nomor di panduan dan di dokumen
mahasiswa bisa beda (panduan: 1.3 Batasan Masalah, 1.4 Tujuan; dokumen lama: 1.3 Tujuan).
`guideline_id` adalah id chunk di ChromaDB/snapshot (lengkap dengan akhiran _instruction).
`butuh_bukti`: default True; False untuk kriteria bernada "tidak memuat..." yang buktinya
berupa ketiadaan (tidak ada kutipan yang bisa diminta).
`wajib`: True = tegas di panduan, False = anjuran/kualitas ("sebaiknya", "yang baik").
"""

CRITERIA = {
    "LATAR BELAKANG": {
        "guideline_id": "bab1_sub1-1_instruction",
        "kriteria": [
            {"id": "L1", "wajib": True,
             "deskripsi": "Membahas problem domain \"{problem_domain}\" (bidang/objeknya), "
                          "bukan topik lain."},
            {"id": "L2", "wajib": True,
             "deskripsi": "Menunjukkan tingkat urgensi problem domain tersebut, mis. tren/"
                          "kenaikan jumlah, dampak, atau tingkat keparahannya."},
            {"id": "L3", "wajib": False, "butuh_bukti": False,
             "deskripsi": "Fokus pada domain, tidak didominasi uraian solusi \"{uniqueness}\"."},
        ],
    },
    "PERMASALAHAN": {
        "guideline_id": "bab1_sub1-2_instruction",
        "kriteria": [
            {"id": "P1", "wajib": True,
             "deskripsi": "Mendeskripsikan problem \"{problem}\" dengan jelas dan deskriptif."},
            {"id": "P2", "wajib": True,
             "deskripsi": "Menguraikan faktor-faktor yang membuat problem tersebut sulit "
                          "atau belum terselesaikan."},
            {"id": "P3", "wajib": False,
             "deskripsi": "Meyakinkan pembaca bahwa problem ini mendesak dan butuh solusi "
                          "dari penelitian ini."},
        ],
    },
    "TUJUAN": {
        "guideline_id": "bab1_sub1-4_instruction",
        "kriteria": [
            {"id": "T1", "wajib": True,
             "deskripsi": "Menyatakan dengan jelas apa yang dilakukan/dibangun untuk "
                          "problem \"{problem}\"."},
            {"id": "T2", "wajib": False, "butuh_bukti": False,
             "deskripsi": "Tidak memuat uraian latar belakang atau pengulangan "
                          "permasalahan (itu tempatnya di 1.1 dan 1.2)."},
            {"id": "T3", "wajib": True,
             "deskripsi": "Mengandung klaim orisinalitas: menyebut \"{uniqueness}\" sebagai "
                          "pendekatan/algoritma/pemodelan/teknik/metode yang baru atau unik."},
            {"id": "T4", "wajib": False,
             "deskripsi": "Klaim orisinalitas dihubungkan dengan problem \"{problem}\": "
                          "ada kalimat yang menyatakan pendekatan/metode baru UNTUK problem "
                          "tersebut DENGAN memakai solusinya. Kata persisnya boleh berbeda "
                          "dari pola panduan (\"untuk mengatasi ... dengan memakai ...\")."},
            {"id": "T5", "wajib": False,
             "deskripsi": "Kalimat berikutnya menerangkan secara argumentatif fitur unik "
                          "\"{uniqueness}\" sebagai solusi."},
        ],
    },
}

# Judul subbab yang berbeda di dokumen mahasiswa tapi padanannya sama di panduan.
ALIAS = {
    "LATAR BELAKANG MASALAH": "LATAR BELAKANG",
    "RUMUSAN MASALAH": "PERMASALAHAN",
    "TUJUAN PENELITIAN": "TUJUAN",
    "TUJUAN PROYEK AKHIR": "TUJUAN",
}


def get_criteria(subbab_title: str) -> dict | None:
    t = subbab_title.strip().upper()
    return CRITERIA.get(ALIAS.get(t, t))
