"""
vector_store.py — pipelines/sistematika/

Modul BERSAMA untuk koneksi ChromaDB dan pemanggilan embedding model
(dipakai ingest, retrieval runtime, dan script tes).

nomic-embed-text v1.5 lewat Ollama:
  1. PREFIX TASK: "search_document: " untuk teks terindeks, "search_query: "
     untuk pencarian. Prefix lain yang didukung nomic ("classification: ",
     "clustering: ") tersedia lewat parameter kind.
  2. NORMALISASI L2 di sisi sini (aman meski endpoint sudah menormalisasi).
  3. COSINE: koleksi dibuat hnsw:space=cosine. Ruang jarak tidak bisa
     diubah setelah dibuat -> koleksi lama harus di-reset.
  4. HEADER KONTEKSTUAL di teks yang di-embed (lihat build_embed_text),
     teks asli di `documents` tidak berubah.
  5. KOLEKSI DIKUNCI KE MODEL: model & metrik dicek saat dibuka.

Perubahan versi ini:
  * Endpoint /api/embed (batch) menggantikan /api/embeddings. Butuh
    Ollama versi baru.
  * upsert_chunks() meng-embed SEMUA chunk dulu, baru reset + tulis, jadi
    kalau Ollama gagal di tengah jalan koleksi lama tetap utuh. Chunk
    basi (id yang sudah tak ada di panduan baru) otomatis dipangkas.
  * get_collection(read_only=True) memakai client.get_collection sehingga
    salah ketik nama koleksi menghasilkan error, bukan koleksi kosong baru.
  * Pengecekan ruang jarak tahan terhadap versi Chroma yang tidak lagi
    mengembalikan "hnsw:space" di collection.metadata.
  * query_collection(): filter chunk_type (default "instruction") dan
    deduplikasi per judul sub-bab sebelum ranking.
  * normalize_title tidak lagi memakan angka di awal judul ("3D Modeling").
  * lookup_by_title punya fallback fuzzy.

Prasyarat: ollama pull nomic-embed-text:v1.5
"""

import difflib
import math
import re
from typing import List, Optional

import chromadb
import requests

OLLAMA_URL = "http://localhost:11434/api/embed"
EMBED_MODEL = "nomic-embed-text:v1.5"

PREFIXES = {
    "document": "search_document: ",
    "query": "search_query: ",
    "classification": "classification: ",
    "clustering": "clustering: ",
}
NUM_CTX = 4096  # nomic v1.5 mendukung sampai 8192 token
MAX_EMBED_CHARS = 6000
BATCH_SIZE = 16


def _l2_normalize(v: List[float]) -> List[float]:
    norm = math.sqrt(sum(x * x for x in v)) or 1.0
    return [x / norm for x in v]


def embed_texts(texts: List[str], kind: str = "document") -> List[List[float]]:
    """Embed banyak teks (batch) -> list vektor ternormalisasi."""
    if kind not in PREFIXES:
        raise ValueError(f"kind harus salah satu dari {list(PREFIXES)}, bukan '{kind}'")
    prefix = PREFIXES[kind]
    out: List[List[float]] = []
    for i in range(0, len(texts), BATCH_SIZE):
        batch = [prefix + t[:MAX_EMBED_CHARS] for t in texts[i : i + BATCH_SIZE]]
        response = requests.post(
            OLLAMA_URL,
            json={
                "model": EMBED_MODEL,
                "input": batch,
                "options": {"num_ctx": NUM_CTX},
            },
            timeout=120,
        )
        if response.status_code == 404:
            raise RuntimeError(
                f"Model '{EMBED_MODEL}' atau endpoint /api/embed tidak ditemukan di "
                f"Ollama. Jalankan: ollama pull {EMBED_MODEL} (dan update Ollama "
                f"kalau versinya lama)."
            )
        response.raise_for_status()
        embs = response.json()["embeddings"]
        if len(embs) != len(batch):
            raise RuntimeError(
                f"Ollama mengembalikan {len(embs)} vektor untuk {len(batch)} teks"
            )
        out.extend(_l2_normalize(e) for e in embs)
    return out


def embed_text(text: str, kind: str = "document") -> List[float]:
    """Embed satu teks (ternormalisasi)."""
    return embed_texts([text], kind)[0]


def build_embed_text(text: str, metadata: dict) -> str:
    """Teks yang benar-benar di-embed: header konteks + isi chunk."""
    bab = metadata.get("bab_title", "")
    num = metadata.get("subbab_num", "")
    title = metadata.get("subbab_title", "")
    label = (
        "Contoh penulisan"
        if metadata.get("chunk_type") == "example"
        else "Panduan penulisan"
    )
    header = f"{label} sub-bab {num} {title} pada BAB {bab}.".strip()
    return f"{header}\n{text}"


def _collection_space(collection) -> str:
    """Ruang jarak koleksi: dari metadata, lalu configuration; 'unknown' kalau tak terbaca."""
    meta = collection.metadata or {}
    if meta.get("hnsw:space"):
        return meta["hnsw:space"]
    try:
        cfg = getattr(collection, "configuration_json", None) or getattr(
            collection, "configuration", None
        )
        if isinstance(cfg, dict):
            return (cfg.get("hnsw") or {}).get("space") or "unknown"
    except Exception:
        pass
    return "unknown"


def get_collection(
    chroma_dir: str,
    collection_name: str,
    check_space: bool = True,
    read_only: bool = False,
):
    """
    Buka koleksi ChromaDB persisten di `chroma_dir`.

    read_only=True  -> client.get_collection: error kalau koleksi tidak ada
                       (dipakai retrieval/view/test).
    read_only=False -> get_or_create_collection (dipakai ingest).
    check_space=False -> jangan tolak koleksi lama (L2 / model lain).
    """
    client = chromadb.PersistentClient(path=chroma_dir)
    if read_only:
        collection = client.get_collection(name=collection_name)
    else:
        collection = client.get_or_create_collection(
            name=collection_name,
            metadata={"hnsw:space": "cosine", "embed_model": EMBED_MODEL},
        )
    if check_space:
        space = _collection_space(collection)
        model = (collection.metadata or {}).get("embed_model")
        if model != EMBED_MODEL or space not in ("cosine", "unknown"):
            raise RuntimeError(
                f"Koleksi '{collection_name}' dibuat dengan space={space}, "
                f"model={model}; kode ini butuh space=cosine, model={EMBED_MODEL}. "
                f"Ingest ulang: python ingest_guideline.py --reset"
            )
    return collection


def reset_collection(chroma_dir: str, collection_name: str) -> None:
    client = chromadb.PersistentClient(path=chroma_dir)
    try:
        client.delete_collection(collection_name)
    except Exception:
        pass  # belum ada, tidak masalah


def upsert_chunks(
    chunks: List[dict],
    chroma_dir: str,
    collection_name: str,
    reset: bool = False,
    prune: bool = True,
) -> int:
    """Embed semua chunk, baru (opsional) reset koleksi, lalu tulis.

    Embedding dilakukan PALING DULU: kalau gagal, koleksi lama tidak tersentuh.
    prune=True menghapus chunk lama yang id-nya tidak ada di `chunks`.
    """
    ids = [c["id"] for c in chunks]
    documents = [c["text"] for c in chunks]
    metadatas = [{**c["metadata"], "embed_model": EMBED_MODEL} for c in chunks]
    embeddings = embed_texts(
        [build_embed_text(c["text"], c["metadata"]) for c in chunks], kind="document"
    )

    if reset:
        reset_collection(chroma_dir, collection_name)
    collection = get_collection(chroma_dir, collection_name)
    collection.upsert(
        ids=ids, documents=documents, metadatas=metadatas, embeddings=embeddings
    )

    if prune:
        keep = set(ids)
        stale = [i for i in collection.get(include=[])["ids"] if i not in keep]
        if stale:
            collection.delete(ids=stale)
    return len(ids)


def query_similar(
    query_text: str,
    chroma_dir: str,
    collection_name: str,
    top_k: int = 5,
    kind: str = "query",
    where: Optional[dict] = None,
) -> dict:
    """Top-K chunk paling mirip (format mentah Chroma; distances = cosine distance)."""
    collection = get_collection(chroma_dir, collection_name, read_only=True)
    query_vector = embed_text(query_text, kind=kind)
    # jumlah chunk yang benar-benar memenuhi filter (bukan total koleksi),
    # supaya n_results tidak melebihi yang tersedia
    available = (
        len(collection.get(where=where, include=[])["ids"])
        if where
        else collection.count()
    )
    kwargs = {
        "query_embeddings": [query_vector],
        "n_results": max(1, min(top_k, available)),
    }
    if where:
        kwargs["where"] = where
    return collection.query(**kwargs)


def normalize_title(s: str) -> str:
    """Normalisasi judul: buang nomor di depan ("1.4 ", "2. "), tanda baca, kapital.
    Angka yang menempel pada kata ("3D Modeling") tidak dibuang."""
    s = re.sub(r"^\d+(?:\.\d+)+\.?\s+|^\d+\.\s+", "", s or "")
    s = re.sub(r"[^A-Za-z0-9 ]", " ", s)
    return " ".join(s.upper().split())


def query_collection(
    query_text: str,
    chroma_dir: str,
    collection_name: str,
    top_k: int = 5,
    kind: str = "query",
    chunk_type: Optional[str] = "instruction",
    dedupe: bool = True,
) -> List[dict]:
    """Seperti query_similar tapi list of dict per match.

    chunk_type: filter metadata (None = semua tipe).
    dedupe: satu sub-bab hanya muncul sekali (jarak terbaik), lalu dipotong top_k.
    """
    where = {"chunk_type": chunk_type} if chunk_type else None
    fetch_k = top_k * 4 if dedupe else top_k
    res = query_similar(query_text, chroma_dir, collection_name, fetch_k, kind, where)

    matches: List[dict] = []
    seen = set()
    for cid, doc, meta, dist in zip(
        res["ids"][0], res["documents"][0], res["metadatas"][0], res["distances"][0]
    ):
        key = normalize_title(meta.get("subbab_title", ""))
        if dedupe and key in seen:
            continue
        seen.add(key)
        matches.append(
            {
                "rank": len(matches) + 1,
                "chunk_id": cid,
                "bab_title": meta.get("bab_title"),
                "subbab_num": meta.get("subbab_num"),
                "subbab_title": meta.get("subbab_title"),
                "chunk_type": meta.get("chunk_type"),
                "distance": round(float(dist), 4),
                "matched_text": doc,
            }
        )
        if len(matches) == top_k:
            break
    return matches


def lookup_by_title(
    subbab_title: str,
    chroma_dir: str,
    collection_name: str,
    chunk_type: Optional[str] = None,
    fuzzy_cutoff: float = 0.88,
) -> List[dict]:
    """Ambil chunk panduan lewat judul sub-bab (tanpa similarity search).

    Exact match dulu, lalu fuzzy. Hasil kosong = tidak ada padanan -> pemanggil
    sebaiknya fallback ke query_collection().
    """
    collection = get_collection(
        chroma_dir, collection_name, check_space=False, read_only=True
    )
    kwargs = {"include": ["documents", "metadatas"]}
    if chunk_type:
        kwargs["where"] = {"chunk_type": chunk_type}
    stored = collection.get(**kwargs)

    rows = [
        (normalize_title(meta.get("subbab_title", "")), cid, meta, doc)
        for cid, doc, meta in zip(
            stored["ids"], stored["documents"], stored["metadatas"]
        )
    ]
    target = normalize_title(subbab_title)
    if target not in {r[0] for r in rows}:
        close = difflib.get_close_matches(
            target, list({r[0] for r in rows}), n=1, cutoff=fuzzy_cutoff
        )
        if not close:
            return []
        target = close[0]
    return [
        {"chunk_id": cid, "metadata": meta, "text": doc}
        for norm, cid, meta, doc in rows
        if norm == target
    ]
