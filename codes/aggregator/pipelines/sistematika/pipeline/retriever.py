"""Tahap RAG: ambil kandidat chunk panduan untuk satu subbab dokumen mahasiswa.

Dua mode, keluarannya sama:
  live   : embed query lewat Ollama (nomic-embed-text), cari di ChromaDB.
  cached : pakai kandidat yang sudah tersimpan di retrieval_test_from_student_doc.json
           (cadangan demo kalau Ollama/ChromaDB bermasalah).

Bentuk satu kandidat:
  {rank, chunk_id, subbab_num, subbab_title, distance, text, example}
`text` = chunk instruction, `example` = chunk contoh pasangannya (boleh kosong).
Query yang dipakai = judul subbab saja (mode "label" di tesmu: recall@3 = 1.0).
"""
import difflib
import json
import re
import urllib.request

OLLAMA_EMBED_URL = "http://localhost:11434/api/embed"
QUERY_PREFIX = "search_query: "  # prefix nomic-embed-text untuk query


class SnapshotIndex:
    """Isi production_snapshot.json: header koleksi + teks tiap chunk."""

    def __init__(self, path):
        data = json.load(open(path, encoding="utf-8"))
        self.header = {k: v for k, v in data.items() if k != "items"}
        self.items = {i["id"]: i for i in data["items"]}

    def text(self, chunk_id: str) -> str:
        return self.items.get(chunk_id, {}).get("document_text", "")

    def example_of(self, chunk_id: str) -> str:
        return self.text(chunk_id.replace("_instruction", "_example"))

    def meta(self, chunk_id: str) -> dict:
        return self.items.get(chunk_id, {}).get("metadata", {})


def _embed(text: str, model: str) -> list[float]:
    body = json.dumps({"model": model, "input": [QUERY_PREFIX + text]}).encode("utf-8")
    req = urllib.request.Request(
        OLLAMA_EMBED_URL, data=body, headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req, timeout=60) as r:
        vec = json.load(r)["embeddings"][0]
    norm = sum(x * x for x in vec) ** 0.5 or 1.0
    return [x / norm for x in vec]  # vektor di koleksi sudah dinormalisasi (norm ~ 1)


class LiveRetriever:
    kind = "live"

    def __init__(self, chroma_dir: str, collection: str, embed_model: str):
        import chromadb  # impor di sini supaya mode cached tidak butuh chromadb

        self.col = chromadb.PersistentClient(path=chroma_dir).get_collection(collection)
        self.embed_model = embed_model

    def retrieve(self, query: str, k: int = 3, label_hint: str | None = None) -> list[dict]:
        vec = _embed(query, self.embed_model)
        r = self.col.query(
            query_embeddings=[vec],
            n_results=k,
            where={"chunk_type": "instruction"},
            include=["documents", "metadatas", "distances"],
        )
        out = []
        for i, (cid, doc, meta, dist) in enumerate(
            zip(r["ids"][0], r["documents"][0], r["metadatas"][0], r["distances"][0]), 1
        ):
            ex = self.col.get(ids=[cid.replace("_instruction", "_example")])
            out.append(
                {
                    "rank": i,
                    "chunk_id": cid,
                    "subbab_num": meta.get("subbab_num"),
                    "subbab_title": meta.get("subbab_title"),
                    "distance": round(float(dist), 4),
                    "text": doc,
                    "example": ex["documents"][0] if ex.get("ids") else "",
                }
            )
        return out


def _judul(label: str) -> str:
    """'BAB 4 > 4.2 KARATERISTIK DATA (dari dokumen mahasiswa)' -> 'KARATERISTIK DATA'."""
    t = re.sub(r"\(dari dokumen mahasiswa\)\s*$", "", label or "").split(">")[-1]
    return " ".join(re.sub(r"^\s*\d+(?:\.\d+)*\s*", "", t).split()).upper()


class CachedRetriever:
    """Kandidat dari file cache; kalau subbab tidak ada di cache (dokumen lain, nomor/judul beda), dicari
    lewat judul: (1) entri cache dengan judul yang sama, (2) peringkat kemiripan judul terhadap chunk panduan
    (difflib, tanpa embedding). Hasil (2) bukan retrieval vektor: field `sumber` = 'fallback_judul', dan
    `distance` = 1 - kemiripan judul."""
    kind = "cached"

    def __init__(self, retrieval_json_path: str, snapshot: SnapshotIndex):
        data = json.load(open(retrieval_json_path, encoding="utf-8"))
        self.results = data["modes"]["label"]["results"]
        self.snap = snapshot

    def _dari_cache(self, hit: dict, k: int, sumber: str) -> list[dict]:
        out = []
        for c in hit["candidates"][:k]:
            cid = c["chunk_id"]
            out.append(
                {
                    "rank": c["rank"],
                    "chunk_id": cid,
                    "subbab_num": c.get("subbab_num"),
                    "subbab_title": c.get("subbab_title"),
                    "distance": c.get("distance"),
                    "text": self.snap.text(cid),
                    "example": self.snap.example_of(cid),
                    "sumber": sumber,
                }
            )
        return out

    def _fallback_judul(self, query: str, k: int) -> list[dict]:
        q = " ".join(query.split()).upper()
        rows = []
        for cid, it in self.snap.items.items():
            if not cid.endswith("_instruction"):
                continue
            m = it.get("metadata", {})
            t = " ".join(str(m.get("subbab_title", "")).split()).upper()
            rows.append((difflib.SequenceMatcher(None, q, t).ratio(), cid, m))
        rows.sort(key=lambda r: -r[0])
        return [
            {"rank": i, "chunk_id": cid, "subbab_num": m.get("subbab_num"), "subbab_title": m.get("subbab_title"),
             "distance": round(1 - r, 4), "text": self.snap.text(cid), "example": self.snap.example_of(cid),
             "sumber": "fallback_judul"}
            for i, (r, cid, m) in enumerate(rows[:k], 1)
        ]

    def retrieve(self, query: str, k: int = 3, label_hint: str | None = None) -> list[dict]:
        hit = next((r for r in self.results if label_hint and r["query_label"].startswith(label_hint)), None)
        if hit is not None:
            return self._dari_cache(hit, k, "cache")
        q = " ".join(query.split()).upper()
        hit = next((r for r in self.results if _judul(r["query_label"]) == q), None)
        if hit is not None:
            return self._dari_cache(hit, k, "cache_judul_sama")
        return self._fallback_judul(query, k)


def make_retriever(kind, snapshot: SnapshotIndex, retrieval_json=None, chroma_dir=None):
    if kind == "cached":
        return CachedRetriever(retrieval_json, snapshot)
    h = snapshot.header
    return LiveRetriever(chroma_dir or h["chroma_dir"], h["collection_name"], h["embed_model"])
