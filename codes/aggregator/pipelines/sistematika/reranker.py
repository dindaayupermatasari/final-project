"""
reranker.py — pipelines/sistematika/

Tahap "Cross Encoder Re-ranking" di diagram sistem: menerima kandidat Top-K
dari dense retriever, menilai ulang tiap pasangan (query, chunk panduan)
secara langsung, lalu mengurutkan ulang.

Model default multilingual (teks Indonesia). Cross-encoder yang hanya
bahasa Inggris (mis. ms-marco) akan lemah di sini.

  pip install sentence-transformers --break-system-packages
  # pertama kali jalan: model diunduh dari HuggingFace (bge-reranker-v2-m3 ~2 GB)

Alternatif lebih ringan (multilingual, ~120 MB):
  cross-encoder/mmarco-mMiniLMv2-L12-H384-v1
"""

from typing import List, Optional

from vector_store import build_embed_text

DEFAULT_MODEL = "BAAI/bge-reranker-v2-m3"


class CrossEncoderReranker:
    def __init__(
        self,
        model_name: str = DEFAULT_MODEL,
        max_length: int = 1024,
        device: Optional[str] = None,
    ):
        try:
            from sentence_transformers import CrossEncoder
        except ImportError as e:
            raise RuntimeError(
                "Butuh sentence-transformers: "
                "pip install sentence-transformers --break-system-packages"
            ) from e
        self.model_name = model_name
        self.model = CrossEncoder(model_name, max_length=max_length, device=device)

    def score(
        self, query: str, passages: List[str], batch_size: int = 8
    ) -> List[float]:
        if not passages:
            return []
        pairs = [(query, p) for p in passages]
        scores = self.model.predict(
            pairs, batch_size=batch_size, show_progress_bar=False
        )
        return [float(s) for s in scores]


def candidate_passage(match: dict, use_header: bool = True) -> str:
    """Teks kandidat yang dinilai reranker.

    use_header=True memakai header yang sama dengan yang di-embed di indeks
    (judul bab/sub-bab + isi). use_header=False hanya isi chunk -- ujian yang
    lebih jujur untuk mode content, karena judul panduan tidak ikut membantu.
    """
    text = match["matched_text"]
    if not use_header:
        return text
    return build_embed_text(text, match)


def rerank_matches(
    query: str,
    matches: List[dict],
    reranker: CrossEncoderReranker,
    use_header: bool = True,
) -> List[dict]:
    """Urutkan ulang kandidat; simpan peringkat lama di `retriever_rank`."""
    passages = [candidate_passage(m, use_header) for m in matches]
    scores = reranker.score(query, passages)
    out = [
        {**m, "retriever_rank": m["rank"], "rerank_score": round(s, 4)}
        for m, s in zip(matches, scores)
    ]
    out.sort(key=lambda m: m["rerank_score"], reverse=True)
    for i, m in enumerate(out, start=1):
        m["rank"] = i
    return out
