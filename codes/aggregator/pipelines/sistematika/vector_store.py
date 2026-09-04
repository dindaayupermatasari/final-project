"""
vector_store.py — pipelines/sistematika/

Modul BERSAMA untuk koneksi ChromaDB dan pemanggilan embedding model.
Ini yang paling penting untuk dipisah sendiri: baik proses ingest (sekali
di awal / saat panduan direvisi) MAUPUN proses retrieval saat runtime
(tiap mahasiswa upload dokumen) sama-sama butuh fungsi embed_text() dan
koneksi ke Chroma collection yang SAMA. Kalau logic ini nyangkut di script
ingest, retriever.py nanti harus import script sekali-jalan itu -- rapuh.
"""

from typing import List

import chromadb
import requests

OLLAMA_URL = "http://localhost:11434/api/embeddings"
EMBED_MODEL = "nomic-embed-text"


def embed_text(text: str) -> List[float]:
    """Panggil Ollama untuk mengubah teks jadi vektor embedding."""
    response = requests.post(
        OLLAMA_URL,
        json={"model": EMBED_MODEL, "prompt": text},
        timeout=60,
    )
    response.raise_for_status()
    return response.json()["embedding"]


def get_collection(chroma_dir: str, collection_name: str):
    """
    Buka (atau buat) koleksi ChromaDB yang persist ke disk di `chroma_dir`.
    Dipanggil oleh ingest_guideline.py (untuk nulis) dan nanti retriever.py
    (untuk baca/query) -- keduanya pakai chroma_dir & collection_name yang
    sama supaya nyambung ke data yang sama.
    """
    client = chromadb.PersistentClient(path=chroma_dir)
    return client.get_or_create_collection(name=collection_name)


def upsert_chunks(chunks: List[dict], chroma_dir: str, collection_name: str) -> int:
    """Embed tiap chunk lalu simpan (upsert) ke koleksi ChromaDB. Return jumlah chunk yang masuk."""
    collection = get_collection(chroma_dir, collection_name)

    ids, documents, metadatas, embeddings = [], [], [], []
    for chunk in chunks:
        ids.append(chunk["id"])
        documents.append(chunk["text"])
        metadatas.append(chunk["metadata"])
        embeddings.append(embed_text(chunk["text"]))

    collection.upsert(
        ids=ids, documents=documents, metadatas=metadatas, embeddings=embeddings
    )
    return len(ids)


def query_similar(
    query_text: str, chroma_dir: str, collection_name: str, top_k: int = 5
) -> dict:
    """
    Cari top-K chunk panduan paling mirip dengan query_text (misal: isi
    Bab 1 dokumen mahasiswa yang sedang dievaluasi). Ini yang akan dipakai
    retriever.py nanti saat runtime.
    """
    collection = get_collection(chroma_dir, collection_name)
    query_vector = embed_text(query_text)
    return collection.query(query_embeddings=[query_vector], n_results=top_k)
