"""Klien LLM tunggal untuk seluruh pipeline penilaian.

Pakai format OpenAI-compatible, jadi penyedia bisa diganti lewat .env
(Alibaba Model Studio, OpenRouter, Ollama lokal, dll) tanpa ubah kode.

.env (jangan di-commit, masukkan ke .gitignore):
    LLM_API_KEY=...
    LLM_BASE_URL=https://...        # isi sesuai penyedia
    LLM_MODEL=nama-model-qwen       # cek nama persisnya di dokumentasi penyedia
Contoh Groq:
    LLM_BASE_URL=https://api.groq.com/openai/v1
    LLM_MODEL=qwen/qwen3-32b        # cek daftar terbaru: GET {base_url}/models
Untuk Ollama lokal: LLM_BASE_URL=http://localhost:11434/v1, LLM_API_KEY=ollama
"""
import hashlib
import json
import os
import re
import time
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI, RateLimitError

load_dotenv()

_client = None


class DailyLimitError(RuntimeError):
    """Kuota token harian penyedia habis. Menunggu beberapa detik tidak membantu, jadi tidak diulang."""


# Cache jawaban LLM di disk (satu baris JSON per panggilan). Kunci = model + temperature + system + user.
# Karena temperature 0 dan seed tetap, prompt yang sama memberi jawaban yang sama, jadi menjalankan ulang
# dokumen yang sama (mis. setelah kena batas harian) tidak memakai token lagi untuk bagian yang sudah selesai.
# Matikan dengan LLM_CACHE=0. Lokasi bisa diganti lewat LLM_CACHE_FILE.
_CACHE_FILE = Path(os.environ.get("LLM_CACHE_FILE") or Path(__file__).resolve().parent.parent / "debug_output" / "llm_cache.jsonl")
_cache = None


def _cache_on() -> bool:
    return os.environ.get("LLM_CACHE", "1") != "0"


def _cache_key(system: str, user: str, temperature: float) -> str:
    raw = "\x00".join([os.environ.get("LLM_MODEL", ""), str(temperature), system, user])
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _cache_get(key: str):
    global _cache
    if _cache is None:
        _cache = {}
        if _CACHE_FILE.exists():
            for line in _CACHE_FILE.read_text(encoding="utf-8").splitlines():
                try:
                    row = json.loads(line)
                    _cache[row["k"]] = row["raw"]  # baris yang lebih baru menimpa yang lama
                except (ValueError, KeyError):
                    continue
    return _cache.get(key)


def _cache_put(key: str, raw: str) -> None:
    _cache_get(key)  # pastikan _cache sudah dimuat
    _cache[key] = raw
    _CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(_CACHE_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps({"k": key, "raw": raw}, ensure_ascii=False) + "\n")


def _get_client() -> OpenAI:
    global _client
    if _client is None:
        _client = OpenAI(
            api_key=os.environ["LLM_API_KEY"],
            base_url=os.environ["LLM_BASE_URL"],
        )
    return _client


def call_llm(system: str, user: str, temperature: float = 0.0, max_retries: int = 4, fresh: bool = False) -> str:
    """Panggil LLM; kalau kena rate limit (429), tunggu lalu coba lagi. Batas HARIAN langsung DailyLimitError.
    fresh=True melewati cache (dipakai saat mengulang karena jawaban sebelumnya bukan JSON valid)."""
    key = _cache_key(system, user, temperature)
    if _cache_on() and not fresh:
        hit = _cache_get(key)
        if hit is not None:
            return hit
    for attempt in range(max_retries + 1):
        try:
            resp = _get_client().chat.completions.create(
                model=os.environ["LLM_MODEL"],
                temperature=temperature,
                seed=42,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
            )
            raw = resp.choices[0].message.content
            if _cache_on():
                _cache_put(key, raw)
            return raw
        except RateLimitError as e:
            if "per day" in str(e).lower() or "(tpd)" in str(e).lower():
                raise DailyLimitError(str(e)) from e
            if attempt == max_retries:
                raise
            time.sleep(5 * (attempt + 1))


def parse_json(text: str) -> dict:
    """Ambil objek JSON dari jawaban model, termasuk kalau dibungkus ```json."""
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)  # Qwen3 thinking
    text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"Tidak ada JSON di jawaban model: {text[:200]}")
    return json.loads(text[start : end + 1])


def call_llm_json(system: str, user: str, retries: int = 1) -> dict:
    """call_llm + parse JSON, ulangi sekali kalau JSON-nya rusak."""
    last_err = None
    for attempt in range(retries + 1):
        raw = call_llm(system, user, fresh=attempt > 0)
        try:
            return parse_json(raw)
        except (ValueError, json.JSONDecodeError) as e:
            last_err = e
    raise last_err


if __name__ == "__main__":
    # Tes koneksi: harus mencetak JSON {"ok": true}
    print(call_llm_json("Balas hanya dengan JSON.", 'Balas {"ok": true}'))
