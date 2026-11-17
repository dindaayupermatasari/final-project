import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import * as pdfjs from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import "../styles/pdfviewer.css";

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

const GAP = 14;

// Satu halaman: kanvas hanya digambar saat mendekati layar, dan dikosongkan lagi saat jauh (hemat memori untuk buku ratusan halaman).
function Halaman({ pdf, nomor, lebar, tinggi, root, sorot, aktif, onPilih }) {
  const kotak = useRef(null);
  const kanvas = useRef(null);
  const [dekat, setDekat] = useState(nomor <= 2);

  useEffect(() => {
    const io = new IntersectionObserver(([e]) => setDekat(e.isIntersecting), { root: root.current, rootMargin: "900px 0px" });
    io.observe(kotak.current);
    return () => io.disconnect();
  }, [root]);

  useEffect(() => {
    const c = kanvas.current;
    if (!dekat || !pdf) { c.width = c.height = 0; return; }
    let batal = false, tugas = null;
    (async () => {
      const page = await pdf.getPage(nomor);
      if (batal) return;
      const skala = lebar / page.getViewport({ scale: 1 }).width;
      const dpr = Math.min(window.devicePixelRatio || 1, 2);
      const vp = page.getViewport({ scale: skala * dpr });
      c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      tugas = page.render({ canvasContext: c.getContext("2d"), viewport: vp });
      try { await tugas.promise; } catch { /* dibatalkan saat ukuran berubah */ }
    })();
    return () => { batal = true; tugas?.cancel(); };
  }, [dekat, pdf, nomor, lebar]);

  return (
    <div ref={kotak} className="pdfp" data-hal={nomor} style={{ width: lebar, height: tinggi }}>
      <canvas ref={kanvas} style={{ width: lebar, height: tinggi }} />
      {sorot.map(({ k, jenis, r, pertama }, i) => (
        <button
          key={`${k}-${i}`} type="button" tabIndex={-1}
          className={`sr ${jenis}${k === aktif ? " on" : ""}`}
          style={{ left: `${r[0] * 100}%`, top: `${r[1] * 100}%`, width: `${(r[2] - r[0]) * 100}%`, height: `${(r[3] - r[1]) * 100}%` }}
          onClick={() => onPilih(k)} aria-label={`Temuan ${k + 1}`}
        >
          {pertama && <sup>{k + 1}</sup>}
        </button>
      ))}
    </div>
  );
}

/**
 * Pratinjau PDF yang bisa digulir dari halaman pertama sampai akhir, dengan sorotan temuan di atasnya.
 *  kunci    penanda dokumen; PDF dimuat ulang saat berubah
 *  muat     fungsi async yang mengembalikan ArrayBuffer PDF (memakai token login)
 *  ukuran   [[lebar, tinggi], ...] per halaman (dari /sorotan), supaya tinggi halaman sudah pasti sebelum digambar
 *  temuan   [{ k, jenis, hal, rects: [{p, r:[x0,y0,x1,y1]}] }]  temuan yang ditandai (k = nomor urut - 1)
 *  aktif    k temuan yang sedang dipilih (digulir ke sana saat berubah)
 */
export default function PdfViewer({ kunci, muat, ukuran, temuan, aktif, tandaiGulir, onPilih, onGagal }) {
  const root = useRef(null);
  const [pdf, setPdf] = useState(null);
  const [zoom, setZoom] = useState(1);
  const [lebarWadah, setLebarWadah] = useState(0);
  const [hal, setHal] = useState(1);

  useEffect(() => {
    let batal = false, tugas = null;
    muat()
      .then((buf) => {
        if (batal) return null;
        tugas = pdfjs.getDocument({ data: buf });
        return tugas.promise;
      })
      .then((d) => { if (d && !batal) setPdf(d); })
      .catch(() => { if (!batal) onGagal?.(); });
    return () => { batal = true; tugas?.destroy(); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [kunci]);

  useLayoutEffect(() => {
    const ro = new ResizeObserver(([e]) => setLebarWadah(Math.floor(e.contentRect.width)));
    ro.observe(root.current);
    return () => ro.disconnect();
  }, []);

  const jumlah = pdf?.numPages ?? ukuran?.length ?? 0;
  const lebar = Math.max(240, Math.round((lebarWadah - 2 * GAP) * zoom));
  const tinggiHal = (n) => {
    const [w, h] = ukuran?.[n - 1] || ukuran?.[0] || [595, 842];
    return Math.round((lebar * h) / w);
  };

  // sorotan dikelompokkan per halaman
  const perHalaman = useMemo(() => {
    const m = new Map();
    (temuan || []).forEach(({ k, jenis, rects }) => rects.forEach(({ p, r }, i) => {
      if (!m.has(p)) m.set(p, []);
      m.get(p).push({ k, jenis, r, pertama: i === 0 });
    }));
    return m;
  }, [temuan]);

  const tujuan = useMemo(() => (temuan || []).find((t) => t.k === aktif), [temuan, aktif]);

  // Gulung ke temuan aktif saat dipilih (tandaiGulir berubah tiap kali pengguna memilih temuan atau bab).
  useEffect(() => {
    if (!tujuan || !root.current || !lebarWadah) return;
    const p = tujuan.rects[0]?.p ?? tujuan.hal;
    if (!p) return;
    let atas = GAP;
    for (let n = 1; n < p; n++) atas += tinggiHal(n) + GAP;
    if (tujuan.rects[0]) atas += tujuan.rects[0].r[1] * tinggiHal(p);
    root.current.scrollTo({ top: Math.max(0, atas - root.current.clientHeight / 3), behavior: "smooth" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tandaiGulir, lebarWadah > 0]);

  const saatGulir = (e) => {
    const el = e.currentTarget;
    const tengah = el.scrollTop + el.clientHeight / 3;
    let atas = GAP, n = 1;
    while (n < jumlah && atas + tinggiHal(n) < tengah) { atas += tinggiHal(n) + GAP; n++; }
    setHal(n);
  };

  const loncat = (n) => {
    let atas = GAP;
    for (let i = 1; i < n; i++) atas += tinggiHal(i) + GAP;
    root.current.scrollTo({ top: atas - GAP, behavior: "smooth" });
  };

  return (
    <div className="pdfv">
      <div className="pdfbar">
        <label>
          Halaman
          <input
            type="number" min={1} max={jumlah || 1} value={hal}
            onChange={(e) => { const n = Math.min(Math.max(+e.target.value || 1, 1), jumlah || 1); setHal(n); loncat(n); }}
          />
          <span>dari {jumlah || "…"}</span>
        </label>
        <div>
          <button type="button" onClick={() => setZoom((z) => Math.max(0.6, +(z - 0.2).toFixed(1)))} aria-label="Perkecil">−</button>
          <span>{Math.round(zoom * 100)}%</span>
          <button type="button" onClick={() => setZoom((z) => Math.min(2.2, +(z + 0.2).toFixed(1)))} aria-label="Perbesar">+</button>
        </div>
      </div>
      <div className="pdfscroll" ref={root} onScroll={saatGulir} tabIndex={0}>
        {lebarWadah > 0 && jumlah > 0 && (
          <div className="pdfpages" style={{ width: lebar, gap: GAP, padding: GAP }}>
            {Array.from({ length: jumlah }, (_, i) => i + 1).map((n) => (
              <Halaman
                key={n} pdf={pdf} nomor={n} lebar={lebar} tinggi={tinggiHal(n)} root={root}
                sorot={perHalaman.get(n) || []} aktif={aktif} onPilih={onPilih}
              />
            ))}
          </div>
        )}
        {!pdf && <p className="pdfmuat">Membuka dokumen…</p>}
      </div>
    </div>
  );
}
