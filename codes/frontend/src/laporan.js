// Membuat laporan hasil pemeriksaan sebagai DOCX atau PDF. Pustaka dimuat saat dibutuhkan saja,
// jadi tidak menambah ukuran bundel halaman lain.
// data: { file, score, bab: [{ n, s, f: [{ t, lokasi, d }] }] }

const nilai = (n) => (n == null ? "-" : `${Math.round(n)}%`);
const namaBerkas = (file, ext) => {
  const dasar = String(file || "dokumen").replace(/\.[^.]+$/, "").replace(/[^\w\- ]+/g, "").trim().replace(/\s+/g, "-");
  return `laporan-penscek-${dasar || "dokumen"}.${ext}`;
};

function simpan(blob, nama) {
  const el = document.createElement("a");
  el.href = URL.createObjectURL(blob);
  el.download = nama;
  el.click();
  setTimeout(() => URL.revokeObjectURL(el.href), 1000);
}

export async function unduhDocx({ file, score, bab }) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = await import("docx");
  const anak = [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("Laporan PENScek")] }),
    new Paragraph({ children: [new TextRun({ text: file, bold: true })] }),
    new Paragraph({ spacing: { after: 240 }, children: [new TextRun(`Kesesuaian keseluruhan: ${nilai(score)}`)] }),
  ];
  bab.forEach((x) => {
    anak.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240 }, children: [new TextRun(`${x.n} (${nilai(x.s)})`)] }));
    if (!x.f.length) {
      anak.push(new Paragraph({ children: [new TextRun({ text: "Tidak ada temuan.", italics: true })] }));
      return;
    }
    x.f.forEach((f, i) => anak.push(new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: `${i + 1}. ${f.t}`, bold: true }), new TextRun({ text: `, ${f.lokasi}: ${f.d}` })],
    })));
  });
  const blob = await Packer.toBlob(new Document({ sections: [{ children: anak }] }));
  simpan(blob, namaBerkas(file, "docx"));
}

export async function unduhPdf({ file, score, bab }) {
  const { jsPDF } = await import("jspdf");
  const pdf = new jsPDF({ unit: "pt", format: "a4" });
  const M = 56, W = pdf.internal.pageSize.getWidth() - M * 2, H = pdf.internal.pageSize.getHeight() - M;
  let y = M;

  const tulis = (teks, { size = 11, bold = false, gap = 6, italic = false } = {}) => {
    pdf.setFont("helvetica", bold ? "bold" : italic ? "italic" : "normal");
    pdf.setFontSize(size);
    pdf.splitTextToSize(String(teks), W).forEach((baris) => {
      if (y + size > H) { pdf.addPage(); y = M; }
      pdf.text(baris, M, y + size);
      y += size * 1.35;
    });
    y += gap;
  };

  tulis("Laporan PENScek", { size: 22, bold: true, gap: 8 });
  tulis(file, { size: 12, bold: true, gap: 2 });
  tulis(`Kesesuaian keseluruhan: ${nilai(score)}`, { gap: 16 });
  bab.forEach((x) => {
    tulis(`${x.n} (${nilai(x.s)})`, { size: 14, bold: true, gap: 4 });
    if (!x.f.length) { tulis("Tidak ada temuan.", { italic: true, gap: 14 }); return; }
    x.f.forEach((f, i) => tulis(`${i + 1}. ${f.t}, ${f.lokasi}: ${f.d}`, { gap: 6 }));
    y += 8;
  });
  simpan(pdf.output("blob"), namaBerkas(file, "pdf"));
}
