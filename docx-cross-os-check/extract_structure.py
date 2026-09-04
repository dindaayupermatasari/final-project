"""
Extract structural info (margins, fonts, alignment, mirror margins) from a
.docx file. Run this on the same file across different OSes/runners, then
diff the resulting JSON with compare_structure.py.

Usage:
    python extract_structure.py path/to/sample.docx
Output:
    structure_<PlatformName>.json  (e.g. structure_Windows.json, structure_Darwin.json)
"""
import sys
import json
import platform

from docx import Document
from docx.shared import Emu
from docx.oxml.ns import qn


def emu_to_cm(val):
    return round(Emu(val).cm, 3) if val is not None else None


def has_mirror_margins(section):
    # python-docx has no high-level API for this; mirror margins only
    # shows up as a <w:mirrorMargins/> element inside <w:sectPr> in the
    # raw XML, so we check for it directly.
    sect_pr = section._sectPr
    return sect_pr.find(qn('w:mirrorMargins')) is not None


def extract(path):
    doc = Document(path)
    result = {"platform": platform.system(), "sections": [], "paragraphs": []}

    for s in doc.sections:
        result["sections"].append({
            "left_margin": emu_to_cm(s.left_margin),
            "right_margin": emu_to_cm(s.right_margin),
            "top_margin": emu_to_cm(s.top_margin),
            "bottom_margin": emu_to_cm(s.bottom_margin),
            "gutter": emu_to_cm(s.gutter),
            "mirror_margins": has_mirror_margins(s),
        })

    for p in doc.paragraphs:
        runs_info = []
        for r in p.runs:
            runs_info.append({
                "text": r.text[:30],
                "font_name": r.font.name,
                "font_size": r.font.size.pt if r.font.size else None,
                "bold": r.bold,
                "italic": r.italic,
            })
        result["paragraphs"].append({
            "style": p.style.name if p.style else None,
            "alignment": str(p.alignment),
            "line_spacing": p.paragraph_format.line_spacing,
            "space_before": p.paragraph_format.space_before.pt if p.paragraph_format.space_before else None,
            "space_after": p.paragraph_format.space_after.pt if p.paragraph_format.space_after else None,
            "runs": runs_info,
        })

    return result


def main():
    if len(sys.argv) != 2:
        print("Usage: python extract_structure.py path/to/sample.docx")
        sys.exit(1)

    result = extract(sys.argv[1])
    out_name = f"structure_{result['platform']}.json"
    with open(out_name, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    print(f"Saved {out_name}")


if __name__ == "__main__":
    main()
