"""
Compare two structure_*.json files produced by extract_structure.py
(e.g. one from Windows, one from macOS) and report differences.

Usage:
    python compare_structure.py structure_Windows.json structure_Darwin.json
"""
import sys
import json


def compare(path_a, path_b):
    a = json.load(open(path_a, encoding="utf-8"))
    b = json.load(open(path_b, encoding="utf-8"))

    print(f"Comparing {a['platform']} vs {b['platform']}\n")

    diffs_found = False

    # Sections (margins, mirror margin)
    if a["sections"] != b["sections"]:
        diffs_found = True
        print("=== DIFFERENCE in sections (margins / mirror margins) ===")
        for i, (sa, sb) in enumerate(zip(a["sections"], b["sections"])):
            if sa != sb:
                print(f"  Section {i}:")
                for key in sa:
                    if sa[key] != sb[key]:
                        print(f"    {key}: {a['platform']}={sa[key]!r}  vs  {b['platform']}={sb[key]!r}")
        print()

    # Paragraphs
    if len(a["paragraphs"]) != len(b["paragraphs"]):
        diffs_found = True
        print(f"=== DIFFERENCE in paragraph count: "
              f"{a['platform']}={len(a['paragraphs'])} vs {b['platform']}={len(b['paragraphs'])} ===\n")
    else:
        para_diffs = []
        for i, (pa, pb) in enumerate(zip(a["paragraphs"], b["paragraphs"])):
            if pa != pb:
                para_diffs.append((i, pa, pb))
        if para_diffs:
            diffs_found = True
            print(f"=== DIFFERENCE in {len(para_diffs)} paragraph(s) ===")
            for i, pa, pb in para_diffs[:20]:  # cap output
                print(f"  Paragraph {i}:")
                for key in pa:
                    if pa[key] != pb[key]:
                        print(f"    {key}: {a['platform']}={pa[key]!r}  vs  {b['platform']}={pb[key]!r}")
            if len(para_diffs) > 20:
                print(f"  ... and {len(para_diffs) - 20} more paragraph diffs (truncated)")
            print()

    if not diffs_found:
        print("No structural differences found. Struktur identik antar OS.")
    else:
        print("Ada perbedaan struktur di atas — ini bukti konkret buat dilaporkan ke Pak Ali.")


def main():
    if len(sys.argv) != 3:
        print("Usage: python compare_structure.py structure_A.json structure_B.json")
        sys.exit(1)
    compare(sys.argv[1], sys.argv[2])


if __name__ == "__main__":
    main()
