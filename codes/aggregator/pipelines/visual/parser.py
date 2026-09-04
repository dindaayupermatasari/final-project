from docx import Document


def parse_docx(filepath: str) -> Document:
    try:
        doc = Document(filepath)
    except Exception as e:
        raise ValueError(f"Gagal parsing file '{filepath}': {e}")
    return doc
