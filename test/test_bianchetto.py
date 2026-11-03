"""Test: bianchetto e oscura non devono annerire il testo con il contorno (grassetto simulato, scritte CAD).

MuPDF riscrive la pagina perdendo lo spessore della linea di quel testo: senza la correzione di
restore_stroke_widths le scritte diventano macchie nere. Il PDF di prova imita quelli di JasperReports
passati da Ghostscript (scala 0,1 e poi 10, «2 Tr»), con dati inventati.
"""
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402,F401  (rende importabile anche il programma)
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402

PAGE = b"""q 0.1 0 0 0.1 0 0 cm
0 g 0 G 0.33333 w 2 J
q 10 0 0 10 0 0 cm BT /F1 10 Tf 1 0 0 1 362 730 Tm 2 Tr [(20261)1(24420)1]TJ -45 0 Td 0 Tr [(Richiesta:)]TJ ET Q
q 10 0 0 10 0 0 cm BT /F1 11 Tf 1 0 0 1 362 712 Tm 2 Tr (ROSSI*MARIO) Tj ET Q
Q
q 1 0 0 RG 0.3 w BT 1 Tr /F1 30 Tf 72 600 Td (Contorno \\(rosso\\)) Tj ET Q
q 10 w 0 0 1 RG 50 100 m 500 100 l S Q
q 0 g BT /F1 14 Tf 72 300 Td (CANCELLAMI) Tj ET Q
q 20 0 0 10 400 300 cm BI /W 2 /H 1 /CS /RGB /BPC 8 ID AEI\xff\x00\x00 EI Q"""   # nei dati c'è «EI»: va saltato


def make_pdf():
    objs = {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
           b"/Resources << /Font << /F1 5 0 R >> >> >>",
        4: b"<< /Length %d >>\nstream\n" % len(PAGE) + PAGE + b"\nendstream",
        5: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>",
    }
    out = bytearray(b"%PDF-1.7\n")
    offsets = {}
    for n in sorted(objs):
        offsets[n] = len(out)
        out += b"%d 0 obj\n" % n + objs[n] + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 6\n0000000000 65535 f \n" + b"".join(b"%010d 00000 n \n" % offsets[n] for n in sorted(objs))
    out += b"trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % xref
    return pymupdf.open("pdf", bytes(out))


def widths(page):
    return sorted((t, round(w, 3)) for t, _, w in tp.stroked_spans(page))


def darker_outside(before, after, area, dpi):
    z = dpi / 72
    inside = (area + (-1, -1, 1, 1)) * pymupdf.Matrix(z, z)
    b, a, w = before.samples, after.samples, before.width
    return sum(1 for y in range(before.height) for x in range(w)
               if not inside.contains(pymupdf.Point(x, y)) and sum(a[(y * w + x) * 3:(y * w + x) * 3 + 3])
               < sum(b[(y * w + x) * 3:(y * w + x) * 3 + 3]) - 150)


# lettura del contenuto: stringhe con parentesi e «\», immagini in linea
ops = [op for op, *_ in tp.content_ops(PAGE)]
assert ops.count(b"BT") == 4 and ops.count(b"Tj") == 3 and ops.count(b"TJ") == 2 and b"ID" in ops and b"EI" not in ops, ops
assert [w for *_, w in tp.stroked_text_ops(PAGE)] == [0.33333, 0.33333, 0.3], list(tp.stroked_text_ops(PAGE))
print("lettura degli operatori (stringhe, immagini in linea, spessori del testo a contorno): ok")

dpi = 110
for fill, name in (((1, 1, 1), "bianchetto"), ((0, 0, 0), "oscura")):
    doc = make_pdf()
    page = doc[0]
    original = widths(page)
    assert original == [("2026124420", 0.333), ("Contorno (rosso)", 0.3), ("ROSSI*MARIO", 0.333)], original
    # un timbro e una casella di testo del programma sulla stessa pagina
    stamp = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), 0)
    stamp.clear_with(90)
    xref = page.insert_image(pymupdf.Rect(400, 500, 440, 540), stream=stamp.tobytes("png"))
    tp.mark_stamp_image(doc, xref)
    tp.insert_text_box(page, {"text": "Nota", "font": "Helvetica", "size": 12, "color": "#1f3a93",
                              "origin": [72, 700], "rotate": 0}, None, "helv")
    before = page.get_pixmap(dpi=dpi)
    area = page.search_for("CANCELLAMI")[0] + (-2, -2, 2, 2)
    tp.erase(page, [area], fill=fill, images=pymupdf.PDF_REDACT_IMAGE_PIXELS,
             graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
    doc = pymupdf.open("pdf", doc.tobytes(garbage=3, deflate=True))
    page = doc[0]
    assert not page.search_for("CANCELLAMI"), "il testo sotto è cancellato davvero"
    assert widths(page) == original, (widths(page), original)
    darker = darker_outside(before, page.get_pixmap(dpi=dpi), area, dpi)
    assert darker == 0, f"{darker} pixel diventati più scuri fuori dall'area"
    assert len(tp.placed_stamps(page)) == 1 and len(tp.placed_texts(page)) == 1, "timbro e casella restano"
    print(f"{name}: testo con il contorno con lo spessore di prima, niente si annerisce fuori dall'area: ok")

# cancellando metà di una scritta in grassetto, il resto mantiene lo spessore
doc = make_pdf()
page = doc[0]
r = page.search_for("ROSSI")[0]
tp.erase(page, [pymupdf.Rect(r.x0, r.y0, (r.x0 + r.x1) / 2, r.y1)], fill=(1, 1, 1),
         images=pymupdf.PDF_REDACT_IMAGE_PIXELS, graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
rest = [(t, round(w, 3)) for t, _, w in tp.stroked_spans(page) if "MARIO" in t]
assert rest and all(w == 0.333 for _, w in rest) and rest[0][0] != "ROSSI*MARIO", rest
print(f"metà scritta cancellata: resta {rest[0][0]!r} con lo spessore giusto: ok")
print("TUTTI I TEST OK")
