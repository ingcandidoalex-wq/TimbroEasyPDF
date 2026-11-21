"""Test del disegno a mano libera sul PDF: motore (insert_stroke/strokes_on/erase_strokes/stroke_hit),
opacità dell'evidenziatore, e il giro completo nell'app (penna disegna, gomma cancella, Annulla ripristina)."""
import shutil
import sys
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402

tp.LICENSE_VERSION = 0
CFG = WORK / "cfg_disegno"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()


def quasi_bianco(px):
    return min(px[:3]) > 245


# ---------------------------------------------------------------- 1. motore: disegna, ritrova, cancella
doc = pymupdf.open()
page = doc.new_page(width=300, height=400)
xref = tp.insert_stroke(page, [(50, 50), (250, 50)], (0, 0, 1), 8, 1.0)
assert xref, "insert_stroke deve restituire lo xref del nuovo tratto"
tratti = tp.strokes_on(page)
assert len(tratti) == 1 and tratti[0][1] == xref, tratti
pm = page.get_pixmap(dpi=72)
assert not quasi_bianco(pm.pixel(150, 50)), "il tratto doveva lasciare inchiostro al centro"
tp.erase_strokes(page, [xref])
assert tp.strokes_on(page) == [], "dopo erase_strokes non devono restare tratti"
assert quasi_bianco(page.get_pixmap(dpi=72).pixel(150, 50)), "dopo la gomma il punto deve tornare bianco"
print("1. motore insert/strokes_on/erase OK")

# ---------------------------------------------------------------- 2. stroke_hit
linea = [(50, 50), (250, 50)]
assert tp.stroke_hit(linea, 8, (150, 52), 1), "un punto sulla linea deve risultare colpito"
assert not tp.stroke_hit(linea, 8, (150, 120), 1), "un punto lontano non deve risultare colpito"
print("2. stroke_hit OK")

# ---------------------------------------------------------------- 3. evidenziatore: 0.4 più chiaro del pieno
giallo = tp.hex_to_rgb("#ffd600")
d1 = pymupdf.open(); p1 = d1.new_page(width=200, height=120)
tp.insert_stroke(p1, [(20, 60), (180, 60)], giallo, 16, 1.0)
d2 = pymupdf.open(); p2 = d2.new_page(width=200, height=120)
tp.insert_stroke(p2, [(20, 60), (180, 60)], giallo, 16, tp.HIGHLIGHT_OPACITY)
b_pieno = p1.get_pixmap(dpi=72).pixel(100, 60)[2]
b_evid = p2.get_pixmap(dpi=72).pixel(100, 60)[2]
assert b_evid > b_pieno + 40, f"l'evidenziatore (0.4) doveva essere più chiaro: blu pieno={b_pieno}, evid={b_evid}"
print(f"3. evidenziatore semitrasparente OK (blu pieno={b_pieno}, evidenziatore={b_evid})")

# ---------------------------------------------------------------- 4. app: penna disegna, gomma cancella, annulla
root = tk.Tk()
root.geometry("900x1200+40+40")
app = tp.App(root)
src = CFG / "foglio.pdf"
d = pymupdf.open(); d.new_page(width=300, height=400); src.write_bytes(d.tobytes())
assert app.open_file(str(src)) is True
app.set_zoom_mode("page")                     # pagina intera: nessuno scorrimento, coordinate semplici
app.set_tool("draw_pen")                      # con uno strumento attivo la vista è a pagina singola
root.update()


def ev(px, py):
    """Evento finto alle coordinate-finestra che corrispondono al punto (px, py) della pagina vista."""
    ox, oy = app.page_origin
    z = app.cur_zoom
    r = app.doc[app.current].rect
    return type("E", (), {"x": ox + (px - r.x0) * z, "y": oy + (py - r.y0) * z})()


def disegna(p_from, p_to):
    app._on_press(ev(*p_from))
    mx = ((p_from[0] + p_to[0]) / 2, (p_from[1] + p_to[1]) / 2)
    app._on_drag(ev(*mx))
    app._on_drag(ev(*p_to))
    app._on_release(ev(*p_to))
    root.update()


page = app.doc[app.current]
disegna((60, 120), (240, 120))
assert len(tp.strokes_on(page)) == 1, f"la penna doveva aggiungere 1 tratto, trovati {len(tp.strokes_on(page))}"
print("4. penna: tratto disegnato sulla pagina OK")

# ---------------------------------------------------------------- 5. gomma: cancella il tratto
app.set_tool("draw_erase")
root.update()
app._on_press(ev(150, 120))                   # proprio sopra il tratto
app._on_release(ev(150, 120))
root.update()
assert tp.strokes_on(app.doc[app.current]) == [], "la gomma doveva togliere il tratto"
print("5. gomma: tratto cancellato OK")

# ---------------------------------------------------------------- 6. Annulla ripristina il tratto cancellato
app.undo()
root.update()
assert len(tp.strokes_on(app.doc[app.current])) == 1, "Ctrl+Z doveva ripristinare il tratto cancellato"
app.set_tool(None)
assert app.tool is None
print("6. Annulla ripristina il disegno OK")

root.destroy()
print("TUTTI I TEST OK")
