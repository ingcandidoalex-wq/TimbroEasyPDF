"""Test: schede, trascinamento pagine, modifica testo, caselle, bianchetto, ricerca, OCR, Word."""
import shutil
import sys
import time
import tkinter as tk
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402  (rende importabile anche il programma)
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402
tp.LICENSE_VERSION = 0   # niente finestra della licenza nei test
from shot import grab  # noqa: E402

CFG = WORK / "cfg_nuove"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
errors = []
answers = {"askyesnocancel": False, "askyesno": True}
tp.messagebox.askyesno = lambda *a, **k: answers["askyesno"]
tp.messagebox.askyesnocancel = lambda *a, **k: answers["askyesnocancel"]
tp.messagebox.showinfo = lambda *a, **k: None
tp.messagebox.showwarning = lambda *a, **k: None
tp.messagebox.showerror = lambda *a, **k: errors.append(a)
OUT = WORK / "out_nuove"
shutil.rmtree(OUT, ignore_errors=True)
OUT.mkdir()
ARIAL = r"C:\Windows\Fonts\arial.ttf"

# ---------------------------------------------------------------- documenti di prova
rel = pymupdf.open()
for n in (1, 2):
    p = rel.new_page(width=595, height=842)
    p.insert_text((72, 100), f"Relazione pagina {n}", fontsize=20, fontname="helv")
    p.insert_text((72, 150), "Importo lavori: euro 12.500,00", fontsize=12, fontname="F1", fontfile=ARIAL)
    p.insert_text((72, 170), "Riga successiva da non toccare", fontsize=12, fontname="F1", fontfile=ARIAL)
rel.save(OUT / "relazione.pdf")
tav = pymupdf.open()
for n in (1, 2, 3):
    tav.new_page(width=842, height=595).insert_text((72, 100), f"Tavola {n}", fontsize=30, fontname="helv")
tav.save(OUT / "tavola.pdf")
# scansione: pagina dritta e pagina ruotata di 90Â° (dritta a schermo)
scan = pymupdf.open()
expected_scan = {}
for rot in (0, 90):
    page = scan.new_page(width=595, height=842)
    page.set_rotation(rot)
    src = pymupdf.open()
    sp = src.new_page(width=page.rect.width, height=page.rect.height)
    sp.insert_text((60, 160), "Il sottoscritto ingegnere dichiara che l'opera e conforme.", fontsize=12,
                   fontname="tiro")
    sp.insert_text((60, 185), "Verbale di collaudo numero 4587", fontsize=12, fontname="tiro")
    expected_scan[rot] = sp.search_for("conforme")[0]
    page.insert_image(page.rect * page.derotation_matrix, pixmap=sp.get_pixmap(dpi=200), rotate=page.rotation)
scan.save(OUT / "scansione.pdf")

tp.set_dpi_awareness()
root = tk.Tk()
app = tp.App(root)
root.geometry("1400x940+20+20")
root.attributes("-topmost", True)   # winfo_containing deve trovare le nostre finestre
root.update()


def pump(seconds=0.0):
    end = time.time() + seconds
    while True:
        root.update()
        if time.time() >= end:
            break
        time.sleep(0.02)


def names():
    return [t.name for t in app.tabs]


# ---------------------------------------------------------------- 1. schede
assert app.open_file(str(OUT / "relazione.pdf")) and app.open_file(str(OUT / "tavola.pdf"))
pump()
assert names() == ["relazione.pdf", "tavola.pdf"] and app.tab is app.tabs[1]
assert app.open_file(str(OUT / "relazione.pdf")) and len(app.tabs) == 2 and app.tab is app.tabs[0], "giÃ  aperto"
rel_tab, tav_tab = app.tabs
app.rotate(90)                                     # modifica solo relazione
assert rel_tab.dirty and not tav_tab.dirty and len(rel_tab.undo_stack) == 1 and not tav_tab.undo_stack
app.switch_tab(tav_tab)
app.undo()                                         # niente da annullare in tavola
assert rel_tab.doc[0].rotation == 90
app.switch_tab(rel_tab)
app.undo()
assert rel_tab.doc[0].rotation == 0
assert "relazione.pdf" in root.title()
print("schede: apertura multipla, giÃ  aperto, annulla separato: ok")


# ---------------------------------------------------------------- 2. trascinamento miniature
def thumb_xy(i):
    _x0, y0, _x1, y1, _s = app.thumb_layout[i]
    return app.thumbs.winfo_width() // 2, int((y0 + y1) / 2 - app.thumbs.canvasy(0))


def thumb_drag(i, target, tx, ty, wait=0.0, target_after=None):
    """Trascina la miniatura i fino al punto (tx, ty) del widget target."""
    th = app.thumbs
    x, y = thumb_xy(i)
    rx, ry = th.winfo_rootx(), th.winfo_rooty()
    th.event_generate("<Button-1>", x=x, y=y, rootx=rx + x, rooty=ry + y)
    pump()
    gx, gy = target.winfo_rootx() + tx, target.winfo_rooty() + ty
    for k in range(1, 9):
        px, py = rx + x + (gx - rx - x) * k // 8, ry + y + (gy - ry - y) * k // 8
        th.event_generate("<B1-Motion>", x=px - rx, y=py - ry, rootx=px, rooty=py, state=0x100)
        pump()
    if wait:
        pump(wait)
    if target_after:   # dopo il cambio di scheda: ci si sposta sulle miniature
        target, tx, ty = target_after
        gx, gy = target.winfo_rootx() + tx, target.winfo_rooty() + ty
        th.event_generate("<B1-Motion>", x=gx - rx, y=gy - ry, rootx=gx, rooty=gy, state=0x100)
        pump()
    th.event_generate("<ButtonRelease-1>", x=gx - rx, y=gy - ry, rootx=gx, rooty=gy)
    pump()


def texts(tab):
    return [tab.doc[i].get_text().split("\n")[0] for i in range(len(tab.doc))]


# riordino nella stessa scheda: Tavola 1 in fondo
app.switch_tab(tav_tab)
pump()
last = app.thumb_layout[-1]
thumb_drag(0, app.thumbs, app.thumbs.winfo_width() // 2, int(last[3] + 10 - app.thumbs.canvasy(0)))
assert texts(tav_tab) == ["Tavola 2", "Tavola 3", "Tavola 1"], texts(tav_tab)
app.undo()
assert texts(tav_tab) == ["Tavola 1", "Tavola 2", "Tavola 3"]

# copia su un'altra scheda rilasciando sulla scheda: va in fondo, la scheda attiva non cambia
frame_rel = next(f for f in app._tab_frames if f.doc_tab is rel_tab)
thumb_drag(1, frame_rel, 20, 10)
assert app.tab is tav_tab and texts(rel_tab)[-1] == "Tavola 2" and len(rel_tab.doc) == 3, texts(rel_tab)
assert rel_tab.dirty and texts(tav_tab) == ["Tavola 1", "Tavola 2", "Tavola 3"], "l'origine non cambia"
grab(root, OUT / "schede.png")

# fermandosi sulla scheda questa si apre: rilascio fra la pagina 1 e la 2 di relazione
app.selection = {0, 2}                              # due pagine selezionate
app._paint_selection()
pump()
frame_rel = next(f for f in app._tab_frames if f.doc_tab is rel_tab)
thumb_drag(0, frame_rel, 20, 10, wait=1.0,
           target_after=(app.thumbs, app.thumbs.winfo_width() // 2, 2))
# il rilascio a y=0 della tela cade prima della prima miniatura visibile: posizione 0
assert app.tab is rel_tab, "la scheda doveva aprirsi fermandosi sopra"
assert texts(rel_tab)[:2] == ["Tavola 1", "Tavola 3"] and len(rel_tab.doc) == 5, texts(rel_tab)
assert app.selection == {0, 1}
print("trascinamento: riordino, copia su scheda, apertura della scheda e inserimento nel punto: ok")

# rilascio sul "+" -> nuovo documento con quella pagina
thumb_drag(3, app._plus, 10, 10)                    # [T1, T3, R1, R2, T2]: la pagina R2
assert len(app.tabs) == 3 and texts(app.tabs[2]) == ["Relazione pagina 2"] and app.tab is rel_tab
nuovo = app.tabs[2]
app.close_tab(nuovo)                                 # modificato: risposta "non salvare"
assert len(app.tabs) == 2
app.cmd_new()
assert app.tab.doc is not None and len(app.tab.doc) == 0
assert not app.cmd_save(), "un documento vuoto non si salva"
app.open_file(str(OUT / "scansione.pdf"))           # il documento vuoto intatto lascia il posto
assert names() == ["relazione.pdf", "tavola.pdf", "scansione.pdf"], names()
print("nuovo documento dal Â«+Â», chiusura scheda, documento vuoto: ok")

# ---------------------------------------------------------------- 3. modifica di una riga
app.switch_tab(rel_tab)
for _ in range(3):
    app.undo()                                       # torna alla relazione originale
assert texts(rel_tab) == ["Relazione pagina 1", "Relazione pagina 2"], texts(rel_tab)
app.go_to(0)
app.set_zoom_mode("page")
pump()


def to_window(fx, fy):
    pr = app.doc[app.current].rect
    ox, oy = app.page_origin
    return (int(ox + fx * pr.width * app.cur_zoom - app.view.canvasx(0)),
            int(oy + fy * pr.height * app.cur_zoom - app.view.canvasy(0)))


def at_point(px, py):
    """Punto della pagina vista (in punti) -> coordinate della finestra."""
    pr = app.doc[app.current].rect
    return to_window(px / pr.width, py / pr.height)


# un timbro sulla pagina, per verificare che la modifica del testo non lo tocchi
stamp_png = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 60, 30), 1)
stamp_png.clear_with(0)
stamp_png.set_rect(stamp_png.irect, (200, 30, 30, 255))
app.library.add(stamp_png, "Prova", width_mm=60, remove_white=False)
app._stamp_selected()
app.set_stamp_mode(True)
x, y = at_point(200, 152)                            # sopra la riga da correggere
app.view.event_generate("<Motion>", x=x, y=y)
app.view.event_generate("<ButtonPress-1>", x=x, y=y)
app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
pump()
assert len(tp.placed_stamps(app.doc[0])) == 1

# la finestra vera di modifica: una volta a schermo
line = next(l for l in tp.text_lines(app.doc[0]) if l.text.startswith("Importo"))
assert tp.match_font(line.font, line.flags, app.fonts) == "Arial", line.font
dlg = tp.TextDialog(root, app.fonts, "Modifica testo", {"text": line.text, "font": "Arial", "size": 12,
                                                          "color": line.color}, original="Testo originale: Arial",
                    whiteout=False)
pump()
grab(dlg.top, OUT / "dialogo_testo.png")
dlg.top.destroy()


class FakeDialog:
    """Sostituisce TextDialog: risponde subito con il risultato preparato."""
    result_next = None
    seen = None

    def __init__(self, parent, fonts, title, params, original=None, whiteout=None):
        FakeDialog.seen = dict(params=params, whiteout=whiteout, title=title)
        self.top = tk.Toplevel(parent)
        self.result = FakeDialog.result_next
        parent.after(20, self.top.destroy)


tp.TextDialog = FakeDialog
app.set_tool("edit_text")
x, y = at_point(120, 146)
app.view.event_generate("<Motion>", x=x, y=y)
pump()
assert app.view.find_withtag("hover"), "la riga sotto il mouse deve evidenziarsi"
FakeDialog.result_next = {"text": "Importo lavori: euro 13.750,00 (rivisto)", "font": "Arial", "size": 12,
                          "color": "#000000", "whiteout": False}
app.view.event_generate("<ButtonPress-1>", x=x, y=y)
app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
pump()
assert FakeDialog.seen["params"]["text"] == "Importo lavori: euro 12.500,00" and FakeDialog.seen["whiteout"] is False, (FakeDialog.seen, app.doc[0].get_image_info(xrefs=True), [(i[0], app.doc.xref_get_key(i[0], tp.STAMP_KEY)) for i in app.doc[0].get_images(full=True)])
page = app.doc[0]
assert not page.search_for("12.500"), "il vecchio testo deve sparire davvero"
assert page.search_for("13.750,00 (rivisto)"), "testo nuovo con spazi normali"
assert "Riga successiva da non toccare" in page.get_text().replace("\xa0", " "), "la riga vicina non va toccata"
new = page.search_for("13.750")[0]
assert abs(new.y1 - 152) < 4, new
stamps = tp.placed_stamps(page)
assert len(stamps) == 1, "il timbro deve restare riconoscibile dopo la modifica del testo"
boxes = tp.placed_texts(page)
assert len(boxes) == 1 and boxes[0].params["text"].endswith("(rivisto)")
assert boxes[0].index < stamps[0].index, "la riga corretta deve restare sotto il timbro in primo piano"
print("modifica riga: vecchio testo eliminato, nuovo al suo posto, riga vicina e timbro intatti: ok")

# ---------------------------------------------------------------- 4. caselle di testo
app.set_tool("add_text")
FakeDialog.result_next = {"text": "Visto, si approva\nIl tecnico", "font": "Times New Roman", "size": 14,
                          "color": "#1f3a93", "whiteout": False}
x, y = at_point(300, 600)
app.view.event_generate("<ButtonPress-1>", x=x, y=y)
app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
pump()
boxes = tp.placed_texts(app.doc[0])
assert len(boxes) == 2
box = boxes[-1]
assert abs(box.rect.x0 - 300) < 2 and abs(box.rect.y0 - 600) < 2, box.rect
assert app.cfg["text_style"]["font"] == "Times New Roman", "lo stile usato si ricorda"
app.set_tool(None)
pump()
cx, cy = (box.rect.x0 + box.rect.x1) / 2, (box.rect.y0 + box.rect.y1) / 2
x0, y0 = at_point(cx, cy)
x1, y1 = at_point(cx + 100, cy + 50)
app.view.event_generate("<Motion>", x=x0, y=y0)
app.view.event_generate("<ButtonPress-1>", x=x0, y=y0)
for k in range(1, 9):
    app.view.event_generate("<B1-Motion>", x=x0 + (x1 - x0) * k // 8, y=y0 + (y1 - y0) * k // 8, state=0x100)
    pump()
app.view.event_generate("<ButtonRelease-1>", x=x1, y=y1)
pump()
moved = tp.placed_texts(app.doc[0])[-1]
assert abs(moved.rect.x0 - 400) < 2 and abs(moved.rect.y0 - 650) < 2, moved.rect
hit = app.doc[0].search_for("Visto, si")[0]
assert abs(hit.x0 - 400) < 3 and 648 < hit.y0 < 670, hit
FakeDialog.result_next = dict(FakeDialog.result_next, text="Visto, si approva con riserva\nIl tecnico")
app._on_double_click(type("E", (), {"x": x1, "y": y1})())
pump()
again = tp.placed_texts(app.doc[0])[-1]
assert again.params["text"].startswith("Visto, si approva con riserva") and abs(again.rect.x0 - 400) < 2
assert app.doc[0].search_for("con riserva")

# casella su pagina ruotata: dritta a schermo
app.go_to(1)
app.rotate(90)
app.set_tool("add_text")
FakeDialog.result_next = {"text": "Nota sulla pagina ruotata", "font": "Arial", "size": 12, "color": "#000000",
                          "whiteout": False}
x, y = at_point(100, 100)
app.view.event_generate("<ButtonPress-1>", x=x, y=y)
app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
pump()
pg = app.doc[1]
hit = pg.search_for("ruotata")[0] * pg.rotation_matrix
assert hit.width > hit.height and abs(hit.y1 - 100 - 12) < 6, ("testo dritto vicino al clic", hit)
app.set_tool(None)
print("caselle di testo: aggiunta, spostamento, correzione, pagina ruotata: ok")

# ---------------------------------------------------------------- 5. bianchetto
app.go_to(0)
pump()
app.set_tool("whiteout")
a = at_point(60, 160)
b = at_point(330, 178)
app.view.event_generate("<ButtonPress-1>", x=a[0], y=a[1])
app.view.event_generate("<B1-Motion>", x=b[0], y=b[1], state=0x100)
app.view.event_generate("<ButtonRelease-1>", x=b[0], y=b[1])
pump()
page = app.doc[0]
assert "Riga successiva" not in page.get_text().replace("\xa0", " "), "il testo sotto il bianchetto va eliminato"
assert len(tp.placed_stamps(page)) == 1 and len(tp.placed_texts(page)) == 2, "timbri e caselle restano"
app.set_tool(None)
grab(root, OUT / "testo.png")
print("bianchetto: testo eliminato, timbro e caselle intatti: ok")

# ---------------------------------------------------------------- 6. salvataggio
out = OUT / "relazione_modificata.pdf"
tp.filedialog.asksaveasfilename = lambda **k: str(out)
assert app.cmd_save()
size_kb = out.stat().st_size // 1024
res = pymupdf.open(out)
assert len(tp.placed_texts(res[0])) == 2 and len(tp.placed_stamps(res[0])) == 1
assert "Visto, si approva con riserva" in res[0].get_text(), res[0].get_text()
assert size_kb < 200, f"i font vanno ridotti ai caratteri usati ({size_kb} KB)"
# dopo il salvataggio le caselle si spostano e correggono ancora
box = tp.placed_texts(app.doc[0])[-1]
app.sel_stamp = box.index
FakeDialog.result_next = dict(FakeDialog.result_next, text="Testo corretto dopo il salvataggio")
app.edit_text_box(box)
assert app.doc[0].search_for("dopo il salvataggio")
print(f"salvataggio ({size_kb} KB): caselle e timbri ancora modificabili: ok")

# ---------------------------------------------------------------- 7. ricerca
app.find_var.set("Relazione pagina")
app.find(1)
pump()
assert app.find_lbl.cget("text") == "1 di 2" and app.view.find_withtag("find")
app.find(1)
assert app.current == 1 and app.find_lbl.cget("text") == "2 di 2"
app.clear_find()
print("ricerca: ok")

# ---------------------------------------------------------------- 8. OCR
scan_tab = next(t for t in app.tabs if t.name == "scansione.pdf")
app.switch_tab(scan_tab)
assert not app.doc[0].get_text().strip()
state = {"done": False}
t0 = time.time()
app.cmd_ocr(on_done=lambda: state.update(done=True), all_pages=True)
while not state["done"] and time.time() - t0 < 120:
    pump(0.05)
assert state["done"], "OCR non terminato"
for i, rot in enumerate((0, 90)):
    pg = app.doc[i]
    hit = pg.search_for("conforme")
    assert hit, f"pagina {i + 1}: testo non riconosciuto: {pg.get_text()!r}"
    view = hit[0] * pg.rotation_matrix
    exp = expected_scan[rot]
    assert abs(view.x0 - exp.x0) < 4 and abs(view.y1 - exp.y1) < 6, (rot, view, exp)
print(f"OCR su pagina dritta e ruotata in {time.time() - t0:.1f} s, testo allineato alla scansione: ok")

# modifica di una riga riconosciuta: la casella Â«copri con il biancoÂ» Ã¨ proposta giÃ  spuntata
app.go_to(0)
app.set_tool("edit_text")
line = next(l for l in tp.text_lines(app.doc[0]) if "collaudo" in l.text)
c = (line.rect.tl + line.rect.br) / 2
x, y = at_point(c.x, c.y)
FakeDialog.result_next = {"text": "Verbale n. 4600", "font": "Times New Roman", "size": 12,
                          "color": "#000000", "whiteout": True}
app.view.event_generate("<ButtonPress-1>", x=x, y=y)
app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
pump()
app.set_tool(None)
assert FakeDialog.seen["whiteout"] is True, "su una scansione il bianco va proposto"
assert app.doc[0].search_for("4600"), app.doc[0].get_text()
assert not app.doc[0].search_for("4587"), "il vecchio testo riconosciuto deve sparire"
r = line.rect
pix = app.doc[0].get_pixmap(clip=pymupdf.Rect(r.x0 + r.width * 0.6, r.y0 + 2, r.x1 - 1, r.y1 - 2))
assert min(pix.samples) > 200, "dove c'era il resto della riga la scansione deve essere coperta di bianco"
grab(root, OUT / "ocr.png")
print("correzione su scansione con bianco sotto: ok")

# ---------------------------------------------------------------- 9. Word
docx = OUT / "scansione.docx"
tp.filedialog.asksaveasfilename = lambda **k: str(docx)
answers["askyesno"] = False                          # non aprire Word alla fine
t0 = time.time()
app.cmd_export_word()
while not docx.exists() and time.time() - t0 < 180:
    pump(0.2)
pump(0.5)
assert docx.exists(), f"docx non creato: {errors}"
xml = zipfile.ZipFile(docx).read("word/document.xml").decode("utf-8")
assert "conforme" in xml and "4600" in xml, xml[:500]
print(f"esportazione in Word in {time.time() - t0:.0f} s: testo presente nel .docx: ok")

assert not errors, errors
root.destroy()
print("TUTTI I TEST OK")

