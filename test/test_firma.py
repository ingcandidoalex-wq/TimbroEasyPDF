"""Test: spostare/eliminare timbri già apposti, firma disegnata e scritta."""
import math
import shutil
import sys
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402  (rende importabile anche il programma)
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402
tp.LICENSE_VERSION = 0   # niente finestra della licenza nei test
tp.tk.Menu.tk_popup = lambda self, *a, **k: None   # i menu del tasto destro non vanno aperti sullo schermo
from shot import grab  # noqa: E402

CFG = WORK / "cfg_firma"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.showinfo = lambda *a, **k: None
OUT = WORK / "out"
OUT.mkdir(exist_ok=True)

# documento: 3 pagine uguali + 1 ruotata
doc = pymupdf.open()
for i in range(4):
    p = doc.new_page(width=595, height=842)
    p.insert_text((50, 80), f"Pagina {i + 1}", fontsize=26)
    if i == 3:
        p.set_rotation(90)
doc.save(OUT / "quattro.pdf")

tp.set_dpi_awareness()
root = tk.Tk()
app = tp.App(root)
root.geometry("1360x900+30+30")
root.update()

# ---------------------------------------------------------------- firma disegnata
dlg = tp.SignatureDialog(root)
root.update()
pad = dlg.pad
W, H = dlg.W, dlg.H
pts = [(0.15 + 0.7 * t / 60, 0.55 + 0.18 * math.sin(t / 4)) for t in range(61)]   # un'onda
pad.event_generate("<Button-1>", x=int(pts[0][0] * W), y=int(pts[0][1] * H))
for fx, fy in pts[1:]:
    pad.event_generate("<B1-Motion>", x=int(fx * W), y=int(fy * H), state=0x100)
pad.event_generate("<ButtonRelease-1>", x=int(pts[-1][0] * W), y=int(pts[-1][1] * H))
pad.event_generate("<Button-1>", x=int(0.5 * W), y=int(0.3 * H))     # un punto
root.update()
assert len(dlg.strokes) == 2 and len(dlg.strokes[0]) > 20, [len(s) for s in dlg.strokes]
grab(dlg.top, OUT / "firma_disegna.png")
dlg._undo_stroke()
assert len(dlg.strokes) == 1
dlg.name_var.set("Firma a mano")
dlg._save()
pix, name, width_mm, _src = dlg.result
assert pix.alpha and name == "Firma a mano" and 15 <= width_mm <= 150
print(f"firma disegnata: {pix.width}x{pix.height}px, {width_mm} mm: ok")
app.library.add(pix, name, width_mm=width_mm, remove_white=False)

# ---------------------------------------------------------------- firma scritta
dlg = tp.SignatureDialog(root)
root.update()
print("font installati:", len(dlg.fonts), "| primo proposto:", dlg.font_var.get())
dlg.tabs.select(1)
dlg.text_var.set("Ing. Mario Bianchi")
dlg.size_var.set("28")
dlg._set_color(tp.INK_COLORS["Nero"])
dlg._render_preview()
root.update()
grab(dlg.top, OUT / "firma_scrivi.png")
dlg._save()
pix_t, name_t, width_t, _src = dlg.result
_, box28 = tp.text_signature("Ing. Mario Bianchi", dlg.fonts.get(dlg.font_var.get()), 28, (0, 0, 0))
_, box14 = tp.text_signature("Ing. Mario Bianchi", dlg.fonts.get(dlg.font_var.get()), 14, (0, 0, 0))
assert abs(box28.width / box14.width - 2) < 0.15, "la dimensione in pt deve scalare la firma"
assert abs(width_t - box28.width / tp.MM) < 0.2
print(f"firma scritta: {width_t} mm a 28 pt (14 pt: {box14.width / tp.MM:.0f} mm): ok")
app.library.add(pix_t, "Firma scritta", width_mm=width_t, remove_white=False)
app._stamp_selected()

# ---------------------------------------------------------------- timbri apposti: sposta / elimina
app.open_file(str(OUT / "quattro.pdf"))
app.set_zoom_mode("page")
root.update()


def to_window(fx, fy):
    pr = app.doc[app.current].rect
    ox, oy = app.page_origin
    return (int(ox + fx * pr.width * app.cur_zoom - app.view.canvasx(0)),
            int(oy + fy * pr.height * app.cur_zoom - app.view.canvasy(0)))


def click(fx, fy):
    x, y = to_window(fx, fy)
    app.view.event_generate("<Motion>", x=x, y=y)
    app.view.event_generate("<ButtonPress-1>", x=x, y=y)
    app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
    root.update()


def drag(f_from, f_to):
    x0, y0 = to_window(*f_from)
    x1, y1 = to_window(*f_to)
    app.view.event_generate("<Motion>", x=x0, y=y0)
    app.view.event_generate("<ButtonPress-1>", x=x0, y=y0)
    for k in range(1, 11):
        app.view.event_generate("<B1-Motion>", x=x0 + (x1 - x0) * k // 10, y=y0 + (y1 - y0) * k // 10,
                                state=0x100)
        root.update()
    app.view.event_generate("<ButtonRelease-1>", x=x1, y=y1)
    root.update()


def centers(page):
    return [((s.rect.x0 + s.rect.x1) / 2 / page.rect.width, (s.rect.y0 + s.rect.y1) / 2 / page.rect.height)
            for s in tp.placed_stamps(page)]


app.select_stamp(1)                                  # firma scritta
app.set_stamp_mode(True)
for i in range(4):
    app.go_to(i)
    root.update()
    click(0.7, 0.85)
app.select_stamp(0)                                  # firma a mano, solo pagina 1
app.go_to(0)
root.update()
click(0.3, 0.5)
app.set_stamp_mode(False)
assert len(tp.placed_stamps(app.doc[0])) == 2

# hover + trascinamento sulla pagina 1
app.go_to(0)
root.update()
x, y = to_window(0.7, 0.85)
app.view.event_generate("<Motion>", x=x, y=y)
root.update()
assert app.view.find_withtag("mark") and str(app.view.cget("cursor")) == "fleur", "evidenziazione al passaggio"
drag((0.7, 0.85), (0.5, 0.6))
c = centers(app.doc[0])
assert any(abs(cx - 0.5) < 0.02 and abs(cy - 0.6) < 0.02 for cx, cy in c), c
assert app.sel_stamp is not None and app.view.find_withtag("mark")
grab(root, OUT / "sposta_timbro.png")
app.undo()
assert any(abs(cx - 0.7) < 0.02 and abs(cy - 0.85) < 0.02 for cx, cy in centers(app.doc[0]))
app.redo()

# pagina ruotata: trascinamento verso sinistra
app.go_to(3)
root.update()
drag((0.7, 0.85), (0.3, 0.85))
assert any(abs(cx - 0.3) < 0.02 and abs(cy - 0.85) < 0.02 for cx, cy in centers(app.doc[3])), centers(app.doc[3])

# eliminazione con clic + Canc (non deve eliminare la pagina)
app.go_to(0)
root.update()
click(0.3, 0.5)
assert app.sel_stamp is not None
app.view.focus_force()
app.view.event_generate("<Delete>")
root.update()
assert len(app.doc) == 4 and len(tp.placed_stamps(app.doc[0])) == 1, "Canc deve togliere solo il timbro"
print("sposta (anche su pagina ruotata), annulla, elimina con Canc: ok")

# ---------------------------------------------------------------- dopo salvataggio e riapertura
out = OUT / "firmato.pdf"
tp.filedialog.asksaveasfilename = lambda **k: str(out)
assert app.cmd_save()
app.open_file(str(out))
root.update()
# le pagine 2 e 3 hanno timbri identici: dopo garbage=3 possono condividere lo stesso stream
app.go_to(1)
root.update()
drag((0.7, 0.85), (0.35, 0.2))
assert any(abs(cx - 0.35) < 0.02 and abs(cy - 0.2) < 0.02 for cx, cy in centers(app.doc[1])), centers(app.doc[1])
assert all(abs(cx - 0.7) < 0.02 for cx, cy in centers(app.doc[2])), "la pagina 3 non deve cambiare"
app.go_to(2)
root.update()
x, y = to_window(0.7, 0.85)
app.view.event_generate("<Button-3>", x=x, y=y, rootx=0, rooty=0)
root.update()
app.delete_selected_stamp()
assert not tp.placed_stamps(app.doc[2])
assert app.cmd_save()
res = pymupdf.open(out)
print("timbri per pagina nel PDF finale:", [len(tp.placed_stamps(p)) for p in res])
assert [len(tp.placed_stamps(p)) for p in res] == [1, 1, 0, 1]
print("riapertura: timbri ancora modificabili, pagine indipendenti: ok")
res[0].get_pixmap(dpi=50).save(str(OUT / "firmato_p1.png"))

root.destroy()
print("TUTTI I TEST OK")
