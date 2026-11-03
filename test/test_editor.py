"""Test: rotazione e ritocco dei timbri (galleria e pagina), anteprima del testo, cartelle ricordate."""
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402  (rende importabile anche il programma)
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402
from shot import grab  # noqa: E402

tp.LICENSE_VERSION = 0   # niente finestra della licenza nei test
tp.tk.Menu.tk_popup = lambda self, *a, **k: None
CFG = WORK / "cfg_editor"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
errors = []
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.showinfo = lambda *a, **k: None
tp.messagebox.showerror = lambda *a, **k: errors.append(a)
OUT = WORK / "out_editor"
shutil.rmtree(OUT, ignore_errors=True)
OUT.mkdir()
tp.set_dpi_awareness()

RED, BLUE, WHITE = (220, 30, 30), (30, 30, 220), (255, 255, 255)


def image(w, h, fill, alpha=False):
    """Immagine RGB (o RGBA) da una funzione fill(x, y) -> (r, g, b)."""
    data = bytearray()
    for y in range(h):
        for x in range(w):
            data += bytes(fill(x, y)) + (b"\xff" if alpha else b"")
    return pymupdf.Pixmap(pymupdf.csRGB, w, h, bytes(data), alpha)


def px(pix, x, y):
    i = (y * pix.width + x) * pix.n
    return tuple(pix.samples[i:i + pix.n])


def close(a, b, tol=40):
    return all(abs(u - v) <= tol for u, v in zip(a, b))


# ---------------------------------------------------------------- 1. rotazione senza perdite
src = image(3, 2, lambda x, y: (x * 80, y * 100, 7))            # ogni pixel diverso
r90 = tp.rotate_pixmap(src, 90)
assert (r90.width, r90.height) == (2, 3)
# in senso orario: la colonna di sinistra (dal basso in alto) diventa la riga in alto
assert px(r90, 0, 0) == px(src, 0, 1) and px(r90, 1, 0) == px(src, 0, 0) and px(r90, 0, 2) == px(src, 2, 1)
r270 = tp.rotate_pixmap(src, -90)
assert px(r270, 0, 0) == px(src, 2, 0) and px(r270, 1, 2) == px(src, 0, 1)
assert tp.rotate_pixmap(r90, 90).samples == tp.rotate_pixmap(src, 180).samples
full = src
for _ in range(4):
    full = tp.rotate_pixmap(full, 90)
assert full.samples == src.samples, "quattro quarti di giro: identica"
rgba = image(5, 3, lambda x, y: (x * 50, 0, 0), alpha=True)
assert tp.rotate_pixmap(tp.rotate_pixmap(rgba, 90), -90).samples == rgba.samples and rgba.alpha
gray = pymupdf.Pixmap(pymupdf.csGRAY, 4, 2, bytes(range(8)), False)
assert tp.rotate_pixmap(gray, 180).samples == bytes(range(8))[::-1]
print("rotazione dei pixel di 90°, 180°, 270° (RGB, RGBA, grigio) senza perdite: ok")

# trasparenza separata e ricomposta: i bordi sfumati non si scuriscono
sig, _ = tp.text_signature("Prova", None, 30, (0.1, 0.2, 0.6))
colors, mask = tp.split_alpha(sig)
assert colors.n == 3 and not colors.alpha and mask.n == 1
back = tp.join_alpha(colors, mask)
assert pymupdf.Pixmap(None, back).samples == pymupdf.Pixmap(None, sig).samples
worst = max(abs(a - b) for a, b in zip(back.samples, sig.samples))
assert worst <= 2, worst
assert tp.join_alpha(colors, tp.opaque_mask(colors)) is colors, "maschera tutta opaca: niente alpha"
print(f"trasparenza separata e ricomposta (scarto massimo {worst}): ok")


# ---------------------------------------------------------------- app
def pump(seconds=0.0):
    end = time.time() + seconds
    while True:
        root.update()
        if time.time() >= end:
            break
        time.sleep(0.02)


root = tk.Tk()
app = tp.App(root)
root.geometry("1400x940+20+20")
root.attributes("-topmost", True)
root.update()
lib = app.library

# ---------------------------------------------------------------- 2. rotazione dei timbri della galleria
halves = image(80, 40, lambda x, y: RED if x < 40 else BLUE)       # sinistra rossa, destra blu
halves.save(str(OUT / "meta.png"))
item = lib.add(str(OUT / "meta.png"), "Metà", width_mm=40)
app._stamp_selected()
original_png = app.stamp.original_png
app.rotate_stamp(90)
st = app.stamp
assert item["rotate"] == 90 and (st.pixmap.width, st.pixmap.height) == (40, 80)
assert item["width_mm"] == 20, "40 × 20 mm diventa 20 × 40 mm: le misure vere restano"
assert close(px(st.pixmap, 20, 5), RED) and close(px(st.pixmap, 20, 75), BLUE), "a destra: la sinistra va in alto"
assert st.original_png == original_png, "l'immagine originale resta diritta"
app.rotate_stamp(-90)
assert item["rotate"] == 0 and item["width_mm"] == 40 and app.stamp.pixmap.width == 80
app.rotate_stamp(90)
again = tp.StampLibrary(app.cfg)                                  # come alla prossima apertura
it2 = next(i for i in again.items if i["id"] == item["id"])
assert it2["rotate"] == 90 and again.stamps[item["id"]].pixmap.width == 40, "la rotazione resta salvata"

typed_pix, box = tp.text_signature("03/11/2026", None, 20, (0.1, 0.1, 0.1))
typed_src = {"kind": "typed", "text": "03/11/2026", "font": "Helvetica", "size": 20.0, "color": "#111111"}
typed = lib.add(typed_pix, "Data", width_mm=23, remove_white=False, source=typed_src)
app._stamp_selected()
app.rotate_stamp(-90)
assert typed["rotate"] == 270 and typed["source"] == typed_src, "ruotata, la data resta modificabile come testo"
app.rotate_stamp(90)
print("timbro della galleria ruotato a destra e a sinistra, misure e testo conservati, salvato: ok")


# ---------------------------------------------------------------- 3. editor del timbro
class AutoEditor(tp.StampEditor):
    action = None
    seen = {}

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.top.after(80, self._auto)

    def _auto(self):
        AutoEditor.action(self)
        if self.top.winfo_exists():
            self._save()


class AutoText(tp.TextDialog):
    """La finestra vera del testo: scrive, controlla l'anteprima, conferma."""
    answer = {}
    seen = {}

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.top.after(60, self._auto)

    def _auto(self):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", AutoText.answer["text"])
        self.font_var.set(AutoText.answer["font"])
        self.size_var.set(str(AutoText.answer["size"]))
        self._set_color(AutoText.answer["color"])
        self.top.after(300, self._check)

    def _check(self):
        AutoText.seen.update(font=self.text.cget("font"), fg=self.text.cget("foreground"),
                             note=self.format_note.cget("text"))
        AutoText.seen["preview"] = AutoText.check() if AutoText.check else None
        grab(self.top, OUT / "testo_formattato.png")
        if self.top.master is root:                 # sulla pagina: la finestra principale con l'anteprima
            grab(root, OUT / f"anteprima_{AutoText.answer['text']}.png")
            v = app.view
            x0, y0, x1, y1 = v.bbox("textpreview")
            dx, dy = v.winfo_rootx() - v.canvasx(0), v.winfo_rooty() - v.canvasy(0)
            t = self.top
            AutoText.seen["covers"] = not (t.winfo_rootx() + t.winfo_width() < x0 + dx or t.winfo_rootx() > x1 + dx
                                           or t.winfo_rooty() + t.winfo_height() < y0 + dy
                                           or t.winfo_rooty() - 30 > y1 + dy)
        self._ok()


AutoText.check = None
tp.StampEditor, tp.TextDialog = AutoEditor, AutoText


def drag(canvas, points):
    canvas.event_generate("<ButtonPress-1>", x=points[0][0], y=points[0][1])
    for x, y in points[1:]:
        canvas.event_generate("<B1-Motion>", x=x, y=y, state=0x100)
    canvas.event_generate("<ButtonRelease-1>", x=points[-1][0], y=points[-1][1])


white_img = image(200, 100, lambda x, y: WHITE)
white_img.save(str(OUT / "bianco.png"))
plain = lib.add(str(OUT / "bianco.png"), "Bianco", width_mm=50, remove_white=False)
app._stamp_selected()


def edit_plain(ed):
    c = ed.canvas
    ox, oy, f = ed.ox, ed.oy, ed.f
    to_canvas = lambda x, y: (round(ox + x * f), round(oy + y * f))   # noqa: E731
    ed.size_scale.set(10)
    drag(c, [to_canvas(20 + 4 * k, 50) for k in range(20)])          # matita: riga orizzontale a metà
    AutoEditor.seen["after_pencil"] = ed.colors
    assert ed.edited and ed.undo_stack
    ed.undo()
    AutoEditor.seen["undo"] = ed.colors.samples == white_img.samples
    ed.redo()
    ed.tool_var.set("text")
    ed._tool_changed()
    AutoText.answer = {"text": "OK", "font": "Arial", "size": 8, "color": "#b3261e"}
    AutoText.check = lambda: bool(ed.canvas.find_withtag("textpreview"))
    x, y = to_canvas(120, 10)
    ed.canvas.event_generate("<ButtonPress-1>", x=x, y=y)
    grab(ed.top, OUT / "editor.png")
    ed.rotate(90)


AutoEditor.action = edit_plain
app.cmd_edit_stamp()
pencil = AutoEditor.seen["after_pencil"]
assert close(px(pencil, 60, 50), (31, 58, 147)), px(pencil, 60, 50)   # il blu della matita
assert px(pencil, 60, 90) == WHITE, "lontano dal tratto non cambia niente"
assert AutoEditor.seen["undo"], "annulla riporta l'immagine di prima"
assert AutoText.seen["preview"], "anteprima del testo sul timbro mentre si scrive"
st = app.stamp
assert plain["rotate"] == 0 and "source" not in plain, "ritoccata: la rotazione è dentro l'immagine"
assert (st.pixmap.width, st.pixmap.height) == (100, 200) and plain["width_mm"] == 25
reds = sum(1 for k in range(0, len(st.pixmap.samples), 3) if close(st.pixmap.samples[k:k + 3], (179, 38, 30), 50))
assert reds > 10, "la scritta rossa c'è"
print(f"editor: matita, annulla/ripristina, testo con anteprima ({reds} pixel rossi), rotazione: ok")

# gomma su un timbro trasparente: i pixel diventano trasparenti
sig_item = lib.add(sig, "Firma", width_mm=40, remove_white=False,
                   source={"kind": "typed", "text": "Prova", "font": "Helvetica", "size": 30.0, "color": "#1a3399"})
app._stamp_selected()


def erase_all(ed):
    ed.tool_var.set("eraser")
    ed._tool_changed()
    ed.size_scale.set(10)
    c, w, h = ed.canvas, ed.colors.width, ed.colors.height
    for y in range(0, h, 6):
        drag(c, [(round(ed.ox + x * ed.f), round(ed.oy + y * ed.f)) for x in range(0, w + 1, 8)])
    AutoEditor.seen["note"] = ed.note.cget("text")
    AutoEditor.seen["source_btn"] = str(ed.source_btn.state())


AutoEditor.action = erase_all
app.cmd_edit_stamp()
alpha = pymupdf.Pixmap(None, app.stamp.pixmap).samples
assert alpha.count(0) > 0.95 * len(alpha), "quasi tutto cancellato"
assert "immagine ritoccata" in AutoEditor.seen["note"] and "disabled" in AutoEditor.seen["source_btn"]
assert "source" not in sig_item, "dopo un ritocco la firma resta un'immagine"
print("gomma: la firma diventa trasparente, il testo non è più modificabile (con avviso): ok")

# solo ruotata nell'editor: la firma resta modificabile come testo
app.select_stamp(lib.items.index(typed))
AutoEditor.action = lambda ed: ed.rotate(90)
app.cmd_edit_stamp()
assert typed["rotate"] == 90 and typed["source"] == typed_src
print("firma solo ruotata nell'editor: testo ancora modificabile: ok")

# editor di immagini esterno: quando il file cambia, l'immagine si aggiorna da sola
opened = []
tp.os.startfile = lambda *a: opened.append(a)


def external(ed):
    ed._open_external()
    path = ed._ext["path"]
    edited = image(100, 50, lambda x, y: BLUE)      # come se in Paint fosse stata ritagliata e colorata
    time.sleep(0.05)
    edited.save(str(path))
    end = time.time() + 5
    while time.time() < end and ed.colors.width != 100:
        ed.top.update()
        time.sleep(0.05)
    AutoEditor.seen["external"] = (ed.colors.width, px(ed.colors, 50, 25))


app.select_stamp(lib.items.index(item))     # «Metà», 40 × 80 px ruotata, larga 20 mm
AutoEditor.action = external
app.cmd_edit_stamp()
assert opened and opened[0][1] == "edit", opened
assert AutoEditor.seen["external"] == (100, BLUE), AutoEditor.seen["external"]
assert abs(item["width_mm"] - 100 * 20 / 40) < 0.1, "ritagliata altrove: stessa scala"
print("immagine aperta nell'editor esterno e ricaricata da sola quando la si salva: ok")

# ---------------------------------------------------------------- 4. timbri già apposti
doc = pymupdf.open()
for rot in (0, 90):
    p = doc.new_page(width=595, height=842)
    p.insert_text((72, 100), "Pagina", fontsize=20)
    p.set_rotation(rot)
doc.save(OUT / "pagine.pdf")
assert app.open_file(str(OUT / "pagine.pdf"))
app.set_zoom_mode("page")
pump(0.2)
halves.save(str(OUT / "meta.png"))
item2 = lib.add(str(OUT / "meta.png"), "Metà 2", width_mm=60)
app._stamp_selected()


def at(x, y):
    """Punto della pagina vista (in punti) -> coordinate della finestra della vista."""
    ox, oy = app.page_origin
    return round(ox + x * app.cur_zoom - app.view.canvasx(0)), round(oy + y * app.cur_zoom - app.view.canvasy(0))


def stamp_colors(page, stamp):
    """Colori in alto, in basso, a sinistra e a destra del timbro come lo si vede."""
    pix = page.get_pixmap(dpi=72, clip=stamp.rect)
    w, h = pix.width, pix.height
    return {"top": px(pix, w // 2, 2), "bottom": px(pix, w // 2, h - 3),
            "left": px(pix, 2, h // 2), "right": px(pix, w - 3, h // 2)}


for number in (0, 1):
    app.go_to(number)
    pump(0.1)
    page = app.doc[number]
    pr = page.rect
    app.set_tool("stamp")
    app._place_stamp(*at(pr.width / 2, 200))
    app.set_tool(None)
    page = app.doc[number]
    stamp = tp.placed_stamps(page)[-1]
    before = stamp_colors(page, stamp)
    assert close(before["left"], RED) and close(before["right"], BLUE), (number, before)
    assert tp.stamp_view_angle(page, stamp) == 0
    center = ((stamp.rect.x0 + stamp.rect.x1) / 2, (stamp.rect.y0 + stamp.rect.y1) / 2)
    w, h = stamp.rect.width, stamp.rect.height
    app.rotate_placed(stamp, 90)
    page = app.doc[number]
    stamp = tp.placed_stamps(page)[-1]
    after = stamp_colors(page, stamp)
    assert close(after["top"], RED) and close(after["bottom"], BLUE), (number, after)
    assert abs(stamp.rect.width - h) < 0.5 and abs(stamp.rect.height - w) < 0.5
    assert abs((stamp.rect.x0 + stamp.rect.x1) / 2 - center[0]) < 0.5
    assert tp.stamp_view_angle(page, stamp) == 90
    app.rotate_placed(stamp, -90)
    stamp = tp.placed_stamps(app.doc[number])[-1]
    assert tp.stamp_view_angle(app.doc[number], stamp) == 0
print("timbro apposto ruotato a destra e a sinistra attorno al centro, anche su pagina ruotata: ok")

# menu del tasto destro su un timbro apposto
labels = []
real_popup = tp.tk.Menu.tk_popup
tp.tk.Menu.tk_popup = lambda self, *a, **k: labels.extend(self.entrycget(i, "label") for i in range(self.index("end") + 1)
                                                           if self.type(i) == "command")
r = stamp.rect
cx, cy = app._canvas_rect(r)[0] + 10 - app.view.canvasx(0), app._canvas_rect(r)[1] + 10 - app.view.canvasy(0)
app._on_view_menu(type("E", (), {"x": cx, "y": cy, "x_root": 0, "y_root": 0})())
assert labels == ["Modifica questo timbro…", "Ruota a destra di 90°", "Ruota a sinistra di 90°",
                  "Elimina questo timbro"], labels
tp.tk.Menu.tk_popup = real_popup

# modifica del timbro apposto sulla pagina ruotata: resta dov'era, cambia solo lui
page = app.doc[1]
index_before = stamp.index
count = len(page.get_contents())
AutoEditor.action = lambda ed: (ed.rotate(90), AutoEditor.seen.update(width=ed.width_mm))
app.edit_placed_stamp(stamp)
page = app.doc[1]
new = next(s for s in tp.placed_stamps(page) if s.index == index_before)
assert len(page.get_contents()) == count, "stesso numero di stream: sostituito, non aggiunto"
assert abs((new.rect.x0 + new.rect.x1) / 2 - (r.x0 + r.x1) / 2) < 1 and abs(new.rect.width - r.height) < 1
colors_now = stamp_colors(page, new)
assert close(colors_now["top"], RED) and close(colors_now["bottom"], BLUE), colors_now
assert tp.stamp_view_angle(page, new) == 0, "la nuova immagine è diritta: la rotazione è nei pixel"
assert app.stamp.pixmap.width == 80, "il timbro della galleria non cambia"
print("timbro apposto modificato (ruotato nell'editor): stesso posto, stesso ordine, galleria intatta: ok")

# firma scritta apposta: si riscrive il testo direttamente sul PDF
app.select_stamp(lib.items.index(typed))
app.go_to(0)
pump(0.1)
app.set_tool("stamp")
app._place_stamp(*at(300, 400))
app.set_tool(None)
page = app.doc[0]
placed = tp.placed_stamps(page)[-1]
_, xref = tp.placed_stamp_image(page, placed)
assert tp.stamp_source(app.doc, xref) == (typed_src, 90), "il PDF ricorda il testo della firma e la rotazione"
real_sig = tp.SignatureDialog


class AutoSig(real_sig):
    def __init__(self, parent, source=None, name=None):
        super().__init__(parent, source=source, name=name)
        AutoSig.seen = (self.tabs.index("current"), self.text_var.get())
        parent.after(60, self._auto)

    def _auto(self):
        self.text_var.set("04/11/2026")
        self.top.after(200, self._save)


tp.SignatureDialog = AutoSig
AutoEditor.action = lambda ed: ed._edit_source()
app.edit_placed_stamp(placed)
assert AutoSig.seen == (1, "03/11/2026"), AutoSig.seen
page = app.doc[0]
placed = tp.placed_stamps(page)[-1]
_, xref = tp.placed_stamp_image(page, placed)
src_new, rot_new = tp.stamp_source(app.doc, xref)
assert src_new["text"] == "04/11/2026" and rot_new == 90, (src_new, rot_new)
assert typed["source"]["text"] == "03/11/2026", "la galleria non cambia"
print("data apposta riscritta sul PDF (03/11 -> 04/11), rotazione conservata: ok")

# timbri apposti con la versione precedente (senza il testo nel PDF): lo si ritrova dalla galleria
app.select_stamp(lib.items.index(typed))
app.set_tool("stamp")
app._place_stamp(*at(300, 600))
app.set_tool(None)
page = app.doc[0]
placed = tp.placed_stamps(page)[-1]
_, xref = tp.placed_stamp_image(page, placed)
app.doc.xref_set_key(xref, tp.SOURCE_KEY, "null")
assert tp.stamp_source(app.doc, xref) is None
AutoEditor.action = lambda ed: AutoEditor.seen.update(source=ed.source)
app.edit_placed_stamp(placed)
assert AutoEditor.seen["source"] == typed_src, "riconosciuta dalla galleria"
print("firma apposta senza testo nel PDF riconosciuta dalla galleria: ok")
tp.SignatureDialog = real_sig

# ---------------------------------------------------------------- 5. testo: anteprima sulla pagina
app.go_to(0)
app.set_zoom_mode("page")
pump(0.2)
AutoText.answer = {"text": "Anteprima", "font": "Times New Roman", "size": 22, "color": "#1f3a93"}
AutoText.check = lambda: bool(app.view.find_withtag("textpreview"))
app.set_tool("add_text")
x, y = at(100, 500)
app.add_text_at(x, y)
assert AutoText.seen["preview"], "mentre si scrive il testo si vede sulla pagina"
assert not AutoText.seen["covers"], "la finestra del testo non copre l'anteprima"
assert "Times New Roman" in str(AutoText.seen["font"]) and "22" in str(AutoText.seen["font"]), AutoText.seen
assert AutoText.seen["fg"] == "#1f3a93"
assert not app.view.find_withtag("textpreview"), "chiusa la finestra, l'anteprima sparisce"
box = tp.placed_texts(app.doc[0])[-1]
assert box.params["text"] == "Anteprima" and abs(box.rect.x0 - 100) < 2, box
app.set_tool(None)
# il riquadro mostra il testo grande fin dove ci sta, e lo dice
AutoText.answer = dict(AutoText.answer, text="Grande", size=60)
app.set_tool("add_text")
app.add_text_at(x, y + 200)
assert "36" in str(AutoText.seen["font"]) and "60 pt" in AutoText.seen["note"], AutoText.seen
app.set_tool(None)
# l'anteprima coincide con il testo scritto davvero
page = app.doc[0]
params = dict(box.params)
prev, clip = tp.text_preview(page, params, app.fonts.get(params["font"]), 1)
real = page.get_pixmap(clip=clip, alpha=False)
ink_prev = sum(1 for a in pymupdf.Pixmap(None, prev).samples if a > 128)
ink_real = sum(1 for k in range(0, len(real.samples), 3) if real.samples[k] < 128)
assert ink_prev and abs(ink_prev - ink_real) <= 0.15 * ink_real, (ink_prev, ink_real)
print(f"testo: formattato nel riquadro e con anteprima sulla pagina, uguale a quello scritto ({ink_prev}/{ink_real}): ok")

# ---------------------------------------------------------------- 6. cartelle ricordate
folders = {k: OUT / k for k in ("aperti", "salvati")}
for f in folders.values():
    f.mkdir()
shutil.copy(OUT / "pagine.pdf", folders["aperti"] / "da_aprire.pdf")
asked = {}
tp.filedialog.askopenfilenames = lambda **k: (asked.update(open=k["initialdir"]),
                                              (str(folders["aperti"] / "da_aprire.pdf"),))[1]
app.cmd_open()
assert app.cfg["dirs"]["open"] == str(folders["aperti"])
tp.filedialog.asksaveasfilename = lambda **k: (asked.update(save=k["initialdir"]),
                                               str(folders["salvati"] / "firmato.pdf"))[1]
app.cmd_save()
assert asked["save"] == str(folders["aperti"]), "la prima volta: la cartella del documento"
app.open_file(str(OUT / "pagine.pdf"))
app.cmd_save()
assert asked["save"] == str(folders["salvati"]), "poi: l'ultima cartella di salvataggio"
app.cmd_open()
assert asked["open"] == str(folders["aperti"]), "aprire riparte da dove si è aperto, non da dove si è salvato"
print("cartelle ricordate per aprire e salvare, separate: ok")

assert not errors, errors
tp.StampEditor, tp.TextDialog = AutoEditor.__mro__[1], AutoText.__mro__[1]
root.destroy()
print("TUTTI I TEST OK")
