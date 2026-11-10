"""Test: modifica timbri salvati, unisci/riordina schede, dividi, barre e pannelli agganciabili."""
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
tp.LICENSE_VERSION = 0   # niente finestra della licenza nei test
from shot import grab  # noqa: E402

CFG = WORK / "cfg_layout"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
errors = []
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.askyesnocancel = lambda *a, **k: False
tp.messagebox.showinfo = lambda *a, **k: None
tp.messagebox.showerror = lambda *a, **k: errors.append(a)
tp.os.startfile = lambda *a: None
OUT = WORK / "out_layout"
shutil.rmtree(OUT, ignore_errors=True)
OUT.mkdir()

for name, n in (("alfa", 2), ("beta", 3), ("gamma", 1)):
    d = pymupdf.open()
    for i in range(n):
        d.new_page().insert_text((72, 100), f"{name} {i + 1}", fontsize=24)
    d.save(OUT / f"{name}.pdf")

tp.set_dpi_awareness()


def new_app(geometry="1400x940+20+20"):
    root = tk.Tk()
    app = tp.App(root)
    for _k in ("text", "pages", "stamps"):   # i test della disposizione vogliono tutto a vista
        app._set_visible(_k, True)
    if geometry:
        root.geometry(geometry)
    root.attributes("-topmost", True)
    root.update()
    return root, app


root, app = new_app()


def pump(seconds=0.0):
    end = time.time() + seconds
    while True:
        root.update()
        if time.time() >= end:
            break
        time.sleep(0.02)


# ---------------------------------------------------------------- 1. modifica dei timbri salvati
pix, box = tp.text_signature("Ing. Mario Bianchi", None, 28, (0.1, 0.2, 0.6))
src = {"kind": "typed", "text": "Ing. Mario Bianchi", "font": "Segoe Script", "size": 28.0, "color": "#1f3a93"}
item = app.library.add(pix, "Firma scritta", width_mm=round(box.width / tp.MM, 1), remove_white=False, source=src)
app._stamp_selected()
item["overlay"] = False                    # un'impostazione da conservare
old_width, old_px = item["width_mm"], app.stamp.pixmap.width

seen = {}
Real = tp.SignatureDialog


class AutoEdit(Real):
    """La finestra vera: si controlla cosa mostra, si cambia il testo e si salva."""
    def __init__(self, parent, source=None, name=None):
        super().__init__(parent, source=source, name=name)
        parent.after(80, self._auto)

    def _auto(self):
        seen.update(tab=self.tabs.index("current"), text=self.text_var.get(), font=self.font_var.get(),
                    size=self.size_var.get(), name=self.name_var.get(), color=self.color, strokes=len(self.strokes),
                    title=self.top.title())
        grab(self.top, OUT / "modifica_firma.png")
        if self.tabs.index("current") == 1:
            self.text_var.set("Ing. Mario Bianchi Rossi")
            self.size_var.set("20")
        else:
            self._undo_stroke()
        self.name_var.set(self.name_var.get() + " (mod.)")
        self._save()


class AutoEditor(tp.StampEditor):
    """La finestra vera di modifica del timbro: esegue action(editor) e salva."""
    action = None

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.top.after(80, self._auto)

    def _auto(self):
        AutoEditor.action(self)
        self._save()


tp.SignatureDialog = AutoEdit
tp.StampEditor = AutoEditor
AutoEditor.action = lambda ed: ed._edit_source()       # «Modifica il testo della firma…»
app.cmd_edit_stamp()
assert seen["tab"] == 1 and seen["text"] == "Ing. Mario Bianchi" and seen["size"] == "28", seen
assert seen["title"] == "Modifica firma" and seen["color"] == "#1f3a93", seen
it = app.library.items[0]
assert it is item and it["id"] == item["id"] and len(app.library.items) == 1, "stessa voce, non una nuova"
assert it["name"] == "Firma scritta (mod.)" and it["source"]["text"] == "Ing. Mario Bianchi Rossi"
assert it["overlay"] is False, "le impostazioni restano"
new_px = app.stamp.pixmap.width
assert new_px != old_px and abs(it["width_mm"] - new_px * old_width / old_px) < 0.2, \
    "firma riscritta: stessa scala (mm per pixel), quindi la misura segue il testo"
print(f"firma scritta modificata: stessa voce, {old_width} -> {it['width_mm']} mm alla stessa scala: ok")

# firma disegnata: riaperta con i suoi tratti
strokes = [[[0.1 + 0.02 * k, 0.5 + 0.1 * (k % 2)] for k in range(20)], [[0.6, 0.3]], [[0.7, 0.6], [0.8, 0.7]]]
pix2, _ = tp.drawn_signature([[(x * 680, y * 230) for x, y in s] for s in strokes], (680, 230), 3, (0, 0, 0))
drawn = app.library.add(pix2, "Firma a mano", width_mm=50, remove_white=False,
                        source={"kind": "drawn", "color": "#111111", "thickness": 3, "strokes": strokes})
app._stamp_selected()
old_px = app.stamp.pixmap.width
app.cmd_edit_stamp()
assert seen["tab"] == 0 and seen["strokes"] == 3, seen
assert len(drawn["source"]["strokes"]) == 2, "un tratto annullato"
assert abs(drawn["width_mm"] - app.stamp.pixmap.width * 50 / old_px) < 0.2, "i tratti rimasti non cambiano misura"
print("firma disegnata riaperta con i suoi 3 tratti e modificata alla stessa scala: ok")

# timbro caricato da immagine: si sostituisce l'immagine
img = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 80, 40), 0)
img.clear_with(120)
img.save(str(OUT / "timbro1.png"))
img2 = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 60, 60), 0)
img2.clear_with(60)
img2.save(str(OUT / "timbro2.png"))
stamp_item = app.library.add(str(OUT / "timbro1.png"), "Timbro", width_mm=42)
app._stamp_selected()
tp.filedialog.askopenfilename = lambda **k: str(OUT / "timbro2.png")
AutoEditor.action = lambda ed: ed._replace()           # «Sostituisci immagine…»
app.cmd_edit_stamp()
assert stamp_item["width_mm"] == 42 and abs(app.stamp.aspect - 1) < 0.01, "immagine nuova, misura conservata"
assert [i["name"] for i in app.library.items] == ["Firma scritta (mod.)", "Firma a mano (mod.)", "Timbro"]
assert app.cfg["dirs"]["image"] == str(OUT), "la cartella delle immagini dei timbri viene ricordata"
tp.SignatureDialog = Real
tp.StampEditor = AutoEditor.__mro__[1]
print("timbro da immagine: immagine sostituita mantenendo nome e misura: ok")

# ---------------------------------------------------------------- 2. schede: riordino e unione
for name in ("alfa", "beta", "gamma"):
    app.open_file(str(OUT / f"{name}.pdf"))
pump()
assert [t.name for t in app.tabs] == ["alfa.pdf", "beta.pdf", "gamma.pdf"]


def frame_of(name):
    return next(f for f in app._tab_frames if f.doc_tab.name == name)


def drag_tab(name, target_name, after=True):
    """Trascina la scheda name oltre il centro della scheda target_name."""
    f, t = frame_of(name), frame_of(target_name)
    x0, y0 = f.winfo_rootx() + 10, f.winfo_rooty() + 8
    x1 = t.winfo_rootx() + (t.winfo_width() - 5 if after else 5)
    f.event_generate("<Button-1>", x=10, y=8, rootx=x0, rooty=y0)
    for k in range(1, 9):
        x = x0 + (x1 - x0) * k // 8
        f.event_generate("<B1-Motion>", x=x - f.winfo_rootx(), y=8, rootx=x, rooty=y0, state=0x100)
        pump()
    f.event_generate("<ButtonRelease-1>", x=x1 - f.winfo_rootx(), y=8, rootx=x1, rooty=y0)
    pump()


drag_tab("gamma.pdf", "alfa.pdf", after=False)        # gamma davanti a tutti
assert [t.name for t in app.tabs] == ["gamma.pdf", "alfa.pdf", "beta.pdf"], [t.name for t in app.tabs]
class AutoMerge:
    def __init__(self, parent, tabs):
        self.top = tk.Toplevel(parent)
        self.result = list(tabs)
        parent.after(20, self.top.destroy)


tp.MergeDialog = AutoMerge
app.cmd_merge_tabs()
merged = app.tab
assert merged.name.startswith("Unione") and merged.dirty and len(app.tabs) == 4
order = [p.get_text().strip() for p in merged.doc]
assert order == ["gamma 1", "alfa 1", "alfa 2", "beta 1", "beta 2", "beta 3"], order
print("schede riordinate trascinandole e unite con un clic nell'ordine delle schede: ok")

# ---------------------------------------------------------------- 3. dividi
dlg = tp.SplitDialog(root, 6, "Unione", str(OUT), per=1)
dlg.per_var.set("2")
dlg.name_var.set("Pratica 12/B")                      # la barra non è ammessa nei nomi di file
pump()
assert "3 file" in dlg.preview.cget("text") and "Pratica 12B_01.pdf" in dlg.preview.cget("text"), \
    dlg.preview.cget("text")
grab(dlg.top, OUT / "dividi.png")
dlg._ok()
assert dlg.result == (2, "Pratica 12B", 1, str(OUT)), dlg.result


class AutoSplit:
    result_next = None

    def __init__(self, parent, pages, name, folder, per=1):
        self.top = tk.Toplevel(parent)
        self.result = AutoSplit.result_next
        parent.after(20, self.top.destroy)


tp.SplitDialog = AutoSplit
split_dir = OUT / "divisi"
AutoSplit.result_next = (2, "Pratica", 1, str(split_dir))
app.selection = {0}
app.cmd_split()
files = sorted(p.name for p in split_dir.glob("*.pdf"))
assert files == ["Pratica_01.pdf", "Pratica_02.pdf", "Pratica_03.pdf"], files
assert [len(pymupdf.open(split_dir / f)) for f in files] == [2, 2, 2]
assert pymupdf.open(split_dir / "Pratica_02.pdf")[0].get_text().strip() == "alfa 2"
AutoSplit.result_next = (1, "Pagina", 9, str(split_dir))   # pagine singole, numerazione da 9
app.cmd_split()
files = sorted(p.name for p in split_dir.glob("Pagina_*.pdf"))
assert files == [f"Pagina_{n:02d}.pdf" for n in range(9, 15)], files
assert app.cfg["split_per"] == 1
print("divisione in coppie e in pagine singole con nome e numero progressivo: ok")

# ---------------------------------------------------------------- 4. barre e pannelli
grab(root, OUT / "disposizione_iniziale.png")


def drag_grip(widget, x_root, y_root):
    gx, gy = widget.winfo_rootx() + 4, widget.winfo_rooty() + 4
    widget.event_generate("<Button-1>", x=4, y=4, rootx=gx, rooty=gy)
    for k in range(1, 9):
        x, y = gx + (x_root - gx) * k // 8, gy + (y_root - gy) * k // 8
        widget.event_generate("<B1-Motion>", x=x - widget.winfo_rootx(), y=y - widget.winfo_rooty(),
                              rootx=x, rooty=y, state=0x100)
        pump()
    widget.event_generate("<ButtonRelease-1>", x=x_root - widget.winfo_rootx(), y=y_root - widget.winfo_rooty(),
                          rootx=x_root, rooty=y_root)
    pump(0.3)


rx, ry, rw, rh = root.winfo_rootx(), root.winfo_rooty(), root.winfo_width(), root.winfo_height()
text_bar, main_bar = app.bars["text"], app.bars["main"]
drag_grip(text_bar.grip, rx + 8, ry + rh // 2)            # barra del testo a sinistra
assert app.layout["bars"]["text"] == ["left", 0] and text_bar.side == "left"
assert str(text_bar.pack_info()["in"]) == str(app.docks["left"])
assert text_bar.winfo_height() > text_bar.winfo_width(), "in verticale"
drag_grip(main_bar.grip, rx + rw // 2, ry + rh - 10)      # barra principale in basso
assert app.layout["bars"]["main"] == ["bottom", 0]
assert main_bar.winfo_rooty() > app.view.winfo_rooty(), "sotto la pagina"
head = app.side_panels["thumbs"].winfo_children()[0]
drag_grip(head, rx + rw - 30, ry + rh // 2)               # miniature a destra
panes = [str(p) for p in app.paned.panes()]
assert panes.index(str(app.center)) < panes.index(str(app.side_panels["thumbs"])), panes
assert app.layout["panels"]["thumbs"][0] == "right"
# miniature più larghe trascinando il divisorio: le miniature si ingrandiscono
before = app.thumb_w
i = panes.index(str(app.side_panels["thumbs"])) - 1   # divisorio a sinistra del pannello
x, y = app.paned.sash_coord(i)
app.paned.sash_place(i, x - 120, y)
pump(0.5)
assert app.thumb_w > before + 100, (before, app.thumb_w)
app.go_to(0)
pump(0.5)
grab(root, OUT / "disposizione_spostata.png")
print("barra a sinistra (verticale), barra in basso, miniature a destra e più larghe: ok")

# memoria alla riapertura
thumbs_w = app.side_panels["thumbs"].winfo_width()
root.geometry("1300x880+60+40")
pump(0.3)
app._save_layout()
app._save_window()
root.destroy()
root, app = new_app(geometry=None)
pump(0.5)
assert app.layout["bars"]["text"] == ["left", 0] and app.layout["bars"]["main"] == ["bottom", 0]
assert app.bars["text"].side == "left" and app.bars["main"].side == "bottom"
assert app.layout["panels"]["thumbs"][0] == "right"
assert abs(app.side_panels["thumbs"].winfo_width() - thumbs_w) < 6, (app.side_panels["thumbs"].winfo_width(), thumbs_w)
assert root.geometry() == "1300x880+60+40", root.geometry()
print("riapertura: barre, pannelli, larghezze e finestra come prima: ok")
app.reset_layout()
pump(0.3)
assert app.layout == tp.default_layout()
for key in ("text", "pages", "stamps"):   # il ripristino li richiude: li riapro per controllare la disposizione
    app._set_visible(key, True)
pump(0.3)
assert app.bars["text"].side == "top"
panes = [str(p) for p in app.paned.panes()]
assert panes[0] == str(app.side_panels["thumbs"]) and panes[-1] == str(app.side_panels["stamps"])
print("ripristino della disposizione iniziale: ok")

assert not errors, errors
root.destroy()
print("TUTTI I TEST OK")
