"""Test: archivio dei timbri, pannelli in basso, galleria, vista d'insieme, testo trascinabile,
file trascinati da Esplora risorse, unione selettiva, licenza."""
import ctypes
import json
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

tp.LICENSE_VERSION = 0
tp.tk.Menu.tk_popup = lambda self, *a, **k: None
errors = []
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.showinfo = lambda *a, **k: None
tp.messagebox.showwarning = lambda *a, **k: errors.append(("warning",) + a)
tp.messagebox.showerror = lambda *a, **k: errors.append(("error",) + a)
OUT = WORK / "out_v2"
shutil.rmtree(OUT, ignore_errors=True)
OUT.mkdir()
tp.set_dpi_awareness()


def use_pc(name):
    """Ogni «PC» ha la sua cartella di configurazione (come %APPDATA% su un computer diverso)."""
    cfg = OUT / name
    cfg.mkdir(exist_ok=True)
    tp.CONFIG_DIR, tp.CONFIG_FILE = cfg, cfg / "config.json"
    tp.STAMPS_DIR, tp.OLD_STAMP_FILE = cfg / "timbri", cfg / "timbro.png"
    return cfg


def new_app(geometry="1400x940+20+20"):
    root = tk.Tk()
    app = tp.App(root)
    root.geometry(geometry)
    root.attributes("-topmost", True)
    root.update()
    return root, app


def pump(root, seconds=0.0):
    end = time.time() + seconds
    while True:
        root.update()
        if time.time() >= end:
            break
        time.sleep(0.02)


def png(color, w=80, h=40):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, w, h), 0)
    pix.clear_with(255)
    pix.set_rect(pymupdf.IRect(5, 5, w - 5, h - 5), color)
    return pix.tobytes("png")


# ---------------------------------------------------------------- 1. archivio: passaggio dal vecchio formato
pc1 = use_pc("pc1")
old = [{"id": "timbro1", "name": "Timbro", "width_mm": 44.0, "remove_white": True, "white_tolerance": 10,
        "overlay": True, "last_pos": [0.7, 0.8]},
       {"id": "abc123", "name": "Firma", "width_mm": 50.0, "remove_white": False, "white_tolerance": 10,
        "overlay": False, "last_pos": None,
        "source": {"kind": "typed", "text": "Ing. Rossi", "font": "Arial", "size": 20.0, "color": "#1f3a93"}}]
for item, color in zip(old, ((200, 30, 30), (30, 30, 200))):
    folder = tp.STAMPS_DIR / item["id"]
    folder.mkdir(parents=True)
    (folder / "originale.png").write_bytes(png(color))
    (folder / "timbro.png").write_bytes(png(color))
tp.CONFIG_FILE.write_text(json.dumps({"stamps": old, "current_stamp": "abc123", "last_dir": str(OUT)}),
                          encoding="utf-8")
root, app = new_app()
lib = app.library
archive = pc1 / tp.ARCHIVE_NAME
assert lib.available and archive.exists(), "l'archivio va creato"
assert [i["name"] for i in lib.items] == ["Timbro", "Firma"] and lib.current["id"] == "abc123"
assert lib.items[0]["last_pos"] == [0.7, 0.8] and lib.items[1]["source"]["text"] == "Ing. Rossi"
cfg = json.loads(tp.CONFIG_FILE.read_text(encoding="utf-8"))
assert "stamps" not in cfg and not tp.STAMPS_DIR.exists() and (pc1 / "timbri_versione_precedente").exists()
items, images = tp.read_archive(archive)
assert len(items) == 2 and "last_pos" not in items[0] and images["timbro1"][0].startswith(b"\x89PNG")
print("vecchi timbri trasferiti nell'archivio (un solo file), impostazioni e posizioni conservate: ok")

# un clic di timbro (posizione) non riscrive l'archivio; un timbro nuovo sì
before = archive.stat().st_mtime_ns
lib.items[0]["last_pos"] = [0.1, 0.2]
lib.save()
assert archive.stat().st_mtime_ns == before
lib.add(pymupdf.Pixmap(png((0, 150, 0))), "Verde", width_mm=30)
assert archive.stat().st_mtime_ns != before and len(tp.read_archive(archive)[0]) == 3
print("l'archivio si riscrive solo quando cambiano i timbri: ok")

# password: archivio cifrato, ricordata su questo PC (protezione di Windows)
lib.set_password("segreta123", remember=True)
assert pymupdf.open(archive).needs_pass, "cifrato"
try:
    tp.read_archive(archive, "sbagliata")
    raise AssertionError("con la password sbagliata non si deve aprire")
except PermissionError:
    pass
remembered = json.loads(tp.CONFIG_FILE.read_text(encoding="utf-8"))["archive_pw"]
assert "segreta123" not in remembered and tp.unprotect_secret(remembered) == "segreta123"
root.destroy()
root, app = new_app()
assert app.library.available and len(app.library.items) == 3, "riaperto senza chiedere la password"
root.destroy()
print("password: archivio cifrato AES-256, ricordata cifrata da Windows, riapertura automatica: ok")

# senza password ricordata: la si chiede; annullando l'archivio resta intatto
cfg = json.loads(tp.CONFIG_FILE.read_text(encoding="utf-8"))
cfg.pop("archive_pw")
tp.CONFIG_FILE.write_text(json.dumps(cfg), encoding="utf-8")
asked = []


class AutoPassword:
    answers = []

    def __init__(self, parent, message, confirm=False):
        asked.append(message)
        self.top = tk.Toplevel(parent)
        self.result = AutoPassword.answers.pop(0)
        parent.after(20, self.top.destroy)


tp.PasswordDialog = AutoPassword
AutoPassword.answers = [None]
data_before = archive.read_bytes()
root, app = new_app()
assert not app.library.available and asked and not app.library.items
pump(root, 0.5)
assert errors and "password" in errors[-1][2], errors
errors.clear()
app.library.save()
assert archive.read_bytes() == data_before, "un archivio non aperto non va sovrascritto"
grab(root, OUT / "archivio_chiuso.png")
AutoPassword.answers = [("sbagliata", False), ("segreta123", False)]
asked.clear()
app._on_gallery_double(type("E", (), {"x": 50, "y": 50})())
assert app.library.available and len(app.library.items) == 3 and len(asked) == 2
assert "errata" in asked[1]
root.destroy()
print("password chiesta (errata, poi giusta); annullando l'archivio non viene toccato: ok")

# ---------------------------------------------------------------- 2. archivio condiviso fra due PC
shared = OUT / "Dropbox" / "Timbri"
shared.mkdir(parents=True)
AutoPassword.answers = [("segreta123", True)]
root, app = new_app()
tp.filedialog.askdirectory = lambda **k: str(shared)
app._archive_move()
assert (shared / tp.ARCHIVE_NAME).exists() and not archive.exists()
assert json.loads(tp.CONFIG_FILE.read_text(encoding="utf-8"))["archive"] == str(shared / tp.ARCHIVE_NAME)
root.destroy()
print("archivio spostato nella cartella condivisa: ok")

pc2 = use_pc("pc2")                                  # secondo PC, con un suo timbro
root, app = new_app()
assert app.library.available and (pc2 / tp.ARCHIVE_NAME).exists(), "su un PC nuovo l'archivio si crea da sé"
app.library.add(pymupdf.Pixmap(png((150, 0, 150))), "Timbro PC2")
tp.filedialog.askopenfilename = lambda **k: str(shared / tp.ARCHIVE_NAME)
tp.messagebox.askyesnocancel = lambda *a, **k: True   # porta anche i timbri di questo PC
AutoPassword.answers = [("segreta123", True)]
app._archive_use()
names = [i["name"] for i in app.library.items]
assert names == ["Timbro", "Firma", "Verde", "Timbro PC2"], names
assert len(tp.read_archive(shared / tp.ARCHIVE_NAME, "segreta123")[0]) == 4
root.destroy()
print("secondo PC: usa l'archivio condiviso e vi aggiunge il suo timbro: ok")

use_pc("pc1")                                        # il primo PC vede la modifica tornando alla finestra
root, app = new_app()
assert len(app.library.items) == 4
items, images = tp.read_archive(shared / tp.ARCHIVE_NAME, "segreta123")
images["nuovo"] = (png((0, 0, 0)), png((0, 0, 0)))
items.append({"id": "nuovo", "name": "Dal PC2", "width_mm": 20.0})
time.sleep(0.05)
tp.write_archive(shared / tp.ARCHIVE_NAME, items, images, "segreta123")
app._archive_checked = 0
app._on_focus_in(None)
assert [i["name"] for i in app.library.items][-1] == "Dal PC2", "archivio condiviso riletto"
print("modifica fatta da un altro PC vista tornando alla finestra: ok")

# cartella dell'archivio non raggiungibile: nessun archivio vuoto al suo posto
cfg = json.loads(tp.CONFIG_FILE.read_text(encoding="utf-8"))
root.destroy()
cfg["archive"] = str(OUT / "Z_disco_assente" / tp.ARCHIVE_NAME)
tp.CONFIG_FILE.write_text(json.dumps(cfg), encoding="utf-8")
root, app = new_app()
pump(root, 0.5)
assert not app.library.available and not (OUT / "Z_disco_assente").exists()
errors.clear()
root.destroy()
cfg["archive"] = str(shared / tp.ARCHIVE_NAME)
tp.CONFIG_FILE.write_text(json.dumps(cfg), encoding="utf-8")
print("cartella non raggiungibile: segnalata, nessun archivio vuoto creato: ok")

# ---------------------------------------------------------------- 3. interfaccia
tp.PasswordDialog = AutoPassword
AutoPassword.answers = [("segreta123", True)]
use_pc("pc3")
root, app = new_app()
docs = []
for name, n in (("alfa", 3), ("beta", 2)):
    d = pymupdf.open()
    for i in range(n):
        d.new_page().insert_text((72, 100), f"{name} {i + 1}", fontsize=30)
    d.save(OUT / f"{name}.pdf")
    docs.append(OUT / f"{name}.pdf")

# file trascinati da Esplora risorse: messaggio WM_DROPFILES vero
kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32.PostMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
names = "".join(str(p) + "\0" for p in docs + [OUT / "note.txt"]) + "\0"
data = (20).to_bytes(4, "little") + bytes(12) + (1).to_bytes(4, "little") + names.encode("utf-16-le")
handle = kernel32.GlobalAlloc(0x42, len(data))
ctypes.memmove(kernel32.GlobalLock(handle), data, len(data))
kernel32.GlobalUnlock(handle)
shown = []
tp.messagebox.showinfo = lambda *a, **k: shown.append(a)
user32.PostMessageW(root.winfo_id(), 0x0233, handle, None)
pump(root, 0.5)
assert [t.name for t in app.tabs] == ["alfa.pdf", "beta.pdf"], [t.name for t in app.tabs]
assert shown and "note.txt" in shown[-1][1], "i file che non sono PDF vengono segnalati"
tp.messagebox.showinfo = lambda *a, **k: None
print("PDF trascinati da Esplora risorse aperti in nuove schede: ok")

# divisori ben visibili, galleria allungabile
assert app.paned.cget("bg") == tp.SASH_BG and int(app.paned.cget("sashwidth")) >= 6
h0 = int(app.gallery.cget("height"))
grip = app.gallery.master.winfo_children()[1]
gy = grip.winfo_rooty()
grip.event_generate("<Button-1>", x=5, y=2, rooty=gy)
grip.event_generate("<B1-Motion>", x=5, y=122, rooty=gy + 120, state=0x100)
grip.event_generate("<ButtonRelease-1>", x=5, y=122, rooty=gy + 120)
pump(root, 0.2)
assert int(app.gallery.cget("height")) == h0 + 120 and app.layout["gallery_h"] == h0 + 120
print(f"galleria dei timbri allungata trascinando la barretta: {h0} -> {h0 + 120} px: ok")

# scritte del pannello dei timbri: vanno a capo secondo la larghezza
help_label = app._stamp_sections[2]
wrap_before = int(help_label.cget("wraplength"))
panes = [str(p) for p in app.paned.panes()]
i = panes.index(str(app.side_panels["stamps"])) - 1
x, y = app.paned.sash_coord(i)
app.paned.sash_place(i, x - 150, y)
pump(root, 0.4)
assert int(help_label.cget("wraplength")) > wrap_before + 100, (wrap_before, help_label.cget("wraplength"))
print("pannello dei timbri più largo: le scritte si riadattano alla larghezza: ok")


def drag(widget, x_root, y_root):
    gx, gy = widget.winfo_rootx() + 4, widget.winfo_rooty() + 4
    widget.event_generate("<Button-1>", x=4, y=4, rootx=gx, rooty=gy)
    for k in range(1, 9):
        x, y = gx + (x_root - gx) * k // 8, gy + (y_root - gy) * k // 8
        widget.event_generate("<B1-Motion>", x=x - widget.winfo_rootx(), y=y - widget.winfo_rooty(),
                              rootx=x, rooty=y, state=0x100)
        pump(root)
    widget.event_generate("<ButtonRelease-1>", x=x_root - widget.winfo_rootx(), y=y_root - widget.winfo_rooty(),
                          rootx=x_root, rooty=y_root)
    pump(root, 0.3)


rx, ry, rw, rh = root.winfo_rootx(), root.winfo_rooty(), root.winfo_width(), root.winfo_height()
# barra trascinata verso il basso restando sul lato sinistro: va in basso, non a sinistra
text_bar = app.bars["text"]
drag(text_bar.grip, text_bar.grip.winfo_rootx() + 4, ry + rh - 20)
assert app.layout["bars"]["text"][0] == "bottom", app.layout["bars"]
# pannelli agganciati in basso: miniature in fila, timbri a colonne affiancate
drag(app.side_panels["thumbs"].winfo_children()[0], rx + 60, ry + rh - 30)
assert app.layout["panels"]["thumbs"][0] == "bottom" and app.thumbs_horizontal
boxes = app.thumb_layout
assert boxes[1][0] > boxes[0][2] and abs(boxes[1][1] - boxes[0][1]) < 30, "miniature in fila"
drag(app.side_panels["stamps"].winfo_children()[0], rx + rw // 2, ry + rh - 30)
assert app.layout["panels"]["stamps"][0] == "bottom" and app._stamps_horizontal
gallery_box, selected_box = app._stamp_sections[0], app._stamp_sections[1]
assert selected_box.winfo_rootx() > gallery_box.winfo_rootx() + gallery_box.winfo_width() - 5, "colonne affiancate"
assert str(app.side_panels["thumbs"]) in [str(p) for p in app.vpaned.panes()]
pump(root, 0.6)
grab(root, OUT / "pannelli_in_basso.png")
print("barra e pannelli agganciati in basso (miniature in fila, timbri a colonne): ok")
saved_layout = json.loads(json.dumps(app.layout))
app._save_layout()
root.destroy()
root, app = new_app()
assert app.layout["panels"] == saved_layout["panels"] and app.thumbs_horizontal and app._stamps_horizontal
app.reset_layout()
pump(root, 0.3)
assert not app.thumbs_horizontal and not app._stamps_horizontal
print("disposizione in basso ricordata alla riapertura e ripristinabile: ok")

# vista d'insieme: due colonne, una pagina copiata da alfa a beta trascinandola
for p in docs:
    app.open_file(str(p))
pump(root, 0.3)
app.toggle_overview(True)
pump(root, 0.8)
assert app.overview and not app.view_area.winfo_ismapped() and len(app.ov_layout) == 2
alfa_tab, beta_tab = app.tabs
_tab, _x0, _x1, boxes = app.ov_layout[0]
bx0, by0, bx1, by1, _s = boxes[2]                      # "alfa 3"
sx, sy = int((bx0 + bx1) / 2 - app.ov.canvasx(0)), int((by0 + by1) / 2 - app.ov.canvasy(0))
_tab, tx0, tx1, tboxes = app.ov_layout[1]
tx, ty = int((tx0 + tx1) / 2 - app.ov.canvasx(0)), int(tboxes[0][1] + 5 - app.ov.canvasy(0))   # prima di "beta 1"
ov = app.ov
ov.event_generate("<Button-1>", x=sx, y=sy, rootx=ov.winfo_rootx() + sx, rooty=ov.winfo_rooty() + sy)
for k in range(1, 9):
    x, y = sx + (tx - sx) * k // 8, sy + (ty - sy) * k // 8
    ov.event_generate("<B1-Motion>", x=x, y=y, rootx=ov.winfo_rootx() + x, rooty=ov.winfo_rooty() + y, state=0x100)
    pump(root)
ov.event_generate("<ButtonRelease-1>", x=tx, y=ty, rootx=ov.winfo_rootx() + tx, rooty=ov.winfo_rooty() + ty)
pump(root, 0.8)
assert [p.get_text().strip() for p in beta_tab.doc] == ["alfa 3", "beta 1", "beta 2"]
assert len(alfa_tab.doc) == 3, "dall'origine si copia"
grab(root, OUT / "vista_insieme.png")
_tab, _x0, _x1, boxes = app.ov_layout[1]
bx0, by0, bx1, by1, _s = boxes[2]
app._ov_double(type("E", (), {"x": int((bx0 + bx1) / 2 - ov.canvasx(0)), "y": int((by0 + by1) / 2 - ov.canvasy(0))})())
pump(root, 0.3)
assert not app.overview and app.tab is beta_tab and app.current == 2 and app.view_area.winfo_ismapped()
print("vista d'insieme: pagina copiata fra le colonne, doppio clic apre la pagina: ok")

# unione: scegliendo i documenti e l'ordine
dialog = tp.MergeDialog(root, app.tabs)
dialog._move(1, -1)                                  # beta prima di alfa
pump(root)
grab(dialog.top, OUT / "unisci.png")
dialog._ok()
assert dialog.result == [beta_tab, alfa_tab]
dialog = tp.MergeDialog(root, app.tabs)
dialog.rows[1][1].set(False)
dialog._ok()
assert dialog.result is None, "servono almeno due documenti"
dialog.top.destroy()
print("unione: scelta dei documenti e del loro ordine: ok")

# testo: con lo strumento «Aggiungi testo» attivo la casella si trascina; un clic la corregge
class FakeText:
    result_next = None

    def __init__(self, parent, fonts, title, params, original=None, whiteout=None):
        self.top = tk.Toplevel(parent)
        self.result = FakeText.result_next
        parent.after(20, self.top.destroy)


tp.TextDialog = FakeText
app.go_to(0)
app.set_zoom_mode("page")
pump(root, 0.3)


def at(px, py):
    pr = app.doc[app.current].rect
    ox, oy = app.page_origin
    return int(ox + px * app.cur_zoom - app.view.canvasx(0)), int(oy + py * app.cur_zoom - app.view.canvasy(0))


app.set_tool("add_text")
FakeText.result_next = {"text": "Nota", "font": "Arial", "size": 14, "color": "#000000", "whiteout": False}
x, y = at(100, 300)
app.view.event_generate("<ButtonPress-1>", x=x, y=y)
app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
pump(root, 0.2)
box = tp.placed_texts(app.doc[0])[0]
x0, y0 = at((box.rect.x0 + box.rect.x1) / 2, (box.rect.y0 + box.rect.y1) / 2)
x1, y1 = at((box.rect.x0 + box.rect.x1) / 2 + 150, (box.rect.y0 + box.rect.y1) / 2 + 100)
app.view.event_generate("<ButtonPress-1>", x=x0, y=y0)
for k in range(1, 9):
    app.view.event_generate("<B1-Motion>", x=x0 + (x1 - x0) * k // 8, y=y0 + (y1 - y0) * k // 8, state=0x100)
    pump(root)
app.view.event_generate("<ButtonRelease-1>", x=x1, y=y1)
pump(root, 0.2)
moved = tp.placed_texts(app.doc[0])
assert len(moved) == 1 and abs(moved[0].rect.x0 - 250) < 2 and abs(moved[0].rect.y0 - 400) < 2, moved
FakeText.result_next = {"text": "Nota corretta", "font": "Arial", "size": 14, "color": "#000000",
                        "whiteout": False}
app.view.event_generate("<ButtonPress-1>", x=x1, y=y1)
app.view.event_generate("<ButtonRelease-1>", x=x1, y=y1)
pump(root, 0.2)
assert tp.placed_texts(app.doc[0])[0].params["text"] == "Nota corretta"
assert app.tool == "add_text", "si resta nello strumento"
app.set_tool(None)
print("testo spostato trascinandolo anche con «Aggiungi testo» attivo; un clic lo corregge: ok")

# una sola finestra: un'altra apertura (doppio clic su un PDF) passa il file a quella già aperta
gamma = OUT / "gamma.pdf"
d = pymupdf.open()
d.new_page().insert_text((72, 100), "gamma 1", fontsize=30)
d.save(gamma)
assert (tp.CONFIG_DIR / tp.INSTANCE_FILE_NAME).exists()
assert tp.send_to_running_instance([str(gamma)]), "la finestra aperta deve ricevere il file"
pump(root, 0.8)
assert app.tab.name == "gamma.pdf", [t.name for t in app.tabs]
info = json.loads((tp.CONFIG_DIR / tp.INSTANCE_FILE_NAME).read_text(encoding="utf-8"))
import socket
with socket.create_connection(("127.0.0.1", info["port"])) as s:   # messaggio senza la chiave giusta: ignorato
    s.sendall(json.dumps({"token": "x", "files": [str(docs[0])]}).encode())
    s.shutdown(socket.SHUT_WR)
    assert s.recv(8) == b""
entries = tp.pdf_handler_entries()
command = next(v for k, n, v in entries if k.endswith(r"shell\open\command"))
assert command.endswith('"%1"') and "timbro_pdf.py" in command
assert (r"Software\Classes\.pdf\OpenWithProgids", tp.PROG_ID, "") in entries
assert all(k.startswith("Software\\") for k, _n, _v in entries), "solo HKEY_CURRENT_USER\\Software"
print("una sola finestra: i PDF aperti dopo arrivano come schede; voci per Apri con pronte: ok")

# licenza: finestra vera, indirizzo PayPal copiabile; al primo avvio va accettata
about = tp.AboutDialog(root)
pump(root)
grab(about.top, OUT / "licenza.png")
about._copy()
assert root.clipboard_get() == tp.DONATION_EMAIL
about._accept()
root.destroy()


class AutoAccept:
    def __init__(self, parent, must_accept=False):
        assert must_accept
        self.top = tk.Toplevel(parent)
        self.accepted = True
        parent.after(20, self.top.destroy)


tp.AboutDialog = AutoAccept
tp.LICENSE_VERSION = 1
use_pc("pc4")
root, app = new_app()
pump(root, 0.8)
assert app.cfg["license_accepted"] == 1, "condizioni accettate al primo avvio"
root.destroy()
print("licenza: finestra con autore, condizioni e donazione; accettazione al primo avvio: ok")

assert not errors, errors
print("TUTTI I TEST OK")
