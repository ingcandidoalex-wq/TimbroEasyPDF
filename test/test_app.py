"""Test end-to-end di timbro_pdf: libreria timbri, sfondo, primo/secondo piano, pagine, salvataggio."""
import ctypes
import ctypes.wintypes as wt
import json
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

# configurazione isolata: non tocca i timbri reali dell'utente
CFG = WORK / "cfg"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.showinfo = lambda *a, **k: None

OUT = WORK / "out"
OUT.mkdir(exist_ok=True)


def make_doc(path, specs):
    doc = pymupdf.open()
    for w, h, rot, label in specs:
        p = doc.new_page(width=w, height=h)
        p.draw_rect(pymupdf.Rect(30, 30, w - 30, h - 30), color=(0, 0, 0.8))
        p.insert_text((50, 80), label, fontsize=26)
        p.draw_rect(pymupdf.Rect(w - 230, h - 130, w - 30, h - 30), color=(0, 0, 0))
        p.insert_text((w - 220, h - 108), "Il tecnico", fontsize=11)
        if rot:
            p.set_rotation(rot)
    doc.save(path)


def make_stamp_image(path, text, color, transparent_bg=False):
    d = pymupdf.open()
    p = d.new_page(width=200, height=120)
    if transparent_bg:
        # PNG con sfondo trasparente ma un riquadro bianco interno
        p.draw_rect(pymupdf.Rect(70, 40, 130, 80), color=None, fill=(1, 1, 1))
    p.draw_oval(pymupdf.Rect(6, 6, 194, 114), color=color, width=4)
    p.insert_text((38, 55), text, fontsize=20, color=color)
    p.insert_text((52, 82), "PROVA", fontsize=20, color=color)
    pix = p.get_pixmap(dpi=300, alpha=transparent_bg)
    if path.suffix == ".jpg":
        pix.save(str(path), jpg_quality=90)
    else:
        pix.save(str(path))


main_pdf, extra_pdf = OUT / "main.pdf", OUT / "extra.pdf"
make_doc(main_pdf, [(595, 842, 0, "Pagina A"), (842, 595, 0, "Pagina B orizz."), (595, 842, 90, "Pagina C rot90")])
make_doc(extra_pdf, [(595, 842, 0, "Extra 1"), (595, 842, 0, "Extra 2")])
blue_jpg, red_png = OUT / "timbro_blu.jpg", OUT / "firma_rossa.png"
make_stamp_image(blue_jpg, "TIMBRO DI", (0.1, 0.2, 0.7))
make_stamp_image(red_png, "FIRMA DI", (0.75, 0.1, 0.1), transparent_bg=True)

# timbro salvato dalla prima versione (un solo file + larghezza in config)
make_stamp_image(tp.OLD_STAMP_FILE, "VECCHIO", (0.1, 0.5, 0.1))
tp.CONFIG_FILE.write_text(json.dumps({"stamp_width_mm": 35.0}), encoding="utf-8")


def grab(root, path):
    """Cattura la finestra con PrintWindow (funziona anche se coperta)."""
    root.update()
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    for f in (user32.GetDC, gdi32.CreateCompatibleDC, gdi32.CreateCompatibleBitmap, gdi32.SelectObject):
        f.restype = ctypes.c_void_p
    user32.GetDC.argtypes = [ctypes.c_void_p]
    gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
    gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
    gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    user32.PrintWindow.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
    gdi32.GetDIBits.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint,
                                ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
    hwnd = int(root.wm_frame(), 16)
    rect = wt.RECT()
    user32.GetWindowRect(ctypes.c_void_p(hwnd), ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    hdc = user32.GetDC(None)
    mdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(mdc, bmp)
    user32.PrintWindow(ctypes.c_void_p(hwnd), mdc, 2)

    class BMI(ctypes.Structure):
        _fields_ = [("biSize", ctypes.c_uint32), ("biWidth", ctypes.c_int32), ("biHeight", ctypes.c_int32),
                    ("biPlanes", ctypes.c_uint16), ("biBitCount", ctypes.c_uint16),
                    ("biCompression", ctypes.c_uint32), ("rest", ctypes.c_uint32 * 6)]

    bmi = BMI(40, w, -h, 1, 32, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bmi), 0)
    raw = buf.raw
    rgb = bytearray(w * h * 3)
    rgb[0::3], rgb[1::3], rgb[2::3] = raw[2::4], raw[1::4], raw[0::4]
    pymupdf.Pixmap(pymupdf.csRGB, w, h, bytes(rgb), False).save(str(path))


def alpha_at(pix, fx, fy):
    """Opacità del pixel a frazione (fx, fy) dell'immagine."""
    x, y = int(fx * (pix.width - 1)), int(fy * (pix.height - 1))
    return pix.pixel(x, y)[3] if pix.alpha else 255


def transparent_share(pix):
    a = pix.samples[pix.n - 1::pix.n]
    return a.count(0) / len(a)


def new_app():
    root = tk.Tk()
    app = tp.App(root)
    root.geometry("1360x900+30+30")
    root.update()
    return root, app


def add_stamp(app, path, name):
    tp.filedialog.askopenfilename = lambda **k: str(path)
    tp.simpledialog.askstring = lambda *a, **k: name
    app.cmd_add_stamp()


tp.set_dpi_awareness()

# ---------------------------------------------------------------- 1. libreria
root, app = new_app()
assert [i["name"] for i in app.library.items] == ["Timbro"], app.library.items
assert app.stamp_item["width_mm"] == 35.0 and app.stamp_w_var.get() == "35"
assert not tp.OLD_STAMP_FILE.exists(), "il vecchio timbro deve essere trasferito nella libreria"
print("timbro della prima versione importato: ok")

add_stamp(app, blue_jpg, "Timbro blu")
add_stamp(app, red_png, "Firma rossa")
assert [i["name"] for i in app.library.items] == ["Timbro", "Timbro blu", "Firma rossa"]
assert app.stamp_item["name"] == "Firma rossa"
assert len(app.gallery.find_all()) > 3

# sfondo bianco su PNG già trasparente: esterno resta trasparente, riquadro bianco interno diventa trasparente
pix = app.stamp.pixmap
assert alpha_at(pix, 0.02, 0.02) == 0 and alpha_at(pix, 0.5, 0.5) == 0, "PNG: bianco interno da togliere"
app.remove_white_var.set(False)
app._on_white_settings()
root.update()
app._apply_white_settings() if app._white_job else None
pix = app.stamp.pixmap
assert alpha_at(pix, 0.02, 0.02) == 0 and alpha_at(pix, 0.5, 0.5) == 255, "senza rimozione: riquadro bianco opaco"
app.remove_white_var.set(True)
app._apply_white_settings()

# JPG: con/senza rimozione e tolleranza
app.select_stamp(1)
assert app.stamp_item["name"] == "Timbro blu"
assert app.stamp.pixmap.alpha and alpha_at(app.stamp.pixmap, 0.02, 0.02) == 0
assert alpha_at(app.stamp.pixmap, 0.5, 0.95) > 200, "il bordo blu deve restare opaco"
share_default = transparent_share(app.stamp.pixmap)
app.tolerance_scale.set(80)
app._apply_white_settings()
share_high = transparent_share(app.stamp.pixmap)
assert share_high >= share_default, (share_default, share_high)
app.remove_white_var.set(False)
app._apply_white_settings()
assert not app.stamp.pixmap.alpha, "senza rimozione il JPG resta opaco"
app.remove_white_var.set(True)
app.tolerance_scale.set(10)
app._apply_white_settings()
print(f"sfondo bianco: trasparente {share_default:.0%} (tolleranza 10) / {share_high:.0%} (tolleranza 80): ok")

# impostazioni diverse per timbro
app.stamp_w_var.set("50")
app.select_stamp(2)
app.stamp_w_var.set("30")
app.overlay_var.set(False)
app._on_overlay()
app.select_stamp(1)
assert app.stamp_w_var.get() == "50" and app.overlay_var.get() is True
app.select_stamp(2)
assert app.stamp_w_var.get() == "30" and app.overlay_var.get() is False
app.select_stamp(1)
print("impostazioni separate per ogni timbro: ok")

# ---------------------------------------------------------------- 2. pagine
assert app.open_file(str(main_pdf))
root.update()
tp.filedialog.askopenfilenames = lambda **k: (str(extra_pdf),)
app.cmd_add(at_end=True)


def order():
    return [app.doc[i].get_text().split("\n")[0] for i in range(len(app.doc))]


assert order() == ["Pagina A", "Pagina B orizz.", "Pagina C rot90", "Extra 1", "Extra 2"], order()
app.go_to(1)
app.rotate(90)
assert app.doc[1].rotation == 90
app.go_to(4)
app.move_pages(-1)
assert order() == ["Pagina A", "Pagina B orizz.", "Pagina C rot90", "Extra 2", "Extra 1"], order()
app.selection = {0, 1}
app.move_pages(1)
app.undo()
assert order() == ["Pagina A", "Pagina B orizz.", "Pagina C rot90", "Extra 2", "Extra 1"], order()
print("unione/rotazione/spostamento/annulla: ok")

# ---------------------------------------------------------------- 3. timbri con clic
app.set_zoom_mode("page")
app.set_stamp_mode(True)
expected = {}


def click(i, fx, fy):
    app.go_to(i)
    root.update()
    pr = app.doc[i].rect
    ox, oy = app.page_origin
    vx, vy = fx * pr.width, fy * pr.height
    x = int(ox + vx * app.cur_zoom - app.view.canvasx(0))
    y = int(oy + vy * app.cur_zoom - app.view.canvasy(0))
    app.view.event_generate("<Motion>", x=x, y=y)
    root.update()
    assert app.view.find_withtag("ghost"), "anteprima del timbro assente"
    app.view.event_generate("<ButtonPress-1>", x=x, y=y)
    app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
    root.update()
    return vx, vy


for i in (0, 1, 2):                       # timbro blu, primo piano, 50 mm
    expected[i] = click(i, 0.80, 0.85)
app.select_stamp(2)                       # firma rossa, secondo piano, 30 mm
click(0, 0.30, 0.85)
assert len(app.doc[0].get_image_info()) == 2
grab(root, OUT / "screenshot_libreria.png")

app.select_stamp(1)
app.selection = {3, 4}
app.repeat_last_stamp()
assert all(len(app.doc[i].get_image_info()) == 1 for i in (3, 4))
app.undo()
assert all(len(app.doc[i].get_image_info()) == 0 for i in (3, 4))
app.redo()
print("timbri con clic, ripeti, annulla/ripristina: ok")

out_pdf = OUT / "risultato.pdf"
tp.filedialog.asksaveasfilename = lambda **k: str(out_pdf)
assert app.cmd_save()
res = pymupdf.open(out_pdf)
xrefs = {img[0] for pg in res for img in pg.get_images()}
assert len(xrefs) == 2, f"un'immagine per timbro, trovate {len(xrefs)}"
for i, (vx, vy) in expected.items():
    pg = res[i]
    info = [im for im in pg.get_image_info() if abs(pymupdf.Rect(im["bbox"]).width - 50 * tp.MM) < 2
            or abs(pymupdf.Rect(im["bbox"]).height - 50 * tp.MM) < 2]
    r = pymupdf.Rect(info[0]["bbox"]) * pg.rotation_matrix
    r.normalize()
    assert abs((r.x0 + r.x1) / 2 - vx) < 2 and abs((r.y0 + r.y1) / 2 - vy) < 2, (i, r, vx, vy)

# primo/secondo piano: il timbro in secondo piano è disegnato prima del contenuto della pagina
contents = res[0].get_contents()
streams = [res.xref_stream(x) for x in contents]
first_do = next(k for k, s in enumerate(streams) if b" Do" in s)
assert first_do == 0, "la firma in secondo piano deve stare nel primo stream"
assert b" Do" in streams[-1], "il timbro in primo piano deve stare nell'ultimo stream"
print("salvataggio, posizioni, primo/secondo piano: ok")

res[0].get_pixmap(dpi=60).save(str(OUT / "pagina1.png"))
root.destroy()

# ---------------------------------------------------------------- 4. riavvio: tutto ricordato
root, app = new_app()
assert app.stamp_item["name"] == "Timbro blu", "all'avvio deve essere selezionato l'ultimo timbro usato"
assert app.stamp_w_var.get() == "50"
assert [i["name"] for i in app.library.items] == ["Timbro", "Timbro blu", "Firma rossa"]
assert app.library.items[2]["width_mm"] == 30 and app.library.items[2]["overlay"] is False
make_doc(OUT / "nuovo.pdf", [(595, 842, 0, "Nuovo 1"), (595, 842, 0, "Nuovo 2")])
app.open_file(str(OUT / "nuovo.pdf"))
assert str(app.repeat_btn.cget("state")) == "normal"
app.repeat_last_stamp()
assert all(len(app.doc[i].get_image_info()) == 1 for i in range(2)), "ripeti con la posizione dell'ultima volta"
app.select_stamp(0)
app.cmd_delete_stamp()
assert [i["name"] for i in app.library.items] == ["Timbro blu", "Firma rossa"]
print("riavvio: ultimo timbro, impostazioni e posizione ricordati; eliminazione: ok")
root.destroy()
print("TUTTI I TEST OK")
