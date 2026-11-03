"""Test della 2.2: rotazione di gradi qualsiasi, scorciatoie Ctrl, icona, email all'autore, aggiornamenti."""
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path
from urllib.parse import unquote

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402  (rende importabile anche il programma)
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402

tp.LICENSE_VERSION = 0
tp.tk.Menu.tk_popup = lambda self, *a, **k: None
CFG = WORK / "cfg_v22"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
errors, infos = [], []
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.showinfo = lambda *a, **k: infos.append(a[1] if len(a) > 1 else k.get("message"))
tp.messagebox.showerror = lambda *a, **k: errors.append(a)
OUT = WORK / "out_v22"
shutil.rmtree(OUT, ignore_errors=True)
OUT.mkdir()
tp.set_dpi_awareness()
RED, BLUE = (220, 30, 30), (30, 30, 220)


def image(w, h, fill):
    data = bytearray()
    for y in range(h):
        for x in range(w):
            data += bytes(fill(x, y))
    return pymupdf.Pixmap(pymupdf.csRGB, w, h, bytes(data), False)


def close(a, b, tol=45):
    return all(abs(u - v) <= tol for u, v in zip(a, b))


# ---------------------------------------------------------------- 1. rotazione di un angolo qualsiasi
assert tp.parse_angle("10") == 10 and tp.parse_angle("12,5°") == 12.5 and tp.parse_angle("0") == 90
assert tp.parse_angle("abc", None) is None and tp.parse_angle("400", None) is None
halves = image(80, 40, lambda x, y: RED if x < 40 else BLUE)
r30 = tp.rotate_pixmap(halves, 30)
assert r30.alpha and (r30.width, r30.height) == (89, 75), (r30.width, r30.height)
assert r30.pixel(1, 1)[3] == 0, "gli angoli restano trasparenti"
assert close(r30.pixel(r30.width // 2 - 15, r30.height // 2 - 10)[:3], RED), "la parte sinistra resta rossa"
mask = pymupdf.Pixmap(pymupdf.csGRAY, 80, 40, b"\xff" * 3200, False)
m30 = tp.rotate_pixmap(mask, 30)
assert m30.n == 1 and m30.pixel(1, 1)[0] == 0 and m30.pixel(m30.width // 2, m30.height // 2)[0] == 255
assert tp.rotate_pixmap(halves, 90).samples == tp.rotate_pixmap(halves, 450).samples, "90° esatti come prima"
print("rotazione di un angolo qualsiasi: si allarga, angoli trasparenti (anche per le maschere): ok")


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

halves.save(str(OUT / "meta.png"))
item = lib.add(str(OUT / "meta.png"), "Metà", width_mm=40, remove_white=False)
app._stamp_selected()
app.rotate_step_var.set("10")
assert app.rotate_step == 10 and app.cfg["rotate_step"] == 10, "i gradi si ricordano"
app.rotate_stamp(app.rotate_step)
st = app.stamp
assert item["rotate"] == 10 and st.pixmap.alpha and st.pixmap.width > 80
assert abs(item["width_mm"] - st.pixmap.width * 40 / 80) < 0.05, "stessa scala: l'ingombro cresce"
app.rotate_stamp(-app.rotate_step)
assert item["rotate"] == 0 and app.stamp.pixmap.width == 80 and abs(item["width_mm"] - 40) < 0.05
for _ in range(36):
    app.rotate_stamp(10)
assert item["rotate"] == 0 and abs(item["width_mm"] - 40) < 0.2, item
print("timbro della galleria ruotato di 10° avanti e indietro (e 36 volte = giro completo), misure conservate: ok")

# nell'editor: 15° con i gradi scritti accanto ai pulsanti


class AutoEditor(tp.StampEditor):
    action = None

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self.top.after(80, self._auto)

    def _auto(self):
        AutoEditor.action(self)
        self._save()


tp.StampEditor = AutoEditor


def turn15(ed):
    assert ed.step_var.get() == "10", "l'editor parte dai gradi del pannello"
    ed.step_var.set("15")
    ed.rotate(ed.rotate_step)
    AutoEditor.seen = (ed.colors.width, ed.mask is not None and ed.mask.pixel(0, 0)[0] == 0, ed.rotation)


AutoEditor.action = turn15
app.cmd_edit_stamp()
width, transparent_corner, rotation = AutoEditor.seen
assert width > 80 and transparent_corner and rotation == 15
assert item["rotate"] == 15 and app.rotate_step == 15, "solo ruotato: resta l'originale + 15°; i gradi tornano al pannello"
print("editor: rotazione di 15° scritti nel campo, angoli trasparenti, gradi ricordati: ok")
tp.StampEditor = AutoEditor.__mro__[1]

# timbro apposto ruotato di 10°
doc = pymupdf.open()
doc.new_page(width=595, height=842).insert_text((72, 100), "Pagina", fontsize=20)
doc.save(OUT / "pagina.pdf")
assert app.open_file(str(OUT / "pagina.pdf"))
app.set_zoom_mode("page")
pump(0.2)
app.rotate_stamp(-15)
app.set_tool("stamp")
ox, oy = app.page_origin
app._place_stamp(round(ox + 300 * app.cur_zoom - app.view.canvasx(0)), round(oy + 400 * app.cur_zoom - app.view.canvasy(0)))
app.set_tool(None)
stamp = tp.placed_stamps(app.doc[0])[-1]
center = ((stamp.rect.x0 + stamp.rect.x1) / 2, (stamp.rect.y0 + stamp.rect.y1) / 2)
app.rotate_placed(stamp, 10)
stamp = tp.placed_stamps(app.doc[0])[-1]
assert tp.stamp_view_angle(app.doc[0], stamp) == 10, tp.stamp_view_angle(app.doc[0], stamp)
assert abs((stamp.rect.x0 + stamp.rect.x1) / 2 - center[0]) < 0.5 and abs((stamp.rect.y0 + stamp.rect.y1) / 2 - center[1]) < 0.5
labels = []
real_popup = tp.tk.Menu.tk_popup
tp.tk.Menu.tk_popup = lambda self, *a, **k: labels.extend(self.entrycget(i, "label") for i in range(self.index("end") + 1)
                                                           if self.type(i) == "command")
x0, y0, x1, y1 = app._canvas_rect(stamp.rect)
app._on_view_menu(type("E", (), {"x": (x0 + x1) / 2 - app.view.canvasx(0), "y": (y0 + y1) / 2 - app.view.canvasy(0),
                                 "x_root": 0, "y_root": 0})())
tp.tk.Menu.tk_popup = real_popup
assert "Ruota a destra di 15°" in labels, labels
print("timbro apposto ruotato di 10° attorno al centro; il tasto destro propone i gradi del pannello: ok")

# ---------------------------------------------------------------- 2. scorciatoie Ctrl (anche con il Bloc Maiusc)
called = []
for name in ("cmd_open", "cmd_ocr", "cmd_extract", "cmd_merge_tabs", "cmd_split", "cmd_export_word", "cmd_add",
             "cmd_add_stamp", "cmd_create_signature", "cmd_edit_stamp", "repeat_last_stamp", "cmd_archive"):
    setattr(app, name, lambda *a, n=name, **k: called.append((n, k.get("at_end"))))
app._bind_keys()   # le scorciatoie si legano di nuovo, alle funzioni finte qui sopra


def key(keysym, shift=False):
    called.clear()
    app.view.focus_force()
    pump()
    app.view.event_generate(f"<Control-{keysym}>", state=0x4 | (0x1 if shift else 0))
    pump()
    return list(called)


assert key("o") == [("cmd_open", None)] and key("O") == [("cmd_open", None)], "Ctrl+O apre anche con il Bloc Maiusc"
assert key("O", shift=True) == [("cmd_ocr", None)], "Ctrl+Maiusc+O: OCR"
assert key("e") == [("cmd_extract", None)] and key("u") == [("cmd_merge_tabs", None)] and key("d") == [("cmd_split", None)]
assert key("W", shift=True) == [("cmd_export_word", None)]
assert key("i") == [("cmd_add", True)] and key("I", shift=True) == [("cmd_add", False)]
assert key("N", shift=True) == [("cmd_add_stamp", None)] and key("F", shift=True) == [("cmd_create_signature", None)]
assert key("M", shift=True) == [("cmd_edit_stamp", None)] and key("P", shift=True) == [("repeat_last_stamp", None)]
key("b")
assert app.tool == "whiteout"
key("B", shift=True)
assert app.tool == "blackout"
key("t")
assert app.tool == "add_text"
key("m")
assert app.tool == "edit_text"
app.set_tool(None)
app.set_zoom_mode("page")
app.view.event_generate("<Control-Key-2>", state=0x4)
pump()
assert app.zoom_mode == "width"
app.view.event_generate("<Control-Key-1>", state=0x4)
pump()
assert app.zoom_mode is None and app.zoom == 1.0
menubar = root.nametowidget(root["menu"])
accelerators = {}
for i in range(menubar.index("end") + 1):
    sub = root.nametowidget(menubar.entrycget(i, "menu"))
    for j in range(sub.index("end") + 1):
        if sub.type(j) in ("command", "checkbutton"):
            accelerators[sub.entrycget(j, "label")] = sub.entrycget(j, "accelerator")
without = [label for label, acc in accelerators.items() if not acc]
assert len(accelerators) - len(without) >= 40, (len(accelerators), without)
assert accelerators["Esporta in Word…"] == "Ctrl+Maiusc+W" and accelerators["Esci"] == "Ctrl+Q"
print(f"scorciatoie Ctrl: {len(accelerators) - len(without)} voci dei menu con la scorciatoia, Bloc Maiusc compreso: ok")

# ---------------------------------------------------------------- 3. icona, email, aggiornamenti
icon = tp.icon_pixmap(64)
band = icon.pixel(55, 12)   # la fascetta «PDF» sporge a destra del foglio (che finisce a x = 44)
assert band[0] > 150 and band[1] < 80 and band[3] == 255, band
assert icon.pixel(62, 40)[3] == 0, "fuori dalla fascetta e dal timbro resta trasparente"
icon.save(str(OUT / "icona.png"))

url = tp.feedback_mailto("Segnalazione di un problema", "Il bianchetto annerisce il testo")
assert url.startswith("mailto:ing.candido.alex@gmail.com?subject=") and " " not in url
assert "Segnalazione di un problema" in unquote(url) and f"TimbroEasyPDF {tp.APP_VERSION}" in unquote(url)
started = []
tp.os.startfile = lambda *a: started.append(a)
dialog = tp.FeedbackDialog(root)
dialog.text.insert("1.0", "Proposta: più timbri")
dialog._mail()
assert started and started[0][0].startswith("mailto:") and "Proposta" in unquote(started[0][0])
dialog._copy()
assert root.clipboard_get() == tp.FEEDBACK_EMAIL
dialog.top.destroy()
print("email all'autore: messaggio pronto con versione e Windows, indirizzo copiabile: ok")

assert tp.version_tuple("v2.10") > tp.version_tuple("2.9") > tp.version_tuple("2.2") == (2, 2)
root.after_cancel(app._update_job)   # il controllo automatico dell'avvio non deve sovrapporsi a queste prove
opened = []
import webbrowser  # noqa: E402
webbrowser.open = lambda u: opened.append(u)
tp.UPDATE_URL = "https://esempio.invalid/ultima"


def check(release=None, error=None, manual=True):
    infos.clear()
    opened.clear()

    def fake(timeout=8):
        if error:
            raise error
        return release
    tp.latest_release = fake
    app.cfg.data.pop("update_checked", None)
    app.check_updates(manual=manual)
    pump(1.5)


check(("2.3", "https://github.com/x/releases/tag/v2.3", "Novità varie"))
assert opened == ["https://github.com/x/releases/tag/v2.3"], opened
check(("2.2", "u", ""))
assert opened == [] and "ultima versione" in infos[0], infos
import urllib.error  # noqa: E402
check(error=urllib.error.HTTPError("u", 404, "Not Found", None, None))
assert "Non ci sono ancora versioni" in infos[0], infos
check(error=OSError("rete assente"), manual=False)
assert infos == [] and opened == [], "all'avvio, senza rete, nessun messaggio"
check(("9.0", "u", ""), manual=False)
assert opened == ["u"] and app.cfg["update_checked"] == time.strftime("%Y-%m-%d")
opened.clear()
app.check_updates()          # già controllato oggi: non si ricontrolla
pump(0.5)
assert opened == []
app.auto_update_var.set(False)
app.cfg["check_updates"] = False
app.cfg.data.pop("update_checked", None)
app.check_updates()
pump(0.5)
assert opened == [], "controllo all'avvio disattivato dal menu"
print("aggiornamenti: versione nuova -> pagina di download; già aggiornato; nessuna versione; senza rete; "
      "una volta al giorno; disattivabile: ok")

assert not errors, errors
root.destroy()
print("TUTTI I TEST OK")
