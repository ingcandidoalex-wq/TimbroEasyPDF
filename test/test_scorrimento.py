"""Test della modalità di scorrimento continuo: persistenza, layout, sincronizzazione della pagina
corrente con lo scroll, passaggio automatico a pagina singola quando si modifica, niente salto di
pagina con la rotella."""
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
CFG = WORK / "cfg_scroll"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.messagebox.askyesno = lambda *a, **k: True
tp.messagebox.showinfo = lambda *a, **k: None
tp.messagebox.showwarning = lambda *a, **k: None
tp.set_dpi_awareness()

OUT = WORK / "out_scroll"
OUT.mkdir(exist_ok=True)
DOC = OUT / "multi.pdf"
doc = pymupdf.open()
for n, (w, h, rot) in enumerate([(595, 842, 0), (842, 595, 0), (595, 842, 90), (595, 842, 0), (595, 842, 0)]):
    p = doc.new_page(width=w, height=h)
    p.insert_text((60, 90), f"Pagina {n + 1}", fontsize=28)
    if rot:
        p.set_rotation(rot)
doc.save(DOC)


def new_app():
    root = tk.Tk()
    app = tp.App(root)
    root.geometry("1200x820+30+30")
    root.update()
    return root, app


def pump(root, n=6):
    for _ in range(n):
        root.update()
        root.after(10, root.quit)
        root.mainloop()


def to_continuous(app, root):
    app.scroll_mode_var.set("continuous")
    app._apply_scroll_mode()
    pump(root)


# ---------------------------------------------------------------- 1. persistenza dell'impostazione
root, app = new_app()
app.open_file(str(DOC))
root.update()
assert app.scroll_mode == "single", "il default deve essere pagina singola"
to_continuous(app, root)
assert app.cfg["scroll_mode"] == "continuous", "la scelta non è stata salvata in config"
root.destroy()

root, app = new_app()                      # nuova App: deve ricordare la scelta
assert app.scroll_mode == "continuous", "la modalità non è stata ricordata al riavvio"
print("1. persistenza OK")

# ---------------------------------------------------------------- 2. layout di tutte le pagine
app.open_file(str(DOC))
pump(root)
assert app._continuous_active(), "il continuo dovrebbe essere attivo"
assert len(app.page_tops) == len(app.doc) == 5, f"page_tops: {len(app.page_tops)}"
assert all(b > a for a, b in zip(app.page_tops, app.page_tops[1:])), "page_tops non crescenti"
alt_tot = sum(app.page_layout[i][3] for i in range(5)) + tp.CONT_GAP * 4 + 2 * tp.PAD
assert abs(app.scroll_size[1] - alt_tot) < max(app.view.winfo_height(), alt_tot) * 0.02 \
    or app.scroll_size[1] >= alt_tot - 1, f"altezza scrollregion {app.scroll_size[1]} vs {alt_tot}"
print("2. layout continuo OK")

# ---------------------------------------------------------------- 3. la corrente segue lo scroll (>50%)
target = 3  # pagina indice 3 (quarta)
_ox, oy, _w, h = app.page_layout[target]
sh = app.scroll_size[1]
vh = app.view.winfo_height()
app.view.yview_moveto(max((oy + h / 2 - vh / 2) / sh, 0))   # centra la pagina target
pump(root)
assert app.current == target, f"la corrente doveva essere {target}, è {app.current}"
assert app.page_var.get() == str(target + 1), f"casella pagina: {app.page_var.get()}"
assert app.thumbs.itemcget(f"bg{target}", "outline") == "#1a73e8", "miniatura corrente non evidenziata"
print("3. sincronizzazione scroll->corrente OK")

# ---------------------------------------------------------------- 4. modifica -> pagina singola -> continuo
app.set_tool("whiteout")
pump(root)
assert not app._continuous_active(), "con uno strumento attivo il continuo non deve essere attivo"
assert app.current == target, "la modifica deve partire sulla pagina guardata"
assert isinstance(app.page_origin, tuple), "page_origin non impostato in pagina singola"
app.set_tool(None)
pump(root)
assert app._continuous_active(), "uscendo dallo strumento il continuo deve riprendere"
print("4. modifica -> singola -> continuo OK")

# ---------------------------------------------------------------- 5. la rotella non salta pagina nel continuo
before = app._last_flip
app.view.event_generate("<MouseWheel>", delta=-120, x=400, y=300)
pump(root)
assert app._last_flip == before, "nel continuo la rotella non deve fare il salto di pagina (flip)"
print("5. rotella senza salto di pagina OK")

# ---------------------------------------------------------------- 6. doppio clic su timbro -> singola -> continuo
app.view.yview_moveto(0)
pump(root)


class Ev:
    def __init__(self, x, y):
        self.x, self.y = x, y


app._on_double_click(Ev(60, int(app.page_layout[0][1] + 10)))   # area vuota: resta nel continuo
pump(root)
assert app._continuous_active(), "il doppio clic sul vuoto non deve cambiare modalità"
print("6. doppio clic OK")

root.destroy()
print("TUTTI I TEST OK")
