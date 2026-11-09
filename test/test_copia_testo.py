"""Test della selezione e copia del testo del PDF: parola sotto il cursore, selezione a rettangolo,
copia negli appunti con Ctrl+C."""
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
CFG = WORK / "cfg_copia"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()

OUT = WORK / "out_copia"
OUT.mkdir(exist_ok=True)
DOC = OUT / "testo.pdf"
doc = pymupdf.open()
page = doc.new_page(width=595, height=842)
page.insert_text((70, 120), "Relazione tecnica di calcolo", fontsize=20)
page.insert_text((70, 160), "strutturale per il progetto", fontsize=20)
doc.save(DOC)


def pump(root, n=5):
    for _ in range(n):
        root.update()
        root.after(10, root.quit)
        root.mainloop()


root = tk.Tk()
app = tp.App(root)
root.geometry("900x1000+20+20")
root.update()
app.open_file(str(DOC))
pump(root)

page = app.doc[0]
words = page.get_text("words")
by_text = {w[4]: w for w in words}


def widget_xy(word):
    r = pymupdf.Rect(word[:4]) * page.rotation_matrix
    x0, y0, x1, y1 = app._canvas_rect(r)
    return (x0 + x1) / 2 - app.view.canvasx(0), (y0 + y1) / 2 - app.view.canvasy(0)


# ---------------------------------------------------------------- 1. parola sotto il cursore
wx, wy = widget_xy(by_text["tecnica"])
assert app._word_at(wx, wy)[4] == "tecnica", app._word_at(wx, wy)
print("1. parola sotto il cursore OK")

# ---------------------------------------------------------------- 2. doppio clic: una parola + copia
assert app._select_word(wx, wy) is True
assert app._clip_text == "tecnica", repr(app._clip_text)
app.copy_selection()
assert root.clipboard_get() == "tecnica", repr(root.clipboard_get())
print("2. selezione parola e copia OK")

# ---------------------------------------------------------------- 3. selezione a rettangolo su più parole
x0, y0 = widget_xy(by_text["Relazione"])
x1, y1 = widget_xy(by_text["calcolo"])
app._text_sel = {"start": (x0, y0)}
app._update_text_selection(x1, y1)
assert app._clip_text == "Relazione tecnica di calcolo", repr(app._clip_text)
app.copy_selection()
assert root.clipboard_get() == "Relazione tecnica di calcolo"
print("3. selezione a rettangolo e copia OK")

# ---------------------------------------------------------------- 4. la selezione si azzera cambiando pagina/strumento
app.set_tool("add_text")
assert app._clip_text == "", "cambiando strumento la selezione deve azzerarsi"
app.set_tool(None)
print("4. azzeramento selezione OK")

root.destroy()
print("TUTTI I TEST OK")
