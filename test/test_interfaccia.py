"""Test dell'interfaccia 2.4: barre/pannello a scomparsa (persistenti), cambio tema, menù contestuale
adattivo, formato del nuovo documento."""
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
CFG = WORK / "cfg_interfaccia"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()


def new_app():
    root = tk.Tk()
    app = tp.App(root)
    root.geometry("1500x900+20+20")
    root.update()
    return root, app


def labels(menu):
    return [menu.entrycget(i, "label") for i in range(menu.index("end") + 1) if menu.type(i) == "command"]


# ---------------------------------------------------------------- 1. visibilità: default chiusi, persistente
root, app = new_app()
assert not app._is_visible("text") and not app._is_visible("pages") and not app._is_visible("stamps"), \
    "all'avvio i gruppi sono chiusi"
assert not app.bars["text"].winfo_ismapped() and not app.side_panels["stamps"].winfo_ismapped()
app._set_visible("text", True)
app._set_visible("stamps", True)
root.update()
assert app.bars["text"].winfo_ismapped() and app.side_panels["stamps"].winfo_ismapped(), "aperti"
root.destroy()

root, app = new_app()                                   # riapertura: lo stato è ricordato
assert app._is_visible("text") and app._is_visible("stamps") and not app._is_visible("pages")
print("1. barre/pannello a scomparsa, default chiusi e stato ricordato: OK")

# ---------------------------------------------------------------- 2. cambio tema
start = app._style.theme_use()
other = next((t for t in ("clam", "classic", "alt", "default") if t in app._style.theme_names() and t != start), None)
if other:
    app.set_theme(other)
    assert app._style.theme_use() == other and app.cfg["theme"] == other, "tema applicato e salvato"
    app.set_theme(start)
print("2. cambio tema grafico OK")

# ---------------------------------------------------------------- 3. menù contestuale adattivo
m = tk.Menu(root, tearoff=False)
app._fill_context_menu(m, type("E", (), {"x": 100, "y": 100})())
assert "Apri un PDF…" in labels(m) and "Nuovo documento…" in labels(m), labels(m)   # nessun documento

doc = pymupdf.open()
doc.new_page(width=595, height=842)
doc[0].insert_text((70, 120), "Relazione tecnica di calcolo", fontsize=20)
p = WORK / "ctx.pdf"
doc.save(p)
app.open_file(str(p))
root.update()
m = tk.Menu(root, tearoff=False)
app._fill_context_menu(m, type("E", (), {"x": 100, "y": 100})())
voci = labels(m)
assert "Aggiungi testo qui" in voci and "Stampa…" in voci and "Ruota la pagina a destra" in voci, voci
assert "Copia il testo selezionato" not in voci, "niente copia senza selezione"
app._clip_text = "Relazione"
m = tk.Menu(root, tearoff=False)
app._fill_context_menu(m, type("E", (), {"x": 100, "y": 100})())
assert "Copia il testo selezionato" in labels(m), labels(m)
print("3. menù contestuale cambia secondo lo stato OK")

# ---------------------------------------------------------------- 4. nuovo documento con formato scelto
app.cmd_new("A3", True)
page = app.doc[0]
assert len(app.doc) == 1 and abs(page.rect.width - 1191) < 2 and abs(page.rect.height - 842) < 2, \
    (page.rect.width, page.rect.height)
assert app.cfg["new_format"] == "A3" and app.cfg["new_landscape"] is True
print("4. nuovo documento A3 orizzontale OK")

root.destroy()
print("TUTTI I TEST OK")
