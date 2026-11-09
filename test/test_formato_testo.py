"""Test della formattazione del testo: grassetto/corsivo (variante del font) e sottolineato
(riga disegnata nello stesso stream del testo, che resta selezionabile)."""
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
CFG = WORK / "cfg_formato"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()

# ---------------------------------------------------------------- 1. famiglia e variante
fonts = {"Arial": "a.ttf", "Arial Bold": "ab.ttf", "Arial Italic": "ai.ttf", "Arial Bold Italic": "abi.ttf",
         "Calibri": "c.ttf"}
assert tp.font_base("Arial Bold Italic") == "Arial"
assert tp.font_base("Calibri") == "Calibri"
assert tp.resolve_variant(fonts, "Arial", True, False) == "Arial Bold"
assert tp.resolve_variant(fonts, "Arial", True, True) == "Arial Bold Italic"
assert tp.resolve_variant(fonts, "Arial", False, True) == "Arial Italic"
assert tp.resolve_variant(fonts, "Calibri", True, False) is None   # variante non installata
print("1. famiglia/variante OK")

# ---------------------------------------------------------------- 2. sottolineatura nello stream del testo
doc = pymupdf.open()
page = doc.new_page(width=400, height=200)
ff = tp.installed_fonts().get("Arial")
before = set(page.get_contents())
params = {"text": "Riga sottolineata", "size": 20, "color": "#000000", "origin": [40, 60],
          "rotate": 0, "underline": True}
xref = tp.insert_text_box(page, params, ff, "F0")
assert "sottolineata" in page.get_text(), "il testo deve restare estraibile/selezionabile"
# la casella di testo (testo + riga) è un unico stream marcato
boxes = tp.placed_texts(page)
assert len(boxes) == 1 and boxes[0].params.get("underline") is True, "underline non salvato nei parametri"
# nello stream ci sono sia il testo (BT) sia la riga (operatore 'l' di line-to)
stream = doc.xref_stream(xref)
assert b"BT" in stream and b" l\n" in stream.replace(b"\r", b"\n") or b" l " in stream, "manca la riga nello stream"
print("2. sottolineatura nello stesso stream OK")

# ---------------------------------------------------------------- 3. dialogo: G riseleziona la variante
root = tk.Tk()
root.withdraw()
allfonts = tp.installed_fonts()
start = "Arial" if "Arial" in allfonts else sorted(allfonts)[0]
dlg = tp.TextDialog(root, allfonts, "Testo", {"text": "ciao", "font": start, "size": 14, "color": "#000000"},
                    preview=None)
dlg.bold_var.set(True)
dlg._apply_variant()
if "Arial Bold" in allfonts:
    assert dlg.font_var.get() == "Arial Bold", dlg.font_var.get()
dlg.underline_var.set(True)
dlg.text.delete("1.0", "end")
dlg.text.insert("1.0", "testo in grassetto")
dlg.size_var.set("18")
dlg._ok()
assert dlg.result and dlg.result["underline"] is True, "il risultato deve avere underline"
assert dlg.result["text"] == "testo in grassetto"
root.destroy()
print("3. pulsanti G/S del dialogo OK")

print("TUTTI I TEST OK")
