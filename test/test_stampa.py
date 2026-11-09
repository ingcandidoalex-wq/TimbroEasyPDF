"""Test della stampa: enumerazione stampanti, resa GDI su «Microsoft Print to PDF», parsing
dell'intervallo di pagine del dialogo."""
import shutil
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402

tp.LICENSE_VERSION = 0
CFG = WORK / "cfg_stampa"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()

# ---------------------------------------------------------------- 1. enumerazione stampanti
default, printers = tp.list_printers()
assert isinstance(default, str) and isinstance(printers, list), "list_printers deve dare (str, list)"
print(f"1. stampanti: {len(printers)} trovate, predefinita = {default!r}")

# ---------------------------------------------------------------- 2. resa GDI su file (Print to PDF)
doc = pymupdf.open()
for n in range(4):
    p = doc.new_page(width=595, height=842)
    p.insert_text((60, 90), f"Pagina {n + 1}", fontsize=24)
    p.draw_rect(pymupdf.Rect(150, 300, 445, 500), color=(0, 0, 0.7), fill=(0, 0, 0.7))   # blocco blu pieno

PRINT_TO_PDF = "Microsoft Print to PDF"
if PRINT_TO_PDF in printers:
    out = WORK / "stampa_prova.pdf"
    out.unlink(missing_ok=True)
    seen = []
    ok = tp.print_document(PRINT_TO_PDF, doc, [0, 2, 3], copies=1,
                           on_progress=lambda n, t: seen.append(n) or True, output=str(out))
    for _ in range(60):
        if out.exists() and out.stat().st_size > 0:
            break
        time.sleep(0.2)
    assert ok, "print_document ha restituito False"
    assert out.exists() and out.stat().st_size > 1000, "il PDF stampato non è stato creato"
    res = pymupdf.open(out)
    assert len(res) == 3, f"attese 3 pagine stampate, trovate {len(res)}"
    assert seen == [0, 1, 2], f"avanzamento inatteso: {seen}"
    # il blocco blu non deve risultare rosso: nel centro il canale blu deve dominare
    pm = res[0].get_pixmap(dpi=60)
    r, g, b = pm.pixel(int(pm.width * 0.5), int(pm.height * 0.47))[:3]
    assert b > r + 50, f"colori invertiti nella stampa (r={r}, b={b})"
    print("2. resa GDI su PDF OK (3 pagine, colori corretti)")
else:
    print("2. «Microsoft Print to PDF» non disponibile: salto la prova di resa")

# ---------------------------------------------------------------- 3. parsing dell'intervallo pagine
root = tk.Tk()
root.withdraw()
dlg = tp.PrintDialog(root, current=0, selected=[0], total=10, printers=printers or ["(nessuna)"],
                     default=default)
dlg.range_var.set("1-3, 5, 9-10")
assert dlg._pages_from_range() == [0, 1, 2, 4, 8, 9], dlg._pages_from_range()
dlg.range_var.set("8-100")   # fuori intervallo: scartate quelle oltre il totale
assert dlg._pages_from_range() == [7, 8, 9], dlg._pages_from_range()
dlg.top.destroy()
root.destroy()
print("3. intervallo pagine OK")

print("TUTTI I TEST OK")
