"""Test della compressione PDF (stile PDF24): motore compress_pdf a livelli, dialogo e comando dell'app.

Crea un PDF con dentro un'immagine ad alta risoluzione (rumore, poco comprimibile) e verifica che i
livelli «media» e «massima» riducano davvero la dimensione abbassando la risoluzione delle immagini."""
import os
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402

tp.LICENSE_VERSION = 0
CFG = WORK / "cfg_compress"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()


def pdf_con_immagine(w=1700, h=2400):
    """PDF A4 con un'immagine di rumore (incomprimibile) a piena pagina, ad alta risoluzione (~205 DPI)."""
    pix = pymupdf.Pixmap(pymupdf.csRGB, w, h, os.urandom(w * h * 3), 0)
    d = pymupdf.open()
    page = d.new_page(width=595, height=842)
    page.insert_image(page.rect, pixmap=pix)
    data = d.tobytes()
    d.close()
    return data


# ---------------------------------------------------------------- 1. human_size
assert tp.human_size(0) == "0 B"
assert tp.human_size(1500).endswith("KB")
assert tp.human_size(3_000_000).endswith("MB")
print("1. human_size OK")

# ---------------------------------------------------------------- 2. compress_pdf: preset e DPI a mano
data = pdf_con_immagine()
originale = len(data)
out_leg, out_med, out_max, out_dpi = CFG / "leg.pdf", CFG / "med.pdf", CFG / "max.pdf", CFG / "dpi.pdf"
b1, a_leg = tp.compress_pdf(data, str(out_leg), img_dpi=0, text_dpi=0)                 # leggera
b2, a_med = tp.compress_pdf(data, str(out_med), img_dpi=150, text_dpi=0)               # media
b3, a_max = tp.compress_pdf(data, str(out_max), img_dpi=96, text_dpi=200, quality=55)  # massima
b4, a_dpi = tp.compress_pdf(data, str(out_dpi), img_dpi=72, text_dpi=0)                # DPI immagini a mano
assert b1 == b2 == b3 == b4 == originale, "la dimensione di partenza deve essere quella del documento"
for o in (out_leg, out_med, out_max, out_dpi):
    assert len(pymupdf.open(o)) == 1, f"{o.name}: attesa 1 pagina"
assert a_med < a_leg, f"150 DPI ({a_med}) doveva battere la sola ricompattazione ({a_leg})"
assert a_max < a_med, f"96 DPI ({a_max}) doveva battere 150 DPI ({a_med})"
assert a_dpi < a_med, f"72 DPI a mano ({a_dpi}) doveva battere 150 DPI ({a_med})"
print(f"2. compress_pdf OK  (orig {tp.human_size(originale)} · 150dpi {tp.human_size(a_med)} · "
      f"96dpi {tp.human_size(a_max)} · 72dpi {tp.human_size(a_dpi)})")

# ---------------------------------------------------------------- 3. il documento di partenza non viene toccato
assert len(data) == originale, "compress_pdf non deve modificare i byte di partenza"
print("3. documento di partenza intatto OK")

# ---------------------------------------------------------------- 4. CompressDialog: preset e DPI a mano
root = tk.Tk()
root.withdraw()
dlg = tp.CompressDialog(root)
assert dlg.level.get() == "media", "il livello predefinito dovrebbe essere «media»"
assert str(dlg.img_spin.cget("state")) == "disabled", "con un preset i campi DPI sono bloccati"
assert dlg.img_dpi.get() == "150" and dlg.text_dpi.get() == "0", "preset media = 150/0 DPI"
dlg.level.set("personalizzato")
dlg._on_level()
assert str(dlg.img_spin.cget("state")) == "normal", "con «Personalizzato» i campi DPI sono liberi"
dlg.img_dpi.set("72")
dlg.text_dpi.set("150")
dlg._ok()
assert dlg.result == {"img_dpi": 72, "text_dpi": 150, "quality": 75}, dlg.result
print("4. CompressDialog (preset + DPI a mano) OK")

# ---------------------------------------------------------------- 5. App.cmd_compress salva una copia più leggera
app = tp.App(root)
root.update()
src = CFG / "sorgente.pdf"
src.write_bytes(data)
assert app.open_file(str(src)) is True
root.update()

out = CFG / "copia-compressa.pdf"
out.unlink(missing_ok=True)
esito = {}


class FakeDialog:
    def __init__(self, parent):
        self.result = {"img_dpi": 96, "text_dpi": 200, "quality": 55}
        self.top = tk.Toplevel(parent)
        self.top.after(1, self.top.destroy)   # wait_window ritorna appena la finestra si chiude


tp.CompressDialog = FakeDialog
tp.filedialog.asksaveasfilename = lambda **k: str(out)
tp.messagebox.showinfo = lambda *a, **k: esito.__setitem__("ok", True)
tp.messagebox.showerror = lambda *a, **k: esito.__setitem__("err", (a, k))

app.cmd_compress()
for _ in range(300):
    root.update()
    if esito:
        break
    time.sleep(0.05)

assert "err" not in esito, f"compressione fallita: {esito.get('err')}"
assert esito.get("ok"), "la compressione non è terminata entro il tempo atteso"
assert out.exists() and out.stat().st_size > 0, "la copia compressa non è stata creata"
assert out.stat().st_size < src.stat().st_size, "la copia doveva essere più piccola dell'originale"
assert len(pymupdf.open(out)) == 1
assert app.path == str(src), "il documento aperto non deve cambiare percorso dopo la compressione"
print(f"5. App.cmd_compress OK  ({tp.human_size(src.stat().st_size)} -> {tp.human_size(out.stat().st_size)})")

root.destroy()
print("TUTTI I TEST OK")
