"""Test della conversione di documenti Word/Excel in PDF e della loro apertura come scheda.

I passi 4 e 5 fanno una conversione VERA e servono Word/Excel (o LibreOffice) installati: se mancano,
vengono saltati. I passi 1-3 non richiedono Office (usano un convertitore simulato)."""
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
CFG = WORK / "cfg_office"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()

# ---------------------------------------------------------------- 1. estensioni e tipi
assert {".docx", ".doc", ".rtf"} <= tp.WORD_EXTS, tp.WORD_EXTS
assert {".xlsx", ".xls", ".csv"} <= tp.EXCEL_EXTS, tp.EXCEL_EXTS
assert ".docx" in tp.OFFICE_EXTS and ".xlsx" in tp.OFFICE_EXTS
print("1. estensioni Office OK")

# ---------------------------------------------------------------- 2. nessun motore -> errore, non eccezione
_wi, _ei, _lo = tp.word_installed, tp.excel_installed, tp.libreoffice_exe
tp.word_installed = lambda: False
tp.excel_installed = lambda: False
tp.libreoffice_exe = lambda: None
err = tp.office_to_pdf("qualcosa.docx", CFG / "vuoto.pdf")
assert isinstance(err, str) and "LibreOffice" in err, err
assert tp.office_available() is False
tp.word_installed, tp.excel_installed, tp.libreoffice_exe = _wi, _ei, _lo
print("2. nessun motore: errore gestito OK")

# ---------------------------------------------------------------- 3. apertura come scheda (convertitore simulato)
def finto_convert(src, pdf, timeout=300):
    d = pymupdf.open()
    d.new_page(width=595, height=842).insert_text((72, 72), "CONVERTITO", fontsize=20)
    d.save(str(pdf))
    d.close()
    return None

_orig = tp.office_to_pdf
tp.office_to_pdf = finto_convert
root = tk.Tk()
app = tp.App(root)
root.update()
assert app.open_file(str(CFG / "Relazione.docx")) is True
root.update()
assert app.doc is not None and len(app.doc) == 1, "non si è aperta una scheda"
assert "CONVERTITO" in app.doc[0].get_text(), "il PDF convertito non è stato caricato"
assert app.tab.label == "Relazione", f"etichetta scheda inattesa: {app.tab.label!r}"
assert app.path is None, "una conversione dà un PDF non ancora salvato (path None)"
tp.office_to_pdf = _orig
root.destroy()
print("3. apertura Word come scheda (convertitore simulato) OK")

# ---------------------------------------------------------------- 4. conversione VERA con Word (RTF -> PDF)
if tp.word_installed():
    rtf = CFG / "prova.rtf"
    rtf.write_text(r"{\rtf1\ansi Ciao dal collaudo Word.}", encoding="ascii")
    out = CFG / "prova_word.pdf"
    err = tp.office_to_pdf(str(rtf), out)
    assert err is None, f"Word->PDF fallita: {err}"
    assert out.exists() and len(pymupdf.open(out)) >= 1, "il PDF da Word non è stato creato"
    print("4. conversione reale Word (RTF->PDF) OK")
else:
    print("4. Word non installato: salto la conversione reale")

# ---------------------------------------------------------------- 5. conversione VERA con Excel (CSV -> PDF)
if tp.excel_installed():
    csv = CFG / "prova.csv"
    csv.write_text("Nome,Valore\nAlfa,10\nBeta,20\nGamma,30\n", encoding="utf-8")
    out = CFG / "prova_excel.pdf"
    err = tp.office_to_pdf(str(csv), out)
    assert err is None, f"Excel->PDF fallita: {err}"
    assert out.exists() and len(pymupdf.open(out)) >= 1, "il PDF da Excel non è stato creato"
    print("5. conversione reale Excel (CSV->PDF) OK")
else:
    print("5. Excel non installato: salto la conversione reale")

print("TUTTI I TEST OK")
