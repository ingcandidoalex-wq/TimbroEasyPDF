"""Test dei file firmati .p7m (firma digitale CAdES) e della conversione in blocco di una cartella.

Crea veri involucri PKCS#7/CMS con asn1crypto (DER e base64), ne estrae il contenuto, lo apre come scheda
e converte una cartella di file in PDF. Le parti che richiedono Word/Excel sono saltate se mancano."""
import shutil
import sys
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402
import base64  # noqa: E402
import pymupdf  # noqa: E402
import timbro_pdf as tp  # noqa: E402
from asn1crypto import cms  # noqa: E402

tp.LICENSE_VERSION = 0
CFG = WORK / "cfg_p7m"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()


def un_pdf(testo="CONTENUTO FIRMATO"):
    d = pymupdf.open()
    d.new_page(width=595, height=842).insert_text((72, 72), testo, fontsize=18)
    data = d.tobytes()
    d.close()
    return data


def una_png(path):
    d = pymupdf.open()
    d.new_page(width=200, height=120).insert_text((20, 60), "IMMAGINE", fontsize=16)
    pix = d[0].get_pixmap()
    pix.save(str(path))
    d.close()


def fai_p7m(inner_bytes):
    """Involucro PKCS#7 SignedData (DER) con dentro inner_bytes come eContent."""
    signed = cms.SignedData({
        "version": "v1",
        "digest_algorithms": [],
        "encap_content_info": {"content_type": "data", "content": inner_bytes},
        "signer_infos": [],
    })
    return cms.ContentInfo({"content_type": "signed_data", "content": signed}).dump()


# ---------------------------------------------------------------- 1. estensioni e sniff
assert ".p7m" in tp.P7M_EXTS and tp.OFFICE_EXTS <= tp.CONVERTIBLE_EXTS
assert tp.sniff_ext(b"%PDF-1.7 ...") == ".pdf"
assert tp.sniff_ext(b"\x89PNG\r\n\x1a\n...") == ".png"
assert tp.sniff_ext(b"\xff\xd8\xff\xe0 jfif") == ".jpg"
print("1. estensioni e riconoscimento tipo OK")

# ---------------------------------------------------------------- 2. estrazione da .p7m (DER)
pdf_bytes = un_pdf()
p7m_der = CFG / "relazione.pdf.p7m"
p7m_der.write_bytes(fai_p7m(pdf_bytes))
estratto = tp.extract_p7m(str(p7m_der))
assert estratto == pdf_bytes, "il contenuto estratto dal .p7m (DER) non coincide"
assert tp.inner_name_for(str(p7m_der), estratto) == "relazione.pdf"
print("2. estrazione da .p7m binario (DER) OK")

# ---------------------------------------------------------------- 3. estrazione da .p7m in base64 + nome da indovinare
p7m_b64 = CFG / "firma.p7m"            # niente estensione interna nel nome -> si indovina dai byte
p7m_b64.write_bytes(base64.b64encode(fai_p7m(pdf_bytes)))
estratto2 = tp.extract_p7m(str(p7m_b64))
assert estratto2 == pdf_bytes, "il contenuto estratto dal .p7m (base64) non coincide"
assert tp.inner_name_for(str(p7m_b64), estratto2) == "firma.pdf", tp.inner_name_for(str(p7m_b64), estratto2)
print("3. estrazione da .p7m base64 e nome indovinato OK")

# ---------------------------------------------------------------- 4. .p7m non valido -> errore chiaro, non eccezione grezza
brutto = CFG / "rotto.p7m"
brutto.write_bytes(b"questo non e' un involucro firmato")
try:
    tp.extract_p7m(str(brutto))
    assert False, "doveva sollevare ValueError"
except ValueError as e:
    assert "p7m" in str(e).lower() or "documento" in str(e).lower()
print("4. .p7m non valido: errore gestito OK")

# ---------------------------------------------------------------- 5. convert_file_to_pdf: PDF, immagine, .p7m
out_pdf = CFG / "da_pdf.pdf"
src_pdf = CFG / "sorgente.pdf"
src_pdf.write_bytes(pdf_bytes)
assert tp.convert_file_to_pdf(str(src_pdf), str(out_pdf)) is None and len(pymupdf.open(out_pdf)) == 1

png = CFG / "foto.png"
una_png(png)
out_png = CFG / "da_png.pdf"
assert tp.convert_file_to_pdf(str(png), str(out_png)) is None and len(pymupdf.open(out_png)) == 1

out_p7m = CFG / "da_p7m.pdf"
assert tp.convert_file_to_pdf(str(p7m_der), str(out_p7m)) is None
assert "CONTENUTO FIRMATO" in pymupdf.open(out_p7m)[0].get_text()
print("5. convert_file_to_pdf (PDF, immagine, .p7m) OK")

# ---------------------------------------------------------------- 6. apertura .p7m come scheda non salvata
root = tk.Tk()
app = tp.App(root)
root.update()
assert app.open_file(str(p7m_der)) is True
root.update()
assert app.doc is not None and "CONTENUTO FIRMATO" in app.doc[0].get_text()
assert app.path is None, "il contenuto di un .p7m è un PDF non ancora salvato"
assert app.tab.label == "relazione", f"etichetta inattesa: {app.tab.label!r}"
print("6. apertura .p7m come scheda OK")

# ---------------------------------------------------------------- 7. conversione di una cartella (dialoghi simulati)
cartella = CFG / "lotto"
shutil.rmtree(cartella, ignore_errors=True)
cartella.mkdir()
una_png(cartella / "disegno.png")
(cartella / "atto.pdf.p7m").write_bytes(fai_p7m(un_pdf("ATTO")))
(cartella / "nota.txt").write_text("non convertibile", encoding="utf-8")   # dev'essere ignorato

tp.filedialog.askdirectory = lambda **k: str(cartella)
tp.messagebox.askyesno = lambda *a, **k: "aprire la cartella" not in (a[1] if len(a) > 1 else k.get("message", ""))
tp.messagebox.showinfo = lambda *a, **k: None
tp.messagebox.showwarning = lambda *a, **k: None
tp.messagebox.showerror = lambda *a, **k: None

app.cmd_convert_folder()
root.update()
assert (cartella / "disegno.pdf").exists(), "immagine non convertita"
assert (cartella / "atto.pdf").exists(), ".p7m non convertito"
assert not (cartella / "nota.pdf").exists(), "il .txt non doveva essere convertito"
assert "ATTO" in pymupdf.open(cartella / "atto.pdf")[0].get_text()
print("7. conversione della cartella OK")

# ---------------------------------------------------------------- 8. conversione cartella: niente sovrascrittura
(cartella / "disegno.pdf").write_bytes(b"gia' esistente")   # esiste gia'
una_png(cartella / "disegno.png")                            # stesso nome base
# rimuovo gli altri cosi' resta solo disegno.png da convertire di nuovo
(cartella / "atto.pdf.p7m").unlink()
app.cmd_convert_folder()
root.update()
assert (cartella / "disegno (2).pdf").exists(), "doveva creare «disegno (2).pdf» senza sovrascrivere"
print("8. conversione cartella senza sovrascrivere OK")

root.destroy()
print("TUTTI I TEST OK")
