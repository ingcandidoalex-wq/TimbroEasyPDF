"""Crea la guida illustrata (tp.GUIDE_NAME), con le schermate del programma vero.

Uso:  python crea_guida.py
Apre il programma con documenti, timbro e firma di esempio (in una cartella temporanea: i timbri
e le impostazioni dell'utente non vengono toccati), fotografa le finestre, evidenzia le parti
spiegate e impagina tutto in un PDF accanto a questo script.
"""

import ctypes
import ctypes.wintypes as wt
import math
import shutil
import tempfile
import time
import tkinter as tk
from pathlib import Path

import pymupdf

import timbro_pdf as tp

HERE = Path(__file__).resolve().parent
WORK = Path(tempfile.mkdtemp(prefix="guida_timbropdf_"))
SHOTS = WORK / "immagini"
SHOTS.mkdir()
FONTS = Path(r"C:\Windows\Fonts")
RED = (0.85, 0.12, 0.12)

# ---------------------------------------------------------------- programma in una cartella di prova
tp.CONFIG_DIR, tp.CONFIG_FILE = WORK / "config", WORK / "config" / "config.json"
(WORK / "config").mkdir()
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = WORK / "config" / "timbri", WORK / "config" / "timbro.png"
tp.LICENSE_VERSION = 0
tp.UPDATE_URL = ""   # niente controllo degli aggiornamenti mentre si fanno le schermate
tp.tk.Menu.tk_popup = lambda self, *a, **k: None
for name in ("showinfo", "showwarning", "showerror"):
    setattr(tp.messagebox, name, lambda *a, **k: None)
tp.messagebox.askyesno = lambda *a, **k: True


def grab(window, name, marks=(), crop=None):
    """Fotografa la finestra (anche se coperta); marks: [(widget o (x0, y0, x1, y1) nello schermo, numero)]."""
    window.update()
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
    hwnd = int(window.wm_frame(), 16)
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

    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(BMI(40, w, -h, 1, 32, 0)), 0)
    raw = buf.raw
    rgb = bytearray(w * h * 3)
    rgb[0::3], rgb[1::3], rgb[2::3] = raw[2::4], raw[1::4], raw[0::4]
    pix = pymupdf.Pixmap(pymupdf.csRGB, w, h, bytes(rgb), False)
    # a volte Windows restituisce neri i bordi invisibili della finestra (a destra e in basso): si tolgono
    black = bytes(3)
    right = next(k for k in range(w) if any(rgb[(y * w + w - 1 - k) * 3:(y * w + w - k) * 3] != black
                                             for y in range(40, h, 25)))
    bottom = next(k for k in range(h) if any(rgb[((h - 1 - k) * w + x) * 3:((h - 1 - k) * w + x + 1) * 3] != black
                                              for x in range(0, w, 25)))
    if (right or bottom) and not crop:
        crop = (rect.left, rect.top, rect.right - right, rect.bottom - bottom)
    # evidenziazioni: riquadri rossi con un numero
    doc = pymupdf.open()
    page = doc.new_page(width=w, height=h)
    page.insert_image(page.rect, pixmap=pix)
    for target, number in marks:
        if isinstance(target, tuple):
            x0, y0, x1, y1 = target
        else:
            x0, y0 = target.winfo_rootx(), target.winfo_rooty()
            x1, y1 = x0 + target.winfo_width(), y0 + target.winfo_height()
        r = pymupdf.Rect(x0 - rect.left, y0 - rect.top, x1 - rect.left, y1 - rect.top) + (2, 2, -2, -2)
        page.draw_rect(r, color=RED, width=3)
        c = pymupdf.Point(min(r.x0 + 16, r.x1 - 16), min(r.y0 + 16, r.y1 - 16))
        page.draw_circle(c, 13, color=RED, fill=RED)
        page.insert_text(c + (-4.5 if number < 10 else -9, 5.5), str(number), fontsize=15, color=(1, 1, 1),
                         fontname="hebo")
    clip = pymupdf.Rect(crop) - (rect.left, rect.top, rect.left, rect.top) if crop else None
    page.get_pixmap(clip=clip).save(str(SHOTS / f"{name}.png"))


# ---------------------------------------------------------------- documenti, timbro e firma di esempio
def sample_documents():
    rel = pymupdf.open()
    paragraphs = ("La presente relazione illustra i criteri di progetto e le verifiche strutturali "
                  "dell'intervento, eseguite secondo le norme tecniche vigenti.",
                  "Le azioni considerate comprendono i carichi permanenti, i sovraccarichi d'esercizio, "
                  "la neve e l'azione sismica; le combinazioni sono quelle previste dalla normativa.",
                  "Gli elementi in calcestruzzo armato sono stati verificati allo stato limite ultimo e "
                  "di esercizio; i risultati sono riportati nelle tabelle allegate.")
    for n in range(1, 4):
        page = rel.new_page(width=595, height=842)
        page.insert_text((60, 80), "RELAZIONE TECNICA DI CALCOLO", fontsize=18, fontname="hebo")
        page.insert_text((60, 104), f"Pratica edilizia n. 1234/2026  –  pagina {n}", fontsize=10, fontname="helv")
        page.draw_line((60, 114), (535, 114), color=(0.3, 0.3, 0.3))
        y = 150
        for k, text in enumerate(paragraphs):
            page.insert_textbox(pymupdf.Rect(60, y, 535, y + 70), f"{n}.{k + 1}  {text}", fontsize=11,
                                fontname="tiro")
            y += 80
        page.insert_text((60, y + 20), "Importo dei lavori: euro 125.400,00", fontsize=11, fontname="tiro")
        for i in range(6):
            page.draw_rect(pymupdf.Rect(60, y + 50 + i * 22, 535, y + 72 + i * 22), color=(0.6, 0.6, 0.65))
            page.insert_text((68, y + 66 + i * 22), f"Elemento {i + 1}        verifica soddisfatta", fontsize=9,
                             fontname="helv")
    rel.save(WORK / "Relazione tecnica.pdf")
    tav = pymupdf.open()
    for n in range(1, 3):
        page = tav.new_page(width=842, height=595)
        page.draw_rect(pymupdf.Rect(20, 20, 822, 575), color=(0, 0, 0), width=1.5)
        for i in range(8):
            page.draw_rect(pymupdf.Rect(80 + i * 85, 150, 140 + i * 85, 210), color=(0.2, 0.2, 0.6))
            page.draw_line((110 + i * 85, 210), (110 + i * 85, 420), color=(0.4, 0.4, 0.4))
        page.draw_rect(pymupdf.Rect(560, 480, 822, 575), color=(0, 0, 0))
        page.insert_text((570, 505), f"TAVOLA S{n:02d} – PIANTA FONDAZIONI", fontsize=11, fontname="hebo")
        page.insert_text((570, 530), "Scala 1:100", fontsize=9, fontname="helv")
    tav.save(WORK / "Tavola strutturale.pdf")
    scan = pymupdf.open()
    src = pymupdf.open()
    p = src.new_page(width=595, height=842)
    p.insert_text((70, 120), "VERBALE DI SOPRALLUOGO", fontsize=16, fontname="tibo")
    p.insert_textbox(pymupdf.Rect(70, 150, 525, 400), "In data odierna si è svolto il sopralluogo presso il "
                     "cantiere. Le opere risultano eseguite in conformità al progetto approvato.",
                     fontsize=12, fontname="tiro")
    page = scan.new_page(width=595, height=842)
    page.insert_image(page.rect, pixmap=p.get_pixmap(dpi=150, colorspace=pymupdf.csGRAY))
    scan.save(WORK / "Verbale scansionato.pdf")
    return [WORK / "Relazione tecnica.pdf", WORK / "Tavola strutturale.pdf", WORK / "Verbale scansionato.pdf"]


def sample_stamp():
    """Un timbro tondo di fantasia, su fondo bianco come una scansione."""
    doc = pymupdf.open()
    page = doc.new_page(width=200, height=200)
    blue = (0.12, 0.23, 0.58)
    page.draw_circle((100, 100), 92, color=blue, width=4)
    page.draw_circle((100, 100), 66, color=blue, width=2)
    text = "ORDINE DEGLI INGEGNERI  *  PROVINCIA DI ESEMPIO  *"
    for k, ch in enumerate(text):
        a = math.radians(-90 + k * 360 / len(text))
        x, y = 100 + 78 * math.cos(a), 100 + 78 * math.sin(a)
        page.insert_text((x - 3, y + 4), ch, fontsize=10, fontname="hebo", color=blue,
                         morph=(pymupdf.Point(x, y), pymupdf.Matrix(math.degrees(a) + 90)))
    for dy, line, size in ((-14, "Ing.", 12), (6, "MARIO ROSSI", 13), (26, "n. A0000", 11)):
        w = pymupdf.get_text_length(line, "hebo", size)
        page.insert_text((100 - w / 2, 100 + dy), line, fontsize=size, fontname="hebo", color=blue)
    return page.get_pixmap(dpi=200)


# ---------------------------------------------------------------- schermate
def settle(root, seconds=0.7, app=None):
    if app:   # nella barra di stato niente percorsi della cartella temporanea
        app.status("Pronto.")
    end = time.time() + seconds
    while time.time() < end:
        root.update()
        time.sleep(0.03)


def shots():
    tp.set_dpi_awareness()
    root = tk.Tk()
    app = tp.App(root)
    for key in ("text", "pages", "stamps"):   # nella guida le barre e il pannello si mostrano aperti
        app._set_visible(key, True)
    root.geometry("1420x900+20+20")
    settle(root, 1)
    docs = sample_documents()
    app.library.add(sample_stamp(), "Timbro", width_mm=38)
    pix, box = tp.text_signature("Mario Rossi", app.fonts.get("Segoe Script"), 26, tp.hex_to_rgb("#1f3a93"))
    app.library.add(pix, "Firma", width_mm=round(box.width / tp.MM, 1), remove_white=False,
                    source={"kind": "typed", "text": "Mario Rossi", "font": "Segoe Script", "size": 26.0,
                            "color": "#1f3a93"})
    app.select_stamp(0)
    for d in docs[:2]:
        app.open_file(str(d))
    app.switch_tab(app.tabs[0])
    app.set_zoom_mode("width")
    settle(root)

    def page_point(fx, fy):
        pr = app.doc[app.current].rect
        ox, oy = app.page_origin
        return (int(ox + fx * pr.width * app.cur_zoom - app.view.canvasx(0)),
                int(oy + fy * pr.height * app.cur_zoom - app.view.canvasy(0)))

    def click(x, y):
        app.view.event_generate("<Motion>", x=x, y=y)
        app.view.event_generate("<ButtonPress-1>", x=x, y=y)
        app.view.event_generate("<ButtonRelease-1>", x=x, y=y)
        settle(root, 0.3)

    # 1. la finestra
    settle(root, 0.3, app)
    grab(root, "finestra", [(app.bars["main"], 1), (app.bars["text"], 2), (app.tabbar, 3),
                            (app.side_panels["thumbs"], 4), (app.view_area, 5), (app.side_panels["stamps"], 6)])
    # 2. modalità timbro: anteprima che segue il mouse, poi timbro e firma apposti
    app.view.yview_moveto(0.55)
    settle(root, 0.3)
    app.set_stamp_mode(True)
    x, y = page_point(0.70, 0.80)
    app.view.event_generate("<Motion>", x=x, y=y)
    settle(root, 0.4, app)
    grab(root, "modo_timbro", [(app.mode_btn, 1), (app.gallery, 2)])
    click(x, y)
    app.select_stamp(1)
    x2, y2 = page_point(0.40, 0.83)
    click(x2, y2)
    app.set_stamp_mode(False)
    app.view.event_generate("<Motion>", x=x, y=y)
    settle(root, 0.4, app)
    grab(root, "timbri_apposti", [((app.view.winfo_rootx() + x - 90, app.view.winfo_rooty() + y - 90,
                                    app.view.winfo_rootx() + x + 90, app.view.winfo_rooty() + y + 90), 1)])
    # 3. firma: finestra di creazione/modifica
    dialog = tp.SignatureDialog(root, source=app.library.items[1]["source"], name="Firma")
    settle(root)
    grab(dialog.top, "firma")
    dialog.top.destroy()
    # 3b. modifica del timbro: un tratto di evidenziatore
    app.select_stamp(0)
    item = app.library.items[0]
    editor = tp.StampEditor(root, app.fonts, pymupdf.Pixmap(app.stamp.original_png), item["width_mm"],
                            f"Modifica timbro «{item['name']}»", name=item["name"], rotation=item["rotate"],
                            remove_white=item["remove_white"], tolerance=item["white_tolerance"])
    settle(root)
    editor.tool_var.set("highlight")
    editor._tool_changed()
    w, h, f = editor.colors.width, editor.colors.height, editor.f
    points = [(round(editor.ox + (0.12 + 0.76 * k / 20) * w * f), round(editor.oy + 0.86 * h * f)) for k in range(21)]
    editor.canvas.event_generate("<ButtonPress-1>", x=points[0][0], y=points[0][1])
    for x, y in points[1:]:
        editor.canvas.event_generate("<B1-Motion>", x=x, y=y, state=0x100)
    editor.canvas.event_generate("<ButtonRelease-1>", x=points[-1][0], y=points[-1][1])
    settle(root)
    grab(editor.top, "editor_timbro")
    editor.top.destroy()
    # 4. testo: riga evidenziata, finestra di modifica, casella aggiunta
    app.view.yview_moveto(0)
    settle(root, 0.3)
    app.set_tool("edit_text")
    line = next(l for l in tp.text_lines(app.doc[0]) if l.text.startswith("Importo"))
    c = (line.rect.tl + line.rect.br) / 2
    hx, hy = page_point(c.x / app.doc[0].rect.width, c.y / app.doc[0].rect.height)
    app.view.event_generate("<Motion>", x=hx, y=hy)
    settle(root, 0.4, app)
    grab(root, "modifica_testo", [(app.bars["text"].items[1], 1)])
    dialog = tp.TextDialog(root, app.fonts, "Modifica testo", {"text": line.text, "font": "Times New Roman",
                                                               "size": 11, "color": "#000000"},
                           original="Testo originale: Times-Roman, 11.0 pt", whiteout=False)
    settle(root)
    grab(dialog.top, "dialogo_testo")
    dialog.top.destroy()
    app.set_tool(None)
    # 5. vista d'insieme
    app.toggle_overview(True)
    settle(root, 2.5, app)
    grab(root, "vista_insieme", [(app.overview_area, 1)])
    app.toggle_overview(False)
    # 6. unione e divisione
    dialog = tp.MergeDialog(root, app.tabs)
    settle(root)
    grab(dialog.top, "unisci")
    dialog.top.destroy()
    dialog = tp.SplitDialog(root, len(app.doc), "Relazione tecnica", r"D:\Documenti\Pratiche\1234-2026", per=2)
    settle(root)
    grab(dialog.top, "dividi")
    dialog.top.destroy()
    # 7. OCR e ricerca
    app.open_file(str(docs[2]))
    settle(root)
    dialog = tp.OcrDialog(root, 0, [0], 1)
    settle(root)
    grab(dialog.top, "ocr")
    dialog.top.destroy()
    done = {"ok": False}
    app.cmd_ocr(on_done=lambda: done.update(ok=True), all_pages=True)
    while not done["ok"]:
        settle(root, 0.2)
    app.find_var.set("sopralluogo")
    app.find(1)
    settle(root, 0.7, app)
    grab(root, "ricerca", [(app.find_entry.master, 1), ((app.view.winfo_rootx() + 60, app.view.winfo_rooty() + 20,
                                                         app.view.winfo_rootx() + app.view.winfo_width() - 60,
                                                         app.view.winfo_rooty() + 260), 2)])
    app.clear_find()
    # 8. disposizione: barre e pannelli spostati
    app.switch_tab(app.tabs[0])
    app.layout["bars"]["text"] = ["left", 0]
    app.layout["panels"]["thumbs"] = ["bottom", 0]
    app.layout["heights"]["thumbs"] = 190
    app._apply_layout()
    settle(root, 1.2, app)
    grab(root, "disposizione", [(app.bars["main"].grip, 1), (app.side_panels["thumbs"].winfo_children()[0], 2)])
    app.reset_layout()
    settle(root, 0.6)
    # 9. archivio dei timbri, licenza
    dialog = tp.ArchiveDialog(root, Path(r"D:\Dropbox\Timbri") / tp.ARCHIVE_NAME, 2, True, True)
    settle(root)
    grab(dialog.top, "archivio")
    dialog.top.destroy()
    dialog = tp.AboutDialog(root)
    settle(root)
    grab(dialog.top, "licenza")
    dialog.top.destroy()
    root.destroy()


# ---------------------------------------------------------------- impaginazione
SECTIONS = [
    ("La finestra", "finestra", 100, """
<p>{APP} apre, ruota, unisce, timbra, firma e modifica i PDF. La finestra è divisa in zone:</p>
<ol>
<li><b>Barra principale</b>: apri, nuovo, salva (solo icona), aggiungi file; annulla e ripristina; pagine,
zoom, rotazione, ordine delle pagine e ricerca. A destra gli interruttori <b>Testo</b>, <b>Pagine</b> e
<b>Timbri</b> aprono e chiudono i gruppi di strumenti e il pannello dei timbri (partono chiusi). Se la
finestra è stretta, la barra va a capo da sola.</li>
<li><b>Barra del testo</b> (interruttore «Testo»): modifica e aggiunta di testo, bianchetto, oscura, OCR,
esportazione in Word. <b>Barra delle pagine</b> (interruttore «Pagine»): unisci schede, dividi, vista
d'insieme.</li>
<li><b>Schede</b>: ogni PDF aperto ha la sua scheda. Il pallino ● indica modifiche da salvare.
Trascina una scheda per cambiarne l'ordine.</li>
<li><b>Miniature</b>: clic per andare alla pagina, Ctrl+clic e Maiusc+clic per sceglierne più
d'una; trascinale per riordinarle o portarle in un altro PDF.</li>
<li><b>Pagina</b>: rotellina per scorrere, Ctrl+rotellina per lo zoom. Dal menu <b>Visualizza</b> scegli
se scorrere una pagina per volta o in modo continuo (tutte le pagine di seguito).</li>
<li><b>Timbri e firme</b>: la galleria dei timbri salvati e le impostazioni di quello scelto.</li>
</ol>
<p>Tutte le modifiche valgono per la scheda attiva e vanno salvate con <b>Salva…</b> (Ctrl+S).
Per aprire i PDF puoi anche trascinarli da Esplora risorse sulla finestra. Le finestre per aprire,
salvare, esportare e scegliere le immagini dei timbri ripartono ognuna dall'ultima cartella usata per
quell'operazione.</p>"""),
    ("Apporre un timbro", "modo_timbro", 100, """
<p>Scegli il timbro nella galleria (2) con un clic o con i tasti 1–9, imposta la larghezza in millimetri e
premi <b>Attiva modalità timbro</b> (1) o il tasto T. Il timbro segue il mouse: clicca sul foglio per
apporlo. Puoi cambiare pagina e cliccare di nuovo dove serve. Esc per uscire.</p>
<p><b>Rimuovi sfondo bianco</b> rende trasparente il bianco di una scansione; <b>Primo piano</b> o
<b>Secondo piano</b> mettono il timbro sopra o sotto testo e disegni. <b>Ripeti su tutte le pagine</b>
lo copia nella stessa posizione sulle altre pagine (o su quelle selezionate).</p>"""),
    ("Spostare, ruotare e modificare timbri e testi", "timbri_apposti", 100, """
<p>Fuori dalla modalità timbro, passando sopra un timbro o una casella di testo compare un bordo
tratteggiato (1): <b>trascinalo</b> per spostarlo, <b>doppio clic</b> per modificarlo, <b>Canc</b> per
eliminarlo. Con il <b>tasto destro</b> un timbro si ruota a destra o a sinistra attorno al suo centro.
Funziona anche sui PDF già salvati e riaperti. Ctrl+Z annulla.</p>
<p>Il doppio clic su un timbro apposto apre la stessa finestra di modifica della galleria, ma cambia
solo quel timbro, che resta dov'era. Su una firma o una data scritta da tastiera puoi correggere
direttamente il testo (per esempio la data), senza toglierla e rimetterla.</p>"""),
    ("Firme", "firma", 75, """
<p><b>Crea firma…</b> nel pannello dei timbri: disegnala con mouse, penna o dito, oppure scrivila con la
tastiera scegliendo font, dimensione e colore. Viene salvata nella galleria e si appone come un timbro.
Lo stesso vale per una data o una dicitura: scrivila con il font che preferisci.</p>"""),
    ("Modificare e ruotare i timbri", "editor_timbro", 85, """
<p><b>Modifica…</b> (o doppio clic nella galleria) apre il timbro, di qualunque tipo:</p>
<ul>
<li><b>Ruota a sinistra / a destra</b> dei gradi scritti fra i due pulsanti (90, oppure 10, 45…): le misure
vere nel PDF restano le stesse; con angoli diversi da 90° gli angoli dell'immagine restano trasparenti;</li>
<li><b>Matita</b>, <b>Evidenziatore</b> (semitrasparente) e <b>Gomma</b>, con colore e spessore;</li>
<li><b>Testo</b>: clicca dove scrivere e scegli font, dimensione e colore, vedendo subito il risultato;</li>
<li><b>Apri con l'editor di immagini…</b> lo apre nel programma di Windows (di solito Paint): quando lo
salvi lì, nella finestra si aggiorna da solo;</li>
<li><b>Sostituisci immagine…</b> mette un'altra immagine mantenendo nome e impostazioni.</li>
</ul>
<p>Le firme create nel programma si possono anche <b>riscrivere</b> o <b>ridisegnare</b>; dopo un ritocco
con matita, gomma o testo restano un'immagine. Annulla e Ripristina (Ctrl+Z, Ctrl+Y) tornano indietro.
I pulsanti ↺ e ↻ sotto la galleria ruotano al volo il timbro scelto dei gradi indicati accanto («Ruota di
… °»); gli stessi gradi valgono per i timbri già sulla pagina (tasto destro). Le modifiche valgono per i
timbri apposti da quel momento in poi.</p>"""),
    ("Correggere il testo del PDF", "modifica_testo", 100, """
<p>Con <b>Modifica testo</b> (1) le righe si colorano passandoci sopra: clicca quella da correggere.
Il vecchio testo viene eliminato davvero e il nuovo è scritto nello stesso punto, con il font più simile.
Nelle scansioni spunta <b>Copri con il bianco lo sfondo</b>.</p>
<p><b>Aggiungi testo</b>: clicca dove scrivere. Le caselle si trascinano per spostarle anche con lo
strumento attivo; un clic le corregge. <b>Bianchetto</b> e <b>Oscura</b>: trascina un rettangolo,
quanto c'è sotto viene cancellato davvero.</p>
<p><b>Selezionare e copiare il testo</b>: senza strumenti attivi, trascina il mouse sul testo della
pagina (o fai doppio clic su una parola) per evidenziarlo, poi <b>Ctrl+C</b> per copiarlo negli
appunti.</p>"""),
    ("La finestra del testo", "dialogo_testo", 75, """
<p>Scrivi il testo (Invio va a capo, Ctrl+Invio conferma), scegli font, dimensione e colore: il testo
nel riquadro prende subito il font e il colore scelti e, mentre scrivi, compare anche sulla pagina, nel
punto e nella misura in cui verrà scritto. Con i pulsanti <b>G</b>, <b>C</b> e <b>S</b> lo rendi
grassetto, corsivo o sottolineato. La finestra si sposta da sola per non coprirlo.
L'ultimo stile usato viene ricordato.</p>"""),
    ("Vista d'insieme di tutti i PDF", "vista_insieme", 100, """
<p><b>Vista d'insieme</b> (F4) mostra al posto della pagina le miniature di tutti i PDF aperti, una
colonna per documento (1). Trascina le pagine da una colonna all'altra per copiarle, nella stessa
colonna per spostarle. Doppio clic su una pagina per aprirla; Esc o F4 per tornare.</p>"""),
    ("Unire i documenti", "unisci", 60, """
<p><b>Unisci schede…</b>: spunta i PDF aperti da unire e scegline l'ordine con le frecce. Il risultato
si apre in una nuova scheda, da salvare con <b>Salva…</b>. Puoi anche trascinare singole pagine sulla
scheda di un altro PDF.</p>"""),
    ("Dividere un documento", "dividi", 65, """
<p><b>Dividi…</b>: scegli quante pagine per file (singole, doppie o un numero a piacere), il nome, il
numero iniziale e la cartella. I file escono numerati: <i>Relazione tecnica_01.pdf</i>,
<i>Relazione tecnica_02.pdf</i>…</p>"""),
    ("Riconoscere il testo delle scansioni (OCR)", "ocr", 60, """
<p><b>Riconosci testo (OCR)…</b>: la pagina resta identica, ma sopra viene aggiunto il testo riconosciuto,
invisibile. Da quel momento si può cercare, copiare, correggere ed esportare in Word.</p>"""),
    ("Cercare", "ricerca", 100, """
<p>Scrivi nella casella <b>Cerca</b> (1) e premi Invio: i risultati si evidenziano sulla pagina (2);
▲ ▼ passano al precedente e al successivo.</p>
<p><b>Esporta in Word…</b> crea un documento Word usando Microsoft Word installato sul PC.
<b>File › Stampa…</b> (Ctrl+P) manda il documento alla stampante: scegli stampante, quali pagine
(tutte, la corrente o un intervallo) e il numero di copie.</p>
<p><b>File › Apri Word o Excel (converti in PDF)…</b> apre un documento Word o un foglio Excel
convertendolo in PDF, pronto da timbrare (serve Word o Excel installato sul PC, oppure LibreOffice).
Puoi anche trascinare il file Word/Excel sulla finestra.</p>"""),
    ("Sistemare barre e pannelli", "disposizione", 100, """
<p>Gli interruttori <b>Testo</b>, <b>Pagine</b>, <b>Timbri</b> in alto a destra (e le stesse voci nel menu
<b>Visualizza</b>) mostrano e nascondono i gruppi di strumenti e il pannello dei timbri. Trascina la
maniglia ⋮ di una barra (1) verso un lato della finestra per agganciarla lì; trascina l'intestazione di un
pannello (2) per spostarlo; i divisori grigi si trascinano per allargare i pannelli. Da
<b>Visualizza › Aspetto</b> cambi il tema grafico (Windows moderno o classico, chiaro, standard…).
Col <b>tasto destro</b> sulla pagina compare un menù che cambia secondo il punto e lo stato. Tutto viene
ricordato; <b>Visualizza › Ripristina disposizione</b> torna all'inizio.</p>"""),
    ("Archivio dei timbri e altri PC", "archivio", 70, """
<p>Tutti i timbri e le firme stanno in un unico file, <i>timbri.tpdf</i>, creato da solo se manca.
Da <b>Timbro › Archivio dei timbri</b> puoi:</p>
<ul>
<li><b>spostarlo in una cartella condivisa</b> (per esempio Dropbox) e, sugli altri PC, scegliere
<b>Usa un archivio esistente</b> indicando lo stesso file: tutti i PC avranno gli stessi timbri;</li>
<li><b>proteggerlo con una password</b>: il file viene cifrato (AES-256) e senza password non si apre.
Ogni PC può ricordarla, custodita da Windows per il tuo utente.</li>
</ul>"""),
    ("Informazioni e licenza", "licenza", 65, """
<p>Dal menu <b>?</b> trovi le condizioni d'uso, l'autore e il modo per sostenere il programma.
<b>Scrivi all'autore</b> apre la posta con un messaggio pronto per inviare suggerimenti o segnalare un
problema. Una volta al giorno il programma controlla se è uscita una versione nuova e, se c'è, propone di
aprire la pagina da cui scaricarla (si disattiva dallo stesso menu).
<b>File › Usa {APP} per aprire i PDF</b> lo registra fra i programmi per i PDF e apre la pagina
di Windows dove sceglierlo: i PDF aperti con doppio clic arrivano come nuove schede.</p>"""),
]

SHORTCUTS = """
<p>Le scorciatoie sono scritte anche accanto alle voci dei menu.</p>
<table>
<tr><td>Ctrl+O</td><td>Apri (anche più PDF)</td><td>Ctrl+N</td><td>Nuovo documento</td></tr>
<tr><td>Ctrl+I</td><td>Aggiungi PDF/immagini in fondo</td><td>Ctrl+Maiusc+I</td><td>Inserisci dopo la pagina corrente</td></tr>
<tr><td>Ctrl+S</td><td>Salva</td><td>Ctrl+E</td><td>Estrai le pagine selezionate</td></tr>
<tr><td>Ctrl+U</td><td>Unisci le schede</td><td>Ctrl+D</td><td>Dividi il documento</td></tr>
<tr><td>Ctrl+Maiusc+W</td><td>Esporta in Word</td><td>Ctrl+W</td><td>Chiudi la scheda</td></tr>
<tr><td>Ctrl+Z / Ctrl+Y</td><td>Annulla / Ripristina</td><td>Ctrl+Tab</td><td>Scheda successiva</td></tr>
<tr><td>Ctrl+R</td><td>Ruota la pagina a destra</td><td>Ctrl+Maiusc+R</td><td>Ruota la pagina a sinistra</td></tr>
<tr><td>Ctrl+↑ / Ctrl+↓</td><td>Sposta le pagine</td><td>Canc</td><td>Elimina timbro o pagine</td></tr>
<tr><td>Ctrl+0 / 1 / 2</td><td>Pagina intera / 100% / larghezza</td><td>Ctrl+rotellina</td><td>Zoom</td></tr>
<tr><td>Ctrl+M</td><td>Modifica testo</td><td>Ctrl+T</td><td>Aggiungi testo</td></tr>
<tr><td>Ctrl+B</td><td>Bianchetto</td><td>Ctrl+Maiusc+B</td><td>Oscura</td></tr>
<tr><td>Ctrl+Maiusc+O</td><td>Riconosci testo (OCR)</td><td>Ctrl+F</td><td>Cerca</td></tr>
<tr><td>Ctrl+C</td><td>Copia il testo selezionato</td><td>Ctrl+P</td><td>Stampa</td></tr>
<tr><td>T</td><td>Modalità timbro</td><td>1–9</td><td>Scegli il timbro</td></tr>
<tr><td>Ctrl+Maiusc+N</td><td>Timbro da immagine</td><td>Ctrl+Maiusc+F</td><td>Crea firma</td></tr>
<tr><td>Ctrl+Maiusc+M</td><td>Modifica il timbro</td><td>Ctrl+Maiusc+→ / ←</td><td>Ruota il timbro</td></tr>
<tr><td>F2</td><td>Rinomina il timbro</td><td>Ctrl+Maiusc+P</td><td>Ripeti il timbro su tutte</td></tr>
<tr><td>F4</td><td>Vista d'insieme</td><td>Esc</td><td>Esci dallo strumento</td></tr>
<tr><td>F1</td><td>Guida rapida</td><td>Maiusc+F1</td><td>Questa guida</td></tr>
<tr><td>PagSu / PagGiù</td><td>Pagina precedente / successiva</td><td>Ctrl+Q</td><td>Esci</td></tr>
</table>"""

CSS = """
@font-face { font-family: Segoe; src: url(segoeui.ttf); }
@font-face { font-family: Segoe; src: url(segoeuib.ttf); font-weight: bold; }
@font-face { font-family: Segoe; src: url(segoeuii.ttf); font-style: italic; }
* { font-family: Segoe; }
body { font-size: 10.5pt; line-height: 1.35; color: #222; }
h1 { font-size: 26pt; color: #1a3d7c; margin: 0; }
h2 { font-size: 15pt; color: #1a3d7c; margin-top: 16pt; margin-bottom: 4pt; }
p { margin-top: 3pt; margin-bottom: 5pt; }
li { margin-bottom: 3pt; }
.sub { font-size: 12pt; color: #555; }
.note { font-size: 9pt; color: #666; }
table { border-collapse: collapse; width: 100%; }
td { border: 1px solid #c8c8d0; padding: 4pt; font-size: 9.5pt; }
"""


def build_pdf(target):
    archive = pymupdf.Archive(str(SHOTS))
    for font in ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf"):
        shutil.copy(FONTS / font, SHOTS / font)
    tp.icon_pixmap(160).save(str(SHOTS / "icona.png"))
    html = [f"""<p style="text-align:center; margin-top:110pt"><img src="icona.png" width="110"/></p>
<h1 style="text-align:center">{tp.APP_NAME}</h1>
<p class="sub" style="text-align:center">Guida illustrata – versione {tp.APP_VERSION}</p>
<p class="sub" style="text-align:center; margin-top:30pt">Ideato dall'{tp.AUTHOR}<br/>
con l'ausilio dell'intelligenza artificiale</p>
<p class="note" style="text-align:center; margin-top:200pt">Le schermate mostrano documenti, timbri e
nomi di esempio.</p>
<p style="page-break-before:always"></p>"""]
    for title, image, width, text in SECTIONS:
        text = text.replace("{APP}", tp.APP_NAME)
        html.append(f"<h2>{title}</h2>{text}<p style='text-align:center'>"
                    f"<img src='{image}.png' style='width:{width}%'/></p>")
    html.append(f"<h2>Scorciatoie da tastiera</h2>{SHORTCUTS}")
    story = pymupdf.Story(html="".join(html), user_css=CSS, archive=archive)
    tmp = WORK / "guida.pdf"
    writer = pymupdf.DocumentWriter(str(tmp))
    mediabox = pymupdf.paper_rect("a4")
    more = True
    while more:
        device = writer.begin_page(mediabox)
        more, _ = story.place(mediabox + (54, 50, -54, -56))
        story.draw(device)
        writer.end_page()
    writer.close()
    doc = pymupdf.open(tmp)
    for page in doc:   # piè di pagina, tranne la copertina
        if page.number:
            page.insert_text((54, 815), f"{tp.APP_NAME} – Guida illustrata", fontsize=8, color=(0.45, 0.45, 0.5))
            page.insert_text((520, 815), f"pagina {page.number + 1}", fontsize=8, color=(0.45, 0.45, 0.5))
    doc.set_metadata({"title": f"{tp.APP_NAME} – Guida illustrata", "author": tp.AUTHOR})
    doc.save(target, garbage=3, deflate=True)
    return len(doc)


if __name__ == "__main__":
    shots()
    pages = build_pdf(HERE / tp.GUIDE_NAME)
    print(f"Pronta: {HERE / tp.GUIDE_NAME} ({pages} pagine)")
    shutil.rmtree(WORK, ignore_errors=True)
