"""TimbroEasyPDF — visualizza, ruota e unisci PDF e applica il tuo timbro con un clic.

Avvio:       python timbro_pdf.py [file.pdf ...]
Dipendenze:  pip install pymupdf   (l'OCR usa i file della cartella tessdata,
             l'esportazione in Word usa Microsoft Word installato sul PC)
"""

import array
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import uuid
from collections import namedtuple
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont

import pymupdf

APP_NAME = "TimbroEasyPDF"
OLD_APP_NAMES = ("TimbroPDF",)   # nomi precedenti: impostazioni e timbri passano alla cartella nuova
APP_MUTEX = f"{APP_NAME}_in_esecuzione"   # lo cerca l'installatore per sapere se il programma è aperto
APP_VERSION = "2.1"
CONFIG_DIR = Path(os.environ.get("APPDATA") or Path.home()) / APP_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"
STAMPS_DIR = CONFIG_DIR / "timbri"           # una cartella per ogni timbro salvato
OLD_STAMP_FILE = CONFIG_DIR / "timbro.png"   # timbro unico della prima versione
# nell'exe i file allegati stanno nella cartella temporanea di PyInstaller
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
TESSDATA_DIR = BASE_DIR / "tessdata"          # lingue dell'OCR (Tesseract)
GUIDE_NAME = f"Guida {APP_NAME}.pdf"          # guida illustrata, accanto al programma

MM = 72 / 25.4                # punti PDF in un millimetro
THUMB_W, THUMB_H = 130, 170    # riquadro massimo delle miniature (px)
PAD = 16                       # margine attorno alla pagina nella vista
PREVIEW_W, PREVIEW_H = 210, 110
TILE_W, TILE_H = 97, 62        # riquadri della galleria dei timbri
GALLERY_H = 2 * (TILE_H + 26) + 6
MAX_STAMP_PX = 1600            # lato massimo dell'immagine del timbro salvata
MAX_RENDER_PX = 40_000_000     # limite di pixel per il rendering della pagina
UNDO_LIMIT = 30
UNDO_MAX_BYTES = 400 * 1024 * 1024
OCR_DPI = 300                  # risoluzione a cui si legge la pagina per l'OCR
OCR_MAX_PX = 60_000_000        # per le tavole grandi la risoluzione scende
OCR_LANGUAGES = {"Italiano e inglese": "ita+eng", "Italiano": "ita", "Inglese": "eng"}

TABBAR_BG, TAB_BG, TAB_ACTIVE_BG, TAB_DROP_BG = "#d3d4da", "#e6e7ec", "#ffffff", "#bcd4f6"
TEXT_TOOLS = ("edit_text", "add_text", "whiteout", "blackout")
TOOL_CURSORS = {"stamp": "crosshair", "edit_text": "xterm", "add_text": "xterm", "whiteout": "tcross",
                "blackout": "tcross"}
TOOL_HINTS = {
    "stamp": "Modalità timbro: clicca sul foglio per apporre il timbro. Esc per uscire.",
    "edit_text": "Modifica testo: passa sulle righe (si evidenziano in blu) e clicca quella da correggere. Esc per uscire.",
    "add_text": "Aggiungi testo: clicca dove vuoi scrivere. Esc per uscire.",
    "whiteout": "Bianchetto: trascina un rettangolo sull'area da cancellare. Esc per uscire.",
    "blackout": "Oscura: trascina un rettangolo sull'area da oscurare. Esc per uscire.",
}

PDF_TYPES = [("PDF", "*.pdf"), ("Tutti i file", "*.*")]
DOCX_TYPES = [("Documento Word", "*.docx")]
ADD_TYPES = [("PDF e immagini", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff *.bmp"), ("Tutti i file", "*.*")]
IMAGE_TYPES = [("Immagini", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.gif"), ("Tutti i file", "*.*")]

HELP_TEXT = """\
TIMBRO
1. «Aggiungi timbro…»: scegli PNG/JPG del timbro e dagli un nome. Puoi
   salvarne quanti vuoi (es. timbro e firma): l'immagine viene copiata
   nella cartella dell'app e resta disponibile anche se sposti il file.
   Clic su un timbro della galleria (o tasti 1–9) per sceglierlo.
   Ogni timbro ricorda le sue impostazioni e l'ultima posizione usata;
   all'avvio è selezionato l'ultimo timbro usato.
2. «Rimuovi sfondo bianco» rende trasparente il bianco attorno al timbro;
   alza la tolleranza se resta un alone grigio (tipico delle scansioni).
3. Imposta la larghezza in millimetri (l'altezza segue le proporzioni).
4. Scegli primo piano (timbro sopra il contenuto) o secondo piano
   (timbro sotto testo e linee, come un timbro vero sotto la firma).
   Nelle pagine scansionate il secondo piano resta nascosto sotto l'immagine.
5. Premi «Attiva modalità timbro» (o il tasto T): il timbro segue il mouse,
   clicca sul foglio per apporlo centrato in quel punto.
   Puoi cambiare pagina e continuare a cliccare. Esc per uscire.
6. «Ripeti su tutte le pagine» copia il timbro nella stessa posizione
   su tutte le altre pagine (o solo su quelle selezionate). In un nuovo
   documento usa la posizione dell'ultima volta.
7. Timbri già apposti: fuori dalla modalità timbro passaci sopra col mouse,
   trascinali per spostarli, doppio clic per modificarli; tasto destro per
   ruotarli a destra o a sinistra, oppure per eliminarli (anche Canc).

FIRMA
«Crea firma…»: disegnala con mouse, penna o dito, oppure scrivila con la
tastiera scegliendo font, dimensione e colore. Viene salvata nella galleria
e si appone come un timbro.

MODIFICA DEI TIMBRI
«Modifica…» (o doppio clic nella galleria) apre il timbro, qualunque sia:
ruotalo a destra o a sinistra, ritoccalo con matita, evidenziatore e gomma,
aggiungi una scritta con il font che vuoi, oppure aprilo nell'editor di
immagini di Windows (es. Paint): quando lo salvi lì, si aggiorna da solo.
Le firme create qui si possono anche riscrivere o ridisegnare; dopo un
ritocco a mano restano un'immagine. Nome e impostazioni restano gli stessi.
Le modifiche valgono per i timbri apposti da quel momento in poi; quelli già
sulla pagina si modificano a uno a uno con doppio clic (o tasto destro).
I pulsanti ↺ e ↻ ruotano al volo il timbro selezionato nella galleria.

DISPOSIZIONE
Trascina la maniglia ⋮ di una barra di pulsanti verso un lato della finestra
per agganciarla lì; trascina l'intestazione di un pannello (Pagine, Timbri)
per portarlo a destra o a sinistra, e il divisorio per allargarlo.
Disposizione e finestra vengono ricordate; Visualizza ▸ Ripristina disposizione.

ARCHIVIO DEI TIMBRI
Tutti i timbri stanno in un unico file (timbri.tpdf), creato da solo se manca.
Timbro ▸ Archivio dei timbri: spostalo in una cartella condivisa (es. Dropbox) e
sugli altri PC scegli «Usa un archivio esistente»; proteggilo con una password
(cifratura AES-256), che ogni PC può ricordare.

VISTA D'INSIEME E FILE TRASCINATI
F4 (o «Vista d'insieme»): le miniature di tutti i PDF aperti, una colonna per
PDF; trascina le pagine fra le colonne per copiarle. Doppio clic per aprirne una.
Trascina i PDF da Esplora risorse sulla finestra per aprirli.
File ▸ Usa TimbroEasyPDF per aprire i PDF: lo registra e apre la pagina di Windows
dove sceglierlo; i PDF aperti con doppio clic arrivano come nuove schede.

UNIRE E DIVIDERE
«Unisci schede…»: scegli quali PDF aperti unire e in che ordine; il risultato
si apre in una nuova scheda.
«Dividi…»: salva il documento in file da 1, 2 o N pagine, con il nome che
scegli e un numero progressivo (nome_01.pdf, nome_02.pdf…).

SCHEDE E PAGINE
• Ogni PDF aperto ha la sua scheda; le modifiche valgono solo per la scheda
  attiva, che ha il suo «Annulla». «Salva…» salva la scheda attiva.
• Clic sulle miniature per selezionare, Ctrl+clic / Maiusc+clic per più pagine.
• Trascina le miniature per riordinare le pagine. Per combinare PDF trascinale
  sulla scheda di un altro PDF: rilasciate sulla scheda vanno in fondo; se
  aspetti un attimo sulla scheda, questa si apre e le rilasci nel punto voluto.
  Rilasciate sul «+» creano un nuovo documento. Le pagine vengono copiate.
• Ruota, sposta ed elimina agiscono sulle pagine selezionate.
• «Aggiungi…» unisce altri PDF o immagini in fondo al documento
  (dal menu File anche dopo la pagina corrente).

TESTO
• Modifica testo: clicca una riga per correggerla. Il vecchio testo viene
  eliminato davvero e quello nuovo scritto nello stesso punto. Nelle
  scansioni spunta «Copri con il bianco lo sfondo».
• Aggiungi testo: clicca dove vuoi scrivere. Mentre scrivi il testo si vede
  già sulla pagina, con il font, la dimensione e il colore scelti. Le caselle
  si spostano trascinandole, doppio clic per correggerle, Canc per eliminarle.
• Le finestre per aprire, salvare, esportare e scegliere le immagini dei
  timbri ripartono ognuna dall'ultima cartella usata per quell'operazione.
• Bianchetto / Oscura: trascina un rettangolo; quanto c'è sotto viene
  cancellato davvero (i timbri restano: si eliminano a parte).
• OCR: riconosce il testo delle scansioni e lo aggiunge invisibile sopra la
  pagina, che resta identica: poi si può cercare, copiare, esportare in Word.
• Esporta in Word: usa Microsoft Word installato sul PC.

SCORCIATOIE
Ctrl+O apri · Ctrl+N nuovo · Ctrl+S salva · Ctrl+W chiudi scheda · Ctrl+Tab scheda dopo
Ctrl+Z annulla · Ctrl+Y ripristina · Ctrl+F cerca
Ctrl+R ruota a destra · Ctrl+Maiusc+R ruota a sinistra · Canc elimina pagine
PagSu/PagGiù cambia pagina · Ctrl+rotella zoom · trascina per spostare la vista
"""


def set_dpi_awareness():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


def place_dialog(top, parent):
    """Centra la finestra di dialogo sopra la finestra principale (o sullo schermo) e la rende modale."""
    top.update_idletasks()
    if parent.winfo_viewable():
        x = parent.winfo_rootx() + (parent.winfo_width() - top.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - top.winfo_height()) // 3
    else:   # all'avvio la finestra principale non è ancora a schermo
        x = (top.winfo_screenwidth() - top.winfo_width()) // 2
        y = (top.winfo_screenheight() - top.winfo_height()) // 3
    top.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    top.lift()
    top.grab_set()


def keep_clear(top, area):
    """Sposta la finestra di dialogo perché non copra area (x0, y0, x1, y1 sullo schermo), es. un'anteprima."""
    top.update_idletasks()
    match = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", top.geometry())
    if not match:
        return
    w, h, x, y = map(int, match.groups())
    h += 40                                   # barra del titolo
    x0, y0, x1, y1 = area
    if x + w < x0 or x > x1 or y + h < y0 or y > y1:
        return
    sw, sh, gap = top.winfo_screenwidth(), top.winfo_screenheight() - 40, 20
    for nx, ny in ((x, y1 + gap), (x, y0 - gap - h), (x1 + gap, y), (x0 - gap - w, y)):
        if nx >= 0 and ny >= 0 and nx + w <= sw and ny + h <= sh:
            top.geometry(f"+{round(nx)}+{round(ny)}")
            return


def photo_from_pixmap(pix):
    """Converte un Pixmap in immagine Tk (PNG se ha trasparenza, altrimenti PPM che è più veloce)."""
    return tk.PhotoImage(data=pix.tobytes("png" if pix.alpha else "ppm"))


def draw_checker(c, x0, y0, x1, y1, n=10, tags=()):
    """Scacchiera grigio chiaro: rende visibili le parti trasparenti di un'immagine."""
    x0, y0, x1, y1 = int(x0), int(y0), int(x1), int(y1)
    for y in range(y0, y1, n):
        for x in range(x0 + n * ((y - y0) // n % 2), x1, 2 * n):
            c.create_rectangle(x, y, min(x + n, x1), min(y + n, y1), fill="#e9e9ee", outline="", tags=tags)


def icon_pixmap(size=64):
    """Icona dell'app: un foglio con un timbro rosso."""
    doc = pymupdf.open()
    page = doc.new_page(width=64, height=64)
    page.draw_rect(pymupdf.Rect(10, 3, 46, 57), color=(0.35, 0.35, 0.4), fill=(1, 1, 1), width=2.5)
    for y in (13, 19, 25):
        page.draw_line((16, y), (40, y), color=(0.6, 0.6, 0.65), width=2)
    page.draw_circle((44, 45), 15, color=(0.78, 0.1, 0.1), fill=(1, 0.92, 0.92), width=4)
    page.draw_circle((44, 45), 7, color=(0.78, 0.1, 0.1), width=2)
    return page.get_pixmap(matrix=pymupdf.Matrix(size / 64, size / 64), alpha=True)


def white_lut(tolerance):
    """Tabella luminosità -> opacità: i pixel chiari diventano trasparenti, con bordi sfumati.

    tolerance va da 0 a 100: più è alta, più grigi chiari vengono considerati sfondo.
    """
    hi = 252 - 0.8 * tolerance
    lo = hi - 70
    return bytes(0 if v >= hi else 255 if v <= lo else round(255 * (hi - v) / (hi - lo)) for v in range(256))


def flatten_on_white(pix):
    """L'immagine come appare appoggiata su carta bianca, senza canale di trasparenza."""
    if not pix.alpha:
        return pix
    doc = pymupdf.open()
    page = doc.new_page(width=pix.width, height=pix.height)
    page.insert_image(page.rect, pixmap=pix)
    return page.get_pixmap(alpha=False)


# ---------------------------------------------------------------- firme
SCRIPT_FONTS = ("Segoe Script", "Ink Free", "Lucida Handwriting Italic", "Brush Script MT Italic",
                "Freestyle Script", "Kunstler Script", "Vladimir Script", "Edwardian Script ITC",
                "Monotype Corsiva", "Mistral", "Segoe Print", "Gabriola")   # font calligrafici, in cima all'elenco
INK_COLORS = {"Blu": "#1f3a93", "Nero": "#111111", "Rosso": "#b3261e"}
SIGNATURE_ZOOM = 4        # pixel per punto delle firme create (≈ 290 dpi)
PAD_WIDTH_MM = 120        # larghezza "reale" del riquadro in cui si disegna la firma


def installed_fonts():
    """Font installati in Windows (di sistema e dell'utente): nome -> file."""
    fonts = {}
    try:
        import winreg
    except ImportError:
        return fonts
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    places = ((winreg.HKEY_LOCAL_MACHINE, Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"),
              (winreg.HKEY_CURRENT_USER, Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"))
    for hive, folder in places:
        try:
            key = winreg.OpenKey(hive, key_path)
        except OSError:
            continue
        with key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                name, value, _ = winreg.EnumValue(key, i)
                name = re.sub(r"\s*\((TrueType|OpenType)\)\s*$", "", name).strip()
                file = Path(value) if os.path.isabs(value) else folder / value
                if name and file.suffix.lower() in (".ttf", ".otf", ".ttc") and file.exists():
                    fonts[name] = str(file)
    return fonts


def hex_to_rgb(color):
    return tuple(int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))


def ink_bbox(pix):
    """Riquadro dei pixel non trasparenti, oppure None se l'immagine è vuota."""
    w, n = pix.width, pix.n
    alpha = pix.samples[n - 1::n]
    rows = [alpha[y * w:(y + 1) * w] for y in range(pix.height)]
    used = [y for y, row in enumerate(rows) if row.strip(b"\0")]
    if not used:
        return None
    x0 = min(len(rows[y]) - len(rows[y].lstrip(b"\0")) for y in used)
    x1 = max(len(rows[y].rstrip(b"\0")) for y in used)
    return pymupdf.Rect(x0, used[0], x1, used[-1] + 1)


def render_ink(page, margin=2):
    """Quanto disegnato sulla pagina, su fondo trasparente e ritagliato attorno all'inchiostro.

    Restituisce (pixmap, riquadro in punti) oppure (None, None) se la pagina è vuota.
    """
    box = ink_bbox(page.get_pixmap(alpha=True))
    if box is None:
        return None, None
    clip = (box + (-margin, -margin, margin, margin)) & page.rect
    return page.get_pixmap(matrix=pymupdf.Matrix(SIGNATURE_ZOOM, SIGNATURE_ZOOM), clip=clip, alpha=True), clip


def smooth(points, iterations=2):
    """Arrotonda un tratto disegnato col mouse (algoritmo di Chaikin)."""
    for _ in range(iterations):
        if len(points) < 3:
            break
        out = [points[0]]
        for (x0, y0), (x1, y1) in zip(points, points[1:]):
            out += [(0.75 * x0 + 0.25 * x1, 0.75 * y0 + 0.25 * y1), (0.25 * x0 + 0.75 * x1, 0.25 * y0 + 0.75 * y1)]
        out.append(points[-1])
        points = out
    return points


def drawn_signature(strokes, size, width, color):
    """Firma disegnata a mano: strokes sono elenchi di punti in pixel del riquadro grande size."""
    doc = pymupdf.open()
    page = doc.new_page(width=size[0], height=size[1])
    shape = page.new_shape()
    for stroke in strokes:
        points = smooth(stroke) if len(stroke) > 1 else stroke + [(stroke[0][0] + 0.3, stroke[0][1])]
        shape.draw_polyline(points)
        shape.finish(color=color, width=width, closePath=False, lineCap=1, lineJoin=1)
    shape.commit()
    return render_ink(page)


def text_signature(text, fontfile, fontsize, color):
    """Firma scritta con un font: il riquadro restituito è in punti alla dimensione scelta."""
    font = pymupdf.Font(fontfile=fontfile) if fontfile else pymupdf.Font("helv")
    pad = fontsize  # i font calligrafici escono spesso dai margini nominali delle lettere
    doc = pymupdf.open()
    page = doc.new_page(width=font.text_length(text, fontsize=fontsize) + 2 * pad, height=fontsize * 3)
    writer = pymupdf.TextWriter(page.rect)
    writer.append((pad, fontsize * 2), text, font=font, fontsize=fontsize)
    writer.write_text(page, color=color)
    return render_ink(page)


class SignatureDialog:
    """Crea una firma da salvare tra i timbri: disegnata a mano oppure scritta con un font.

    Con source (i dati salvati di una firma creata qui) la riapre per modificarla.
    """

    def __init__(self, parent, source=None, name=None):
        self.result = None                      # (pixmap, nome, larghezza in mm, dati per rimodificarla)
        self.s = s = parent.winfo_fpixels("1i") / 96
        self.color = (source or {}).get("color", INK_COLORS["Blu"])
        self.strokes, self.lines = [], []
        self.fonts = installed_fonts()
        self._preview_job = None
        self._preview_photo = None
        self.W, self.H = round(680 * s), round(230 * s)

        top = self.top = tk.Toplevel(parent)
        top.title("Modifica firma" if source else "Crea firma")
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=12)
        body.pack(fill=tk.BOTH, expand=True)
        self.tabs = ttk.Notebook(body)
        self.tabs.pack(fill=tk.X)
        draw, write = ttk.Frame(self.tabs, padding=8), ttk.Frame(self.tabs, padding=8)
        self.tabs.add(draw, text="  Disegna  ")
        self.tabs.add(write, text="  Scrivi con la tastiera  ")

        # disegno a mano libera
        ttk.Label(draw, text="Firma nel riquadro con il mouse, la penna o il dito.").pack(anchor="w")
        self.pad = tk.Canvas(draw, width=self.W, height=self.H, bg="white", cursor="pencil",
                             highlightthickness=1, highlightbackground="#b8b8c0")
        self.pad.pack(pady=6)
        self.pad.create_line(24 * s, self.H * 0.72, self.W - 24 * s, self.H * 0.72, fill="#d4d4dc", dash=(6, 4))
        self.pad.bind("<Button-1>", self._stroke_start)
        self.pad.bind("<B1-Motion>", self._stroke_move)
        row = ttk.Frame(draw)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Spessore:").pack(side=tk.LEFT)
        self.thickness = ttk.Scale(row, from_=1, to=8, length=round(160 * s))
        self.thickness.set(3)
        self.thickness.configure(command=lambda _v: self.pad.itemconfigure("ink", width=self._thickness_px()))
        self.thickness.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(row, text="Annulla tratto", command=self._undo_stroke).pack(side=tk.RIGHT)
        ttk.Button(row, text="Cancella tutto", command=self._clear).pack(side=tk.RIGHT, padx=4)

        # testo da tastiera
        r1 = ttk.Frame(write)
        r1.pack(fill=tk.X)
        ttk.Label(r1, text="Testo:").pack(side=tk.LEFT)
        self.text_var = tk.StringVar()
        self.text_entry = ttk.Entry(r1, textvariable=self.text_var, font=("Segoe UI", 12))
        self.text_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        r2 = ttk.Frame(write)
        r2.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(r2, text="Font:").pack(side=tk.LEFT)
        names = [n for n in SCRIPT_FONTS if n in self.fonts]
        names += sorted((n for n in self.fonts if n not in names), key=str.lower)
        self.font_var = tk.StringVar(value=names[0] if names else "Helvetica")
        ttk.Combobox(r2, textvariable=self.font_var, values=names or ["Helvetica"], state="readonly",
                     width=34, height=20).pack(side=tk.LEFT, padx=(6, 14))
        ttk.Label(r2, text="Dimensione (pt):").pack(side=tk.LEFT)
        self.size_var = tk.StringVar(value="28")
        ttk.Spinbox(r2, from_=6, to=150, increment=1, width=5, textvariable=self.size_var).pack(side=tk.LEFT,
                                                                                            padx=(6, 0))
        self.PH = self.H - round(34 * s)
        self.preview = tk.Canvas(write, width=self.W, height=self.PH, bg="white",
                                 highlightthickness=1, highlightbackground="#b8b8c0")
        self.preview.pack(pady=(8, 0))
        for var in (self.text_var, self.font_var, self.size_var):
            var.trace_add("write", self._schedule_preview)

        # colore e nome, comuni alle due schede
        common = ttk.Frame(body)
        common.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(common, text="Colore:").pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value=next((n for n, c in INK_COLORS.items() if c == self.color), ""))
        for color_name in INK_COLORS:
            ttk.Radiobutton(common, text=color_name, value=color_name, variable=self.color_var,
                            command=lambda: self._set_color(INK_COLORS[self.color_var.get()])).pack(side=tk.LEFT,
                                                                                                    padx=(6, 0))
        ttk.Button(common, text="Altro colore…", command=self._pick_color).pack(side=tk.LEFT, padx=(10, 6))
        self.swatch = tk.Label(common, width=3, bg=self.color, relief=tk.SOLID, bd=1)
        self.swatch.pack(side=tk.LEFT)
        self.name_var = tk.StringVar(value=name or "Firma")
        ttk.Entry(common, textvariable=self.name_var, width=18).pack(side=tk.RIGHT)
        ttk.Label(common, text="Nome nella galleria:").pack(side=tk.RIGHT, padx=(20, 6))

        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(buttons, text="Chiudi", command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Salva modifiche" if source else "Salva nella galleria",
                   command=self._save).pack(side=tk.RIGHT, padx=6)
        if source:
            self._load_source(source)

        top.bind("<Control-z>", lambda e: self._undo_stroke())
        top.bind("<Escape>", lambda e: top.destroy())
        self.tabs.bind("<<NotebookTabChanged>>",
                       lambda e: self.text_entry.focus_set() if self.tabs.index("current") == 1 else None)
        self._render_preview()
        place_dialog(top, parent)

    def _load_source(self, source):
        """Rimette nella finestra una firma salvata: il testo con il suo font, o i tratti disegnati."""
        if source["kind"] == "typed":
            self.tabs.select(1)
            self.text_var.set(source["text"])
            if source["font"] in self.fonts:
                self.font_var.set(source["font"])
            self.size_var.set(f"{source['size']:g}")
        else:
            self.thickness.set(source.get("thickness", 3))
            for stroke in source["strokes"]:
                points = [(x * self.W, y * self.H) for x, y in stroke]
                self.strokes.append(points)
                self.lines.append(self.pad.create_line(*[c for p in points for c in p],
                                                       *([points[0][0] + 0.5, points[0][1]] if len(points) == 1
                                                         else []),
                                                       fill=self.color, width=self._thickness_px(),
                                                       capstyle=tk.ROUND, joinstyle=tk.ROUND, smooth=True,
                                                       splinesteps=8, tags="ink"))

    # disegno
    def _thickness_px(self):
        return float(self.thickness.get()) * self.s

    def _stroke_start(self, event):
        self.strokes.append([(event.x, event.y)])
        self.lines.append(self.pad.create_line(event.x, event.y, event.x + 0.5, event.y, fill=self.color,
                                               width=self._thickness_px(), capstyle=tk.ROUND, joinstyle=tk.ROUND,
                                               smooth=True, splinesteps=8, tags="ink"))

    def _stroke_move(self, event):
        if not self.strokes:
            return
        stroke = self.strokes[-1]
        x, y = stroke[-1]
        if (event.x - x) ** 2 + (event.y - y) ** 2 < 4:
            return
        stroke.append((event.x, event.y))
        self.pad.coords(self.lines[-1], *[c for point in stroke for c in point])

    def _undo_stroke(self):
        if self.strokes:
            self.strokes.pop()
            self.pad.delete(self.lines.pop())

    def _clear(self):
        self.strokes.clear()
        self.lines.clear()
        self.pad.delete("ink")

    # colore
    def _pick_color(self):
        color = colorchooser.askcolor(color=self.color, parent=self.top, title="Colore della firma")[1]
        if color:
            self.color_var.set("")
            self._set_color(color)

    def _set_color(self, color):
        self.color = color
        self.swatch.configure(bg=color)
        self.pad.itemconfigure("ink", fill=color)
        self._schedule_preview()

    # testo
    def _text_signature(self):
        text = self.text_var.get().strip()
        try:
            size = float(self.size_var.get().replace(",", "."))
        except ValueError:
            return None, None
        if not text or not 4 <= size <= 300:
            return None, None
        return text_signature(text, self.fonts.get(self.font_var.get()), size, hex_to_rgb(self.color))

    def _schedule_preview(self, *_):
        if self._preview_job:
            self.top.after_cancel(self._preview_job)
        self._preview_job = self.top.after(120, self._render_preview)

    def _render_preview(self):
        self._preview_job = None
        c = self.preview
        c.delete("all")
        pix, box = self._text_signature()
        if pix is None:
            c.create_text(self.W / 2, self.PH / 2, text="Scrivi il testo della firma qui sopra", fill="#999")
            return
        real = self.s * 96 / 72 / SIGNATURE_ZOOM   # dimensione reale sullo schermo, se ci sta
        f = min(real, (self.W - 20) / pix.width, (self.PH - 30) / pix.height)
        self._preview_photo = photo_from_pixmap(pymupdf.Pixmap(pix, max(1, round(pix.width * f)),
                                                               max(1, round(pix.height * f))))
        c.create_image(self.W / 2, (self.PH - 14) / 2, image=self._preview_photo)
        c.create_text(8, self.PH - 6, anchor="sw", fill="#888",
                      text=f"Nel PDF: {box.width / MM:.0f} × {box.height / MM:.0f} mm"
                           + ("" if f >= real else "  (anteprima rimpicciolita)"))

    def _save(self):
        if self.tabs.index("current") == 0:
            if not self.strokes:
                messagebox.showinfo(APP_NAME, "Disegna prima la firma nel riquadro.", parent=self.top)
                return
            pix, box = drawn_signature(self.strokes, (self.W, self.H), self._thickness_px(), hex_to_rgb(self.color))
            width_mm = min(max(box.width / self.W * PAD_WIDTH_MM, 15), 150)
            # tratti in frazioni del riquadro: si ridisegnano uguali anche con un'altra scala dello schermo
            source = {"kind": "drawn", "color": self.color, "thickness": round(float(self.thickness.get()), 2),
                      "strokes": [[[round(x / self.W, 4), round(y / self.H, 4)] for x, y in stroke]
                                  for stroke in self.strokes]}
        else:
            pix, box = self._text_signature()
            if pix is None:
                messagebox.showinfo(APP_NAME, "Scrivi prima il testo della firma.", parent=self.top)
                return
            width_mm = box.width / MM
            source = {"kind": "typed", "text": self.text_var.get().strip(), "font": self.font_var.get(),
                      "size": float(self.size_var.get().replace(",", ".")), "color": self.color}
        self.result = (pix, self.name_var.get().strip() or "Firma", round(width_mm, 1), source)
        self.top.destroy()


# ---------------------------------------------------------------- ritocco dei timbri
EDIT_TOOLS = (("pencil", "✎ Matita"), ("highlight", "▬ Evidenziatore"), ("eraser", "◌ Gomma"), ("text", "T  Testo"))
EDIT_HINTS = {
    "pencil": "Matita: disegna tenendo premuto il tasto sinistro del mouse.",
    "highlight": "Evidenziatore: passa sopra tenendo premuto il tasto; il colore è semitrasparente.",
    "eraser": "Gomma: quello su cui passi diventa trasparente (bianco, se il timbro toglie lo sfondo bianco).",
    "text": "Testo: clicca dove deve iniziare la scritta, poi scegli font, dimensione e colore.",
}
EDIT_COLORS = {"pencil": "#1f3a93", "highlight": "#ffd600", "text": "#111111"}
EDIT_SWATCHES = ("#111111", "#1f3a93", "#b3261e", "#2e7d32", "#ffd600")
BRUSH_PT = {"pencil": 0.35, "highlight": 1.6, "eraser": 1.2}   # spessore in punti per ogni tacca del cursore
HIGHLIGHT_OPACITY = 0.4
EDIT_UNDO = 20


def stamp_pixmap(pix):
    """Immagine pronta per un timbro: RGB (con o senza trasparenza), lato massimo MAX_STAMP_PX."""
    if pix.colorspace is None or pix.colorspace.n != 3:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    longest = max(pix.width, pix.height)
    if longest > MAX_STAMP_PX:
        f = MAX_STAMP_PX / longest
        pix = pymupdf.Pixmap(pix, round(pix.width * f), round(pix.height * f))
    return pix


def rotate_pixmap(pix, angle):
    """L'immagine ruotata di un multiplo di 90° in senso orario (negativo: antiorario), senza perdite."""
    angle %= 360
    if not angle:
        return pix
    if pix.n not in (1, 4):   # RGB: si aggiunge l'alpha, così ogni pixel occupa 4 byte e si sposta in blocco
        return pymupdf.Pixmap(rotate_pixmap(pymupdf.Pixmap(pix, 1), angle), 0)
    w, h = pix.width, pix.height
    a = array.array("I" if pix.n == 4 else "B")
    a.frombytes(pix.samples)
    if angle == 180:
        out = a[::-1]
    else:
        out = array.array(a.typecode)
        for x in range(w) if angle == 90 else range(w - 1, -1, -1):
            column = a[x::w]                      # dall'alto in basso
            out.extend(column[::-1] if angle == 90 else column)
        w, h = h, w
    return pymupdf.Pixmap(pix.colorspace or pymupdf.csGRAY, w, h, out.tobytes(), pix.alpha)


def split_alpha(pix):
    """(colori, trasparenza): l'immagine RGB senza alpha e la maschera grigia (None se è opaca).

    MuPDF tiene i colori già moltiplicati per l'alpha: qui tornano ai valori veri, così ritoccando
    l'immagine i bordi sfumati non si scuriscono.
    """
    if not pix.alpha:
        return pix, None
    alpha = pymupdf.Pixmap(None, pix).samples
    rgb = bytearray(pymupdf.Pixmap(pix, 0).samples)
    for m in re.finditer(rb"\x00+", alpha):       # trasparente: il colore non conta, meglio bianco
        rgb[3 * m.start():3 * m.end()] = b"\xff" * (3 * (m.end() - m.start()))
    for m in re.finditer(rb"[\x01-\xfe]", alpha):
        i, a = m.start(), alpha[m.start()]
        rgb[3 * i:3 * i + 3] = bytes(min(255, (v * 255 + a // 2) // a) for v in rgb[3 * i:3 * i + 3])
    return (pymupdf.Pixmap(pymupdf.csRGB, pix.width, pix.height, bytes(rgb), False),
            pymupdf.Pixmap(pymupdf.csGRAY, pix.width, pix.height, alpha, False))


def join_alpha(colors, mask):
    """Il contrario di split_alpha: senza maschera (o se è tutta opaca) l'immagine resta senza alpha."""
    if mask is None or mask.samples.count(255) == len(mask.samples):
        return colors
    return pymupdf.Pixmap(colors, mask)


def opaque_mask(pix):
    return pymupdf.Pixmap(pymupdf.csGRAY, pix.width, pix.height, b"\xff" * (pix.width * pix.height), False)


def paint(pix, draw):
    """Disegna sull'immagine (senza alpha) con draw(page): la pagina è grande come l'immagine, 1 punto = 1 pixel."""
    doc = pymupdf.open()
    page = doc.new_page(width=pix.width, height=pix.height)
    page.insert_image(page.rect, pixmap=pix)
    draw(page)
    return page.get_pixmap(colorspace=pix.colorspace, alpha=False)


def stroke_drawer(points, color, width, opacity=1.0):
    """Un tratto a mano libera (punti in pixel dell'immagine), con le estremità arrotondate."""
    def draw(page):
        shape = page.new_shape()
        shape.draw_polyline(points if len(points) > 1 else points + [(points[0][0] + 0.3, points[0][1])])
        shape.finish(color=color, width=width, closePath=False, lineCap=1, lineJoin=1, stroke_opacity=opacity)
        shape.commit()
    return draw


def text_drawer(x, y, params, fontfile, scale, color):
    """Testo con l'angolo in alto a sinistra in (x, y); scale: pixel dell'immagine per punto di carattere."""
    size = params["size"] * scale
    ascender = text_font(fontfile).ascender

    def draw(page):
        page.insert_text((x, y + ascender * size), params["text"], fontsize=size, lineheight=LINE_HEIGHT,
                         fontname="F0" if fontfile else "helv", fontfile=fontfile, color=color)
    return draw


def without_white(pix, tolerance):
    """Come Stamp.process: lo sfondo bianco diventa trasparente (senza cache: per le anteprime)."""
    flat = flatten_on_white(pix)
    s = flat.samples
    out = pymupdf.Pixmap(flat, 1)
    out.set_alpha(bytes(map(min, s[0::3], s[1::3], s[2::3])).translate(white_lut(tolerance)))
    return out


class StampEditor:
    """Ritocco di un timbro: ruota, matita, evidenziatore, gomma e testo, oppure l'editor di immagini di Windows.

    original è l'immagine di partenza e rotation la rotazione (gradi in senso orario) con cui la si
    vede. Le firme create nel programma (source) si possono anche riscrivere o ridisegnare finché non
    si ritocca l'immagine a mano; il disegno rifatto si ruota di source_rotation.
    remove_white None nasconde l'opzione dello sfondo (timbri già apposti: l'immagine è quella finale).
    result: None se non è cambiato niente, altrimenti dict con image (l'immagine finale, ruotata),
    original (None se ritoccata a mano), rotation, width_mm, source, remove_white, name.
    """

    def __init__(self, parent, fonts, original, width_mm, title, name=None, rotation=0, source=None,
                 source_rotation=0, remove_white=None, tolerance=10):
        self.result = None
        self.fonts = fonts
        self.original = stamp_pixmap(original)
        self.rotation = rotation % 360
        self.width_mm = width_mm
        self.source, self.source_rotation = source, source_rotation
        self.edited = False            # pixel ritoccati: la firma non si ricrea più dal suo testo o disegno
        self.changed = False
        self.tolerance = tolerance
        self.colors, self.mask = split_alpha(rotate_pixmap(self.original, self.rotation))
        self.undo_stack, self.redo_stack = [], []
        self.tool_colors = dict(EDIT_COLORS)
        self.text_style = {"font": "Arial" if "Arial" in fonts else next(iter(sorted(fonts)), "Helvetica"),
                           "size": 10.0}
        self.stroke = None
        self.f, self.ox, self.oy, self._geometry = 1.0, 0, 0, None
        self._photo = self._text_photo = None
        self._ext, self._ext_job = None, None
        self._initial = (name, bool(remove_white))
        s = parent.winfo_fpixels("1i") / 96
        self.CW, self.CH = round(720 * s), round(360 * s)

        top = self.top = tk.Toplevel(parent)
        top.title(title)
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=10)
        body.pack(fill=tk.BOTH, expand=True)

        bar = ttk.Frame(body)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="↺ Ruota a sinistra", command=lambda: self.rotate(-90)).pack(side=tk.LEFT)
        ttk.Button(bar, text="Ruota a destra ↻", command=lambda: self.rotate(90)).pack(side=tk.LEFT, padx=(4, 14))
        self.tool_var = tk.StringVar(value="pencil")
        for key, label in EDIT_TOOLS:
            ttk.Radiobutton(bar, text=label, value=key, variable=self.tool_var, style="Toolbutton",
                            command=self._tool_changed).pack(side=tk.LEFT, padx=1, ipadx=4)
        ttk.Button(bar, text="Ripristina ↷", command=self.redo).pack(side=tk.RIGHT)
        ttk.Button(bar, text="↶ Annulla", command=self.undo).pack(side=tk.RIGHT, padx=4)

        row = ttk.Frame(body)
        row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(row, text="Colore:").pack(side=tk.LEFT)
        self.swatch = tk.Label(row, width=3, relief=tk.SOLID, bd=1, cursor="hand2")
        self.swatch.pack(side=tk.LEFT, padx=(6, 8))
        self.swatch.bind("<Button-1>", lambda e: self._pick_color())
        for color in EDIT_SWATCHES:
            chip = tk.Label(row, width=2, bg=color, relief=tk.RAISED, bd=1, cursor="hand2")
            chip.pack(side=tk.LEFT, padx=1)
            chip.bind("<Button-1>", lambda e, c=color: self._set_color(c))
        ttk.Label(row, text="Spessore:").pack(side=tk.LEFT, padx=(18, 0))
        self.size_scale = ttk.Scale(row, from_=1, to=10, length=round(150 * s))
        self.size_scale.set(3)
        self.size_scale.pack(side=tk.LEFT, padx=(6, 0))
        self.size_label = ttk.Label(row, foreground="#555")
        self.size_label.pack(side=tk.RIGHT)

        self.canvas = tk.Canvas(body, width=self.CW, height=self.CH, bg="white", highlightthickness=1,
                                highlightbackground="#b8b8c0")
        self.canvas.pack(pady=(8, 0))
        for seq, fn in (("<Button-1>", self._press), ("<B1-Motion>", self._drag), ("<ButtonRelease-1>", self._release),
                        ("<Motion>", self._hover), ("<Leave>", lambda e: self.canvas.delete("brush"))):
            self.canvas.bind(seq, fn)
        self.hint = ttk.Label(body, foreground="#555")
        self.hint.pack(anchor="w", pady=(4, 0))
        self.note = ttk.Label(body, foreground="#a15c00", wraplength=self.CW, justify=tk.LEFT)
        self.note.pack(anchor="w")

        row = ttk.Frame(body)
        row.pack(fill=tk.X, pady=(8, 0))
        self.source_btn = None
        if source:
            self.source_btn = ttk.Button(row, command=self._edit_source,
                                         text="Modifica il testo della firma…" if source["kind"] == "typed"
                                         else "Ridisegna la firma…")
            self.source_btn.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row, text="Apri con l'editor di immagini…", command=self._open_external).pack(side=tk.LEFT)
        ttk.Button(row, text="Sostituisci immagine…", command=self._replace).pack(side=tk.LEFT, padx=6)
        self.white_var = tk.BooleanVar(value=bool(remove_white))
        if remove_white is not None:
            ttk.Checkbutton(row, text="Rimuovi sfondo bianco", variable=self.white_var,
                            command=self._redraw).pack(side=tk.RIGHT)

        bottom = ttk.Frame(body)
        bottom.pack(fill=tk.X, pady=(12, 0))
        self.name_var = tk.StringVar(value=name or "")
        if name is not None:
            ttk.Label(bottom, text="Nome nella galleria:").pack(side=tk.LEFT)
            ttk.Entry(bottom, textvariable=self.name_var, width=24).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(bottom, text="Chiudi senza salvare", command=self._close).pack(side=tk.RIGHT)
        ttk.Button(bottom, text="Salva", command=self._save).pack(side=tk.RIGHT, padx=6)

        top.bind("<Control-z>", lambda e: self.undo())
        top.bind("<Control-y>", lambda e: self.redo())
        top.bind("<Escape>", lambda e: self._close())
        top.protocol("WM_DELETE_WINDOW", self._close)
        top.bind("<Destroy>", self._on_destroy)
        self._tool_changed()
        self._redraw()
        place_dialog(top, parent)

    # ------------------------------------------------------------ stato e annulla
    @property
    def tool(self):
        return self.tool_var.get()

    def _state(self):
        return self.colors, self.mask, self.rotation, self.width_mm, self.edited, self.original, self.source

    def _push(self):
        self.undo_stack = self.undo_stack[-(EDIT_UNDO - 1):] + [self._state()]
        self.redo_stack.clear()

    def _restore(self, state):
        self.colors, self.mask, self.rotation, self.width_mm, self.edited, self.original, self.source = state
        self._after_change()

    def undo(self):
        if self.undo_stack:
            self.redo_stack.append(self._state())
            self._restore(self.undo_stack.pop())

    def redo(self):
        if self.redo_stack:
            self.undo_stack.append(self._state())
            self._restore(self.redo_stack.pop())

    def _after_change(self):
        self.changed = True
        if self.source_btn:
            self.source_btn.state(["disabled"] if self.edited else ["!disabled"])
        self.note.configure(text="La firma ora è un'immagine ritoccata: il testo (o il disegno) non si può più "
                                 "modificare come prima. Annulla (Ctrl+Z) per tornare indietro."
                            if self.edited and self.source else "")
        self._redraw()

    def _pixels_changed(self):
        self.edited = True
        self._after_change()

    # ------------------------------------------------------------ vista
    def _px_per_pt(self):
        return self.colors.width / (self.width_mm * MM)

    def _brush_px(self):
        return max(1.0, float(self.size_scale.get()) * BRUSH_PT.get(self.tool, 1) * self._px_per_pt())

    def _to_image(self, x, y):
        return (x - self.ox) / self.f, (y - self.oy) / self.f

    def _redraw(self):
        c = self.canvas
        pix = join_alpha(self.colors, self.mask)
        w, h = pix.width, pix.height
        self.f = min((self.CW - 24) / w, (self.CH - 24) / h, 4)
        dw, dh = max(1, round(w * self.f)), max(1, round(h * self.f))
        self.ox, self.oy = (self.CW - dw) // 2, (self.CH - dh) // 2
        if self._geometry != (dw, dh):   # la scacchiera si ridisegna solo se cambia il riquadro
            self._geometry = (dw, dh)
            c.delete("checker")
            draw_checker(c, self.ox, self.oy, self.ox + dw, self.oy + dh, tags="checker")
        small = pymupdf.Pixmap(pix, dw, dh)
        if self.white_var.get():
            small = without_white(small, self.tolerance)
        self._photo = photo_from_pixmap(small)
        c.delete("image", "frame")
        c.create_image(self.ox, self.oy, anchor="nw", image=self._photo, tags="image")
        c.create_rectangle(self.ox - 1, self.oy - 1, self.ox + dw, self.oy + dh, outline="#c8c8d0", tags="frame")
        c.tag_raise("brush")
        self.size_label.configure(text=f"Nel PDF: {self.width_mm:.0f} × {self.width_mm * h / w:.0f} mm")

    def _tool_changed(self):
        self.hint.configure(text=EDIT_HINTS[self.tool])
        color = self.tool_colors.get(self.tool)
        self.swatch.configure(bg=color or "#ffffff", text="" if color else "—")
        self.canvas.configure(cursor="xterm" if self.tool == "text" else "crosshair")
        self.canvas.delete("brush")

    def _set_color(self, color):
        if self.tool in self.tool_colors:
            self.tool_colors[self.tool] = color
            self.swatch.configure(bg=color)

    def _pick_color(self):
        if self.tool in self.tool_colors:
            color = colorchooser.askcolor(color=self.tool_colors[self.tool], parent=self.top, title="Colore")[1]
            self.top.grab_set()
            if color:
                self._set_color(color)

    def _hover(self, event):
        self.canvas.delete("brush")
        if self.tool != "text":
            r = self._brush_px() * self.f / 2
            self.canvas.create_oval(event.x - r, event.y - r, event.x + r, event.y + r, outline="#555",
                                    dash=(2, 2), tags="brush")

    # ------------------------------------------------------------ disegno
    def _press(self, event):
        if self.tool == "text":
            self._add_text(event.x, event.y)
            return
        self.stroke = [self._to_image(event.x, event.y)]
        color = "#d8d8de" if self.tool == "eraser" else self.tool_colors[self.tool]
        self._line = self.canvas.create_line(event.x, event.y, event.x + 0.5, event.y, fill=color,
                                             width=self._brush_px() * self.f, capstyle=tk.ROUND,
                                             joinstyle=tk.ROUND, tags="stroke",
                                             stipple="gray50" if self.tool == "highlight" else "")

    def _drag(self, event):
        self._hover(event)
        if not self.stroke:
            return
        x, y = self._to_image(event.x, event.y)
        lx, ly = self.stroke[-1]
        if (x - lx) ** 2 + (y - ly) ** 2 < (1.5 / self.f) ** 2:
            return
        self.stroke.append((x, y))
        self.canvas.coords(self._line, *[v for px, py in self.stroke
                                         for v in (self.ox + px * self.f, self.oy + py * self.f)])

    def _release(self, _event):
        points, self.stroke = self.stroke, None
        self.canvas.delete("stroke")
        if not points:
            return
        width = self._brush_px()
        self._push()
        if self.tool == "eraser":
            self.mask = paint(self.mask or opaque_mask(self.colors), stroke_drawer(points, (0, 0, 0), width))
            self.colors = paint(self.colors, stroke_drawer(points, (1, 1, 1), width))
        else:
            opacity = HIGHLIGHT_OPACITY if self.tool == "highlight" else 1.0
            self.colors = paint(self.colors, stroke_drawer(points, hex_to_rgb(self.tool_colors[self.tool]), width,
                                                           opacity))
            if self.mask is not None:
                self.mask = paint(self.mask, stroke_drawer(points, (1, 1, 1), width, opacity))
        self._pixels_changed()

    def _add_text(self, x, y):
        ix, iy = self._to_image(x, y)
        if not (0 <= ix <= self.colors.width and 0 <= iy <= self.colors.height):
            return
        scale = self._px_per_pt()

        def preview(params):
            self.canvas.delete("textpreview")
            if not params:
                return
            doc = pymupdf.open()
            page = doc.new_page(width=self.colors.width, height=self.colors.height)
            text_drawer(ix, iy, params, self.fonts.get(params["font"]), scale, hex_to_rgb(params["color"]))(page)
            self._text_photo = photo_from_pixmap(page.get_pixmap(matrix=pymupdf.Matrix(self.f, self.f), alpha=True))
            self.canvas.create_image(self.ox, self.oy, anchor="nw", image=self._text_photo, tags="textpreview")

        dialog = TextDialog(self.top, self.fonts, "Testo sul timbro",
                            dict(self.text_style, text="", color=self.tool_colors["text"]), preview=preview)
        c = self.canvas
        keep_clear(dialog.top, (c.winfo_rootx(), c.winfo_rooty(), c.winfo_rootx() + self.CW,
                                c.winfo_rooty() + self.CH))
        self.top.wait_window(dialog.top)
        self.top.grab_set()
        self.canvas.delete("textpreview")
        r = dialog.result
        if not r:
            return
        self.text_style = {"font": r["font"], "size": r["size"]}
        self.tool_colors["text"] = r["color"]
        self._tool_changed()
        fontfile = self.fonts.get(r["font"])
        self._push()
        self.colors = paint(self.colors, text_drawer(ix, iy, r, fontfile, scale, hex_to_rgb(r["color"])))
        if self.mask is not None:
            self.mask = paint(self.mask, text_drawer(ix, iy, r, fontfile, scale, (1, 1, 1)))
        self._pixels_changed()

    def rotate(self, angle):
        self._push()
        w, h = self.colors.width, self.colors.height
        self.colors = rotate_pixmap(self.colors, angle)
        if self.mask is not None:
            self.mask = rotate_pixmap(self.mask, angle)
        self.rotation = (self.rotation + angle) % 360
        if angle % 180:   # la larghezza nel PDF diventa l'altezza di prima: le misure vere non cambiano
            self.width_mm = self.width_mm * h / w
        self._after_change()

    # ------------------------------------------------------------ immagini nuove
    def _edit_source(self):
        """Riapre la firma con il suo testo o il suo disegno; quella rifatta resta alla stessa scala."""
        if self.edited:
            return
        dialog = SignatureDialog(self.top, source=self.source, name=self.name_var.get().strip() or "Firma")
        self.top.wait_window(dialog.top)
        self.top.grab_set()
        if not dialog.result:
            return
        pix, name, _width, source = dialog.result
        mm_per_px = self.width_mm / self.colors.width
        turned = (self.source_rotation + self.rotation) % 180
        self._push()
        self.original = rotate_pixmap(stamp_pixmap(pix), self.source_rotation)
        self.source = source
        self.colors, self.mask = split_alpha(rotate_pixmap(self.original, self.rotation))
        self.width_mm = min(max((pix.height if turned else pix.width) * mm_per_px, 3), 600)
        if self._initial[0] is not None:
            self.name_var.set(name)
        self._after_change()

    def _load_image(self, pix, keep_scale):
        """Mette un'immagine nuova al posto di quella attuale: keep_scale conserva i mm per pixel
        (es. ritagliata nell'editor esterno), altrimenti la larghezza nel PDF."""
        mm_per_px = self.width_mm / self.colors.width
        had_alpha = self.mask is not None
        self._push()
        self.colors, self.mask = split_alpha(stamp_pixmap(pix))
        if keep_scale:
            self.width_mm = min(max(self.colors.width * mm_per_px, 3), 600)
        if had_alpha and self.mask is None and self._initial[0] is not None and not self.white_var.get():
            self.white_var.set(True)   # l'editor ha perso la trasparenza: il bianco si toglie qui
        self._pixels_changed()

    def _replace(self):
        path = filedialog.askopenfilename(parent=self.top, title="Nuova immagine del timbro",
                                          initialdir=self.image_dir(), filetypes=IMAGE_TYPES)
        self.top.grab_set()
        if not path:
            return
        try:
            pix = pymupdf.Pixmap(path)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile aprire l'immagine:\n{exc}", parent=self.top)
            return
        StampEditor.last_image_dir = str(Path(path).parent)
        self._load_image(pix, keep_scale=False)

    last_image_dir = None   # cartella delle immagini: la imposta e la legge il programma (cartelle ricordate)

    @classmethod
    def image_dir(cls):
        return cls.last_image_dir or str(Path.home())

    def _open_external(self):
        """Apre l'immagine nell'editor di Windows (di solito Paint); quando la si salva lì, qui si aggiorna."""
        folder = Path(tempfile.gettempdir()) / APP_NAME
        folder.mkdir(exist_ok=True)
        name = re.sub(r'[\\/:*?"<>|]+', "_", self.name_var.get().strip()) or "timbro"
        path = folder / f"{name}.png"
        try:
            join_alpha(self.colors, self.mask).save(str(path))
            mtime = path.stat().st_mtime
            try:
                os.startfile(str(path), "edit")
            except OSError:
                os.startfile(str(path))
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Impossibile aprire l'editor di immagini:\n{exc}", parent=self.top)
            return
        self._ext = {"path": path, "seen": mtime, "loaded": mtime}
        self.note.configure(text=f"Immagine aperta nell'editor di immagini ({path.name}): modificala e salvala "
                                 "lì (Ctrl+S); qui si aggiorna da sola. Poi premi «Salva».")
        if not self._ext_job:
            self._watch_external()

    def _watch_external(self):
        ext = self._ext
        try:
            mtime = ext["path"].stat().st_mtime
        except OSError:
            mtime = ext["seen"]
        if mtime != ext["seen"]:          # appena cambiato: si aspetta che l'editor finisca di scriverlo
            ext["seen"] = mtime
        elif mtime != ext["loaded"]:
            try:
                pix = pymupdf.Pixmap(str(ext["path"]))
            except Exception:
                pix = None
            if pix is not None:
                ext["loaded"] = mtime
                self._load_image(pix, keep_scale=True)
        self._ext_job = self.top.after(700, self._watch_external)

    # ------------------------------------------------------------ chiusura
    def _on_destroy(self, event):
        if event.widget is self.top and self._ext_job:
            self.top.after_cancel(self._ext_job)
            self._ext_job = None

    def _close(self):
        name, white = self._initial
        if (self.changed or self.name_var.get().strip() != (name or "").strip() or self.white_var.get() != white) \
                and not messagebox.askyesno(APP_NAME, "Chiudere senza salvare le modifiche al timbro?",
                                            parent=self.top):
            return
        self.top.destroy()

    def _save(self):
        name, white = self._initial
        new_name = self.name_var.get().strip() or name
        if self.changed or new_name != name or self.white_var.get() != white:
            self.result = {"image": join_alpha(self.colors, self.mask),
                           "original": None if self.edited else self.original, "rotation": self.rotation,
                           "width_mm": round(self.width_mm, 1), "source": None if self.edited else self.source,
                           "remove_white": self.white_var.get(), "name": new_name}
        self.top.destroy()


# ---------------------------------------------------------------- timbri già apposti
# Ogni timbro è uno stream a sé ("q <matrice> cm /<nome> Do Q") e la sua immagine porta la
# chiave STAMP_KEY: così i timbri si riconoscono anche dopo il salvataggio e la riapertura.
STAMP_KEY = "TimbroPDF"   # segni interni nei PDF: restano uguali anche col nome nuovo del programma
# nell'immagine di una firma creata nel programma: il suo testo (o disegno) e la rotazione già applicata
SOURCE_KEY = "TimbroPDFFirma"
STAMP_STREAM = re.compile(rb"^\s*q\s+((?:\S+\s+){6})cm\s+/(\S+)\s+Do\s+Q\s*$")
PlacedStamp = namedtuple("PlacedStamp", "index cm name rect")  # rect: pagina come la si vede


def placed_stamps(page):
    """Timbri apposti da TimbroPDF sulla pagina, nell'ordine in cui sono disegnati."""
    doc = page.parent
    names = {img[7] for img in page.get_images(full=True) if doc.xref_get_key(img[0], STAMP_KEY)[1] == "true"}
    if not names:
        return []
    to_view = page.transformation_matrix * page.rotation_matrix
    found = []
    for index, xref in enumerate(page.get_contents()):
        match = STAMP_STREAM.match(doc.xref_stream(xref) or b"")
        if match and match.group(2).decode() in names:
            cm = pymupdf.Matrix(*map(float, match.group(1).split()))
            found.append(PlacedStamp(index, cm, match.group(2).decode(), pymupdf.Rect(0, 0, 1, 1) * cm * to_view))
    return found


def set_contents(page, xrefs):
    page.parent.xref_set_key(page.xref, "Contents", "[%s]" % " ".join(f"{x} 0 R" for x in xrefs))


def remove_placed_stamp(page, stamp):
    xrefs = page.get_contents()
    del xrefs[stamp.index]
    set_contents(page, xrefs)


def _restamp(page, stamp, cm):
    """Ridisegna il timbro con un'altra matrice.

    Si crea uno stream nuovo invece di modificare quello esistente: dopo il salvataggio
    timbri identici su pagine diverse possono condividere lo stesso stream.
    """
    doc = page.parent
    xref = doc.get_new_xref()
    doc.update_object(xref, "<<>>")
    numbers = " ".join(f"{v:.4f}" for v in (cm.a, cm.b, cm.c, cm.d, cm.e, cm.f))
    doc.update_stream(xref, f"\nq\n{numbers} cm\n/{stamp.name} Do\nQ\n".encode())
    xrefs = page.get_contents()
    xrefs[stamp.index] = xref
    set_contents(page, xrefs)


def move_placed_stamp(page, stamp, dx, dy):
    """Sposta il timbro di (dx, dy) punti nella pagina come la si vede."""
    to_pdf = ~(page.transformation_matrix * page.rotation_matrix)
    delta = pymupdf.Point(dx, dy) * to_pdf - pymupdf.Point(0, 0) * to_pdf
    cm = pymupdf.Matrix(stamp.cm)
    cm.e += delta.x
    cm.f += delta.y
    _restamp(page, stamp, cm)


def rotate_placed_stamp(page, stamp, angle):
    """Ruota il timbro attorno al suo centro, di angle gradi in senso orario come lo si vede (multipli di 90)."""
    c = pymupdf.Point(0.5, 0.5) * stamp.cm
    turn = pymupdf.Matrix(-angle)   # nello spazio del PDF l'asse y va in su: l'orario è negativo
    turn = pymupdf.Matrix(*(round(v) for v in (turn.a, turn.b, turn.c, turn.d)), 0, 0)
    to_center, back = pymupdf.Matrix(1, 0, 0, 1, -c.x, -c.y), pymupdf.Matrix(1, 0, 0, 1, c.x, c.y)
    _restamp(page, stamp, stamp.cm * to_center * turn * back)


def stamp_view_angle(page, stamp):
    """Di quanti gradi, in senso orario, l'immagine del timbro appare ruotata nella pagina vista."""
    m = stamp.cm * page.transformation_matrix * page.rotation_matrix
    return round(math.degrees(math.atan2(m.b, m.a)) / 90) * 90 % 360


def placed_stamp_image(page, stamp):
    """(immagine del timbro apposto, con la sua trasparenza; xref dell'immagine)."""
    doc = page.parent
    xref, smask = next((img[0], img[1]) for img in page.get_images(full=True) if img[7] == stamp.name)
    pix = pymupdf.Pixmap(doc, xref)
    if pix.colorspace is None or pix.colorspace.n != 3:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    if smask and not pix.alpha:
        pix = pymupdf.Pixmap(pix, pymupdf.Pixmap(doc, smask))
    return pix, xref


def mark_stamp_image(doc, xref, source=None, rotate=0):
    """Segna l'immagine come timbro; per le firme create nel programma ricorda come rifarle."""
    doc.xref_set_key(xref, STAMP_KEY, "true")
    if source:
        doc.xref_set_key(xref, SOURCE_KEY, pymupdf.get_pdf_str(json.dumps({"source": source, "rotate": rotate},
                                                                           ensure_ascii=False)))


def stamp_source(doc, xref):
    """(source, rotazione) della firma salvati nell'immagine di un timbro apposto, oppure None."""
    kind, value = doc.xref_get_key(xref, SOURCE_KEY)
    if kind != "string":
        return None
    try:
        data = json.loads(value)
        return data["source"], int(data.get("rotate", 0))
    except (ValueError, KeyError, TypeError):
        return None


def replace_placed_stamp(page, stamp, pix, width_mm, source=None, rotate=0):
    """Mette un'immagine nuova al posto del timbro apposto: stesso centro e stesso posto nell'ordine di
    disegno (primo o secondo piano); l'immagine appare diritta, come nella finestra di modifica."""
    doc = page.parent
    r = stamp.rect
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    w = width_mm * MM
    h = w * pix.height / pix.width
    rect = pymupdf.Rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2) * page.derotation_matrix
    rect.normalize()
    before = page.get_contents()
    old = before[stamp.index]
    xref = page.insert_image(rect, stream=pix.tobytes("png"), rotate=page.rotation)
    mark_stamp_image(doc, xref, source, rotate)
    name = next(img[7] for img in page.get_images(full=True) if img[0] == xref)
    new = next(x for x in page.get_contents() if x not in before
               and (m := STAMP_STREAM.match(doc.xref_stream(x) or b"")) and m.group(2).decode() == name)
    xrefs = [x for x in page.get_contents() if x != new]
    xrefs[xrefs.index(old)] = new
    set_contents(page, xrefs)


def render_with_contents(page, xrefs, **kwargs):
    """Disegna la pagina usando solo gli stream indicati, poi ripristina il contenuto originale."""
    original = page.get_contents()
    set_contents(page, xrefs)
    try:
        return page.get_pixmap(**kwargs)
    finally:
        set_contents(page, original)


# ---------------------------------------------------------------- testo
# Anche le caselle di testo scritte da TimbroPDF sono uno stream a sé: i parametri (testo, font,
# posizione…) stanno nella chiave TEXT_KEY dello stream, così si spostano e si correggono anche
# dopo il salvataggio. Posizioni e riquadri sono sulla pagina non ruotata, come in get_text().
TEXT_KEY = "TimbroPDFTesto"
TEXT_COLORS = {"Nero": "#000000", "Blu": "#1f3a93", "Rosso": "#b3261e"}
LINE_HEIGHT = 1.2      # distanza tra le righe, in multipli della dimensione del font
TEXT_SHIFT = re.compile(rb"^q 1 0 0 1 (\S+) (\S+) cm\n(.*)\nQ\n$", re.S)
FONT_ALIASES = {"helvetica": "arial", "arialmt": "arial", "times": "timesnewroman", "timesroman": "timesnewroman",
                "courier": "couriernew"}
PlacedText = namedtuple("PlacedText", "index params rect")   # rect: pagina come la si vede
TextLine = namedtuple("TextLine", "rect text font size color flags origin rotate")


def text_font(fontfile):
    return pymupdf.Font(fontfile=fontfile) if fontfile else pymupdf.Font("helv")


def font_resource(name, token):
    """Nome del font nelle risorse della pagina.

    PyMuPDF riusa un font già presente con lo stesso nome, e dopo il salvataggio i font sono
    ridotti ai soli caratteri usati: per questo il nome cambia a ogni apertura del documento.
    """
    return "TP" + re.sub(r"[^A-Za-z0-9]", "", name)[:20] + token


def match_font(pdf_name, flags, fonts):
    """Font installato più simile a quello di una riga del PDF (None se non c'è niente di adatto)."""
    def norm(s):
        return re.sub(r"[^a-z0-9]", "", s.lower())

    family, _, style = re.sub(r"^[A-Z]{6}\+", "", pdf_name).replace(",", "-").partition("-")
    family = norm(re.sub(r"(PSMT|PS|MT)$", "", family))
    family = FONT_ALIASES.get(family, family)
    style = style.lower()
    bold = bool(flags & 16) or any(w in style for w in ("bold", "black", "heavy", "semibold", "demi"))
    italic = bool(flags & 2) or any(w in style for w in ("italic", "oblique"))
    suffix = ("bold" if bold else "") + ("italic" if italic else "")
    generic = "timesnewroman" if flags & 4 else "couriernew" if flags & 8 else "arial"
    index = {norm(name): name for name in fonts}
    for fam in (family, generic):
        for candidate in (fam + suffix, fam + "regular", fam):
            if candidate in index:
                return index[candidate]
    return None


def text_box_rect(params, font):
    """Riquadro della casella di testo sulla pagina non ruotata."""
    lines = params["text"].split("\n")
    size = params["size"]
    width = max(font.text_length(line, fontsize=size) for line in lines)
    local = pymupdf.Rect(0, -font.ascender * size, width,
                         (len(lines) - 1) * LINE_HEIGHT * size - font.descender * size)
    # il testo scorre lungo u e va a capo lungo v (l'asse y della pagina va verso il basso)
    a = math.radians(params["rotate"])
    u = round(math.cos(a), 6), round(-math.sin(a), 6)
    v = round(math.sin(a), 6), round(math.cos(a), 6)
    return local * pymupdf.Matrix(u[0], u[1], v[0], v[1], *params["origin"])


def fix_space_mapping(page):
    """PyMuPDF associa lo spazio dei font TrueType al carattere U+00A0: copiando il testo,
    cercandolo o convertendolo in Word gli spazi diventerebbero spazi non separabili."""
    doc = page.parent
    for xref, _ext, _type, _base, name, *_ in page.get_fonts(full=True):
        if not name.startswith("TP"):
            continue
        kind, value = doc.xref_get_key(xref, "ToUnicode")
        if kind != "xref":
            continue
        cmap = int(value.split()[0])
        data = doc.xref_stream(cmap)
        fixed = re.sub(rb"beginbfchar.*?endbfchar",
                       lambda m: re.sub(rb"(<[0-9a-fA-F]+>\s*)<00[aA]0>", rb"\1<0020>", m.group(0)),
                       data, flags=re.S)
        if fixed != data:
            doc.update_stream(cmap, fixed)


def insert_text_box(page, params, fontfile, resname):
    """Scrive una casella di testo: uno stream nuovo, marcato con i suoi parametri. Restituisce lo xref."""
    doc = page.parent
    params = dict(params, rect=[round(v, 2) for v in text_box_rect(params, text_font(fontfile))])
    before = set(page.get_contents())
    page.insert_text(params["origin"], params["text"], fontsize=params["size"], lineheight=LINE_HEIGHT,
                     fontname=resname if fontfile else "helv", fontfile=fontfile,
                     color=hex_to_rgb(params["color"]), rotate=params["rotate"])
    xref = [x for x in page.get_contents() if x not in before and b"BT" in doc.xref_stream(x)][-1]
    doc.xref_set_key(xref, TEXT_KEY, pymupdf.get_pdf_str(json.dumps(params, ensure_ascii=False)))
    if fontfile:
        fix_space_mapping(page)
    return xref


def placed_texts(page):
    """Caselle di testo scritte da TimbroPDF sulla pagina."""
    doc = page.parent
    found = []
    for index, xref in enumerate(page.get_contents()):
        kind, value = doc.xref_get_key(xref, TEXT_KEY)
        if kind != "string":
            continue
        try:
            params = json.loads(value)
            rect = pymupdf.Rect(params["rect"]) * page.rotation_matrix
        except (ValueError, KeyError, TypeError):
            continue
        found.append(PlacedText(index, params, rect))
    return found


def placed_objects(page):
    """Timbri e caselle di testo di TimbroPDF, nell'ordine in cui sono disegnati."""
    return sorted(placed_stamps(page) + placed_texts(page), key=lambda o: o.index)


def move_placed_text(page, box, dx, dy):
    """Sposta la casella di (dx, dy) punti nella pagina vista.

    Lo stream viene solo racchiuso in una traslazione: il testo non si riscrive, così resta
    identico anche se il font non è installato su questo PC.
    """
    doc = page.parent
    to_pdf = ~(page.transformation_matrix * page.rotation_matrix)
    d = pymupdf.Point(dx, dy) * to_pdf - pymupdf.Point(0, 0) * to_pdf
    du = pymupdf.Point(dx, dy) * page.derotation_matrix - pymupdf.Point(0, 0) * page.derotation_matrix
    xrefs = page.get_contents()
    data = doc.xref_stream(xrefs[box.index])
    shifted = TEXT_SHIFT.match(data)
    if shifted:
        tx, ty, data = float(shifted.group(1)) + d.x, float(shifted.group(2)) + d.y, shifted.group(3)
    else:
        tx, ty = d.x, d.y
    params = dict(box.params)
    params["origin"] = [params["origin"][0] + du.x, params["origin"][1] + du.y]
    params["rect"] = [round(v, 2) for v in pymupdf.Rect(params["rect"]) + (du.x, du.y, du.x, du.y)]
    xref = doc.get_new_xref()
    doc.update_object(xref, "<<>>")
    doc.update_stream(xref, b"q 1 0 0 1 %.4f %.4f cm\n" % (tx, ty) + data + b"\nQ\n")
    doc.xref_set_key(xref, TEXT_KEY, pymupdf.get_pdf_str(json.dumps(params, ensure_ascii=False)))
    xrefs[box.index] = xref
    set_contents(page, xrefs)


def replace_placed_text(page, box, params, fontfile, resname):
    """Riscrive la casella con i nuovi parametri, nella stessa posizione dell'ordine di disegno."""
    xrefs = page.get_contents()
    previous = xrefs[box.index - 1] if box.index else None
    del xrefs[box.index]
    set_contents(page, xrefs)
    new = insert_text_box(page, params, fontfile, resname)
    xrefs = [x for x in page.get_contents() if x != new]
    xrefs.insert(xrefs.index(previous) + 1 if previous in xrefs else 0, new)
    set_contents(page, xrefs)


def place_with_content(page, xref):
    """Sposta lo stream subito dopo il contenuto proprio della pagina, sotto timbri e caselle in primo piano."""
    contents = page.get_contents()
    ours = {contents[o.index] for o in placed_objects(page)} - {xref}
    xrefs = [x for x in contents if x != xref]
    last = max((i for i, x in enumerate(xrefs) if x not in ours), default=-1)
    xrefs.insert(last + 1, xref)
    set_contents(page, xrefs)


def text_lines(page):
    """Righe di testo della pagina, escluse le caselle di TimbroPDF, con lo stile del pezzo più lungo."""
    boxes = [pymupdf.Rect(t.params["rect"]) for t in placed_texts(page)]
    lines = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = [s for s in line["spans"] if s["text"].strip()]
            if not spans:
                continue
            rect = pymupdf.Rect(line["bbox"])
            center = pymupdf.Point((rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2)
            if any(b.contains(center) for b in boxes):
                continue
            main = max(spans, key=lambda s: len(s["text"].strip()))
            dx, dy = line["dir"]
            rotate = round(math.degrees(math.atan2(-dy, dx)) / 90) * 90 % 360
            text = "".join(s["text"] for s in line["spans"]).replace("\xa0", " ").strip()
            lines.append(TextLine(rect, text, main["font"],
                                  main["size"], f"#{main['color'] & 0xFFFFFF:06x}", main["flags"],
                                  tuple(spans[0]["origin"]), rotate))
    return lines


def text_band(line):
    """Striscia centrale della riga: basta per eliminarne le lettere senza toccare le righe vicine."""
    r = line.rect
    if line.rotate in (0, 180):
        c, h = (r.y0 + r.y1) / 2, r.height * 0.15
        return pymupdf.Rect(r.x0, c - h, r.x1, c + h)
    c, w = (r.x0 + r.x1) / 2, r.width * 0.15
    return pymupdf.Rect(c - w, r.y0, c + w, r.y1)


def erase(page, areas, fill=False, images=pymupdf.PDF_REDACT_IMAGE_NONE,
          graphics=pymupdf.PDF_REDACT_LINE_ART_NONE):
    """Elimina davvero il testo nelle aree (pagina non ruotata); fill colora le aree, False no.

    apply_redactions riscrive tutto il contenuto della pagina in un solo stream: timbri e
    caselle di testo di TimbroPDF si staccano prima e si riattaccano dopo, con le loro risorse,
    così restano spostabili e non vengono toccati.
    """
    doc = page.parent
    contents = page.get_contents()
    ours = {o.index for o in placed_objects(page)}
    saved = {}
    for i in ours:
        # "/fzImg0 Do", "/TPArial12ab 12 Tf", "/gs0 gs": le risorse usate dallo stream
        for name, op in re.findall(rb"/([^\s/\[\]()<>{}%]+)\s+(?:[-+.\d]+\s+)?(Do|Tf|gs)\b",
                                   doc.xref_stream(contents[i]) or b""):
            key = "Resources/%s/%s" % ({b"Do": "XObject", b"Tf": "Font", b"gs": "ExtGState"}[op], name.decode())
            kind, value = doc.xref_get_key(page.xref, key)
            if kind != "null":
                saved[key] = value
    if ours:
        set_contents(page, [x for i, x in enumerate(contents) if i not in ours])
    for area in areas:
        page.add_redact_annot(area, fill=fill)
    page.apply_redactions(images=images, graphics=graphics, text=pymupdf.PDF_REDACT_TEXT_REMOVE)
    if ours:
        first = min((i for i in range(len(contents)) if i not in ours), default=len(contents))
        set_contents(page, [contents[i] for i in sorted(ours) if i < first] + page.get_contents()
                     + [contents[i] for i in sorted(ours) if i > first])
        for key, value in saved.items():
            if doc.xref_get_key(page.xref, key)[0] == "null":
                doc.xref_set_key(page.xref, key, value)


def image_under(page, rect):
    """True se sotto il riquadro c'è un'immagine (una scansione, una foto), esclusi i timbri."""
    doc = page.parent
    center = pymupdf.Point((rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2)
    for info in page.get_image_info(xrefs=True):
        if info["xref"] and doc.xref_get_key(info["xref"], STAMP_KEY)[1] == "true":
            continue
        if pymupdf.Rect(info["bbox"]).contains(center):
            return True
    return False


# ---------------------------------------------------------------- OCR e Word
def ocr_page(page, language):
    """Aggiunge sopra la pagina il testo riconosciuto, invisibile: la pagina resta identica ma il
    testo si può cercare, copiare ed esportare. Restituisce il numero di parole trovate."""
    r = page.rect
    dpi = min(OCR_DPI, int(72 * (OCR_MAX_PX / (r.width * r.height)) ** 0.5))
    pix = page.get_pixmap(dpi=dpi)
    ocr = pymupdf.open("pdf", pix.pdfocr_tobytes(language=language, tessdata=str(TESSDATA_DIR)))
    layer = ocr[0]
    words = len(layer.get_text("words"))
    if not words:
        return 0
    layer.add_redact_annot(layer.rect)     # dell'OCR serve solo il testo: via l'immagine
    layer.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_REMOVE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                           text=pymupdf.PDF_REDACT_TEXT_NONE)
    # lo strato è dritto come la pagina vista: va ruotato come la pagina
    page.show_pdf_page(page.rect * page.derotation_matrix, ocr, 0, rotate=page.rotation)
    return words


# Word va pilotato con parametri "per riferimento"; si chiude solo se non ha altri documenti aperti
WORD_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $file = $env:TIMBRO_PDF; $confirm = $false; $readonly = $true; $recent = $false
    $doc = $word.Documents.Open([ref]$file, [ref]$confirm, [ref]$readonly, [ref]$recent)
    $out = $env:TIMBRO_DOCX; $format = 16
    $doc.SaveAs2([ref]$out, [ref]$format)
    $nosave = 0
    $doc.Close([ref]$nosave)
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
"""


def word_installed():
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                       r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"))
        return True
    except (ImportError, OSError):
        return False


def pdf_to_docx(pdf, docx, timeout=900):
    """Converte il PDF con Microsoft Word. Restituisce None se va bene, altrimenti il messaggio d'errore.

    Word non accetta percorsi più lunghi di 255 caratteri: usare file in una cartella temporanea.
    """
    exe = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    env = dict(os.environ, TIMBRO_PDF=str(pdf), TIMBRO_DOCX=str(docx))
    try:
        res = subprocess.run([str(exe), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                              "-Command", WORD_SCRIPT], env=env, capture_output=True, text=True,
                             encoding="oem", errors="replace", timeout=timeout,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        return "Word non ha risposto in tempo."
    except OSError as exc:
        return str(exc)
    if res.returncode != 0 or not Path(docx).exists():
        return (res.stderr or res.stdout or "errore sconosciuto").strip()
    return None


_TK_FAMILIES = None


def tk_font_style(name, widget):
    """(famiglia, peso, stile) di Tk per un font installato: «Arial Bold Italic» -> Arial, bold, italic."""
    global _TK_FAMILIES
    if _TK_FAMILIES is None:
        _TK_FAMILIES = set(tkfont.families(widget))
    for suffix, weight, slant in (("", "normal", "roman"), (" Bold Italic", "bold", "italic"),
                                  (" Bold Oblique", "bold", "italic"), (" Bold", "bold", "roman"),
                                  (" Italic", "normal", "italic"), (" Oblique", "normal", "italic")):
        base = name[:len(name) - len(suffix)] if suffix else name
        if name.endswith(suffix) and base in _TK_FAMILIES:
            return base, weight, slant
    return "Segoe UI", "normal", "roman"


def text_preview(page, params, fontfile, zoom):
    """La casella di testo disegnata da sola, su fondo trasparente, come apparirà sulla pagina:
    (pixmap, riquadro nella pagina vista)."""
    box = text_box_rect(params, text_font(fontfile)) * page.rotation_matrix
    doc = pymupdf.open()
    tmp = doc.new_page(width=page.cropbox.width, height=page.cropbox.height)
    tmp.set_rotation(page.rotation)
    tmp.insert_text(params["origin"], params["text"], fontsize=params["size"], lineheight=LINE_HEIGHT,
                    fontname="F0" if fontfile else "helv", fontfile=fontfile, color=hex_to_rgb(params["color"]),
                    rotate=params["rotate"])
    clip = (box + (-2, -2, 2, 2)) & tmp.rect
    return tmp.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip, alpha=True), clip


class TextDialog:
    """Testo da scrivere sul PDF: contenuto, font, dimensione e colore.

    whiteout None nasconde la casella «Copri con il bianco»; True/False la mostra già impostata.
    preview(dati) riceve a ogni modifica testo, font, dimensione e colore (None quando la finestra si
    chiude o i dati non sono validi), per mostrare il risultato mentre si scrive.
    """

    def __init__(self, parent, fonts, title, params, original=None, whiteout=None, preview=None):
        self.result = None   # dict(text, font, size, color, whiteout)
        self.color = params["color"]
        self.preview = preview
        self._job = None
        s = parent.winfo_fpixels("1i") / 96
        top = self.top = tk.Toplevel(parent)
        top.title(title)
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=12)
        body.pack(fill=tk.BOTH, expand=True)
        if original:
            ttk.Label(body, text=original, foreground="#555").pack(anchor="w", pady=(0, 6))
        # riquadro di misura fissa: il testo si vede con il suo font, la finestra non cambia misura
        box = tk.Frame(body, width=round(560 * s), height=round(130 * s))
        box.pack_propagate(False)
        box.pack(fill=tk.X)
        self.text = tk.Text(box, font=("Segoe UI", 11), wrap="none", undo=True, padx=6, pady=4)
        self.text.pack(fill=tk.BOTH, expand=True)
        self.text.insert("1.0", params["text"])
        self.text.tag_add("sel", "1.0", "end-1c")
        self.text.edit_modified(False)
        self.text.bind("<<Modified>>", self._on_modified)
        self.format_note = ttk.Label(body, foreground="#777")
        self.format_note.pack(anchor="e")

        row = ttk.Frame(body)
        row.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(row, text="Font:").pack(side=tk.LEFT)
        names = sorted(fonts, key=str.lower) or ["Helvetica"]
        font = params["font"] if params["font"] in names else "Arial" if "Arial" in names else names[0]
        self.font_var = tk.StringVar(value=font)
        ttk.Combobox(row, textvariable=self.font_var, values=names, state="readonly", width=32,
                     height=20).pack(side=tk.LEFT, padx=(6, 14))
        ttk.Label(row, text="Dimensione (pt):").pack(side=tk.LEFT)
        self.size_var = tk.StringVar(value=f"{params['size']:g}")
        ttk.Spinbox(row, from_=4, to=200, increment=0.5, width=6, textvariable=self.size_var).pack(side=tk.LEFT,
                                                                                                padx=(6, 0))
        row = ttk.Frame(body)
        row.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(row, text="Colore:").pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value=next((n for n, c in TEXT_COLORS.items() if c == self.color.lower()), ""))
        for name in TEXT_COLORS:
            ttk.Radiobutton(row, text=name, value=name, variable=self.color_var,
                            command=lambda: self._set_color(TEXT_COLORS[self.color_var.get()])).pack(side=tk.LEFT,
                                                                                                     padx=(6, 0))
        ttk.Button(row, text="Altro colore…", command=self._pick_color).pack(side=tk.LEFT, padx=(10, 6))
        self.swatch = tk.Label(row, width=3, bg=self.color, relief=tk.SOLID, bd=1)
        self.swatch.pack(side=tk.LEFT)
        self.whiteout_var = tk.BooleanVar(value=bool(whiteout))
        if whiteout is not None:
            ttk.Checkbutton(body, variable=self.whiteout_var,
                            text="Copri con il bianco lo sfondo della riga (scansioni, testo sopra immagini)"
                            ).pack(anchor="w", pady=(10, 0))
        ttk.Label(body, text="Invio va a capo · Ctrl+Invio conferma · Esc annulla",
                  foreground="#777").pack(anchor="w", pady=(8, 0))

        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(buttons, text="Annulla", command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="OK", command=self._ok).pack(side=tk.RIGHT, padx=6)
        top.bind("<Control-Return>", lambda e: self._ok() or "break")
        top.bind("<Escape>", lambda e: top.destroy())
        top.bind("<Destroy>", self._on_destroy)
        for var in (self.font_var, self.size_var):
            var.trace_add("write", self._changed)
        self._refresh()
        place_dialog(top, parent)
        self.text.focus_set()

    def _pick_color(self):
        color = colorchooser.askcolor(color=self.color, parent=self.top, title="Colore del testo")[1]
        if color:
            self.color_var.set("")
            self._set_color(color)

    def _set_color(self, color):
        self.color = color
        self.swatch.configure(bg=color)
        self._changed()

    def _on_modified(self, _event):
        if self.text.edit_modified():
            self.text.edit_modified(False)
            self._changed()

    def _changed(self, *_):
        if self._job:
            self.top.after_cancel(self._job)
        self._job = self.top.after(120, self._refresh)

    def _values(self):
        """(testo, dimensione): la dimensione è None se non è un numero valido."""
        text = self.text.get("1.0", "end-1c").replace("\t", "    ").rstrip()
        try:
            size = float(self.size_var.get().replace(",", "."))
        except ValueError:
            size = None
        return text, size if size and 4 <= size <= 400 else None

    def _refresh(self):
        """Il testo nel riquadro prende font, colore e (fin dove ci sta) dimensione scelti; poi l'anteprima."""
        self._job = None
        text, size = self._values()
        family, weight, slant = tk_font_style(self.font_var.get(), self.top)
        shown = round(min(max(size or 11, 8), 36))
        light = sum(hex_to_rgb(self.color)) > 2.55   # testo chiaro: su fondo scuro, se no non si vede
        self.text.configure(font=(family, shown, weight, slant), foreground=self.color,
                            insertbackground=self.color, background="#5f6368" if light else "white")
        self.format_note.configure(text="" if not size or shown == round(size) else
                                   f"Nel riquadro il testo è mostrato a {shown} pt; nel PDF sarà di {size:g} pt.")
        if self.preview:
            self.preview({"text": text, "font": self.font_var.get(), "size": size, "color": self.color}
                         if text.strip() and size else None)

    def _on_destroy(self, event):
        if event.widget is self.top:
            if self._job:
                self.top.after_cancel(self._job)
            if self.preview:
                self.preview(None)

    def _ok(self):
        text = self.text.get("1.0", "end-1c").replace("\t", "    ").rstrip()
        try:
            size = float(self.size_var.get().replace(",", "."))
        except ValueError:
            size = 0
        if not text.strip():
            messagebox.showinfo(APP_NAME, "Scrivi il testo.", parent=self.top)
            return
        if not 4 <= size <= 400:
            messagebox.showinfo(APP_NAME, "Dimensione non valida (da 4 a 400 punti).", parent=self.top)
            return
        self.result = {"text": text, "font": self.font_var.get(), "size": size, "color": self.color,
                       "whiteout": self.whiteout_var.get()}
        self.top.destroy()


class OcrDialog:
    """Scelta delle pagine e della lingua per l'OCR."""

    def __init__(self, parent, current, selected, total):
        self.result = None   # (pagine, lingua, salta le pagine con testo)
        self.choices = {"current": [current], "selected": selected, "all": list(range(total))}
        top = self.top = tk.Toplevel(parent)
        top.title("Riconosci testo (OCR)")
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=14)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, wraplength=420, justify=tk.LEFT, text=(
            "Il testo riconosciuto viene aggiunto invisibile sopra la scansione: la pagina resta identica, "
            "ma il testo si può cercare, selezionare, copiare ed esportare in Word.")).pack(anchor="w")
        ttk.Label(body, text="Pagine:").pack(anchor="w", pady=(12, 0))
        self.scope = tk.StringVar(value="all")
        options = [("current", f"Pagina corrente ({current + 1})")]
        if len(selected) > 1:
            options.append(("selected", f"Pagine selezionate ({len(selected)})"))
        options.append(("all", f"Tutte le pagine ({total})"))
        for value, text in options:
            ttk.Radiobutton(body, text=text, value=value, variable=self.scope).pack(anchor="w", padx=(12, 0))
        row = ttk.Frame(body)
        row.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(row, text="Lingua:").pack(side=tk.LEFT)
        self.language = tk.StringVar(value=next(iter(OCR_LANGUAGES)))
        ttk.Combobox(row, textvariable=self.language, values=list(OCR_LANGUAGES), state="readonly",
                     width=20).pack(side=tk.LEFT, padx=(6, 0))
        self.skip = tk.BooleanVar(value=True)
        ttk.Checkbutton(body, variable=self.skip,
                        text="Salta le pagine che contengono già testo").pack(anchor="w", pady=(12, 0))
        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(buttons, text="Annulla", command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Avvia", command=self._ok).pack(side=tk.RIGHT, padx=6)
        top.bind("<Return>", lambda e: self._ok())
        top.bind("<Escape>", lambda e: top.destroy())
        place_dialog(top, parent)

    def _ok(self):
        self.result = (self.choices[self.scope.get()], OCR_LANGUAGES[self.language.get()], self.skip.get())
        self.top.destroy()


class ProgressWindow:
    """Finestra di avanzamento modale; «Annulla» imposta cancelled (se can_cancel)."""

    def __init__(self, parent, title, maximum=None, can_cancel=True):
        self.cancelled = False
        top = self.top = tk.Toplevel(parent)
        top.title(title)
        top.transient(parent)
        top.resizable(False, False)
        top.protocol("WM_DELETE_WINDOW", self._cancel if can_cancel else lambda: None)
        body = ttk.Frame(top, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        self.label = ttk.Label(body, width=52)
        self.label.pack(anchor="w")
        self.bar = ttk.Progressbar(body, length=380, maximum=maximum or 100,
                                   mode="determinate" if maximum else "indeterminate")
        self.bar.pack(fill=tk.X, pady=(8, 0))
        if not maximum:
            self.bar.start(12)
        if can_cancel:
            ttk.Button(body, text="Annulla", command=self._cancel).pack(anchor="e", pady=(12, 0))
        place_dialog(top, parent)

    def update(self, value, text):
        self.bar["value"] = value
        self.label["text"] = text
        self.top.update()

    def _cancel(self):
        self.cancelled = True
        self.label["text"] = "Interruzione…"

    def close(self):
        self.top.grab_release()
        self.top.destroy()


# ---------------------------------------------------------------- disposizione della finestra
# barre: lato e ordine nel lato; pannelli: lato e ordine dal bordo (0 = il più esterno);
# larghezze in pixel (0 = la minima necessaria al contenuto)
DEFAULT_LAYOUT = {"bars": {"main": ["top", 0], "text": ["top", 1]},
                  "panels": {"thumbs": ["left", 0], "stamps": ["right", 0]},
                  "widths": {"thumbs": 180, "stamps": 0},      # pannelli a sinistra/destra
                  "heights": {"thumbs": 210, "stamps": 0},     # pannelli in alto/in basso
                  "gallery_h": 0}                              # altezza della galleria dei timbri
SIDE_NAMES = {"top": "in alto", "bottom": "in basso", "left": "a sinistra", "right": "a destra"}
SASH_BG = "#b9bbc4"            # colore dei divisori trascinabili fra i pannelli


def default_layout():
    return json.loads(json.dumps(DEFAULT_LAYOUT))


def bar_button(group, text, command, width=None):
    ttk.Button(group, text=text, command=command, width=width).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)


def on_screen(geometry):
    """True se la finestra "LxA+X+Y" cade almeno in parte su uno degli schermi collegati."""
    m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$", geometry or "")
    if not m:
        return False
    w, _h, x, y = map(int, m.groups())
    try:
        import ctypes
        metric = ctypes.windll.user32.GetSystemMetrics
        vx, vy, vw, vh = metric(76), metric(77), metric(78), metric(79)   # schermo virtuale (tutti i monitor)
    except Exception:
        return True
    return vx - w + 100 <= x <= vx + vw - 100 and vy - 10 <= y <= vy + vh - 100


class ToolBar(ttk.Frame):
    """Barra di pulsanti agganciabile a uno dei quattro lati della finestra.

    I pulsanti stanno in gruppi che restano orizzontali: con la barra in verticale i gruppi
    si impilano. La barra si sposta trascinando la maniglia all'inizio.
    """

    def __init__(self, master, key, on_grip):
        super().__init__(master, padding=(4, 3))
        self.key = key
        self.side = None
        self.items = []            # gruppi (Frame) e separatori (None)
        self._separators = []
        bg = ttk.Style().lookup("TFrame", "background") or master.cget("bg")
        self.grip = tk.Label(self, text="⋮", cursor="fleur", fg="#999", bg=bg, font=("Segoe UI", 12, "bold"), padx=3)
        for seq, phase in (("<Button-1>", "press"), ("<B1-Motion>", "motion"), ("<ButtonRelease-1>", "release")):
            self.grip.bind(seq, lambda e, p=phase: on_grip(self, e, p))

    def group(self):
        g = ttk.Frame(self)
        self.items.append(g)
        return g

    def separator(self):
        self.items.append(None)

    def orient(self, side):
        """Dispone i gruppi in riga (in alto/in basso) o in colonna (a sinistra/a destra)."""
        if side == self.side:
            return
        self.side = side
        vertical = side in ("left", "right")
        for s in self._separators:
            s.destroy()
        self._separators = []
        self.grip.pack_forget()
        for g in self.items:
            if g is not None:
                g.pack_forget()
        self.grip.configure(text="⋯" if vertical else "⋮")
        self.grip.pack(side=tk.TOP if vertical else tk.LEFT, fill=tk.X if vertical else tk.Y)
        for g in self.items:
            if g is None:
                s = ttk.Separator(self, orient=tk.HORIZONTAL if vertical else tk.VERTICAL)
                if vertical:
                    s.pack(side=tk.TOP, fill=tk.X, pady=5)
                else:
                    s.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
                self._separators.append(s)
            elif vertical:
                g.pack(side=tk.TOP, fill=tk.X, pady=1)
            else:
                g.pack(side=tk.LEFT, padx=1)


class PasswordDialog:
    """Password dell'archivio dei timbri; con confirm la si sceglie (scritta due volte)."""

    def __init__(self, parent, message, confirm=False):
        self.result = None   # (password, ricordala su questo PC)
        self.confirm = confirm
        top = self.top = tk.Toplevel(parent)
        top.title("Archivio dei timbri")
        if parent.winfo_viewable():
            top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=14)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text=message, wraplength=420, justify=tk.LEFT).pack(anchor="w")
        self.pw = tk.StringVar()
        self.pw2 = tk.StringVar()
        grid = ttk.Frame(body)
        grid.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(grid, text="Password:").grid(row=0, column=0, sticky="w", pady=2)
        entry = ttk.Entry(grid, textvariable=self.pw, show="•", width=30)
        entry.grid(row=0, column=1, padx=(8, 0))
        if confirm:
            ttk.Label(grid, text="Ripetila:").grid(row=1, column=0, sticky="w", pady=2)
            ttk.Entry(grid, textvariable=self.pw2, show="•", width=30).grid(row=1, column=1, padx=(8, 0))
        self.remember = tk.BooleanVar(value=True)
        ttk.Checkbutton(body, variable=self.remember,
                        text="Ricordala su questo PC (resta cifrata da Windows, solo per il tuo utente)"
                        ).pack(anchor="w", pady=(10, 0))
        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(buttons, text="Annulla", command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="OK", command=self._ok).pack(side=tk.RIGHT, padx=6)
        top.bind("<Return>", lambda e: self._ok())
        top.bind("<Escape>", lambda e: top.destroy())
        place_dialog(top, parent)
        entry.focus_set()

    def _ok(self):
        pw = self.pw.get()
        if not pw:
            messagebox.showinfo(APP_NAME, "Scrivi la password.", parent=self.top)
            return
        if self.confirm and len(pw) < 6:
            messagebox.showinfo(APP_NAME, "Usa almeno 6 caratteri.", parent=self.top)
            return
        if self.confirm and pw != self.pw2.get():
            messagebox.showinfo(APP_NAME, "Le due password non coincidono.", parent=self.top)
            return
        self.result = (pw, self.remember.get())
        self.top.destroy()


class ArchiveDialog:
    """Dove sta l'archivio dei timbri e com'è protetto; i pulsanti scelgono cosa fare (action)."""

    def __init__(self, parent, path, count, protected, available):
        self.action = None
        top = self.top = tk.Toplevel(parent)
        top.title("Archivio dei timbri")
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=14)
        body.pack(fill=tk.BOTH, expand=True)
        state = ("Protetto da password (cifrato AES-256)" if protected else "Non protetto da password")
        if not available:
            state = "Non aperto (manca la password o il file non è raggiungibile)"
        ttk.Label(body, text=f"File:  {path}", wraplength=520, justify=tk.LEFT).pack(anchor="w")
        ttk.Label(body, text=f"Timbri:  {count}   ·   {state}").pack(anchor="w", pady=(4, 0))
        ttk.Label(body, wraplength=520, justify=tk.LEFT, foreground="#555", text=(
            "Per avere gli stessi timbri su più PC sposta l'archivio in una cartella condivisa (es. Dropbox); "
            "sugli altri PC scegli «Usa un archivio esistente» e indica lo stesso file. Con la password, "
            "chi copia il file senza conoscerla non può vedere né usare i timbri.")).pack(anchor="w", pady=(10, 4))
        actions = [("open_folder", "Apri la cartella dell'archivio"),
                   ("move", "Sposta l'archivio in un'altra cartella…"),
                   ("use", "Usa un archivio esistente (es. quello di un altro PC)…")]
        if not available:
            actions.insert(0, ("unlock", "Apri l'archivio…"))
        elif protected:
            actions += [("password", "Cambia la password…"), ("unprotect", "Togli la password")]
        else:
            actions.append(("password", "Proteggi con una password…"))
        for action, text in actions:
            ttk.Button(body, text=text, command=lambda a=action: self._choose(a)).pack(fill=tk.X, pady=2)
        ttk.Button(body, text="Chiudi", command=top.destroy).pack(anchor="e", pady=(10, 0))
        top.bind("<Escape>", lambda e: top.destroy())
        place_dialog(top, parent)

    def _choose(self, action):
        self.action = action
        self.top.destroy()


class MergeDialog:
    """Scelta dei PDF aperti da unire (spunta) e del loro ordine (frecce)."""

    def __init__(self, parent, tabs):
        self.result = None   # schede scelte, nell'ordine
        self.rows = [(tab, tk.BooleanVar(value=True)) for tab in tabs]
        top = self.top = tk.Toplevel(parent)
        top.title("Unisci documenti")
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=14)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="Spunta i documenti da unire; con le frecce scegli l'ordine.\n"
                             "Il risultato si apre in una nuova scheda.").pack(anchor="w")
        self.list_frame = ttk.Frame(body)
        self.list_frame.pack(fill=tk.X, pady=(10, 0))
        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(buttons, text="Tutti", command=lambda: self._check_all(True)).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Nessuno", command=lambda: self._check_all(False)).pack(side=tk.LEFT, padx=4)
        ttk.Button(buttons, text="Annulla", command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Unisci", command=self._ok).pack(side=tk.RIGHT, padx=6)
        top.bind("<Escape>", lambda e: top.destroy())
        self._draw()
        place_dialog(top, parent)

    def _draw(self):
        for w in self.list_frame.winfo_children():
            w.destroy()
        for k, (tab, var) in enumerate(self.rows):
            row = ttk.Frame(self.list_frame)
            row.pack(fill=tk.X, pady=1)
            ttk.Checkbutton(row, variable=var, text=f"{k + 1}.  {tab.name}   ({len(tab.doc)} pag.)").pack(side=tk.LEFT)
            down = ttk.Button(row, text="▼", width=3, command=lambda k=k: self._move(k, 1))
            up = ttk.Button(row, text="▲", width=3, command=lambda k=k: self._move(k, -1))
            down.pack(side=tk.RIGHT)
            up.pack(side=tk.RIGHT, padx=(16, 2))
            if k == 0:
                up.state(["disabled"])
            if k == len(self.rows) - 1:
                down.state(["disabled"])

    def _move(self, k, step):
        rows = self.rows
        rows[k], rows[k + step] = rows[k + step], rows[k]
        self._draw()

    def _check_all(self, value):
        for _tab, var in self.rows:
            var.set(value)

    def _ok(self):
        chosen = [tab for tab, var in self.rows if var.get()]
        if len(chosen) < 2:
            messagebox.showinfo(APP_NAME, "Spunta almeno due documenti.", parent=self.top)
            return
        self.result = chosen
        self.top.destroy()


AUTHOR = "ing. Alessandro Candido"
DONATION_EMAIL = "a75lex@libero.it"
DONATION_URL = "https://www.paypal.com/"   # l'indirizzo si incolla come destinatario
LICENSE_VERSION = 1   # va aumentato quando cambiano le condizioni: si richiede di accettarle di nuovo
# bozza delle condizioni d'uso: da far rivedere prima di distribuire il programma
LICENSE_TEXT = f"""\
CONDIZIONI DI LICENZA D'USO

{APP_NAME} è stato ideato dall'{AUTHOR} e realizzato con l'ausilio dell'intelligenza artificiale.
© 2026 Alessandro Candido.

1. Uso. Il programma serve a consultare e modificare documenti PDF di cui si ha la disponibilità.

2. Timbri e firme. Chi usa il programma è l'unico responsabile dei timbri e delle firme che inserisce:
devono essere i propri, oppure deve essere autorizzato a usarli. È vietato usare il programma per
falsificare documenti, timbri o firme.

3. Nessuna garanzia. Il programma è fornito «così com'è», senza garanzie di alcun tipo. Controlla sempre
i documenti prodotti prima di usarli o di inviarli e conserva una copia degli originali.

4. Responsabilità. Nei limiti consentiti dalla legge, l'autore non risponde di danni diretti o indiretti
derivanti dall'uso del programma o dal suo mancato funzionamento, compresa la perdita di dati.

5. Dati. Il programma lavora sul tuo PC e non invia i documenti a nessuno. L'esportazione in Word usa
Microsoft Word installato sul PC. L'archivio dei timbri resta dove lo salvi tu.

6. Componenti di terze parti, ciascuno soggetto alla propria licenza:
   • PyMuPDF / MuPDF – Artifex Software (AGPL-3.0 o licenza commerciale)
   • Tesseract OCR, dati delle lingue – Apache License 2.0
   • Python – Python Software Foundation License
   • Tcl/Tk – licenza di tipo BSD
"""


class AboutDialog:
    """Condizioni di licenza, autore e donazione. Con must_accept serve accettarle per usare il programma."""

    def __init__(self, parent, must_accept=False):
        self.accepted = False
        top = self.top = tk.Toplevel(parent)
        top.title(f"{APP_NAME} – Informazioni e licenza")
        if parent.winfo_viewable():
            top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        head = ttk.Frame(body)
        head.pack(fill=tk.X)
        self._icon = photo_from_pixmap(icon_pixmap(56))
        tk.Label(head, image=self._icon).pack(side=tk.LEFT, padx=(0, 12))
        titles = ttk.Frame(head)
        titles.pack(side=tk.LEFT, fill=tk.X)
        ttk.Label(titles, text=f"{APP_NAME}  {APP_VERSION}", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(titles, text=f"Ideato dall'{AUTHOR}, con l'ausilio dell'intelligenza artificiale.",
                  font=("Segoe UI", 10)).pack(anchor="w")
        frame = ttk.Frame(body)
        frame.pack(fill=tk.BOTH, pady=(12, 0))
        text = tk.Text(frame, width=88, height=19, wrap="word", font=("Segoe UI", 9), relief=tk.SOLID, bd=1,
                       padx=10, pady=8, bg="#fbfbfc")
        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(side=tk.LEFT, fill=tk.BOTH)
        text.insert("1.0", LICENSE_TEXT)
        text.configure(state=tk.DISABLED)

        donate = ttk.LabelFrame(body, text=" Ti è utile? ", padding=10)
        donate.pack(fill=tk.X, pady=(12, 0))
        ttk.Label(donate, wraplength=600, justify=tk.LEFT, text=(
            "Se il programma ti fa risparmiare tempo puoi sostenerne lo sviluppo con una donazione "
            f"PayPal all'indirizzo  {DONATION_EMAIL}\n«Apri PayPal» copia l'indirizzo e apre il sito di PayPal: "
            "incollalo come destinatario.")).pack(anchor="w")
        row = ttk.Frame(donate)
        row.pack(anchor="w", pady=(8, 0))
        ttk.Button(row, text="Apri PayPal", command=self._donate).pack(side=tk.LEFT)
        ttk.Button(row, text="Copia l'indirizzo PayPal", command=self._copy).pack(side=tk.LEFT, padx=6)
        self.copied = ttk.Label(row, foreground="#2e7d32")
        self.copied.pack(side=tk.LEFT, padx=6)

        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(14, 0))
        if must_accept:
            ttk.Button(buttons, text="Non accetto (esci)", command=top.destroy).pack(side=tk.RIGHT)
            ttk.Button(buttons, text="Accetto le condizioni", command=self._accept).pack(side=tk.RIGHT, padx=6)
            top.protocol("WM_DELETE_WINDOW", top.destroy)
        else:
            ttk.Button(buttons, text="Chiudi", command=self._accept).pack(side=tk.RIGHT)
            top.bind("<Escape>", lambda e: self._accept())
        place_dialog(top, parent)

    def _donate(self):
        import webbrowser
        self._copy()   # pronto da incollare come destinatario
        webbrowser.open(DONATION_URL)

    def _copy(self):
        self.top.clipboard_clear()
        self.top.clipboard_append(DONATION_EMAIL)
        self.copied.configure(text="Indirizzo copiato.")

    def _accept(self):
        self.accepted = True
        self.top.destroy()


def enable_file_drop(widget, callback):
    """Accetta i file trascinati da Esplora risorse sulla finestra (messaggio WM_DROPFILES di Windows).

    callback riceve l'elenco dei percorsi. Restituisce l'oggetto da tenere in vita, o None se non si può.
    """
    if sys.platform != "win32":
        return None
    import ctypes
    from ctypes import wintypes
    user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
    lresult = ctypes.c_ssize_t
    wndproc_type = ctypes.WINFUNCTYPE(lresult, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    set_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
    set_long.restype = ctypes.c_void_p
    set_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
    user32.CallWindowProcW.restype = lresult
    user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM,
                                       wintypes.LPARAM]
    shell32.DragQueryFileW.restype = wintypes.UINT
    shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
    shell32.DragFinish.argtypes = [ctypes.c_void_p]
    shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
    wm_dropfiles, gwlp_wndproc = 0x0233, -4
    hwnd = widget.winfo_id()
    old = None
    pending = []

    def proc(h, msg, wparam, lparam):
        if msg == wm_dropfiles:
            files = []
            try:
                for i in range(shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)):
                    size = shell32.DragQueryFileW(wparam, i, None, 0) + 1
                    buffer = ctypes.create_unicode_buffer(size)
                    shell32.DragQueryFileW(wparam, i, buffer, size)
                    files.append(buffer.value)
            finally:
                shell32.DragFinish(wparam)
            # qui niente tkinter: Tk sta elaborando i messaggi e una chiamata annidata manda in crash Python
            pending.append(files)
            return 0
        return user32.CallWindowProcW(old, h, msg, wparam, lparam)

    def poll():
        while pending:
            callback(pending.pop(0))
        widget.after(250, poll)

    hook = wndproc_type(proc)
    old = set_long(hwnd, gwlp_wndproc, ctypes.cast(hook, ctypes.c_void_p))
    shell32.DragAcceptFiles(hwnd, True)
    widget.after(250, poll)
    return hook


# ---------------------------------------------------------------- una sola finestra, programma per i PDF
INSTANCE_FILE_NAME = "istanza.json"   # porta e chiave della finestra già aperta


def send_to_running_instance(paths):
    """Se TimbroPDF è già aperto gli passa i file da aprire (anche nessuno: si porta in primo piano).

    Restituisce True se la finestra già aperta li ha ricevuti: allora questa apertura può chiudersi.
    """
    import socket
    try:
        info = json.loads((CONFIG_DIR / INSTANCE_FILE_NAME).read_text(encoding="utf-8"))
        message = json.dumps({"token": info["token"], "files": [os.path.abspath(p) for p in paths]})
        with socket.create_connection(("127.0.0.1", int(info["port"])), timeout=2) as s:
            s.sendall(message.encode("utf-8"))
            s.shutdown(socket.SHUT_WR)
            return s.recv(8) == b"ok"
    except (OSError, ValueError, KeyError, TypeError):
        return False


class InstanceServer:
    """Riceve, solo dal PC stesso (127.0.0.1), i file aperti con doppio clic mentre il programma è già aperto.

    La chiave casuale nel file di istanza (nella cartella dell'utente) evita messaggi da altri programmi.
    """

    def __init__(self):
        import socket
        self.token = uuid.uuid4().hex
        self.pending = []          # elenchi di file arrivati, raccolti dalla finestra
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(5)
        self.file = CONFIG_DIR / INSTANCE_FILE_NAME
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self.file.write_text(json.dumps({"port": self.sock.getsockname()[1], "token": self.token,
                                         "pid": os.getpid()}), encoding="utf-8")
        threading.Thread(target=self._serve, daemon=True).start()

    def _serve(self):
        while True:
            try:
                conn, _addr = self.sock.accept()
            except OSError:
                return
            with conn:
                try:
                    conn.settimeout(3)
                    data = b""
                    while chunk := conn.recv(65536):
                        data += chunk
                    message = json.loads(data.decode("utf-8"))
                    if message.get("token") == self.token:
                        self.pending.append([str(f) for f in message.get("files", [])])
                        conn.sendall(b"ok")
                except (OSError, ValueError, AttributeError):
                    pass

    def close(self):
        try:
            if json.loads(self.file.read_text(encoding="utf-8")).get("token") == self.token:
                self.file.unlink()
        except (OSError, ValueError):
            pass
        self.sock.close()


PROG_ID = f"{APP_NAME}.Documento"


def migrate_config_dir():
    """Con il nome nuovo cambia la cartella delle impostazioni: vi passano quelle del nome vecchio
    (configurazione e archivio dei timbri). Solo la prima volta e solo nella posizione predefinita."""
    base = Path(os.environ.get("APPDATA") or Path.home())
    if CONFIG_DIR != base / APP_NAME or CONFIG_DIR.exists():
        return
    for old_name in OLD_APP_NAMES:
        old = base / old_name
        if not old.is_dir():
            continue
        try:
            os.replace(old, CONFIG_DIR)
        except OSError:   # cartella in uso (es. versione vecchia aperta): se ne fa una copia
            try:
                shutil.copytree(old, CONFIG_DIR, ignore=shutil.ignore_patterns(INSTANCE_FILE_NAME))
            except OSError:
                return
        (CONFIG_DIR / INSTANCE_FILE_NAME).unlink(missing_ok=True)
        cfg_file = CONFIG_DIR / "config.json"
        try:   # un archivio tenuto nella vecchia cartella ora sta in quella nuova
            data = json.loads(cfg_file.read_text(encoding="utf-8"))
            if str(data.get("archive") or "").startswith(str(old)):
                data["archive"] = str(CONFIG_DIR) + data["archive"][len(str(old)):]
                cfg_file.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except (OSError, ValueError):
            pass
        return


def open_command():
    """Comando con cui Windows apre un PDF con TimbroPDF (l'exe, oppure Python con questo script)."""
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"', sys.executable
    python = Path(sys.executable).with_name("pythonw.exe")
    python = python if python.exists() else Path(sys.executable)
    return f'"{python}" "{Path(__file__).resolve()}"', str(python)


def pdf_handler_entries():
    """Voci del registro (in HKEY_CURRENT_USER, solo per questo utente) che mettono TimbroPDF fra i
    programmi che aprono i PDF: (chiave, nome del valore, valore)."""
    command, icon = open_command()
    open_with = f'{command} "%1"'
    entries = [(rf"Software\Classes\{PROG_ID}", "", "Documento PDF"),
               (rf"Software\Classes\{PROG_ID}", "FriendlyTypeName", f"Documento PDF ({APP_NAME})"),
               (rf"Software\Classes\{PROG_ID}\DefaultIcon", "", f"{icon},0"),
               (rf"Software\Classes\{PROG_ID}\shell\open\command", "", open_with),
               (r"Software\Classes\.pdf\OpenWithProgids", PROG_ID, ""),
               (rf"Software\{APP_NAME}\Capabilities", "ApplicationName", APP_NAME),
               (rf"Software\{APP_NAME}\Capabilities", "ApplicationDescription",
                "Visualizza, timbra, firma e modifica i PDF"),
               (rf"Software\{APP_NAME}\Capabilities\FileAssociations", ".pdf", PROG_ID),
               (r"Software\RegisteredApplications", APP_NAME, rf"Software\{APP_NAME}\Capabilities")]
    if getattr(sys, "frozen", False):
        exe = Path(sys.executable).name
        entries += [(rf"Software\Classes\Applications\{exe}\shell\open\command", "", open_with),
                    (rf"Software\Classes\Applications\{exe}\SupportedTypes", ".pdf", "")]
    return entries


def _notify_associations_changed():
    import ctypes
    ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)   # SHCNE_ASSOCCHANGED


def register_pdf_handler():
    import winreg
    for old_name in OLD_APP_NAMES:   # le voci del vecchio nome punterebbero al vecchio programma
        unregister_pdf_handler(old_name, notify=False)
    for key, name, value in pdf_handler_entries():
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key) as k:
            winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)
    _notify_associations_changed()


def unregister_pdf_handler(app_name=APP_NAME, notify=True):
    import winreg
    prog_id = f"{app_name}.Documento"

    def delete_tree(path):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
                while True:
                    try:
                        child = winreg.EnumKey(k, 0)
                    except OSError:
                        break
                    delete_tree(rf"{path}\{child}")
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
        except FileNotFoundError:
            pass

    for path, name in ((r"Software\Classes\.pdf\OpenWithProgids", prog_id),
                       (r"Software\RegisteredApplications", app_name)):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.DeleteValue(k, name)
        except FileNotFoundError:
            pass
    for path in (rf"Software\Classes\{prog_id}", rf"Software\{app_name}",
                 rf"Software\Classes\Applications\{app_name}.exe"):
        delete_tree(path)
    if notify:
        _notify_associations_changed()


def is_default_pdf_handler():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer"
                                                      r"\FileExts\.pdf\UserChoice") as k:
            return winreg.QueryValueEx(k, "ProgId")[0] == PROG_ID
    except OSError:
        return False


def split_names(name, start, count):
    """Nomi dei file della divisione: nome_01.pdf, nome_02.pdf… con le cifre che servono all'ultimo."""
    digits = max(2, len(str(start + count - 1)))
    return [f"{name}_{n:0{digits}d}.pdf" for n in range(start, start + count)]


class SplitDialog:
    """Divisione del documento in più file: pagine per file, nome, numero iniziale e cartella."""

    def __init__(self, parent, pages, name, folder, per=1):
        self.result = None   # (pagine per file, nome, numero iniziale, cartella)
        self.pages = pages
        top = self.top = tk.Toplevel(parent)
        top.title("Dividi documento")
        top.transient(parent)
        top.resizable(False, False)
        body = ttk.Frame(top, padding=14)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text=f"Il documento ha {pages} pagine. Pagine in ogni file:").pack(anchor="w")
        row = ttk.Frame(body)
        row.pack(fill=tk.X, pady=(4, 0))
        self.per_var = tk.StringVar(value=str(per) if per in (1, 2) else "n")
        self.custom_var = tk.StringVar(value=str(per if per > 2 else 3))
        ttk.Radiobutton(row, text="1 (pagine singole)", value="1", variable=self.per_var).pack(side=tk.LEFT)
        ttk.Radiobutton(row, text="2 (pagine doppie)", value="2", variable=self.per_var).pack(side=tk.LEFT, padx=12)
        ttk.Radiobutton(row, text="Altro:", value="n", variable=self.per_var).pack(side=tk.LEFT)
        ttk.Spinbox(row, from_=1, to=max(pages, 1), width=5, textvariable=self.custom_var,
                    command=lambda: self.per_var.set("n")).pack(side=tk.LEFT, padx=(4, 0))

        grid = ttk.Frame(body)
        grid.pack(fill=tk.X, pady=(12, 0))
        self.name_var = tk.StringVar(value=name)
        self.start_var = tk.StringVar(value="1")
        self.folder_var = tk.StringVar(value=folder)
        ttk.Label(grid, text="Nome dei file:").grid(row=0, column=0, sticky="w", pady=2)
        name_entry = ttk.Entry(grid, textvariable=self.name_var, width=44)
        name_entry.grid(row=0, column=1, columnspan=2, sticky="we", padx=(8, 0))
        ttk.Label(grid, text="Numero iniziale:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Spinbox(grid, from_=0, to=99999, width=7, textvariable=self.start_var).grid(row=1, column=1, sticky="w",
                                                                                     padx=(8, 0))
        ttk.Label(grid, text="Cartella:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(grid, textvariable=self.folder_var, width=36).grid(row=2, column=1, sticky="we", padx=(8, 0))
        ttk.Button(grid, text="Sfoglia…", command=self._browse).grid(row=2, column=2, padx=(6, 0))
        self.preview = ttk.Label(body, foreground="#555", wraplength=470, justify=tk.LEFT)
        self.preview.pack(anchor="w", pady=(12, 0))
        for var in (self.per_var, self.custom_var, self.name_var, self.start_var, self.folder_var):
            var.trace_add("write", self._update_preview)
        buttons = ttk.Frame(body)
        buttons.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(buttons, text="Annulla", command=top.destroy).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Dividi e salva", command=self._ok).pack(side=tk.RIGHT, padx=6)
        top.bind("<Return>", lambda e: self._ok())
        top.bind("<Escape>", lambda e: top.destroy())
        self._update_preview()
        place_dialog(top, parent)
        name_entry.focus_set()
        name_entry.select_range(0, tk.END)

    def _browse(self):
        folder = filedialog.askdirectory(parent=self.top, initialdir=self.folder_var.get(),
                                         title="Cartella in cui salvare i file")
        if folder:
            self.folder_var.set(folder)

    def _values(self):
        """(pagine per file, nome, numero iniziale, cartella); ValueError se qualcosa non va."""
        per = int(self.custom_var.get() if self.per_var.get() == "n" else self.per_var.get())
        start = int(self.start_var.get())
        name = re.sub(r'[\\/:*?"<>|]', "", self.name_var.get()).strip()
        folder = self.folder_var.get().strip()
        if per < 1 or start < 0 or not name or not folder:
            raise ValueError
        return per, name, start, folder

    def _update_preview(self, *_):
        try:
            per, name, start, folder = self._values()
        except ValueError:
            self.preview.configure(text="Controlla i valori: servono un nome, una cartella e numeri validi.")
            return
        count = -(-self.pages // per)
        names = split_names(name, start, count)
        last = f" … {names[-1]}" if count > 1 else ""
        self.preview.configure(text=f"Verranno creati {count} file:  {names[0]}{last}")

    def _ok(self):
        try:
            self.result = self._values()
        except ValueError:
            messagebox.showinfo(APP_NAME, "Controlla i valori: servono un nome, una cartella e numeri validi.",
                                parent=self.top)
            return
        self.top.destroy()


def load_as_pdf(path, parent):
    """Apre un PDF (o un'immagine, convertita in PDF) interamente in memoria.

    Restituisce None se l'utente annulla la richiesta della password.
    """
    src = pymupdf.open(path)
    if src.is_pdf:
        src.close()
        doc = pymupdf.open(stream=Path(path).read_bytes(), filetype="pdf")
    else:
        data = src.convert_to_pdf()
        src.close()
        doc = pymupdf.open(stream=data, filetype="pdf")
    if doc.needs_pass:
        while True:
            pw = simpledialog.askstring(APP_NAME, f"«{Path(path).name}» è protetto da password.\nPassword:",
                                        show="*", parent=parent)
            if pw is None:
                return None
            if doc.authenticate(pw):
                break
        # copia decifrata in memoria: semplifica annulla e salvataggio
        doc = pymupdf.open(stream=doc.tobytes(encryption=pymupdf.PDF_ENCRYPT_NONE), filetype="pdf")
    return doc


class Config:
    def __init__(self):
        self.data = {"last_dir": str(Path.home()), "stamps": [], "current_stamp": None}
        try:
            self.data.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass

    def __getitem__(self, key):
        return self.data.get(key)

    def __setitem__(self, key, value):
        self.data[key] = value
        self.save()

    def save(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            CONFIG_FILE.write_text(json.dumps(self.data, indent=2), encoding="utf-8")
        except OSError:
            pass


ARCHIVE_NAME = "timbri.tpdf"
ARCHIVE_TYPES = [("Archivio dei timbri", "*.tpdf"), ("Tutti i file", "*.*")]


def read_archive(path, password=None):
    """Legge l'archivio dei timbri: (elenco con nomi e impostazioni, {id: (originale PNG, timbro PNG)}).

    PermissionError se è protetto e la password manca o è sbagliata.
    """
    doc = pymupdf.open(stream=Path(path).read_bytes(), filetype="pdf")
    if doc.needs_pass and not doc.authenticate(password or ""):
        raise PermissionError("password")
    items = json.loads(doc.embfile_get("timbri.json").decode("utf-8"))
    images = {item["id"]: (doc.embfile_get(f"{item['id']}-originale.png"), doc.embfile_get(f"{item['id']}-timbro.png"))
              for item in items}
    return items, images


def write_archive(path, items, images, password=None):
    """Scrive l'archivio: un PDF con i timbri allegati, cifrato AES-256 se c'è una password."""
    doc = pymupdf.open()
    page = doc.new_page(width=420, height=150)
    page.insert_text((30, 60), f"Archivio dei timbri di {APP_NAME}", fontsize=16)
    page.insert_text((30, 90), f"Si apre con {APP_NAME}, menu Timbro > Archivio dei timbri.", fontsize=10)
    doc.embfile_add("timbri.json", json.dumps(items, ensure_ascii=False).encode("utf-8"))
    for stamp_id, (original, processed) in images.items():
        doc.embfile_add(f"{stamp_id}-originale.png", original)
        doc.embfile_add(f"{stamp_id}-timbro.png", processed)
    if password:
        data = doc.tobytes(garbage=3, deflate=True, encryption=pymupdf.PDF_ENCRYPT_AES_256,
                           owner_pw=password, user_pw=password)
    else:
        data = doc.tobytes(garbage=3, deflate=True)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    os.replace(tmp, path)


def _dpapi(data, protect):
    """Cifra/decifra con la protezione dati di Windows: funziona solo per questo utente su questo PC."""
    import ctypes
    from ctypes import wintypes

    class Blob(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buffer = ctypes.create_string_buffer(data, len(data))
    blob_in, blob_out = Blob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char))), Blob()
    crypt32 = ctypes.windll.crypt32
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    fn.argtypes = [ctypes.POINTER(Blob), ctypes.c_wchar_p if protect else ctypes.c_void_p, ctypes.c_void_p,
                   ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(Blob)]
    if not fn(ctypes.byref(blob_in), "TimbroPDF" if protect else None, None, None, None, 1, ctypes.byref(blob_out)):
        raise OSError("protezione dati di Windows non riuscita")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


def protect_secret(text):
    import base64
    return base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode("ascii")


def unprotect_secret(value):
    """La password ricordata su questo PC, oppure None."""
    if not value:
        return None
    try:
        import base64
        return _dpapi(base64.b64decode(value), False).decode("utf-8")
    except Exception:
        return None


class Stamp:
    """Un timbro in memoria: l'immagine caricata (originale) e quella usata per timbrare.

    Si conserva l'originale, così lo sfondo si può togliere o rimettere in ogni momento.
    """

    def __init__(self, original=None, processed=None):
        self.original_png = original
        self.pixmap = None
        self.png = None
        self._photos = {}
        self._original = None
        self._flat = None      # originale appoggiato su bianco
        self._darkest = None   # per ogni pixel di _flat, il valore del canale più scuro
        if processed:
            try:
                self._set(pymupdf.Pixmap(processed))
            except Exception:
                pass

    @property
    def loaded(self):
        return self.pixmap is not None

    @property
    def aspect(self):
        return self.pixmap.height / self.pixmap.width

    def _set(self, pix):
        self.pixmap = pix
        self.png = pix.tobytes("png")
        self._photos.clear()

    def import_file(self, source, remove_white, tolerance, rotate=0):
        """source: percorso di un'immagine oppure un Pixmap già pronto (es. una firma creata nell'app)."""
        pix = stamp_pixmap(source if isinstance(source, pymupdf.Pixmap) else pymupdf.Pixmap(source))
        self.original_png = pix.tobytes("png")
        self._original = self._flat = self._darkest = None
        self.process(remove_white, tolerance, rotate)

    def process(self, remove_white, tolerance, rotate=0):
        """Ricrea l'immagine usata per timbrare partendo dall'originale (rotate: gradi in senso orario)."""
        if self._original is None:
            self._original = pymupdf.Pixmap(self.original_png)
        if remove_white:
            if self._darkest is None:
                self._flat = flatten_on_white(self._original)
                s = self._flat.samples
                self._darkest = bytes(map(min, s[0::3], s[1::3], s[2::3]))
            pix = pymupdf.Pixmap(self._flat, 1)
            pix.set_alpha(self._darkest.translate(white_lut(tolerance)))
        else:
            pix = self._original
        self._set(rotate_pixmap(pix, rotate))

    def photo(self, width, height):
        """Immagine Tk del timbro alle dimensioni indicate (in pixel), con cache."""
        key = (max(1, width), max(1, height))
        if key not in self._photos:
            if len(self._photos) > 30:
                self._photos.clear()
            self._photos[key] = photo_from_pixmap(pymupdf.Pixmap(self.pixmap, *key))
        return self._photos[key]

    def fit_photo(self, max_w, max_h):
        """Immagine Tk del timbro rimpicciolita per stare in un riquadro max_w × max_h."""
        s = min(max_w / self.pixmap.width, max_h / self.pixmap.height, 1)
        return self.photo(round(self.pixmap.width * s), round(self.pixmap.height * s))


class StampLibrary:
    """Timbri salvati, tutti in un unico file (ARCHIVE_NAME): immagini, nomi e impostazioni.

    L'archivio si può tenere in una cartella condivisa (es. Dropbox) per avere gli stessi timbri
    su più PC, e proteggere con una password. Su ogni PC restano solo la posizione dell'archivio,
    la password ricordata (cifrata da Windows), il timbro scelto e le ultime posizioni usate.
    """

    DEFAULTS = {"width_mm": 40.0, "remove_white": True, "white_tolerance": 10, "overlay": True, "rotate": 0,
                "last_pos": None}
    LOCAL_KEYS = ("last_pos",)   # restano su questo PC, non nell'archivio

    def __init__(self, cfg, ask_password=None, on_error=None):
        self.cfg = cfg
        self.ask_password = ask_password   # (percorso, sbagliata) -> (password, ricordala) oppure None
        self.on_error = on_error           # messaggio se l'archivio non si riesce a scrivere
        self.password = unprotect_secret(cfg["archive_pw"])
        self.items, self.stamps = [], {}
        self.current = None
        self.available = False             # False: archivio protetto non aperto o non raggiungibile
        self.error = None
        self.mtime = None
        self._written = None               # contenuto dell'ultimo salvataggio: si riscrive solo se cambia
        self._version = 0                  # cresce quando cambia un'immagine
        self.load()

    @property
    def path(self):
        return Path(self.cfg["archive"]) if self.cfg["archive"] else CONFIG_DIR / ARCHIVE_NAME

    @property
    def stamp(self):
        return self.stamps[self.current["id"]] if self.current else None

    def load(self):
        """Legge l'archivio; la prima volta vi porta i timbri salvati dalle versioni precedenti."""
        path = self.path
        self.items, self.stamps, self.available, self.error = [], {}, False, None
        if not path.exists():
            # se manca si crea (vuoto), ma solo se la cartella c'è: un disco o una cartella condivisa
            # non raggiungibili non vanno scambiati per un archivio vuoto
            if self.cfg["archive"] and not path.parent.is_dir():
                self.error = f"La cartella dell'archivio dei timbri non è raggiungibile:\n{path.parent}"
                return
            self.available = True
            self._migrate_old()
            if not path.exists():
                self._version += 1
                self.save()
        else:
            password, remember, wrong = self.password, False, False
            while True:
                try:
                    items, images = read_archive(path, password)
                    break
                except PermissionError:
                    answer = self.ask_password(path, wrong) if self.ask_password else None
                    if not answer:
                        self.error = "L'archivio dei timbri è protetto da password e non è stato aperto."
                        return
                    (password, remember), wrong = answer, True
                except Exception as exc:
                    self.error = f"Impossibile leggere l'archivio dei timbri:\n{path}\n\n{exc}"
                    return
            if remember:
                self.cfg["archive_pw"] = protect_secret(password)
            self.password = password
            self.available = True
            for item in items:
                stamp = Stamp(*images[item["id"]])
                if stamp.loaded:
                    self.items.append(dict(self.DEFAULTS, **item))
                    self.stamps[item["id"]] = stamp
            self._written = self._snapshot()
            self.mtime = path.stat().st_mtime
        positions = self.cfg["last_pos"] or {}
        for item in self.items:
            item["last_pos"] = positions.get(item["id"])
        self.current = next((i for i in self.items if i["id"] == self.cfg["current_stamp"]),
                            self.items[0] if self.items else None)

    def _snapshot(self):
        return json.dumps(self._shared_items(), sort_keys=True), self._version, self.password

    def _shared_items(self):
        return [{k: v for k, v in item.items() if k not in self.LOCAL_KEYS} for item in self.items]

    def save(self):
        """Salva le scelte di questo PC e, se è cambiato qualcosa nei timbri, riscrive l'archivio."""
        self.cfg.data["current_stamp"] = self.current["id"] if self.current else None
        self.cfg.data["last_pos"] = {i["id"]: i["last_pos"] for i in self.items if i.get("last_pos")}
        self.cfg.save()
        if not self.available or self._snapshot() == self._written:
            return
        try:
            write_archive(self.path, self._shared_items(),
                          {i["id"]: (self.stamps[i["id"]].original_png, self.stamps[i["id"]].png)
                           for i in self.items}, self.password)
        except OSError as exc:
            if self.on_error:
                self.on_error(f"Impossibile salvare l'archivio dei timbri:\n{self.path}\n\n{exc}")
            return
        self._written = self._snapshot()
        self.mtime = self.path.stat().st_mtime

    def changed_on_disk(self):
        """True se un altro PC ha modificato l'archivio (cartella condivisa) dopo l'ultima lettura."""
        try:
            return self.available and self.mtime is not None and self.path.stat().st_mtime != self.mtime
        except OSError:
            return False

    def add(self, image, name, **settings):
        """image: percorso o Pixmap; settings può contenere "source", i dati per rimodificare una firma."""
        item = dict(self.DEFAULTS, id=uuid.uuid4().hex[:10], name=name, **settings)
        stamp = Stamp()
        stamp.import_file(image, item["remove_white"], item["white_tolerance"], item["rotate"])
        self.items.append(item)
        self.stamps[item["id"]] = stamp
        self.current = item
        self._version += 1
        self.save()
        return item

    def apply_edit(self, item, result):
        """Salva quanto fatto nella finestra di modifica (StampEditor.result): restano id e impostazioni.

        Se l'immagine è stata solo ruotata o la firma riscritta, si conserva l'immagine diritta e la
        rotazione a parte, così la firma resta modificabile come testo; se è stata ritoccata a mano,
        la nuova immagine contiene già la rotazione.
        """
        item.update(name=result["name"], width_mm=result["width_mm"], remove_white=result["remove_white"])
        if result["original"] is None:
            image, item["rotate"] = result["image"], 0
            item.pop("source", None)
        else:
            image, item["rotate"] = result["original"], result["rotation"]
            if result["source"]:
                item["source"] = result["source"]
        self.stamps[item["id"]].import_file(image, item["remove_white"], item["white_tolerance"], item["rotate"])
        self._version += 1
        self.save()

    def rotate(self, item, angle):
        """Ruota il timbro di angle gradi in senso orario; le misure vere nel PDF restano quelle."""
        if angle % 180:
            item["width_mm"] = round(item["width_mm"] * self.stamps[item["id"]].aspect, 1)
        item["rotate"] = (item["rotate"] + angle) % 360
        self.reprocess(item)

    def reprocess(self, item):
        """Rifà l'immagine per timbrare dopo un cambio delle impostazioni dello sfondo o della rotazione."""
        self.stamps[item["id"]].process(item["remove_white"], item["white_tolerance"], item["rotate"])
        self._version += 1
        self.save()

    def remove(self, item):
        index = self.items.index(item)
        self.items.remove(item)
        self.stamps.pop(item["id"], None)
        if self.current is item:
            self.current = self.items[min(index, len(self.items) - 1)] if self.items else None
        self.save()

    def move_to(self, path):
        """Porta l'archivio in un altro file (es. in una cartella condivisa) e cancella quello vecchio."""
        old = self.path
        write_archive(path, self._shared_items(),
                      {i["id"]: (self.stamps[i["id"]].original_png, self.stamps[i["id"]].png) for i in self.items},
                      self.password)
        self.cfg["archive"] = str(path)
        self._written = self._snapshot()
        self.mtime = Path(path).stat().st_mtime
        if Path(old).resolve() != Path(path).resolve():
            Path(old).unlink(missing_ok=True)

    def use_archive(self, path, merge=False):
        """Passa a un archivio esistente; merge porta anche i timbri attuali che lì mancano."""
        keep = [(i, self.stamps[i["id"]]) for i in self.items] if merge and self.available else []
        old_path, old_pw = self.cfg["archive"], self.cfg["archive_pw"]
        self.cfg["archive"] = str(path)
        self.password = None
        self.cfg.data.pop("archive_pw", None)
        self.load()
        if not self.available:   # password non data o file illeggibile: si resta dove si era
            error = self.error
            self.cfg.data["archive_pw"] = old_pw
            self.cfg["archive"] = old_path
            self.password = unprotect_secret(old_pw)
            self.load()
            self.error = error
            return False
        ids = {i["id"] for i in self.items}
        added = [(i, s) for i, s in keep if i["id"] not in ids]
        for item, stamp in added:
            self.items.append(item)
            self.stamps[item["id"]] = stamp
        if added:
            self._version += 1
            self.save()
        return True

    def set_password(self, password, remember=True):
        """Protegge l'archivio con la password (None la toglie) e lo riscrive cifrato."""
        self.password = password or None
        if password and remember:
            self.cfg["archive_pw"] = protect_secret(password)
        else:
            self.cfg.data.pop("archive_pw", None)
            self.cfg.save()
        self.save()

    def _migrate_old(self):
        """Le versioni precedenti salvavano ogni timbro in una cartella (la prima un solo timbro.png)."""
        old_items = [dict(self.DEFAULTS, **i) for i in self.cfg["stamps"] or []]
        if not old_items and OLD_STAMP_FILE.exists():
            item = dict(self.DEFAULTS, id="timbro1", name="Timbro", width_mm=float(self.cfg["stamp_width_mm"] or 40))
            stamp = Stamp()
            try:
                stamp.import_file(str(OLD_STAMP_FILE), item["remove_white"], item["white_tolerance"])
                self.items.append(item)
                self.stamps[item["id"]] = stamp
            except Exception:
                pass
        for item in old_items:
            folder = STAMPS_DIR / item["id"]
            try:
                stamp = Stamp((folder / "originale.png").read_bytes(), (folder / "timbro.png").read_bytes())
            except OSError:
                continue
            if stamp.loaded:
                self.items.append(item)
                self.stamps[item["id"]] = stamp
        positions = {i["id"]: i["last_pos"] for i in self.items if i.get("last_pos")}
        if not self.items:
            return
        # save() annota anche il timbro scelto: deve restare quello di prima
        self.current = next((i for i in self.items if i["id"] == self.cfg["current_stamp"]), None)
        self._version += 1
        self.save()
        try:   # i vecchi file si tolgono solo se l'archivio si rilegge bene
            if len(read_archive(self.path, self.password)[0]) != len(self.items):
                return
        except Exception:
            return
        self.cfg.data.pop("stamps", None)
        self.cfg.data["last_pos"] = positions
        self.cfg.save()
        if STAMPS_DIR.exists():
            backup = CONFIG_DIR / "timbri_versione_precedente"
            shutil.rmtree(backup, ignore_errors=True)
            shutil.move(str(STAMPS_DIR), str(backup))
        OLD_STAMP_FILE.unlink(missing_ok=True)


class DocTab:
    """Un PDF aperto in una scheda, con il suo annulla/ripristina e la posizione nella vista."""

    _new_count = 0

    def __init__(self, doc, path=None, label="Nuovo documento"):
        self.doc = doc
        self.path = path
        self.label = label
        self.number = 0
        if doc is not None and path is None:
            DocTab._new_count += 1
            self.number = DocTab._new_count
        self.dirty = False
        self.undo_stack, self.redo_stack = [], []
        self.current = 0
        self.selection = {0} if doc is not None and len(doc) else set()
        self.anchor = 0
        self.stamp_xrefs = {}      # id timbro -> immagine già presente nel documento
        self.last_stamp = None     # (pagina, x relativa, y relativa, id timbro)
        self.thumb_photos = {}
        self.view_pos = (0.0, 0.0)
        self.thumbs_pos = 0.0
        self.thumb_w = 0           # misura delle miniature in thumb_photos
        self.ov_photos = {}        # miniature della vista d'insieme
        self.fonts_added = False   # testo scritto con font di sistema: al salvataggio si riducono

    @property
    def doc(self):
        return self._doc

    @doc.setter
    def doc(self, doc):
        self._doc = doc
        self.font_token = uuid.uuid4().hex[:6]   # vedi font_resource()

    @property
    def name(self):
        return Path(self.path).name if self.path else f"{self.label} {self.number}"


def _per_tab(name):
    """Attributo dell'App che appartiene in realtà alla scheda attiva."""
    return property(lambda self: getattr(self.tab, name), lambda self, value: setattr(self.tab, name, value))


class App:
    doc = _per_tab("doc")
    path = _per_tab("path")
    dirty = _per_tab("dirty")
    undo_stack = _per_tab("undo_stack")
    redo_stack = _per_tab("redo_stack")
    current = _per_tab("current")
    selection = _per_tab("selection")
    anchor = _per_tab("anchor")
    stamp_xrefs = _per_tab("stamp_xrefs")
    last_stamp = _per_tab("last_stamp")
    thumb_photos = _per_tab("thumb_photos")

    def __init__(self, root, path=None):
        self.root = root
        migrate_config_dir()
        self.cfg = Config()
        self.library = None            # si apre dopo aver costruito la finestra (può chiedere la password)
        self._archive_checked = 0.0
        self.tabs = []
        self.no_tab = DocTab(None)     # "scheda" usata quando non ci sono PDF aperti
        self.tab = self.no_tab
        self._tab_frames = []
        self._plus = None
        self._drop_tab = None
        self._page_drag = None         # trascinamento di miniature
        self.overview = False          # vista d'insieme delle miniature di tutti i PDF aperti
        self.ov_layout = []            # per colonna: (scheda, x0, x1, riquadri delle pagine)
        self._ov_job = self._ov_render_job = None
        self._ov_drag = None
        self._tab_drag = None          # trascinamento di una scheda per riordinarle
        self._dock_drag = None         # trascinamento di una barra o di un pannello
        self.layout = self._load_layout()
        self._normal_geometry = None   # posizione della finestra quando non è ingrandita
        self.thumb_w = THUMB_W         # larghezza delle miniature: segue quella del pannello
        self._thumb_resize_job = None
        self._gallery_cols = 2
        self._fonts = None
        self.zoom_mode = "width"       # "width", "page" oppure None (zoom libero)
        self.zoom = 1.0                # zoom libero, 1.0 = dimensione reale
        self.cur_zoom = 1.0            # pixel per punto della pagina mostrata
        self.page_origin = (0, 0)
        self.scroll_size = (1, 1)
        self.page_photo = None
        self.thumb_layout = []         # per pagina: (y0, y1, x0, scala)
        self.thumb_total = 1           # lunghezza della fila di miniature (per lo scorrimento)
        self._thumb_job = None
        self._render_job = None
        self._cfg_job = None
        self._white_job = None
        self._syncing = False
        self._last_flip = 0.0
        self._mouse = None
        self._ghost_photo = None
        self._preview_photo = None
        self._gallery_photos = {}
        self._loading_settings = False
        self.tool = None               # None, "stamp" oppure uno di TEXT_TOOLS
        self.page_stamps = []          # timbri e caselle di testo già sulla pagina mostrata
        self.sel_stamp = None          # indice (in /Contents) del timbro/casella selezionato
        self._sel_page = None
        self._drag = None              # trascinamento di un timbro o di una casella
        self._page_lines = None        # righe di testo della pagina mostrata (modifica testo)
        self._area = None              # rettangolo del bianchetto in corso
        self._find = None              # ricerca: testo, documento, risultati, posizione
        self.px_per_pt = root.winfo_fpixels("1i") / 72

        self._build_ui()
        self._bind_keys()
        root.protocol("WM_DELETE_WINDOW", self.quit)
        self.library = StampLibrary(self.cfg, self._ask_archive_password, self._archive_error)
        if self.library.error:
            root.after(300, lambda: messagebox.showwarning(
                APP_NAME, self.library.error + "\n\nPuoi aprirlo da Timbro ▸ Archivio dei timbri.", parent=root))
        root.bind("<FocusIn>", self._on_focus_in, add=True)
        try:   # file trascinati da Esplora risorse
            self._drop_hook = enable_file_drop(root, self._on_files_dropped)
        except Exception:
            self._drop_hook = None
        try:   # PDF aperti con doppio clic mentre il programma è già aperto: arrivano qui
            self.instance = InstanceServer()
            root.after(300, self._poll_instance)
        except OSError:
            self.instance = None
        self._load_stamp_settings()
        self._refresh_stamp_panel()
        self._doc_changed(structure=True)
        paths = [path] if isinstance(path, str) else list(path or [])
        if paths:
            root.after(100, lambda: [self.open_file(p) for p in paths])
        root.after(400, self._check_license)

    @property
    def stamp_mode(self):
        return self.tool == "stamp"

    @property
    def fonts(self):
        """Font installati (nome -> file), letti una volta sola."""
        if self._fonts is None:
            self._fonts = installed_fonts()
        return self._fonts

    @property
    def stamp(self):
        """Timbro selezionato (None se non ce ne sono)."""
        return self.library.stamp

    @property
    def stamp_item(self):
        """Nome e impostazioni del timbro selezionato."""
        return self.library.current

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        root = self.root
        root.title(APP_NAME)
        root.geometry(f"{min(1360, root.winfo_screenwidth() - 60)}x{min(900, root.winfo_screenheight() - 100)}")
        root.minsize(980, 620)
        self._restore_window()
        self._icon = photo_from_pixmap(icon_pixmap(64))
        root.iconphoto(True, self._icon)
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        bg = style.lookup("TFrame", "background") or root.cget("bg")

        self._build_menu()

        # zone in cui agganciare le barre: create prima delle barre, che devono stare sopra di loro
        self.docks = {side: ttk.Frame(root) for side in ("top", "bottom", "left", "right")}
        status = ttk.Frame(root, padding=(8, 2))
        status.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var = tk.StringVar(value="Pronto.")
        ttk.Label(status, textvariable=self.status_var).pack(side=tk.LEFT)
        self.info_var = tk.StringVar()
        ttk.Label(status, textvariable=self.info_var).pack(side=tk.RIGHT)
        self.bars = {"main": self._build_main_bar(), "text": self._build_text_bar()}

        # pannelli e documento, separati da divisori trascinabili (la riga grigia):
        # vpaned ha in alto/in basso i pannelli agganciati lì e in mezzo paned, con quelli a sinistra/destra.
        # I pannelli sono figli della finestra, così possono passare da un divisore all'altro.
        self.vpaned = tk.PanedWindow(root, orient=tk.VERTICAL, sashwidth=7, bd=0, bg=SASH_BG, sashrelief=tk.FLAT)
        self.paned = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashwidth=7, bd=0, bg=SASH_BG, sashrelief=tk.FLAT)
        self.side_panels = {}
        self.thumbs_horizontal = False
        self._stamps_horizontal = False
        self._build_thumbs_panel()
        self._build_stamps_panel(bg)
        self._build_center()
        self._dock_indicator = tk.Frame(root, bg="#1a73e8")
        self._apply_layout()
        root.bind("<Configure>", self._on_root_configure, add=True)

    def _build_main_bar(self):
        bar = ToolBar(self.root, "main", self._on_grip)
        for text, cmd in (("Apri…", self.cmd_open), ("Nuovo", self.cmd_new),
                          ("Aggiungi…", lambda: self.cmd_add(at_end=True)), ("Salva…", self.cmd_save)):
            bar_button(bar.group(), text, cmd)
        bar.separator()
        bar_button(bar.group(), "↶ Annulla", self.undo)
        bar_button(bar.group(), "↷ Ripristina", self.redo)
        bar.separator()
        g = bar.group()
        bar_button(g, "◀", self.prev_page, width=3)
        self.page_var = tk.StringVar()
        entry = ttk.Entry(g, textvariable=self.page_var, width=5, justify="center")
        entry.pack(side=tk.LEFT, padx=2)
        entry.bind("<Return>", self._goto_entry)
        self.page_count_lbl = ttk.Label(g, text="/ 0", width=6)
        self.page_count_lbl.pack(side=tk.LEFT)
        bar_button(g, "▶", self.next_page, width=3)
        bar.separator()
        g = bar.group()
        bar_button(g, "−", lambda: self._zoom_at(1 / 1.25), width=3)
        self.zoom_lbl = ttk.Label(g, text="100%", width=6, anchor="center")
        self.zoom_lbl.pack(side=tk.LEFT)
        bar_button(g, "+", lambda: self._zoom_at(1.25), width=3)
        bar_button(bar.group(), "Larghezza", lambda: self.set_zoom_mode("width"))
        bar_button(bar.group(), "Pagina intera", lambda: self.set_zoom_mode("page"))
        bar.separator()
        for text, cmd in (("↺ Ruota sx", lambda: self.rotate(-90)), ("↻ Ruota dx", lambda: self.rotate(90)),
                          ("▲ Su", lambda: self.move_pages(-1)), ("▼ Giù", lambda: self.move_pages(1)),
                          ("✕ Elimina", self.delete_pages)):
            bar_button(bar.group(), text, cmd)
        return bar

    def _build_text_bar(self):
        bar = ToolBar(self.root, "text", self._on_grip)
        ttk.Label(bar.group(), text="Testo:").pack(side=tk.LEFT, padx=(2, 4))
        self.tool_var = tk.StringVar()
        for value, text in (("edit_text", "✎ Modifica testo"), ("add_text", "+ Aggiungi testo"),
                            ("whiteout", "Bianchetto"), ("blackout", "Oscura")):
            ttk.Radiobutton(bar.group(), text=text, value=value, variable=self.tool_var, style="Toolbutton",
                            command=lambda v=value: self.set_tool(None if self.tool == v else v)
                            ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        bar.separator()
        bar_button(bar.group(), "Riconosci testo (OCR)…", self.cmd_ocr)
        bar_button(bar.group(), "Esporta in Word…", self.cmd_export_word)
        bar.separator()
        bar_button(bar.group(), "Unisci schede…", self.cmd_merge_tabs)
        bar_button(bar.group(), "Dividi…", self.cmd_split)
        self.overview_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar.group(), text="▦ Vista d'insieme", style="Toolbutton", variable=self.overview_var,
                        command=lambda: self.toggle_overview(self.overview_var.get())
                        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=1)
        bar.separator()
        g = bar.group()
        ttk.Label(g, text="Cerca:").pack(side=tk.LEFT, padx=(0, 4))
        self.find_var = tk.StringVar()
        self.find_entry = ttk.Entry(g, textvariable=self.find_var, width=20)
        self.find_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.find_entry.bind("<Return>", lambda e: self.find(1) or "break")
        self.find_entry.bind("<Shift-Return>", lambda e: self.find(-1) or "break")
        self.find_entry.bind("<Escape>", lambda e: self.clear_find() or "break")
        g = bar.group()
        bar_button(g, "▲", lambda: self.find(-1), width=3)
        bar_button(g, "▼", lambda: self.find(1), width=3)
        self.find_lbl = ttk.Label(g, foreground="#555", width=10)
        self.find_lbl.pack(side=tk.LEFT, padx=6)
        return bar

    def _panel_header(self, panel, key, title):
        """Intestazione di un pannello laterale: trascinandola si sposta il pannello a destra o a sinistra."""
        head = tk.Label(panel, text="⋮⋮  " + title, anchor="w", bg="#d9dae0", fg="#444", cursor="fleur",
                        font=("Segoe UI", 9), padx=6, pady=2)
        head.pack(side=tk.TOP, fill=tk.X)
        for seq, phase in (("<Button-1>", "press"), ("<B1-Motion>", "motion"), ("<ButtonRelease-1>", "release")):
            head.bind(seq, lambda e, p=phase: self._on_grip(key, e, p))

    def _build_thumbs_panel(self):
        panel = self.side_panels["thumbs"] = ttk.Frame(self.root)
        self._panel_header(panel, "thumbs", "Pagine")
        left = ttk.Frame(panel)
        left.pack(fill=tk.BOTH, expand=True)
        self.thumbs = tk.Canvas(left, width=THUMB_W + 34, height=120, bg="#e4e4ea", highlightthickness=0)
        # barra verticale con il pannello a lato, orizzontale con il pannello in alto/in basso
        self._thumbs_sb = {"v": ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.thumbs.yview),
                           "h": ttk.Scrollbar(left, orient=tk.HORIZONTAL, command=self.thumbs.xview)}

        def scrolled(bar):
            def command(a, b):
                self._thumbs_sb[bar].set(a, b)
                self._schedule_thumbs()
            return command

        self.thumbs.configure(yscrollcommand=scrolled("v"), xscrollcommand=scrolled("h"))
        self._thumbs_sb["v"].pack(side=tk.RIGHT, fill=tk.Y)
        self.thumbs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.thumbs.bind("<Configure>", self._on_thumbs_configure)
        self.thumbs.bind("<MouseWheel>", self._on_thumbs_wheel)
        self.thumbs.bind("<Button-1>", lambda e: self._on_thumb_click(e, "single"))
        self.thumbs.bind("<Control-Button-1>", lambda e: self._on_thumb_click(e, "toggle"))
        self.thumbs.bind("<Shift-Button-1>", lambda e: self._on_thumb_click(e, "range"))
        self.thumbs.bind("<B1-Motion>", self._on_thumb_motion)
        self.thumbs.bind("<ButtonRelease-1>", self._on_thumb_release)
        self.thumbs.bind("<Button-3>", self._on_thumb_menu)
        self.thumbs.configure(yscrollincrement=20, xscrollincrement=20)

    def _build_stamps_panel(self, bg):
        """Pannello dei timbri, con barre di scorrimento se lo spazio non basta; il contenuto segue il pannello."""
        panel = self.side_panels["stamps"] = ttk.Frame(self.root)
        self._panel_header(panel, "stamps", "Timbri e firme")
        right = ttk.Frame(panel)
        right.pack(fill=tk.BOTH, expand=True)
        self.panel = tk.Canvas(right, highlightthickness=0, bg=bg, yscrollincrement=20, xscrollincrement=20,
                               width=10, height=10)
        psb = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.panel.yview)
        phs = ttk.Scrollbar(right, orient=tk.HORIZONTAL, command=self.panel.xview)
        self.panel.configure(yscrollcommand=psb.set, xscrollcommand=phs.set)
        self.panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner = self._stamps_inner = ttk.Frame(self.panel, padding=(8, 8))
        window = self.panel.create_window(0, 0, window=inner, anchor="nw")

        def fit_panel(_event=None):
            cw, ch = self.panel.winfo_width(), self.panel.winfo_height()
            rw, rh = inner.winfo_reqwidth(), inner.winfo_reqheight()
            # larghezza/altezza 0: quella richiesta dal contenuto
            if self._stamps_horizontal:   # in alto/in basso: le colonne scorrono in orizzontale
                self.panel.itemconfigure(window, width=0, height=0)
                self.panel.configure(scrollregion=(0, 0, rw, 0))
                scroll, other, needed = phs, psb, rw > cw > 1
            else:                         # a lato: il contenuto si allarga col pannello e scorre in verticale
                self.panel.itemconfigure(window, width=max(cw, rw), height=0)
                self.panel.configure(scrollregion=(0, 0, 0, rh))
                scroll, other, needed = psb, phs, rh > ch > 1
                self._wrap_stamp_labels(max(cw, rw))
            other.pack_forget()
            if needed:
                scroll.pack(side=tk.BOTTOM if scroll is phs else tk.RIGHT, fill=tk.X if scroll is phs else tk.Y,
                            before=self.panel)
            else:
                scroll.pack_forget()
                self.panel.xview_moveto(0)
                self.panel.yview_moveto(0)

        self._fit_stamps_panel = fit_panel
        inner.bind("<Configure>", fit_panel)
        self.panel.bind("<Configure>", fit_panel)
        self._build_stamp_panel(inner)
        self.root.bind_all("<MouseWheel>", self._on_panel_wheel, add=True)

    def _build_center(self):
        """Schede dei PDF aperti e vista della pagina."""
        center = self.center = ttk.Frame(self.root)
        self.tabbar = tk.Frame(center, bg=TABBAR_BG)
        self.tabbar.pack(side=tk.TOP, fill=tk.X)
        self.tabbar.bind("<Configure>", lambda e: self._refresh_tabbar())
        self.tabbar.bind("<Double-Button-1>", lambda e: self.cmd_new())
        self._tab_marker = tk.Frame(self.tabbar, bg="#1a73e8", width=3)
        self._refresh_tabbar()
        # vista d'insieme: le miniature di tutti i PDF aperti, al posto della pagina
        ov = self.overview_area = ttk.Frame(center)
        self.ov = tk.Canvas(ov, bg="#5f6068", highlightthickness=0, xscrollincrement=20, yscrollincrement=20)
        ov_v = ttk.Scrollbar(ov, orient=tk.VERTICAL, command=self.ov.yview)
        ov_h = ttk.Scrollbar(ov, orient=tk.HORIZONTAL, command=self.ov.xview)

        def ov_scrolled(bar):
            def command(a, b):
                bar.set(a, b)
                self._schedule_overview_render()
            return command

        self.ov.configure(yscrollcommand=ov_scrolled(ov_v), xscrollcommand=ov_scrolled(ov_h))
        ov_v.pack(side=tk.RIGHT, fill=tk.Y)
        ov_h.pack(side=tk.BOTTOM, fill=tk.X)
        self.ov.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.ov.bind("<Button-1>", lambda e: self._ov_press(e, "single"))
        self.ov.bind("<Control-Button-1>", lambda e: self._ov_press(e, "toggle"))
        self.ov.bind("<Shift-Button-1>", lambda e: self._ov_press(e, "range"))
        self.ov.bind("<B1-Motion>", self._ov_motion)
        self.ov.bind("<ButtonRelease-1>", self._ov_release)
        self.ov.bind("<Double-Button-1>", self._ov_double)
        self.ov.bind("<MouseWheel>", self._ov_wheel)
        self.ov.bind("<Configure>", lambda e: self._schedule_overview_render())
        area = self.view_area = ttk.Frame(center)
        area.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.view = tk.Canvas(area, bg="#5f6068", highlightthickness=0,
                              xscrollincrement=20, yscrollincrement=20)
        vsb = ttk.Scrollbar(area, orient=tk.VERTICAL, command=self.view.yview)
        hsb = ttk.Scrollbar(area, orient=tk.HORIZONTAL, command=self.view.xview)
        self.view.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.view.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v = self.view
        v.bind("<Configure>", self._on_view_configure)
        v.bind("<Motion>", self._on_motion)
        v.bind("<Leave>", self._on_leave)
        v.bind("<Button-1>", self._on_press)
        v.bind("<B1-Motion>", self._on_drag)
        v.bind("<ButtonRelease-1>", self._on_release)
        v.bind("<Double-Button-1>", self._on_double_click)
        v.bind("<Button-2>", lambda e: v.scan_mark(e.x, e.y))
        v.bind("<B2-Motion>", lambda e: v.scan_dragto(e.x, e.y, gain=1))
        v.bind("<MouseWheel>", self._on_wheel)
        v.bind("<Button-3>", self._on_view_menu)

    # ------------------------------------------------------------------ disposizione
    def _load_layout(self):
        layout = default_layout()
        saved = self.cfg["layout"] or {}
        for section in ("bars", "panels"):
            for key, value in (saved.get(section) or {}).items():
                if key in layout[section] and isinstance(value, list) and len(value) == 2 and value[0] in SIDE_NAMES:
                    layout[section][key] = [value[0], int(value[1])]
        for section in ("widths", "heights"):
            for key, value in (saved.get(section) or {}).items():
                if key in layout[section] and isinstance(value, int) and 0 <= value < 3000:
                    layout[section][key] = value
        if isinstance(saved.get("gallery_h"), int) and 0 <= saved["gallery_h"] < 2000:
            layout["gallery_h"] = saved["gallery_h"]
        return layout

    def _save_layout(self):
        for key, panel in self.side_panels.items():
            if panel.winfo_ismapped() and panel.winfo_width() > 1:
                if self.layout["panels"][key][0] in ("left", "right"):
                    self.layout["widths"][key] = panel.winfo_width()
                else:
                    self.layout["heights"][key] = panel.winfo_height()
        self.cfg["layout"] = self.layout

    def _apply_layout(self):
        """Aggancia barre e pannelli secondo self.layout."""
        bars, panels = self.layout["bars"], self.layout["panels"]
        for dock in self.docks.values():
            dock.pack_forget()
        self.vpaned.pack_forget()
        for bar in self.bars.values():
            bar.pack_forget()
        for side in ("top", "bottom", "left", "right"):
            here = sorted((b for b in self.bars.values() if bars[b.key][0] == side), key=lambda b: bars[b.key][1])
            if not here:
                continue
            dock = self.docks[side]
            horizontal = side in ("top", "bottom")
            dock.pack(side=side, fill=tk.X if horizontal else tk.Y)
            for bar in here:
                bar.orient(side)
                bar.pack(in_=dock, side=tk.TOP if horizontal else tk.LEFT, fill=tk.X if horizontal else tk.Y)
                bar.lift()
        self.vpaned.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        for paned in (self.vpaned, self.paned):
            for pane in paned.panes():
                paned.forget(pane)

        def on(side, outer_first=True):   # pannelli di un lato, dal bordo verso il centro o viceversa
            keys = [k for k, v in panels.items() if v[0] == side]
            return sorted(keys, key=lambda k: panels[k][1] if outer_first else -panels[k][1])

        for key in on("top"):
            self._add_side_panel(key, self.vpaned)
        self.vpaned.add(self.paned, stretch="always", minsize=200)
        for key in on("bottom", outer_first=False):
            self._add_side_panel(key, self.vpaned)
        for key in on("left"):
            self._add_side_panel(key, self.paned)
        self.paned.add(self.center, stretch="always", minsize=300)
        for key in on("right", outer_first=False):
            self._add_side_panel(key, self.paned)
        for widget in (self.paned, self.center, *self.side_panels.values()):
            widget.lift()   # sopra ai divisori di cui fanno parte
        self._gallery_height(self.layout["gallery_h"] or GALLERY_H)

    def _add_side_panel(self, key, paned):
        horizontal = paned is self.vpaned   # in alto/in basso il contenuto si dispone in orizzontale
        if key == "thumbs":
            self._orient_thumbs(horizontal)
        else:
            self._orient_stamps(horizontal)
        panel = self.side_panels[key]
        panel.update_idletasks()
        if horizontal:
            minimum = 120 if key == "thumbs" else self._stamps_inner.winfo_reqheight() + 30
            size = max(self.layout["heights"].get(key) or 0, minimum)
            paned.add(panel, stretch="never", minsize=min(minimum, 400), height=size)
        else:
            minimum = 100 if key == "thumbs" else self._stamps_inner.winfo_reqwidth() + 4
            size = max(self.layout["widths"].get(key) or 0, minimum)
            paned.add(panel, stretch="never", minsize=minimum, width=size)

    def _orient_thumbs(self, horizontal):
        """Miniature in colonna (pannello a lato) o in fila (pannello in alto/in basso)."""
        if horizontal == self.thumbs_horizontal:
            return
        self.thumbs_horizontal = horizontal
        for bar in self._thumbs_sb.values():
            bar.pack_forget()
        if horizontal:
            self._thumbs_sb["h"].pack(side=tk.BOTTOM, fill=tk.X, before=self.thumbs)
        else:
            self._thumbs_sb["v"].pack(side=tk.RIGHT, fill=tk.Y, before=self.thumbs)
        self.thumbs.xview_moveto(0)
        self.thumbs.yview_moveto(0)
        self._layout_thumbs()

    def _orient_stamps(self, horizontal):
        """Pannello dei timbri in colonna (a lato) o a colonne affiancate (in alto/in basso)."""
        if horizontal == self._stamps_horizontal and self._stamp_sections[0].winfo_manager():
            return
        self._stamps_horizontal = horizontal
        sections, columns = self._stamp_sections, self._stamp_columns
        for w in (*sections, *columns):
            w.pack_forget()
        gallery, selected, help_label = sections
        if horizontal:
            gallery.pack(side=tk.LEFT, fill=tk.Y)
            selected.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))
            for k, column in enumerate(columns):
                column.pack(side=tk.LEFT, fill=tk.Y, anchor="n", padx=(0 if k == 0 else 14, 0))
            help_label.configure(wraplength=260)
            self.overlay_note.configure(wraplength=PREVIEW_W)
            help_label.pack(side=tk.LEFT, anchor="n", padx=(12, 0))
        else:
            gallery.pack(side=tk.TOP, fill=tk.X)
            selected.pack(side=tk.TOP, fill=tk.X, pady=(8, 0))
            for column in columns:
                column.pack(side=tk.TOP, fill=tk.X)
            help_label.pack(side=tk.TOP, anchor="w", pady=(10, 0))
        self._fit_stamps_panel()

    def _wrap_stamp_labels(self, width):
        """Le scritte del pannello dei timbri vanno a capo secondo la sua larghezza."""
        help_label = self._stamp_sections[2]
        if int(help_label.cget("wraplength")) != max(width - 24, 120):
            help_label.configure(wraplength=max(width - 24, 120))
            self.overlay_note.configure(wraplength=max(width - 44, 120))

    def _gallery_height(self, height):
        height = min(max(int(height), TILE_H + 30), 1200)
        if int(self.gallery.cget("height")) != height:
            self.gallery.configure(height=height)
            self._draw_gallery()

    def _on_gallery_grip(self, event, phase):
        """Trascinando la barretta sotto la galleria la si allunga o la si accorcia."""
        if phase == "press":
            self._gallery_drag = (event.y_root, int(self.gallery.cget("height")))
        elif phase == "motion" and self._gallery_drag:
            y0, h0 = self._gallery_drag
            self._gallery_height(h0 + event.y_root - y0)
        elif self._gallery_drag:
            self._gallery_drag = None
            self.layout["gallery_h"] = int(self.gallery.cget("height"))
            self._save_layout()

    def reset_layout(self):
        self.layout = default_layout()
        self._apply_layout()
        self._save_layout()
        self.status("Barre e pannelli rimessi nella posizione iniziale.")

    def _on_grip(self, item, event, phase):
        """Trascinamento della maniglia di una barra (ToolBar) o dell'intestazione di un pannello (chiave)."""
        root = self.root
        rx, ry = event.x_root - root.winfo_rootx(), event.y_root - root.winfo_rooty()
        if phase == "press":
            self._dock_drag = {"item": item, "x": event.x_root, "y": event.y_root, "side": None, "moved": False}
            return
        d = self._dock_drag
        if not d:
            return
        if phase == "motion":
            if not d["moved"] and abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 8:
                return
            d["moved"] = True
            side = self._dock_side(item, event.x_root, event.y_root)
            d.update(side=side, pos=(rx, ry))
            what = "la barra" if isinstance(item, ToolBar) else "il pannello"
            if side:
                self._show_dock_indicator(side)
                self.status(f"Rilascia per agganciare {what} {SIDE_NAMES[side]}.")
            else:
                self._dock_indicator.place_forget()
                self.status(f"Porta {what} verso un bordo della finestra: in alto, in basso, a sinistra o a destra.")
            return
        self._dock_drag = None
        self._dock_indicator.place_forget()
        if d["moved"] and d["side"]:
            self._dock(item, d["side"], d["pos"])

    def _dock_side(self, item, x_root, y_root):
        """Lato a cui agganciare, secondo la zona della finestra in cui si trova il mouse (None: nessuno).

        Le fasce in alto e in basso hanno la precedenza, così si raggiungono anche partendo da un lato.
        """
        root = self.root
        rx, ry = x_root - root.winfo_rootx(), y_root - root.winfo_rooty()
        w, h = root.winfo_width(), root.winfo_height()
        if not isinstance(item, ToolBar):   # finché il mouse resta sopra al pannello non lo si sposta
            panel = self.side_panels[item]
            if (0 <= x_root - panel.winfo_rootx() < panel.winfo_width()
                    and 0 <= y_root - panel.winfo_rooty() < panel.winfo_height()):
                return None
        band_y, band_x = max(90, h * 0.15), max(110, w * 0.12)
        if ry < band_y:
            return "top"
        if ry > h - band_y:
            return "bottom"
        if rx < band_x:
            return "left"
        if rx > w - band_x:
            return "right"
        return None

    def _show_dock_indicator(self, side):
        t = 6
        place = {"top": dict(x=0, y=0, relwidth=1, height=t), "bottom": dict(x=0, rely=1, y=-t, relwidth=1, height=t),
                 "left": dict(x=0, y=0, width=t, relheight=1), "right": dict(relx=1, x=-t, y=0, width=t, relheight=1)}
        self._dock_indicator.place(**place[side])
        self._dock_indicator.lift()

    def _dock(self, item, side, pos):
        """Aggancia la barra o il pannello al lato, nel punto di rilascio rispetto a quelli già lì."""
        root = self.root
        self._save_layout()   # tiene le larghezze attuali dei pannelli
        if isinstance(item, ToolBar):
            bars = self.layout["bars"]
            axis = 1 if side in ("top", "bottom") else 0
            others = sorted((b for b in self.bars.values() if b is not item and bars[b.key][0] == side),
                            key=lambda b: bars[b.key][1])

            def middle(b):
                return (b.winfo_rooty() - root.winfo_rooty() + b.winfo_height() / 2 if axis
                        else b.winfo_rootx() - root.winfo_rootx() + b.winfo_width() / 2)

            # nello stesso lato le barre vanno dall'alto in basso o da sinistra a destra
            others.insert(sum(1 for b in others if middle(b) < pos[axis]), item)
            for i, b in enumerate(others):
                bars[b.key] = [side, i]
            what = "Barra agganciata"
        else:
            panels = self.layout["panels"]
            others = sorted((k for k, v in panels.items() if k != item and v[0] == side), key=lambda k: panels[k][1])
            axis = 1 if side in ("top", "bottom") else 0

            def middle(k):
                p = self.side_panels[k]
                return (p.winfo_rooty() - root.winfo_rooty() + p.winfo_height() / 2 if axis
                        else p.winfo_rootx() - root.winfo_rootx() + p.winfo_width() / 2)

            # l'indice conta dal bordo della finestra verso il centro
            if side in ("left", "top"):
                k = sum(1 for o in others if middle(o) < pos[axis])
            else:
                k = sum(1 for o in others if middle(o) > pos[axis])
            others.insert(k, item)
            for i, key in enumerate(others):
                panels[key] = [side, i]
            what = "Pannello spostato"
        self._apply_layout()
        self._save_layout()
        self.status(f"{what} {SIDE_NAMES[side]}. La posizione sarà ricordata alla prossima apertura "
                    "(Visualizza ▸ Ripristina disposizione per tornare a quella iniziale).")

    def _restore_window(self):
        saved = self.cfg["window"] or {}
        if on_screen(saved.get("geometry")):
            self.root.geometry(saved["geometry"])
            self._normal_geometry = saved["geometry"]
        if saved.get("zoomed"):
            self.root.state("zoomed")

    def _on_root_configure(self, event):
        if event.widget is self.root and self.root.state() == "normal":
            self._normal_geometry = self.root.geometry()

    def _save_window(self):
        self.cfg["window"] = {"geometry": self._normal_geometry or self.root.geometry(),
                              "zoomed": self.root.state() == "zoomed"}

    def _thumb_box(self):
        """Riquadro massimo (larghezza, altezza) di una miniatura: segue la larghezza del pannello a lato,
        o la sua altezza quando è in alto/in basso."""
        if self.thumbs_horizontal:
            h = max(self.thumbs.winfo_height() - 44, 60)
            return round(h * THUMB_W / THUMB_H), h
        return self.thumb_w, round(self.thumb_w * THUMB_H / THUMB_W)

    def _on_thumbs_configure(self, event):
        """Le miniature si ingrandiscono o rimpiccioliscono insieme al pannello."""
        width = max(event.width - 34, 60)
        size = (event.height, self.thumbs_horizontal) if self.thumbs_horizontal else (width, False)
        if size == getattr(self, "_thumbs_size", None):
            self._schedule_thumbs()
            return
        self._thumbs_size = size
        if not self.thumbs_horizontal:
            self.thumb_w = width
        if self._thumb_resize_job:
            self.root.after_cancel(self._thumb_resize_job)
        self._thumb_resize_job = self.root.after(150, self._resize_thumbs)

    def _resize_thumbs(self):
        self._thumb_resize_job = None
        start = self.thumbs.xview()[0] if self.thumbs_horizontal else self.thumbs.yview()[0]
        self._layout_thumbs()
        (self.thumbs.xview_moveto if self.thumbs_horizontal else self.thumbs.yview_moveto)(start)

    def _build_stamp_panel(self, parent):
        gf = ttk.LabelFrame(parent, text=" Timbri salvati ", padding=8)
        self.gallery = tk.Canvas(gf, width=PREVIEW_W, height=GALLERY_H, bg="white", yscrollincrement=10,
                                 highlightthickness=1, highlightbackground="#b8b8c0")
        self.gallery.pack(fill=tk.X)
        self.gallery.bind("<Button-1>", self._on_gallery_click)
        self.gallery.bind("<Double-Button-1>", self._on_gallery_double)
        self.gallery.bind("<Button-3>", self._on_gallery_menu)
        self.gallery.bind("<MouseWheel>", self._on_gallery_wheel)
        self.gallery.bind("<Configure>", self._on_gallery_resize)
        # barretta sotto la galleria: trascinandola la si allunga
        grip = tk.Frame(gf, height=7, bg=SASH_BG, cursor="sb_v_double_arrow")
        grip.pack(fill=tk.X, pady=(2, 0))
        self._gallery_drag = None
        for seq, phase in (("<Button-1>", "press"), ("<B1-Motion>", "motion"), ("<ButtonRelease-1>", "release")):
            grip.bind(seq, lambda e, p=phase: self._on_gallery_grip(e, p))
        rows = [ttk.Frame(gf) for _ in range(3)]
        for k, row in enumerate(rows):
            row.pack(fill=tk.X, pady=(6 if k == 0 else 4, 0))
        for row, buttons in zip(rows, ((("Aggiungi…", self.cmd_add_stamp), ("Crea firma…", self.cmd_create_signature)),
                                       (("Modifica…", self.cmd_edit_stamp), ("Rinomina…", self.cmd_rename_stamp)),
                                       (("↺ Ruota", lambda: self.rotate_stamp(-90)),
                                        ("Ruota ↻", lambda: self.rotate_stamp(90)),
                                        ("Elimina", self.cmd_delete_stamp)))):
            for k, (text, cmd) in enumerate(buttons):
                ttk.Button(row, text=text, command=cmd).pack(side=tk.LEFT, expand=True, fill=tk.X,
                                                             padx=(4 if k else 0, 0))

        # tre colonne: impilate con il pannello a lato, affiancate con il pannello in alto/in basso
        lf = ttk.LabelFrame(parent, text=" Timbro selezionato ", padding=8)
        c1, c2, c3 = ttk.Frame(lf), ttk.Frame(lf), ttk.Frame(lf)
        self.stamp_preview = tk.Canvas(c1, width=PREVIEW_W, height=PREVIEW_H, bg="white",
                                       highlightthickness=1, highlightbackground="#b8b8c0")
        self.stamp_preview.pack()

        self.remove_white_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(c2, text="Rimuovi sfondo bianco", variable=self.remove_white_var,
                        command=self._on_white_settings).pack(anchor="w", pady=(8, 0))
        tol = ttk.Frame(c2)
        tol.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(tol, text="Tolleranza:").pack(side=tk.LEFT)
        self.tolerance_scale = ttk.Scale(tol, from_=0, to=100,
                                         command=lambda _v: self._on_white_settings(delay=150))
        self.tolerance_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

        size = ttk.Frame(c2)
        size.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(size, text="Larghezza (mm):").pack(side=tk.LEFT)
        self.stamp_w_var = tk.StringVar()
        spin = ttk.Spinbox(size, from_=3, to=600, increment=1, width=6, textvariable=self.stamp_w_var)
        spin.pack(side=tk.RIGHT)
        spin.bind("<Return>", lambda e: self.view.focus_set())
        self.stamp_scale = ttk.Scale(c2, from_=5, to=150, command=self._on_scale)
        self.stamp_scale.pack(fill=tk.X, pady=(4, 0))
        self.stamp_info = ttk.Label(c2, text="", foreground="#555")
        self.stamp_info.pack(anchor="w", pady=(2, 0))
        self.stamp_w_var.trace_add("write", self._on_stamp_width)

        ttk.Label(c3, text="Posizione rispetto al contenuto:").pack(anchor="w", pady=(10, 0))
        self.overlay_var = tk.BooleanVar(value=True)
        for value, text in ((True, "Primo piano (sopra testo e disegni)"),
                            (False, "Secondo piano (sotto testo e linee)")):
            ttk.Radiobutton(c3, text=text, variable=self.overlay_var, value=value,
                            command=self._on_overlay).pack(anchor="w")
        self.overlay_note = ttk.Label(c3, wraplength=PREVIEW_W, foreground="#a15c00", justify=tk.LEFT)
        self.overlay_note.pack(anchor="w")
        self._paint_overlay_note()

        self.mode_btn = tk.Button(c3, command=self.toggle_stamp_mode, font=("Segoe UI", 10, "bold"),
                                  relief=tk.RAISED, bd=1, padx=6, pady=8, cursor="hand2")
        self._btn_colors = (self.mode_btn.cget("bg"), self.mode_btn.cget("fg"))
        self.mode_btn.pack(fill=tk.X, pady=(12, 0))
        self._paint_mode_btn()

        self.repeat_btn = ttk.Button(c3, text="Ripeti su tutte le pagine", command=self.repeat_last_stamp)
        self.repeat_btn.pack(fill=tk.X, pady=(6, 0))

        help_label = ttk.Label(parent, wraplength=PREVIEW_W + 10, justify=tk.LEFT, foreground="#555", text=(
            "Clic su un timbro per sceglierlo (o tasti 1–9), doppio clic per modificarlo. "
            "Ogni timbro ricorda le sue impostazioni.\n\n"
            "In modalità timbro clicca sul foglio per apporlo; cambia pagina e clicca di nuovo dove vuoi. "
            "Esc esce dalla modalità.\n\n"
            "Fuori dalla modalità timbro trascina un timbro già apposto per spostarlo, doppio clic per "
            "modificarlo; tasto destro per ruotarlo o eliminarlo (anche Canc). Ctrl+Z annulla.\n\n"
            "Le modifiche al PDF vanno salvate con «Salva…»."))
        self._stamp_sections = (gf, lf, help_label)
        self._stamp_columns = (c1, c2, c3)
        self._orient_stamps(False)

    def _build_menu(self):
        m = tk.Menu(self.root)
        f = tk.Menu(m, tearoff=False)
        f.add_command(label="Apri… (anche più file, ognuno nella sua scheda)", accelerator="Ctrl+O",
                      command=self.cmd_open)
        f.add_command(label="Nuovo documento vuoto", accelerator="Ctrl+N", command=self.cmd_new)
        f.add_command(label="Aggiungi PDF/immagini in fondo…", command=lambda: self.cmd_add(at_end=True))
        f.add_command(label="Inserisci PDF/immagini dopo la pagina corrente…",
                      command=lambda: self.cmd_add(at_end=False))
        f.add_separator()
        f.add_command(label="Salva con nome…", accelerator="Ctrl+S", command=self.cmd_save)
        f.add_command(label="Estrai pagine selezionate…", command=self.cmd_extract)
        f.add_command(label="Esporta in Word…", command=self.cmd_export_word)
        f.add_separator()
        f.add_command(label="Unisci tutte le schede in un nuovo documento", command=self.cmd_merge_tabs)
        f.add_command(label="Dividi il documento in più file…", command=self.cmd_split)
        f.add_separator()
        f.add_command(label=f"Usa {APP_NAME} per aprire i PDF (programma predefinito)…", command=self.cmd_default_app)
        f.add_command(label=f"Togli {APP_NAME} dai programmi per i PDF…", command=self.cmd_remove_default_app)
        f.add_separator()
        f.add_command(label="Chiudi scheda", accelerator="Ctrl+W", command=self.close_tab)
        f.add_command(label="Esci", command=self.quit)
        m.add_cascade(label="File", menu=f)

        e = tk.Menu(m, tearoff=False)
        e.add_command(label="Annulla", accelerator="Ctrl+Z", command=self.undo)
        e.add_command(label="Ripristina", accelerator="Ctrl+Y", command=self.redo)
        e.add_separator()
        e.add_command(label="Ruota a destra", accelerator="Ctrl+R", command=lambda: self.rotate(90))
        e.add_command(label="Ruota a sinistra", accelerator="Ctrl+Maiusc+R", command=lambda: self.rotate(-90))
        e.add_command(label="Ruota tutte le pagine a destra", command=lambda: self.rotate(90, all_pages=True))
        e.add_command(label="Ruota tutte le pagine a sinistra", command=lambda: self.rotate(-90, all_pages=True))
        e.add_separator()
        e.add_command(label="Sposta su", command=lambda: self.move_pages(-1))
        e.add_command(label="Sposta giù", command=lambda: self.move_pages(1))
        e.add_command(label="Elimina pagine selezionate", accelerator="Canc", command=self.delete_pages)
        e.add_command(label="Seleziona tutte le pagine", accelerator="Ctrl+A", command=self.select_all)
        m.add_cascade(label="Pagine", menu=e)

        v = tk.Menu(m, tearoff=False)
        v.add_command(label="Zoom avanti", accelerator="Ctrl++", command=lambda: self._zoom_at(1.25))
        v.add_command(label="Zoom indietro", accelerator="Ctrl+-", command=lambda: self._zoom_at(1 / 1.25))
        v.add_command(label="Adatta alla larghezza", command=lambda: self.set_zoom_mode("width"))
        v.add_command(label="Pagina intera", accelerator="Ctrl+0", command=lambda: self.set_zoom_mode("page"))
        v.add_command(label="Dimensione reale (100%)", command=self.set_real_size)
        v.add_command(label="Vista d'insieme delle miniature di tutti i PDF", accelerator="F4",
                      command=self.toggle_overview)
        v.add_separator()
        v.add_command(label="Ripristina disposizione di barre e pannelli", command=self.reset_layout)
        v.add_separator()
        v.add_command(label="Scheda successiva", accelerator="Ctrl+Tab", command=lambda: self.next_tab(1))
        v.add_command(label="Scheda precedente", accelerator="Ctrl+Maiusc+Tab", command=lambda: self.next_tab(-1))
        m.add_cascade(label="Visualizza", menu=v)

        x = tk.Menu(m, tearoff=False)
        x.add_command(label="Modifica testo", command=lambda: self.set_tool("edit_text"))
        x.add_command(label="Aggiungi testo", command=lambda: self.set_tool("add_text"))
        x.add_command(label="Bianchetto (copri con il bianco)", command=lambda: self.set_tool("whiteout"))
        x.add_command(label="Oscura (copri con il nero)", command=lambda: self.set_tool("blackout"))
        x.add_separator()
        x.add_command(label="Riconosci testo (OCR)…", command=self.cmd_ocr)
        x.add_command(label="Cerca", accelerator="Ctrl+F", command=self.focus_find)
        m.add_cascade(label="Testo", menu=x)

        t = tk.Menu(m, tearoff=False)
        t.add_command(label="Aggiungi timbro da immagine…", command=self.cmd_add_stamp)
        t.add_command(label="Crea firma (disegnata o scritta)…", command=self.cmd_create_signature)
        t.add_command(label="Modifica timbro selezionato (ruota, matita, gomma, testo)…",
                      command=self.cmd_edit_stamp)
        t.add_command(label="Ruota timbro selezionato a destra", command=lambda: self.rotate_stamp(90))
        t.add_command(label="Ruota timbro selezionato a sinistra", command=lambda: self.rotate_stamp(-90))
        t.add_command(label="Rinomina timbro selezionato…", command=self.cmd_rename_stamp)
        t.add_command(label="Elimina timbro selezionato", command=self.cmd_delete_stamp)
        t.add_separator()
        t.add_command(label="Attiva/disattiva modalità timbro", accelerator="T", command=self.toggle_stamp_mode)
        t.add_command(label="Ripeti timbro su tutte/selezionate", command=self.repeat_last_stamp)
        t.add_separator()
        t.add_command(label="Archivio dei timbri (posizione, altri PC, password)…", command=self.cmd_archive)
        m.add_cascade(label="Timbro", menu=t)

        h = tk.Menu(m, tearoff=False)
        h.add_command(label="Guida rapida", accelerator="F1", command=self.show_help)
        h.add_command(label="Guida illustrata (PDF)", command=self.open_guide)
        h.add_separator()
        h.add_command(label="Informazioni, licenza e donazioni…", command=self.show_about)
        m.add_cascade(label="?", menu=h)
        self.root.configure(menu=m)

    def _bind_keys(self):
        typing = (tk.Entry, ttk.Entry, tk.Spinbox, ttk.Spinbox, ttk.Scale)

        def bind(seq, fn, skip_when_typing=False):
            def handler(e):
                # solo nella finestra principale, non nelle finestre di dialogo
                if not hasattr(e.widget, "winfo_toplevel") or e.widget.winfo_toplevel() is not self.root:
                    return None
                if skip_when_typing and isinstance(e.widget, typing):
                    return None
                fn()
                return "break"
            self.root.bind_all(seq, handler)

        bind("<Control-o>", self.cmd_open)
        bind("<Control-n>", self.cmd_new)
        bind("<Control-w>", self.close_tab)
        bind("<Control-f>", self.focus_find)
        bind("<Control-s>", self.cmd_save)
        bind("<Control-z>", self.undo, True)
        bind("<Control-y>", self.redo, True)
        bind("<Control-r>", lambda: self.rotate(90))
        bind("<Control-R>", lambda: self.rotate(-90))
        bind("<Control-a>", self.select_all, True)
        bind("<Prior>", self.prev_page)
        bind("<Next>", self.next_page)
        bind("<Left>", self.prev_page, True)
        bind("<Right>", self.next_page, True)
        bind("<Home>", lambda: self.go_to(0), True)
        bind("<End>", lambda: self.go_to(len(self.doc) - 1 if self.doc else 0), True)
        bind("<Delete>", self._on_delete_key, True)
        bind("<t>", self.toggle_stamp_mode, True)
        bind("<T>", self.toggle_stamp_mode, True)
        bind("<Escape>", lambda: self.toggle_overview(False) if self.overview else self.set_tool(None))
        bind("<F4>", self.toggle_overview)

        def tab_key(e):
            if hasattr(e.widget, "winfo_toplevel") and e.widget.winfo_toplevel() is self.root:
                self.next_tab(-1 if e.state & 0x0001 else 1)
                return "break"
            return None
        self.root.bind_all("<Control-Tab>", tab_key)
        for k in range(1, 10):
            bind(f"<Key-{k}>", lambda k=k: self.select_stamp(k - 1), True)
        bind("<F1>", self.show_help)
        for seq in ("<Control-plus>", "<Control-equal>", "<Control-KP_Add>"):
            bind(seq, lambda: self._zoom_at(1.25))
        for seq in ("<Control-minus>", "<Control-KP_Subtract>"):
            bind(seq, lambda: self._zoom_at(1 / 1.25))
        bind("<Control-0>", lambda: self.set_zoom_mode("page"))

    def status(self, text):
        self.status_var.set(text)

    def _dir(self, kind, fallback=None):
        """Cartella da proporre: l'ultima usata per quel tipo di operazione ("open", "add", "save", "split",
        "word", "image"), altrimenti fallback (es. la cartella del documento), altrimenti l'ultima usata."""
        for folder in ((self.cfg["dirs"] or {}).get(kind), fallback, self.cfg["last_dir"]):
            if folder and Path(folder).is_dir():
                return str(folder)
        return str(Path.home())

    def _remember_dir(self, kind, folder):
        if folder:
            self.cfg.data["last_dir"] = str(folder)
            self.cfg["dirs"] = dict(self.cfg["dirs"] or {}, **{kind: str(folder)})

    # ------------------------------------------------------------ schede
    def _refresh_tabbar(self):
        """Disegna le schede; le ricrea solo se l'elenco dei PDF aperti è cambiato."""
        bar = self.tabbar
        if [f.doc_tab for f in self._tab_frames] != self.tabs:
            for w in self._tab_frames + [self._plus]:
                if w is not None:
                    w.destroy()
            self._tab_frames = []
            for tab in self.tabs:
                frame = tk.Frame(bar, padx=10, pady=4, cursor="hand2")
                name = tk.Label(frame, font=("Segoe UI", 9), padx=0, pady=0)
                close = tk.Label(frame, text="✕", font=("Segoe UI", 8), padx=3, fg="#666", cursor="hand2")
                name.pack(side=tk.LEFT)
                close.pack(side=tk.LEFT, padx=(8, 0))
                frame.pack(side=tk.LEFT, padx=(0, 1), pady=(4, 0))
                for w in (frame, name, close):
                    w.doc_tab = tab
                    w.bind("<Button-1>", lambda e, t=tab: self._on_tab_press(e, t))
                    w.bind("<B1-Motion>", self._on_tab_motion)
                    w.bind("<ButtonRelease-1>", self._on_tab_release)
                    w.bind("<Button-2>", lambda e, t=tab: self.close_tab(t))
                    w.bind("<Button-3>", lambda e, t=tab: self._on_tab_menu(e, t))
                    w.bind("<Enter>", lambda e, t=tab: self.status(t.path or t.name))
                close.bind("<Button-1>", lambda e, t=tab: self.close_tab(t))
                close.bind("<Enter>", lambda e: e.widget.configure(fg="#c62828"), add=True)
                close.bind("<Leave>", lambda e: e.widget.configure(fg="#666"))
                frame.parts = (name, close)
                self._tab_frames.append(frame)
            self._plus = tk.Label(bar, text=" + ", font=("Segoe UI", 12, "bold"), bg=TABBAR_BG, fg="#444",
                                  cursor="hand2")
            self._plus.pack(side=tk.LEFT, padx=4)
            self._plus.bind("<Button-1>", lambda e: self.cmd_open())
            self._plus.bind("<Enter>", lambda e: self.status(
                "Apri un PDF in una nuova scheda (doppio clic sulla barra: documento vuoto)."))
        char_px = 7 * self.px_per_pt * 72 / 96
        room = max(bar.winfo_width() - 60, 300) / max(len(self.tabs), 1)
        max_chars = min(max(int(room / char_px) - 5, 8), 34)
        for frame in self._tab_frames:
            tab = frame.doc_tab
            active = tab is self.tab
            bg = TAB_DROP_BG if tab is self._drop_tab else TAB_ACTIVE_BG if active else TAB_BG
            text = tab.name if len(tab.name) <= max_chars else tab.name[:max_chars - 1] + "…"
            name, close = frame.parts
            frame.configure(bg=bg)
            name.configure(text=("● " if tab.dirty else "") + text, bg=bg,
                           font=("Segoe UI", 9, "bold") if active else ("Segoe UI", 9))
            close.configure(bg=bg)
        if self._plus:
            self._plus.configure(bg=TAB_DROP_BG if self._drop_tab == "new" else TABBAR_BG)
        self._schedule_overview()

    def _on_tab_menu(self, event, tab):
        self.switch_tab(tab)
        menu = tk.Menu(self.root, tearoff=False)
        menu.add_command(label="Salva con nome…", command=self.cmd_save)
        menu.add_command(label="Unisci tutte le schede in un nuovo documento", command=self.cmd_merge_tabs)
        menu.add_command(label="Dividi in più file…", command=self.cmd_split)
        menu.add_separator()
        menu.add_command(label="Chiudi", command=lambda: self.close_tab(tab))
        menu.add_command(label="Chiudi le altre schede",
                         command=lambda: [self.close_tab(t) for t in list(self.tabs) if t is not tab])
        menu.tk_popup(event.x_root, event.y_root)

    # ------------------------------------------------------------ vista d'insieme
    def toggle_overview(self, on=None):
        """Mostra al posto della pagina le miniature di tutti i PDF aperti, una colonna per PDF."""
        on = (not self.overview) if on is None else on
        if on and not self.tabs:
            on = False
            self.status("Apri almeno un PDF per la vista d'insieme.")
        self.overview_var.set(on)
        if on == self.overview:
            return
        self.overview = on
        if on:
            self.set_tool(None)
            self.view_area.pack_forget()
            self.overview_area.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            self._layout_overview()
            self.status("Vista d'insieme: trascina le pagine da una colonna all'altra per copiarle, nella stessa "
                        "colonna per spostarle. Doppio clic su una pagina per aprirla. Esc o F4 per tornare.")
        else:
            self.overview_area.pack_forget()
            self.view_area.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
            self._render_page()

    def _schedule_overview(self):
        if self.overview and self._ov_job is None:
            self._ov_job = self.root.after(30, self._layout_overview)

    def _layout_overview(self):
        self._ov_job = None
        if not self.overview:
            return
        if not self.tabs:
            self.toggle_overview(False)
            return
        c = self.ov
        c.delete("all")
        self.ov_layout = []
        tw, th = THUMB_W, THUMB_H
        colw = tw + 46
        x, bottom = 12, 0
        for tab in self.tabs:
            active = tab is self.tab
            boxes, y = [], 52
            for i in range(len(tab.doc)):
                r = tab.doc[i].rect
                s = min(tw / r.width, th / r.height)
                w, h = r.width * s, r.height * s
                x0 = x + (colw - w) / 2
                selected = i in tab.selection
                c.create_rectangle(x0 - 6, y - 6, x0 + w + 6, y + h + 20, width=2, outline=(
                    "#1a73e8" if i == tab.current and active else ""), fill="#bcd4f6" if selected and active else "")
                c.create_rectangle(x0, y, x0 + w, y + h, fill="white", outline="#a0a0a8")
                c.create_text(x0 + w / 2, y + h + 8, text=str(i + 1), fill="#eee", font=("Segoe UI", 8))
                if i in tab.ov_photos:
                    c.create_image(x0, y, anchor="nw", image=tab.ov_photos[i])
                boxes.append((x0, y, x0 + w, y + h, s))
                y += h + 30
            if not len(tab.doc):
                c.create_text(x + colw / 2, 90, text="vuoto:\ntrascina qui\nle pagine", fill="#ddd",
                              justify=tk.CENTER, font=("Segoe UI", 9))
                y = 140
            self.ov_layout.append((tab, x, x + colw, boxes))
            bottom = max(bottom, y)
            x += colw + 10
        # intestazioni e fondi delle colonne, sotto alle miniature
        for tab, x0, x1, _boxes in self.ov_layout:
            active = tab is self.tab
            bg = c.create_rectangle(x0, 4, x1, bottom + 4, outline="", fill="#6f7884" if active else "#686970")
            c.tag_lower(bg)
            name = tab.name if len(tab.name) <= 22 else tab.name[:21] + "…"
            c.create_text((x0 + x1) / 2, 20, text=("● " if tab.dirty else "") + name, fill="white",
                          font=("Segoe UI", 9, "bold" if active else "normal"))
            c.create_text((x0 + x1) / 2, 37, text=f"{len(tab.doc)} pagine", fill="#ddd", font=("Segoe UI", 8))
        c.configure(scrollregion=(0, 0, x, bottom + 10))
        self._schedule_overview_render()

    def _schedule_overview_render(self):
        if self.overview and self._ov_render_job is None:
            self._ov_render_job = self.root.after(20, self._render_overview_step)

    def _render_overview_step(self):
        """Disegna una miniatura visibile alla volta, come per il pannello delle pagine."""
        self._ov_render_job = None
        if not self.overview:
            return
        c = self.ov
        left, right = c.canvasx(0) - 200, c.canvasx(c.winfo_width()) + 200
        top, bottom = c.canvasy(0) - 300, c.canvasy(c.winfo_height()) + 300
        for tab, x0, x1, boxes in self.ov_layout:
            if x1 < left or x0 > right:
                continue
            for i, (bx0, by0, _bx1, by1, s) in enumerate(boxes):
                if by1 < top or i in tab.ov_photos:
                    continue
                if by0 > bottom:
                    break
                pix = tab.doc[i].get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
                tab.ov_photos[i] = photo_from_pixmap(pix)
                c.create_image(bx0, by0, anchor="nw", image=tab.ov_photos[i])
                self._ov_render_job = self.root.after(1, self._render_overview_step)
                return

    def _ov_hit(self, event):
        """(colonna, pagina o None, posizione di inserimento) sotto il mouse; colonna None fuori dalle colonne."""
        x, y = self.ov.canvasx(event.x), self.ov.canvasy(event.y)
        for col, (_tab, x0, x1, boxes) in enumerate(self.ov_layout):
            if x0 <= x <= x1:
                page = next((i for i, b in enumerate(boxes) if b[1] - 6 <= y <= b[3] + 20), None)
                insert = next((i for i, b in enumerate(boxes) if y < (b[1] + b[3]) / 2), len(boxes))
                return col, page, insert
        return None, None, None

    def _ov_press(self, event, mode):
        self._ov_drag = None
        col, page, _ = self._ov_hit(event)
        if col is None or page is None:
            return
        tab = self.ov_layout[col][0]
        self.switch_tab(tab)
        collapse = False
        if mode == "toggle":
            self.selection ^= {page}
            self.anchor = page
        elif mode == "range":
            a, b = sorted((self.anchor, page))
            self.selection = set(range(a, b + 1))
        elif page in self.selection and len(self.selection) > 1:
            collapse = True
        else:
            self.selection = {page}
            self.anchor = page
        self.current = page
        self._paint_selection()
        self._update_controls()
        self._ov_drag = {"tab": tab, "index": page, "x": event.x_root, "y": event.y_root, "active": False,
                         "collapse": collapse, "target": None}

    def _ov_motion(self, event):
        d = self._ov_drag
        if not d:
            return
        if not d["active"]:
            if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 8:
                return
            d["active"] = True
            d["pages"] = sorted(self.selection) if d["index"] in self.selection else [d["index"]]
            n = len(d["pages"])
            tip = d["tip"] = tk.Toplevel(self.root)
            tip.overrideredirect(True)
            tip.attributes("-topmost", True)
            tk.Label(tip, text=f"Pagina {d['pages'][0] + 1}" if n == 1 else f"{n} pagine", bg="#1a73e8", fg="white",
                     font=("Segoe UI", 9, "bold"), padx=8, pady=3).pack()
        d["tip"].geometry(f"+{event.x_root + 16}+{event.y_root + 14}")
        c = self.ov
        c.delete("dropline")
        col, _page, insert = self._ov_hit(event)
        d["target"] = None if col is None else (col, insert)
        if col is not None:
            dst, x0, x1, boxes = self.ov_layout[col]
            y = 52 if not boxes else boxes[insert][1] - 15 if insert < len(boxes) else boxes[-1][3] + 22
            c.create_line(x0 + 8, y, x1 - 8, y, fill="#ffd54f", width=4, tags="dropline")
            self.status(f"Rilascia per {'spostarle' if dst is d['tab'] else 'copiarle'} qui in «{dst.name}».")
        if event.y < 30 or event.y > c.winfo_height() - 30:   # vicino al bordo: scorre
            c.yview_scroll(-1 if event.y < 30 else 1, "units")

    def _ov_release(self, _event):
        d, self._ov_drag = self._ov_drag, None
        if not d:
            return
        if not d["active"]:
            if d["collapse"] and d["tab"] is self.tab:
                self.selection = {d["index"]}
                self.anchor = d["index"]
                self._update_controls()
            return
        d["tip"].destroy()
        self.ov.delete("dropline")
        if not d["target"] or d["tab"] not in self.tabs:
            self.status("Trascinamento annullato.")
            return
        col, insert = d["target"]
        dst = self.ov_layout[col][0]
        if dst is d["tab"]:
            self.switch_tab(dst)
            self._move_pages_to(d["pages"], insert)
        else:
            self._copy_pages(d["tab"], d["pages"], dst, insert)
        self._schedule_overview()

    def _ov_double(self, event):
        col, page, _ = self._ov_hit(event)
        if col is None or page is None:
            return
        self._ov_drag = None
        self.switch_tab(self.ov_layout[col][0])
        self.toggle_overview(False)
        self.go_to(page)

    def _ov_wheel(self, event):
        step = -3 if event.delta > 0 else 3
        (self.ov.xview_scroll if event.state & 0x0001 else self.ov.yview_scroll)(step, "units")
        return "break"

    # trascinando una scheda si cambia l'ordine (è l'ordine usato da «Unisci schede»)
    def _on_tab_press(self, event, tab):
        self._tab_drag = {"tab": tab, "x": event.x_root, "active": False, "index": None}
        self.switch_tab(tab)

    def _on_tab_motion(self, event):
        d = self._tab_drag
        if not d or (not d["active"] and abs(event.x_root - d["x"]) < 12):
            return
        d["active"] = True
        others = [f for f in self._tab_frames if f.doc_tab is not d["tab"]]
        k = sum(1 for f in others if f.winfo_rootx() + f.winfo_width() / 2 < event.x_root)
        d["index"] = k
        bar_x = self.tabbar.winfo_rootx()
        x = (others[k].winfo_rootx() if k < len(others) else
             others[-1].winfo_rootx() + others[-1].winfo_width() if others else bar_x) - bar_x
        self._tab_marker.place(x=max(x - 2, 0), y=3, width=3, relheight=0.8)
        self._tab_marker.lift()
        self.status("Rilascia per spostare la scheda qui.")

    def _on_tab_release(self, _event):
        d, self._tab_drag = self._tab_drag, None
        self._tab_marker.place_forget()
        if not d or not d["active"] or d["tab"] not in self.tabs:
            return
        tabs = [t for t in self.tabs if t is not d["tab"]]
        tabs.insert(d["index"], d["tab"])
        if tabs != self.tabs:
            self.tabs[:] = tabs
            self._refresh_tabbar()
            self.status("Ordine delle schede cambiato: «Unisci schede» le accoda in quest'ordine.")

    def _store_view(self):
        if self.tab is not self.no_tab:
            self.tab.view_pos = (self.view.xview()[0], self.view.yview()[0])
            self.tab.thumbs_pos = self.thumbs.yview()[0]

    def switch_tab(self, tab, force=False):
        if tab is self.tab and not force:
            return
        self._store_view()
        self.tab = tab
        self.sel_stamp = None
        self._drag = None
        self._layout_thumbs()
        self._render_page()
        self._update_controls()
        self.view.xview_moveto(tab.view_pos[0])
        self.view.yview_moveto(tab.view_pos[1])
        self.thumbs.yview_moveto(tab.thumbs_pos)
        self._schedule_thumbs()
        self._redraw_ghost()

    def next_tab(self, step):
        if len(self.tabs) > 1 and self.tab in self.tabs:
            self.switch_tab(self.tabs[(self.tabs.index(self.tab) + step) % len(self.tabs)])

    def _add_tab(self, tab, switch=True):
        # un documento nuovo ancora vuoto e intatto lascia il posto al PDF aperto
        old = self.tab
        unused = old.doc is not None and old.path is None and not old.dirty and not len(old.doc)
        self.tabs.append(tab)
        if not switch:
            self._refresh_tabbar()
            return
        if unused:
            self.tabs.remove(old)
        self.switch_tab(tab)

    def close_tab(self, tab=None):
        """Chiude la scheda (chiedendo se salvare le modifiche). False se l'utente annulla."""
        tab = tab or self.tab
        if tab not in self.tabs:
            return True
        if tab.dirty and len(tab.doc):
            self.switch_tab(tab)
            answer = messagebox.askyesnocancel(APP_NAME, f"«{tab.name}» ha modifiche non salvate.\n"
                                               "Vuoi salvarle prima di chiuderlo?", parent=self.root)
            if answer is None or (answer and not self.cmd_save()):
                return False
        index = self.tabs.index(tab)
        self.tabs.remove(tab)
        if tab is self.tab:
            self.tab = self.no_tab
            self.switch_tab(self.tabs[min(index, len(self.tabs) - 1)] if self.tabs else self.no_tab, force=True)
        else:
            self._refresh_tabbar()
        tab.doc.close()
        return True

    # ------------------------------------------------------------ documento
    def cmd_open(self):
        paths = filedialog.askopenfilenames(parent=self.root, title="Apri PDF (anche più di uno)",
                                            initialdir=self._dir("open"), filetypes=PDF_TYPES)
        for path in paths:
            self.open_file(path)
        if paths:
            self._remember_dir("open", Path(paths[0]).parent)

    def cmd_new(self):
        self._add_tab(DocTab(pymupdf.open()))
        self.status("Nuovo documento vuoto: trascina qui le pagine dalle miniature di un'altra scheda "
                    "(fermati un attimo sulla scheda), oppure usa «Aggiungi…».")

    def open_file(self, path):
        """Apre il PDF in una nuova scheda (se è già aperto, passa alla sua scheda)."""
        key = os.path.normcase(os.path.abspath(path))
        for tab in self.tabs:
            if tab.path and os.path.normcase(os.path.abspath(tab.path)) == key:
                self.switch_tab(tab)
                self.status(f"«{tab.name}» è già aperto.")
                return True
        try:
            doc = load_as_pdf(path, self.root)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile aprire il file:\n{path}\n\n{exc}", parent=self.root)
            return False
        if doc is None:
            return False
        self._add_tab(DocTab(doc, path))
        self.cfg["last_dir"] = str(Path(path).parent)
        self.status(f"Aperto: {path}")
        return True

    def cmd_add(self, at_end=True):
        paths = filedialog.askopenfilenames(parent=self.root, title="Aggiungi PDF o immagini",
                                            initialdir=self._dir("add", self._dir("open")), filetypes=ADD_TYPES)
        if not paths:
            return
        self._remember_dir("add", Path(paths[0]).parent)
        paths = sorted(paths, key=lambda p: Path(p).name.lower())
        if self.doc is None:
            if not self.open_file(paths.pop(0)) or not paths:
                return
        sources = []
        for p in paths:
            try:
                src = load_as_pdf(p, self.root)
            except Exception as exc:
                messagebox.showerror(APP_NAME, f"Impossibile aprire «{Path(p).name}»:\n{exc}", parent=self.root)
                continue
            if src is not None:
                sources.append(src)
        if not sources:
            return
        self._snapshot()
        pos = first = len(self.doc) if at_end else self.current + 1
        for src in sources:
            self.doc.insert_pdf(src, start_at=pos)
            pos += len(src)
            src.close()
        self.go_to(first, structure=True)
        self.selection = set(range(first, pos))
        self._paint_selection()
        self._update_controls()
        self.status(f"Aggiunte {pos - first} pagine da {len(sources)} file.")

    def cmd_save(self):
        if self.doc is not None and not len(self.doc):
            messagebox.showinfo(APP_NAME, "Il documento è vuoto: trascinaci prima qualche pagina "
                                "da un'altra scheda, oppure usa «Aggiungi…».", parent=self.root)
            return False
        if not self.doc:
            return False
        # si propone la cartella dell'ultimo salvataggio (la prima volta quella del documento)
        name = Path(self.path).stem + ".pdf" if self.path else f"{self.tab.name}.pdf"
        folder = self._dir("save", str(Path(self.path).parent) if self.path else None)
        path = filedialog.asksaveasfilename(parent=self.root, title=f"Salva «{self.tab.name}»", initialdir=folder,
                                            initialfile=name, defaultextension=".pdf", filetypes=PDF_TYPES)
        if not path:
            return False
        if self.tab.fonts_added:
            try:
                self.doc.subset_fonts()   # dei font di sistema usati si tengono solo i caratteri scritti
            except Exception:
                pass
        data = self.doc.tobytes(garbage=3, deflate=True)
        if not self._write_file(path, data):
            return False
        # ricarica dai byte salvati: stato pulito e coerente con il file
        self.doc = pymupdf.open(stream=data, filetype="pdf")
        self.stamp_xrefs.clear()
        self.tab.fonts_added = False
        self.path = path
        self.dirty = False
        self._remember_dir("save", Path(path).parent)
        self._doc_changed()
        self.status(f"Salvato: {path}")
        return True

    def _write_file(self, path, data):
        tmp = Path(path + ".tmp")
        try:
            tmp.write_bytes(data)
            os.replace(tmp, path)
            return True
        except OSError as exc:
            tmp.unlink(missing_ok=True)
            messagebox.showerror(APP_NAME, f"Impossibile salvare il file:\n{path}\n\n{exc}\n\n"
                                 "Forse è aperto in un altro programma?", parent=self.root)
            return False

    def cmd_extract(self):
        if not self.doc:
            return
        pages = sorted(self._targets())
        stem = Path(self.path).stem if self.path else "documento"
        path = filedialog.asksaveasfilename(parent=self.root, title="Salva pagine estratte",
                                            initialdir=self._dir("save", str(Path(self.path).parent) if self.path
                                                                 else None),
                                            initialfile=f"{stem}_estratto.pdf",
                                            defaultextension=".pdf", filetypes=PDF_TYPES)
        if not path:
            return
        out = pymupdf.open()
        for i in pages:
            out.insert_pdf(self.doc, from_page=i, to_page=i)
        if self._write_file(path, out.tobytes(garbage=3, deflate=True)):
            self._remember_dir("save", Path(path).parent)
            self.status(f"Estratte {len(pages)} pagine in {path}")

    def cmd_merge_tabs(self):
        """Accoda in un nuovo documento tutti i PDF aperti, nell'ordine delle schede."""
        tabs = [t for t in self.tabs if len(t.doc)]
        if len(tabs) < 2:
            messagebox.showinfo(APP_NAME, "Apri almeno due PDF (ognuno nella sua scheda) per unirli.",
                                parent=self.root)
            return
        dialog = MergeDialog(self.root, tabs)   # quali unire e in che ordine
        self.root.wait_window(dialog.top)
        if not dialog.result:
            return
        tabs = dialog.result
        doc = pymupdf.open()
        for tab in tabs:
            doc.insert_pdf(tab.doc)
        merged = DocTab(doc, label="Unione")
        merged.dirty = True
        merged.fonts_added = any(t.fonts_added for t in tabs)
        self._add_tab(merged)
        self.status(f"Uniti {len(tabs)} documenti ({len(doc)} pagine) in «{merged.name}», nell'ordine delle "
                    "schede: salvalo con «Salva…».")

    def cmd_split(self):
        """Salva il documento in più file da 1, 2 o N pagine, con nome e numero progressivo."""
        if not self.doc:
            return
        stem = Path(self.path).stem if self.path else self.tab.name
        folder = self._dir("split", str(Path(self.path).parent) if self.path else None)
        dialog = SplitDialog(self.root, len(self.doc), stem, folder, per=self.cfg["split_per"] or 1)
        self.root.wait_window(dialog.top)
        if not dialog.result:
            return
        per, name, start, folder = dialog.result
        n = len(self.doc)
        chunks = [(a, min(a + per, n) - 1) for a in range(0, n, per)]
        paths = [Path(folder) / file for file in split_names(name, start, len(chunks))]
        existing = [p for p in paths if p.exists()]
        if existing and not messagebox.askyesno(
                APP_NAME, f"{len(existing)} file con questi nomi esistono già (es. «{existing[0].name}»).\n"
                          "Vuoi sovrascriverli?", parent=self.root):
            return
        try:
            Path(folder).mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Impossibile usare la cartella:\n{folder}\n\n{exc}", parent=self.root)
            return
        self.cfg["split_per"] = per
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            for (a, b), path in zip(chunks, paths):
                out = pymupdf.open()
                out.insert_pdf(self.doc, from_page=a, to_page=b)
                if self.tab.fonts_added:
                    try:
                        out.subset_fonts()
                    except Exception:
                        pass
                if not self._write_file(str(path), out.tobytes(garbage=3, deflate=True)):
                    return
        finally:
            self.root.configure(cursor="")
        self._remember_dir("split", folder)
        self.status(f"Creati {len(paths)} file in {folder}: da «{paths[0].name}» a «{paths[-1].name}».")
        if messagebox.askyesno(APP_NAME, f"Creati {len(paths)} file in:\n{folder}\n\nVuoi aprire la cartella?",
                               parent=self.root):
            os.startfile(folder)

    def quit(self):
        for tab in [t for t in self.tabs if t.dirty]:
            if not self.close_tab(tab):
                return
        self.library.save()  # impostazioni dei timbri ancora in attesa di salvataggio
        self._save_layout()
        self._save_window()
        if self.instance:
            self.instance.close()
        self.root.destroy()

    # --------------------------------------------------------- annulla
    def _state(self, tab=None):
        tab = tab or self.tab
        return tab.doc.tobytes(), tab.current, set(tab.selection)

    def _snapshot(self, tab=None):
        """Da chiamare prima di ogni modifica del documento (della scheda attiva, se tab è None)."""
        tab = tab or self.tab
        stack = tab.undo_stack
        stack.append(self._state(tab))
        while len(stack) > 1 and (len(stack) > UNDO_LIMIT or sum(len(s[0]) for s in stack) > UNDO_MAX_BYTES):
            stack.pop(0)
        tab.redo_stack.clear()
        tab.dirty = True

    def _restore(self, state):
        data, self.current, self.selection = state
        self.doc = pymupdf.open(stream=data, filetype="pdf")
        self.stamp_xrefs.clear()
        self.sel_stamp = None
        self.dirty = True
        self._doc_changed(structure=True)
        self._ensure_thumb_visible(self.current)

    def undo(self):
        if not self.undo_stack:
            self.status("Niente da annullare.")
            return
        self.redo_stack.append(self._state())
        self._restore(self.undo_stack.pop())
        self.status("Annullato.")

    def redo(self):
        if not self.redo_stack:
            return
        self.undo_stack.append(self._state())
        self._restore(self.redo_stack.pop())
        self.status("Ripristinato.")

    # ------------------------------------------------------- pagine
    def _targets(self):
        return self.selection if self.selection else {self.current}

    def rotate(self, delta, all_pages=False):
        if not self.doc:
            return
        pages = list(range(len(self.doc))) if all_pages else sorted(self._targets())
        self._snapshot()
        for i in pages:
            page = self.doc[i]
            page.set_rotation((page.rotation + delta) % 360)
        self._doc_changed(pages=pages)
        self.status(f"Ruotate {len(pages)} pagine." if len(pages) > 1 else f"Ruotata pagina {pages[0] + 1}.")

    def delete_pages(self):
        if not self.doc:
            return
        pages = sorted(self._targets())
        if len(pages) == len(self.doc):
            messagebox.showwarning(APP_NAME, "Non puoi eliminare tutte le pagine.", parent=self.root)
            return
        what = f"la pagina {pages[0] + 1}" if len(pages) == 1 else f"{len(pages)} pagine"
        if not messagebox.askyesno(APP_NAME, f"Eliminare {what}?", parent=self.root):
            return
        self._snapshot()
        for i in reversed(pages):
            self.doc.delete_page(i)
        self.go_to(min(pages[0], len(self.doc) - 1), structure=True)
        self.status(f"Eliminate {len(pages)} pagine.")

    def move_pages(self, direction):
        if not self.doc:
            return
        pages = sorted(self._targets())
        if (direction < 0 and pages[0] == 0) or (direction > 0 and pages[-1] == len(self.doc) - 1):
            return
        self._snapshot()
        if direction < 0:
            for i in pages:
                self.doc.move_page(i, i - 1)
        else:
            for i in reversed(pages):
                self.doc.move_page(i, i + 2 if i + 2 < len(self.doc) else -1)
        self.selection = {i + direction for i in pages}
        if self.current in pages:
            self.current += direction
        self._doc_changed(structure=True)
        self._ensure_thumb_visible(self.current)

    def select_all(self):
        if self.doc:
            self.selection = set(range(len(self.doc)))
            self._paint_selection()
            self._update_controls()

    # ------------------------------------------------------- navigazione
    def go_to(self, index, at_bottom=False, structure=False):
        if not self.doc:
            return
        self.current = max(0, min(index, len(self.doc) - 1))
        self.selection = {self.current}
        self.anchor = self.current
        if structure:
            self._doc_changed(structure=True)
        else:
            self._render_page()
            self._paint_selection()
            self._update_controls()
        self.view.yview_moveto(1.0 if at_bottom else 0.0)
        self._ensure_thumb_visible(self.current)
        self._redraw_ghost()

    def prev_page(self):
        if self.doc and self.current > 0:
            self.go_to(self.current - 1)

    def next_page(self):
        if self.doc and self.current < len(self.doc) - 1:
            self.go_to(self.current + 1)

    def _goto_entry(self, _event):
        try:
            self.go_to(int(self.page_var.get()) - 1)
        except ValueError:
            pass
        self.view.focus_set()

    def _doc_changed(self, structure=False, pages=()):
        n = len(self.doc) if self.doc else 0
        self.current = min(self.current, max(n - 1, 0))
        self.selection = {i for i in self.selection if i < n} or ({self.current} if n else set())
        if self._find and self._find["doc"] is self.doc:
            self.clear_find(keep_text=True)
        if structure:
            self.thumb_photos.clear()
            self.tab.ov_photos.clear()
        else:
            for i in pages:
                self.thumb_photos.pop(i, None)
                self.tab.ov_photos.pop(i, None)
        self._layout_thumbs()
        self._render_page()
        self._update_controls()

    def _update_controls(self):
        n = len(self.doc) if self.doc else 0
        self.page_var.set(str(self.current + 1) if n else "")
        self.page_count_lbl.configure(text=f"/ {n}")
        name = self.tab.name if self.tab is not self.no_tab else ""
        self.root.title(f"{name}{' *' if self.dirty else ''} — {APP_NAME}" if name else APP_NAME)
        self._refresh_tabbar()
        self.zoom_lbl.configure(text=f"{self.cur_zoom / self.px_per_pt * 100:.0f}%" if n else "")
        sel = len(self.selection)
        self.repeat_btn.configure(text=f"Ripeti sulle {sel} pagine selezionate" if sel > 1
                                  else "Ripeti su tutte le pagine",
                                  state=tk.NORMAL if (n and self._repeat_position()) else tk.DISABLED)
        if n:
            page = self.doc[self.current]
            w, h = page.rect.width / MM, page.rect.height / MM
            rot = f" · ruotata {page.rotation}°" if page.rotation else ""
            sel_txt = f" · {sel} selezionate" if sel > 1 else ""
            self.info_var.set(f"Pagina {self.current + 1} di {n} · {w:.0f} × {h:.0f} mm{rot}{sel_txt}")
        else:
            self.info_var.set("")

    # ------------------------------------------------------- miniature
    def _layout_thumbs(self):
        """Riquadri delle miniature: in colonna, o in fila con il pannello in alto/in basso.

        thumb_layout ha per ogni pagina (x0, y0, x1, y1, scala) dell'immagine sulla tela.
        """
        c = self.thumbs
        c.delete("all")
        self.thumb_layout = []
        horizontal = self.thumbs_horizontal
        tw, th = self._thumb_box()
        if self.tab.thumb_w != (tw, th):   # miniature già pronte ma di un'altra misura
            self.thumb_photos.clear()
            self.tab.thumb_w = (tw, th)
        cw = tw + 34
        pos = 10
        for i in range(len(self.doc) if self.doc else 0):
            r = self.doc[i].rect
            s = min(tw / r.width, th / r.height)
            w, h = r.width * s, r.height * s
            if horizontal:
                x0, y0 = pos, 10 + (th - h) / 2
                c.create_rectangle(x0 - 6, 5, x0 + w + 6, 10 + th + 22, fill="", outline="", width=2, tags=f"bg{i}")
                c.create_text(x0 + w / 2, 10 + th + 11, text=str(i + 1), fill="#333", font=("Segoe UI", 9))
                pos += w + 24
            else:
                x0, y0 = (cw - w) / 2, pos
                c.create_rectangle(5, y0 - 5, cw - 5, y0 + h + 22, fill="", outline="", width=2, tags=f"bg{i}")
                c.create_text(cw / 2, y0 + h + 11, text=str(i + 1), fill="#333", font=("Segoe UI", 9))
                pos += h + 34
            c.create_rectangle(x0, y0, x0 + w, y0 + h, fill="white", outline="#a0a0a8")
            if i in self.thumb_photos:
                c.create_image(x0, y0, anchor="nw", image=self.thumb_photos[i])
            self.thumb_layout.append((x0, y0, x0 + w, y0 + h, s))
        if self.doc is not None and not len(self.doc):
            c.create_text(cw / 2 if not horizontal else 120, 70 if not horizontal else 50, justify=tk.CENTER,
                          text="Documento vuoto:\ntrascina qui\nle pagine da\nun'altra scheda",
                          fill="#666", font=("Segoe UI", 9))
        self.thumb_total = max(pos, 1)
        c.configure(scrollregion=(0, 0, pos, th + 40) if horizontal else (0, 0, cw, pos))
        self._paint_selection()
        self._schedule_thumbs()

    def _schedule_thumbs(self):
        if self._thumb_job is None:
            self._thumb_job = self.root.after(20, self._render_thumbs_step)

    def _visible_range(self):
        """Tratto visibile delle miniature lungo la direzione in cui scorrono (con un margine)."""
        c = self.thumbs
        if self.thumbs_horizontal:
            return c.canvasx(0) - 300, c.canvasx(c.winfo_width()) + 300
        return c.canvasy(0) - 300, c.canvasy(c.winfo_height()) + 300

    def _render_thumbs_step(self):
        """Disegna una miniatura visibile alla volta, così l'interfaccia resta reattiva."""
        self._thumb_job = None
        if not self.doc:
            return
        lo, hi = self._visible_range()
        main = (0, 2) if self.thumbs_horizontal else (1, 3)
        for i, box in enumerate(self.thumb_layout):
            if box[main[1]] < lo or i in self.thumb_photos:
                continue
            if box[main[0]] > hi:
                break
            s = box[4]
            pix = self.doc[i].get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
            self.thumb_photos[i] = photo_from_pixmap(pix)
            self.thumbs.create_image(box[0], box[1], anchor="nw", image=self.thumb_photos[i])
            self._thumb_job = self.root.after(1, self._render_thumbs_step)
            return

    def _paint_selection(self):
        for i in range(len(self.thumb_layout)):
            self.thumbs.itemconfigure(f"bg{i}", fill="#bcd4f6" if i in self.selection else "",
                                      outline="#1a73e8" if i == self.current else "")

    def _ensure_thumb_visible(self, i):
        if i >= len(self.thumb_layout):
            return
        c = self.thumbs
        x0, y0, x1, y1, _s = self.thumb_layout[i]
        if self.thumbs_horizontal:
            a0, a1, start, size, moveto = x0, x1, c.canvasx(0), c.winfo_width(), c.xview_moveto
        else:
            a0, a1, start, size, moveto = y0, y1 + 20, c.canvasy(0), c.winfo_height(), c.yview_moveto
        if a0 - 10 < start:
            moveto((a0 - 10) / self.thumb_total)
        elif a1 + 10 > start + size:
            moveto((a1 + 10 - size) / self.thumb_total)

    def _on_thumbs_wheel(self, event):
        scroll = self.thumbs.xview_scroll if self.thumbs_horizontal else self.thumbs.yview_scroll
        scroll(-3 if event.delta > 0 else 3, "units")
        return "break"

    def _thumb_at(self, event):
        x, y = self.thumbs.canvasx(event.x), self.thumbs.canvasy(event.y)
        for i, (x0, y0, x1, y1, _s) in enumerate(self.thumb_layout):
            if (x0 - 8 <= x <= x1 + 8) if self.thumbs_horizontal else (y0 - 6 <= y <= y1 + 26):
                return i
        return None

    def _on_thumb_click(self, event, mode):
        self._page_drag = None
        i = self._thumb_at(event)
        if i is None:
            return
        self.view.focus_set()
        self.sel_stamp = None
        collapse = False
        if mode == "toggle":
            self.selection ^= {i}
            self.anchor = i
        elif mode == "range":
            a, b = sorted((self.anchor, i))
            self.selection = set(range(a, b + 1))
        elif i in self.selection and len(self.selection) > 1:
            collapse = True     # resta la selezione multipla, che si può trascinare; al rilascio si riduce
        else:
            self.selection = {i}
            self.anchor = i
        if self.current != i:
            self.current = i
            self._render_page()
            self.view.yview_moveto(0)
        self._paint_selection()
        self._update_controls()
        self._page_drag = {"tab": self.tab, "index": i, "x": event.x_root, "y": event.y_root, "active": False,
                           "collapse": collapse, "target": None, "hover_tab": None}

    # trascinamento delle miniature: riordina le pagine o le copia in un'altra scheda
    def _on_thumb_motion(self, event):
        d = self._page_drag
        if not d:
            return
        if not d["active"]:
            if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 8:
                return
            pages = sorted(self.selection) if d["index"] in self.selection else [d["index"]]
            d.update(active=True, pages=pages)
            tip = d["tip"] = tk.Toplevel(self.root)
            tip.overrideredirect(True)
            tip.attributes("-topmost", True)
            tk.Label(tip, text=f"Pagina {pages[0] + 1}" if len(pages) == 1 else f"{len(pages)} pagine",
                     bg="#1a73e8", fg="white", font=("Segoe UI", 9, "bold"), padx=8, pady=3).pack()
        d["tip"].geometry(f"+{event.x_root + 16}+{event.y_root + 14}")
        self._update_drop_target(d, event.x_root, event.y_root)

    def _drop_widget(self, x, y):
        """Zona della finestra sotto il punto (coordinate dello schermo). Non si usa winfo_containing:
        l'etichetta che segue il mouse a volte resta un istante sotto il puntatore."""
        def inside(w):
            return (w is not None and w.winfo_ismapped() and 0 <= x - w.winfo_rootx() < w.winfo_width()
                    and 0 <= y - w.winfo_rooty() < w.winfo_height())
        return next((w for w in (*self._tab_frames, self._plus, self.tabbar, self.thumbs, self.view) if inside(w)),
                    None)

    def _update_drop_target(self, d, x, y):
        widget = self._drop_widget(x, y)
        tab = getattr(widget, "doc_tab", None)
        self.thumbs.delete("dropline")
        target, hint = None, "Rilascia sulle miniature per spostare le pagine, sulla scheda di un altro PDF per copiarle."
        if tab is not None:
            target = ("tab", tab)
            hint = ("Rilascia per copiarle in fondo a «%s», oppure aspetta: la scheda si apre e scegli il punto."
                    % tab.name if tab is not d["tab"] else hint)
        elif widget is self.tabbar or widget is self._plus:
            target, hint = ("new",), "Rilascia per creare un nuovo documento con queste pagine."
        elif widget is self.thumbs and self.doc is not None:
            th = self.thumbs
            if self.thumbs_horizontal:
                k = self._insert_index(th.canvasx(x - th.winfo_rootx()))
                start, size, pos, scroll = th.winfo_rootx(), th.winfo_width(), x, th.xview_scroll
            else:
                k = self._insert_index(th.canvasy(y - th.winfo_rooty()))
                start, size, pos, scroll = th.winfo_rooty(), th.winfo_height(), y, th.yview_scroll
            target = ("thumbs", k)
            hint = ("Rilascia per spostarle qui." if d["tab"] is self.tab
                    else f"Rilascia per copiarle qui in «{self.tab.name}».")
            self._draw_drop_line(k)
            if pos < start + 30 or pos > start + size - 30:   # vicino al bordo: scorre
                scroll(-1 if pos < start + 30 else 1, "units")
        elif widget is self.view and d["tab"] is not self.tab and self.doc is not None:
            target, hint = ("thumbs", len(self.doc)), f"Rilascia per copiarle in fondo a «{self.tab.name}»."
        d["target"] = target
        self.status(hint)
        # fermandosi su un'altra scheda la si apre, per scegliere dove inserire le pagine
        over = target[1] if target and target[0] == "tab" and target[1] is not self.tab else None
        if over is not d["hover_tab"]:
            d["hover_tab"] = over
            if over:
                self.root.after(700, lambda: self._spring_tab(d, over))
        drop = "new" if target == ("new",) else over
        if drop is not self._drop_tab:
            self._drop_tab = drop
            self._refresh_tabbar()

    def _spring_tab(self, d, tab):
        if self._page_drag is d and d["hover_tab"] is tab and tab in self.tabs:
            d["hover_tab"] = None
            self._drop_tab = None
            self.switch_tab(tab)
            self.status(f"Scheda «{tab.name}»: rilascia sulle miniature nel punto in cui inserire le pagine.")

    def _insert_index(self, pos):
        """Posizione di inserimento fra le miniature per la coordinata della tela lungo la fila."""
        a0, a1 = (0, 2) if self.thumbs_horizontal else (1, 3)
        for i, box in enumerate(self.thumb_layout):
            if pos < (box[a0] + box[a1]) / 2:
                return i
        return len(self.thumb_layout)

    def _draw_drop_line(self, k):
        layout = self.thumb_layout
        if self.thumbs_horizontal:
            x = 8 if not layout else layout[k][0] - 12 if k < len(layout) else layout[-1][2] + 12
            th = self._thumb_box()[1]
            self.thumbs.create_line(max(x, 3), 6, max(x, 3), th + 30, fill="#1a73e8", width=3, tags="dropline")
        else:
            y = 10 if not layout else layout[k][1] - 12 if k < len(layout) else layout[-1][3] + 28
            self.thumbs.create_line(8, max(y, 3), self.thumb_w + 26, max(y, 3), fill="#1a73e8", width=3,
                                    tags="dropline")

    def _on_thumb_release(self, _event):
        d, self._page_drag = self._page_drag, None
        if not d:
            return
        if not d["active"]:
            if d["collapse"] and d["tab"] is self.tab:
                self.selection = {d["index"]}
                self.anchor = d["index"]
                self._paint_selection()
                self._update_controls()
            return
        d["tip"].destroy()
        self.thumbs.delete("dropline")
        self._drop_tab = None
        self._refresh_tabbar()
        target, src, pages = d["target"], d["tab"], d["pages"]
        if src not in self.tabs or not target:
            self.status("Trascinamento annullato.")
            return
        if target[0] == "thumbs":
            if src is self.tab:
                self._move_pages_to(pages, target[1])
            else:
                self._copy_pages(src, pages, self.tab, target[1])
        elif target[0] == "tab":
            if target[1] is not src:
                self._copy_pages(src, pages, target[1], len(target[1].doc))
            else:
                self.status("Trascinamento annullato.")
        else:
            self._copy_pages(src, pages, None, 0)

    def _move_pages_to(self, pages, pos):
        """Sposta le pagine (della scheda attiva) davanti alla pagina pos."""
        rest = [i for i in range(len(self.doc)) if i not in pages]
        k = pos - sum(1 for p in pages if p < pos)
        order = rest[:k] + pages + rest[k:]
        if order == list(range(len(self.doc))):
            return
        self._snapshot()
        self.doc.select(order)
        self.selection = set(range(k, k + len(pages)))
        self.current = self.anchor = k
        self._doc_changed(structure=True)
        self._ensure_thumb_visible(k)
        self.status("Pagine spostate. (Ctrl+Z per annullare)")

    def _copy_pages(self, src, pages, dst, pos):
        """Copia le pagine della scheda src nella scheda dst, davanti alla pagina pos.

        dst None: crea un nuovo documento con quelle pagine.
        """
        part = pymupdf.open(stream=src.doc.tobytes(), filetype="pdf")
        part.select(pages)
        what = f"la pagina {pages[0] + 1}" if len(pages) == 1 else f"{len(pages)} pagine"
        if dst is None:
            doc = pymupdf.open()
            doc.insert_pdf(part)
            dst = DocTab(doc)
            dst.dirty = True
            self._add_tab(dst, switch=False)
            self.status(f"Creato «{dst.name}» con {what} di «{src.name}»: clicca la sua scheda per vederlo.")
            return
        self._snapshot(dst)
        dst.doc.insert_pdf(part, start_at=pos if pos < len(dst.doc) else -1)
        dst.thumb_photos.clear()
        dst.ov_photos.clear()
        if dst is self.tab:
            self.current = self.anchor = pos
            self.selection = set(range(pos, pos + len(pages)))
            self._doc_changed(structure=True)
            self._ensure_thumb_visible(pos)
        else:
            self._refresh_tabbar()
        where = "" if dst is self.tab else " (in fondo)"
        self.status(f"{'Copiata' if len(pages) == 1 else 'Copiate'} {what} da «{src.name}» in «{dst.name}»{where}. "
                    "Ctrl+Z in quella scheda per annullare.")

    def _on_thumb_menu(self, event):
        i = self._thumb_at(event)
        if i is None:
            return
        if i not in self.selection:
            self._on_thumb_click(event, "single")
        self._page_drag = None
        menu = tk.Menu(self.root, tearoff=False)
        menu.add_command(label="Ruota a destra", command=lambda: self.rotate(90))
        menu.add_command(label="Ruota a sinistra", command=lambda: self.rotate(-90))
        menu.add_separator()
        menu.add_command(label="Sposta su", command=lambda: self.move_pages(-1))
        menu.add_command(label="Sposta giù", command=lambda: self.move_pages(1))
        menu.add_separator()
        menu.add_command(label="Estrai…", command=self.cmd_extract)
        menu.add_command(label="Elimina", command=self.delete_pages)
        menu.tk_popup(event.x_root, event.y_root)

    # ------------------------------------------------------- vista pagina
    def _compute_zoom(self, page):
        r = page.rect
        cw = max(self.view.winfo_width() - 2 * PAD, 50)
        ch = max(self.view.winfo_height() - 2 * PAD, 50)
        if self.zoom_mode == "width":
            z = cw / r.width
        elif self.zoom_mode == "page":
            z = min(cw / r.width, ch / r.height)
        else:
            z = self.zoom * self.px_per_pt
        max_z = (MAX_RENDER_PX / (r.width * r.height)) ** 0.5
        return max(0.05, min(z, max_z))

    def _render_page(self, without=None):
        """Mostra la pagina corrente; without = timbro apposto da non disegnare (mentre lo si trascina)."""
        c = self.view
        c.delete("all")
        self.page_photo = None
        self.page_stamps = []
        self._page_lines = None
        if not self.doc:
            text = ("Documento vuoto.\nTrascina qui le pagine dalle miniature di un'altra scheda,\n"
                    "oppure usa «Aggiungi…»." if self.doc is not None else
                    "Apri un PDF con «Apri…» (Ctrl+O)\noppure unisci più file con «Aggiungi…»\n\n"
                    "Puoi aprire più PDF insieme: ognuno va nella sua scheda.")
            c.configure(scrollregion=(0, 0, 0, 0))
            c.create_text(max(c.winfo_width(), 400) / 2, max(c.winfo_height(), 300) / 2, justify=tk.CENTER,
                          fill="#f0f0f0", font=("Segoe UI", 13), text=text)
            return
        page = self.doc[self.current]
        z = self.cur_zoom = self._compute_zoom(page)
        matrix = pymupdf.Matrix(z, z)
        if without:
            others = [x for i, x in enumerate(page.get_contents()) if i != without.index]
            pix = render_with_contents(page, others, matrix=matrix, alpha=False)
        else:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
        self.page_photo = photo_from_pixmap(pix)
        cw, ch = c.winfo_width(), c.winfo_height()
        ox = max(PAD, (cw - pix.width) // 2)
        oy = max(PAD, (ch - pix.height) // 2)
        self.page_origin = (ox, oy)
        self.scroll_size = (max(cw, pix.width + 2 * PAD), max(ch, pix.height + 2 * PAD))
        c.create_rectangle(ox + 4, oy + 4, ox + pix.width + 4, oy + pix.height + 4, fill="#3c3d44", outline="")
        c.create_image(ox, oy, anchor="nw", image=self.page_photo)
        c.configure(scrollregion=(0, 0, *self.scroll_size))
        c.configure(cursor=TOOL_CURSORS.get(self.tool, "arrow"))
        if self._sel_page != self.current:
            self.sel_stamp, self._sel_page = None, self.current
        self.page_stamps = placed_objects(page)
        if self.sel_stamp not in {s.index for s in self.page_stamps}:
            self.sel_stamp = None
        self._draw_stamp_marks()
        self._draw_find_marks()
        self._redraw_ghost()

    def _on_view_configure(self, _event):
        if self._render_job:
            self.root.after_cancel(self._render_job)
        self._render_job = self.root.after(80, self._render_page)

    def _view_to_page(self, x, y):
        """Coordinate della finestra -> punto sulla pagina così come la si vede (in punti PDF)."""
        r = self.doc[self.current].rect
        ox, oy = self.page_origin
        return (self.view.canvasx(x) - ox) / self.cur_zoom + r.x0, (self.view.canvasy(y) - oy) / self.cur_zoom + r.y0

    def set_zoom_mode(self, mode):
        self.zoom_mode = mode
        self._render_page()
        self._update_controls()

    def set_real_size(self):
        self.zoom_mode = None
        self.zoom = 1.0
        self._render_page()
        self._update_controls()

    def _zoom_at(self, factor, x=None, y=None):
        if not self.doc:
            return
        if x is None:
            x, y = self.view.winfo_width() / 2, self.view.winfo_height() / 2
        px, py = self._view_to_page(x, y)
        self.zoom_mode = None
        self.zoom = min(max(self.cur_zoom / self.px_per_pt * factor, 0.1), 8.0)
        self._render_page()
        # mantiene sotto il cursore lo stesso punto della pagina
        r = self.doc[self.current].rect
        ox, oy = self.page_origin
        sw, sh = self.scroll_size
        self.view.xview_moveto((ox + (px - r.x0) * self.cur_zoom - x) / sw)
        self.view.yview_moveto((oy + (py - r.y0) * self.cur_zoom - y) / sh)
        self._update_controls()
        self._redraw_ghost()

    def _on_wheel(self, event):
        if not self.doc:
            return
        step = -1 if event.delta > 0 else 1
        if event.state & 0x0004:  # Ctrl: zoom
            self._zoom_at(1.15 if step < 0 else 1 / 1.15, event.x, event.y)
            return "break"
        if event.state & 0x0001:  # Maiusc: scorrimento orizzontale
            self.view.xview_scroll(step * 3, "units")
        else:
            top, bottom = self.view.yview()
            now = time.monotonic()
            if step > 0 and bottom >= 0.999 and self.current < len(self.doc) - 1:
                if now - self._last_flip > 0.4:
                    self._last_flip = now
                    self.go_to(self.current + 1)
            elif step < 0 and top <= 0.001 and self.current > 0:
                if now - self._last_flip > 0.4:
                    self._last_flip = now
                    self.go_to(self.current - 1, at_bottom=True)
            else:
                self.view.yview_scroll(step * 3, "units")
        self._redraw_ghost()
        return "break"

    def _on_press(self, event):
        self.view.focus_set()
        if self.tool == "stamp":
            self._place_stamp(event.x, event.y)
            return
        if self.tool in ("edit_text", "add_text"):
            box = self._text_box_at(event.x, event.y)
            if box:   # casella di testo: trascinandola si sposta, con un clic si corregge
                self.view.delete("hover")
                self._drag = {"stamp": box, "start": self._view_to_page(event.x, event.y), "delta": (0, 0),
                              "lifted": False, "edit": True}
            else:
                self._text_click(event.x, event.y)
            return
        if self.tool in ("whiteout", "blackout"):
            if self.doc and self.doc[self.current].rect.contains(pymupdf.Point(self._view_to_page(event.x, event.y))):
                cx, cy = self.view.canvasx(event.x), self.view.canvasy(event.y)
                self._area = (event.x, event.y)
                self.view.create_rectangle(cx, cy, cx, cy, tags="area", width=2, dash=(4, 3),
                                           outline="#d32f2f" if self.tool == "blackout" else "#1a73e8")
            return
        hit = self._stamp_at(event.x, event.y)
        if hit:
            # un timbro o una casella di testo: si seleziona e, se il mouse si muove, si trascina
            self.sel_stamp = hit.index
            self._drag = {"stamp": hit, "start": self._view_to_page(event.x, event.y), "delta": (0, 0),
                          "lifted": False}
            self.status("Testo selezionato: trascinalo per spostarlo, doppio clic per correggerlo, Canc per "
                        "eliminarlo." if isinstance(hit, PlacedText) else
                        "Timbro selezionato: trascinalo per spostarlo, doppio clic per modificarlo, tasto destro per "
                        "ruotarlo, Canc per eliminarlo.")
        else:
            if self.sel_stamp is not None:
                self.sel_stamp = None
                self._draw_stamp_marks()
            self.view.scan_mark(event.x, event.y)
            self.view.configure(cursor="fleur")

    def _on_drag(self, event):
        if self._area:
            x0, y0 = self._area
            self.view.coords("area", self.view.canvasx(x0), self.view.canvasy(y0),
                             self.view.canvasx(event.x), self.view.canvasy(event.y))
            return
        if self.tool and not (self._drag and self.tool in ("edit_text", "add_text")):
            return
        d = self._drag
        if not d:
            self.view.scan_dragto(event.x, event.y, gain=1)
            return
        px, py = self._view_to_page(event.x, event.y)
        r, pr, z = d["stamp"].rect, self.doc[self.current].rect, self.cur_zoom
        dx = min(max(px - d["start"][0], pr.x0 - r.x0), pr.x1 - r.x1) if r.width < pr.width else 0
        dy = min(max(py - d["start"][1], pr.y0 - r.y0), pr.y1 - r.y1) if r.height < pr.height else 0
        if not d["lifted"]:
            if abs(dx) * z < 3 and abs(dy) * z < 3:
                return
            self._lift_stamp(d)
        d["delta"] = (dx, dy)
        x0, y0, x1, y1 = self._stamp_box(d["stamp"])
        self.view.coords(d["ghost"], x0 + dx * z, y0 + dy * z)
        self.view.coords(d["outline"], x0 + dx * z, y0 + dy * z, x1 + dx * z, y1 + dy * z)

    def _lift_stamp(self, d):
        """Toglie il timbro dalla pagina mostrata e lo fa seguire il mouse."""
        page, stamp, z = self.doc[self.current], d["stamp"], self.cur_zoom
        only = [page.get_contents()[stamp.index]]
        pix = render_with_contents(page, only, matrix=pymupdf.Matrix(z, z), clip=stamp.rect, alpha=True)
        self._render_page(without=stamp)
        d["photo"] = photo_from_pixmap(pix)
        x0, y0, x1, y1 = self._stamp_box(stamp)
        d["ghost"] = self.view.create_image(x0, y0, anchor="nw", image=d["photo"])
        d["outline"] = self.view.create_rectangle(x0, y0, x1, y1, outline="#1a73e8", dash=(4, 3))
        d["lifted"] = True
        self.view.configure(cursor="fleur")

    def _on_release(self, event):
        if self._area:
            self._apply_area(event.x, event.y)
            return
        if self.tool and not (self._drag and self.tool in ("edit_text", "add_text")):
            return
        d, self._drag = self._drag, None
        if not d:
            self.view.configure(cursor="arrow")
        elif d["lifted"] and d["delta"] != (0, 0):
            self._snapshot()
            obj = d["stamp"]
            if isinstance(obj, PlacedText):
                move_placed_text(self.doc[self.current], obj, *d["delta"])
            else:
                move_placed_stamp(self.doc[self.current], obj, *d["delta"])
            self._doc_changed(pages=[self.current])
            self.status(f"{'Testo' if isinstance(obj, PlacedText) else 'Timbro'} spostato. (Ctrl+Z per annullare)")
        elif d["lifted"]:
            self._render_page()
        elif d.get("edit"):   # clic senza trascinare su una casella, con uno strumento di testo
            self.edit_text_box(d["stamp"])
        else:
            self._draw_stamp_marks()

    def _on_motion(self, event):
        self._mouse = (event.x, event.y)
        if self.tool == "stamp":
            self._redraw_ghost()
        elif self.tool in ("edit_text", "add_text"):
            self._hover_text(event.x, event.y)
        elif self.tool is None and self.doc and not self._drag:
            hit = self._stamp_at(event.x, event.y)
            self.view.configure(cursor="fleur" if hit else "arrow")
            self._draw_stamp_marks(hover=hit)

    def _on_leave(self, _event):
        self._mouse = None
        self.view.delete("ghost", "hover")
        if not self._drag:
            self._draw_stamp_marks()

    def _on_double_click(self, event):
        hit = self._stamp_at(event.x, event.y) if self.tool is None else None
        if isinstance(hit, PlacedText):
            self._drag = None
            self.edit_text_box(hit)
        elif hit:
            self._drag = None
            self.edit_placed_stamp(hit)
        else:
            # due clic veloci nello stesso punto (es. timbrando più pagine di fila) valgono come due clic
            self._on_press(event)

    def _on_view_menu(self, event):
        hit = self._stamp_at(event.x, event.y)
        if not hit:
            return
        self.sel_stamp = hit.index
        self._draw_stamp_marks()
        menu = tk.Menu(self.root, tearoff=False)
        if isinstance(hit, PlacedText):
            menu.add_command(label="Modifica testo…", command=lambda: self.edit_text_box(hit))
            menu.add_command(label="Elimina questo testo", command=self.delete_selected_stamp)
        else:
            menu.add_command(label="Modifica questo timbro…", command=lambda: self.edit_placed_stamp(hit))
            menu.add_command(label="Ruota a destra", command=lambda: self.rotate_placed(hit, 90))
            menu.add_command(label="Ruota a sinistra", command=lambda: self.rotate_placed(hit, -90))
            menu.add_separator()
            menu.add_command(label="Elimina questo timbro", command=self.delete_selected_stamp)
        menu.tk_popup(event.x_root, event.y_root)

    # ------------------------------------------------------- timbri e caselle di testo già sulla pagina
    def _stamp_at(self, x, y):
        """Timbro o casella di testo sotto il mouse (il più in alto, se si sovrappongono)."""
        if not (self.doc and self.page_stamps):
            return None
        point = pymupdf.Point(self._view_to_page(x, y))
        return next((s for s in reversed(self.page_stamps) if s.rect.contains(point)), None)

    def _canvas_rect(self, r):
        """Rettangolo della pagina vista -> coordinate della tela."""
        z, (ox, oy), pr = self.cur_zoom, self.page_origin, self.doc[self.current].rect
        return ox + (r.x0 - pr.x0) * z, oy + (r.y0 - pr.y0) * z, ox + (r.x1 - pr.x0) * z, oy + (r.y1 - pr.y0) * z

    def _stamp_box(self, stamp):
        """Riquadro del timbro apposto in coordinate della vista."""
        return self._canvas_rect(stamp.rect)

    def _draw_stamp_marks(self, hover=None):
        """Bordo blu sul timbro apposto selezionato (continuo) e su quello sotto il mouse (tratteggiato)."""
        self.view.delete("mark")
        if self.tool or self._drag or not self.doc:
            return
        for s in self.page_stamps:
            if s.index == self.sel_stamp:
                self.view.create_rectangle(*self._stamp_box(s), outline="#1a73e8", width=2, tags="mark")
            elif hover and s.index == hover.index:
                self.view.create_rectangle(*self._stamp_box(s), outline="#1a73e8", dash=(4, 3), tags="mark")

    def delete_selected_stamp(self):
        stamp = next((s for s in self.page_stamps if s.index == self.sel_stamp), None)
        if not stamp:
            return
        self._snapshot()
        remove_placed_stamp(self.doc[self.current], stamp)
        self.sel_stamp = None
        self._doc_changed(pages=[self.current])
        self.status(f"{'Testo' if isinstance(stamp, PlacedText) else 'Timbro'} eliminato. (Ctrl+Z per annullare)")

    def _on_delete_key(self):
        if self.sel_stamp is not None:
            self.delete_selected_stamp()
        else:
            self.delete_pages()

    def rotate_placed(self, stamp, angle):
        """Ruota un timbro già apposto attorno al suo centro (solo questo, su questa pagina)."""
        self._snapshot()
        rotate_placed_stamp(self.doc[self.current], stamp, angle)
        self.sel_stamp = stamp.index
        self._doc_changed(pages=[self.current])
        self.status(f"Timbro ruotato {'a destra' if angle > 0 else 'a sinistra'}. (Ctrl+Z per annullare)")

    def edit_placed_stamp(self, stamp):
        """Modifica un timbro già apposto (solo questo): ruota, ritocca e, per le firme create nel
        programma, riscrive il testo o ridisegna; il timbro resta dov'era, primo o secondo piano compreso."""
        page = self.doc[self.current]
        try:
            pix, xref = placed_stamp_image(page, stamp)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile leggere l'immagine del timbro:\n{exc}", parent=self.root)
            return
        angle = stamp_view_angle(page, stamp)
        source, baked = stamp_source(self.doc, xref) or self._gallery_source(pix) or (None, 0)
        StampEditor.last_image_dir = self._dir("image")
        editor = StampEditor(self.root, self.fonts, rotate_pixmap(pix, angle), stamp.rect.width / MM,
                             "Modifica timbro sulla pagina", source=source, source_rotation=(baked + angle) % 360)
        self.root.wait_window(editor.top)
        self._remember_dir("image", StampEditor.last_image_dir)
        r = editor.result
        if not r:
            return
        self._snapshot()
        replace_placed_stamp(page, stamp, r["image"], r["width_mm"], r["source"],
                             (editor.source_rotation + r["rotation"]) % 360)
        self.sel_stamp = stamp.index
        self._doc_changed(pages=[self.current])
        self.status("Timbro modificato su questa pagina; quello nella galleria non cambia. (Ctrl+Z per annullare)")

    def _gallery_source(self, pix):
        """(source, rotazione) della firma della galleria con la stessa immagine: serve per i timbri apposti
        prima che il PDF ricordasse il testo delle firme."""
        alpha = pymupdf.Pixmap(None, pix).samples if pix.alpha else None
        for item in self.library.items:
            p = self.library.stamps[item["id"]].pixmap
            if not item.get("source") or (p.width, p.height, bool(p.alpha)) != (pix.width, pix.height, bool(pix.alpha)):
                continue
            # i colori possono differire di un'unità per gli arrotondamenti del PDF; la trasparenza no
            if (pymupdf.Pixmap(None, p).samples == alpha) if alpha is not None else p.samples == pix.samples:
                return item["source"], item["rotate"]
        return None

    # ------------------------------------------------------- timbro
    def _stamp_size_pt(self):
        w = self.stamp_item["width_mm"] * MM
        return w, w * self.stamp.aspect

    def _stamp_rect(self, page, px, py):
        """Rettangolo del timbro centrato in (px, py), tenuto dentro la pagina (coordinate della vista)."""
        w, h = self._stamp_size_pt()
        r = page.rect
        px = r.x0 + r.width / 2 if w >= r.width else min(max(px, r.x0 + w / 2), r.x1 - w / 2)
        py = r.y0 + r.height / 2 if h >= r.height else min(max(py, r.y0 + h / 2), r.y1 - h / 2)
        return pymupdf.Rect(px - w / 2, py - h / 2, px + w / 2, py + h / 2)

    def _insert_stamp(self, page, view_rect):
        # insert_image lavora sulla pagina non ruotata: si riporta il rettangolo
        # e si ruota l'immagine perché appaia diritta come la si vede a schermo
        # overlay=False mette il timbro sotto il contenuto esistente della pagina
        rect = view_rect * page.derotation_matrix
        rect.normalize()
        item = self.stamp_item
        xref = self.stamp_xrefs.get(item["id"])
        if xref:
            page.insert_image(rect, xref=xref, rotate=page.rotation, overlay=item["overlay"])
        else:
            xref = page.insert_image(rect, stream=self.stamp.png, rotate=page.rotation, overlay=item["overlay"])
            mark_stamp_image(self.doc, xref, item.get("source"), item["rotate"])
            self.stamp_xrefs[item["id"]] = xref

    def _redraw_ghost(self):
        v = self.view
        v.delete("ghost")
        if not (self.stamp_mode and self.stamp and self.doc and self._mouse):
            return
        page = self.doc[self.current]
        px, py = self._view_to_page(*self._mouse)
        if not page.rect.contains(pymupdf.Point(px, py)):
            return
        r = self._stamp_rect(page, px, py)
        z, (ox, oy), pr = self.cur_zoom, self.page_origin, page.rect
        x0, y0 = ox + (r.x0 - pr.x0) * z, oy + (r.y0 - pr.y0) * z
        w, h = round(r.width * z), round(r.height * z)
        if w * h > 25_000_000:
            return
        self._ghost_photo = self.stamp.photo(w, h)
        v.create_image(x0, y0, anchor="nw", image=self._ghost_photo, tags="ghost")
        v.create_rectangle(x0, y0, x0 + w, y0 + h, outline="#d32f2f", dash=(4, 3), tags="ghost")

    def _place_stamp(self, x, y):
        if not (self.doc and self.stamp):
            return
        page = self.doc[self.current]
        px, py = self._view_to_page(x, y)
        if not page.rect.contains(pymupdf.Point(px, py)):
            return
        rect = self._stamp_rect(page, px, py)
        self._snapshot()
        self._insert_stamp(page, rect)
        pr = page.rect
        center = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
        rx, ry = (center[0] - pr.x0) / pr.width, (center[1] - pr.y0) / pr.height
        self.last_stamp = (self.current, rx, ry, self.stamp_item["id"])
        self.stamp_item["last_pos"] = [round(rx, 4), round(ry, 4)]
        self.library.save()
        self._doc_changed(pages=[self.current])
        self.status(f"«{self.stamp_item['name']}» apposto a pagina {self.current + 1}. (Ctrl+Z per annullare)")

    def _repeat_position(self):
        """(pagina da saltare, x, y relative) per «Ripeti», oppure None.

        Vale l'ultimo clic fatto con questo timbro nel documento aperto; altrimenti
        la posizione usata con questo timbro l'ultima volta (anche in un altro PDF).
        """
        item = self.stamp_item
        if not item:
            return None
        if self.last_stamp and self.last_stamp[3] == item["id"]:
            return self.last_stamp[:3]
        if item["last_pos"]:
            return (None, *item["last_pos"])
        return None

    def repeat_last_stamp(self):
        if not (self.doc and self.stamp):
            return
        position = self._repeat_position()
        if not position:
            messagebox.showinfo(APP_NAME, "Prima apponi il timbro cliccando su una pagina.", parent=self.root)
            return
        skip, rx, ry = position
        pages = sorted(self.selection) if len(self.selection) > 1 else range(len(self.doc))
        pages = [i for i in pages if i != skip]
        if not pages:
            return
        where = "nella stessa posizione" if skip is not None else "nella posizione usata l'ultima volta"
        if not messagebox.askyesno(APP_NAME, f"Apporre «{self.stamp_item['name']}» {where} su {len(pages)} pagine?",
                                   parent=self.root):
            return
        self._snapshot()
        for i in pages:
            page = self.doc[i]
            pr = page.rect
            self._insert_stamp(page, self._stamp_rect(page, pr.x0 + rx * pr.width, pr.y0 + ry * pr.height))
        self._doc_changed(pages=pages)
        self.status(f"Timbro apposto su {len(pages)} pagine.")

    def toggle_stamp_mode(self):
        self.set_tool(None if self.tool == "stamp" else "stamp")

    def set_stamp_mode(self, on):
        if on:
            self.set_tool("stamp")
        elif self.tool == "stamp":
            self.set_tool(None)

    def set_tool(self, tool):
        """Strumento attivo sulla pagina: None (sposta timbri e testi), "stamp" o uno di TEXT_TOOLS."""
        if tool == "stamp" and not self.stamp:
            if messagebox.askyesno(APP_NAME, "Nessun timbro salvato.\nVuoi aggiungere ora l'immagine del timbro?",
                                   parent=self.root):
                self.cmd_add_stamp()
            if not self.stamp:
                tool = self.tool if self.tool != "stamp" else None
        self.tool = tool
        self.tool_var.set(tool if tool in TEXT_TOOLS else "")
        self._area = None
        self._paint_mode_btn()
        self.view.delete("area", "hover")
        self.view.configure(cursor=TOOL_CURSORS.get(tool, "arrow"))
        self._draw_stamp_marks()
        self._redraw_ghost()
        if tool:
            self.view.focus_set()
            self.status(TOOL_HINTS[tool])

    def _paint_mode_btn(self):
        if self.stamp_mode:
            self.mode_btn.configure(text="● MODALITÀ TIMBRO ATTIVA\nclicca sul foglio · Esc per uscire",
                                    bg="#2e7d32", fg="white", activebackground="#256628", activeforeground="white")
        else:
            bg, fg = self._btn_colors
            self.mode_btn.configure(text="Attiva modalità timbro  (T)", bg=bg, fg=fg,
                                    activebackground=bg, activeforeground=fg)

    # ------------------------------------------------------- testo
    def _lines(self):
        if self._page_lines is None:
            self._page_lines = text_lines(self.doc[self.current]) if self.doc else []
        return self._page_lines

    def _line_at(self, x, y):
        """Riga di testo del PDF sotto il mouse (la più piccola, se più righe si sovrappongono)."""
        page = self.doc[self.current]
        point = pymupdf.Point(self._view_to_page(x, y)) * page.derotation_matrix
        hits = [line for line in self._lines() if (line.rect + (-1, -1, 1, 1)).contains(point)]
        return min(hits, key=lambda line: line.rect.width * line.rect.height, default=None)

    def _text_box_at(self, x, y):
        hit = self._stamp_at(x, y)
        return hit if isinstance(hit, PlacedText) else None

    def _hover_text(self, x, y):
        """Evidenzia la casella o la riga che si modificherebbe cliccando."""
        self.view.delete("hover")
        if not self.doc:
            return
        box = self._text_box_at(x, y)
        if box:
            self.view.create_rectangle(*self._canvas_rect(box.rect), outline="#1a73e8", dash=(4, 3), width=2,
                                       tags="hover")
        elif self.tool == "edit_text":
            line = self._line_at(x, y)
            if line:
                page = self.doc[self.current]
                self.view.create_rectangle(*self._canvas_rect(line.rect * page.rotation_matrix), outline="#1a73e8",
                                           fill="#1a73e8", stipple="gray12", width=1, tags="hover")

    def _text_click(self, x, y):
        if not self.doc:
            return
        box = self._text_box_at(x, y)
        if box:
            self.edit_text_box(box)
        elif self.tool == "add_text":
            self.add_text_at(x, y)
        else:
            line = self._line_at(x, y)
            if line:
                self.edit_line(line)
            else:
                self.status("Qui non c'è testo: passa col mouse sulle righe, quelle modificabili si colorano di blu. "
                            "Nelle scansioni usa prima «Riconosci testo (OCR)».")

    def _write_text(self, page, params, replace=None):
        """Scrive o riscrive una casella di testo; restituisce lo xref del nuovo stream (None se riscritta)."""
        fontfile = self.fonts.get(params["font"])
        resname = font_resource(params["font"], self.tab.font_token)
        if fontfile:
            self.tab.fonts_added = True
        if replace:
            replace_placed_text(page, replace, params, fontfile, resname)
            return None
        return insert_text_box(page, params, fontfile, resname)

    def _ask_text(self, title, params, place, area, original=None, whiteout=None, cover=None):
        """Finestra del testo con l'anteprima dal vivo sulla pagina; restituisce il risultato della finestra.

        place(dati della finestra) -> parametri completi della casella (origine, rotazione); area: dove sta
        il testo nella pagina vista (la finestra si sposta per non coprirlo); cover: testo che verrà
        sostituito, coperto di bianco nell'anteprima.
        """
        page = self.doc[self.current]
        dialog = TextDialog(self.root, self.fonts, title, params, original=original, whiteout=whiteout,
                            preview=lambda p: self._show_text_preview(page, place(p) if p else None, cover))
        x0, y0, x1, y1 = self._canvas_rect(area)
        dx, dy = self.view.winfo_rootx() - self.view.canvasx(0), self.view.winfo_rooty() - self.view.canvasy(0)
        keep_clear(dialog.top, (x0 + dx, y0 + dy, x1 + dx, y1 + dy))
        self.root.wait_window(dialog.top)
        self._show_text_preview(page, None)
        return dialog.result

    def _show_text_preview(self, page, params, cover=None):
        """Mostra sulla pagina il testo che si sta scrivendo, com'è nel PDF (params None lo toglie)."""
        self.view.delete("textpreview")
        self._text_photo = None
        if not params or not self.doc or page.number != self.current:
            return
        try:
            pix, clip = text_preview(page, params, self.fonts.get(params["font"]), self.cur_zoom)
        except Exception:
            return
        if cover:
            self.view.create_rectangle(*self._canvas_rect(cover), fill="white", outline="", tags="textpreview")
        x0, y0, x1, y1 = self._canvas_rect(clip)
        self._text_photo = photo_from_pixmap(pix)
        self.view.create_image(x0, y0, anchor="nw", image=self._text_photo, tags="textpreview")
        self.view.create_rectangle(x0 - 2, y0 - 2, x1 + 2, y1 + 2, outline="#1a73e8", dash=(4, 3), tags="textpreview")

    def edit_line(self, line):
        """Corregge una riga del PDF: il vecchio testo si elimina e il nuovo si scrive nello stesso punto."""
        page = self.doc[self.current]
        font = match_font(line.font, line.flags, self.fonts) or "Arial"
        seen = line.rect * page.rotation_matrix
        r = self._ask_text("Modifica testo",
                           {"text": line.text, "font": font, "size": round(line.size, 1), "color": line.color},
                           lambda p: dict(p, origin=list(line.origin), rotate=line.rotate), seen,
                           original=f"Testo originale: {re.sub(r'^[A-Z]{6}[+]', '', line.font)}, {line.size:.1f} pt",
                           whiteout=image_under(page, line.rect), cover=seen + (-1, -1, 1, 1))
        if not r:
            return
        self._snapshot()
        if r["whiteout"]:
            erase(page, [line.rect + (-1, -1, 1, 1)], fill=(1, 1, 1), images=pymupdf.PDF_REDACT_IMAGE_PIXELS)
        else:
            erase(page, [text_band(line)])
        params = {k: r[k] for k in ("text", "font", "size", "color")}
        params.update(origin=list(line.origin), rotate=line.rotate)
        # la riga corretta resta dove stava il testo originale: sotto i timbri in primo piano
        place_with_content(page, self._write_text(page, params))
        self._doc_changed(pages=[self.current])
        self.status("Testo corretto: ora è una casella che puoi spostare o ricorreggere. (Ctrl+Z per annullare)")

    def edit_text_box(self, box):
        self._render_page(without=box)   # mentre la si corregge si vede solo l'anteprima
        r = self._ask_text("Modifica casella di testo", box.params, lambda p: dict(box.params, **p), box.rect)
        if not r:
            self._render_page()
            return
        self._snapshot()
        params = dict(box.params, **{k: r[k] for k in ("text", "font", "size", "color")})
        self._write_text(self.doc[self.current], params, replace=box)
        self._doc_changed(pages=[self.current])
        self.status("Casella di testo modificata. (Ctrl+Z per annullare)")

    def add_text_at(self, x, y):
        """Nuova casella di testo con l'angolo in alto a sinistra nel punto cliccato."""
        page = self.doc[self.current]
        px, py = self._view_to_page(x, y)
        if not page.rect.contains(pymupdf.Point(px, py)):
            return
        style = self.cfg["text_style"] or {"font": "Arial", "size": 11, "color": "#000000"}

        def place(p):
            """Il punto cliccato è l'angolo in alto a sinistra della casella."""
            ascender = text_font(self.fonts.get(p["font"])).ascender
            origin = pymupdf.Point(px, py + ascender * p["size"]) * page.derotation_matrix
            return dict(p, origin=[origin.x, origin.y], rotate=page.rotation)

        r = self._ask_text("Aggiungi testo", dict(style, text=""), place,
                           pymupdf.Rect(px, py, px + 40 * style["size"], py + 4 * style["size"]))
        if not r:
            return
        style = {k: r[k] for k in ("font", "size", "color")}
        self.cfg["text_style"] = style
        self._snapshot()
        self._write_text(page, place(dict(style, text=r["text"])))
        self._doc_changed(pages=[self.current])
        self.status("Testo aggiunto: trascinalo per spostarlo, doppio clic per correggerlo. (Ctrl+Z per annullare)")

    def _apply_area(self, x, y):
        """Bianchetto/oscura: elimina testo e immagini sotto il rettangolo e lo colora."""
        (x0, y0), self._area = self._area, None
        self.view.delete("area")
        page = self.doc[self.current]
        rect = pymupdf.Rect(self._view_to_page(x0, y0), self._view_to_page(x, y)).normalize() & page.rect
        if rect.is_empty or rect.width * self.cur_zoom < 3 or rect.height * self.cur_zoom < 3:
            return
        black = self.tool == "blackout"
        self._snapshot()
        erase(page, [(rect * page.derotation_matrix).normalize()], fill=(0, 0, 0) if black else (1, 1, 1),
              images=pymupdf.PDF_REDACT_IMAGE_PIXELS, graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED)
        self._doc_changed(pages=[self.current])
        self.status(f"Area {'oscurata' if black else 'cancellata'}: il contenuto sotto è eliminato davvero "
                    "(timbri e caselle di testo restano). Ctrl+Z per annullare.")

    # ------------------------------------------------------- ricerca
    def focus_find(self):
        self.find_entry.focus_set()
        self.find_entry.select_range(0, tk.END)

    def find(self, step=1):
        text = self.find_var.get().strip()
        if not text or not self.doc:
            self.clear_find(keep_text=True)
            return
        f = self._find
        if not f or f["text"] != text or f["doc"] is not self.doc:
            self.root.configure(cursor="watch")
            self.root.update_idletasks()
            hits = [(i, r) for i in range(len(self.doc)) for r in self.doc[i].search_for(text)]
            self.root.configure(cursor="")
            f = self._find = {"text": text, "doc": self.doc, "hits": hits, "pos": None}
        hits = f["hits"]
        if not hits:
            self.find_lbl.configure(text="Nessun risultato")
            self._draw_find_marks()
            return
        if f["pos"] is None:   # il primo risultato a partire dalla pagina mostrata
            if step > 0:
                f["pos"] = next((k for k, (i, _r) in enumerate(hits) if i >= self.current), 0)
            else:
                f["pos"] = max((k for k, (i, _r) in enumerate(hits) if i <= self.current), default=len(hits) - 1)
        else:
            f["pos"] = (f["pos"] + step) % len(hits)
        page_no, rect = hits[f["pos"]]
        self.find_lbl.configure(text=f"{f['pos'] + 1} di {len(hits)}")
        if page_no != self.current:
            self.go_to(page_no)
        else:
            self._draw_find_marks()
        self._scroll_to(rect * self.doc[page_no].rotation_matrix)

    def clear_find(self, keep_text=False):
        self._find = None
        self.find_lbl.configure(text="")
        self.view.delete("find")
        if not keep_text:
            self.find_var.set("")
            self.view.focus_set()

    def _draw_find_marks(self):
        self.view.delete("find")
        f = self._find
        if not f or f["doc"] is not self.doc or not self.doc:
            return
        page = self.doc[self.current]
        for k, (i, r) in enumerate(f["hits"]):
            if i == self.current:
                active = k == f["pos"]
                self.view.create_rectangle(*self._canvas_rect(r * page.rotation_matrix), tags="find",
                                           outline="#e8710a" if active else "#f4b400", width=2 if active else 1,
                                           fill="#ffd54f", stipple="gray25")

    def _scroll_to(self, rect):
        """Scorre la vista perché il rettangolo (pagina vista) sia visibile."""
        x0, y0, x1, y1 = self._canvas_rect(rect)
        sw, sh = self.scroll_size
        vw, vh = self.view.winfo_width(), self.view.winfo_height()
        left, top = self.view.canvasx(0), self.view.canvasy(0)
        if y0 < top or y1 > top + vh:
            self.view.yview_moveto(max(y0 - vh / 3, 0) / sh)
        if x0 < left or x1 > left + vw:
            self.view.xview_moveto(max(x0 - vw / 3, 0) / sw)

    # ------------------------------------------------------- OCR e Word
    def cmd_ocr(self, on_done=None, all_pages=False):
        """Riconosce il testo delle pagine scansionate; on_done si chiama a lavoro finito."""
        if not self.doc:
            return
        if not (TESSDATA_DIR / "ita.traineddata").exists():
            messagebox.showerror(APP_NAME, f"Mancano i file delle lingue per l'OCR:\n{TESSDATA_DIR}",
                                 parent=self.root)
            return
        if all_pages:
            pages, language, skip = list(range(len(self.doc))), OCR_LANGUAGES["Italiano e inglese"], True
        else:
            dialog = OcrDialog(self.root, self.current, sorted(self.selection), len(self.doc))
            self.root.wait_window(dialog.top)
            if not dialog.result:
                return
            pages, language, skip = dialog.result
        tab, was_dirty = self.tab, self.dirty
        self._snapshot()
        progress = ProgressWindow(self.root, "Riconoscimento del testo (OCR)", maximum=len(pages))
        done, stats = [], {"k": 0, "words": 0, "skipped": 0}

        def step():
            k = stats["k"]
            if progress.cancelled or k >= len(pages):
                return finish()
            page = tab.doc[pages[k]]
            progress.update(k, f"Pagina {pages[k] + 1}  ({k + 1} di {len(pages)})…")
            if skip and page.get_text("words"):
                stats["skipped"] += 1
            else:
                try:
                    stats["words"] += ocr_page(page, language)
                except Exception as exc:
                    progress.close()
                    messagebox.showerror(APP_NAME, f"OCR non riuscito a pagina {pages[k] + 1}:\n{exc}",
                                         parent=self.root)
                    progress.cancelled = True
                    return finish(closed=True)
                done.append(pages[k])
            stats["k"] += 1
            self.root.after(1, step)

        def finish(closed=False):
            if not closed:
                progress.close()
            if not done:   # niente di cambiato: via l'annulla appena aggiunto
                tab.undo_stack.pop()
                tab.dirty = was_dirty
            for i in done:
                tab.thumb_photos.pop(i, None)
                tab.ov_photos.pop(i, None)
            if tab is self.tab:
                self._doc_changed(pages=done)
            skipped = f", {stats['skipped']} saltate perché contengono già testo" if stats["skipped"] else ""
            if done or not stats["skipped"]:
                self.status(f"OCR: {len(done)} pagine lette, {stats['words']} parole riconosciute{skipped}.")
            else:
                messagebox.showinfo(APP_NAME, "Le pagine scelte contengono già testo: non c'era niente da "
                                    "riconoscere.\nPer rifarlo togli la spunta «Salta le pagine che contengono "
                                    "già testo».", parent=self.root)
            if on_done and not progress.cancelled:
                on_done()

        self.root.after(10, step)

    def cmd_export_word(self):
        if not self.doc:
            return
        if not word_installed():
            messagebox.showerror(APP_NAME, "Per esportare in Word serve Microsoft Word installato su questo PC.",
                                 parent=self.root)
            return
        if not any(page.get_text("words") for page in self.doc):
            answer = messagebox.askyesnocancel(APP_NAME, "Il PDF non contiene testo: è probabilmente una scansione "
                                               "e in Word arriverebbero solo immagini.\n\nVuoi prima riconoscere "
                                               "il testo con l'OCR?", parent=self.root)
            if answer is None:
                return
            if answer:
                self.cmd_ocr(on_done=self.cmd_export_word, all_pages=True)
                return
        stem = Path(self.path).stem if self.path else "documento"
        path = filedialog.asksaveasfilename(parent=self.root, title="Esporta in Word",
                                            initialdir=self._dir("word", str(Path(self.path).parent) if self.path
                                                                 else None),
                                            initialfile=stem + ".docx", defaultextension=".docx",
                                            filetypes=DOCX_TYPES)
        if not path:
            return
        # Word non accetta percorsi oltre 255 caratteri: si lavora in una cartella temporanea
        work = Path(tempfile.mkdtemp(prefix="tpw"))
        pdf, docx = work / "documento.pdf", work / "documento.docx"
        pdf.write_bytes(self.doc.tobytes(garbage=3, deflate=True))
        progress = ProgressWindow(self.root, "Esporta in Word", can_cancel=False)
        progress.update(0, "Conversione con Microsoft Word in corso…")
        result = {}
        worker = threading.Thread(target=lambda: result.update(error=pdf_to_docx(pdf, docx)), daemon=True)
        worker.start()

        def wait():
            if worker.is_alive():
                self.root.after(200, wait)
                return
            progress.close()
            error = result.get("error")
            if error is None:
                try:
                    shutil.move(str(docx), path)
                except OSError as exc:
                    error = f"Impossibile salvare il file:\n{path}\n\n{exc}\n\nForse è aperto in Word?"
            shutil.rmtree(work, ignore_errors=True)
            if error:
                messagebox.showerror(APP_NAME, f"Esportazione in Word non riuscita:\n\n{error}", parent=self.root)
                return
            self._remember_dir("word", Path(path).parent)
            self.status(f"Esportato in Word: {path}")
            if messagebox.askyesno(APP_NAME, f"Documento Word creato:\n{path}\n\nVuoi aprirlo?", parent=self.root):
                os.startfile(path)

        self.root.after(200, wait)

    # ------------------------------------------------------- elenco timbri
    def cmd_add_stamp(self):
        if not self._library_ready():
            return
        path = filedialog.askopenfilename(parent=self.root, title="Scegli l'immagine del timbro",
                                          initialdir=self._dir("image"), filetypes=IMAGE_TYPES)
        if not path:
            return
        self._remember_dir("image", Path(path).parent)
        name = simpledialog.askstring(APP_NAME, "Nome del timbro:", initialvalue=Path(path).stem, parent=self.root)
        if name is None:
            return
        name = name.strip() or Path(path).stem
        self.root.configure(cursor="watch")
        self.root.update_idletasks()
        try:
            self.library.add(path, name)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile caricare il timbro:\n{exc}", parent=self.root)
            return
        finally:
            self.root.configure(cursor="")
        self._stamp_selected()
        self.gallery.yview_moveto(1.0)
        self.status(f"Timbro «{name}» salvato: resterà disponibile anche alle prossime aperture.")

    def cmd_create_signature(self):
        if not self._library_ready():
            return
        dialog = SignatureDialog(self.root)
        self.root.wait_window(dialog.top)
        if not dialog.result:
            return
        pix, name, width_mm, source = dialog.result
        try:
            # la firma ha già lo sfondo trasparente: niente rimozione del bianco; source serve a rimodificarla
            self.library.add(pix, name, width_mm=width_mm, remove_white=False, source=source)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile salvare la firma:\n{exc}", parent=self.root)
            return
        self._stamp_selected()
        self.gallery.yview_moveto(1.0)
        self.status(f"Firma «{name}» salvata nella galleria: attiva la modalità timbro e clicca dove apporla.")

    def cmd_edit_stamp(self):
        """Modifica il timbro selezionato senza crearne uno nuovo: restano posto in galleria e impostazioni.

        Ogni timbro si può ruotare e ritoccare (matita, evidenziatore, gomma, testo, editor esterno);
        le firme create qui si possono anche riscrivere o ridisegnare.
        """
        item = self.stamp_item
        if not item or not self._library_ready():
            return
        StampEditor.last_image_dir = self._dir("image")
        editor = StampEditor(self.root, self.fonts, pymupdf.Pixmap(self.stamp.original_png), item["width_mm"],
                             f"Modifica timbro «{item['name']}»", name=item["name"], rotation=item["rotate"],
                             source=item.get("source"), remove_white=item["remove_white"],
                             tolerance=item["white_tolerance"])
        self.root.wait_window(editor.top)
        self._remember_dir("image", StampEditor.last_image_dir)
        if not editor.result:
            return
        try:
            self.library.apply_edit(item, editor.result)
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile aggiornare il timbro:\n{exc}", parent=self.root)
            return
        self._stamp_image_changed(item)
        self.status(f"Timbro «{item['name']}» modificato. I timbri già apposti nei PDF restano com'erano "
                    "(si modificano con il tasto destro sul timbro).")

    def rotate_stamp(self, angle):
        """Ruota il timbro selezionato della galleria (vale per i timbri apposti da ora in poi)."""
        item = self.stamp_item
        if not item or not self._library_ready():
            return
        self.library.rotate(item, angle)
        self._stamp_image_changed(item)
        self.status(f"«{item['name']}» ruotato {'a destra' if angle > 0 else 'a sinistra'}: vale per i timbri che "
                    "apponi da ora in poi.")

    def _stamp_image_changed(self, item):
        for tab in self.tabs:   # nei documenti aperti la nuova immagine va inserita da capo
            tab.stamp_xrefs.pop(item["id"], None)
        self._gallery_photos.pop(item["id"], None)
        self._stamp_selected()

    def _on_gallery_double(self, event):
        k = self._gallery_index(event)
        if k is None:
            if not self.library.available:
                self._library_ready()
            return
        self.select_stamp(k)
        self.cmd_edit_stamp()

    def cmd_rename_stamp(self):
        item = self.stamp_item
        if not item:
            return
        name = simpledialog.askstring(APP_NAME, "Nuovo nome del timbro:", initialvalue=item["name"],
                                      parent=self.root)
        if name and name.strip():
            item["name"] = name.strip()
            self.library.save()
            self._draw_gallery()

    def cmd_delete_stamp(self):
        item = self.stamp_item
        if not item:
            return
        if not messagebox.askyesno(APP_NAME, f"Eliminare il timbro «{item['name']}» dall'elenco?\n\n"
                                   "I PDF già timbrati non cambiano.", parent=self.root):
            return
        self.library.remove(item)
        for tab in self.tabs:
            tab.stamp_xrefs.pop(item["id"], None)
        self._gallery_photos.pop(item["id"], None)
        if not self.stamp:
            self.set_stamp_mode(False)
        self._stamp_selected()

    def select_stamp(self, index):
        if not 0 <= index < len(self.library.items) or self.library.items[index] is self.stamp_item:
            return
        self.library.current = self.library.items[index]
        self.library.save()
        self._stamp_selected()
        self.status(f"Timbro selezionato: {self.stamp_item['name']}")

    def _stamp_selected(self):
        self._load_stamp_settings()
        self._refresh_stamp_panel()
        self._update_controls()

    def _load_stamp_settings(self):
        """Mostra nei controlli le impostazioni salvate del timbro selezionato."""
        item = self.stamp_item or StampLibrary.DEFAULTS
        self._loading_settings = True
        self.remove_white_var.set(item["remove_white"])
        self.tolerance_scale.set(item["white_tolerance"])
        self.tolerance_scale.state(["!disabled"] if item["remove_white"] else ["disabled"])
        self.stamp_w_var.set(f"{item['width_mm']:g}")
        self.overlay_var.set(item["overlay"])
        self._paint_overlay_note()
        self._loading_settings = False

    def _tile_origin(self, index):
        cols = self._gallery_cols
        return 5 + index % cols * (TILE_W + 6), 5 + index // cols * (TILE_H + 26)

    def _on_gallery_resize(self, event):
        """Con il pannello più largo la galleria mostra più timbri per riga."""
        cols = max(2, (event.width - 6) // (TILE_W + 6))
        if cols != self._gallery_cols:
            self._gallery_cols = cols
            self._draw_gallery()

    def _draw_gallery(self):
        if self.library is None:
            return
        c = self.gallery
        c.delete("all")
        items = self.library.items
        width = max(c.winfo_width(), PREVIEW_W)
        if not items:
            text = ("Nessun timbro salvato.\nUsa «Aggiungi…» o «Crea firma…»." if self.library.available else
                    "Archivio dei timbri non aperto.\nDoppio clic qui per aprirlo.")
            c.create_text(width / 2, GALLERY_H / 2, justify=tk.CENTER, fill="#888", text=text)
            c.configure(scrollregion=(0, 0, width, GALLERY_H))
            return
        for k, item in enumerate(items):
            x0, y0 = self._tile_origin(k)
            selected = item is self.stamp_item
            c.create_rectangle(x0 - 2, y0 - 2, x0 + TILE_W + 2, y0 + TILE_H + 20, width=2 if selected else 1,
                               fill="#d6e4fb" if selected else "white", outline="#1a73e8" if selected else "#d0d0d8")
            self._draw_checker(c, x0 + 1, y0 + 1, x0 + TILE_W - 1, y0 + TILE_H - 1, 8)
            if item["id"] not in self._gallery_photos:
                self._gallery_photos[item["id"]] = self.library.stamps[item["id"]].fit_photo(TILE_W - 8, TILE_H - 8)
            c.create_image(x0 + TILE_W / 2, y0 + TILE_H / 2, image=self._gallery_photos[item["id"]])
            name = item["name"] if len(item["name"]) <= 13 else item["name"][:12] + "…"
            c.create_text(x0 + TILE_W / 2, y0 + TILE_H + 9, text=f"{k + 1}. {name}" if k < 9 else name,
                          font=("Segoe UI", 8), fill="#222")
        height = 6 + -(-len(items) // self._gallery_cols) * (TILE_H + 26)
        c.configure(scrollregion=(0, 0, width, max(GALLERY_H, height)))

    _draw_checker = staticmethod(draw_checker)

    def _gallery_index(self, event):
        x, y = self.gallery.canvasx(event.x), self.gallery.canvasy(event.y)
        for k in range(len(self.library.items)):
            x0, y0 = self._tile_origin(k)
            if x0 - 2 <= x <= x0 + TILE_W + 2 and y0 - 2 <= y <= y0 + TILE_H + 20:
                return k
        return None

    def _on_gallery_click(self, event):
        k = self._gallery_index(event)
        if k is not None:
            self.select_stamp(k)

    def _on_gallery_menu(self, event):
        k = self._gallery_index(event)
        if k is None:
            return
        self.select_stamp(k)
        menu = tk.Menu(self.root, tearoff=False)
        menu.add_command(label="Modifica… (ruota, matita, gomma, testo)", command=self.cmd_edit_stamp)
        menu.add_command(label="Ruota a destra", command=lambda: self.rotate_stamp(90))
        menu.add_command(label="Ruota a sinistra", command=lambda: self.rotate_stamp(-90))
        menu.add_separator()
        menu.add_command(label="Rinomina…", command=self.cmd_rename_stamp)
        menu.add_command(label="Elimina", command=self.cmd_delete_stamp)
        menu.tk_popup(event.x_root, event.y_root)

    def _on_gallery_wheel(self, event):
        self.gallery.yview_scroll(-3 if event.delta > 0 else 3, "units")
        return "break"

    def _on_panel_wheel(self, event):
        if str(event.widget).startswith(str(self.panel)) and self.panel.yview() != (0.0, 1.0):
            self.panel.yview_scroll(-3 if event.delta > 0 else 3, "units")

    # ------------------------------------------------------- impostazioni del timbro
    def _on_white_settings(self, delay=0):
        if self._loading_settings:
            return
        if self._white_job:
            self.root.after_cancel(self._white_job)
        self._white_job = self.root.after(delay, self._apply_white_settings)

    def _apply_white_settings(self):
        self._white_job = None
        remove, tolerance = self.remove_white_var.get(), round(self.tolerance_scale.get())
        self.tolerance_scale.state(["!disabled"] if remove else ["disabled"])
        item = self.stamp_item
        if not item:
            return
        item.update(remove_white=remove, white_tolerance=tolerance)
        self.library.reprocess(item)
        for tab in self.tabs:   # l'immagine cambia: nei documenti aperti va reinserita
            tab.stamp_xrefs.pop(item["id"], None)
        self._gallery_photos.pop(item["id"], None)
        self._refresh_stamp_panel()

    def _on_overlay(self):
        if self.stamp_item:
            self.stamp_item["overlay"] = self.overlay_var.get()
            self.library.save()
        self._paint_overlay_note()

    def _paint_overlay_note(self):
        self.overlay_note.configure(text="" if self.overlay_var.get() else
                                    "Nelle pagine scansionate il timbro in secondo piano resta nascosto "
                                    "sotto l'immagine.")

    def _refresh_stamp_panel(self):
        self._draw_gallery()
        c = self.stamp_preview
        c.delete("all")
        self._draw_checker(c, 0, 0, PREVIEW_W + 2, PREVIEW_H + 2)
        if not self.stamp:
            c.create_text(PREVIEW_W / 2, PREVIEW_H / 2, text="Nessun timbro selezionato", fill="#888")
            self.stamp_info.configure(text="")
            self._redraw_ghost()
            return
        self._preview_photo = self.stamp.fit_photo(PREVIEW_W - 16, PREVIEW_H - 16)
        c.create_image(PREVIEW_W / 2 + 1, PREVIEW_H / 2 + 1, image=self._preview_photo)
        self._update_stamp_info()
        self._redraw_ghost()

    def _update_stamp_info(self):
        if self.stamp:
            w_mm = self.stamp_item["width_mm"]
            dpi = self.stamp.pixmap.width / (w_mm / 25.4)
            self.stamp_info.configure(text=f"Dimensione: {w_mm:g} × {w_mm * self.stamp.aspect:.0f} mm  ·  "
                                           f"{dpi:.0f} dpi")

    def _on_scale(self, value):
        if self._syncing:
            return
        self._syncing = True
        self.stamp_w_var.set(str(round(float(value))))
        self._syncing = False

    def _on_stamp_width(self, *_):
        try:
            w = float(self.stamp_w_var.get().replace(",", ".").strip())
        except ValueError:
            return
        if not 3 <= w <= 600:
            return
        if not self._syncing:
            self._syncing = True
            self.stamp_scale.set(min(max(w, 5), 150))
            self._syncing = False
        item = self.stamp_item
        if not item:
            return
        if item["width_mm"] != w:
            item["width_mm"] = w
            if self._cfg_job:
                self.root.after_cancel(self._cfg_job)
            self._cfg_job = self.root.after(500, self.library.save)
        self._update_stamp_info()
        self._redraw_ghost()

    # ------------------------------------------------------- varie
    # ------------------------------------------------------- archivio dei timbri
    def _ask_archive_password(self, path, wrong):
        message = (("Password errata.\n\n" if wrong else "") +
                   f"L'archivio dei timbri è protetto da password:\n{path}")
        dialog = PasswordDialog(self.root, message)
        self.root.wait_window(dialog.top)
        return dialog.result

    def _archive_error(self, message):
        messagebox.showerror(APP_NAME, message, parent=self.root)

    def _library_ready(self):
        """True se l'archivio dei timbri è aperto; altrimenti prova ad aprirlo (chiedendo la password)."""
        if self.library.available:
            return True
        self.library.load()
        self._library_reloaded()
        if not self.library.available:
            messagebox.showwarning(APP_NAME, self.library.error, parent=self.root)
        return self.library.available

    def _library_reloaded(self):
        for tab in self.tabs:   # le immagini possono essere cambiate
            tab.stamp_xrefs.clear()
        self._gallery_photos.clear()
        self._stamp_selected()

    def _on_focus_in(self, _event):
        """Tornando alla finestra: se un altro PC ha cambiato l'archivio condiviso, si rilegge."""
        now = time.monotonic()
        if self.library is None or now - self._archive_checked < 2:
            return
        self._archive_checked = now
        if self.library.changed_on_disk():
            self.library.load()
            self._library_reloaded()
            self.status("Timbri aggiornati: l'archivio condiviso è stato modificato da un altro PC.")

    def cmd_archive(self):
        lib = self.library
        dialog = ArchiveDialog(self.root, lib.path, len(lib.items), bool(lib.password), lib.available)
        self.root.wait_window(dialog.top)
        action = dialog.action
        if action == "open_folder":
            lib.path.parent.mkdir(parents=True, exist_ok=True)
            os.startfile(lib.path.parent)
        elif action == "unlock":
            self._library_ready()
        elif action == "move":
            self._archive_move()
        elif action == "use":
            self._archive_use()
        elif action == "password":
            if not self._library_ready():
                return
            dialog = PasswordDialog(self.root, "Scegli la password dell'archivio dei timbri.\nSenza di essa "
                                    "l'archivio non si apre su nessun PC: annotala in un posto sicuro.", confirm=True)
            self.root.wait_window(dialog.top)
            if dialog.result:
                lib.set_password(*dialog.result)
                self.status("Archivio dei timbri protetto da password (cifrato AES-256).")
        elif action == "unprotect":
            if messagebox.askyesno(APP_NAME, "Togliere la password? L'archivio resterà leggibile da chiunque abbia "
                                   "il file.", parent=self.root):
                lib.set_password(None)
                self.status("Password dell'archivio dei timbri tolta.")

    def _archive_move(self):
        if not self._library_ready():
            return
        folder = filedialog.askdirectory(parent=self.root, initialdir=str(self.library.path.parent),
                                         title="Cartella in cui tenere l'archivio dei timbri (es. su Dropbox)")
        if not folder:
            return
        target = Path(folder) / ARCHIVE_NAME
        if target.exists() and target.resolve() != self.library.path.resolve():
            messagebox.showinfo(APP_NAME, "In quella cartella c'è già un archivio dei timbri.\n"
                                "Per usarlo scegli «Usa un archivio esistente».", parent=self.root)
            return
        try:
            self.library.move_to(target)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Impossibile spostare l'archivio:\n{exc}", parent=self.root)
            return
        self.status(f"Archivio dei timbri spostato in {target}. Sugli altri PC: Timbro ▸ Archivio dei timbri "
                    "▸ Usa un archivio esistente.")

    def _archive_use(self):
        path = filedialog.askopenfilename(parent=self.root, title="Archivio dei timbri da usare",
                                          initialdir=str(self.library.path.parent), filetypes=ARCHIVE_TYPES)
        if not path or Path(path).resolve() == self.library.path.resolve():
            return
        merge = False
        if self.library.available and self.library.items:
            merge = messagebox.askyesnocancel(
                APP_NAME, f"Su questo PC ci sono {len(self.library.items)} timbri. Vuoi aggiungerli all'archivio "
                          "scelto (quelli che non ci sono già)?\n\nNo: si usano solo i timbri dell'archivio scelto; "
                          "quelli di questo PC restano nel file di prima.", parent=self.root)
            if merge is None:
                return
        if self.library.use_archive(path, merge=merge):
            self.status(f"Ora i timbri vengono dall'archivio {path}.")
        else:
            messagebox.showwarning(APP_NAME, self.library.error, parent=self.root)
        self._library_reloaded()

    def show_help(self):
        messagebox.showinfo(f"{APP_NAME} — Guida rapida", HELP_TEXT, parent=self.root)

    def show_about(self):
        dialog = AboutDialog(self.root)
        self.root.wait_window(dialog.top)

    def _check_license(self):
        """Al primo avvio (e quando cambiano) le condizioni d'uso vanno lette e accettate."""
        if (self.cfg["license_accepted"] or 0) >= LICENSE_VERSION:
            return
        dialog = AboutDialog(self.root, must_accept=True)
        self.root.wait_window(dialog.top)
        if dialog.accepted:
            self.cfg["license_accepted"] = LICENSE_VERSION
        else:
            self.root.destroy()

    def open_guide(self):
        """La guida illustrata è un PDF: si apre in una scheda di TimbroPDF stesso."""
        guide = BASE_DIR / GUIDE_NAME
        if not guide.exists():
            messagebox.showinfo(APP_NAME, f"La guida illustrata non si trova:\n{guide}", parent=self.root)
            return
        try:   # copia di sola lettura: così il file allegato al programma non viene modificato per sbaglio
            doc = pymupdf.open(guide)
            tab = DocTab(pymupdf.open(stream=doc.tobytes(), filetype="pdf"), label="Guida")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossibile aprire la guida:\n{exc}", parent=self.root)
            return
        self._add_tab(tab)
        self.set_zoom_mode("width")

    def _poll_instance(self):
        """File aperti con doppio clic (da Esplora risorse) mentre il programma era già aperto."""
        while self.instance.pending:
            files = self.instance.pending.pop(0)
            for f in files:
                self.open_file(f)
            root = self.root
            root.deiconify()
            if root.state() == "iconic":
                root.state("normal")
            root.lift()
            root.attributes("-topmost", True)   # in primo piano, poi torna normale
            root.after(300, lambda: root.attributes("-topmost", False))
            root.focus_force()
        self.root.after(300, self._poll_instance)

    def cmd_default_app(self):
        if is_default_pdf_handler():
            messagebox.showinfo(APP_NAME, f"{APP_NAME} è già il programma predefinito per aprire i PDF.",
                                parent=self.root)
            return
        if not messagebox.askyesno(APP_NAME, "Windows non permette ai programmi di diventare predefiniti da "
                                   f"soli: {APP_NAME} si registra fra i programmi che aprono i PDF e poi si apre "
                                   "la pagina di Windows «App predefinite», dove basta sceglierlo per «.pdf».\n\n"
                                   "Procedere?", parent=self.root):
            return
        try:
            register_pdf_handler()
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Registrazione non riuscita:\n{exc}", parent=self.root)
            return
        try:
            os.startfile(f"ms-settings:defaultapps?registeredAppUser={APP_NAME}")
        except OSError:
            os.startfile("ms-settings:defaultapps")
        messagebox.showinfo(APP_NAME, f"Nella finestra di Windows clicca su «.pdf» e scegli {APP_NAME}.\n\n"
                            f"In alternativa: tasto destro su un PDF ▸ Apri con ▸ Scegli un'altra app ▸ {APP_NAME}, "
                            "con «Usa sempre questa app».\n\nSe sposti il file del programma, ripeti questo comando. "
                            "Facendo doppio clic su altri PDF, si aprono come nuove schede nella finestra già aperta.",
                            parent=self.root)

    def cmd_remove_default_app(self):
        if not messagebox.askyesno(APP_NAME, f"Togliere {APP_NAME} dai programmi che aprono i PDF?",
                                   parent=self.root):
            return
        try:
            unregister_pdf_handler()
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Operazione non riuscita:\n{exc}", parent=self.root)
            return
        self.status(f"{APP_NAME} tolto dai programmi per i PDF: scegli in Windows quale usare al suo posto.")

    def _on_files_dropped(self, files):
        """File trascinati da Esplora risorse sulla finestra: i PDF e le immagini si aprono in nuove schede."""
        accepted = (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")
        others = [Path(f).name for f in files if Path(f).suffix.lower() not in accepted]
        for f in files:
            if Path(f).suffix.lower() in accepted:
                self.open_file(f)
        self.root.lift()
        if others:
            messagebox.showinfo(APP_NAME, "Questi file non sono PDF né immagini e non sono stati aperti:\n"
                                + "\n".join(others), parent=self.root)


def main():
    # già aperto? gli si passano i file (doppio clic su un PDF) e questa apertura finisce qui
    if send_to_running_instance(sys.argv[1:]):
        return
    try:   # segnale per l'installatore: il programma è aperto, va chiuso prima di aggiornarlo
        import ctypes
        ctypes.windll.kernel32.CreateMutexW(None, False, APP_MUTEX)
    except Exception:
        pass
    set_dpi_awareness()
    root = tk.Tk()
    App(root, sys.argv[1:])
    root.mainloop()


if __name__ == "__main__":
    main()
