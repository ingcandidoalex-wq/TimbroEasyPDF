# TimbroEasyPDF

Programma **gratuito** per Windows per visualizzare, timbrare, firmare e modificare i PDF.
Ideato dall'ing. Alessandro Candido con l'ausilio dell'intelligenza artificiale.

## Scaricare

Dalla pagina delle versioni, [**Release**](https://github.com/ingcandidoalex-wq/TimbroEasyPDF/releases/latest),
scarica uno di questi file:

- `Setup TimbroEasyPDF <versione>.exe`: l'installatore, con menu Start, disinstallazione e, se vuoi,
  apertura dei PDF con doppio clic;
- `TimbroEasyPDF.exe`: un solo file, senza installazione. Basta copiarlo e avviarlo.

Funziona su Windows 10 e 11 a 64 bit. Il programma controlla da solo, una volta al giorno, se è uscita
una versione nuova.

> **«Windows ha protetto il PC»?** Al primo avvio Windows può mostrare questo avviso, perché il programma
> è nuovo e Microsoft non lo conosce ancora. Clicca **Ulteriori informazioni** e poi **Esegui comunque**.

## Che cosa fa

- **Timbri e firme**: galleria di timbri da immagine e di firme disegnate o scritte con un font. Si
  appongono con un clic, anche su tutte le pagine in un colpo solo. Si ruotano dei gradi che vuoi e si
  ritoccano con matita, evidenziatore, gomma e testo. Quelli già sulla pagina si spostano, si ruotano e
  si modificano.
- **Pagine e documenti**: più PDF aperti in schede. Le pagine si trascinano da un documento all'altro,
  si ruotano, si riordinano e si eliminano. C'è anche la vista d'insieme delle miniature di tutti i PDF.
- **Testo**: correzione delle righe del PDF e caselle di testo nuove, con l'anteprima mentre scrivi.
  Bianchetto e oscuramento cancellano davvero quello che c'è sotto.
- **OCR**: riconosce il testo delle scansioni (in italiano e in inglese), che poi si può cercare,
  copiare ed esportare.
- **Esportazione in Word**, se Microsoft Word è installato.
- **Unire e dividere**: si scelgono i PDF da unire e il loro ordine, oppure si divide un documento in
  file da 1, 2 o N pagine con un numero progressivo.
- **Archivio dei timbri** in un solo file, che si può tenere in una cartella condivisa (es. Dropbox) per
  usarlo su più PC e proteggere con una password.

La guida illustrata è inclusa: menu **?** ▸ **Guida illustrata**.

## Suggerimenti e segnalazioni

Scrivi a **ing.candido.alex@gmail.com**, oppure dal programma: menu **?** ▸ **Scrivi all'autore**.
Se segnali un problema con un PDF, allegalo solo se non contiene dati riservati.

## Sostieni il progetto

Se il programma ti è utile puoi fare una donazione con PayPal all'indirizzo **a75lex@libero.it**.

## Riservatezza

Il programma lavora sul tuo PC e non invia i documenti a nessuno. L'unico collegamento a Internet è il
controllo giornaliero dell'ultima versione pubblicata su GitHub, che si disattiva dal menu **?**.

## Dal codice sorgente

Serve Python 3.14 con PyMuPDF e PyInstaller:

```
python -m pip install pymupdf pyinstaller
python timbro_pdf.py                  avvia il programma
python test\esegui_test.py            esegue i test
python crea_guida.py                  rifà la guida illustrata
python build_exe.py                   crea exe e installatore in dist (l'installatore richiede Inno Setup 6)
```

## Licenza

Copyright © 2026 Alessandro Candido.

TimbroEasyPDF è software libero: puoi ridistribuirlo e modificarlo alle condizioni della
**GNU Affero General Public License, versione 3 o successiva**, il cui testo completo è nel file
[LICENSE](LICENSE). È distribuito **senza alcuna garanzia**.

Usa questi componenti, ciascuno con la sua licenza:

- [PyMuPDF / MuPDF](https://pymupdf.readthedocs.io/) di Artifex Software: AGPL-3.0;
- i dati delle lingue di [Tesseract OCR](https://github.com/tesseract-ocr/tessdata_fast): Apache License 2.0;
- Python: Python Software Foundation License;
- Tcl/Tk: licenza di tipo BSD.
