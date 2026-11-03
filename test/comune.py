"""Percorsi comuni ai test: il programma è nella cartella sopra questa; configurazioni e file di prova
vanno in una cartella temporanea (non nel progetto, che è sincronizzato da Dropbox), mai in quella vera."""
import sys
import tempfile
from pathlib import Path

PROGRAMMA = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROGRAMMA))
WORK = Path(tempfile.gettempdir()) / "TimbroEasyPDF-test"
WORK.mkdir(exist_ok=True)

import timbro_pdf  # noqa: E402

timbro_pdf.UPDATE_URL = ""   # durante i test il programma non controlla gli aggiornamenti su Internet
