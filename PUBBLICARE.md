# Pubblicare una nuova versione

Il programma cerca gli aggiornamenti fra le **Release** di questo repository: legge il numero
dell'etichetta dell'ultima versione pubblicata (es. `v2.3`) e propone di aprirne la pagina. Per
pubblicare una versione:

1. Alza `APP_VERSION` in `timbro_pdf.py` (es. `"2.3"`). Se cambiano le condizioni d'uso, alza anche
   `LICENSE_VERSION`: tutti dovranno riaccettarle.
2. Esegui i test: `python test\esegui_test.py`.
3. Rifai la guida: `python crea_guida.py`.
4. Costruisci: `python build_exe.py`. In `dist` trovi `TimbroEasyPDF.exe`, `Setup TimbroEasyPDF 2.3.exe`
   e la guida. Il setup «(studio)», che si crea solo con `--studio`, **non va mai pubblicato**: installa
   il certificato autofirmato.
5. Registra e carica il codice: `git commit`, poi `git tag v2.3` e `git push --follow-tags`.
6. Su GitHub apri **Releases** ▸ **Draft a new release**:
   - scegli l'etichetta `v2.3` e come titolo «TimbroEasyPDF 2.3»;
   - scrivi le novità in poche righe: il programma le mostra a chi aggiorna;
   - allega `Setup TimbroEasyPDF 2.3.exe`, `TimbroEasyPDF.exe` e `Guida TimbroEasyPDF.pdf` dalla
     cartella `dist`;
   - premi **Publish release**.
7. Pubblica sulla pagina Facebook il link
   `https://github.com/ingcandidoalex-wq/TimbroEasyPDF/releases/latest`, che porta sempre all'ultima
   versione. Conviene non allegare l'exe al post: spesso Facebook lo blocca.
