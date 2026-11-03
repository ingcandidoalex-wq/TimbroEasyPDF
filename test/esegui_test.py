"""Esegue i test (test_*.py) uno dopo l'altro e ne riassume l'esito.

Uso:  python test\\esegui_test.py               (tutti)
      python test\\esegui_test.py layout v2     (solo quelli che contengono queste parole nel nome)

I test aprono finestre vere del programma e usano una configurazione a parte (vedi comune.py):
i timbri e le impostazioni reali non vengono toccati. test_nuove prova anche l'esportazione in Word,
che serve installato. Il resoconto completo di ogni test resta in <temp>\\TimbroEasyPDF-test\\<nome>.log
"""
import os
import subprocess
import sys
import time
from pathlib import Path

from comune import WORK

HERE = Path(__file__).parent
LIMITE = 600   # secondi: oltre, il test si considera bloccato


def run(test):
    t0 = time.time()
    log = WORK / f"{test.stem}.log"
    try:
        r = subprocess.run([sys.executable, "-X", "faulthandler", str(test)], cwd=HERE, timeout=LIMITE,
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        out, ok = r.stdout + r.stderr, r.returncode == 0 and "TUTTI I TEST OK" in r.stdout
    except subprocess.TimeoutExpired as e:
        out, ok = f"{e.stdout or ''}{e.stderr or ''}\nBLOCCATO dopo {LIMITE} s", False
    log.write_text(out, encoding="utf-8")
    print(f"{'ok     ' if ok else 'FALLITO'} {test.stem:12} {time.time() - t0:5.0f} s")
    if not ok:
        print("        " + "\n        ".join(out.strip().splitlines()[-25:]))
    return ok


if __name__ == "__main__":
    words = sys.argv[1:]
    tests = [t for t in sorted(HERE.glob("test_*.py")) if not words or any(w in t.stem for w in words)]
    results = [run(t) for t in tests]
    print(f"\n{sum(results)} su {len(results)} riusciti; resoconti in {WORK}")
    sys.exit(0 if all(results) else 1)
