"""Genera lo script di Inno Setup come build_exe.py, senza compilarlo, e ne mostra le parti sulla firma
e sul certificato. Serve a controllare il setup senza costruire niente.

Uso:  python test\\mostra_setup.py            (setup da pubblicare)
      python test\\mostra_setup.py --studio   (setup per i PC dello studio, con il certificato)
"""
import sys
import tempfile
import types
from pathlib import Path

import comune  # noqa: F401  (rende importabile il programma)
import build_exe as b

captured = {}
b.subprocess = types.SimpleNamespace(run=lambda cmd, check: captured.setdefault("cmd", cmd))
b.shutil = types.SimpleNamespace(copy=lambda *a: None)
with tempfile.TemporaryDirectory() as work:
    b.build_setup("ISCC.exe", Path(work) / "app", b.HERE / "timbro.ico", work, studio="--studio" in sys.argv)
    text = (Path(work) / "installer.iss").read_text(encoding="utf-8-sig")
    cer = [c.name for c in Path(work).glob("*.cer")]
print("opzioni di ISCC:", captured["cmd"][1:-1])
print("certificato esportato:", cer)
for line in text.splitlines():
    if any(k in line for k in ("#define", "SignTool", "certificato", "certutil", "Cert", "[UninstallRun]", "[Code]",
                               "OutputBaseFilename", "LICENSE")):
        print(line if "CertFile \"" not in line else line.split("\"")[0] + '"…"')
left = sorted({w for w in text.split("@")[1::2] if w.isupper()})
print("segnaposto rimasti:", left or "nessuno")
