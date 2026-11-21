"""Test dello sblocco a firma digitale: codice macchina, verifica del token (valido/manomesso/scaduto/altro pc)
e promemoria saltato quando il programma è sbloccato. Usa una coppia di chiavi usa-e-getta al posto di quella vera."""
import base64
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from comune import WORK  # noqa: E402
import timbro_pdf as tp  # noqa: E402
from cryptography.hazmat.primitives import serialization  # noqa: E402
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey  # noqa: E402

tp.LICENSE_VERSION = 0
CFG = WORK / "cfg_licenza"
shutil.rmtree(CFG, ignore_errors=True)
CFG.mkdir()
tp.CONFIG_DIR, tp.CONFIG_FILE = CFG, CFG / "config.json"
tp.STAMPS_DIR, tp.OLD_STAMP_FILE = CFG / "timbri", CFG / "timbro.png"
tp.set_dpi_awareness()

# coppia di chiavi usa-e-getta: sostituisce quella vera (LICENSE_PUBKEY) per i test
_priv = Ed25519PrivateKey.generate()
_pub = _priv.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw)
tp.LICENSE_PUBKEY = base64.b64encode(_pub).decode("ascii")


def firma(codice, scadenza):
    f = _priv.sign(tp._token_payload(codice, scadenza))
    return base64.urlsafe_b64encode(scadenza.encode("ascii") + f).decode("ascii")


def fra_giorni(n):
    return time.strftime("%Y%m%d", time.localtime(time.time() + n * 86400))


# ---------------------------------------------------------------- 1. codice macchina: la trasformazione
tp.hardware_fingerprint = lambda: ("AABBCC", "SN12")
atteso = "".join(str(ord(c)) for c in ("AABBCCSN12" + "AABB"))   # base + prime 4 in fondo, poi ASCII decimale
assert tp.codice_macchina() == atteso, tp.codice_macchina()
print("1. codice macchina (prime 4 accodate + ASCII decimale) OK")

# ---------------------------------------------------------------- 2. verifica del token
code = "123456789"
tok = firma(code, fra_giorni(365))
assert tp.verifica_token(code, tok), "il token valido deve passare"
assert not tp.verifica_token(code, tok[:-3] + "AAA"), "un token manomesso deve fallire"
assert not tp.verifica_token(code, firma(code, fra_giorni(-1))), "un token scaduto deve fallire"
assert not tp.verifica_token(code + "0", tok), "un token di un altro pc deve fallire"
assert not tp.verifica_token(code, ""), "un token vuoto deve fallire"
print("2. verifica token (valido/manomesso/scaduto/altro pc/vuoto) OK")

# ---------------------------------------------------------------- 3. app: sblocco e promemoria
tp.codice_macchina = lambda: "999888777"
root = tk.Tk()
root.geometry("400x300+60+60")
app = tp.App(root)
assert app.is_unlocked() is False, "senza password non deve risultare sbloccato"
app.cfg["license_key"] = firma("999888777", fra_giorni(365))
assert app.is_unlocked() is True, "con una password valida deve risultare sbloccato"
app._nag_job = "x"
app._show_nag()                              # da sbloccato non deve aprire nulla né ripianificare
assert app._nag_job is None, "da sbloccato il promemoria non deve ripianificarsi"
app.cfg["license_key"] = firma("999888777", fra_giorni(-1))   # scaduta
app._codice_cache = None
assert app.is_unlocked() is False, "una password scaduta non deve più sbloccare"
root.destroy()
print("3. app: sblocco valido, promemoria saltato, scadenza riporta il blocco OK")

print("TUTTI I TEST OK")
