"""Genera la password di sblocco di TimbroEasyPDF per un dato codice di computer.

Solo per l'autore: serve la chiave privata (chiave_licenza_privata.pem), che NON sta nel repository.

Uso:
    python genera_sblocco.py <codice del computer>     password valida un anno
    python genera_sblocco.py <codice> --giorni 730     durata personalizzata (in giorni)
    python genera_sblocco.py                           usa il codice di QUESTO pc (per prova)

La password stampata va rispedita via email al cliente: lui la incolla nel programma e i messaggi
spariscono fino alla scadenza. La chiave privata va custodita e mai pubblicata: chi ce l'ha può
generare sblocchi per qualsiasi computer. Se la perdi, non puoi più generare password.
"""
import argparse
import base64
import sys
import time
from pathlib import Path

from cryptography.hazmat.primitives import serialization

from timbro_pdf import APP_NAME, LICENSE_GIORNI, _token_payload, codice_macchina, verifica_token

HERE = Path(__file__).parent
CHIAVE = HERE / "chiave_licenza_privata.pem"


def carica_chiave(percorso):
    try:
        return serialization.load_pem_private_key(Path(percorso).read_bytes(), password=None)
    except FileNotFoundError:
        sys.exit(f"Chiave privata non trovata: {percorso}\n"
                 "È il file creato insieme alla chiave pubblica del programma: tienilo da parte e non pubblicarlo.")


def genera(codice, giorni, chiave):
    """(token, scadenza AAAAMMGG): firma «codice|scadenza» con la chiave privata e lo impacchetta in base64."""
    scadenza = time.strftime("%Y%m%d", time.localtime(time.time() + giorni * 86400))
    firma = chiave.sign(_token_payload(codice, scadenza))
    token = base64.urlsafe_b64encode(scadenza.encode("ascii") + firma).decode("ascii")
    return token, scadenza


def main():
    ap = argparse.ArgumentParser(description=f"Password di sblocco di {APP_NAME}.")
    ap.add_argument("codice", nargs="?", help="il codice del computer del cliente (se manca, usa questo pc)")
    ap.add_argument("--giorni", type=int, default=LICENSE_GIORNI,
                    help=f"durata dello sblocco in giorni (predefinito {LICENSE_GIORNI})")
    ap.add_argument("--chiave", default=str(CHIAVE), help="percorso della chiave privata")
    args = ap.parse_args()
    codice = (args.codice or codice_macchina()).strip()
    chiave = carica_chiave(args.chiave)
    token, scadenza = genera(codice, args.giorni, chiave)
    print("Codice:   ", codice)
    print(f"Scadenza:  {scadenza[6:8]}/{scadenza[4:6]}/{scadenza[:4]}")
    print("Password di sblocco (da inviare al cliente):")
    print(token)
    if not verifica_token(codice, token):   # coerenza con la chiave pubblica inclusa nel programma
        print("\nATTENZIONE: la password non supera la verifica con la chiave pubblica del programma: "
              "la chiave privata non corrisponde a LICENSE_PUBKEY.")


if __name__ == "__main__":
    main()
