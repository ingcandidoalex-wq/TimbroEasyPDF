"""Firma digitale (Authenticode) dei programmi creati da build_exe.py.

Usa il certificato per la firma del codice intestato a FIRMATARIO che si trova nell'archivio personale
di Windows (certmgr.msc > Personale > Certificati). Oggi è un certificato autofirmato, attendibile solo
sui PC dove lo si installa (il setup lo fa, se si sceglie l'opzione); se un giorno se ne aggiunge uno
rilasciato da un ente (SSL.com, Actalis...) con lo stesso nome, si usa quello.
Su un PC senza certificato i programmi si creano lo stesso, senza firma.

Uso:  python firma_codice.py <file.exe> [...]     (lo richiama anche Inno Setup per setup e disinstallazione)
"""

import os
import subprocess
import sys
import time

FIRMATARIO = "Alessandro Candido"
# la marca temporale tiene valida la firma anche dopo la scadenza del certificato
TIMESTAMP_SERVERS = ("http://timestamp.digicert.com", "http://timestamp.sectigo.com")

# il certificato più adatto: valido, per la firma del codice, con chiave privata; quello di un ente prima
# di quello autofirmato, poi il più recente
_FIND = r"""
$c = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
     Where-Object { $_.GetNameInfo('SimpleName', $false) -eq $env:FIRMA_NOME -and $_.NotAfter -gt (Get-Date) } |
     Sort-Object @{ Expression = { $_.Issuer -ne $_.Subject } }, NotAfter | Select-Object -Last 1
"""

_SIGN = _FIND + r"""
if (-not $c) { exit 2 }
foreach ($ts in @($env:FIRMA_TS -split ' ') + '') {
    $p = @{ FilePath = $env:FIRMA_FILE; Certificate = $c; HashAlgorithm = 'SHA256'; IncludeChain = 'All' }
    if ($ts) { $p.TimestampServer = $ts }
    $null = Set-AuthenticodeSignature @p -ErrorAction SilentlyContinue
    $s = Get-AuthenticodeSignature $env:FIRMA_FILE
    if ($s.SignerCertificate.Thumbprint -eq $c.Thumbprint -and ($s.TimeStamperCertificate -or -not $ts)) {
        if (-not $ts) { 'senza marca temporale (server non raggiungibili)' }
        exit 0
    }
}
exit 1
"""


def _powershell(script, **env):
    return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                          env={**os.environ, "FIRMA_NOME": FIRMATARIO, **env},
                          capture_output=True, text=True, errors="replace")


def find_certificate():
    """(impronta, autofirmato) del certificato da usare, o None se non c'è."""
    out = _powershell(_FIND + "if ($c) { '{0} {1}' -f $c.Thumbprint, ($c.Issuer -eq $c.Subject) }").stdout.split()
    return (out[0], out[1] == "True") if len(out) == 2 else None


def export_certificate(path):
    """Salva la parte pubblica del certificato (.cer), da installare sui PC che devono fidarsi della firma."""
    _powershell(_FIND + "Export-Certificate -Cert $c -FilePath $env:FIRMA_FILE -Type CERT", FIRMA_FILE=str(path))
    return path


def sign(path, attempts=3):
    """Firma il file; ritenta perché l'antivirus può tenere occupato un exe appena scritto."""
    for attempt in range(attempts):
        r = _powershell(_SIGN, FIRMA_FILE=str(path), FIRMA_TS=" ".join(TIMESTAMP_SERVERS))
        if r.returncode == 0:
            print(f"Firmato: {path}", r.stdout.strip())
            return True
        if r.returncode == 2:
            print(f"Nessun certificato per la firma del codice intestato a {FIRMATARIO}: {path} resta senza firma.")
            return False
        time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Firma non riuscita: {path}\n{r.stdout}{r.stderr}")


if __name__ == "__main__":
    ok = [sign(f) for f in sys.argv[1:]]
    sys.exit(0 if all(ok) else 1)
