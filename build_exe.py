"""Crea il programma da distribuire, in dist:

  <nome>.exe                   un solo file: si usa senza installare (basta copiarlo)
  Setup <nome> <versione>.exe  installatore, con menu Start, disinstallazione e PDF associati
                               (serve Inno Setup 6, gratuito: https://jrsoftware.org)
  Guida <nome>.pdf             la guida illustrata

Se nell'archivio di Windows c'è il certificato per la firma del codice, exe e setup escono firmati
(vedi firma_codice.py). Il setup normale è quello da pubblicare; con --studio se ne crea uno a parte,
«Setup <nome> <versione> (studio).exe», con l'opzione che installa il certificato autofirmato: solo
per i PC dello studio, mai da distribuire.

Uso:  python -m pip install pyinstaller pymupdf
      python build_exe.py               (tutto)
      python build_exe.py --solo-setup  (solo l'installatore, es. se l'exe è aperto e non si può riscrivere)
      python build_exe.py --studio      (in più: l'installatore per i PC dello studio, con il certificato)
"""

import os
import shutil
import struct
import subprocess
import sys
import tempfile
from pathlib import Path

from firma_codice import FIRMATARIO, export_certificate, find_certificate, sign
from timbro_pdf import (APP_MUTEX, APP_NAME, APP_VERSION, AUTHOR, GUIDE_NAME, LICENSE_TEXT, PROG_ID,
                        icon_pixmap)

HERE = Path(__file__).parent
DIST = HERE / "dist"
# identifica il programma per Windows (aggiornamenti e disinstallazione): non va mai cambiato
APP_ID = "{7E3C9A52-4B1D-4F8E-9C26-5A0D3B8E61F4}"
# pymupdf.layout (analisi AI del layout, opzionale) non serve e non funziona dentro l'exe
EXCLUDED = ["pymupdf.layout", "onnxruntime", "numpy"]

INSTALLER = r'''; generato da build_exe.py
#define AppName "@APP@"
#define AppVersion "@VERSION@"
#define AppExe "@APP@.exe"
@FIRMA@
[Setup]
AppId={{@APP_ID@
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=@AUTHOR@
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; senza diritti di amministratore si installa solo per l'utente; si può scegliere di installare per tutti
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=@DIST@
OutputBaseFilename=Setup {#AppName} {#AppVersion}@SUFFIX@
SetupIconFile=@ICON@
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
LicenseFile=@LICENSE@
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ChangesAssociations=yes
AppMutex=@MUTEX@
CloseApplications=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
#ifdef Firma
; firma il setup e il programma di disinstallazione (il comando "firma" lo passa build_exe.py a ISCC)
SignTool=firma
#endif

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
Name: "pdf"; Description: "Aggiungi {#AppName} ai programmi che aprono i PDF"; GroupDescription: "File PDF:"
#ifdef CertFile
Name: "certificato"; Description: "Considera attendibili i programmi firmati da {#Firmatario} (installa il certificato: solo per i PC dello studio)"; GroupDescription: "Firma digitale:"; Flags: unchecked
#endif

[Files]
Source: "@APPDIR@\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "@GUIDE@"; DestDir: "{app}"; Flags: ignoreversion
Source: "@AGPL@"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion
#ifdef CertFile
Source: "{#CertFile}"; DestDir: "{app}"; Flags: ignoreversion; Tasks: certificato
#endif

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autoprograms}\Guida di {#AppName}"; Filename: "{app}\@GUIDE_NAME@"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Registry]
; le stesse voci che il programma scrive con File > Usa {#AppName} per aprire i PDF
Root: HKA; Subkey: "Software\Classes\@PROG@"; ValueType: string; ValueName: ""; ValueData: "Documento PDF"; Flags: uninsdeletekey; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\@PROG@"; ValueType: string; ValueName: "FriendlyTypeName"; ValueData: "Documento PDF ({#AppName})"; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\@PROG@\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"",0"; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\@PROG@\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" ""%1"""; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\.pdf\OpenWithProgids"; ValueType: string; ValueName: "@PROG@"; ValueData: ""; Flags: uninsdeletevalue; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}"; ValueType: none; Flags: uninsdeletekey; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" ""%1"""; Tasks: pdf
Root: HKA; Subkey: "Software\Classes\Applications\{#AppExe}\SupportedTypes"; ValueType: string; ValueName: ".pdf"; ValueData: ""; Tasks: pdf
Root: HKA; Subkey: "Software\{#AppName}"; ValueType: none; Flags: uninsdeletekey; Tasks: pdf
Root: HKA; Subkey: "Software\{#AppName}\Capabilities"; ValueType: string; ValueName: "ApplicationName"; ValueData: "{#AppName}"; Tasks: pdf
Root: HKA; Subkey: "Software\{#AppName}\Capabilities"; ValueType: string; ValueName: "ApplicationDescription"; ValueData: "Visualizza, timbra, firma e modifica i PDF"; Tasks: pdf
Root: HKA; Subkey: "Software\{#AppName}\Capabilities\FileAssociations"; ValueType: string; ValueName: ".pdf"; ValueData: "@PROG@"; Tasks: pdf
Root: HKA; Subkey: "Software\RegisteredApplications"; ValueType: string; ValueName: "{#AppName}"; ValueData: "Software\{#AppName}\Capabilities"; Flags: uninsdeletevalue; Tasks: pdf

[Run]
#ifdef CertFile
; tra le autorità radice (per l'utente Windows chiede conferma) e tra gli autori attendibili;
; niente runhidden dove Windows può chiedere conferma, altrimenti la sua finestra potrebbe restare nascosta
Filename: "{sys}\certutil.exe"; Parameters: "{code:CertScope}-f -addstore Root ""{app}\{#CertName}"""; StatusMsg: "Installazione del certificato..."; Tasks: certificato
Filename: "{sys}\certutil.exe"; Parameters: "{code:CertScope}-f -addstore TrustedPublisher ""{app}\{#CertName}"""; Flags: runhidden; Tasks: certificato
#endif
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
Filename: "ms-settings:defaultapps?registeredAppUser={#AppName}"; Description: "Scegli {#AppName} come programma predefinito per i PDF"; Flags: shellexec nowait postinstall skipifsilent unchecked; Tasks: pdf
#ifdef CertFile

[UninstallRun]
Filename: "{sys}\certutil.exe"; Parameters: "{code:CertScope}-delstore Root {#CertThumb}"; RunOnceId: "CertificatoRadice"; Tasks: certificato
Filename: "{sys}\certutil.exe"; Parameters: "{code:CertScope}-delstore TrustedPublisher {#CertThumb}"; Flags: runhidden; RunOnceId: "CertificatoAutori"; Tasks: certificato

[Code]
// installazione per l'utente: archivi dell'utente; per tutti (amministratore): archivi del computer
function CertScope(Param: String): String;
begin
  if IsAdminInstallMode then Result := '' else Result := '-user ';
end;
#endif
'''


def write_ico(path, sizes=(16, 24, 32, 48, 64, 256)):
    """File .ico con immagini PNG incorporate (formato supportato da Windows Vista in poi)."""
    pngs = [icon_pixmap(s).tobytes("png") for s in sizes]
    header = struct.pack("<HHH", 0, 1, len(pngs))
    offset = 6 + 16 * len(pngs)
    entries = b""
    for size, data in zip(sizes, pngs):
        entries += struct.pack("<BBBBHHII", size % 256, size % 256, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
    path.write_bytes(header + entries + b"".join(pngs))


def pyinstaller(mode, distpath, ico, work):
    """mode: "--onefile" (un solo exe) o "--onedir" (cartella, si avvia più in fretta: per l'installatore)."""
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", mode, "--windowed",
                    "--name", APP_NAME, "--icon", str(ico),
                    "--collect-submodules", "asn1crypto",                       # lettura dei file firmati .p7m
                    "--add-data", f"{HERE / 'tessdata'}{os.pathsep}tessdata",   # lingue dell'OCR
                    "--add-data", f"{HERE / GUIDE_NAME}{os.pathsep}.",           # guida illustrata
                    "--add-data", f"{HERE / 'LICENSE'}{os.pathsep}.",            # testo della licenza AGPL-3.0
                    "--workpath", str(Path(work) / "build"), "--specpath", work, "--distpath", str(distpath),
                    *[arg for mod in EXCLUDED for arg in ("--exclude-module", mod)],
                    str(HERE / "timbro_pdf.py")],
                   cwd=HERE, check=True)


def find_iscc():
    """Il compilatore di Inno Setup, se è installato."""
    candidates = [Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Inno Setup 6" / "ISCC.exe",
                  Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Inno Setup 6" / "ISCC.exe",
                  Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Inno Setup 6" / "ISCC.exe"]
    return next((p for p in candidates if p.exists()), None) or shutil.which("ISCC")


def build_setup(iscc, appdir, ico, work, studio=False):
    """studio: con l'opzione che installa il certificato autofirmato (solo per i PC dello studio)."""
    licence = Path(work) / "licenza.txt"
    licence.write_text(LICENSE_TEXT, encoding="utf-8-sig")
    defines, options = "", []
    cert = find_certificate()
    if cert:
        thumbprint, self_signed = cert
        defines = f'#define Firma\n#define Firmatario "{FIRMATARIO}"\n'
        # Inno Setup firma richiamando firma_codice.py ($q sono le virgolette, $f il file da firmare)
        options.append(f"/Sfirma=$q{sys.executable}$q $q{HERE / 'firma_codice.py'}$q $f")
        if self_signed and studio:   # nessun ente lo garantisce: lo si installa solo dove ci si fida (opzione)
            cer = export_certificate(Path(work) / f"{FIRMATARIO} - firma del codice.cer")
            defines += f'#define CertFile "{cer}"\n#define CertName "{cer.name}"\n#define CertThumb "{thumbprint}"\n'
    suffix = " (studio)" if studio else ""
    script = INSTALLER
    for key, value in (("@FIRMA@", defines), ("@APP_ID@", APP_ID[1:]), ("@APP@", APP_NAME), ("@VERSION@", APP_VERSION),
                       ("@AUTHOR@", AUTHOR[0].upper() + AUTHOR[1:]), ("@DIST@", str(DIST)), ("@ICON@", str(ico)),
                       ("@LICENSE@", str(licence)), ("@MUTEX@", APP_MUTEX), ("@APPDIR@", str(appdir)),
                       ("@GUIDE@", str(HERE / GUIDE_NAME)), ("@GUIDE_NAME@", GUIDE_NAME), ("@PROG@", PROG_ID),
                       ("@AGPL@", str(HERE / "LICENSE")), ("@SUFFIX@", suffix)):
        script = script.replace(key, value)
    # si crea fuori da Dropbox: mentre la sincronizzazione o l'antivirus leggono il file appena scritto,
    # Windows non permette di aggiungervi l'icona (EndUpdateResource); poi lo si copia in dist
    out = Path(work) / "setup"
    script = script.replace(f"OutputDir={DIST}", f"OutputDir={out}")
    iss = Path(work) / "installer.iss"
    iss.write_text(script, encoding="utf-8-sig")
    subprocess.run([str(iscc), "/Q", *options, str(iss)], check=True)
    setup = DIST / f"Setup {APP_NAME} {APP_VERSION}{suffix}.exe"
    shutil.copy(out / setup.name, setup)
    return setup


if __name__ == "__main__":
    ico = HERE / "timbro.ico"
    write_ico(ico)
    DIST.mkdir(exist_ok=True)
    # file temporanei fuori dalla cartella del progetto (che può essere sincronizzata da Dropbox)
    with tempfile.TemporaryDirectory() as work:
        if "--solo-setup" not in sys.argv:
            # creato e firmato fuori da Dropbox, poi copiato in dist
            pyinstaller("--onefile", Path(work) / "file_unico", ico, work)
            exe = Path(work) / "file_unico" / f"{APP_NAME}.exe"
            sign(exe)
            shutil.copy(exe, DIST / exe.name)
            shutil.copy(HERE / GUIDE_NAME, DIST / GUIDE_NAME)   # la guida anche come file a parte
            print("\nPronto:", DIST / exe.name)
        iscc = find_iscc()
        if not iscc:
            print("Inno Setup 6 non è installato: niente installatore (https://jrsoftware.org/isdl.php).")
            sys.exit(0)
        appdir = Path(work) / "cartella"
        pyinstaller("--onedir", appdir, ico, work)
        sign(appdir / APP_NAME / f"{APP_NAME}.exe")
        print("Pronto:", build_setup(iscc, appdir / APP_NAME, ico, work))
        if "--studio" in sys.argv:
            print("Pronto:", build_setup(iscc, appdir / APP_NAME, ico, work, studio=True))
