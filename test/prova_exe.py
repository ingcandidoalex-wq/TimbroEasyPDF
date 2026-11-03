"""Avvia l'exe costruito, aspetta la sua finestra, la chiude normalmente (WM_CLOSE) e riporta i tempi.

Uso:  python test\\prova_exe.py [percorso.exe]       (di norma dist\\TimbroEasyPDF.exe)
      python test\\prova_exe.py --chiudi nome.exe    (chiude le finestre dei processi con quel nome)
"""
import ctypes
import subprocess
import sys
import time
from ctypes import wintypes

from comune import PROGRAMMA

user32 = ctypes.windll.user32
WM_CLOSE = 0x0010


def windows_of(pids):
    found = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def cb(hwnd, _):
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value in pids and user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buf, 256)
            if buf.value:
                found.append((hwnd, buf.value))
        return True

    user32.EnumWindows(cb, 0)
    return found


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD), ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.c_void_p), ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD), ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_wchar * 260)]


def processes():
    """(pid, pid del padre, nome dell'exe) di tutti i processi."""
    k32 = ctypes.windll.kernel32
    k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    snap = k32.CreateToolhelp32Snapshot(2, 0)   # TH32CS_SNAPPROCESS
    e = PROCESSENTRY32W()
    e.dwSize = ctypes.sizeof(e)
    out = []
    ok = k32.Process32FirstW(snap, ctypes.byref(e))
    while ok:
        out.append((e.th32ProcessID, e.th32ParentProcessID, e.szExeFile))
        ok = k32.Process32NextW(snap, ctypes.byref(e))
    k32.CloseHandle(snap)
    return out


def child_pids(pid):
    return {p for p, parent, _ in processes() if parent == pid}


if __name__ == "__main__":
    if sys.argv[1:2] == ["--chiudi"]:
        pids = {p for p, _, name in processes() if name.lower() == sys.argv[2].lower()}
        for hwnd, title in windows_of(pids):
            print("chiudo", title)
            user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
        sys.exit(0)

    exe = sys.argv[1] if len(sys.argv) > 1 else str(PROGRAMMA / "dist" / "TimbroEasyPDF.exe")
    t0 = time.time()
    p = subprocess.Popen([exe])
    win = []
    while time.time() - t0 < 60 and p.poll() is None:
        win = windows_of({p.pid} | child_pids(p.pid))   # il file unico si estrae e riparte come processo figlio
        if win:
            break
        time.sleep(0.2)
    if not win:
        print("nessuna finestra; codice di uscita:", p.poll())
        sys.exit(1)
    print(f"finestra '{win[0][1]}' dopo {time.time() - t0:.1f} s")
    time.sleep(2)
    for hwnd, _ in win:
        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
    try:
        code = p.wait(30)
        print("uscita con codice", code)
        sys.exit(code)
    except subprocess.TimeoutExpired:
        print("non si è chiuso entro 30 s")
        sys.exit(1)
