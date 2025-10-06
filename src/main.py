#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
windows_key_tryer.py
Liest Keys aus einer Textdatei (eine Zeile = ein Key) und versucht sie nacheinander zu installieren und zu aktivieren.
Nutzt slmgr.vbs und wmic zur Prüfung.

Benutzung:
    1) keys.txt erstellen (eine Zeile = ein Key, keine zusätzlichen Kommentare)
    2) Script als Administrator ausführen (Rechtsklick -> "Als Administrator ausführen")
"""

import subprocess
import sys
import time
import os
import ctypes
from pathlib import Path

# Pfad zur Key-Datei (anpassen falls nötig)
KEYFILE = "keys.txt"
# Wartezeit nach jeder Aktion (sekunden)
SLEEP_AFTER_CMD = 3

def is_admin() -> bool:
    """Prüft ob das Script als Administrator läuft (Windows-only)."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def run_command(cmd, timeout=60):
    """
    Führt ein Kommando aus und gibt (returncode, stdout, stderr) zurück.
    cmd sollte eine Liste (recommended) oder String sein.
    """
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, shell=False)
        return proc.returncode, proc.stdout.strip(), proc.stderr.strip()
    except subprocess.TimeoutExpired:
        return -1, "", "TIMEOUT"

def install_key(key: str):
    """Installiert den Produktkey via slmgr.vbs /ipk"""
    slmgr = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "slmgr.vbs")
    cmd = ["cscript", "//nologo", slmgr, "/ipk", key]
    return run_command(cmd)

def attempt_activate():
    """Versucht Windows online zu aktivieren via slmgr.vbs /ato"""
    slmgr = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "slmgr.vbs")
    cmd = ["cscript", "//nologo", slmgr, "/ato"]
    return run_command(cmd, timeout=120)

def check_activation_with_xpr():
    """Prüft /xpr-Ausgabe (gibt True zurück wenn dauerhaft aktiviert)."""
    slmgr = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "slmgr.vbs")
    cmd = ["cscript", "//nologo", slmgr, "/xpr"]
    rc, out, err = run_command(cmd)
    if rc != 0:
        return False, out or err
    # Prüfe typische Texte (englisch + deutsch)
    check_text = (out or "").lower()
    if "permanently activated" in check_text or "is permanently activated" in check_text:
        return True, out
    if "dauerhaft aktiviert" in check_text or "ist dauerhaft aktiviert" in check_text:
        return True, out
    # Manche Systeme geben andere Formulierungen — gib die Ausgabe zurück und false
    return False, out

def check_activation_with_wmic():
    """
    Nutzt WMIC als Fallback. Liest LicenseStatus Werte:
    0 = Unlicensed, 1 = Licensed, 2 = Notification, 3 = Initial Grace Period, 4 = Additional Grace, ...
    Wenn irgendein SoftwareLicensingProduct Eintrag LicenseStatus == 1 hat -> lizenziert.
    """
    # wmic kann mehrere Einträge zurückgeben, daher parse-manipulation
    cmd = ["wmic", "path", "SoftwareLicensingProduct", "where", "PartialProductKey is not null", "get", "Name,LicenseStatus", "/format:list"]
    rc, out, err = run_command(cmd)
    if rc != 0:
        return False, out or err
    text = (out or "").lower()
    if "licensesstatus=1" in text or "licensestatus=1" in text or "licensestatus= 1" in text:
        return True, out
    # manchmal ist Format anders; wir suchen nach "LicenseStatus" gefolgt von 1
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("licensestatus"):
            parts = line.split("=", 1)
            if len(parts) == 2:
                try:
                    val = int(parts[1].strip())
                    if val == 1:
                        return True, out
                except ValueError:
                    pass
    return False, out

def normalize_key_line(line: str) -> str:
    """Entfernt Whitespaces und unerwünschte Zeichen, gibt leere string zurück wenn ungültig."""
    # einfache Normalisierung: nur Buchstaben/Ziffern und Bindestriche behalten
    cleaned = line.strip().replace(" ", "").replace("\uFEFF", "")
    if not cleaned:
        return ""
    # optional: falls Key schon in Form XXXXX-... ohne Bindestriche, du könntest formatieren
    # Akzeptiere nur alphanumerische und '-' Zeichen
    import re
    if re.match(r'^[A-Za-z0-9\-]+$', cleaned):
        return cleaned
    return ""

def try_keys_from_file(filepath: str):
    p = Path(filepath)
    if not p.exists():
        print(f"[ERROR] Key-Datei '{filepath}' nicht gefunden.")
        return False

    with p.open("r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    for idx, raw in enumerate(lines, start=1):
        key = normalize_key_line(raw)
        if not key:
            print(f"[{idx}] Leer/ungültig übersprungen.")
            continue
        print(f"[{idx}] Versuche Key: {key}")

        # Installieren
        rc, out, err = install_key(key)
        print("  -> /ipk returncode:", rc)
        if out:
            print("  STDOUT:", out)
        if err:
            print("  STDERR:", err)
        time.sleep(SLEEP_AFTER_CMD)

        # Aktivieren versuchen
        rc2, out2, err2 = attempt_activate()
        print("  -> /ato returncode:", rc2)
        if out2:
            print("  STDOUT:", out2)
        if err2:
            print("  STDERR:", err2)
        time.sleep(SLEEP_AFTER_CMD)

        # Prüfung: zuerst slmgr /xpr
        ok, detail = check_activation_with_xpr()
        print("  -> /xpr check:", ok)
        if ok:
            print(f"[SUCCESS] Aktiviert mit Key (Zeile {idx}): {key}")
            return True

        # Fallback wmic
        ok2, detail2 = check_activation_with_wmic()
        print("  -> wmic check:", ok2)
        if ok2:
            print(f"[SUCCESS] Aktiviert mit Key (Zeile {idx}): {key}")
            return True

        print(f"[{idx}] Key hat nicht aktiviert. Fahre mit nächstem fort.\n")

    print("[FINISHED] Alle Keys probiert, keine Aktivierung erreicht.")
    return False

if __name__ == "__main__":
    if os.name != "nt":
        print("Dieses Script läuft nur unter Windows.")
        sys.exit(2)

    if not is_admin():
        print("Dieses Script benötigt Administratorrechte. Bitte als Administrator ausführen.")
        sys.exit(3)

    # Optional: Backup / Anzeige des aktuellen Lizenzstatus vor Start
    print("Prüfe aktuellen Aktivierungsstatus...")
    xpr_ok, xpr_out = check_activation_with_xpr()
    print("  /xpr:", xpr_ok)
    print(xpr_out)
    print("Starte Key-Versuche...\n")

    success = try_keys_from_file(KEYFILE)
    if success:
        print("Erfolgreich aktiviert.")
        sys.exit(0)
    else:
        print("Aktivierung nicht erfolgreich.")
        sys.exit(1)
