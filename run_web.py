#!/usr/bin/env python3
"""
Lanceur Praxedo Processor Web
==============================
Démarre Streamlit et ouvre le navigateur automatiquement.

Usage : python run_web.py
"""

import os
import sys
import time
import socket
import threading
import webbrowser
import subprocess
from pathlib import Path


def find_free_port(preferred: int = 8501) -> int:
    """Retourne un port disponible, en essayant d'abord celui fourni."""
    for port in (preferred, 8502, 8503, 8504, 8505, 0):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", port))
            actual = s.getsockname()[1]
            s.close()
            return actual
        except OSError:
            continue
    return preferred


def wait_and_open_browser(url: str, timeout: float = 15.0):
    """Attend que le serveur réponde puis ouvre le navigateur."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.5)
            host, port = url.split("//")[1].split(":")
            port = int(port.split("/")[0])
            s.connect((host, port))
            s.close()
            webbrowser.open(url)
            return
        except (OSError, ValueError):
            time.sleep(0.3)
    # Ouvrir quand même au cas où
    webbrowser.open(url)


def main():
    here = Path(__file__).resolve().parent
    app_file = here / "app.py"

    if not app_file.exists():
        print(f"❌ Fichier introuvable : {app_file}")
        sys.exit(1)

    port = find_free_port(8501)
    url = f"http://localhost:{port}"

    print("=" * 60)
    print("  📊 Praxedo Processor — Version Web")
    print("=" * 60)
    print(f"  🌐 URL : {url}")
    print(f"  📁 App : {app_file}")
    print("  (Ctrl+C pour arrêter)")
    print("=" * 60)
    print()

    # Ouvrir navigateur en arrière-plan
    threading.Thread(
        target=wait_and_open_browser,
        args=(url,),
        daemon=True,
    ).start()

    # Lancer Streamlit en premier plan
    cmd = [
        sys.executable, "-m", "streamlit", "run", str(app_file),
        "--server.port", str(port),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--server.fileWatcherType", "none",
    ]
    try:
        subprocess.run(cmd, cwd=str(here))
    except KeyboardInterrupt:
        print("\n👋 Arrêt de l'application.")


if __name__ == "__main__":
    main()
