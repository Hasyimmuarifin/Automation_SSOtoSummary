import sys
import os
from datetime import datetime
import pathlib

class Tee:
    def __init__(self, filename):
        # Tentukan console asli (fallback jika None)
        self.stdout = sys.__stdout__ if sys.stdout is None else sys.stdout

        # Buat folder log di user home agar selalu bisa write
        log_dir = os.path.join(pathlib.Path.home(), "SSOtoSummary_Logs")
        os.makedirs(log_dir, exist_ok=True)

        filepath = os.path.join(log_dir, filename)

        # Coba buka file
        try:
            self.file = open(filepath, "a", encoding="utf-8")
        except:
            # Jika gagal, disable logging ke file
            self.file = None

    def write(self, data):
        # Tulis ke console
        try:
            if self.stdout:
                self.stdout.write(data)
        except:
            pass

        # Tulis ke file jika tersedia
        if self.file:
            try:
                self.file.write(data)
            except:
                pass

    def flush(self):
        try:
            if self.stdout:
                self.stdout.flush()
        except:
            pass

        if self.file:
            try:
                self.file.flush()
            except:
                pass


def start_logging():
    # Nama file log dengan timestamp
    log_filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    # Redirect stdout/stderr hanya jika bukan None
    new_tee = Tee(log_filename)

    sys.stdout = new_tee
    sys.stderr = new_tee   # aman untuk PyQt6
