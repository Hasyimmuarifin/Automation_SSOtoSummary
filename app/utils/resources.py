import sys, os

def resource_path(relative_path: str) -> str:
    """Dapatkan path ke resource (works untuk .py dan .exe)"""
    if getattr(sys, "frozen", False):  
        base_path = sys._MEIPASS  # folder temp PyInstaller
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)