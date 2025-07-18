# === Modules ===
import tkinter as tk
from tkinter import messagebox, ttk
import requests
import zipfile
import io
import os
from pathlib import Path
import threading
import csv
import sys
from io import StringIO
import winreg as reg
import ctypes

# === Config ===
SHEET_ID = "1c8gg13-GlvBaxH06XiTsfmcyT20Ukdt9Up0ix2JV38E"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv"
DESTINATION_FOLDER = Path("C:/Program Files/WRT")
EXE_NAME = "WindowsRunTool.exe"

REGISTRY_PATH = r"Software\WRT"
VERSION_WRT = "Version"

# === ADMIN ===
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def relaunch_as_admin():
    script_path = os.path.abspath(__file__)
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{script_path}"', None, 1)
    sys.exit(0)

if not is_admin():
    relaunch_as_admin()


# === Reading version from registry ===
try:
    key = reg.OpenKey(reg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, reg.KEY_READ)
    version, regtype = reg.QueryValueEx(key, VERSION_WRT)
    reg.CloseKey(key)
    print(f"Version lue dans le registre : {version}")
except FileNotFoundError:
    print("La valeur 'Version' n'existe pas dans le registre.")
    version = "Aucune version installée"
except Exception as e:
    print(f"Erreur : {e}")
    version = "Erreur de lecture"

# === research ===
def fetch_versions():
    try:
        response = requests.get(CSV_URL)
        response.raise_for_status()
        reader = csv.reader(StringIO(response.text))
        rows = list(reader)
        return rows  # [version, date, lien]
    except Exception as e:
        print("Erreur de récupération:", e)
        return []

# === Remove legacy .exe ===
def remove_old_exe(destination_folder, exe_name):
    exe_path = destination_folder / exe_name
    if exe_path.exists():
        try:
            os.remove(exe_path)
            print(f"{exe_name} supprimé avec succès.")
        except Exception as e:
            print(f"Erreur lors de la suppression de {exe_name} :", e)

# === Download and Unzip ===
def download_and_extract_zip(zip_url, destination_folder, callback):
    try:
        remove_old_exe(destination_folder, EXE_NAME)

        response = requests.get(zip_url)
        response.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(response.content)) as zip_ref:
            zip_ref.extractall(destination_folder)
        callback(success=True)
    except Exception as e:
        print("Erreur de téléchargement ou d'extraction:", e)
        callback(success=False)

# === Interface graphique ===
class UpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WRT - Gestionnaire de mises à jour et d'installation")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        def get_resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except AttributeError:
                base_path = os.path.abspath(".")

            return os.path.join(base_path, relative_path)
        icon_path = get_resource_path("Image/logowrt.ico")
        root.iconbitmap(icon_path)

        # Icon
        def get_resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except AttributeError:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)

        try:
            icon_path = get_resource_path("Image/logowrt.ico")
            root.iconbitmap(icon_path)
        except Exception:
            pass

        self.versions = []
        self.selected_index = tk.IntVar(value=-1)

        self.label = tk.Label(root, text="Sélectionnez une version à installer :", font=("Segoe UI", 12))
        self.label.pack(pady=10)

        self.listbox = tk.Listbox(root, height=10, font=("Segoe UI", 10), activestyle='dotbox')
        self.listbox.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        self.refresh_button = tk.Button(root, text="Rafraîchir", command=self.load_versions)
        self.refresh_button.pack(pady=5)

        self.download_button = tk.Button(root, text="Télécharger et installer", command=self.install_selected_version)
        self.download_button.pack(pady=10)

        self.lbl_version = tk.Label(root, text=f"Version installée : {version}", font=("Arial", 10), borderwidth=0.5, fg="red")
        self.lbl_version.pack(pady=20)

        self.status_label = ttk.Label(root, text="", foreground="green")
        self.status_label.pack(pady=5)

        self.load_versions()

    def load_versions(self):
        self.listbox.delete(0, tk.END)
        self.versions = fetch_versions()
        if not self.versions:
            self.status_label.config(text="Erreur lors du chargement des versions.", foreground="red")
            return

        for version in self.versions:
            ver_str = f"Version {version[0]} - Sortie le {version[1]}"
            self.listbox.insert(tk.END, ver_str)
        self.status_label.config(text="Versions chargées avec succès.", foreground="green")

    def install_selected_version(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une version.")
            return

        index = selected[0]
        _, _, url = self.versions[index]

        self.status_label.config(text="Téléchargement en cours...", foreground="blue")
        self.download_button.config(state=tk.DISABLED)

        def on_complete(success):
            self.root.after(0, lambda: self._on_download_complete(success))

        threading.Thread(target=download_and_extract_zip, args=(url, DESTINATION_FOLDER, on_complete)).start()

    def _on_download_complete(self, success):
        if success:
            self.status_label.config(text="Téléchargement et installation réussis !", foreground="green")
        else:
            self.status_label.config(text="Échec du téléchargement ou de l'installation.", foreground="red")
        self.download_button.config(state=tk.NORMAL)

# === Start ===
if __name__ == "__main__":
    if not DESTINATION_FOLDER.exists():
        os.makedirs(DESTINATION_FOLDER)

    root = tk.Tk()
    app = UpdaterApp(root)
    root.mainloop()
