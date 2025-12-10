import webview
import os
from tkinter import filedialog, messagebox
from generate_facture import generer_facture
from excel_loader import charger_donnees_excel


class Api:
    def __init__(self):
        self.dossier_sortie = ""

    def choisir_dossier(self):
        dossier = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if dossier:
            self.dossier_sortie = dossier
            return f"Dossier sélectionné : {dossier}"
        return "Aucun dossier sélectionné."

    def generer_facture(self, data):
        if not self.dossier_sortie:
            return "Veuillez d'abord choisir un dossier de sortie."

        chemin = generer_facture(data, self.dossier_sortie)
        return f"✅ Facture générée : {chemin}"

    def generer_depuis_excel(self):
        if not self.dossier_sortie:
            return "Veuillez d'abord choisir un dossier de sortie."

        fichier_excel = filedialog.askopenfilename(
            title="Sélectionner le fichier Excel",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not fichier_excel:
            return "Aucun fichier sélectionné."

        for data in charger_donnees_excel(fichier_excel):
            generer_facture(data, self.dossier_sortie)
        return "✅ Toutes les factures ont été générées avec succès !"


# --- Charger ton fichier HTML local ---
base_path = os.path.abspath(os.path.dirname(__file__))
html_file_path = os.path.join(base_path, "web", "index.html")

if not os.path.exists(html_file_path):
    raise FileNotFoundError("Le fichier HTML de l'interface est introuvable : web/index.html")

# --- Lancer la fenêtre WebView ---
api = Api()
webview.create_window(
    title="Fact Gen 1 BNETD",
    url=f"file://{html_file_path}",
    js_api=api,
    width=1280,
    height=720
)
webview.start()
