from docxtpl import DocxTemplate
import os
import sys


ACCENT_MAP = {
    "é": "e",
    "è": "e",
    "ê": "e",
    "ë": "e",
    "à": "a",
    "â": "a",
    "ä": "a",
    "ù": "u",
    "û": "u",
    "ü": "u",
    "ï": "i",
    "î": "i",
    "ô": "o",
    "ö": "o",
    "ç": "c",
}


def _normalize_key(name: str) -> str:
    base = str(name).strip().lower()
    for accented, plain in ACCENT_MAP.items():
        base = base.replace(accented, plain)
    for char in (" ", "-", ".", "/", "\\"):
        base = base.replace(char, "_")
    return base


def resource_path(relative_path):
    """Permet de trouver le bon chemin du template meme apres build .exe"""
    try:
        base_path = sys._MEIPASS  # PyInstaller cree ce dossier temporaire
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def generer_facture(data, dossier_sortie):
    """
    Remplit le template Word avec les placeholders disponibles.
    Les nouvelles variables ajoutees dans le .docx sont initialisées a vide
    si elles ne sont pas fournies dans les donnees.
    """
    template_path = resource_path("Template/Mode\u0300le facture interne projet - V1.docx")
    doc = DocxTemplate(template_path)

    placeholders = doc.get_undeclared_template_variables()
    context = {}
    for key in placeholders:
        if key in data:
            context[key] = data.get(key, "")
            continue
        normalized = _normalize_key(key)
        if key == "lignes":
            context[key] = data.get("lignes", [])
        elif key == "total_ht":
            context[key] = data.get("total_ht", 0)
        else:
            context[key] = data.get(normalized, "")
    doc.render(context)

    os.makedirs(dossier_sortie, exist_ok=True)
    nom_base = (
        data.get("numero_otfi")
        or data.get("code_sous_projet")
        or data.get("code_projet")
        or data.get("nom")
        or "Facture"
    )
    nom_fichier = f"Facture_{nom_base}.docx"
    chemin_fichier = os.path.join(dossier_sortie, nom_fichier)

    doc.save(chemin_fichier)
    return chemin_fichier
