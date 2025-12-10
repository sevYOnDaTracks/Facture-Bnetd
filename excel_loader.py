import os
import re
import sys

import pandas as pd


def resource_path(relative_path):
    """
    Renvoie le bon chemin pour acceder a un fichier (utile pour le .exe).
    """
    try:
        base_path = sys._MEIPASS  # utilise par PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


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

# Normalized column name -> placeholder attendu dans le docx
PLACEHOLDER_ALIASES = {
    "code_projet": "code_projet",
    "codeprojet": "code_projet",
    "code_sous_projet": "code_sous_projet",
    "codesousprojet": "code_sous_projet",
    "date_du_jour": "date_du_jour",
    "datedujour": "date_du_jour",
    "date_emission": "date_emission",
    "dateemission": "date_emission",
    "dept_dir_destinataire": "dept_dir_destinataire",
    "deptdirdestinataire": "dept_dir_destinataire",
    "dept_dir_emettrice": "dept_dir_emettrice",
    "deptdiremettrice": "dept_dir_emettrice",
    "numero_otfi": "numero_otfi",
    "poledestinataire": "pole_destinataire",
    "pole_destinataire": "pole_destinataire",
    "poleemettrice": "pole_emettrice",
    "pole_emettrice": "pole_emettrice",
    "periode_concernee": "période_concernee",
    "période_concernee": "période_concernee",
    "periode": "période_concernee",
    "somme_facture": "somme_facture",
    "montant": "somme_facture",
}


def normalize_key(name: str) -> str:
    """Nettoie un nom de colonne pour matcher les placeholders du template."""
    base = str(name).strip().lower()
    for accented, plain in ACCENT_MAP.items():
        base = base.replace(accented, plain)
    base = re.sub(r"[^a-z0-9]+", "_", base)
    return base.strip("_")


def charger_donnees_excel(fichier_excel):
    """
    Charge un fichier Excel contenant les donnees de facturation.

    Les noms de colonnes sont normalises (minuscules, accents retires, espaces -> underscore)
    puis associes aux placeholders du modele Word :
    code_sous_projet, numero_otfi, pole_emettrice, pole_destinataire,
    dept_dir_emettrice, dept_dir_destinataire, date_emission, date_du_jour,
    période_concernee, somme_facture. Les colonnes de lignes de prestations sont
    supportees sous la forme ligne1_designation, ligne1_type_prestation, ligne1_unite,
    ligne1_quantite, ligne1_prix_unitaire (jusqu'a 5 lignes).
    """
    if not os.path.exists(fichier_excel):
        raise FileNotFoundError(f"Fichier introuvable : {fichier_excel}")

    df = pd.read_excel(fichier_excel, engine="openpyxl").fillna("")

    normalized = [normalize_key(col) for col in df.columns]
    df.columns = normalized

    def to_float(val):
        try:
            return float(str(val).replace(",", "."))
        except Exception:
            return 0.0

    for _, row in df.iterrows():
        record = {col: row[col] for col in df.columns}

        mapped = {}
        for norm_key, placeholder in PLACEHOLDER_ALIASES.items():
            if norm_key in record:
                mapped[placeholder] = record.get(norm_key, "")

        # Champs optionnels utiles pour le nommage de fichier
        mapped["nom"] = record.get("nom", "")
        mapped["prenom"] = record.get("prenom", "")

        # Lignes de prestations : support jusqu'a 5 lignes
        lignes = []
        for idx in range(1, 6):
            prefix = f"ligne{idx}_"
            designation = record.get(prefix + "designation", "")
            type_prest = record.get(prefix + "type_prestation", "")
            unite = record.get(prefix + "unite", "")
            qte_raw = record.get(prefix + "quantite", "")
            pu_raw = record.get(prefix + "prix_unitaire", "")

            if not any([designation, type_prest, unite, qte_raw, pu_raw]):
                continue

            quantite = to_float(qte_raw)
            prix_unitaire = to_float(pu_raw)
            montant = quantite * prix_unitaire
            lignes.append(
                {
                    "numero": len(lignes) + 1,
                    "designation": designation,
                    "type_prestation": type_prest,
                    "unite": unite,
                    "quantite": quantite,
                    "prix_unitaire": prix_unitaire,
                    "montant": montant,
                }
            )

        if lignes:
            mapped["lignes"] = lignes
            mapped["total_ht"] = sum(l["montant"] for l in lignes)

        yield mapped
