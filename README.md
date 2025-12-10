# Facture interne projet - Generator V1

Application de génération de factures basée sur un template Word, avec deux modes : saisie via formulaire et import Excel. Interface Tkinter modernisée, export d'historique et gestion des modèles Excel.

## Fonctionnalités principales
- **Formulaire manuel** : saisie du code projet/sous-projet/OTFI, pôles, départements, dates (pickers), période, montant global et lignes de prestations (quantité x prix, calcul du total HT).
- **Import Excel** : génération batch depuis un fichier `.xlsx` avec les colonnes attendues. Bouton pour télécharger un **template Excel** pré-rempli avec l’exemple de structure (inclut jusqu’à 5 lignes de prestations).
- **Template Word** : rendu via `docxtpl` avec remplissage dynamique des placeholders, y compris le tableau des prestations (boucle `lignes` + `total_ht`).
- **Historique** : enregistrement de chaque facture générée (date/heure, référence, montant, chemin). Filtres avancés (un à la fois) : date exacte, intervalle de dates, montant exact, intervalle de montants, titre/référence. Suppression d’entrée et export de l’historique en Excel.
- **UI** : navigation accueil/formulaire/import, logo, sélection du dossier de sortie, scroll, bouton de réinitialisation du formulaire et ajout/suppression de lignes.

## Pré-requis
- Python 3.10+ recommandé.
- Dépendances principales : `docxtpl`, `pandas`, `openpyxl`, `tkcalendar` (optionnel pour date picker avancé), `pywebview` (pour `main_webview.py` si utilisé).

## Installation
```bash
python -m venv .venv
.venv\Scripts\activate       # Windows
pip install -r requirements.txt   # si présent, sinon installer les libs listées ci-dessus
```

## Lancement (mode Tkinter)
```bash
python main.py
```
1. Choisir le dossier de sortie.
2. Soit **Remplir le formulaire** (ajouter lignes de prestations si besoin).
3. Soit **Charger un fichier Excel** (utiliser le bouton “Template de remplissage” pour récupérer le modèle attendu).

## Colonnes attendues pour l’import Excel
Champs principaux :
- `code_projet`, `code_sous_projet`, `numero_otfi`, `pole_emettrice`, `pole_destinataire`, `dept_dir_emettrice`, `dept_dir_destinataire`, `date_emission`, `date_du_jour`, `periode_concernee`, `somme_facture`.

Lignes de prestations (jusqu’à 5) :
- `ligne1_designation`, `ligne1_type_prestation`, `ligne1_unite`, `ligne1_quantite`, `ligne1_prix_unitaire`
- ... répéter jusqu’à `ligne5_...`

## Historique
- Stocké dans `data/history.json`.
- Filtres exclusifs via un sélecteur (date exacte, intervalle de dates, montant exact, intervalle de montants, titre/référence).
- Actions : suppression d’entrée, export Excel.

## Structure des principaux fichiers
- `main.py` : UI Tkinter, logique formulaire/import, historique.
- `excel_loader.py` : lecture/normalisation Excel, mapping des colonnes, calcul des montants lignes.
- `generate_facture.py` : rendu `docxtpl` avec placeholders, support `lignes` et `total_ht`.
- `web/index.html` + `main_webview.py` : alternative webview (facultatif).
- `Template/` : template Word utilisé pour la génération.

## Remarques
- Les dates sont manipulées en JJ/MM/AAAA pour l’UI ; le template Word doit avoir les placeholders alignés (`date_emission`, `date_du_jour`, etc.).
- L’option `tkcalendar` améliore l’expérience pour la sélection de dates ; sinon un fallback manuel est utilisé.
