# Guide utilisateur – Facture Interne Projet V1

## 1. Prérequis
- Windows avec Python 3.12+ si vous exécutez en mode script (`python main.py`).  
- Sinon, utilisez l’exécutable `dist/FactureGenerator.exe` (pas besoin de Python).
- Fichiers nécessaires (déjà inclus) :  
  - `Template/Modèle facture interne projet - V1.docx` (template Word)  
  - `data/logo.jpg` (icône)  
  - `web/` (vue web facultative)

## 2. Lancer l’application
- **Depuis l’exe** : double-cliquez sur `dist/FactureGenerator.exe` (pas de console).  
- **Depuis Python** : `python main.py` (venv activé avec les dépendances installées).

## 3. Choisir le dossier de sortie
1. Cliquez sur **Choisir...** dans “Dossier de sortie”.  
2. Sélectionnez le répertoire où les factures seront enregistrées.

## 4. Mode Formulaire (facture unique)
1. Cliquez sur **Aller au formulaire**.  
2. Renseignez les champs : code projet/sous-projet/OTFI, pôles, départements, dates (picker), période, somme facture.  
3. Ajoutez des lignes de prestations (désignation, type, unité, quantité, prix unitaire). Les montants et le total HT sont calculés.  
4. Bouton **Generer la facture** : crée le .docx dans le dossier de sortie et ajoute une entrée à l’historique.  
5. Bouton **Reinitialiser** : vide le formulaire et recrée 3 lignes vides.

## 5. Mode Excel (batch)
1. Cliquez sur **Charger un fichier Excel**.  
2. Si besoin, récupérez le **Template de remplissage** pour connaître les colonnes attendues.  
3. Choisissez le fichier `.xlsx` à importer : une facture est générée pour chaque ligne.

### Colonnes Excel attendues
Champs principaux : `code_projet`, `code_sous_projet`, `numero_otfi`, `pole_emettrice`, `pole_destinataire`, `dept_dir_emettrice`, `dept_dir_destinataire`, `date_emission`, `date_du_jour`, `periode_concernee`, `somme_facture`.  
Lignes de prestations (jusqu’à 5) : `ligne1_designation`, `ligne1_type_prestation`, `ligne1_unite`, `ligne1_quantite`, `ligne1_prix_unitaire` … idem jusqu’à `ligne5_...`.

## 6. Historique
- Chaque génération (formulaire ou Excel) ajoute une entrée : date/heure, référence (OTFI/projet/sous-projet), montant, chemin du fichier.  
- Filtres (un à la fois, via sélecteur) :  
  - Date exacte  
  - Intervalle de dates  
  - Montant exact  
  - Intervalle de montants  
  - Titre/référence (fichier ou codes)  
- Actions :  
  - **Supprimer la sélection** : retire l’entrée de l’historique (ne supprime pas le fichier).  
  - **Exporter l’historique (Excel)** : sauvegarde un récapitulatif `.xlsx`.

## 7. Vue web (optionnelle)
- `main_webview.py` + `web/index.html` permettent une interface WebView. Lancer `python main_webview.py` (pywebview requis).

## 8. Dépannage rapide
- Module manquant : installez les dépendances via `pip install -r requirements.txt`.  
- Fichier non trouvé : vérifiez que `Template/` et `data/` sont présents à côté de l’exe ou du script.  
- Console noire avec l’exe : utiliser la version `--windowed` (déjà fourni).  
- Erreur Excel : vérifiez les noms de colonnes et le format `.xlsx`.
