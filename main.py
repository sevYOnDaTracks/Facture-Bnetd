import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import date, datetime
import pandas as pd
try:
    from PIL import Image, ImageTk  # type: ignore
except ImportError:
    Image = None
    ImageTk = None
try:
    from tkcalendar import Calendar  # type: ignore
except ImportError:
    Calendar = None

from excel_loader import charger_donnees_excel
from generate_facture import generer_facture

# Palette et polices pour un rendu plus propre
COLORS = {
    "bg": "#f5f7fb",
    "card": "#ffffff",
    "text": "#0f172a",
    "muted": "#6b7280",
    "accent": "#2563eb",
    "success": "#16a34a",
    "warning": "#f59e0b",
    "border": "#e5e7eb",
}
TITLE_FONT = ("Segoe UI", 16, "bold")
SUBTITLE_FONT = ("Segoe UI", 11)
LABEL_FONT = ("Segoe UI", 10, "bold")
ENTRY_FONT = ("Segoe UI", 10)
BUTTON_FONT = ("Segoe UI", 10, "bold")

APP_DIR = os.path.abspath(os.path.dirname(__file__))
HISTORY_PATH = os.path.join(APP_DIR, "data", "history.json")

root = tk.Tk()
root.title("Facture interne projet V1")
root.geometry("900x780")
root.minsize(760, 700)
root.configure(bg=COLORS["bg"])
root.resizable(True, True)

dossier_sortie = tk.StringVar(value="")
status_message = tk.StringVar(value="Choisissez un dossier de sortie pour commencer.")
label_dossier_value: tk.Label
header_subtitle: tk.Label
home_subtitle: tk.Label
form_desc_label: tk.Label
excel_desc_label: tk.Label
form_info_label: tk.Label
excel_info_label: tk.Label
logo_image = None
ligne_entries = []
entry_date_emission: tk.Entry
entry_date_du_jour: tk.Entry
history_listbox: tk.Listbox
history_data = []
history_filter_start: tk.Entry
history_filter_end: tk.Entry
history_filter_amount_min: tk.Entry
history_filter_amount_max: tk.Entry
history_filter_mode: tk.StringVar
history_filter_date: tk.Entry
history_filter_amount: tk.Entry
history_filter_title: tk.Entry


# Utilitaires
def set_status(message: str) -> None:
    status_message.set(message)


def require_dossier() -> bool:
    if not dossier_sortie.get():
        messagebox.showwarning("Dossier requis", "Merci de choisir un dossier de sortie avant de continuer.")
        return False
    return True


def load_logo():
    """Charge le logo JPEG depuis data/logo.jpg si Pillow est disponible."""
    logo_path = os.path.join(APP_DIR, "data", "logo.jpg")
    if not os.path.exists(logo_path):
        return None
    try:
        if Image is not None and ImageTk is not None:
            img = Image.open(logo_path)
            img.thumbnail((72, 72))
            return ImageTk.PhotoImage(img)
        # Fallback: PhotoImage peut lire certains JPEG selon la version de Tk
        return tk.PhotoImage(file=logo_path)
    except Exception:
        return None


def load_history():
    if not os.path.exists(HISTORY_PATH):
        return []
    try:
        with open(HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_history(data_list):
    os.makedirs(os.path.dirname(HISTORY_PATH), exist_ok=True)
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(data_list, f, ensure_ascii=False, indent=2)


def add_history_entry(path: str, meta: dict):
    entry = {
        "datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file": path,
        "meta": meta,
    }
    history_data.append(entry)
    save_history(history_data)
    refresh_history_ui()


def remove_history_selected():
    selection = history_listbox.curselection()
    if not selection:
        return
    idx = selection[0]
    if idx < 0 or idx >= len(history_data):
        return
    del history_data[idx]
    save_history(history_data)
    refresh_history_ui()


def refresh_history_ui():
    history_listbox.delete(0, "end")
    mode = history_filter_mode.get() if history_filter_mode else ""
    date_exacte = history_filter_date.get().strip() if history_filter_date else ""
    start_val = history_filter_start.get().strip() if history_filter_start else ""
    end_val = history_filter_end.get().strip() if history_filter_end else ""
    amount_exact = history_filter_amount.get().strip() if history_filter_amount else ""
    min_val = history_filter_amount_min.get().strip() if history_filter_amount_min else ""
    max_val = history_filter_amount_max.get().strip() if history_filter_amount_max else ""
    title_val = history_filter_title.get().strip().lower() if history_filter_title else ""

    def parse_filter_date(val: str):
        try:
            return datetime.strptime(val, "%d/%m/%Y").date()
        except Exception:
            return None

    def parse_amount(val: str):
        try:
            return float(val.replace(",", "."))
        except Exception:
            return None

    date_target = parse_filter_date(date_exacte) if date_exacte else None
    start_date = parse_filter_date(start_val) if start_val else None
    end_date = parse_filter_date(end_val) if end_val else None
    amount_target = parse_amount(amount_exact) if amount_exact else None
    min_amount = parse_amount(min_val) if min_val else None
    max_amount = parse_amount(max_val) if max_val else None

    for entry in history_data:
        file_name = os.path.basename(entry.get("file", ""))
        dt_str = entry.get("datetime", "")
        try:
            dt_obj = datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S").date()
        except Exception:
            dt_obj = None

        if start_date and dt_obj and dt_obj < start_date:
            continue
        if end_date and dt_obj and dt_obj > end_date:
            continue
        if date_target and dt_obj and dt_obj != date_target:
            continue

        meta = entry.get("meta", {})
        amount = meta.get("total_ht")
        if amount is None:
            amount = meta.get("somme_facture", 0)
        try:
            amount_val = float(str(amount).replace(",", "."))
        except Exception:
            amount_val = 0.0

        if amount_target is not None and abs(amount_val - amount_target) > 0.001:
            continue
        if min_amount is not None and amount_val < min_amount:
            continue
        if max_amount is not None and amount_val > max_amount:
            continue

        if title_val:
            search_zone = f"{file_name} {meta.get('numero_otfi','')} {meta.get('code_sous_projet','')} {meta.get('code_projet','')}".lower()
            if title_val not in search_zone:
                continue

        ref = meta.get("numero_otfi") or meta.get("code_sous_projet") or meta.get("code_projet") or ""
        amount_txt = f"{amount_val:.2f} FCFA" if amount_val else "-"
        display = f"{dt_str} | {ref} | {amount_txt} | {file_name}"
        history_listbox.insert("end", display)


def apply_history_filter():
    refresh_history_ui()


def clear_history_filter():
    history_filter_mode.set("date_exacte")
    for ent in [
        history_filter_date,
        history_filter_start,
        history_filter_end,
        history_filter_amount,
        history_filter_amount_min,
        history_filter_amount_max,
        history_filter_title,
    ]:
        if ent:
            ent.delete(0, "end")
    render_filter_fields()
    refresh_history_ui()


def export_history_excel():
    """Exporte l'historique des factures en Excel."""
    if not history_data:
        messagebox.showinfo("Historique", "Aucune facture dans l'historique.")
        return

    path = filedialog.asksaveasfilename(
        title="Exporter l'historique en Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        initialfile="historique_factures.xlsx",
    )
    if not path:
        return

    rows = []
    for entry in history_data:
        meta = entry.get("meta", {})
        amount = meta.get("total_ht")
        if amount is None:
            amount = meta.get("somme_facture", "")
        rows.append(
            {
                "date_heure": entry.get("datetime", ""),
                "fichier": entry.get("file", ""),
                "numero_otfi": meta.get("numero_otfi", ""),
                "code_sous_projet": meta.get("code_sous_projet", ""),
                "code_projet": meta.get("code_projet", ""),
                "somme_facture": meta.get("somme_facture", ""),
                "total_ht": amount,
            }
        )
    df = pd.DataFrame(rows)
    try:
        df.to_excel(path, index=False)
        messagebox.showinfo("Export Excel", f"Historique exporte :\n{path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d'exporter :\n{e}")


def export_excel_template():
    """Genere un fichier Excel vierge avec les colonnes attendues."""
    path = filedialog.asksaveasfilename(
        title="Enregistrer le template Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")],
        initialfile="template_factures.xlsx",
    )
    if not path:
        return

    colonnes = [
        "code_projet",
        "code_sous_projet",
        "numero_otfi",
        "pole_emettrice",
        "pole_destinataire",
        "dept_dir_emettrice",
        "dept_dir_destinataire",
        "date_emission",
        "date_du_jour",
        "periode_concernee",
        "somme_facture",
    ]

    # Colonnes pour 5 lignes de prestation
    for idx in range(1, 6):
        colonnes.extend(
            [
                f"ligne{idx}_designation",
                f"ligne{idx}_type_prestation",
                f"ligne{idx}_unite",
                f"ligne{idx}_quantite",
                f"ligne{idx}_prix_unitaire",
            ]
        )

    exemple = {
        "code_projet": "PRJ-001",
        "code_sous_projet": "SP-001",
        "numero_otfi": "OTFI-2025-001",
        "pole_emettrice": "Pole A",
        "pole_destinataire": "Pole B",
        "dept_dir_emettrice": "Direction X",
        "dept_dir_destinataire": "Direction Y",
        "date_emission": "01/01/2025",
        "date_du_jour": "01/01/2025",
        "periode_concernee": "Janvier 2025",
        "somme_facture": 120000,
    }
    exemple.update(
        {
            "ligne1_designation": "Prestation A",
            "ligne1_type_prestation": "Service",
            "ligne1_unite": "Lot",
            "ligne1_quantite": 2,
            "ligne1_prix_unitaire": 50000,
            "ligne2_designation": "Prestation B",
            "ligne2_type_prestation": "Support",
            "ligne2_unite": "H",
            "ligne2_quantite": 5,
            "ligne2_prix_unitaire": 15000,
        }
    )

    df = pd.DataFrame([exemple], columns=colonnes)
    try:
        df.to_excel(path, index=False)
        messagebox.showinfo("Template Excel", f"Template enregistre :\n{path}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de sauvegarder le template :\n{e}")


def open_date_picker(target_entry: tk.Entry, label: str):
    """
    Ouvre un date picker si tkcalendar est disponible, sinon renseigne la date du jour.
    """
    def set_value(value: str):
        target_entry.delete(0, "end")
        target_entry.insert(0, value)

    if Calendar is None:
        # Fallback simple : spinboxes jour/mois/annee
        top = tk.Toplevel(root)
        top.title(f"Choisir {label}")
        top.resizable(False, False)
        today = date.today()
        tk.Label(top, text="Jour / Mois / Annee", font=SUBTITLE_FONT).pack(pady=(8, 4))
        frame = tk.Frame(top)
        frame.pack(padx=10, pady=6)
        sb_day = tk.Spinbox(frame, from_=1, to=31, width=4, font=ENTRY_FONT)
        sb_month = tk.Spinbox(frame, from_=1, to=12, width=4, font=ENTRY_FONT)
        sb_year = tk.Spinbox(frame, from_=2000, to=2100, width=6, font=ENTRY_FONT)
        sb_day.delete(0, "end"); sb_day.insert(0, today.day)
        sb_month.delete(0, "end"); sb_month.insert(0, today.month)
        sb_year.delete(0, "end"); sb_year.insert(0, today.year)
        sb_day.grid(row=0, column=0, padx=4)
        sb_month.grid(row=0, column=1, padx=4)
        sb_year.grid(row=0, column=2, padx=4)

        def apply_spin():
            try:
                d = int(sb_day.get()); m = int(sb_month.get()); y = int(sb_year.get())
                value = f"{d:02d}/{m:02d}/{y}"
                set_value(value)
            except ValueError:
                set_value("")
            top.destroy()

        tk.Button(top, text="Valider", command=apply_spin, bg=COLORS["accent"], fg="white", bd=0, padx=10, pady=6).pack(
            pady=(6, 10)
        )
        top.grab_set()
        set_status("tkcalendar non installe : saisie avec selecteurs simples (format JJ/MM/AAAA).")
        return

    top = tk.Toplevel(root)
    top.title(f"Choisir {label}")
    top.resizable(False, False)
    cal = Calendar(top, selectmode="day", date_pattern="dd/mm/yyyy", locale="fr_FR")
    cal.pack(padx=10, pady=10)

    def apply_date():
        selected = cal.get_date()
        set_value(selected)
        top.destroy()

    tk.Button(top, text="Valider", command=apply_date, bg=COLORS["accent"], fg="white", bd=0, padx=10, pady=6).pack(
        pady=(0, 10)
    )
    top.grab_set()


# Actions coeur
def choisir_dossier() -> None:
    dossier = filedialog.askdirectory(title="Choisir le dossier de sortie")
    if dossier:
        dossier_sortie.set(dossier)
        label_dossier_value.config(text=dossier)
        set_status(f"Dossier selectionne : {dossier}")
    else:
        set_status("Aucun dossier selectionne.")


def generer_manuel() -> None:
    if not require_dossier():
        return

    data = {
        "code_projet": entry_code_projet.get().strip(),
        "code_sous_projet": entry_code_sous_projet.get().strip(),
        "numero_otfi": entry_numero_otfi.get().strip(),
        "pole_emettrice": entry_pole_emettrice.get().strip(),
        "pole_destinataire": entry_pole_destinataire.get().strip(),
        "dept_dir_emettrice": entry_dept_dir_emettrice.get().strip(),
        "dept_dir_destinataire": entry_dept_dir_destinataire.get().strip(),
        "date_emission": entry_date_emission.get().strip(),
        "date_du_jour": entry_date_du_jour.get().strip(),
        "periode_concernee": entry_periode_concernee.get().strip(),
        "somme_facture": entry_somme_facture.get().strip(),
    }

    if not data["code_sous_projet"] and not data["numero_otfi"] and not data["code_projet"]:
        messagebox.showwarning(
            "Champs manquants", "Renseignez au moins le code projet, le code sous-projet ou le numero OTFI."
        )
        return

    lignes = []
    for idx, row in enumerate(ligne_entries, start=1):
        designation = row["designation"].get().strip()
        type_prestation = row["type_prestation"].get().strip()
        unite = row["unite"].get().strip()
        qte_raw = row["quantite"].get().strip()
        pu_raw = row["prix_unitaire"].get().strip()

        if not any([designation, type_prestation, unite, qte_raw, pu_raw]):
            continue

        try:
            quantite = float(qte_raw.replace(",", ".")) if qte_raw else 0
        except ValueError:
            quantite = 0
        try:
            prix_unitaire = float(pu_raw.replace(",", ".")) if pu_raw else 0
        except ValueError:
            prix_unitaire = 0

        montant = quantite * prix_unitaire
        lignes.append(
            {
                "numero": idx,
                "designation": designation,
                "type_prestation": type_prestation,
                "unite": unite,
                "quantite": quantite,
                "prix_unitaire": prix_unitaire,
                "montant": montant,
            }
        )

    total_ht = sum(l["montant"] for l in lignes)
    data["lignes"] = lignes
    data["total_ht"] = total_ht

    chemin = generer_facture(data, dossier_sortie.get())
    add_history_entry(chemin, data)
    set_status(f"Facture enregistree : {chemin}")
    messagebox.showinfo("Facture generee", f"Facture enregistree dans :\n{chemin}")


def generer_depuis_excel() -> None:
    if not require_dossier():
        return

    fichier_excel = filedialog.askopenfilename(
        title="Selectionner le fichier Excel",
        filetypes=[("Excel", "*.xlsx")],
    )
    if not fichier_excel:
        set_status("Generation annulee : aucun fichier Excel selectionne.")
        return

    compteur = 0
    for data in charger_donnees_excel(fichier_excel):
        chemin = generer_facture(data, dossier_sortie.get())
        add_history_entry(chemin, data)
        compteur += 1

    set_status(f"{compteur} facture(s) generee(s) depuis {os.path.basename(fichier_excel)}.")
    messagebox.showinfo("Termine", f"Toutes les factures ont ete generees ({compteur}).")


# Navigation entre ecrans
def show_frame(frame: tk.Frame) -> None:
    for child in content_holder.winfo_children():
        child.pack_forget()
    frame.pack(fill="both", expand=True)
    content_canvas.yview_moveto(0)


# Habillage principal
container = tk.Frame(root, bg=COLORS["bg"])
container.pack(fill="both", expand=True, padx=20, pady=20)

header = tk.Frame(container, bg=COLORS["bg"])
header.pack(fill="x")
logo_image = load_logo()
header_row = tk.Frame(header, bg=COLORS["bg"])
header_row.pack(fill="x")
if logo_image:
    tk.Label(header_row, image=logo_image, bg=COLORS["bg"]).pack(side="left", padx=(0, 12))
else:
    tk.Label(header_row, text="INGRID", font=("Segoe UI", 12, "bold"), bg=COLORS["bg"], fg=COLORS["text"]).pack(
        side="left", padx=(0, 12)
    )
titles = tk.Frame(header_row, bg=COLORS["bg"])
titles.pack(side="left", fill="x", expand=True)
tk.Label(titles, text="Facture Interne Projet - Generator V1", font=TITLE_FONT, bg=COLORS["bg"], fg=COLORS["text"]).pack(anchor="w")
header_subtitle = tk.Label(
    titles,
    text="Assistant de generation de factures base sur votre template Word.",
    font=SUBTITLE_FONT,
    bg=COLORS["bg"],
    fg=COLORS["muted"],
)
header_subtitle.pack(anchor="w", pady=(2, 12))

output_card = tk.Frame(
    container,
    bg=COLORS["card"],
    bd=0,
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=14,
    pady=10,
)
output_card.pack(fill="x")
tk.Label(output_card, text="Dossier de sortie", font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
row_output = tk.Frame(output_card, bg=COLORS["card"])
row_output.pack(fill="x", pady=(6, 0))
tk.Button(
    row_output,
    text="Choisir...",
    command=choisir_dossier,
    bg=COLORS["warning"],
    fg="white",
    activebackground=COLORS["warning"],
    activeforeground="white",
    bd=0,
    padx=12,
    pady=8,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="left")
label_dossier_value = tk.Label(
    row_output,
    text="Aucun dossier selectionne",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=400,
    justify="left",
)
label_dossier_value.pack(side="left", padx=12)

scroll_area = tk.Frame(container, bg=COLORS["bg"])
scroll_area.pack(fill="both", expand=True, pady=(18, 10))

content_canvas = tk.Canvas(scroll_area, bg=COLORS["bg"], highlightthickness=0)
content_scrollbar = tk.Scrollbar(scroll_area, orient="vertical", command=content_canvas.yview)
content_canvas.configure(yscrollcommand=content_scrollbar.set)
content_scrollbar.pack(side="right", fill="y")
content_canvas.pack(side="left", fill="both", expand=True)

content_holder = tk.Frame(content_canvas, bg=COLORS["bg"])
content_window = content_canvas.create_window((0, 0), window=content_holder, anchor="nw")


def _on_holder_config(event):
    content_canvas.configure(scrollregion=content_canvas.bbox("all"))


def _on_canvas_config(event):
    content_canvas.itemconfig(content_window, width=event.width)


def _on_mousewheel(event):
    # Windows wheel delta increments by 120
    content_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


content_holder.bind("<Configure>", _on_holder_config)
content_canvas.bind("<Configure>", _on_canvas_config)
content_canvas.bind_all("<MouseWheel>", _on_mousewheel)

status_bar = tk.Frame(container, bg=COLORS["bg"])
status_bar.pack(fill="x", pady=(6, 0))
tk.Label(status_bar, textvariable=status_message, font=SUBTITLE_FONT, bg=COLORS["bg"], fg=COLORS["muted"]).pack(anchor="w")


# Ecran accueil
home_frame = tk.Frame(content_holder, bg=COLORS["bg"])
home_card = tk.Frame(
    home_frame,
    bg=COLORS["card"],
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=18,
    pady=18,
)
home_card.pack(fill="both", expand=True)
tk.Label(
    home_card,
    text="Bienvenue dans le generateur de factures - projet interne V1",
    font=("Segoe UI", 14, "bold"),
    bg=COLORS["card"],
    fg=COLORS["text"],
).pack(anchor="w")
home_subtitle = tk.Label(
    home_card,
    text="Choisissez le mode qui correspond a votre flux : saisie rapide ou import Excel.",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=520,
    justify="left",
)
home_subtitle.pack(anchor="w", pady=(6, 16))

tiles = tk.Frame(home_card, bg=COLORS["card"])
tiles.pack(fill="both", expand=True)
tiles.grid_columnconfigure(0, weight=1, uniform="tile")
tiles.grid_columnconfigure(1, weight=1, uniform="tile")

tile_formulaire = tk.Frame(
    tiles,
    bg=COLORS["card"],
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=14,
    pady=14,
)
tile_formulaire.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
tk.Label(tile_formulaire, text="Remplir le formulaire", font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
form_desc_label = tk.Label(
    tile_formulaire,
    text="Pour une facture unique, saisissez les informations client et generez en un clic.",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=230,
    justify="left",
)
form_desc_label.pack(anchor="w", pady=(4, 10))
tk.Button(
    tile_formulaire,
    text="Aller au formulaire",
    command=lambda: show_frame(form_frame),
    bg=COLORS["accent"],
    fg="white",
    activebackground=COLORS["accent"],
    activeforeground="white",
    bd=0,
    padx=12,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(anchor="e")

tile_excel = tk.Frame(
    tiles,
    bg=COLORS["card"],
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=14,
    pady=14,
)
tile_excel.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
tk.Label(tile_excel, text="Charger un fichier Excel", font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
excel_desc_label = tk.Label(
    tile_excel,
    text="Importez un fichier Excel et generez automatiquement toutes les factures listees.",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=230,
    justify="left",
)
excel_desc_label.pack(anchor="w", pady=(4, 10))
tk.Button(
    tile_excel,
    text="Lancer l'import Excel",
    command=lambda: show_frame(excel_frame),
    bg=COLORS["success"],
    fg="white",
    activebackground=COLORS["success"],
    activeforeground="white",
    bd=0,
    padx=12,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(anchor="e")

# Historique
history_card = tk.Frame(
    home_frame,
    bg=COLORS["card"],
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=18,
    pady=14,
)
history_card.pack(fill="both", expand=True, pady=(16, 0))
tk.Label(history_card, text="Historique des factures generees", font=("Segoe UI", 13, "bold"), bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
tk.Label(
    history_card,
    text="Liste des fichiers generes avec date et reference. Selectionnez une ligne pour la supprimer de l'historique.",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=520,
    justify="left",
).pack(anchor="w", pady=(4, 10))

filters_row = tk.Frame(history_card, bg=COLORS["card"])
filters_row.pack(fill="x", pady=(0, 10))
tk.Label(filters_row, text="Filtrer par", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
    row=0, column=0, sticky="w", padx=(0, 6)
)
history_filter_mode = tk.StringVar(value="date_exacte")
mode_options = [
    ("Date exacte", "date_exacte"),
    ("Intervalle de dates", "date_intervalle"),
    ("Montant exact", "montant_exact"),
    ("Intervalle de montant", "montant_intervalle"),
    ("Titre / reference", "titre"),
]
tk.OptionMenu(filters_row, history_filter_mode, *[opt[1] for opt in mode_options]).grid(row=0, column=1, padx=(0, 10), sticky="w")

dynamic_filter_area = tk.Frame(filters_row, bg=COLORS["card"])
dynamic_filter_area.grid(row=1, column=0, columnspan=8, sticky="w")

history_filter_date = None
history_filter_start = None
history_filter_end = None
history_filter_amount = None
history_filter_amount_min = None
history_filter_amount_max = None
history_filter_title = None


def render_filter_fields(*args):
    for child in dynamic_filter_area.winfo_children():
        child.destroy()

    mode = history_filter_mode.get()
    global history_filter_date, history_filter_start, history_filter_end
    global history_filter_amount, history_filter_amount_min, history_filter_amount_max, history_filter_title

    history_filter_date = history_filter_start = history_filter_end = None
    history_filter_amount = history_filter_amount_min = history_filter_amount_max = None
    history_filter_title = None

    row_base = 0
    if mode == "date_exacte":
        tk.Label(dynamic_filter_area, text="Date (JJ/MM/AAAA)", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=0, sticky="w", padx=(0, 6)
        )
        history_filter_date = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=14, bd=1, relief="solid", highlightthickness=0)
        history_filter_date.grid(row=row_base, column=1, padx=(0, 6))
        tk.Button(
            dynamic_filter_area,
            text="📅",
            command=lambda: open_date_picker(history_filter_date, "Date exacte historique"),
            bg="#e2e8f0",
            fg=COLORS["text"],
            activebackground="#e2e8f0",
            activeforeground=COLORS["text"],
            bd=0,
            padx=6,
            pady=1,
            cursor="hand2",
        ).grid(row=row_base, column=2, padx=(0, 10))
    elif mode == "date_intervalle":
        tk.Label(dynamic_filter_area, text="Date debut (JJ/MM/AAAA)", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=0, sticky="w", padx=(0, 6)
        )
        history_filter_start = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=14, bd=1, relief="solid", highlightthickness=0)
        history_filter_start.grid(row=row_base, column=1, padx=(0, 6))
        tk.Button(
            dynamic_filter_area,
            text="📅",
            command=lambda: open_date_picker(history_filter_start, "Date debut historique"),
            bg="#e2e8f0",
            fg=COLORS["text"],
            activebackground="#e2e8f0",
            activeforeground=COLORS["text"],
            bd=0,
            padx=6,
            pady=1,
            cursor="hand2",
        ).grid(row=row_base, column=2, padx=(0, 10))
        tk.Label(dynamic_filter_area, text="Date fin (JJ/MM/AAAA)", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=3, sticky="w", padx=(0, 6)
        )
        history_filter_end = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=14, bd=1, relief="solid", highlightthickness=0)
        history_filter_end.grid(row=row_base, column=4, padx=(0, 6))
        tk.Button(
            dynamic_filter_area,
            text="📅",
            command=lambda: open_date_picker(history_filter_end, "Date fin historique"),
            bg="#e2e8f0",
            fg=COLORS["text"],
            activebackground="#e2e8f0",
            activeforeground=COLORS["text"],
            bd=0,
            padx=6,
            pady=1,
            cursor="hand2",
        ).grid(row=row_base, column=5, padx=(0, 10))
    elif mode == "montant_exact":
        tk.Label(dynamic_filter_area, text="Montant exact (FCFA)", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=0, sticky="w", padx=(0, 6)
        )
        history_filter_amount = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=14, bd=1, relief="solid", highlightthickness=0)
        history_filter_amount.grid(row=row_base, column=1, padx=(0, 6))
    elif mode == "montant_intervalle":
        tk.Label(dynamic_filter_area, text="Montant min", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=0, sticky="w", padx=(0, 6)
        )
        history_filter_amount_min = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=14, bd=1, relief="solid", highlightthickness=0)
        history_filter_amount_min.grid(row=row_base, column=1, padx=(0, 10))
        tk.Label(dynamic_filter_area, text="Montant max", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=2, sticky="w", padx=(0, 6)
        )
        history_filter_amount_max = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=14, bd=1, relief="solid", highlightthickness=0)
        history_filter_amount_max.grid(row=row_base, column=3, padx=(0, 6))
    elif mode == "titre":
        tk.Label(dynamic_filter_area, text="Titre / reference (fichier ou code)", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
            row=row_base, column=0, sticky="w", padx=(0, 6)
        )
        history_filter_title = tk.Entry(dynamic_filter_area, font=ENTRY_FONT, width=26, bd=1, relief="solid", highlightthickness=0)
        history_filter_title.grid(row=row_base, column=1, padx=(0, 6))

    tk.Button(
        dynamic_filter_area,
        text="Filtrer",
        command=apply_history_filter,
        bg=COLORS["accent"],
        fg="white",
        activebackground=COLORS["accent"],
        activeforeground="white",
        bd=0,
        padx=10,
        pady=6,
        font=BUTTON_FONT,
        cursor="hand2",
    ).grid(row=row_base + 1, column=0, padx=(0, 6), pady=(8, 0), sticky="w")
    tk.Button(
        dynamic_filter_area,
        text="Reinitialiser filtre",
        command=clear_history_filter,
        bg="#e2e8f0",
        fg=COLORS["text"],
        activebackground="#e2e8f0",
        activeforeground=COLORS["text"],
        bd=0,
        padx=10,
        pady=6,
        font=BUTTON_FONT,
        cursor="hand2",
    ).grid(row=row_base + 1, column=1, pady=(8, 0), sticky="w")


history_filter_mode.trace_add("write", render_filter_fields)
render_filter_fields()

history_listbox = tk.Listbox(history_card, height=8, font=ENTRY_FONT, activestyle="dotbox")
history_scroll = tk.Scrollbar(history_card, orient="vertical", command=history_listbox.yview)
history_listbox.configure(yscrollcommand=history_scroll.set)
history_listbox.pack(side="left", fill="both", expand=True)
history_scroll.pack(side="right", fill="y")

tk.Button(
    history_card,
    text="Supprimer la selection",
    command=remove_history_selected,
    bg="#fee2e2",
    fg="#991b1b",
    activebackground="#fee2e2",
    activeforeground="#991b1b",
    bd=0,
    padx=12,
    pady=8,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(anchor="e", pady=(10, 0))
tk.Button(
    history_card,
    text="Exporter l'historique (Excel)",
    command=export_history_excel,
    bg="#e2e8f0",
    fg=COLORS["text"],
    activebackground="#e2e8f0",
    activeforeground=COLORS["text"],
    bd=0,
    padx=12,
    pady=8,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(anchor="e", pady=(6, 6))


# Ecran formulaire manuel
form_frame = tk.Frame(content_holder, bg=COLORS["bg"])
form_card = tk.Frame(
    form_frame,
    bg=COLORS["card"],
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=18,
    pady=18,
)
form_card.pack(fill="both", expand=True)
tk.Label(form_card, text="Formulaire client", font=("Segoe UI", 14, "bold"), bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
form_info_label = tk.Label(
    form_card,
    text="Completez les champs alignes avec le modele : code projet, code sous-projet, numero OTFI, poles, departements, dates, periode et somme facture. Ajoutez vos lignes de prestations ci-dessous, le montant et le total sont calcules automatiquement.",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=520,
    justify="left",
)
form_info_label.pack(anchor="w", pady=(4, 14))

fields_frame = tk.Frame(form_card, bg=COLORS["card"])
fields_frame.pack(fill="x")
fields_frame.grid_columnconfigure(0, weight=0)
fields_frame.grid_columnconfigure(1, weight=1)


def add_entry(row: int, label: str) -> tk.Entry:
    tk.Label(fields_frame, text=label, font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
        row=row, column=0, sticky="w", pady=6, padx=(0, 10)
    )
    entry = tk.Entry(fields_frame, font=ENTRY_FONT, bd=1, relief="solid", highlightthickness=0)
    entry.grid(row=row, column=1, sticky="ew", pady=6)
    return entry


def add_date_entry(row: int, label: str, target: str):
    tk.Label(fields_frame, text=label, font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["text"]).grid(
        row=row, column=0, sticky="w", pady=6, padx=(0, 10)
    )
    holder = tk.Frame(fields_frame, bg=COLORS["card"])
    holder.grid(row=row, column=1, sticky="ew", pady=6)
    holder.grid_columnconfigure(0, weight=1)
    entry = tk.Entry(holder, font=ENTRY_FONT, bd=1, relief="solid", highlightthickness=0)
    entry.grid(row=0, column=0, sticky="ew")
    btn = tk.Button(
        holder,
        text="📅",
        command=lambda: open_date_picker(entry, label),
        bg="#e2e8f0",
        fg=COLORS["text"],
        activebackground="#e2e8f0",
        activeforeground=COLORS["text"],
        bd=0,
        padx=8,
        pady=1,
        cursor="hand2",
    )
    btn.grid(row=0, column=1, padx=(6, 0))
    if target == "emission":
        global entry_date_emission
        entry_date_emission = entry
    else:
        global entry_date_du_jour
        entry_date_du_jour = entry
    return entry


entry_code_projet = add_entry(0, "Code projet")
entry_code_sous_projet = add_entry(1, "Code sous-projet")
entry_numero_otfi = add_entry(2, "Numero OTFI")
entry_pole_emettrice = add_entry(3, "Pole emettrice")
entry_pole_destinataire = add_entry(4, "Pole destinataire")
entry_dept_dir_emettrice = add_entry(5, "Departement / Direction emettrice")
entry_dept_dir_destinataire = add_entry(6, "Departement / Direction destinataire")
add_date_entry(7, "Date emission (JJ/MM/AAAA)", target="emission")
add_date_entry(8, "Date du jour (JJ/MM/AAAA)", target="jour")
entry_periode_concernee = add_entry(9, "Periode concernee")
entry_somme_facture = add_entry(10, "Somme facture")

tk.Label(form_card, text="Lignes de prestations", font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["text"]).pack(
    anchor="w", pady=(16, 8)
)
lines_table = tk.Frame(form_card, bg=COLORS["card"])
lines_table.pack(fill="both", expand=True)
headers = [
    ("Ligne", 0, 0),
    ("Designation", 1, 3),
    ("Type prestation", 2, 2),
    ("Unite", 3, 1),
    ("Quantite", 4, 1),
    ("Prix unitaire", 5, 2),
    ("Actions", 6, 0),
]
for text, col, weight in headers:
    lbl = tk.Label(lines_table, text=text, font=LABEL_FONT, bg=COLORS["card"], fg=COLORS["muted"])
    lbl.grid(row=0, column=col, padx=4, pady=(0, 2), sticky="w")
    lines_table.grid_columnconfigure(col, weight=weight)


def refresh_ligne_labels():
    for idx, row in enumerate(ligne_entries, start=1):
        row["label"].config(text=f"Ligne {idx}")


def remove_ligne_row(row_ref):
    if row_ref not in ligne_entries:
        return
    # sauvegarder les valeurs des autres lignes
    remaining_values = []
    for row in ligne_entries:
        if row is row_ref:
            continue
        remaining_values.append(
            {
                "designation": row["designation"].get(),
                "type_prestation": row["type_prestation"].get(),
                "unite": row["unite"].get(),
                "quantite": row["quantite"].get(),
                "prix_unitaire": row["prix_unitaire"].get(),
            }
        )
    # nettoyer
    clear_ligne_rows()
    ligne_entries.clear()
    # reconstruire avec les valeurs restantes
    for vals in remaining_values:
        add_ligne_row(prefill=vals)


def clear_ligne_rows():
    for child in list(lines_table.winfo_children()):
        info = child.grid_info()
        if info and int(info.get("row", 0)) > 0:
            child.destroy()


def add_ligne_row(prefill=None):
    row_index = len(ligne_entries) + 1
    row = row_index  # grid row (header est 0)

    label_num = tk.Label(lines_table, text=f"Ligne {row_index}", font=SUBTITLE_FONT, bg=COLORS["card"], fg=COLORS["muted"])
    label_num.grid(row=row, column=0, padx=(0, 8), sticky="w")

    def make_entry(col, width, value=""):
        entry = tk.Entry(lines_table, font=ENTRY_FONT, width=width, bd=1, relief="solid", highlightthickness=0)
        if value:
            entry.insert(0, value)
        entry.grid(row=row, column=col, padx=4, pady=2, sticky="ew")
        return entry

    designation = make_entry(1, 30, value=(prefill or {}).get("designation", ""))
    type_prestation = make_entry(2, 22, value=(prefill or {}).get("type_prestation", ""))
    unite = make_entry(3, 12, value=(prefill or {}).get("unite", ""))
    quantite = make_entry(4, 8, value=str((prefill or {}).get("quantite", "")))
    prix_unitaire = make_entry(5, 12, value=str((prefill or {}).get("prix_unitaire", "")))

    remove_btn = tk.Button(
        lines_table,
        text="Suppr.",
        command=lambda: remove_ligne_row(row_dict),
        bg="#fecdd3",
        fg="#9f1239",
        activebackground="#fecdd3",
        activeforeground="#9f1239",
        bd=0,
        padx=8,
        pady=4,
        font=("Segoe UI", 9, "bold"),
        cursor="hand2",
    )
    remove_btn.grid(row=row, column=6, padx=(6, 0))

    row_dict = {
        "label": label_num,
        "designation": designation,
        "type_prestation": type_prestation,
        "unite": unite,
        "quantite": quantite,
        "prix_unitaire": prix_unitaire,
    }
    ligne_entries.append(row_dict)


def reset_form():
    for entry in [
        entry_code_projet,
        entry_code_sous_projet,
        entry_numero_otfi,
        entry_pole_emettrice,
        entry_pole_destinataire,
        entry_dept_dir_emettrice,
        entry_dept_dir_destinataire,
        entry_periode_concernee,
        entry_somme_facture,
    ]:
        entry.delete(0, "end")
    entry_date_emission.delete(0, "end")
    entry_date_du_jour.delete(0, "end")

    clear_ligne_rows()
    ligne_entries.clear()
    for _ in range(3):
        add_ligne_row()

    set_status("Formulaire reinitialise. Choisissez un dossier et remplissez les informations.")


# init avec 3 lignes vides
for _ in range(3):
    add_ligne_row()

tk.Button(
    form_card,
    text="Ajouter une ligne",
    command=add_ligne_row,
    bg="#e2e8f0",
    fg=COLORS["text"],
    activebackground="#e2e8f0",
    activeforeground=COLORS["text"],
    bd=0,
    padx=12,
    pady=8,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(anchor="w", pady=(10, 0))

actions_form = tk.Frame(form_card, bg=COLORS["card"])
actions_form.pack(fill="x", pady=(18, 0))
tk.Button(
    actions_form,
    text="Reinitialiser",
    command=reset_form,
    bg="#e2e8f0",
    fg=COLORS["text"],
    activebackground="#e2e8f0",
    activeforeground=COLORS["text"],
    bd=0,
    padx=12,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="left", padx=(0, 8))

tk.Button(
    actions_form,
    text="Retour a l'accueil",
    command=lambda: show_frame(home_frame),
    bg="#cbd5e1",
    fg=COLORS["text"],
    activebackground="#cbd5e1",
    activeforeground=COLORS["text"],
    bd=0,
    padx=12,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="left")
tk.Button(
    actions_form,
    text="Generer la facture",
    command=generer_manuel,
    bg=COLORS["accent"],
    fg="white",
    activebackground=COLORS["accent"],
    activeforeground="white",
    bd=0,
    padx=14,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="right")


# Ecran import Excel
excel_frame = tk.Frame(content_holder, bg=COLORS["bg"])
excel_card = tk.Frame(
    excel_frame,
    bg=COLORS["card"],
    highlightbackground=COLORS["border"],
    highlightthickness=1,
    padx=18,
    pady=18,
)
excel_card.pack(fill="both", expand=True)
tk.Label(excel_card, text="Generation depuis Excel", font=("Segoe UI", 14, "bold"), bg=COLORS["card"], fg=COLORS["text"]).pack(anchor="w")
excel_info_label = tk.Label(
    excel_card,
    text="Importez un fichier Excel (.xlsx) avec les colonnes : code_projet, code_sous_projet, numero_otfi, pole_emettrice, pole_destinataire, dept_dir_emettrice, dept_dir_destinataire, date_emission, date_du_jour, periode_concernee, somme_facture.",
    font=SUBTITLE_FONT,
    bg=COLORS["card"],
    fg=COLORS["muted"],
    wraplength=520,
    justify="left",
)
excel_info_label.pack(anchor="w", pady=(4, 14))

excel_actions = tk.Frame(excel_card, bg=COLORS["card"])
excel_actions.pack(fill="x", pady=(10, 0))
tk.Button(
    excel_actions,
    text="Template de remplissage",
    command=export_excel_template,
    bg="#e2e8f0",
    fg=COLORS["text"],
    activebackground="#e2e8f0",
    activeforeground=COLORS["text"],
    bd=0,
    padx=12,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="left", padx=(0, 10))
tk.Button(
    excel_actions,
    text="Retour a l'accueil",
    command=lambda: show_frame(home_frame),
    bg="#cbd5e1",
    fg=COLORS["text"],
    activebackground="#cbd5e1",
    activeforeground=COLORS["text"],
    bd=0,
    padx=12,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="left")
tk.Button(
    excel_actions,
    text="Charger un fichier Excel et generer",
    command=generer_depuis_excel,
    bg=COLORS["success"],
    fg="white",
    activebackground=COLORS["success"],
    activeforeground="white",
    bd=0,
    padx=14,
    pady=10,
    font=BUTTON_FONT,
    cursor="hand2",
).pack(side="right")


def update_wraplength() -> None:
    width = content_canvas.winfo_width()
    if width <= 200:
        return

    main_width = max(width - 120, 260)
    card_width = max(width - 220, 240)
    tile_width = max((width // 2) - 60, 200)

    home_subtitle.config(wraplength=main_width)
    form_info_label.config(wraplength=main_width)
    excel_info_label.config(wraplength=main_width)
    label_dossier_value.config(wraplength=card_width)
    form_desc_label.config(wraplength=tile_width)
    excel_desc_label.config(wraplength=tile_width)


show_frame(home_frame)
update_wraplength()
history_data = load_history()
refresh_history_ui()
root.bind("<Configure>", lambda event: update_wraplength())
root.mainloop()
