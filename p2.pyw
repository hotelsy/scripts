import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from num2words import num2words
import json
import subprocess
import shutil
import sys

# =========================================================================
# --- Configuration et chemins ---
# =========================================================================

# ⭐ CHEMIN CRITIQUE : VOTRE PARTAGE RÉSEAU
BASE_CHEMIN_RESEAU = r"\\192.168.100.20\dossier partage de serveur\proforma"


#  Le chemin vers le fichier de prix doit pointer vers le même répertoire que p1.pyw et p2.pyw
# (déplacé dans settings/ ; p1.pyw l'y écrit). Repli sur l'ancien emplacement si pas encore migré.
FICHIER_PRIX_JSON = Path(__file__).resolve().parent / "settings" / "prix_chambres.json"
if not FICHIER_PRIX_JSON.exists() and (Path(__file__).resolve().parent / "prix_chambres.json").exists():
    FICHIER_PRIX_JSON = Path(__file__).resolve().parent / "prix_chambres.json"

# ⭐ Informations de l'entreprise émettrice (saisies au premier lancement)
# ⭐ Dossier "settings" (à côté du script) : contient tous les fichiers JSON de configuration
DOSSIER_SETTINGS = Path(__file__).resolve().parent / "settings"
FICHIER_ENTREPRISE_JSON = DOSSIER_SETTINGS / "entreprise.json"


# Tous les fichiers de travail et de configuration utiliseront ce chemin
DOSSIER_PROFORMA = Path(BASE_CHEMIN_RESEAU)
DOSSIER_ANCIENS = DOSSIER_PROFORMA / "anciens"  # Archive des anciennes proformas
FICHIER_CLIENTS_JSON = DOSSIER_SETTINGS / "clients_details.json"
FICHIER_NUMERO_PROFORMA_JSON = DOSSIER_SETTINGS / "proforma_counter.json"
FICHIER_NOM_PDF_TEMP_JSON = DOSSIER_SETTINGS / "last_proforma_pdf.json"

# ⭐ VÉRIFIEZ ET MODIFIEZ CE CHEMIN SI NÉCESSAIRE
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"






# --- Fonction utilitaire pour charger les prix HT ---
def charger_prix_config_puht():
    """Charge les prix PUHT sauvegardés par p1.pyw depuis le JSON."""
    try:
        with open(FICHIER_PRIX_JSON, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
            # Le JSON de p1.pyw stocke sous la forme : {"chambre_single": {"ttc": "12345", "puht": "11325.68"}}
            puht_config = {}
            for key, values in data.items():
                 # Tente de convertir en float (les valeurs sont stockées comme chaînes dans p1.pyw)
                 try:
                     puht_config[key] = float(values.get('puht', 0))
                 except (TypeError, ValueError):
                     puht_config[key] = 0.0
                 
            return puht_config

    except (FileNotFoundError, json.JSONDecodeError):
        # Afficher un avertissement si le fichier manque ou est corrompu, mais continue avec 0
        messagebox.showwarning("Avertissement", "Fichier de configuration des prix (prix_chambres.json) manquant ou corrompu. Les prix sont à 0.")
        return {}
# --- Fin de la fonction utilitaire ---


# --- Données de facturation (TARIFS CHARGÉS DYNAMIQUEMENT) ---

# ⭐ CHARGEMENT DYNAMIQUE DES PRIX HT
prix_ht_dynamique = charger_prix_config_puht()

# Les clés dans le dictionnaire prix_ht_dynamique proviennent de p1.pyw (ex: "chambre_single")
CHAMBRES_DATA = {
    "Ch. Single": {"puht": prix_ht_dynamique.get("chambre_single", 0), "taxe_sejour": 200},
    "Ch. Double": {"puht": prix_ht_dynamique.get("chambre_double", 0), "taxe_sejour": 400},
    "Ch. lit jumeaux": {"puht": prix_ht_dynamique.get("chambre_twin", 0), "taxe_sejour": 400},
    "Suite Junior.Single": {"puht": prix_ht_dynamique.get("suite_junior_single", 0), "taxe_sejour": 200},
    "Suite Junior.Double": {"puht": prix_ht_dynamique.get("suite_junior_double", 0), "taxe_sejour": 400},
    "Suite Sup.Single": {"puht": prix_ht_dynamique.get("suite_sup_single", 0), "taxe_sejour": 200},
    "Suite Sup.Double": {"puht": prix_ht_dynamique.get("suite_sup_double", 0), "taxe_sejour": 400},
}
TVA_TAUX = 0.09
PD_PUHT = 1500
LIT_SUP_PUHT = round(3500 / 1.09, 2)
DINER_TVA = 0.09
EXTRA_TVA = 0.09
DATE_FORMAT = "%d/%m/%Y"







# --- Variables globales de l'application ---
mode_multi_date = None
chambre_widgets = []
clients_details = {}
entreprise_infos = {}
# ⭐ NOUVEAU : Mémoriser le chemin du dernier DOCX généré
last_generated_docx_path = None
btn_modifier_pdf = None # Référence au bouton pour contrôle d'état
# ⭐ NOUVELLE VARIABLE GLOBALE POUR LE DROIT DE TIMBRE
droit_timbre_var = None

root = tk.Tk()
root.title("Générateur de Facture Proforma")

entreprise_var = tk.StringVar()
def validate_numeric_input(P):
    """Permet seulement les chiffres (ou vide) ou un seul point/virgule."""
    if P == "": return True
    try:
        float(P.replace(',', '.'))
        return True
    except ValueError:
        return False
    
vcmd = root.register(validate_numeric_input)

# =========================================================================
# --- Fonctions de Configuration et Utilitaires (Partie 1) ---
# =========================================================================

def show_splash(root_main):
    """Affiche l'écran de démarrage."""
    splash = tk.Toplevel(root_main)
    splash.title("Chargement...")
    splash.overrideredirect(True)
    
    TRANSPARENT_COLOR = '#FF00FF'
    splash.config(bg=TRANSPARENT_COLOR)

    try:
        # Assurez-vous que 'start.png' est dans le même répertoire
        img = tk.PhotoImage(file="start.png")
        label = tk.Label(splash, image=img, bg=TRANSPARENT_COLOR)
        label.image = img
    except tk.TclError:
        label = tk.Label(splash, text="Chargement de l'application...", font=("Arial", 14), width=40, height=10, bg='white')

    label.pack()

    try:
        splash.wm_attributes('-transparentcolor', TRANSPARENT_COLOR)
    except tk.TclError:
        if 'label' in locals():
            label.config(bg='white')

    w, h = label.winfo_reqwidth(), label.winfo_reqheight()
    screen_width = root_main.winfo_screenwidth()
    screen_height = root_main.winfo_screenheight()
    x = (screen_width // 2) - (w // 2)
    y = (screen_height // 2) - (h // 2)
    splash.geometry(f'{w}x{h}+{x}+{y}')

    splash.after(3000, splash.destroy)
    root_main.withdraw()
    root_main.after(3000, root_main.deiconify)

# =========================================================================
# --- Informations de l'entreprise (entreprise.json) ---
# =========================================================================

CHAMPS_ENTREPRISE = [
    # (clé, libellé affiché dans le formulaire, obligatoire)
    ("nom", "Nom de la société", True),
    ("adresse", "Adresse", False),
    ("tel", "Téléphone", False),
    ("mobile", "Mobile", False),
    ("email", "Email", False),
    ("rc", "N° RC", False),
    ("nif", "NIF", False),
    ("article", "N° Article", False),
    ("banque", "Compte N° (banque / agence)", False),
    ("rib", "RIB", False),
]

def migrer_anciens_fichiers():
    """Copie (une seule fois) les anciens fichiers JSON vers settings/ s'ils n'y sont pas encore.
    Les originaux ne sont pas supprimés. Garde la continuité du compteur et de la liste des clients."""
    anciens = [
        (DOSSIER_PROFORMA / "clients_details.json", FICHIER_CLIENTS_JSON),
        (DOSSIER_PROFORMA / "proforma_counter.json", FICHIER_NUMERO_PROFORMA_JSON),
        (Path(__file__).resolve().parent / "entreprise.json", FICHIER_ENTREPRISE_JSON),
    ]
    for ancien, nouveau in anciens:
        try:
            if ancien.is_file() and not nouveau.exists():
                DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
                shutil.copy2(str(ancien), str(nouveau))
        except Exception:
            pass

def charger_entreprise():
    """Charge entreprise.json. Retourne {} s'il est absent ou invalide."""
    try:
        with open(FICHIER_ENTREPRISE_JSON, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}
    except Exception as e:
        messagebox.showwarning("Avertissement", f"Impossible de lire entreprise.json. Détail: {e}")
        return {}

def sauvegarder_entreprise(data):
    try:
        DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
        with open(FICHIER_ENTREPRISE_JSON, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        messagebox.showerror("Erreur de sauvegarde", f"Impossible de sauvegarder entreprise.json: {e}")
        return False

def construire_bloc_entreprise(infos):
    """Construit le texte du bloc d'informations de l'entreprise (les champs vides sont ignorés)."""
    modeles = [
        ("adresse", "Adresse: {}"),
        ("tel", "Tel : {}"),
        ("mobile", "Mobile : {}"),
        ("email", "Email : {}"),
        ("rc", "N° RC : {}"),
        ("nif", "NIF : {}"),
        ("article", "N° Article : {}"),
        ("banque", "Compte N° : {}"),
        ("rib", "RIB : {}"),
    ]
    lignes = []
    for cle, modele in modeles:
        valeur = str(infos.get(cle, "")).strip()
        if valeur:
            lignes.append(modele.format(valeur))
    return "\n".join(lignes)

def demander_infos_entreprise(parent, initial=None, premiere_fois=False):
    """Fenêtre de saisie des informations de l'entreprise.
    Retourne le dictionnaire saisi, ou None si l'utilisateur annule."""
    initial = initial or {}
    resultat = {"data": None}

    dlg = tk.Toplevel(parent)
    dlg.title("Informations de l'entreprise")
    dlg.resizable(False, False)
    dlg.transient(parent)

    texte = ("Première ouverture : veuillez saisir les informations de l'entreprise\n"
             "qui émet les proformas. Elles seront enregistrées dans entreprise.json."
             if premiere_fois else
             "Modifiez les informations de l'entreprise émettrice.")
    tk.Label(dlg, text=texte, justify="left", font=("Arial", 10)).grid(
        row=0, column=0, columnspan=2, sticky="w", padx=15, pady=(12, 8))

    entrees = {}
    for i, (cle, libelle, obligatoire) in enumerate(CHAMPS_ENTREPRISE, start=1):
        tk.Label(dlg, text=libelle + (" *" if obligatoire else "") + " :").grid(
            row=i, column=0, sticky="e", padx=(15, 5), pady=3)
        e = tk.Entry(dlg, width=48)
        e.insert(0, str(initial.get(cle, "")))
        e.grid(row=i, column=1, sticky="we", padx=(0, 15), pady=3)
        entrees[cle] = e

    def valider():
        data = {cle: e.get().strip() for cle, e in entrees.items()}
        if not data["nom"]:
            messagebox.showwarning("Champ obligatoire", "Le nom de la société est obligatoire.", parent=dlg)
            return
        resultat["data"] = data
        dlg.destroy()

    def annuler():
        dlg.destroy()

    cadre = tk.Frame(dlg)
    cadre.grid(row=len(CHAMPS_ENTREPRISE) + 1, column=0, columnspan=2, pady=12)
    tk.Button(cadre, text="Enregistrer", command=valider, bg="#28a745", fg="white",
              font=("Arial", 10, "bold"), width=14).pack(side="left", padx=6)
    tk.Button(cadre, text="Quitter" if premiere_fois else "Annuler", command=annuler,
              width=14).pack(side="left", padx=6)

    dlg.protocol("WM_DELETE_WINDOW", annuler)
    dlg.bind("<Return>", lambda ev: valider())
    dlg.update_idletasks()
    x = (dlg.winfo_screenwidth() - dlg.winfo_reqwidth()) // 2
    y = (dlg.winfo_screenheight() - dlg.winfo_reqheight()) // 3
    dlg.geometry(f"+{x}+{y}")
    entrees["nom"].focus_set()
    dlg.grab_set()
    parent.wait_window(dlg)
    return resultat["data"]

def modifier_infos_entreprise():
    """Ouvre la fenêtre de modification depuis le menu Paramètres."""
    global entreprise_infos
    data = demander_infos_entreprise(root, entreprise_infos)
    if data and sauvegarder_entreprise(data):
        entreprise_infos = data
        messagebox.showinfo("Enregistré", "Les informations de l'entreprise ont été enregistrées.")

def charger_clients():
    global clients_details
    try:
        DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
        with open(FICHIER_CLIENTS_JSON, 'r', encoding='utf-8') as f:
            clients_details = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        clients_details = {}
    except Exception as e:
        messagebox.showwarning("Avertissement", f"Impossible de charger les données clients. Détail: {e}")
        clients_details = {}

def sauvegarder_clients():
    try:
        DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
        with open(FICHIER_CLIENTS_JSON, 'w', encoding='utf-8') as f:
            json.dump(clients_details, f, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("Erreur de sauvegarde", f"Impossible de sauvegarder les données clients: {e}")

def update_client_list():
    charger_clients()

def load_proforma_counter():
    """Charge le dernier numéro utilisé (un seul entier)."""
    try:
        DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
        with open(FICHIER_NUMERO_PROFORMA_JSON, 'r', encoding='utf-8') as f:
            return int(f.read().strip())
    except (FileNotFoundError, ValueError):
        return 0
    except Exception as e:
        messagebox.showwarning("Avertissement", f"Impossible de charger le compteur. Réinitialisation à 0. Détail: {e}")
        return 0

def save_proforma_counter(number):
    """Sauvegarde le numéro de proforma qui vient d'être utilisé."""
    try:
        DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
        with open(FICHIER_NUMERO_PROFORMA_JSON, 'w', encoding='utf-8') as f:
            f.write(str(number))
    except Exception as e:
        messagebox.showwarning("Avertissement de sauvegarde", f"Impossible de sauvegarder le compteur: {e}")

def get_next_proforma_data(increment=True):
    """Retourne le prochain numéro de proforma séquentiel."""
    last_number = load_proforma_counter()
    next_number = last_number + 1
    if increment:
        save_proforma_counter(next_number)
    return next_number

def archiver_anciennes_proformas():
    """Déplace les anciennes proformas (.docx / .pdf) du dossier principal vers le sous-dossier 'anciens'.
    À appeler juste avant d'enregistrer une NOUVELLE proforma, pour que le dossier principal
    ne contienne que la dernière. Les fichiers .json de configuration ne sont jamais déplacés."""
    echecs = []
    try:
        DOSSIER_ANCIENS.mkdir(parents=True, exist_ok=True)
        for f in DOSSIER_PROFORMA.iterdir():
            if not f.is_file():
                continue
            if f.suffix.lower() not in (".docx", ".pdf"):
                continue
            if f.name.startswith("~$"):  # fichier temporaire Word
                continue
            dest = DOSSIER_ANCIENS / f.name
            # Évite d'écraser un fichier déjà archivé portant le même nom
            if dest.exists():
                horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
                dest = DOSSIER_ANCIENS / f"{f.stem}_{horodatage}{f.suffix}"
            try:
                shutil.move(str(f), str(dest))
            except Exception:
                echecs.append(f.name)  # ex : fichier encore ouvert dans Word
    except Exception as e:
        messagebox.showwarning("Archivage", f"Impossible d'archiver les anciennes proformas : {e}")
        return

    if echecs:
        messagebox.showwarning(
            "Archivage incomplet",
            "Ces fichiers n'ont pas pu être déplacés vers 'anciens' (probablement ouverts dans Word ou un lecteur PDF) :\n\n"
            + "\n".join(echecs)
            + "\n\nFermez-les puis déplacez-les manuellement.")

def ouvrir_dossier():
    out = DOSSIER_PROFORMA
    out.mkdir(parents=True, exist_ok=True)
    try:
        os.startfile(out.resolve())
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d'ouvrir le dossier : {e}")

def ajouter_logo_entete(doc):
    section = doc.sections[0]
    paragraph = section.header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    
    logo_path = trouver_chemin_cachet("logo.png")
    
    if logo_path:
        try:
              # Utilisation de Inches pour la taille
              run.add_picture(logo_path, width=Inches(1.2))
        except Exception:
            run.add_text("[Logo manquant: Erreur de chargement. Vérifiez le fichier logo.png.]")
    else:
          run.add_text("[Logo manquant: logo.png non trouvé.]")

def trouver_chemin_cachet(image_path):
    """Tente de trouver le fichier d'image dans le dossier courant ou dans le chemin absolu du script."""
    if Path(image_path).is_file():
        return image_path
    
    try:
        # Tente de trouver le fichier dans le dossier du script
        script_dir = Path(__file__).resolve().parent
        chemin_absolu = script_dir / image_path
        
        if chemin_absolu.is_file():
            return str(chemin_absolu)
    except NameError:
        # Tente de trouver le fichier dans le répertoire de travail courant
        chemin_absolu = Path.cwd() / image_path
        if chemin_absolu.is_file():
            return str(chemin_absolu)
    
    return None

def set_cell_border(cell, **kwargs):
    """Ajoute des bordures personnalisées à une cellule DOCX. (Non utilisé dans le code fourni, mais utile pour l'esthétique)"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for edge in ('top', 'bottom', 'left', 'right'):
        if edge in kwargs:
            tag = 'w:{}'.format(edge)
            element = OxmlElement(tag)
            for key in kwargs[edge]:
                value = str(kwargs[edge][key])
                element.set(qn('w:{}'.format(key)), value)
            tcPr.append(element)

# --- Fenêtre de gestion des clients (Partie 1) ---
def gerer_clients():
    """Ouvre une fenêtre pour ajouter/modifier/supprimer les profils clients."""
    
    def on_selection(event):
        try:
            selection_indices = list_clients.curselection()
            if not selection_indices: return
            selection = list_clients.get(selection_indices[0])
            details = clients_details.get(selection, {})
            
            for entry in [entry_nom, entry_nif, entry_nis, entry_article, entry_rc, entry_adresse, entry_code_postal, entry_wilaya, entry_pays]:
                entry.delete(0, tk.END)

            if details:
                entry_nom.insert(0, selection)
                entry_nif.insert(0, details.get('nif', ''))
                entry_nis.insert(0, details.get('nis', ''))
                entry_article.insert(0, details.get('article', ''))
                entry_rc.insert(0, details.get('rc', ''))
                entry_adresse.insert(0, details.get('adresse', ''))
                entry_code_postal.insert(0, details.get('code_postal', ''))
                entry_wilaya.insert(0, details.get('wilaya', ''))
                entry_pays.insert(0, details.get('pays', 'Algérie'))
            else:
                entry_pays.insert(0, "Algérie")
        except tk.TclError:
            pass

    def sauvegarder_profil():
        nom = entry_nom.get().strip()
        if not nom:
            messagebox.showerror("Erreur", "Le nom de l'entreprise est obligatoire.")
            return

        clients_details[nom] = {
            'nif': entry_nif.get().strip(),
            'nis': entry_nis.get().strip(),
            'article': entry_article.get().strip(),
            'rc': entry_rc.get().strip(),
            'adresse': entry_adresse.get().strip(),
            'code_postal': entry_code_postal.get().strip(),
            'wilaya': entry_wilaya.get().strip(),
            'pays': entry_pays.get().strip()
        }
        sauvegarder_clients()
        update_listbox()
        update_client_list()
        messagebox.showinfo("Succès", f"Le profil '{nom}' a été sauvegardé.")

    def supprimer_profil():
        try:
            selection = list_clients.get(list_clients.curselection())
            if selection:
                if messagebox.askyesno("Confirmer", f"Êtes-vous sûr de vouloir supprimer le profil '{selection}' ?"):
                    del clients_details[selection]
                    sauvegarder_clients()
                    update_listbox()
                    update_client_list()
                    reinitialiser_champs()
        except tk.TclError:
            messagebox.showwarning("Attention", "Veuillez sélectionner un profil à supprimer.")

    def choisir_profil():
        try:
            selection = list_clients.get(list_clients.curselection())
            if selection:
                entreprise_var.set(selection)
                client_window.destroy()
            else:
                messagebox.showwarning("Attention", "Veuillez sélectionner un profil.")
        except tk.TclError:
            messagebox.showwarning("Attention", "Veuillez sélectionner un profil.")

    def reinitialiser_champs():
        for entry in [entry_nom, entry_nif, entry_nis, entry_article, entry_rc, entry_adresse, entry_code_postal, entry_wilaya, entry_pays]:
            entry.delete(0, tk.END)
        entry_pays.insert(0, "Algérie")

    def update_listbox(event=None):
        list_clients.delete(0, tk.END)
        filtre = entry_recherche.get().strip().lower()
        charger_clients()
        
        for nom in sorted(clients_details.keys()):
            if not filtre or filtre in nom.lower():
                list_clients.insert(tk.END, nom)

    client_window = tk.Toplevel(root)
    client_window.title("Gérer les profils d'entreprises")
    client_window.transient(root)
    
    window_width = 550
    window_height = 650
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    client_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    frame_saisie = tk.LabelFrame(client_window, text="Détails du Profil (Création/Modification)", padx=10, pady=10)
    frame_saisie.pack(pady=10, padx=10, fill="x")

    tk.Label(frame_saisie, text="Nom Entreprise:").grid(row=0, column=0, sticky='w', pady=2)
    entry_nom = tk.Entry(frame_saisie, width=40)
    entry_nom.grid(row=0, column=1, sticky='we', padx=5, pady=2)
    
    tk.Label(frame_saisie, text="NIF:").grid(row=1, column=0, sticky='w', pady=2)
    entry_nif = tk.Entry(frame_saisie, width=40)
    entry_nif.grid(row=1, column=1, sticky='we', padx=5, pady=2)

    tk.Label(frame_saisie, text="NIS:").grid(row=2, column=0, sticky='w', pady=2)
    entry_nis = tk.Entry(frame_saisie, width=40)
    entry_nis.grid(row=2, column=1, sticky='we', padx=5, pady=2)

    tk.Label(frame_saisie, text="N° Article:").grid(row=3, column=0, sticky='w', pady=2)
    entry_article = tk.Entry(frame_saisie, width=40)
    entry_article.grid(row=3, column=1, sticky='we', padx=5, pady=2)

    tk.Label(frame_saisie, text="RC:").grid(row=4, column=0, sticky='w', pady=2)
    entry_rc = tk.Entry(frame_saisie, width=40)
    entry_rc.grid(row=4, column=1, sticky='we', padx=5, pady=2)
    
    tk.Label(frame_saisie, text="Adresse (Rue, etc.):").grid(row=5, column=0, sticky='w', pady=2)
    entry_adresse = tk.Entry(frame_saisie, width=40)
    entry_adresse.grid(row=5, column=1, sticky='we', padx=5, pady=2)
    
    tk.Label(frame_saisie, text="Code Postal:").grid(row=6, column=0, sticky='w', pady=2)
    entry_code_postal = tk.Entry(frame_saisie, width=40)
    entry_code_postal.grid(row=6, column=1, sticky='we', padx=5, pady=2)
    
    tk.Label(frame_saisie, text="Wilaya:").grid(row=7, column=0, sticky='w', pady=2)
    entry_wilaya = tk.Entry(frame_saisie, width=40)
    entry_wilaya.grid(row=7, column=1, sticky='we', padx=5, pady=2)

    tk.Label(frame_saisie, text="Pays:").grid(row=8, column=0, sticky='w', pady=2)
    entry_pays = tk.Entry(frame_saisie, width=40)
    entry_pays.insert(0, "Algérie")
    entry_pays.grid(row=8, column=1, sticky='we', padx=5, pady=2)
    
    frame_buttons = tk.Frame(frame_saisie)
    frame_buttons.grid(row=9, column=0, columnspan=2, pady=10)
    tk.Button(frame_buttons, text="💾 Sauvegarder", command=sauvegarder_profil).pack(side='left', padx=5)
    tk.Button(frame_buttons, text="✨ Nouveau", command=reinitialiser_champs).pack(side='left', padx=5)

    frame_liste = tk.LabelFrame(client_window, text="Profils Existants", padx=10, pady=5)
    frame_liste.pack(padx=10, pady=5, fill="both", expand=True)
    
    recherche_frame = tk.Frame(frame_liste)
    recherche_frame.pack(fill='x', pady=(0, 5))
    tk.Label(recherche_frame, text="Rechercher:").pack(side='left')
    
    entry_recherche = tk.Entry(recherche_frame, width=30)
    entry_recherche.pack(side='left', fill='x', expand=True, padx=5)
    entry_recherche.bind("<KeyRelease>", update_listbox)

    list_clients = tk.Listbox(frame_liste, height=8)
    list_clients.pack(side="left", fill="both", expand=True)
    list_clients.bind('<<ListboxSelect>>', on_selection)

    scrollbar = tk.Scrollbar(frame_liste, orient="vertical")
    scrollbar.config(command=list_clients.yview)
    scrollbar.pack(side="right", fill="y")
    list_clients.config(yscrollcommand=scrollbar.set)
    
    action_buttons_frame = tk.Frame(frame_liste)
    action_buttons_frame.pack(fill='x', pady=5)
    
    btn_delete = tk.Button(action_buttons_frame, text="❌ Supprimer le Profil", fg="red", command=supprimer_profil)
    btn_delete.pack(side='left', expand=True, padx=2)
    
    btn_choisir = tk.Button(action_buttons_frame, text="✅ Choisir ce profil", command=choisir_profil, bg="green", fg="white", font=("Arial", 10, "bold"))
    btn_choisir.pack(side='right', expand=True, padx=2)
    
    update_listbox()

# =========================================================================
# --- Logique de l'interface (Chambres, Dates, Compteurs) (Partie 2) ---
# =========================================================================

def calculer_total_nuits_gui():
    """Calcule le nombre total de nuitées facturées."""
    total_nuits = 0
    global_arr_str = entry_arrivee.get().strip()
    global_dep_str = entry_depart.get().strip()
    
    total_chambres_reel = 0
    try:
        total_chambres_reel = sum(int((c['qte_entry'].get() or "1").split('.')[0]) for c in chambre_widgets)
    except:
        total_chambres_reel = 0

    global_nuits = 0
    try:
        d1 = datetime.strptime(global_arr_str, DATE_FORMAT).date()
        d2 = datetime.strptime(global_dep_str, DATE_FORMAT).date()
        global_nuits = (d2 - d1).days
    except:
        global_nuits = 0
        
    if global_nuits < 0: global_nuits = 0

    if mode_multi_date.get():
        for i, ch in enumerate(chambre_widgets):
            try:
                qte_similaire = int((ch['qte_entry'].get() or "1").split('.')[0])
            except:
                qte_similaire = 1
                
            if i == 0:
                nb_nuits = global_nuits
            else:
                try:
                    d_arr_local_str = ch['date_arrivee'].get().strip()
                    d_dep_local_str = ch['date_depart'].get().strip()
                    d_arr_local = datetime.strptime(d_arr_local_str, DATE_FORMAT).date()
                    d_dep_local = datetime.strptime(d_dep_local_str, DATE_FORMAT).date()
                    
                    nb_nuits = (d_dep_local - d_arr_local).days
                    if nb_nuits < 0: nb_nuits = 0
                except:
                    nb_nuits = 0
            
            if nb_nuits > 0:
                total_nuits += nb_nuits * qte_similaire
                
    else:
        if global_nuits > 0:
            total_nuits = global_nuits * total_chambres_reel
            
    return total_nuits

def update_compteurs_gui():
    """Met à jour les labels affichant le total des chambres et des nuitées."""
    total_chambres = 0
    for c in chambre_widgets:
        try:
            qte = int((c['qte_entry'].get() or "1").split('.')[0])
            total_chambres += qte if qte > 0 else 1
        except ValueError:
            total_chambres += 0
            
    for i, ch in enumerate(chambre_widgets):
        ch["label"].config(text=f"Chambre {i+1}")
        ch["btn_remove"].config(state="disabled" if i == 0 else "normal")
            
    compteur_chambres_label.config(text=f"Nombre de chambres : {total_chambres}")
    
    total_nuits = calculer_total_nuits_gui()
    compteur_nuits_label.config(text=f"Nombre de nuitées : {total_nuits}")

def update_all_dates(event=None):
    """Synchronise les dates globales avec la première chambre."""
    global_arr = entry_arrivee.get().strip()
    global_dep = entry_depart.get().strip()
    
    if chambre_widgets:
        ch = chambre_widgets[0]
        ch['date_arrivee'].config(state="normal")
        ch['date_depart'].config(state="normal")
        
        ch['date_arrivee'].delete(0, tk.END)
        ch['date_arrivee'].insert(0, global_arr)
        ch['date_depart'].delete(0, tk.END)
        ch['date_depart'].insert(0, global_dep)
        
        ch['date_arrivee'].config(state="disabled")
        ch['date_depart'].config(state="disabled")
        
    update_compteurs_gui()

def pre_remplir_dates_chambre(ch):
    """Pré-remplit les dates d'une chambre avec celles de la 1ère chambre (dates globales),
    uniquement si les champs sont vides. L'utilisateur peut ensuite les modifier."""
    global_arr = entry_arrivee.get().strip()
    global_dep = entry_depart.get().strip()

    ch['date_arrivee'].config(state="normal")
    ch['date_depart'].config(state="normal")

    if not ch['date_arrivee'].get().strip() and global_arr:
        ch['date_arrivee'].delete(0, tk.END)
        ch['date_arrivee'].insert(0, global_arr)
    if not ch['date_depart'].get().strip() and global_dep:
        ch['date_depart'].delete(0, tk.END)
        ch['date_depart'].insert(0, global_dep)

def toggle_multi_dates(prefill=True):
    """Active ou désactive les champs de date individuels.
    Quand le mode est activé, les chambres 2, 3... reprennent par défaut les dates de la chambre 1."""
    is_multi = mode_multi_date.get()
    
    global_arr = entry_arrivee.get().strip()
    global_dep = entry_depart.get().strip()

    for i, ch in enumerate(chambre_widgets):
        
        if is_multi and i > 0:
            ch['dates_frame'].grid()
        else:
            ch['dates_frame'].grid_remove()
            
        if i == 0:
            ch['date_arrivee'].config(state="normal")
            ch['date_depart'].config(state="normal")
            
            ch['date_arrivee'].delete(0, tk.END)
            ch['date_arrivee'].insert(0, global_arr)
            ch['date_depart'].delete(0, tk.END)
            ch['date_depart'].insert(0, global_dep)
            
            ch['date_arrivee'].config(state="disabled")
            ch['date_depart'].config(state="disabled")
            
        else:
            if is_multi:
                ch['date_arrivee'].config(state="normal")
                ch['date_depart'].config(state="normal")
                if prefill:
                    pre_remplir_dates_chambre(ch)
            else:
                ch['date_arrivee'].config(state="disabled")
                ch['date_depart'].config(state="disabled")

    update_compteurs_gui()

# =========================================================================
# --- Fonctions Dynamiques de l'Interface (Partie 3) ---
# =========================================================================

def remove_extra(frame, extra_data):
    """Supprime un champ extra d'une chambre."""
    for ch in chambre_widgets:
        # Remonter jusqu'au frame parent du conteneur des extras
        if ch['frame'] == frame.master.master:
            ch['extras_list'] = [e for e in ch['extras_list'] if e != extra_data]
            break
    frame.destroy()

def ajouter_extra_field(parent_frame, extras_list):
    """Ajoute un champ de saisie pour un extra (libellé, montant, par nuit)."""
    extra_frame = tk.Frame(parent_frame)
    
    tk.Label(extra_frame, text="Libellé:").pack(side='left', padx=(5, 0))
    libelle_entry = tk.Entry(extra_frame, width=15)
    libelle_entry.insert(0, "Divers")
    libelle_entry.pack(side='left', padx=(5, 10))

    tk.Label(extra_frame, text="Montant TTC:").pack(side='left')
    montant_entry = tk.Entry(extra_frame, width=8, validate='key', vcmd=(vcmd, '%P'))
    montant_entry.insert(0, "0")
    montant_entry.pack(side='left', padx=(5, 5))
    
    is_per_night_var = tk.BooleanVar(value=False)
    per_night_check = tk.Checkbutton(extra_frame, text="Par Nuit", variable=is_per_night_var)
    per_night_check.pack(side='left', padx=(10, 5))
    
    extra_data = {
        "libelle": libelle_entry,
        "montant": montant_entry,
        "is_per_night": is_per_night_var
    }
    btn_remove_extra = tk.Button(extra_frame, text="x", fg="red", width=2, command=lambda f=extra_frame, d=extra_data: remove_extra(f, d))
    btn_remove_extra.pack(side='left', padx=(5, 0))

    extra_frame.grid(row=len(parent_frame.grid_slaves()), column=0, columnspan=5, sticky='w', pady=1)
    
    extras_list.append(extra_data)

def remove_chambre(frame):
    """Supprime un bloc de chambre de l'interface."""
    global chambre_widgets
    chambre_widgets = [c for c in chambre_widgets if c["frame"] != frame]
    frame.destroy()
    update_compteurs_gui()
    toggle_multi_dates()

def ajouter_chambre():
    """Ajoute un nouveau bloc de saisie de chambre."""
    frame = tk.Frame(chambres_container, bd=1, relief=tk.SOLID, padx=6, pady=6)
    frame.pack(pady=5, fill='x')

    frame.grid_columnconfigure(0, weight=0)
    frame.grid_columnconfigure(1, weight=0)

    header = tk.Frame(frame)
    header.grid(row=0, column=0, columnspan=12, sticky='w')
    label = tk.Label(header, text="Chambre", font=("Arial", 10, "bold"))
    label.pack(side='left')
    btn_remove = tk.Button(header, text="-", fg="red", width=2, command=lambda f=frame: remove_chambre(f))
    btn_remove.pack(side='right')

    tk.Label(frame, text="Type:").grid(row=1, column=0, sticky='e')
    type_var = tk.StringVar()
    combo = ttk.Combobox(frame, textvariable=type_var, values=list(CHAMBRES_DATA.keys()), width=24)
    combo.current(0)
    combo.grid(row=1, column=1, sticky='w')

    tk.Label(frame, text="Lits suppl.:").grid(row=1, column=2, sticky='e')
    lits_entry = tk.Entry(frame, width=5, validate='key', vcmd=(vcmd, '%P'))
    lits_entry.insert(0, "0")
    lits_entry.grid(row=1, column=3, sticky='w')

    tk.Label(frame, text="Dîner TTC:").grid(row=1, column=4, sticky='e')
    diner_entry = tk.Entry(frame, width=8, validate='key', vcmd=(vcmd, '%P'))
    diner_entry.insert(0, "0")
    diner_entry.grid(row=1, column=5, sticky='w', padx=(0, 15))
    
    tk.Label(frame, text="Qté chambres similaires:").grid(row=1, column=6, sticky='e')
    qte_entry = tk.Entry(frame, width=5, validate='key', vcmd=(vcmd, '%P'))
    qte_entry.insert(0, "1")
    qte_entry.grid(row=1, column=7, sticky='w')
    
    # Liens pour la mise à jour des compteurs
    qte_entry.bind("<KeyRelease>", lambda event: update_compteurs_gui())
    qte_entry.bind("<FocusOut>", lambda event: update_compteurs_gui())
    combo.bind("<<ComboboxSelected>>", lambda event: update_compteurs_gui())
    
    # Cadre des dates individuelles (caché par défaut pour les chambres > 1)
    dates_frame = tk.Frame(frame)
    tk.Label(dates_frame, text="Arrivée (JJ/MM/AAAA):").grid(row=0, column=0)
    
    d_arr = tk.Entry(dates_frame, width=15)
    d_arr.grid(row=0, column=1, padx=3)
    d_arr.bind("<KeyRelease>", lambda event: update_compteurs_gui())
    d_arr.bind("<FocusOut>", lambda event: update_compteurs_gui())
    
    tk.Label(dates_frame, text="Départ (JJ/MM/AAAA):").grid(row=0, column=2)
    
    d_dep = tk.Entry(dates_frame, width=15)
    d_dep.grid(row=0, column=3, padx=3)
    d_dep.bind("<KeyRelease>", lambda event: update_compteurs_gui())
    d_dep.bind("<FocusOut>", lambda event: update_compteurs_gui())
    
    dates_frame.grid(row=2, column=0, columnspan=6, pady=4, sticky='w')

    # Cadre des Extras
    extras_frame_container = tk.LabelFrame(frame, text="Extras", padx=5, pady=5)
    extras_frame_container.grid(row=1, column=8, columnspan=4, rowspan=3, sticky='nsw', padx=(10,0))

    extras_fields_container = tk.Frame(extras_frame_container)
    extras_fields_container.pack(side='top', fill='x')

    extras_list = []
    
    btn_add_extra = tk.Button(extras_fields_container, text="+ Ajouter un extra", command=lambda f=extras_fields_container, l=extras_list: ajouter_extra_field(f, l))
    btn_add_extra.grid(row=0, column=0, columnspan=5, sticky='w', pady=(3, 0))

    # Initialisation de la première chambre (qui suit les dates globales)
    index = len(chambre_widgets)
    if index == 0:
        if entry_arrivee.get():
            d_arr.delete(0, tk.END)
            d_arr.insert(0, entry_arrivee.get())
        if entry_depart.get():
            d_dep.delete(0, tk.END)
            d_dep.insert(0, entry_depart.get())
            
        d_arr.config(state="disabled")
        d_dep.config(state="disabled")

    chambre_widgets.append({
        "frame": frame, "label": label, "btn_remove": btn_remove, "type_var": type_var,
        "lits_entry": lits_entry, "diner_entry": diner_entry, "dates_frame": dates_frame,
        "date_arrivee": d_arr, "date_depart": d_dep, "qte_entry": qte_entry,
        "extras_list": extras_list
    })

    toggle_multi_dates(prefill=False)

    # Nouvelle chambre (2, 3...) en mode "Dates différentes" : dates de la chambre 1 par défaut
    if index > 0 and mode_multi_date.get():
        pre_remplir_dates_chambre(chambre_widgets[-1])

    update_compteurs_gui()

# =========================================================================
# --- FONCTION PRINCIPALE DE GÉNÉRATION DU DOCUMENT WORD (.docx) ---
# =========================================================================

def calculer_droit_timbre(montant_ttc):
    """Calcule le droit de timbre basé sur le Total TTC (après remise)."""
    if montant_ttc < 30000:
        taux = 0.01
    elif montant_ttc < 100000:
        taux = 0.015
    else: # montant_ttc >= 100000
        taux = 0.02
        
    return round(montant_ttc * taux, 2)

def generer_document_docx(numero_proforma, increment_counter=True):
    """
    Génère le contenu du document Word.
    Retourne (doc, nom_fichier_docx) si succès, ou lève une exception.
    """
    global global_nuits # Déclaré globalement même si non modifié ici.
    global_arr = entry_arrivee.get().strip()
    global_dep = entry_depart.get().strip()
    global_nuits = 0
    
    # 1. Validation des dates globales
    try:
        if not global_arr or not global_dep:
            raise ValueError("Veuillez entrer une date d'arrivée et de départ globale valide au format JJ/MM/AAAA.")
            
        d1 = datetime.strptime(global_arr, DATE_FORMAT).date()
        d2 = datetime.strptime(global_dep, DATE_FORMAT).date()
        
        if (d2 - d1).days <= 0:
            raise ValueError("La date de départ doit être après la date d'arrivée globale.")
        global_nuits = (d2 - d1).days
        
    except ValueError as e:
        messagebox.showerror("Erreur de Date", str(e))
        return None, None
    except Exception:
        messagebox.showerror("Erreur", "Une erreur inattendue est survenue lors de la validation des dates.")
        return None, None

    # 2. Validation des données clients et chambres
    selected_name = entreprise_var.get().strip()
    details = clients_details.get(selected_name, {})
    client_contact = entry_client.get().strip()

    if not selected_name and not client_contact:
        messagebox.showerror("Erreur", "Veuillez sélectionner ou entrer le nom d'une entreprise/un client.")
        return None, None
    
    if len(chambre_widgets) == 0:
        messagebox.showerror("Erreur", "Veuillez ajouter au moins une chambre.")
        return None, None

    # 2b. Validation des dates de chaque chambre (mode "Dates différentes par chambre")
    if mode_multi_date.get():
        for i, ch in enumerate(chambre_widgets):
            if i == 0:
                continue  # La chambre 1 suit les dates globales (déjà validées)
            arr_str = ch['date_arrivee'].get().strip()
            dep_str = ch['date_depart'].get().strip()

            if not arr_str or not dep_str:
                messagebox.showerror("Erreur de Date", f"Chambre {i+1} : veuillez saisir la date d'arrivée et la date de départ (JJ/MM/AAAA).")
                (ch['date_arrivee'] if not arr_str else ch['date_depart']).focus_set()
                return None, None

            try:
                d_arr_ch = datetime.strptime(arr_str, DATE_FORMAT).date()
                d_dep_ch = datetime.strptime(dep_str, DATE_FORMAT).date()
            except ValueError:
                messagebox.showerror("Erreur de Date", f"Chambre {i+1} : format de date invalide. Utilisez JJ/MM/AAAA.")
                ch['date_arrivee'].focus_set()
                return None, None

            if (d_dep_ch - d_arr_ch).days <= 0:
                messagebox.showerror("Erreur de Date", f"Chambre {i+1} : la date de départ doit être après la date d'arrivée.")
                ch['date_depart'].focus_set()
                return None, None

    # 3. Logique de numérotation séquentielle
    if increment_counter:
        num = get_next_proforma_data(increment=True)
    else:
        # Ne pas incrémenter si c'est seulement pour la prévisualisation (non utilisé ici)
        num = get_next_proforma_data(increment=False)

    numero_formatte = str(num)
    
    # 4. Création du document
    doc = Document()
    ajouter_logo_entete(doc)
    
    p_title = doc.add_paragraph(entreprise_infos.get("nom", ""))
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.runs[0]
    run.font.name = 'Baskerville Old Face'
    run.font.size = Pt(20)
    run.bold = True

    # Titre Proforma
    date_aujourdhui = datetime.now().strftime(DATE_FORMAT)
    p_facture = doc.add_paragraph()
    run_date = p_facture.add_run(f"Facture Proforma : du {date_aujourdhui}")
    p_facture.add_run(u'\u00A0' * 5)
    run_num = p_facture.add_run(f"N°:{numero_formatte}")
    p_facture.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in [run_date, run_num]:
        run.bold = True
        run.font.size = Pt(14)

    # Tableau d'informations Société/Client
    t_info = doc.add_table(rows=1, cols=2)
    t_info.style = 'Table Grid'
    t_info.cell(0, 0).text = construire_bloc_entreprise(entreprise_infos)

    cell2 = t_info.cell(0, 1)
    
    total_chambres_reel = sum(int((c['qte_entry'].get() or "1").split('.')[0]) for c in chambre_widgets) if chambre_widgets else 0
    total_nuits = calculer_total_nuits_gui()

    adresse_elements = [details.get('adresse', '').strip()]
    wilaya_cp = f"{details.get('wilaya', '').strip()} {details.get('code_postal', '').strip()}".strip()
    if wilaya_cp: adresse_elements.append(wilaya_cp)
    if details.get('pays', '').strip(): adresse_elements.append(details.get('pays', '').strip())
    adresse_complete = ", ".join(filter(None, adresse_elements))

    client_block = f"Détails du prestataire :\n"
    client_block += f"Entreprise : {selected_name}\n"
    # Les champs vides (NIF, NIS, Article, RC) ne sont pas affichés
    for libelle, cle in (("NIF", "nif"), ("NIS", "nis"), ("N° Article", "article"), ("RC", "rc")):
        valeur = str(details.get(cle, "") or "").strip()
        if valeur:
            client_block += f"{libelle} : {valeur}\n"
    if adresse_complete: client_block += f"Adresse : {adresse_complete}\n"
    if client_contact: client_block += f"Client : {client_contact}\n"
        
    client_block += f"Date d'arrivée : {'Dates différentes' if mode_multi_date.get() else global_arr}\n"
    client_block += f"Date de départ : {'Dates différentes' if mode_multi_date.get() else global_dep}\n"
    client_block += f"Nombre de nuitées : {total_nuits}\n"
    client_block += f"Nombre de chambres : {total_chambres_reel}"
    
    cell2.text = client_block

    doc.add_paragraph("")
    p_details = doc.add_paragraph("Détail des prestations")
    p_details.runs[0].bold = True

    # 5. Tableau de facturation (Lignes de détails)
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Light Grid Accent 2'
    headers = ['Date', 'Libellé', 'Qté', 'PUHT', 'Montant HT', 'TVA %']
    for i, h in enumerate(headers):
        table.cell(0, i).text = h

    total_ht = 0
    total_ht_non_remisable = 0  # Dîner + Extras : exclus de la remise
    total_taxe_sejour = 0

    for i, ch in enumerate(chambre_widgets):
        try:
            qte_similaire = int((ch['qte_entry'].get() or "1").split('.')[0])
            if qte_similaire <= 0: qte_similaire = 1
        except ValueError:
            qte_similaire = 1
            
        type_ch = ch['type_var'].get()
        lits_sup = int((ch['lits_entry'].get() or "0").split('.')[0])
        diner_ttc = float((ch['diner_entry'].get() or "0").replace(',', '.'))
        extras = ch.get('extras_list', [])

        if i == 0 or not mode_multi_date.get():
            nb_nuits = global_nuits
            date_affiche = global_arr
        else:
            try:
                d_local_arr = datetime.strptime(ch['date_arrivee'].get().strip(), DATE_FORMAT).date()
                d_local_dep = datetime.strptime(ch['date_depart'].get().strip(), DATE_FORMAT).date()
                nb_nuits = (d_local_dep - d_local_arr).days
                date_affiche = ch['date_arrivee'].get().strip()
            except:
                nb_nuits = 0
                date_affiche = "Date Invalide"
            
        if nb_nuits <= 0: continue

        nb_pers = 2 if type_ch in ("Ch. Double", "Ch. lit jumeaux", "Suite Junior.Double", "Suite Sup.Double") else 1
        info = CHAMBRES_DATA[type_ch]
        puht_chambre = info['puht']
        taxe_sejour = info['taxe_sejour'] * nb_nuits * qte_similaire
        total_taxe_sejour += taxe_sejour

        pd_total = PD_PUHT * nb_pers * nb_nuits * qte_similaire
        puht_diner = round(diner_ttc / (1 + DINER_TVA), 2) if diner_ttc > 0 else 0
        lit_total_ht = LIT_SUP_PUHT * lits_sup * nb_nuits * qte_similaire

        lignes = [
            (date_affiche, type_ch, nb_nuits * qte_similaire, puht_chambre, puht_chambre * nb_nuits * qte_similaire, "9%", True),
            (date_affiche, "Petit déj", nb_pers * nb_nuits * qte_similaire, PD_PUHT, pd_total, "9%", True),
        ]
        if lits_sup > 0:
            lignes.append((date_affiche, "Lit Supplémentaire", lits_sup * nb_nuits * qte_similaire, LIT_SUP_PUHT, lit_total_ht, "9%", True))
        if diner_ttc > 0:
            lignes.append((date_affiche, "Dîner", nb_nuits * qte_similaire, puht_diner, puht_diner * nb_nuits * qte_similaire, "9%", False))
        
        for extra in extras:
            try:
                extra_libelle = extra['libelle'].get().strip()
                extra_ttc_str = extra['montant'].get() or "0"
                extra_ttc = float(extra_ttc_str.replace(',', '.'))
                is_per_night = extra['is_per_night'].get()
            except Exception:
                continue
                
            if extra_libelle and (extra_ttc > 0 or extra_libelle.strip().lower() != "divers"):
                puht_extra = round(extra_ttc / (1 + EXTRA_TVA), 2) if extra_ttc > 0 else 0
                qte_facturation = qte_similaire * nb_nuits if is_per_night else qte_similaire
                extra_total_ht = puht_extra * qte_facturation
                tva_text = f"{EXTRA_TVA * 100:.0f}%" if extra_total_ht > 0 else "0%"

                lignes.append((date_affiche, extra_libelle, qte_facturation, puht_extra, extra_total_ht, tva_text, False))


        for date_val, lib, qte, puht, montant, tva_text, remisable in lignes:
            row = table.add_row().cells
            row[0].text = date_val
            row[1].text = lib
            row[2].text = str(qte)
            row[3].text = f"{puht:,.2f}".replace(",", " ").replace(".", ",")
            row[4].text = f"{montant:,.2f}".replace(",", " ").replace(".", ",")
            row[5].text = tva_text
            
            if montant > 0:
                total_ht += montant
                if not remisable:
                    total_ht_non_remisable += montant

 
    # 6. Totaux, Remise, Droit de timbre et lettres
    tva = round(total_ht * TVA_TAUX, 2)
    total_ttc = round(total_ht + tva + total_taxe_sejour, 2)

    try:
        remise_pct = max(0, float((entry_remise.get().strip() or "0").replace(',', '.')))
    except ValueError:
        remise_pct = 0

    # Montant HT + TVA (sans taxe de séjour)
    montant_sans_taxe = total_ht + tva

    # ⭐ La remise ne s'applique ni à la taxe de séjour, ni au Dîner, ni aux Extras.
    # Partie non remisable (Dîner + Extras) en TTC (HT + TVA)
    montant_non_remisable = round(total_ht_non_remisable * (1 + TVA_TAUX), 2)
    # Partie remisable = chambres, petit déj, lits supplémentaires (HT + TVA)
    montant_remisable = round(montant_sans_taxe - montant_non_remisable, 2)

    # Calcul de la remise uniquement sur la partie remisable
    remise_montant = round(montant_remisable * remise_pct / 100, 2)

    # Total TTC remisé = (total HT + TVA - remise) + taxe de séjour (non remisée)
    total_ttc_remise = round((montant_sans_taxe - remise_montant) + total_taxe_sejour, 2)


    
    # 6b. Calcul Droit de timbre
    droit_timbre_montant = 0
    if droit_timbre_var.get():
        # ⭐ MODIFICATION IMPORTANTE : On utilise total_ttc_remise pour le calcul du droit de timbre
        droit_timbre_montant = calculer_droit_timbre(total_ttc_remise)

    # 6c. Calcul Montant dû (anciennement Total à payer)
    total_a_payer = round(total_ttc_remise + droit_timbre_montant, 2) 

    doc.add_paragraph("")
    t_total = doc.add_table(rows=0, cols=2)
    t_total.style = 'Light Grid Accent 2'
    
    # Ajout des lignes au tableau des totaux
    labels_vals = [
        ("Total HT", total_ht), 
        ("TVA", tva), 
        ("Taxe de séjour", total_taxe_sejour), 
        ("Total TTC", total_ttc)
    ]
    
    if remise_montant > 0:
        labels_vals.append(("Remise", -remise_montant))
        
    if droit_timbre_montant > 0:
        labels_vals.append(("Droit de timbre", droit_timbre_montant))
        
    # ⭐ MODIFICATION 1 : Changement de "Montant à payer" vers "Montant dû"
    labels_vals.append(("Montant dû", total_a_payer))

    for label, val in labels_vals:
        r = t_total.add_row().cells
        # Met en gras "Montant dû"
        is_bold = label in ("Total TTC", "Montant dû") 
        r[0].paragraphs[0].add_run(label).bold = is_bold
        r[1].paragraphs[0].add_run(f"{val:,.2f}".replace(",", " ").replace(".", ",") + " DZD").bold = is_bold

    # 7. Conversion en lettres
    # On utilise total_a_payer (Montant dû)
    montant_a_payer_entier = int(total_a_payer)
    centimes = int(round((total_a_payer - montant_a_payer_entier) * 100))
    
    try:
        dinars_lettres = num2words(montant_a_payer_entier, lang='fr')
    except:
        dinars_lettres = "Zéro"
        
    # ⭐ MODIFICATION 2 : Libellé en lettres mis à jour
    total_a_payer_lettres = f"{dinars_lettres.capitalize()} dinars algériens"
    if centimes > 0:
        centimes_lettres = num2words(centimes, lang='fr')
        total_a_payer_lettres += f" et {centimes_lettres} centimes"

    table_lettres = doc.add_table(rows=1, cols=1)
    table_lettres.style = 'Light Grid Accent 2'
    cell = table_lettres.cell(0, 0)
    # ⭐ MODIFICATION 3 : Affichage du nouveau libellé
    run = cell.paragraphs[0].add_run(f"Le Montant dû en lettres est : {total_a_payer_lettres}")
    run.bold = True

    doc.add_paragraph("")
    doc.add_paragraph("")
    
    fname = f"facture_proforma_{numero_formatte}.docx"
    
    return doc, fname

# =========================================================================
# --- Conversion LibreOffice (DOCX -> PDF) ---
# =========================================================================

def convert_docx_to_pdf_libreoffice(docx_path: Path, pdf_path: Path):
    """
    Convertit DOCX en PDF en appelant LibreOffice.
    """
    if not Path(LIBREOFFICE_PATH).exists():
        raise FileNotFoundError(f"LibreOffice non trouvé. Vérifiez le chemin : {LIBREOFFICE_PATH}")

    # Le chemin de sortie est le dossier du PDF
    output_dir = str(pdf_path.parent)
    
    try:
        result = subprocess.run([
            LIBREOFFICE_PATH,
            '--headless',
            '--convert-to', 'pdf',
            str(docx_path),
            '--outdir', output_dir
        ],
        check=True,
        capture_output=True,
        text=True)

        expected_pdf_name = docx_path.stem + ".pdf"
        temp_pdf_path = Path(output_dir) / expected_pdf_name
        
        if temp_pdf_path.exists():
            # Renommer en cas de besoin si le nom de sortie est différent du nom attendu
            temp_pdf_path.rename(pdf_path)
        else:
            raise Exception("Le fichier PDF n'a pas été généré par LibreOffice. Sortie: " + result.stderr)
            
    except subprocess.CalledProcessError as e:
        raise Exception(f"La conversion LibreOffice a échoué. Erreur: {e.stderr}")
    except Exception as e:
        raise e

# =========================================================================
# --- Fonction d'Estampillage (Commun aux deux méthodes) ---
# =========================================================================

def executer_estampillage(fname_pdf_brut_path):
    """
    Sauvegarde le nom du PDF brut dans un JSON et exécute apply_stamp.py.
    """
    # 1. Sauvegarde du nom du PDF brut dans le JSON pour apply_stamp.py
    try:
        data = {"last_proforma_name": fname_pdf_brut_path.name}
        DOSSIER_SETTINGS.mkdir(parents=True, exist_ok=True)
        with open(FICHIER_NOM_PDF_TEMP_JSON, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showwarning("Avertissement", f"Impossible de sauvegarder le nom du PDF dans le JSON : {e}")
        os.startfile(fname_pdf_brut_path)
        message_label.config(text=f"PDF brut généré : {fname_pdf_brut_path.name}. Cachet non appliqué.", fg="red")
        return

    # 2. Exécution du script d'estampillage (apply_stamp.py)
    message_label.config(text="Création de proforma en cours...", fg="blue") # Message général
    root.update_idletasks()
    
    script_dir = Path(__file__).resolve().parent if '__file__' in locals() else Path(sys.argv[0]).resolve().parent
    stamp_script_path = script_dir / "apply_stamp.py"
    
    if not stamp_script_path.exists():
        messagebox.showerror("Erreur d'Exécution", f"Le script 'apply_stamp.py' est introuvable à l'emplacement: {stamp_script_path}")
        os.startfile(fname_pdf_brut_path)
        message_label.config(text=f"PDF brut généré : {fname_pdf_brut_path.name}. Exécutez le cachet manuellement.", fg="red")
        return

    try:
        # Estimation du nom du fichier final
        final_pdf_name = fname_pdf_brut_path.stem + "_hotel.pdf"
        
        # Utilise subprocess.Popen pour lancer le script en arrière-plan (non bloquant)
        subprocess.Popen([sys.executable, str(stamp_script_path)],
                             cwd=str(script_dir),
                             creationflags=subprocess.DETACHED_PROCESS if os.name == 'nt' else 0)
        
        # ⭐ MESSAGE FINAL MODIFIÉ
        message_label.config(text=f"Facture proforma PDF générée et ouverte : {final_pdf_name}", fg="green")
        
    except Exception as e:
        messagebox.showerror("Erreur d'Exécution", f"Impossible d'exécuter apply_stamp.py. Détail: {e}")
        os.startfile(fname_pdf_brut_path)
        message_label.config(text=f"PDF brut généré : {fname_pdf_brut_path.name}. Cachet manuel requis.", fg="red")


# =========================================================================
# --- Fonction de Génération PDF (AVEC CACHET AUTOMATIQUE) ---
# =========================================================================

def generer_pdf_avec_cachet():
    """Génère le .docx, le convertit en .pdf et exécute apply_stamp.py."""
    global last_generated_docx_path
    
    # 1. Génération et sauvegarde du DOCX
    try:
        doc, fname_docx = generer_document_docx(None, increment_counter=True)
        if doc is None: 
            check_pdf_button_state()
            return
            
        fname_docx_path = DOSSIER_PROFORMA / fname_docx
        fname_pdf_brut_path = fname_docx_path.with_suffix('.pdf')
        
        DOSSIER_PROFORMA.mkdir(parents=True, exist_ok=True)
        archiver_anciennes_proformas()  # Les anciennes proformas vont dans 'anciens'
        doc.save(fname_docx_path)
        last_generated_docx_path = fname_docx_path
        check_pdf_button_state(activate=True)

    except Exception as e:
        messagebox.showerror("Erreur de Génération DOCX", f"Une erreur s'est produite lors de la création du DOCX : {e}")
        check_pdf_button_state()
        return

    # 2. Conversion DOCX en PDF (via LibreOffice)
    message_label.config(text="Création de proforma en cours...", fg="blue") # ⭐ MESSAGE SIMPLIFIÉ
    root.update_idletasks()
    try:
        convert_docx_to_pdf_libreoffice(fname_docx_path, fname_pdf_brut_path)
    except Exception as e:
        messagebox.showerror("ERREUR CRITIQUE DE CONVERSION",
                             f"Impossible de convertir DOCX en PDF. Vérifiez le chemin LibreOffice ({LIBREOFFICE_PATH}). Détail: {e}")
        try: os.remove(fname_docx_path)
        except: pass
        return
    
    # Nettoyage du fichier .docx 
    try:
        os.remove(fname_docx_path)
        last_generated_docx_path = None
        check_pdf_button_state()
    except Exception:
        pass
        
    # 3. Exécuter Estampillage
    executer_estampillage(fname_pdf_brut_path)

# =========================================================================
# --- Fonction : Génération PDF à partir du DOCX Modifié ---
# =========================================================================

def generer_pdf_modifier():
    """Utilise le dernier DOCX (potentiellement modifié) pour créer le PDF estampillé."""
    global last_generated_docx_path
    
    if last_generated_docx_path is None or not last_generated_docx_path.exists():
        messagebox.showwarning("Attention", "Ce bouton est disponible uniquement après avoir cliqué sur 'Visualiser la proforma (Word)'.")
        return

    fname_pdf_brut_path = last_generated_docx_path.with_suffix('.pdf')
    
    # ⭐ MESSAGE DE TRANSITION
    message_label.config(text="Chargement...", fg="blue") 
    root.update_idletasks()

    # 1. Conversion DOCX en PDF (via LibreOffice)
    message_label.config(text="Création de proforma en cours...", fg="blue") # ⭐ MESSAGE SIMPLIFIÉ
    root.update_idletasks()
    try:
        convert_docx_to_pdf_libreoffice(last_generated_docx_path, fname_pdf_brut_path)
    except Exception as e:
        messagebox.showerror("ERREUR CRITIQUE DE CONVERSION",
                             f"Impossible de convertir DOCX en PDF. Le fichier DOCX est-il bien fermé ? Détail: {e}")
        return
        
    # 2. Suppression du DOCX
    try:
        os.remove(last_generated_docx_path)
        last_generated_docx_path = None
        check_pdf_button_state(activate=False)
    except Exception as e:
        messagebox.showwarning("Avertissement", f"Impossible de supprimer le DOCX modifié : {e}")
        
    # 3. Exécuter Estampillage
    executer_estampillage(fname_pdf_brut_path)


def generer_facture():
    """Génère uniquement le .docx et l'ouvre (Visualiser la proforma (Word))."""
    global last_generated_docx_path
    
    # 1. Génération et sauvegarde du DOCX
    try:
        doc, fname = generer_document_docx(None, increment_counter=True)
        if doc is None: 
            check_pdf_button_state()
            return

        out = DOSSIER_PROFORMA
        out.mkdir(parents=True, exist_ok=True)
        fname_final = out / fname
        
        archiver_anciennes_proformas()  # Les anciennes proformas vont dans 'anciens'
        doc.save(fname_final)
        
        # ⭐ Mémoriser le chemin du DOCX et activer le bouton PDF Modifié
        last_generated_docx_path = fname_final
        check_pdf_button_state(activate=True)
        
        os.startfile(fname_final)
        message_label.config(text=f"Facture DOCX générée et ouverte : {fname_final.name}", fg="green")
        
    except Exception as e:
        messagebox.showerror("Erreur de Génération DOCX", f"Erreur : {e}")
        check_pdf_button_state()
        
# =========================================================================
# --- Interface graphique (Disposition) (Partie 2) ---
# =========================================================================
def force_focus_depart(event):
    entry_depart.focus_set()
    return "break"

def force_focus_entreprise(event):
    entry_entreprise.focus_set()
    return "break"

def check_pdf_button_state(activate=False):
    """Gère l'état du bouton PDF Modifié."""
    global btn_modifier_pdf
    
    if btn_modifier_pdf:
        if activate:
            btn_modifier_pdf.config(state="normal", bg="#dc3545", fg="white", text="Créer la proforma (PDF MODIFIÉ)")
        else:
            btn_modifier_pdf.config(state="disabled", bg="gray", fg="black", text="Créer la proforma (PDF Modifié)")

show_splash(root)

# ⭐ DÉBUT DE LA MODIFICATION DE TAILLE
# Constantes pour l'augmentation de taille (approx. 34 cm = 1342 pixels)
LARGEUR_SUPPLEMENTAIRE = 1342 # AJOUT DE 1 CM (39 pixels) suppl. (1303 + 39 = 1342)
HAUTEUR_SUPPLEMENTAIRE = 233 # Ajout de 6 cm de hauteur (approx. 233 pixels)

root.update_idletasks() # Assurez-vous que les widgets sont dessinés
window_width_req = root.winfo_reqwidth()
window_height_req = root.winfo_reqheight()

# Nouvelle taille augmentée
new_window_width = window_width_req + LARGEUR_SUPPLEMENTAIRE
new_window_height = window_height_req + HAUTEUR_SUPPLEMENTAIRE

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# --- CALCUL DE CENTRAGE ---
x = (screen_width // 2) - (new_window_width // 2)
y = (screen_height // 2) - (new_window_height // 2)

root.geometry(f"{new_window_width}x{new_window_height}+{x}+{y}")
# ⭐ FIN DE LA MODIFICATION DE TAILLE

frm_top = tk.Frame(root)
frm_top.pack(fill='x', padx=10, pady=8)

# Ligne 1: Entreprise et Date Arrivée
tk.Label(frm_top, text="Entreprise").grid(row=0, column=0, sticky='e')
entry_entreprise = tk.Entry(frm_top, textvariable=entreprise_var, width=50)
entry_entreprise.grid(row=0, column=1, padx=5, sticky='w')
btn_gerer_clients = tk.Button(frm_top, text="Liste d'entreprises", command=gerer_clients)
btn_gerer_clients.grid(row=0, column=1, sticky='e', padx=(0, 2))
btn_gerer_clients.config(takefocus=0)

tk.Label(frm_top, text="Arrivée (JJ/MM/AAAA)").grid(row=0, column=2, sticky='e')
entry_arrivee = tk.Entry(frm_top, width=15)
entry_arrivee.grid(row=0, column=3, padx=5)
entry_arrivee.bind("<KeyRelease>", lambda event: update_all_dates(event))
entry_arrivee.bind("<FocusOut>", lambda event: update_all_dates(event))
entry_arrivee.bind("<Key-Tab>", force_focus_depart)
entry_arrivee.bind("<Shift-Key-Tab>", force_focus_entreprise)


# Ligne 2: Client, Date Départ et Remise
tk.Label(frm_top, text="Client").grid(row=1, column=0, sticky='e')
entry_client = tk.Entry(frm_top, width=40)
entry_client.grid(row=1, column=1, padx=5, sticky='w')

tk.Label(frm_top, text="Départ (JJ/MM/AAAA)").grid(row=1, column=2, sticky='e')
entry_depart = tk.Entry(frm_top, width=15)
entry_depart.grid(row=1, column=3, padx=5)
entry_depart.bind("<KeyRelease>", lambda event: update_all_dates(event))
entry_depart.bind("<FocusOut>", lambda event: update_all_dates(event))
entry_depart.bind("<Shift-Key-Tab>", lambda event: entry_arrivee.focus_set())
entry_depart.bind("<Key-Tab>", lambda event: entry_remise.focus_set())


tk.Label(frm_top, text="Remise (%)").grid(row=1, column=4, sticky='e')
entry_remise = tk.Entry(frm_top, width=5, validate='key', vcmd=(vcmd, '%P'))
entry_remise.insert(0, "0")
entry_remise.grid(row=1, column=5, padx=5)
entry_remise.bind("<Shift-Key-Tab>", lambda event: entry_depart.focus_set())

# ⭐ NOUVEAU WIDGET : Droit de Timbre
droit_timbre_var = tk.BooleanVar(value=False)
tk.Checkbutton(frm_top, text="Droit de timbre", variable=droit_timbre_var).grid(row=1, column=6, sticky='w', padx=(10, 0))


# Ligne 3: Option Multi-dates
mode_multi_date = tk.BooleanVar(value=False)
tk.Checkbutton(frm_top, text="Dates différentes par chambre", variable=mode_multi_date, command=toggle_multi_dates).grid(row=2, column=0, columnspan=2, pady=4, sticky='w')

# Conteneur des Chambres avec Scrollbar
canvas = tk.Canvas(root, borderwidth=0, height=340)
scroll_y = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
wrapper = tk.Frame(canvas)
wrapper.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=wrapper, anchor="nw")
canvas.configure(yscrollcommand=scroll_y.set)
canvas.pack(side="left", fill="both", expand=True, padx=(10,0), pady=5)
scroll_y.pack(side="right", fill="y", pady=5)
chambres_container = wrapper

# Compteurs
compteur_nuits_label = tk.Label(root, text="Nombre de nuitées : 0", font=("Arial", 10, "bold"))
compteur_nuits_label.pack(pady=(4,0))
compteur_chambres_label = tk.Label(root, text="Nombre de chambres : 0", font=("Arial", 10, "bold"))
compteur_chambres_label.pack(pady=(4,0))

# Bouton Ajouter Chambre
tk.Button(root, text="+ Ajouter une chambre", command=ajouter_chambre).pack(pady=(10, 4))

# Pied de page (Boutons Créer Facture et Ouvrir Dossier)
bottom_frame = tk.Frame(root)
# Aligner bottom_frame à droite
bottom_frame.pack(side='bottom', anchor='e', padx=10, pady=(10, 12))

# 1. Créer la proforma (PDF) - En haut
btn_creer_pdf = tk.Button(bottom_frame, text="Créer la proforma (PDF)", command=generer_pdf_avec_cachet, bg="#28a745", fg="white", font=("Arial", 12, "bold"), height=2, width=30)
btn_creer_pdf.pack(side="top", pady=5, anchor='e')

# 2. Créer la proforma (PDF Modifié)
btn_modifier_pdf = tk.Button(bottom_frame, text="Créer la proforma (PDF Modifié)", command=generer_pdf_modifier, font=("Arial", 12, "bold"), height=2, width=30)
btn_modifier_pdf.bind('<Button-1>', lambda event: messagebox.showwarning("Attention", "Ce bouton est disponible uniquement après avoir cliqué sur 'Visualiser la proforma (Word)'.") if btn_modifier_pdf.cget('state') == 'disabled' else None)
check_pdf_button_state(activate=False)
btn_modifier_pdf.pack(side="top", pady=5, anchor='e')

# 3. Créer la proforma (DOCX) -> Visualiser la proforma (Word)
btn_creer_docx = tk.Button(bottom_frame, text="Visualiser la proforma (Word)", command=generer_facture, bg="#007bff", fg="white", font=("Arial", 12, "bold"), height=2, width=25)
btn_creer_docx.pack(side="top", pady=5, anchor='e')

# 4. Ouvrir le dossier des proformas - En bas
btn_ouvrir_dossier = tk.Button(bottom_frame, text="📂 Ouvrir le dossier des proformas", command=ouvrir_dossier, bg="#ccc", font=("Arial", 10))
btn_ouvrir_dossier.pack(side="top", pady=5, anchor='e')


status_frame = tk.Frame(root)
status_frame.pack(side='bottom', anchor='w', padx=10, pady=(0, 10))
message_label = tk.Label(status_frame, text="", fg="green", font=("Arial", 10))
message_label.pack(side='left')

# --- Menu Paramètres ---
menubar = tk.Menu(root)
menu_parametres = tk.Menu(menubar, tearoff=0)
menu_parametres.add_command(label="Informations de l'entreprise...", command=modifier_infos_entreprise)
menubar.add_cascade(label="Paramètres", menu=menu_parametres)
root.config(menu=menubar)

# --- Démarrage ---
# ⭐ Récupère les anciens fichiers JSON (réseau / dossier du script) vers settings/
migrer_anciens_fichiers()

# ⭐ Premier lancement : demander les informations de l'entreprise
entreprise_infos = charger_entreprise()
if not entreprise_infos.get("nom"):
    root.withdraw()
    _data = demander_infos_entreprise(root, entreprise_infos, premiere_fois=True)
    if not _data or not sauvegarder_entreprise(_data):
        root.destroy()
        sys.exit()
    entreprise_infos = _data
    root.deiconify()

charger_clients()
entreprise_var.set("")
ajouter_chambre()
# Raccourci Entrée
root.bind('<Return>', lambda event: generer_pdf_avec_cachet() if btn_creer_pdf.cget('state') == 'normal' else None) 
root.mainloop()
