import json
import os
import tkinter as tk
from tkinter import messagebox
import subprocess
import platform
import shutil
# Les imports re et Pillow ne sont plus strictement nécessaires pour la logique de prix/remplacement

# ⬇️ Importations pour gérer le PNG avec transparence (gardées pour votre interface)
try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

# 📂 Dossier du script courant
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ⭐ Dossier "settings" : tous les fichiers générés par ce script y sont enregistrés
SETTINGS_DIR = os.path.join(BASE_DIR, "settings")
os.makedirs(SETTINGS_DIR, exist_ok=True)
PRIX_JSON_PATH = os.path.join(SETTINGS_DIR, "prix_chambres.json")
PRIX_TXT_PATH = os.path.join(SETTINGS_DIR, "prix_calcules.txt")

# Migration (une seule fois) : copie les anciens fichiers placés à côté du script vers settings/
for _nom in ("prix_chambres.json", "prix_calcules.txt"):
    _ancien = os.path.join(BASE_DIR, _nom)
    _nouveau = os.path.join(SETTINGS_DIR, _nom)
    try:
        if os.path.isfile(_ancien) and not os.path.exists(_nouveau):
            shutil.copy2(_ancien, _nouveau)
    except Exception:
        pass

# Constantes pour les calculs
TVA = 0.09
PD_HT = 1500  # Petit déjeuner HT
TAXE_SEJOUR = 200  # Taxe de séjour par personne et par nuit

# Configuration des chambres
CHAMBRES_CONFIG = [
    {"label": "Chambre Single", "json_key": "chambre_single", "nb_pers": 1},
    {"label": "Chambre Double", "json_key": "chambre_double", "nb_pers": 2},
    {"label": "Chambre Twin", "json_key": "chambre_twin", "nb_pers": 2},
    {"label": "Suite Junior Single", "json_key": "suite_junior_single", "nb_pers": 1},
    {"label": "Suite Junior Double", "json_key": "suite_junior_double", "nb_pers": 2},
    {"label": "Suite Sup Single", "json_key": "suite_sup_single", "nb_pers": 1},
    {"label": "Suite Sup Double", "json_key": "suite_sup_double", "nb_pers": 2},
]
# NOTE: "pyw_marker" a été retiré car nous n'écrivons plus dans p2.pyw.

ENTRY_FIELDS = {}

# --- Fonctions de Fichier ---

def load_from_json():
    """Charge les données depuis le fichier JSON."""
    json_path = PRIX_JSON_PATH
    if not os.path.exists(json_path):
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump({}, f)
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        return {}

def save_to_json(prices):
    """Sauvegarde les prix dans le fichier JSON."""
    json_path = PRIX_JSON_PATH
    os.makedirs(SETTINGS_DIR, exist_ok=True)
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            # Sauvegarde les prix sous la forme : {"chambre_single": {"ttc": "12000", "puht": "11009.17"}}
            json.dump(prices, f, indent=4, ensure_ascii=False)
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde : {str(e)}")

# --- Fonction de Calcul ---

def calculer_pu(prix_ttc, nuits, nb_chambres, nb_pers):
    """Calcule le PUHT par nuit/chambre."""
    taxe_sejour_total = TAXE_SEJOUR * nuits * nb_pers * nb_chambres
    ttc_sans_taxe_sejour = prix_ttc - taxe_sejour_total
    total_ht = ttc_sans_taxe_sejour / (1 + TVA)
    pd_total = PD_HT * nuits * nb_pers * nb_chambres
    puht_total = total_ht - pd_total
    puht_par_nuit_par_chambre = puht_total / (nuits * nb_chambres)
    return f"{puht_par_nuit_par_chambre:.2f}"

# --- Fonctions principales ---

def creer_fichier_prix():
    """Lit les entrées, calcule, sauvegarde les prix et lance p2.pyw."""
    resultats_complets = load_from_json()
    
    try:
        nuits = 1
        nb_chambres = 1

        for ch in CHAMBRES_CONFIG:
            entry_field = ENTRY_FIELDS[ch["json_key"]]
            entry_value = entry_field.get().strip().replace(" ", "").replace(",", ".")
            json_key = ch["json_key"]

            if entry_value:
                try:
                    prix_ttc = float(entry_value)
                except ValueError:
                    raise ValueError(f"Le prix TTC pour {ch['label']} doit être un nombre valide.")

                pu_calcule = calculer_pu(prix_ttc, nuits, nb_chambres, ch["nb_pers"])
                
                # Mise à jour du résultat JSON
                resultats_complets[json_key] = {
                    "ttc": entry_value, # Gardé en string pour l'affichage de l'interface
                    "puht": pu_calcule # Gardé en string pour la persistance
                }
            # Si le champ est vide, on utilise l'ancienne valeur si elle existe
            elif json_key in resultats_complets:
                 pass # On garde les données existantes dans resultats_complets


        save_to_json(resultats_complets)

        # Création/écriture du fichier texte (conservé)
        txt_path = PRIX_TXT_PATH
        with open(txt_path, "w", encoding="utf-8") as f:
            for key, data in resultats_complets.items():
                f.write(f"{key} - TTC Saisi: {data['ttc']}, PUHT Calculé : {data['puht']}\n")

        # ⭐ NOUVEAU : Lance p2.pyw sans le modifier.
        lancer_p2pyw()

    except ValueError as ve:
        messagebox.showerror("Erreur de Saisie", str(ve))
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur inattendue est survenue : {str(e)}")

def creer_fichier_prix_event(event=None):
    creer_fichier_prix()

def lancer_p2pyw():
    """Lance le fichier p2.pyw après sauvegarde."""
    fichier_p2pyw = os.path.join(BASE_DIR, "p2.pyw")
    
    if not os.path.exists(fichier_p2pyw):
         messagebox.showerror("Erreur", f"Le script p2.pyw est introuvable à l'emplacement: {fichier_p2pyw}")
         return
         
    try:
        messagebox.showinfo("Succès", "Les prix ont été enregistrés.")
        
        # Exécution du fichier p2.pyw
        if platform.system() == "Windows":
            # Utilisez 'start' pour lancer en arrière-plan et ne pas bloquer
            subprocess.Popen(["cmd", "/c", "start", "pythonw", fichier_p2pyw], close_fds=True)
        else:
            subprocess.Popen(["python3", fichier_p2pyw], close_fds=True)
            
        # Ferme la fenêtre de configuration
        root.destroy()
        
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du lancement de p2.pyw : {str(e)}")


# 🎨 Fonction d'Écran de Démarrage (Splash Screen)
# ... [Le code de show_splash_screen() et le bloc principal Tkinter sont conservés] ...

def show_splash_screen():
    if not Image or not ImageTk:
        root.deiconify()
        return

    splash = tk.Toplevel(root)
    
    TRANSPARENT_COLOR = '#f000ff' 
    
    splash.title("Démarrage...")
    splash.overrideredirect(True)
    
    img_path = os.path.join(BASE_DIR, "start.png")
    
    try:
        if not os.path.exists(img_path):
            tk.Label(splash, text="Image start.png non trouvée. (Continuer...)", bg=TRANSPARENT_COLOR).pack(padx=10, pady=10)
            w_img, h_img = 300, 50 
        else:
            image = Image.open(img_path)
            w_img, h_img = image.size
            photo = ImageTk.PhotoImage(image)
            
            splash.photo = photo
            
            tk.Label(splash, image=photo, borderwidth=0, bg=TRANSPARENT_COLOR).pack(expand=True, fill='both')
    
    except Exception as e:
        tk.Label(splash, text=f"Erreur de chargement de l'image 'start.png': {e}", bg=TRANSPARENT_COLOR).pack(padx=10, pady=10)
        w_img, h_img = 350, 50 

    splash.config(bg=TRANSPARENT_COLOR)
    splash.attributes('-transparentcolor', TRANSPARENT_COLOR)

    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    x_img, y_img = (sw - w_img) // 2, (sh - h_img) // 2
    splash.geometry(f"{w_img}x{h_img}+{x_img}+{y_img}")
    
    splash.after(3000, lambda: [splash.destroy(), root.deiconify()])

# --- Démarrage de l'Interface Tkinter ---

root = tk.Tk()
root.withdraw() 
root.title("Gestion des Prix des Chambres")
root.resizable(False, False)

show_splash_screen()

w, h = 600, 600
sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
x, y = (sw - w) // 2, (sh - h) // 2
root.geometry(f"{w}x{h}+{x}+{y}")

prix_precedents = load_from_json()

for config in CHAMBRES_CONFIG:
    tk.Label(root, text=f"{config['label']} (TTC) :").pack(pady=5)
    entry = tk.Entry(root, width=20, font=("Arial", 14))
    entry.pack(pady=5)

    json_key = config['json_key']
    if json_key in prix_precedents and 'ttc' in prix_precedents[json_key]:
        entry.insert(0, prix_precedents[json_key]['ttc'])

    ENTRY_FIELDS[json_key] = entry

btn_calculer = tk.Button(root, text="Valider les tarifs", command=creer_fichier_prix,
                          font=("Arial", 14), bg="#4CAF50", fg="white")
btn_calculer.pack(pady=20)

root.bind('<Return>', creer_fichier_prix_event)

root.mainloop()
