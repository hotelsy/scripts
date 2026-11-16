import tkinter as tk
from tkinter import ttk, messagebox
import json
import os

# Constantes
TVA = 0.09
PD_HT = 1500
TAXE_SEJOUR = 200
JSON_FILE = "choix_utilisateur.json"

# Fonction pour sauvegarder dans un fichier JSON
def sauvegarder_choix(petit_dej_inclus):
    data = {"petit_dejeuner_inclus": petit_dej_inclus}
    with open(JSON_FILE, "w") as f:
        json.dump(data, f, indent=4)

# Fonction de calcul
def faire_calcul():
    try:
        ttc_raw = entry_ttc.get().strip().replace(" ", "").replace(",", ".")
        nuits_raw = entry_nuits.get().strip()
        chambres_raw = entry_chambres.get().strip()

        if not ttc_raw or not nuits_raw or not chambres_raw:
            raise ValueError("Tous les champs sont obligatoires.")

        ttc_total = float(ttc_raw)
        nuits = int(nuits_raw)
        nb_chambres = int(chambres_raw)

        if nuits <= 0 or ttc_total <= 0 or nb_chambres <= 0:
            raise ValueError("Les valeurs doivent être positives.")

        type_chambre = chambre_type.get()
        nb_pers = 1 if type_chambre == "Single" else 2

        # Valeur booléenne pour savoir si on mixe la taxe
        mixer_taxe = bool(var_mixer_taxe.get())

        # Calcul de la taxe de séjour totale
        taxe_sejour_total = TAXE_SEJOUR * nuits * nb_pers * nb_chambres

        # Si la taxe est mixée dans le TTC, on ne la retire pas
        if mixer_taxe:
            ttc_sans_taxe_sejour = ttc_total
        else:
            ttc_sans_taxe_sejour = ttc_total - taxe_sejour_total

        # On retire la TVA pour obtenir le HT
        total_ht = ttc_sans_taxe_sejour / (1 + TVA)

        # Sauvegarde du choix petit déjeuner
        petit_dej_inclus = bool(var_petit_dej.get())
        sauvegarder_choix(petit_dej_inclus)

        # Si le petit déj est inclus, on ne le retire pas
        if petit_dej_inclus:
            puht_total = total_ht
        else:
            pd_total = PD_HT * nuits * nb_pers * nb_chambres
            puht_total = total_ht - pd_total

        puht_par_nuit_par_chambre = puht_total / (nuits * nb_chambres)

        return f"{puht_par_nuit_par_chambre:.2f}".replace(".", ",")

    except ValueError as ve:
        messagebox.showerror("Erreur", str(ve))
        return None
    except Exception:
        messagebox.showerror("Erreur", "Entrée invalide. Vérifie les valeurs.")
        return None

# Fonction globale
def calculer_et_copier(event=None):
    resultat = faire_calcul()
    if not resultat:
        return

    root.clipboard_clear()
    root.clipboard_append(resultat)
    root.update()

    # Réinitialisation
    entry_ttc.delete(0, tk.END)
    entry_nuits.delete(0, tk.END)
    entry_chambres.delete(0, tk.END)
    chambre_type.current(0)
    var_petit_dej.set(0)
    var_mixer_taxe.set(0)
    entry_ttc.focus()

    root.iconify()

# --- Interface ---
root = tk.Tk()
root.title("Calculateur de PU")
root.resizable(False, False)

w, h = 350, 520
sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
x, y = (sw - w) // 2, (sh - h) // 2
root.geometry(f"{w}x{h}+{x}+{y}")

# Widgets
tk.Label(root, text="Total TTC :").pack(pady=(10, 0))
entry_ttc = tk.Entry(root, width=20, font=("Arial", 14))
entry_ttc.pack()

tk.Label(root, text="Nuitées :").pack(pady=(10, 0))
entry_nuits = tk.Entry(root, width=20, font=("Arial", 14))
entry_nuits.pack()

tk.Label(root, text="Nombre de chambres :").pack(pady=(10, 0))
entry_chambres = tk.Entry(root, width=20, font=("Arial", 14))
entry_chambres.pack()

chambre_type = ttk.Combobox(root, values=["Single", "Double"], state="readonly", width=17, font=("Arial", 12))
chambre_type.current(0)
chambre_type.pack(pady=10)

# Checkbox pour petit déjeuner
var_petit_dej = tk.IntVar(value=0)
check_petit_dej = tk.Checkbutton(
    root,
    text="Ajouter le débit de petit déjeuner au débit de la chambre",
    variable=var_petit_dej,
    font=("Arial", 10)
)
check_petit_dej.pack(pady=(10, 0))

# Checkbox pour taxe de séjour
var_mixer_taxe = tk.IntVar(value=0)
check_taxe = tk.Checkbutton(
    root,
    text="Ajouter la taxe de séjour au débit de la chambre",
    variable=var_mixer_taxe,
    font=("Arial", 10)
)
check_taxe.pack(pady=(5, 0))

# Bouton
btn_calculer = tk.Button(
    root,
    text="Calculer",
    command=calculer_et_copier,
    bg="#2b7",
    fg="white",
    font=("Arial", 20, "bold"),
    width=12,
    height=3
)
btn_calculer.pack(pady=20)

# Bind touche Entrée
root.bind('<Return>', calculer_et_copier)

entry_ttc.focus()
root.mainloop()
