import tkinter as tk

def format_number(number):
    """Formate le nombre : supprime .0 inutiles, remplace . par ,."""
    if number.is_integer():
        return f"{int(number)}"
    else:
        return f"{number:.2f}".replace('.', ',')

# Fonction pour calculer le total selon les tranches
def calculer_total(montant):
    if 300 <= montant <= 29999:
        total = montant * 1.01       # +1%
    elif 30000 <= montant <= 99999:
        total = montant * 1.015      # +1,5%
    elif montant >= 100000:
        total = montant * 1.02       # +2%
    else:
        total = montant
    return round(total, 2)

# Fonction pour calculer le droit de timbre (juste le pourcentage)
def calcul_droit(montant):
    if montant < 300:
        return 0.0
    elif 300 <= montant <= 29999:
        return round(montant * 0.01, 2)    # 1%
    elif 30000 <= montant <= 99999:
        return round(montant * 0.015, 2)   # 1,5%
    else:
        return round(montant * 0.02, 2)    # 2%

# Quand l'utilisateur valide la saisie
def traiter_saisie():
    global total_calculé, droit_calculé

    # Lecture + remplacement des virgules
    montant_str = entry.get().replace(" ", "").replace(",", ".")
    try:
        montant = float(montant_str)
    except ValueError:
        return

    # Calculs
    total_calculé = calculer_total(montant)
    droit_calculé = calcul_droit(montant)

    # Formatage pour affichage
    total_str = format_number(total_calculé)
    droit_str = format_number(droit_calculé)

    # Mise à jour des boutons
    btn_total.config(text=f"Copier le Total : {total_str} DA")
    btn_droit.config(text=f"Copier Droit de Timbre : {droit_str} DA")

    # Affichage
    btn_total.pack(pady=5)
    btn_droit.pack(pady=5)

def copier_et_reduire(text):
    """Copie dans le presse-papiers et minimise la fenêtre."""
    root.clipboard_clear()
    root.clipboard_append(text)
    root.iconify()

# --- INTERFACE TKINTER ---

root = tk.Tk()
root.title("Calculateur de droit de timbre")

# Positionnement légèrement à droite du centre
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) + 378
y = (screen_height // 2) - 100
root.geometry(f"400x200+{x}+{y}")

total_calculé = 0
droit_calculé = 0

label = tk.Label(root, text="Entrez le montant TTC :")
label.pack(pady=10)

entry = tk.Entry(root)
entry.pack()
entry.focus_set()

# Valider avec Entrée
entry.bind('<Return>', lambda event: traiter_saisie())

# Boutons (vides au début)
btn_total = tk.Button(root, text="", command=lambda: copier_et_reduire(format_number(total_calculé)))
btn_droit = tk.Button(root, text="", command=lambda: copier_et_reduire(format_number(droit_calculé)))

root.mainloop()
