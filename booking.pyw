import tkinter as tk
from PIL import Image, ImageTk  # Nécessaire pour afficher le logo

def calculate_per_night_tva():
    try:
        # Obtenir les valeurs des champs d'entrée
        raw_booking = entry_booking.get()
        raw_nights = entry_nights.get()
        raw_housing = entry_housing.get()  # Récupérer la valeur du nombre de logements

        # Nettoyage des valeurs
        booking_cleaned = raw_booking.replace(" ", "").replace(",", ".")
        nights_cleaned = raw_nights.replace(" ", "").replace(",", ".")
        housing_cleaned = raw_housing.replace(" ", "").replace(",", ".")

        # Conversion en float
        total_booking = float(booking_cleaned)
        nights = float(nights_cleaned)
        housing = float(housing_cleaned)

        if nights <= 0:
            raise ValueError("Nombre de nuitées invalide")
        if housing <= 0:
            raise ValueError("Nombre de logements invalide")

        # Calcul : montant / nuitée, puis / 1.09, et divisé par le nombre de logements
        per_night = total_booking / nights
        result = (per_night / 1.09) / housing

        # Formatage
        rounded_result = round(result, 2)
        formatted_result = f"{rounded_result:.2f}".replace('.', ',')

        # Copier dans le presse-papiers
        root.clipboard_clear()
        root.clipboard_append(formatted_result)

        # Afficher le résultat
        result_label.config(text=f"Résultat: {formatted_result} DZD (copié)")

        # Effacer les champs booking et nights uniquement, mais garder le nombre de logements
        entry_booking.delete(0, tk.END)
        entry_nights.delete(0, tk.END)

        # Réduire la fenêtre
        root.iconify()

    except ValueError:
        result_label.config(text="Entrée invalide. Réessaye.")
        entry_booking.delete(0, tk.END)
        entry_nights.delete(0, tk.END)
        entry_booking.focus_set()

def on_window_deiconify(event):
    # Placer le focus sur le champ "Montant total du booking" après la restauration de la fenêtre
    entry_booking.focus_set()

# Interface graphique
root = tk.Tk()
root.title("Calculateur de booking")
root.geometry("400x360")  # Hauteur ajustée pour le logo

# Centrer la fenêtre
window_width = 400
window_height = 360
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (window_width // 2) + 378
y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# ===== AJOUT DU LOGO =====
try:
    logo_image = Image.open("logo.png")
    logo_image = logo_image.resize((180, 60), Image.ANTIALIAS)
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = tk.Label(root, image=logo_photo, bg="white")
    logo_label.image = logo_photo  # Garde une référence
    logo_label.pack(pady=(10, 0))
except Exception as e:
    print(f"Erreur lors du chargement du logo : {e}")
# =========================

# Widgets
label1 = tk.Label(root, text="Montant total du booking (DZD) :")
label1.pack(pady=(15, 5))
entry_booking = tk.Entry(root)
entry_booking.pack(pady=5)

label2 = tk.Label(root, text="Nombre de nuitées :")
label2.pack(pady=(10, 5))
entry_nights = tk.Entry(root)
entry_nights.pack(pady=5)

label3 = tk.Label(root, text="Nombre de logements :")
label3.pack(pady=(10, 5))
entry_housing = tk.Entry(root)
entry_housing.insert(0, "1")  # Valeur par défaut est 1
entry_housing.pack(pady=5)

button = tk.Button(root, text="Calculer", command=calculate_per_night_tva)
button.pack(pady=15)

result_label = tk.Label(root, text="")
result_label.pack()

# Entrée clavier
root.bind('<Return>', lambda event: calculate_per_night_tva())

# Lors de la restauration de la fenêtre (dé-reduction)
root.bind("<Map>", on_window_deiconify)

# Initialiser le focus sur le champ "Montant total du booking"
entry_booking.focus_set()

root.mainloop()
