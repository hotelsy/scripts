import tkinter as tk

def calculate_and_minimize():
    try:
        # Obtenir la valeur d'entrée
        raw_input = entry.get()
        
        # Remplacer les espaces par des vides et gérer les caractères décimaux
        cleaned_input = raw_input.replace(" ", "").replace(",", ".")
        
        # Convertir à float
        number = float(cleaned_input)
        
        # Calculer le résultat
        result = number / 1.09
        rounded_result = round(result, 2)
        formatted_result = f"{rounded_result:.2f}".replace('.', ',')

        # Copier le résultat dans le presse-papiers
        root.clipboard_clear()  # Effacer le presse-papiers
        root.clipboard_append(formatted_result)  # Ajouter le résultat au presse-papiers

        # Afficher le résultat
        result_label.config(text=f"Résultat: {formatted_result} (copié)")

        # Effacer le champ d'entrée
        entry.delete(0, tk.END)

        # Minimiser la fenêtre
        root.iconify()

    except ValueError:
        entry.delete(0, tk.END)  # Effacer le champ d'entrée pour nouvelle saisie
        entry.focus_set()  # Récupérer le focus sur le champ d'entrée en cas d'entrée invalide
        entry.insert(0, "Entrée invalide!")  # Retour d'information à l'utilisateur

# Créer la fenêtre principale
root = tk.Tk()
root.title("Calculateur TVA 9%")
root.geometry("400x200")  # Définir la taille de la fenêtre à 400x200 pixels

# Centrer la fenêtre sur l'écran
window_width = 400
window_height = 200

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Ajuster la variable x pour déplacer la fenêtre vers la droite de 10 cm
x = (screen_width // 2) - (window_width // 2) + 378  # Déplacer vers la droite d'environ 378 pixels
y = (screen_height // 2) - (window_height // 2)

root.geometry(f"{window_width}x{window_height}+{x}+{y}")  # Définir la taille et la position

# Créer et placer des widgets
label = tk.Label(root, text="Entrer un chiffre:")
label.pack(pady=10)

entry = tk.Entry(root)
entry.pack(pady=5)
entry.focus_set()  # Mettre le focus sur le champ d'entrée

button = tk.Button(root, text="Diviser", command=calculate_and_minimize)
button.pack(pady=10)

result_label = tk.Label(root, text="")
result_label.pack()

# Lier la touche Entrée à la fonction
root.bind('<Return>', lambda event: calculate_and_minimize())

# Exécuter l'application
root.mainloop()
