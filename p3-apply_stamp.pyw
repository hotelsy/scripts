import fitz # PyMuPDF
import json
from pathlib import Path
import os
import sys

# =================================================================
# --- CHEMIN DE CONFIGURATION (Doit correspondre au script principal p2.pyw) ---
# =================================================================

# Le chemin de base doit être le dossier partagé sur le serveur
BASE_CHEMIN_RESEAU = r"\\192.168.100.20\dossier partage de serveur\proforma" 
DOSSIER_PROFORMA = Path(BASE_CHEMIN_RESEAU) 

# Dossier du script et dossier "settings" (doit correspondre au script principal p2.pyw)
DOSSIER_SCRIPT = Path(__file__).resolve().parent
DOSSIER_SETTINGS = DOSSIER_SCRIPT / "settings"

# Le chemin d'accès au fichier JSON temporaire (doit correspondre au script principal)
FICHIER_NOM_PDF_TEMP_JSON = DOSSIER_SETTINGS / "last_proforma_pdf.json"

# =================================================================
# --- FONCTION D'APPLICATION DU CACHET (MODIFIÉE POUR TOUTES LES PAGES) ---
# =================================================================

def ajouter_image_en_bas_a_gauche(pdf_path, image_path, largeur_image=120, marge_x=30, marge_y=30):
    """
    Ajoute une image (cachet) en bas à gauche de *chaque* page d'un PDF.
    Renomme le fichier de sortie en ajoutant '_hotel' au nom.
    """
    if not Path(image_path).exists():
        print(f"❌ ERREUR: L'image du cachet '{image_path}' est introuvable. Vérifiez que 'stamp.png' est présent.")
        try: os.startfile(pdf_path)
        except Exception: pass
        return

    doc = None
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"❌ ERREUR: Impossible d'ouvrir le PDF '{pdf_path}'. Détail: {e}")
        return

    if len(doc) == 0:
        print("❌ ERREUR: Le document PDF est vide.")
        doc.close()
        return

    # Dimensions de l'image (pour l'aspect ratio)
    try:
        pix = fitz.Pixmap(image_path)
    except Exception as e:
        print(f"❌ ERREUR: Impossible de charger l'image {image_path}. Détail: {e}")
        doc.close()
        return
        
    ratio = pix.height / pix.width
    hauteur_image = largeur_image * ratio

    
    # 🌟 NOUVEAU : Itérer sur toutes les pages
    nb_pages_cachetees = 0
    for i in range(len(doc)):
        page = doc[i]
        
        # Déterminer la position (bas-gauche de la page)
        hauteur_page = page.rect.height

        # Définir le rectangle d'insertion [x0, y0, x1, y1] (bas-gauche)
        x0 = marge_x 
        y1 = hauteur_page - marge_y # Le point y du bas
        x1 = x0 + largeur_image
        y0 = y1 - hauteur_image # Le point y du haut
        rect = fitz.Rect(x0, y0, x1, y1)

        # Insérer l'image
        try:
            page.insert_image(rect, filename=image_path)
            nb_pages_cachetees += 1
        except Exception as e:
            print(f"❌ ERREUR lors de l'insertion de l'image PyMuPDF à la page {i+1} : {e}")
            
    print(f"✅ Cachet ajouté à {nb_pages_cachetees} page(s).")
    
    # Sauvegarde du fichier estampillé
    output_path = Path(pdf_path)
    new_name = output_path.stem + "_hotel" + output_path.suffix
    output_hotel_path = output_path.parent / new_name

    try:
        # Enregistrement du nouveau fichier
        doc.save(output_hotel_path, garbage=4, deflate=True) 
        print(f"✅ Fichier 'hotel' sauvegardé sous: {output_hotel_path.name}")
        
        # ⭐ ÉTAPE CLÉ : Fermer l'objet document pour libérer le verrou sur le PDF brut
        doc.close() 

        # Suppression du fichier PDF brut temporaire
        output_path.unlink()
        print(f"🗑️ Fichier brut supprimé: {output_path.name}")
        
        # Ouvre le PDF final et estampillé
        os.startfile(output_hotel_path)
        
    except Exception as e:
        print(f"❌ ERREUR lors de la sauvegarde/suppression du fichier : {e}")
        # Si la suppression échoue, on essaie quand même d'ouvrir le PDF estampillé
        try:
            os.startfile(output_hotel_path)
        except:
            pass
    finally:
        # S'assurer que le document est fermé en cas d'échec antérieur
        if doc is not None and not doc.is_closed:
             doc.close()
             

# =================================================================
# --- UTILISATION DU SCRIPT (inchangé) ---
# =================================================================
if __name__ == "__main__":
    pdf_name = None
    try:
        with open(FICHIER_NOM_PDF_TEMP_JSON, 'r', encoding='utf-8') as f:
            data = json.load(f)
            pdf_name = data.get("last_proforma_name")
            if not pdf_name:
                raise ValueError("Clé 'last_proforma_name' manquante dans le JSON.")
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        print(f"❌ ERREUR: Impossible de lire le nom du PDF depuis le fichier JSON. Détail: {e}")
        sys.exit(1) 

    CHEMIN_PDF_BRUT = DOSSIER_PROFORMA / pdf_name
    NOM_DE_L_IMAGE = str(DOSSIER_SCRIPT / "stamp.png")
    LARGEUR_VOULUE = 120
    MARGE_DU_BORD_GAUCHE = 30
    MARGE_DU_BORD_BAS = 30

    print(f"Lecture réussie: Application du cachet sur '{pdf_name}'...")

    ajouter_image_en_bas_a_gauche(
        pdf_path=str(CHEMIN_PDF_BRUT), 
        image_path=NOM_DE_L_IMAGE,
        largeur_image=LARGEUR_VOULUE,
        marge_x=MARGE_DU_BORD_GAUCHE,
        marge_y=MARGE_DU_BORD_BAS
    )
