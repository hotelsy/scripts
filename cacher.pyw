import os
import json
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tempfile
import time
import gc
import subprocess
import pythoncom
import win32com.client
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
from reportlab.lib.utils import ImageReader

CONFIG_FILE = "config.json"
SOURCE_DIR_FACTURES = r"\\192.168.100.20\TnOpenG\spool"
SOURCE_DIR_CONFIRM = r"\\192.168.100.20\TnOpenG\spool\Voucher"
TEMP_DIR = r"D:\\temp"
PRINT_TEMP_DIR = r"D:\\temp_print"
OLD_FACTURES_DIR = r"c:\\ancienfactures"

if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

if not os.path.exists(PRINT_TEMP_DIR):
    os.makedirs(PRINT_TEMP_DIR)
    
if not os.path.exists(OLD_FACTURES_DIR):
    os.makedirs(OLD_FACTURES_DIR, exist_ok=True)

def cleanup_print_temp():
    for f in os.listdir(PRINT_TEMP_DIR):
        try:
            os.remove(os.path.join(PRINT_TEMP_DIR, f))
        except:
            pass

cleanup_print_temp()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
STAMP_IMAGE_PATH = os.path.join(SCRIPT_DIR, "stamp.png")
STAMP_X_DEFAULT = 110
STAMP_Y_DEFAULT = 12

class LoadingWindow:
    def __init__(self, parent):
        self.root = tk.Toplevel(parent)
        self.root.title("Chargement...")
        self.root.geometry("300x100")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.root.update_idletasks()
        w, h = 300, 100
        ws, hs = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (ws // 2) - (w // 2), (hs // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        self.label = tk.Label(self.root, text="Traitement des documents...", pady=10)
        self.label.pack()

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", mode="determinate", length=250)
        self.progress_bar.pack(pady=10)

        self.root.grab_set()

    def update_progress(self, current, total):
        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = current
        self.root.update_idletasks()

    def on_close(self):
        messagebox.showwarning("Attention", "Veuillez attendre la fin du traitement.")

    def destroy(self):
        self.root.destroy()

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_config():
    config = {
        "year": year_var.get(),
        "dest_dir": dest_dir.get(),
        "print_last_page": bool(print_last_page_var.get()),
        "stamp_x": STAMP_X_DEFAULT,
        "stamp_y": STAMP_Y_DEFAULT,
        "doc_type": doc_type_var.get(),
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4)

def copy_to_temp(src_path):
    temp_file_path = os.path.join(TEMP_DIR, os.path.basename(src_path))
    shutil.copy2(src_path, temp_file_path)
    return temp_file_path

def add_stamp_to_pdf(input_pdf, output_pdf, image_path, x, y):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    img = ImageReader(image_path)
    for page in reader.pages:
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=A4)
        c.drawImage(img, x, y, mask='auto')
        c.save()
        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)
    with open(output_pdf, "wb") as f_out:
        writer.write(f_out)

def extract_pages_from_pdf(input_pdf, output_pdf, page_range):
    reader = PdfReader(input_pdf)
    num_pages = len(reader.pages)

    if not page_range:
        shutil.copy2(input_pdf, output_pdf)
        return

    if page_range.lower() == "last":
        writer = PdfWriter()
        writer.add_page(reader.pages[-1])
        with open(output_pdf, "wb") as f_out:
            writer.write(f_out)
        return

    writer = PdfWriter()
    normalized_range = page_range.replace(' ', '')
    page_indices = set()
    parts = normalized_range.split(',')

    for part in parts:
        if not part:
            continue
        if '-' in part:
            start_str, end_str = part.split('-')
            start_page = int(start_str) if start_str else 1
            end_page = int(end_str) if end_str else num_pages
            for page_num in range(start_page, end_page + 1):
                page_indices.add(page_num - 1)
        else:
            page_indices.add(int(part) - 1)

    for index in sorted(page_indices):
        if index < num_pages:
            writer.add_page(reader.pages[index])

    with open(output_pdf, "wb") as f_out:
        writer.write(f_out)

# Fonction unifiée pour gérer Word, Excel, Images et PDF
def process_document_file(input_path, output_pdf_path, page_range, image_path=None, x=100, y=100):
    pythoncom.CoInitialize()
    word = None
    doc = None
    excel = None
    workbook = None
    tmp_pdf_export = None
    stamped_tmp_pdf = None
    extracted_tmp_pdf = None
    file_extension = os.path.splitext(input_path)[1].lower()
    input_pdf_for_processing = None

    if file_extension == '.pdf':
        input_pdf_for_processing = input_path
    
    elif file_extension in ['.doc', '.docx', '.rtf']:
        # Traitement Word/RTF (via Word COM)
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(input_path, ReadOnly=False)
            if doc.ProtectionType != -1:
                try:
                    doc.Unprotect()
                except:
                    doc.Close(False)
                    raise
            
            tmp_pdf_export = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_pdf_export.close()
            doc.SaveAs(tmp_pdf_export.name, FileFormat=17) # 17 = wdFormatPDF
            input_pdf_for_processing = tmp_pdf_export.name
        finally:
            if doc:
                doc.Close(False)
            if word:
                word.Quit()
            pythoncom.CoUninitialize()
            gc.collect()
            time.sleep(1)
    
    elif file_extension in ['.xls', '.xlsx']:
        # Traitement Excel (via Excel COM)
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(input_path, ReadOnly=True)
            
            tmp_pdf_export = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            tmp_pdf_export.close()
            workbook.ExportAsFixedFormat(0, tmp_pdf_export.name, 0) # 0 = xlTypePDF
            input_pdf_for_processing = tmp_pdf_export.name
        finally:
            if workbook:
                workbook.Close(False)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()
            gc.collect()
            time.sleep(1)

    elif file_extension in ['.png', '.jpeg', '.jpg']:
        # Conversion Image en PDF (via ReportLab)
        tmp_pdf_export = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        tmp_pdf_export.close()
        
        c = canvas.Canvas(tmp_pdf_export.name, pagesize=A4)
        try:
            img = ImageReader(input_path)
        except Exception as e:
            raise ValueError(f"Erreur de lecture de l'image ({input_path}): {e}")

        img_width, img_height = img.getSize()
        a4_width, a4_height = A4
        
        # Calcul de la mise à l'échelle pour s'adapter à la page A4
        aspect_ratio = img_height / img_width
        scale_w = a4_width / img_width
        scale_h = a4_height / img_height
        
        scale = min(scale_w, scale_h)
        width = img_width * scale
        height = img_height * scale
            
        x_center = (a4_width - width) / 2
        y_center = (a4_height - height) / 2
        
        c.drawImage(img, x_center, y_center, width=width, height=height)
        c.save()
        input_pdf_for_processing = tmp_pdf_export.name
        pythoncom.CoUninitialize()
    
    else:
        raise ValueError(f"Format de fichier non pris en charge pour la conversion PDF: {file_extension}")

    # Section de traitement du PDF (Cachet et Extraction de pages)
    try:
        current_pdf_path = input_pdf_for_processing

        if image_path:
            stamped_tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            stamped_tmp_pdf.close()
            add_stamp_to_pdf(current_pdf_path, stamped_tmp_pdf.name, image_path, x, y)
            current_pdf_path = stamped_tmp_pdf.name

        if page_range:
            extracted_tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            extracted_tmp_pdf.close()
            extract_pages_from_pdf(current_pdf_path, extracted_tmp_pdf.name, page_range)
            current_pdf_path = extracted_tmp_pdf.name

        shutil.move(current_pdf_path, output_pdf_path)

    finally:
        # Nettoyage des fichiers temporaires
        if tmp_pdf_export and os.path.exists(tmp_pdf_export.name):
            try: os.remove(tmp_pdf_export.name)
            except: pass
        if stamped_tmp_pdf and os.path.exists(stamped_tmp_pdf.name):
            try: os.remove(stamped_tmp_pdf.name)
            except: pass
        if extracted_tmp_pdf and os.path.exists(extracted_tmp_pdf.name):
            try: os.remove(extracted_tmp_pdf.name)
            except: pass

def move_factures_to_archive(factures_folder, archive_folder):
    """Déplace tous les fichiers du dossier de factures vers le dossier d'archives."""
    if not os.path.exists(factures_folder):
        return

    if not os.path.exists(archive_folder):
        os.makedirs(archive_folder, exist_ok=True)
        
    for filename in os.listdir(factures_folder):
        src_path = os.path.join(factures_folder, filename)
        if os.path.isfile(src_path):
            try:
                base, ext = os.path.splitext(filename)
                dst_path = os.path.join(archive_folder, filename)
                counter = 1
                while os.path.exists(dst_path):
                    dst_path = os.path.join(archive_folder, f"{base}_{counter}{ext}")
                    counter += 1
                
                shutil.move(src_path, dst_path)
            except Exception as e:
                print(f"Erreur lors du déplacement de {filename} vers l'archive : {e}")

def clear_factures_folder():
    default_dir = r"C:\factures"
    factures_folder = dest_dir.get()
    
    if os.path.abspath(factures_folder).lower() != os.path.abspath(default_dir).lower():
        return
        
    move_factures_to_archive(factures_folder, OLD_FACTURES_DIR)

def save_factures():
    DEFAULT_DEST_DIR = r"C:\factures"
    current_doc_type = doc_type_var.get()
    custom_range = page_range_var.get().strip()

    if custom_range:
        page_range_setting = custom_range
    elif bool(print_last_page_var.get()):
        page_range_setting = "Last"
    else:
        page_range_setting = ""

    files_to_process = []
    is_choice_mode = (current_doc_type == "Choix")

    if is_choice_mode:
        global selected_files_list
        files_to_process = selected_files_list
        if not files_to_process:
            messagebox.showwarning("Attention", "Veuillez sélectionner au moins un fichier.")
            return
    else:
        factures = text_area.get("1.0", tk.END).strip().split("\n")
        factures = [f.strip() for f in factures if f.strip()]
        if not factures:
            messagebox.showwarning("Attention", "Veuillez entrer au moins un numéro.")
            return
        files_to_process = factures

    if not dest_dir.get():
        dest_dir.set(DEFAULT_DEST_DIR)

    current_dest_dir = dest_dir.get()

    if not os.path.exists(current_dest_dir):
        os.makedirs(current_dest_dir, exist_ok=True)

    clear_factures_folder()

    if not os.path.exists(STAMP_IMAGE_PATH):
        messagebox.showerror("Erreur Cachet", "stamp.png introuvable.")
        return

    image_path = STAMP_IMAGE_PATH
    x = STAMP_X_DEFAULT
    y = STAMP_Y_DEFAULT

    loading_win = LoadingWindow(root)
    total_factures = len(files_to_process)
    factures_traitees = 0
    erreurs = []
    prefix_base = filename_prefix_var.get().strip()
    year = year_var.get().strip()

    try:
        for i, item in enumerate(files_to_process):
            loading_win.update_progress(i, total_factures)
            temp_file = None
            file_name = ""
            src_path = ""

            try:
                if is_choice_mode:
                    src_path = item
                    file_name = os.path.basename(src_path)
                    filename_base = f"{prefix_base}{os.path.splitext(file_name)[0]}" 
                    file_to_process = src_path

                elif current_doc_type == "Facture":
                    num = item
                    file_name = f"F{year}{num}.rtf"
                    filename_base = f"{prefix_base}f{year}{num}"
                    src_path = os.path.join(SOURCE_DIR_FACTURES, file_name)
                    temp_file = copy_to_temp(src_path)
                    file_to_process = temp_file

                elif current_doc_type == "Reçu":
                    num = item
                    file_name = f"R{year}{num}.rtf"
                    filename_base = f"{prefix_base}r{year}{num}"
                    src_path = os.path.join(SOURCE_DIR_FACTURES, file_name)
                    temp_file = copy_to_temp(src_path)
                    file_to_process = temp_file

                elif current_doc_type == "Lettre de confirmation":
                    num = item
                    file_name = f"TNVOUCHHT_{num}.doc"
                    filename_base = f"{prefix_base}lettre{num}"
                    src_path = os.path.join(SOURCE_DIR_CONFIRM, file_name)
                    temp_file = copy_to_temp(src_path)
                    file_to_process = temp_file

                else:
                    erreurs.append(f"Type inconnu : {current_doc_type}")
                    continue

                if not os.path.exists(file_to_process):
                    erreurs.append(f"Introuvable : {src_path}")
                    continue

                dst = os.path.join(current_dest_dir, f"{filename_base}.pdf")
                counter = 1
                while os.path.exists(dst):
                    dst = os.path.join(current_dest_dir, f"{filename_base}_{counter}.pdf")
                    counter += 1

                # Appel à la fonction mise à jour
                process_document_file(file_to_process, dst, page_range_setting, image_path, x, y)
                factures_traitees += 1

            except Exception as e:
                erreurs.append(f"{file_name} : {e}")

            finally:
                if temp_file and os.path.exists(temp_file):
                    try: os.remove(temp_file)
                    except: pass

        loading_win.update_progress(total_factures, total_factures)

    finally:
        loading_win.destroy()

    save_config()

    message = (
        f"Les factures sont sauvegardées temporairement dans {DEFAULT_DEST_DIR} "
        f"et seront déplacées automatiquement lors de votre prochaine sauvegarde vers {OLD_FACTURES_DIR}."
    )

    if erreurs:
        messagebox.showwarning(
            "Partiel", 
            f"{factures_traitees} documents sauvegardés.\n{len(erreurs)} erreurs :\n\n" + 
            "\n".join(erreurs) + 
            f"\n\n{message}"
        )
    else:
        messagebox.showinfo(
            "Succès", 
            f"{factures_traitees} documents sauvegardés.\n\n"
            f"{message}"
        )
        
    time.sleep(1)
    try:
        subprocess.Popen(f'explorer "{current_dest_dir}"')
    except Exception as e:
        messagebox.showwarning("Erreur", f"Impossible d'ouvrir le dossier de destination : {e}")

def print_factures():
    current_doc_type = doc_type_var.get()
    custom_range = page_range_var.get().strip()

    if custom_range:
        page_range_setting = custom_range
    elif bool(print_last_page_var.get()):
        page_range_setting = "Last"
    else:
        page_range_setting = ""

    files_to_process = []
    is_choice_mode = (current_doc_type == "Choix")

    if is_choice_mode:
        global selected_files_list
        files_to_process = selected_files_list
        if not files_to_process:
            messagebox.showwarning("Attention", "Sélectionnez un fichier.")
            return
    else:
        factures = text_area.get("1.0", tk.END).strip().split("\n")
        factures = [f.strip() for f in factures if f.strip()]
        if not factures:
            messagebox.showwarning("Attention", "Entrez au moins un numéro.")
            return
        files_to_process = factures

    loading_win = LoadingWindow(root)
    total_factures = len(files_to_process)
    year = year_var.get().strip()
    erreurs = []
    x = STAMP_X_DEFAULT
    y = STAMP_Y_DEFAULT

    try:
        for i, item in enumerate(files_to_process):
            loading_win.update_progress(i, total_factures)
            temp_file = None
            file_name = ""
            tmp_pdf_temp_dest = None
            src_path = ""

            try:
                if is_choice_mode:
                    src_path = item
                    file_name = os.path.basename(src_path)
                    file_to_process = src_path

                elif current_doc_type == "Facture":
                    num = item
                    file_name = f"F{year}{num}.rtf"
                    src_path = os.path.join(SOURCE_DIR_FACTURES, file_name)
                    temp_file = copy_to_temp(src_path)
                    file_to_process = temp_file

                elif current_doc_type == "Reçu":
                    num = item
                    file_name = f"R{year}{num}.rtf"
                    src_path = os.path.join(SOURCE_DIR_FACTURES, file_name)
                    temp_file = copy_to_temp(src_path)
                    file_to_process = temp_file

                elif current_doc_type == "Lettre de confirmation":
                    num = item
                    file_name = f"TNVOUCHHT_{num}.doc"
                    src_path = os.path.join(SOURCE_DIR_CONFIRM, file_name)
                    temp_file = copy_to_temp(src_path)
                    file_to_process = temp_file

                else:
                    erreurs.append(f"Type inconnu.")
                    continue

                if not os.path.exists(file_to_process):
                    erreurs.append(f"Introuvable : {src_path}")
                    continue

                tmp_pdf_temp_dest = os.path.join(
                    PRINT_TEMP_DIR,
                    f"print_{int(time.time()*1000)}.pdf"
                )

                # Appel à la fonction mise à jour
                process_document_file(
                    file_to_process,
                    tmp_pdf_temp_dest,
                    page_range_setting,
                    None,
                    x, y
                )

                os.startfile(tmp_pdf_temp_dest, "print")

            except Exception as e:
                erreurs.append(f"{file_name} : {e}")

            finally:
                if temp_file and os.path.exists(temp_file):
                    try: os.remove(temp_file)
                    except: pass

        loading_win.update_progress(total_factures, total_factures)

    finally:
        loading_win.destroy()

    if erreurs:
        messagebox.showwarning("Erreurs", "\n".join(erreurs))
    else:
        messagebox.showinfo("Succès", f"{total_factures} documents envoyés à l'impression.")

def update_file_display(files_list):
    if not files_list:
        selected_file_path_var.set("Aucun fichier sélectionné")
    elif len(files_list) == 1:
        selected_file_path_var.set(os.path.basename(files_list[0]))
    else:
        selected_file_path_var.set(f"{len(files_list)} fichiers sélectionnés.")

def choose_files():
    global selected_files_list
    # Mise à jour des types de fichiers acceptés
    files_selected = filedialog.askopenfilenames(
        title="Sélectionner des documents",
        filetypes=[("Documents supportés", "*.doc *.docx *.rtf *.pdf *.xls *.xlsx *.png *.jpeg *.jpg"),
                   ("Tous les fichiers", "*.*")]
    )
    if files_selected:
        selected_files_list = list(files_selected)
        update_file_display(selected_files_list)

def add_more_files():
    global selected_files_list
    # Mise à jour des types de fichiers acceptés
    files_added = filedialog.askopenfilenames(
        title="Ajouter des documents",
        filetypes=[("Documents supportés", "*.doc *.docx *.rtf *.pdf *.xls *.xlsx *.png *.jpeg *.jpg"),
                   ("Tous les fichiers", "*.*")]
    )
    if files_added:
        selected_files_list = list(set(selected_files_list).union(files_added))
        update_file_display(selected_files_list)

def choose_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        dest_dir.set(folder_selected)

def show_context_menu(event):
    context_menu = tk.Menu(root, tearoff=0)
    context_menu.add_command(label="Coller", command=paste_from_clipboard)
    context_menu.tk_popup(event.x_root, event.y_root)

def paste_from_clipboard():
    try:
        clipboard_content = root.clipboard_get()
        text_area.insert(tk.END, clipboard_content)
    except:
        pass

def update_last_page_checkbox(*args):
    if page_range_var.get().strip():
        check_last_page.config(state=tk.DISABLED)
        print_last_page_var.set(0)
    else:
        check_last_page.config(state=tk.NORMAL)

def update_gui(*args):
    if doc_type_var.get() == "Choix":
        frame_year.pack_forget()
        label_text_area.pack_forget()
        text_area.pack_forget()
        frame_file_choice.pack(before=frame_page_range, pady=5)
        update_file_display(selected_files_list)
    else:
        frame_file_choice.pack_forget()
        frame_year.pack(after=frame_type, pady=5)
        label_text_area.pack(after=frame_year)
        text_area.pack(after=label_text_area, pady=5)

    update_last_page_checkbox()

# --- Initialisation de la Fenêtre Racine (Utilise tk.Tk standard) ---
root = tk.Tk()
root.title("Gestion des Documents")
root.update_idletasks()
w, h = 550, 500
ws, hs = root.winfo_screenwidth(), root.winfo_screenheight()
x, y = (ws // 2) - (w // 2), (hs // 2)
root.geometry(f"{w}x{h}+{x}+{y}") # C'EST ICI QUE LE CENTRAGE EST FAIT
# -------------------------------------------------------------------

config = load_config()
DEFAULT_DEST_DIR = r"C:\factures"

doc_type_var = tk.StringVar(value=config.get("doc_type", "Facture"))
year_var = tk.StringVar(value=config.get("year", "2025"))
dest_dir = tk.StringVar(value=config.get("dest_dir", DEFAULT_DEST_DIR))
filename_prefix_var = tk.StringVar(value=config.get("filename_prefix", ""))
print_last_page_var = tk.IntVar(value=1 if config.get("print_last_page", False) else 0)
page_range_var = tk.StringVar(value="")
page_range_var.trace_add("write", update_last_page_checkbox)

selected_file_path_var = tk.StringVar(value="Aucun fichier sélectionné")
selected_files_list = []
doc_type_var.trace_add("write", update_gui)

frame_type = tk.Frame(root)
frame_type.pack(pady=5)
tk.Label(frame_type, text="Type de document:").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame_type, text="Facture", variable=doc_type_var, value="Facture").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame_type, text="Reçu", variable=doc_type_var, value="Reçu").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame_type, text="Lettre de confirmation", variable=doc_type_var, value="Lettre de confirmation").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame_type, text="Choix", variable=doc_type_var, value="Choix").pack(side=tk.LEFT, padx=5)

frame_year = tk.Frame(root)
tk.Label(frame_year, text="Année:").pack(side=tk.LEFT, padx=5)
tk.Entry(frame_year, textvariable=year_var, width=6).pack(side=tk.LEFT)

label_text_area = tk.Label(root, text="Entrez les numéros (un par ligne):")
text_area = tk.Text(root, height=10, width=30)
text_area.bind("<Button-3>", show_context_menu)

frame_file_choice = tk.Frame(root, bd=2, relief=tk.GROOVE)
tk.Label(frame_file_choice, text="Documents sélectionnés:").pack(side=tk.TOP, pady=5)

frame_top_line = tk.Frame(frame_file_choice)
frame_top_line.pack(pady=2)

tk.Entry(frame_top_line, textvariable=selected_file_path_var, width=35, state='readonly').pack(side=tk.LEFT, padx=5)
tk.Button(frame_top_line, text="Sélectionner...", command=choose_files, width=15).pack(side=tk.LEFT, padx=5)

frame_bottom_line = tk.Frame(frame_file_choice)
frame_bottom_line.pack(pady=2)
tk.Button(frame_bottom_line, text="Ajouter d'autres...", command=add_more_files, width=15).pack(pady=2, padx=5)

frame_page_range = tk.Frame(root)
frame_page_range.pack(pady=5)
tk.Label(frame_page_range, text="Pages (ex: 1,3-5,8-):").pack(side=tk.LEFT)
tk.Entry(frame_page_range, textvariable=page_range_var, width=20).pack(side=tk.LEFT, padx=5)

frame_last_page = tk.Frame(root)
frame_last_page.pack(pady=5)
check_last_page = tk.Checkbutton(frame_last_page, text="Dernière page uniquement", variable=print_last_page_var)
check_last_page.pack()

frame_prefix = tk.Frame(root)
frame_prefix.pack(pady=5)
tk.Label(frame_prefix, text="Nom de fichier de base (optionnel):").pack(side=tk.LEFT, padx=5)
tk.Entry(frame_prefix, textvariable=filename_prefix_var, width=10).pack(side=tk.LEFT)

frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)
tk.Button(frame_buttons, text="Sauvegarder les documents", command=save_factures, width=30, height=2, bg='light green').pack(side=tk.LEFT, padx=10)
tk.Button(frame_buttons, text="Imprimer les documents", command=print_factures, width=30, height=2, bg='light blue').pack(side=tk.LEFT, padx=10)

update_gui()

root.mainloop()
