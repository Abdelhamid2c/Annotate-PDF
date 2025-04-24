import os
import re
import fitz  # PyMuPDF
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import threading

class YazakiPDFAnnotator(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuration de la fenêtre
        self.title("Yazaki PDF Annotator")
        self.geometry("800x600")
        self.minsize(600, 500)
        ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
        ctk.set_default_color_theme("blue")  # Thèmes: "blue" (default), "green", "dark-blue"
        
        # Variables pour stocker les chemins des fichiers
        self.excel_path = None
        self.pdf_path = None
        self.sheet_name = None
        self.output_path = None
        
        # Variables pour les colonnes par défaut
        self.circuit_column = "Wire Internal Name"
        self.sn_column = "SN FILS SIMPLE"
        
        # Créer l'UI
        self.create_ui()
    
    def create_ui(self):
        # Frame principal
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Logo Yazaki
        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "yazaki_logo.png")
            if os.path.exists(logo_path):
                logo_image = ctk.CTkImage(Image.open(logo_path), size=(200, 100))
                logo_label = ctk.CTkLabel(self.main_frame, image=logo_image, text="")
                logo_label.pack(pady=(0, 20))
            else:
                # Afficher un texte si l'image n'existe pas
                logo_text = ctk.CTkLabel(self.main_frame, text="YAZAKI", font=ctk.CTkFont(size=36, weight="bold"))
                logo_text.pack(pady=(0, 20))
        except Exception as e:
            print(f"Erreur lors du chargement du logo: {e}")
            logo_text = ctk.CTkLabel(self.main_frame, text="YAZAKI", font=ctk.CTkFont(size=36, weight="bold"))
            logo_text.pack(pady=(0, 20))
        
        # Titre
        title = ctk.CTkLabel(self.main_frame, text="Annotation de PDF avec données Excel", 
                            font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(0, 20))
        
        # Section Excel
        excel_frame = ctk.CTkFrame(self.main_frame)
        excel_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        excel_label = ctk.CTkLabel(excel_frame, text="1. Sélectionner le fichier Excel:")
        excel_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        excel_path_frame = ctk.CTkFrame(excel_frame)
        excel_path_frame.pack(fill="x", padx=10, pady=5)
        
        self.excel_path_var = ctk.StringVar()
        excel_path_entry = ctk.CTkEntry(excel_path_frame, textvariable=self.excel_path_var, width=500)
        excel_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        excel_browse_btn = ctk.CTkButton(excel_path_frame, text="Parcourir", command=self.browse_excel)
        excel_browse_btn.pack(side="right")
        
        # Champ pour le nom de la feuille Excel
        sheet_frame = ctk.CTkFrame(excel_frame)
        sheet_frame.pack(fill="x", padx=10, pady=(5, 10))
        
        sheet_label = ctk.CTkLabel(sheet_frame, text="Nom de la feuille (laisser vide pour la première):")
        sheet_label.pack(side="left", padx=(0, 10))
        
        self.sheet_var = ctk.StringVar()
        sheet_entry = ctk.CTkEntry(sheet_frame, textvariable=self.sheet_var, width=200)
        sheet_entry.pack(side="left", fill="x", expand=True)
        
        # Section PDF
        pdf_frame = ctk.CTkFrame(self.main_frame)
        pdf_frame.pack(fill="x", padx=10, pady=(15, 5))
        
        pdf_label = ctk.CTkLabel(pdf_frame, text="2. Sélectionner le fichier PDF à annoter:")
        pdf_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        pdf_path_frame = ctk.CTkFrame(pdf_frame)
        pdf_path_frame.pack(fill="x", padx=10, pady=5)
        
        self.pdf_path_var = ctk.StringVar()
        pdf_path_entry = ctk.CTkEntry(pdf_path_frame, textvariable=self.pdf_path_var, width=500)
        pdf_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        pdf_browse_btn = ctk.CTkButton(pdf_path_frame, text="Parcourir", command=self.browse_pdf)
        pdf_browse_btn.pack(side="right")
        
        # Bouton de traitement
        process_frame = ctk.CTkFrame(self.main_frame)
        process_frame.pack(fill="x", pady=20)
        
        self.process_btn = ctk.CTkButton(
            process_frame, 
            text="Traiter le fichier", 
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            command=self.process_file
        )
        self.process_btn.pack(pady=10)
        
        # Zone de log
        log_frame = ctk.CTkFrame(self.main_frame)
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        log_label = ctk.CTkLabel(log_frame, text="Journal d'exécution:")
        log_label.pack(anchor="w", padx=10, pady=(10, 0))
        
        self.log_text = ctk.CTkTextbox(log_frame)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)
    
    def log(self, message):
        """Ajouter un message au journal"""
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")  # Faire défiler pour voir la dernière ligne
        self.update_idletasks()  # Mettre à jour l'interface
    
    def browse_excel(self):
        """Ouvrir une boîte de dialogue pour sélectionner un fichier Excel"""
        filetypes = [("Fichiers Excel", "*.xlsx *.xls")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.excel_path = file_path
            self.excel_path_var.set(file_path)
            self.log(f"Fichier Excel sélectionné: {file_path}")
    
    def browse_pdf(self):
        """Ouvrir une boîte de dialogue pour sélectionner un fichier PDF"""
        filetypes = [("Fichiers PDF", "*.pdf")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.pdf_path = file_path
            self.pdf_path_var.set(file_path)
            # Créer également le chemin de sortie par défaut
            base_name = os.path.splitext(file_path)[0]
            self.output_path = f"{base_name}_avec_SN_FILS.pdf"
            self.log(f"Fichier PDF sélectionné: {file_path}")
            self.log(f"Fichier de sortie: {self.output_path}")
    
    def process_file(self):
        """Traiter le fichier PDF avec les données Excel"""
        # Vérifier si les fichiers ont été sélectionnés
        if not self.excel_path:
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel.")
            return
        
        if not self.pdf_path:
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier PDF.")
            return
        
        # Récupérer le nom de la feuille
        self.sheet_name = self.sheet_var.get().strip() or None
        
        # Désactiver le bouton pendant le traitement
        self.process_btn.configure(state="disabled", text="Traitement en cours...")
        
        # Lancer le traitement dans un thread séparé pour ne pas bloquer l'UI
        threading.Thread(target=self.run_processing).start()
    
    def run_processing(self):
        """Exécuter le traitement dans un thread séparé"""
        try:
            self.log("Démarrage du traitement...")
            self.log(f"Extraction des numéros de circuit du fichier {self.pdf_path}...")
            
            # Extraire les numéros de circuit
            circuit_info = extract_circuit_numbers(self.pdf_path)
            self.log(f"{len(circuit_info)} numéros de circuit trouvés.")
            
            self.log(f"Recherche des correspondances dans {self.excel_path}...")
            
            # Correspondre avec Excel
            matched_info = match_with_excel(
                circuit_info, 
                self.excel_path, 
                self.circuit_column, 
                self.sn_column,
                sheet_name=self.sheet_name
            )
            
            self.log(f"Ajout des annotations au PDF...")
            
            # Ajouter les annotations
            add_annotations_to_pdf(self.pdf_path, self.output_path, matched_info)
            
            self.log("Traitement terminé avec succès!")
            self.log(f"Résultat enregistré dans: {self.output_path}")
            
            # Proposer d'ouvrir le fichier résultat
            if messagebox.askyesno("Traitement terminé", f"Le fichier a été enregistré dans:\n{self.output_path}\n\nVoulez-vous l'ouvrir maintenant?"):
                try:
                    os.startfile(self.output_path)  # Windows
                except:
                    try:
                        import subprocess
                        subprocess.Popen(['xdg-open', self.output_path])  # Linux
                    except:
                        try:
                            import subprocess
                            subprocess.Popen(['open', self.output_path])  # macOS
                        except:
                            messagebox.showinfo("Information", "Impossible d'ouvrir le fichier automatiquement. Veuillez l'ouvrir manuellement.")
            
        except Exception as e:
            self.log(f"Erreur lors du traitement: {str(e)}")
            messagebox.showerror("Erreur", f"Une erreur s'est produite: {str(e)}")
        
        finally:
            # Réactiver le bouton
            self.process_btn.configure(state="normal", text="Traiter le fichier")


# Fonctions de traitement (reprises du code original)
def extract_circuit_numbers(pdf_path):
    """
    Extraire tous les numéros de circuit du PDF
    """
    circuit_info = []
    circuits_to_skip = set()  # Pour stocker les circuits à ignorer (associés à un J)
    
    # Ouvrir le PDF avec PyMuPDF
    doc = fitz.open(pdf_path)
    
    # Premier passage: identifier tous les numéros de circuit associés à un joint "J"
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        
        # Identifier les joints J et leurs circuits associés pour les ignorer
        for j_match in re.finditer(r'J\d+\s*\n(\d+)', text):
            circuit_to_skip = j_match.group(1)
            circuits_to_skip.add(circuit_to_skip)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        rotation = page.rotation
        
        # Obtenir les dimensions de la page
        page_width = page.rect.width
        page_height = page.rect.height
        
        # Extraire tous les part numbers (numéros suivis d'une lettre majuscule)
        part_numbers = []
        for part_match in re.finditer(r'\b\d+[A-Z]\b', text):
            part_number = part_match.group(0)
            if part_number not in part_numbers:  # Éviter les doublons
                part_numbers.append(part_number)
        
        # Rechercher des patterns comme "7/W0007,COFLRYB-0.35,GY/W"
        for match in re.finditer(r'(\d+)/W\d+,|J\+(\d+)\b', text):
            if match.group(1):  # Cas standard "X/WXXX,"
                circuit_num = match.group(1)
                match_text = match.group(0)
            else:  # Cas "J+XXX"
                circuit_num = match.group(2)
                match_text = match.group(0)
            
            # Vérifier si ce circuit doit être ignoré
            if circuit_num in circuits_to_skip:
                print(f"Circuit ignoré: {circuit_num} (associé à un joint)")
                continue
            
            # Obtenir les coordonnées du texte pour permettre l'annotation
            text_instances = page.search_for(match.group(0))
            
            if text_instances:
                position = text_instances[0]
                x0, y0, x1, y1 = position
                
                # Déterminer si le circuit est sur la moitié gauche ou droite de la page
                is_left_side = (x0 + x1) / 2 < page_width / 2
                
                circuit_info.append({
                    'page_num': page_num,
                    'circuit_number': circuit_num,
                    'match_text': match_text,
                    'rect': position,
                    'rotation': rotation,
                    'is_left_side': is_left_side,
                    'page_width': page_width,
                    'page_height': page_height,
                    'part_numbers': part_numbers  # Ajout des part numbers trouvés dans la page
                })
    
    doc.close()
    return circuit_info

def match_with_excel(circuit_info, excel_path, circuit_col='Numéro Circuit', sn_col='SN FILS SIMPLE', sheet_name=None):
    """
    Correspondre les numéros de circuit avec le fichier Excel, en utilisant les part numbers si disponibles
    """
    # Lire le fichier Excel
    try:
        # Si sheet_name est None, pandas retourne toutes les feuilles dans un dictionnaire
        # Dans ce cas, on prend la première feuille
        excel_data = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        # Vérifier si le résultat est un dictionnaire (cas où sheet_name=None)
        if isinstance(excel_data, dict):
            if not excel_data:  # Vérifier si le dictionnaire est vide
                raise ValueError("Le fichier Excel ne contient aucune feuille")
            
            # Utiliser la première feuille si sheet_name est None
            sheet_name = list(excel_data.keys())[0]
            df = excel_data[sheet_name]
            print(f"Utilisation de la première feuille: '{sheet_name}'")
        else:
            # Si excel_data est déjà un DataFrame, l'utiliser directement
            df = excel_data
            
    except ValueError as e:
        # Si le nom de feuille spécifié n'existe pas
        if "No sheet named" in str(e):
            # Afficher les noms de feuilles disponibles
            xls = pd.ExcelFile(excel_path)
            available_sheets = xls.sheet_names
            raise ValueError(f"Feuille '{sheet_name}' non trouvée. Feuilles disponibles: {', '.join(available_sheets)}")
        else:
            raise e
    
    # Nettoyer les noms de colonnes (enlever tout ce qui suit ":")
    df.columns = [str(col).split(':')[0] for col in df.columns]
    
    # Convertir la colonne de numéro de circuit en type numérique si nécessaire
    if circuit_col in df.columns and df[circuit_col].dtype != 'int64':
        df[circuit_col] = pd.to_numeric(df[circuit_col], errors='coerce')
    
    # Pour chaque circuit, chercher les correspondances dans Excel
    for entry in circuit_info:
        try:
            circuit_num = int(entry['circuit_number'])
            part_numbers = entry.get('part_numbers', [])
            
            # Vérifier si les part numbers de la page existent dans les colonnes de l'Excel
            matching_columns = [col for col in df.columns if col in part_numbers]
            
            if matching_columns:
                # Créer un masque pour les lignes qui contiennent la valeur "1" dans l'une des colonnes part_number
                mask = pd.Series(False, index=df.index)
                
                # Vérifier chaque colonne part_number
                for col in matching_columns:
                    mask = mask | df[col].astype(str).isin(["1", "1.0"])
                
                # Filtrer le DataFrame pour les lignes qui correspondent au part_number=1 et au circuit_num
                filtered_df = df[mask]
                matching_row = filtered_df[filtered_df[circuit_col] == circuit_num]
                
                if not matching_row.empty:
                    entry['sn_fils_simple'] = str(matching_row.iloc[0][sn_col])
                    # Ajouter SN GROUP si disponible
                    if "SN GROUP" in df.columns:
                        entry['sn_group'] = str(matching_row.iloc[0]["SN GROUP"])
                else:
                    # Si pas de correspondance avec part_number, chercher juste par circuit_num
                    matching_row = df[df[circuit_col] == circuit_num]
                    if not matching_row.empty:
                        entry['sn_fils_simple'] = str(matching_row.iloc[0][sn_col])
                        if "SN GROUP" in df.columns:
                            entry['sn_group'] = str(matching_row.iloc[0]["SN GROUP"])
                    else:
                        entry['sn_fils_simple'] = "Non trouvé"
                        entry['sn_group'] = ""
            else:
                # Si aucun part_number correspondant, chercher simplement par circuit_num
                matching_row = df[df[circuit_col] == circuit_num]
                if not matching_row.empty:
                    entry['sn_fils_simple'] = str(matching_row.iloc[0][sn_col])
                    if "SN GROUP" in df.columns:
                        entry['sn_group'] = str(matching_row.iloc[0]["SN GROUP"])
                else:
                    entry['sn_fils_simple'] = "Non trouvé"
                    entry['sn_group'] = ""
        except (ValueError, TypeError):
            # Si le circuit_number n'est pas convertible en entier
            entry['sn_fils_simple'] = "Erreur de format"
            entry['sn_group'] = ""
    
    return circuit_info

def add_annotations_to_pdf(pdf_path, output_path, annotations):
    """
    Ajouter des annotations au PDF existant avec positionnement adapté
    """
    # Ouvrir le PDF
    doc = fitz.open(pdf_path)
    
    # Grouper les annotations par page
    annotations_by_page = {}
    for ann in annotations:
        page_num = ann['page_num']
        if page_num not in annotations_by_page:
            annotations_by_page[page_num] = []
        annotations_by_page[page_num].append(ann)
    
    # Parcourir chaque page et ajouter les annotations
    for page_num in range(len(doc)):
        if page_num in annotations_by_page:
            page = doc[page_num]
            page_rotation = page.rotation
            
            for ann in annotations_by_page[page_num]:
                circuit_num = ann['circuit_number']
                sn_fils = ann['sn_fils_simple']
                sn_group = ann.get('sn_group', '')
                rect = ann['rect']
                is_left_side = ann['is_left_side']
                
                # Calculer la position de l'annotation
                x0, y0, x1, y1 = rect
                
                # Préparer le texte à ajouter
                if sn_group:
                    annotation_text = f"{sn_fils}"#f"{sn_fils} ({sn_group})"
                else:
                    annotation_text = f"{sn_fils}"
                
                # Déterminer la position en fonction de la position du circuit et de la rotation
                if page_rotation == 0:
                    if is_left_side:
                        text_point = fitz.Point(x0 - 90, y0 + (y1 - y0)/2)
                    else:
                        text_point = fitz.Point(x1 + 90, y0 + (y1 - y0)/2)
                
                elif page_rotation == 90:
                    if is_left_side:
                        text_point = fitz.Point(x0 + (x1 - x0)/2 + 50, y0 - 10)
                    else:
                        text_point = fitz.Point(x0 + (x1 - x0)/2 + 50, y1 + 10)
                
                elif page_rotation == 180:
                    if is_left_side:
                        text_point = fitz.Point(x1 + 110, y0 + (y1 - y0)/2)
                    else:
                        text_point = fitz.Point(x0 - 110, y0 + (y1 - y0)/2)
                
                elif page_rotation == 270:
                    if is_left_side:
                        text_point = fitz.Point(x0 + (x1 - x0)/2, y1 + 90)
                    else:
                        text_point = fitz.Point(x0 + (x1 - x0)/2, y0 - 90)
                
                # Ajouter l'annotation (texte en rouge)
                page.insert_text(
                    text_point,
                    annotation_text,
                    fontsize=10,
                    color=(1, 0, 0),  # Rouge (R,G,B)
                    rotate=page_rotation
                )
    
    # Enregistrer le PDF modifié
    doc.save(output_path)
    doc.close()
    
    return True

if __name__ == "__main__":
    # Créer un dossier "assets" s'il n'existe pas
    assets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
    os.makedirs(assets_dir, exist_ok=True)
    
    # Lancer l'application
    app = YazakiPDFAnnotator()
    app.mainloop()