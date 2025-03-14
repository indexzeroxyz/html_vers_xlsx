import json
import pandas as pd
from bs4 import BeautifulSoup
import os
from tkinter import Tk, filedialog

# ✅ Boîte de dialogue pour sélectionner le fichier HTML
root = Tk()
root.withdraw()  # Masquer la fenêtre principale de tkinter
html_file = filedialog.askopenfilename(
    title="Sélectionner le fichier HTML",
    filetypes=[("Fichiers HTML", "*.html")]
)

if not html_file:
    print("🚨 Aucun fichier sélectionné. Opération annulée.")
    exit()

# Lecture du fichier HTML
with open(html_file, 'r', encoding='utf-8') as file:
    soup = BeautifulSoup(file, 'html.parser')

# Trouver la balise <script> avec l'id="reportsData"
script = soup.find('script', id='reportsData')

if script:
    try:
        # Extraction du JSON
        json_data = script.string.strip().replace("var reports = ", "").rstrip(";")

        # Conversion en JSON
        data = json.loads(json_data)

        # ✅ Récupération du nom du projet (sinon fallback au nom du fichier HTML)
        project_name = data.get('ProjectName', os.path.splitext(os.path.basename(html_file))[0])
        project_name_clean = project_name.replace(" ", "_").replace("/", "_").replace("\\", "_")

        # ✅ Définir le chemin du fichier Excel dans le même dossier que le fichier HTML
        output_excel = os.path.join(os.path.dirname(html_file), f"{project_name_clean}_UpgradeReport.xlsx")

        # Dictionnaire pour stocker les données par fichier Revit
        file_data = {}

        # Parcours des fichiers dans "UpgradedModels"
        upgraded_models = data.get('UpgradedModels', {})
        for key, models in upgraded_models.items():
            for model in models:
                model_name = model.get('ModelName', 'Inconnu')

                # Extraction des warnings, errors et document corruption sous "Resolved"
                resolved = model.get('Resolved', {})
                issues = (
                    resolved.get('Warnings', []) + 
                    resolved.get('Errors', []) + 
                    resolved.get('DocumentCorruption', [])  # 🔥 Ajout des "DocumentCorruption"
                )
                
                for issue in issues:
                    message = issue.get('Message', '')

                    # ✅ Identifier le type de problème
                    if '[Warning]' in message:
                        issue_type = '[Warning]'
                    elif '[Error]' in message:
                        issue_type = '[Error]'
                    elif '[DocumentCorruption]' in message:
                        issue_type = '[DocumentCorruption]'
                    else:
                        issue_type = 'Inconnu'

                    # Extraction des related elements
                    related_elements = issue.get('RelatedElements', [])

                    element_ids = ', '.join(str(e.get('ElementID', '')) for e in related_elements)
                    element_names = ', '.join(e.get('ElementName', '') for e in related_elements)
                    category_names = ', '.join(e.get('CategoryName', '') for e in related_elements)

                    # 🔥 Stocker dans le dictionnaire (clé = model_name)
                    if model_name not in file_data:
                        file_data[model_name] = []
                    
                    file_data[model_name].append({
                        'RevitFile': model_name,
                        'IssueType': issue_type,  # ✅ Type de problème (Warning, Error ou DocumentCorruption)
                        'Message': message,
                        'ElementID': element_ids,
                        'ElementName': element_names,
                        'CategoryName': category_names
                    })

        # ✅ Exportation vers Excel avec plusieurs feuilles
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for model_name, data in file_data.items():
                # Nettoyer le nom de la feuille (pas plus de 31 caractères et sans caractères spéciaux)
                safe_sheet_name = model_name[:31].replace('/', '_').replace('\\', '_').replace('*', '_')
                
                # Si le nom existe déjà, ajouter un suffixe
                counter = 1
                while safe_sheet_name in writer.sheets:
                    safe_sheet_name = f"{model_name[:28]}_{counter}"
                    counter += 1
                
                # Créer le DataFrame
                df = pd.DataFrame(data)
                
                # Exporter vers une feuille séparée
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

        print(f"✅ Fichier Excel créé avec succès : {output_excel}")

    except json.JSONDecodeError as e:
        print(f"❌ Erreur lors de la conversion JSON : {e}")

else:
    print("🚨 Balise <script> avec id='reportsData' non trouvée dans le fichier HTML.")