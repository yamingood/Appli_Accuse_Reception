from nicegui import ui, events
import os
import pandas as pd
from datetime import date, datetime
from docxtpl import DocxTemplate
from PyPDF2 import PdfReader
import shutil
import tempfile
import platform

# Configuration des chemins
DOSSIER_BASE = os.path.join(os.path.expanduser("~"), "OneDrive - Cafdoc", "Documents", "DEVS", "Appli_Accuse_Reception")
definir_chemin = lambda *chemins: os.path.join(DOSSIER_BASE, *chemins)
SYSTEME = platform.system()
preview_pdf_active = SYSTEME == "Windows"

# Vérifier ou créer les dossiers nécessaires
def verifier_et_creer_repertoires():
    for dossier in ["template", "accuse_recep", "archive"]:
        os.makedirs(definir_chemin(dossier), exist_ok=True)

# Déplacer fichier vers dossier archive
def deplacer_fichier(temp_path, nom_original):
    archive_path = definir_chemin("archive", nom_original)
    if os.path.exists(archive_path):
        base, ext = os.path.splitext(archive_path)
        i = 1
        while os.path.exists(f"{base}_{i}{ext}"):
            i += 1
        archive_path = f"{base}_{i}{ext}"
    shutil.move(temp_path, archive_path)

# Extraire données
def extraire_donnees(fichier_excel, champs_attendus):
    df = pd.read_excel(fichier_excel)
    df.columns = [col.strip().replace(' ', '_') for col in df.columns]
    champs_manquants = [champ for champ in champs_attendus if champ not in df.columns]
    if champs_manquants:
        raise ValueError(f"Champs manquants : {', '.join(champs_manquants)}")
    return df[champs_attendus].to_dict(orient="records")

# Remplir les templates Word
def remplir_template(fichier_template, dossier_sortie, donnees_liste, dateDuJour):
    os.makedirs(dossier_sortie, exist_ok=True)
    premier_fichier = None
    for i, donnees in enumerate(donnees_liste):
        tpl = DocxTemplate(fichier_template)
        for cle, valeur in donnees.items():
            if pd.isna(valeur):
                donnees[cle] = ""
            elif isinstance(valeur, (datetime, pd.Timestamp)):
                donnees[cle] = valeur.strftime("%d/%m/%Y")
            else:
                donnees[cle] = str(valeur)

        tpl.render(donnees)
        matricule = str(donnees.get("Matricule", f"{i+1}")).strip()
        fichier_sortie = os.path.join(dossier_sortie, f"accuseReception_{matricule}_{dateDuJour.replace('/', '-')}.docx")
        tpl.save(fichier_sortie)
        if i == 0:
            premier_fichier = fichier_sortie
    return premier_fichier

# Convertir le Word en texte pour affichage
def lire_contenu_word(docx_path):
    try:
        import docx
        doc = docx.Document(docx_path)
        return '\n\n'.join([para.text for para in doc.paragraphs])
    except:
        return "(Impossible de lire le contenu)"

# Configuration du thème et styles personnalisés
def setup_custom_styles():
    ui.add_head_html('''
        <style>
            .gradient-bg {
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            }
            .card-shadow {
                box-shadow: 0 10px 25px rgba(0,0,0,0.1);
                border-radius: 16px;
            }
            .upload-zone {
                background: linear-gradient(45deg, #f8fafc, #e2e8f0);
                border: 2px dashed #94a3b8;
                transition: all 0.3s ease;
            }
            .upload-zone:hover {
                border-color: #667eea;
                background: linear-gradient(45deg, #f1f5f9, #ddd6fe);
            }
            .success-card {
                background: linear-gradient(135deg, #10b981, #059669);
            }
            .error-card {
                background: linear-gradient(135deg, #ef4444, #dc2626);
            }
            .info-card {
                background: linear-gradient(135deg, #3b82f6, #2563eb);
            }
            .preview-card {
                background: linear-gradient(135deg, #f8fafc, #f1f5f9);
                border: 1px solid #e2e8f0;
            }
            .title-gradient {
                background: linear-gradient(135deg, #667eea, #764ba2);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
            }
            .pulse-animation {
                animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
            }
            @keyframes pulse {
                0%, 100% { opacity: 1; }
                50% { opacity: .7; }
            }
        </style>
    ''')

# Interface principale améliorée
def create_main_interface():
    verifier_et_creer_repertoires()
    setup_custom_styles()

    with ui.row().classes('w-full justify-center gradient-bg p-8 mb-8'):
        with ui.column().classes('items-center text-white'):
            ui.icon('receipt_long', size='3rem').classes('mb-4')
            ui.label('Générateur d\'Accusés de Réception').classes('text-4xl font-bold text-center')
            ui.label('Automatisation des documents administratifs').classes('text-lg opacity-90 text-center mt-2')

    with ui.column().classes('max-w-4xl mx-auto p-6'):
        with ui.card().classes('card-shadow p-8 mb-6'):
            with ui.column().classes('w-full items-center'):
                ui.icon('cloud_upload', size='2.5rem').classes('text-blue-500 mb-4')
                ui.label('Téléchargement du fichier Excel').classes('text-2xl font-semibold mb-4 text-center')
                ui.label('Glissez-déposez votre fichier Excel (.xlsx) ou cliquez pour sélectionner').classes('text-gray-600 text-center mb-6')
                upload = ui.upload(
                    label='📊 Sélectionner le fichier Excel',
                    auto_upload=True,
                    max_files=1
                ).classes('upload-zone w-full p-8 rounded-xl')

        output = ui.column().classes('w-full')
        return upload, output

def afficher_message(message, type='info', output_container=None):
    config = {
        'success': {'icon': 'check_circle', 'class': 'success-card text-white'},
        'error': {'icon': 'error', 'class': 'error-card text-white'},
        'info': {'icon': 'info', 'class': 'info-card text-white'},
        'processing': {'icon': 'hourglass_empty', 'class': 'bg-gradient-to-r from-yellow-400 to-orange-500 text-white'}
    }
    style = config.get(type, config['info'])
    if output_container:
        with output_container:
            with ui.card().classes(f'card-shadow p-4 mb-4 {style["class"]}'):
                with ui.row().classes('w-full items-center gap-4'):
                    icon_class = 'pulse-animation text-2xl' if type == 'processing' else 'text-2xl'
                    ui.icon(style['icon']).classes(icon_class)
                    ui.label(message).classes('text-lg font-medium')

def afficher_preview(contenu, output_container):
    with output_container:
        with ui.card().classes('preview-card card-shadow p-6 mt-6'):
            with ui.row().classes('w-full items-center mb-4'):
                ui.icon('visibility', size='1.5rem').classes('text-blue-600')
                ui.label('Aperçu du document généré').classes('text-xl font-bold text-gray-800')
            ui.separator().classes('mb-4')
            with ui.scroll_area().classes('h-96 w-full'):
                ui.textarea(value=contenu, readonly=True, placeholder='Contenu du document...').classes('w-full min-h-full bg-white border-0 text-sm leading-relaxed')

def afficher_statistiques(donnees, output_container):
    nb_documents = len(donnees)
    with output_container:
        with ui.card().classes('card-shadow p-6 mb-4 bg-gradient-to-r from-purple-500 to-pink-500 text-white'):
            with ui.row().classes('w-full items-center justify-between'):
                with ui.column():
                    ui.label('Documents générés').classes('text-lg font-medium opacity-90')
                    ui.label(str(nb_documents)).classes('text-3xl font-bold')
                ui.icon('description', size='3rem').classes('opacity-80')

def handle_upload(e: events.UploadEventArguments):
    output.clear()
    afficher_message("🔄 Traitement du fichier en cours...", 'processing', output)
    dateDuJour = date.today().strftime("%d/%m/%Y")
    template_word = definir_chemin("template", "13. Accusé de réception déclaration d'impayés.docx")
    dossier_sortie = definir_chemin("accuse_recep", f"accuses_reception_{dateDuJour.replace('/', '-')}")
    champs_attendus = ["Date_Liq", "Matricule", "Identité_Allocataire", "Identité_Destinataire_bailleur", "Adresse_Ligne_2", "Adresse_Ligne_3", "Adresse_Ligne_4", "Adresse_Ligne_5", "Adresse_Ligne_6", "Adresse_Ligne_7"]
    try:
        temp_path = os.path.join(tempfile.gettempdir(), e.name)
        with open(temp_path, "wb") as f:
            f.write(e.content.read())
        donnees = extraire_donnees(temp_path, champs_attendus)
        premier_fichier = remplir_template(template_word, dossier_sortie, donnees, dateDuJour)
        deplacer_fichier(temp_path, e.name)
        output.clear()
        afficher_statistiques(donnees, output)
        afficher_message(f"✅ {len(donnees)} document(s) généré(s) avec succès dans le dossier :\n{dossier_sortie}", 'success', output)
        if premier_fichier:
            contenu = lire_contenu_word(premier_fichier)
            afficher_preview(contenu, output)
    except Exception as err:
        output.clear()
        afficher_message(f"❌ Erreur lors du traitement : {str(err)}", 'error', output)

upload, output = create_main_interface()
upload.on_upload(handle_upload)
ui.page_title('Générateur d\'Accusés de Réception')
ui.add_head_html('<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>📄</text></svg>">')
ui.run(title="Générateur d'Accusés de Réception", favicon='📄', host='0.0.0.0', port=8080)

