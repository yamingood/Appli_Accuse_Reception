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

# Convertir le Word en texte pour affichage (pas de docx2pdf ici)
def lire_contenu_word(docx_path):
    try:
        import docx
        doc = docx.Document(docx_path)
        return '\n\n'.join([para.text for para in doc.paragraphs])
    except:
        return "(Impossible de lire le contenu)"

# Interface principale
verifier_et_creer_repertoires()
ui.label('📄 Générateur d\'accusés de réception').classes('text-2xl font-bold')

output = ui.column().classes('w-full')

upload = ui.upload(label='Déposer un fichier Excel (.xlsx)', auto_upload=True, max_files=1)


def afficher_message(message, type='info'):
    couleur = {
        'success': 'green',
        'error': 'red',
        'info': 'blue'
    }.get(type, 'gray')
    with ui.row().classes('w-full items-center gap-3').style(f'color:{couleur}').with_parent(output):
        if type == 'success':
            ui.icon('check_circle')
        elif type == 'error':
            ui.icon('error')
        else:
            ui.icon('info')
        ui.label(message).classes('text-md')

def handle_upload(e: events.UploadEventArguments):
    output.clear()
    dateDuJour = date.today().strftime("%d/%m/%Y")
    template_word = definir_chemin("template", "13. Accusé de réception déclaration d'impayés.docx")
    dossier_sortie = definir_chemin("accuse_recep", f"accuses_reception_{dateDuJour.replace('/', '-')}")

    champs_attendus = [
        "Date_Liq", "Matricule", "Identité_Allocataire", "Identité_Destinataire_bailleur",
        "Adresse_Ligne_2", "Adresse_Ligne_3", "Adresse_Ligne_4", "Adresse_Ligne_5",
        "Adresse_Ligne_6", "Adresse_Ligne_7"
    ]

    try:
        temp_path = os.path.join(tempfile.gettempdir(), e.name)
        with open(temp_path, "wb") as f:
            f.write(e.content.read())

        donnees = extraire_donnees(temp_path, champs_attendus)
        premier_fichier = remplir_template(template_word, dossier_sortie, donnees, dateDuJour)
        deplacer_fichier(temp_path, e.name)

        afficher_message(f"✅ Documents générés dans : {dossier_sortie}", 'success')

        if premier_fichier:
            ui.label("🔎 Aperçu du premier document Word :").classes('text-lg font-bold mt-4').with_parent(output)
            contenu = lire_contenu_word(premier_fichier)
            ui.textarea(value=contenu, readonly=True, rows=20).classes('w-full bg-gray-50 border border-gray-300 rounded-md p-2').with_parent(output)

    except Exception as err:
        afficher_message(f"❌ Erreur : {str(err)}", 'error')
# Enregistrement de l’événement
upload.on_upload(handle_upload)

ui.run()
