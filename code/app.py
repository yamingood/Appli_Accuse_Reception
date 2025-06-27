import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import date, datetime
import os
import shutil
from io import BytesIO
from PyPDF2 import PdfReader
import time
from docx2pdf import convert
import platform
import zipfile
import io
import pythoncom  # <-- Ajout essentiel pour gérer COM

SYSTEME = platform.system()
preview_pdf_active = SYSTEME == "Windows"

# Configuration du dossier de base
DOSSIER_BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
definir_chemin = lambda *chemins: os.path.join(DOSSIER_BASE, *chemins)

# Créer les répertoires requis
def verifier_et_creer_repertoires():
    for dossier in ["template", "accuse_recep", "archive"]:
        os.makedirs(definir_chemin(dossier), exist_ok=True)

# Déplacer le fichier Excel dans le dossier archive
def deplacer_fichier(temp_path, nom_original):
    archive_path = definir_chemin("archive", nom_original)
    if os.path.exists(archive_path):
        base, ext = os.path.splitext(archive_path)
        i = 1
        while os.path.exists(f"{base}_{i}{ext}"):
            i += 1
        archive_path = f"{base}_{i}{ext}"
    shutil.move(temp_path, archive_path)

# Extraire les données depuis le fichier Excel
def extraire_donnees(fichier_excel, champs_attendus):
    try:
        df = pd.read_excel(fichier_excel)
        df.columns = [col.strip().replace(' ', '_') for col in df.columns]
        champs_manquants = [champ for champ in champs_attendus if champ not in df.columns]
        if champs_manquants:
            st.error(f"Champs manquants : {', '.join(champs_manquants)}")
            return None
        return df[champs_attendus].to_dict(orient="records")
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        return None

# Remplir les templates Word et convertir en PDF
def remplir_et_convertir(fichier_template, dossier_sortie, donnees_liste, horodatage, progress_bar=None, compteur_txt=None):
    pythoncom.CoInitialize()
    os.makedirs(dossier_sortie, exist_ok=True)
    premier_pdf = None
    total = len(donnees_liste)

    try:
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
            nom_base = f"accuseReception_{matricule}_{horodatage}"
            fichier_docx = os.path.join(dossier_sortie, f"{nom_base}.docx")
            fichier_pdf = fichier_docx.replace(".docx", ".pdf")

            tpl.save(fichier_docx)
            convert(fichier_docx, fichier_pdf)
            time.sleep(0.5)
            os.remove(fichier_docx)

            if i == 0:
                premier_pdf = fichier_pdf

            if progress_bar:
                progress_bar.progress((i + 1) / total)
            if compteur_txt:
                compteur_txt.text(f"📄 Fichiers générés : {i + 1} / {total}")
    finally:
        pythoncom.CoUninitialize()

    return premier_pdf

def creer_zip_depuis_dossier(dossier_path):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(dossier_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=dossier_path)
                zipf.write(file_path, arcname=arcname)
    zip_buffer.seek(0)
    return zip_buffer

# Convertir PDF en texte pour affichage
def convertir_en_pdf_et_lire(pdf_path):
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        return "\n\n".join(page.extract_text() or "" for page in reader.pages)

# UI Streamlit
st.set_page_config(page_title="Accusés de réception", layout="centered")
st.title("📄 Générateur d'accusés de réception")

verifier_et_creer_repertoires()
uploaded_file = st.file_uploader("Téléversez un fichier Excel", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner("Traitement du fichier..."):
        temp_file_path = definir_chemin("temp_uploaded.xlsx")
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        horodatage = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        template_word = definir_chemin("template", "template.docx")
        dossier_sortie = definir_chemin("accuse_recep", f"accuses_reception_{horodatage}")

        champs_attendus = [
            "Date_Liq", "Matricule", "Identité_Allocataire", "Identité_Destinataire_bailleur",
            "Adresse_Ligne_2", "Adresse_Ligne_3", "Adresse_Ligne_4", "Adresse_Ligne_5",
            "Adresse_Ligne_6", "Adresse_Ligne_7","Adresse_Ligne_2_Alloc", "Adresse_Ligne_3_Alloc", "Adresse_Ligne_4_Alloc", "Adresse_Ligne_5_Alloc", "Adresse_Ligne_6_Alloc", "Libellé_Allocataire", "Nom_Prénom_Allocataire"
        ]

        donnees_liste = extraire_donnees(temp_file_path, champs_attendus)
        if donnees_liste:
            progress_bar = st.progress(0)
            compteur_txt = st.empty()

            premier_pdf = remplir_et_convertir(template_word, dossier_sortie, donnees_liste, horodatage, progress_bar, compteur_txt)
            deplacer_fichier(temp_file_path, uploaded_file.name)

            st.success(f"✅ Documents générés dans : `{dossier_sortie}`")

            zip_buffer = creer_zip_depuis_dossier(dossier_sortie)
            zip_filename = f"accuses_reception_{horodatage}.zip"
            st.download_button(
                label="📦 Télécharger tous les accusés (.zip)",
                data=zip_buffer,
                file_name=zip_filename,
                mime="application/zip"
            )

            if premier_pdf and preview_pdf_active:
                st.subheader("🔎 Aperçu du premier accusé généré")
                try:
                    contenu = convertir_en_pdf_et_lire(premier_pdf)
                    st.text_area("Contenu du document (PDF)", value=contenu, height=300)
                except Exception as e:
                    st.warning(f"Impossible de prévisualiser le document : {e}")
            elif not preview_pdf_active:
                st.info("La prévisualisation PDF est désactivée (non supportée sur ce système).")
