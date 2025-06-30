import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
import shutil
from io import BytesIO
from PyPDF2 import PdfReader
import platform
import zipfile
import base64
import subprocess

SYSTEME = platform.system()
preview_pdf_active = SYSTEME in ["Windows", "Linux", "Darwin"]

# Configuration du dossier de base
DOSSIER_BASE = os.path.abspath(os.path.join(os.path.dirname(__file__)))
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

# Conversion avec LibreOffice
def convert_with_libreoffice(docx_path, pdf_path):
    try:
        sortie_dir = os.path.dirname(pdf_path)

        # Tentative automatique de détection
        soffice_path = shutil.which("soffice")
        if not soffice_path:
            soffice_path = r"C:\Program Files (x86)\LibreOffice 4\program\soffice.exe"  # Modifie ce chemin si besoin

        if not os.path.exists(soffice_path):
            raise FileNotFoundError("LibreOffice (soffice) introuvable. Vérifie son installation ou ajoute le chemin dans le PATH.")

        subprocess.run([
            soffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", sortie_dir,
            docx_path
        ], check=True)

        if os.path.exists(docx_path):
            os.remove(docx_path)

        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"Le fichier PDF attendu n'a pas été généré : {pdf_path}")

    except Exception as e:
        st.error(f"Erreur de conversion PDF avec LibreOffice : {e}")

# Remplir les templates Word et convertir en PDF
def remplir_et_convertir(fichier_template, dossier_sortie, donnees_liste, horodatage, progress_bar=None, compteur_txt=None):
    os.makedirs(dossier_sortie, exist_ok=True)
    premier_pdf = None
    total = len(donnees_liste)

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
        fichier_pdf = os.path.join(dossier_sortie, f"{nom_base}.pdf")

        tpl.save(fichier_docx)
        convert_with_libreoffice(fichier_docx, fichier_pdf)

        if i == 0:
            premier_pdf = fichier_pdf

        if progress_bar:
            progress_bar.progress((i + 1) / total)
        if compteur_txt:
            compteur_txt.text(f"📄 Fichiers générés : {i + 1} / {total}")

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
    if "zip_buffer" not in st.session_state or "horodatage" not in st.session_state or uploaded_file.name != st.session_state.get("nom_fichier"):
        with st.spinner("Traitement du fichier..."):
            # Sauvegarde temporaire
            temp_file_path = definir_chemin("temp_uploaded.xlsx")
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Préparation des chemins
            horodatage = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
            template_word = definir_chemin("template", "template.docx")
            dossier_sortie = definir_chemin("accuse_recep", f"accuses_reception_{horodatage}")

            # Champs requis
            champs_attendus = [
                "Date_Liq", "Matricule", "Identité_Allocataire", "Identité_Destinataire_bailleur",
                "Adresse_Ligne_2", "Adresse_Ligne_3", "Adresse_Ligne_4", "Adresse_Ligne_5",
                "Adresse_Ligne_6", "Adresse_Ligne_7", "Adresse_Ligne_2_Alloc", "Adresse_Ligne_3_Alloc",
                "Adresse_Ligne_4_Alloc", "Adresse_Ligne_5_Alloc", "Adresse_Ligne_6_Alloc",
                "Libellé_Allocataire", "Nom_Prénom_Allocataire"
            ]

            donnees_liste = extraire_donnees(temp_file_path, champs_attendus)
            if donnees_liste:
                progress_bar = st.progress(0)
                compteur_txt = st.empty()

                premier_pdf = remplir_et_convertir(template_word, dossier_sortie, donnees_liste, horodatage, progress_bar, compteur_txt)

                deplacer_fichier(temp_file_path, uploaded_file.name)

                st.success(f"✅ Documents générés dans : `{dossier_sortie}`")

                # Stockage dans session_state
                st.session_state.zip_buffer = creer_zip_depuis_dossier(dossier_sortie)
                st.session_state.zip_filename = f"accuses_reception_{horodatage}.zip"
                st.session_state.nom_fichier = uploaded_file.name
                st.session_state.horodatage = horodatage
                st.session_state.premier_pdf = premier_pdf

    # Bouton de téléchargement (aucune régénération ici)
    st.download_button(
        label="📦 Télécharger tous les accusés (.zip)",
        data=st.session_state.zip_buffer,
        file_name=st.session_state.zip_filename,
        mime="application/zip"
    )

    # Affichage PDF si supporté
    if preview_pdf_active and st.session_state.get("premier_pdf"):
        st.subheader("🔎 Aperçu du premier accusé généré (PDF)")
        try:
            with open(st.session_state["premier_pdf"], "rb") as f:
                base64_pdf = base64.b64encode(f.read()).decode("utf-8")
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)
        except Exception as e:
            st.warning(f"Impossible d'afficher le PDF : {e}")
    elif not preview_pdf_active:
        st.info("La prévisualisation PDF est désactivée sur ce système.")

