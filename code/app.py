import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import date, datetime
import os
import shutil
from io import BytesIO
from PyPDF2 import PdfReader
from docx2pdf import convert
import platform
import zipfile
import io

# Détection système
SYSTEME = platform.system()
preview_pdf_active = SYSTEME == "Windows"

# 📁 Dossier base = un niveau au-dessus du dossier courant (../)
DOSSIER_BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
definir_chemin = lambda *chemins: os.path.join(DOSSIER_BASE, *chemins)

# ✅ Crée les dossiers si besoin
def verifier_et_creer_repertoires():
    for dossier in ["template", "accuse_recep", "archive"]:
        os.makedirs(definir_chemin(dossier), exist_ok=True)

# 🔁 Déplace le fichier Excel dans un dossier archive
def deplacer_fichier(temp_path, nom_original):
    archive_path = definir_chemin("archive", nom_original)
    if os.path.exists(archive_path):
        base, ext = os.path.splitext(archive_path)
        i = 1
        while os.path.exists(f"{base}_{i}{ext}"):
            i += 1
        archive_path = f"{base}_{i}{ext}"
    shutil.move(temp_path, archive_path)

# 🔍 Extraction des données
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

# 🧾 Remplissage des templates Word
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

# 📦 Zipper le dossier
def creer_zip_depuis_dossier(dossier_path):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(dossier_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=dossier_path)
                zipf.write(file_path, arcname=arcname)
    zip_buffer.seek(0)
    return zip_buffer

# 🧾 Lire le contenu du PDF pour affichage
def convertir_en_pdf_et_lire(docx_path):
    temp_pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, temp_pdf_path)
    with open(temp_pdf_path, "rb") as f:
        reader = PdfReader(f)
        text = "\n\n".join(page.extract_text() or "" for page in reader.pages)
    return text

# 🖼️ Interface utilisateur
st.set_page_config(page_title="Accusés de réception", layout="centered")
st.title("📄 Générateur d'accusés de réception")

verifier_et_creer_repertoires()

uploaded_file = st.file_uploader("Téléversez un fichier Excel", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner("Traitement du fichier..."):
        temp_file_path = definir_chemin("temp_uploaded.xlsx")
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        dateDuJour = date.today().strftime("%d/%m/%Y")
        template_word = definir_chemin("template", "template.docx")
        dossier_sortie = definir_chemin("accuse_recep", f"accuses_reception_{dateDuJour.replace('/', '-')}")

        champs_attendus = [
            "Date_Liq", "Matricule", "Identité_Allocataire", "Identité_Destinataire_bailleur",
            "Adresse_Ligne_2", "Adresse_Ligne_3", "Adresse_Ligne_4", "Adresse_Ligne_5",
            "Adresse_Ligne_6", "Adresse_Ligne_7"
        ]

        donnees_liste = extraire_donnees(temp_file_path, champs_attendus)

        if donnees_liste:
            premier_fichier = remplir_template(template_word, dossier_sortie, donnees_liste, dateDuJour)
            deplacer_fichier(temp_file_path, uploaded_file.name)
            st.success("✅ Documents générés avec succès !")

            zip_buffer = creer_zip_depuis_dossier(dossier_sortie)
            st.download_button(
                label="📦 Télécharger tous les accusés (.zip)",
                data=zip_buffer,
                file_name=f"accuses_reception_{dateDuJour.replace('/', '-')}.zip",
                mime="application/zip"
            )


