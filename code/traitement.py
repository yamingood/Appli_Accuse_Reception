import pandas as pd
from docx import Document
from datetime import date
import os
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import shutil

# Définition du chemin jusqu'à Documents
dossier_base = os.path.join(os.path.expanduser("~"), "OneDrive - Cafdoc", "Documents", "DEVS", "Appli_Accuse_Reception")

def definir_chemin(*chemins):
    return os.path.join(dossier_base, *chemins)

# Fonction pour s'assurer que tous les répertoires nécessaires existent
def verifier_et_creer_repertoires():
    dossiers_requis = [
        "template",
        "accuse_recep",
        "archive",
    ]
    for dossier in dossiers_requis:
        chemin = definir_chemin(dossier)
        os.makedirs(chemin, exist_ok=True)  # Crée le dossier s'il n'existe pas
        print(f"📂 Vérification : {chemin} - ✅ OK")

# Fonction pour sélectionner un fichier via une boîte de dialogue
def choisir_fichier():
    root = tk.Tk()
    root.withdraw()
    fichier = filedialog.askopenfilename(title="Sélectionnez un fichier Excel",
                                         filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
    return fichier

# Fonction pour déplacer le fichier traité
def deplacer_fichier(fichier):
    try:
        dossier_destination = definir_chemin("archive")
        os.makedirs(dossier_destination, exist_ok=True)
        
        chemin_destination = os.path.join(dossier_destination, os.path.basename(fichier))
        #chemin_destination = dossier_destination

        if not os.path.exists(fichier):  # Vérifie si le fichier source existe
            print(f"❌ Le fichier source n'existe pas : {fichier}")
            return

        if os.path.exists(chemin_destination):  # Vérifie si le fichier destination existe déjà
            print(f"⚠️ Le fichier {chemin_destination} existe déjà. Renommage en cours...")
            base, ext = os.path.splitext(chemin_destination)
            i = 1
            while os.path.exists(f"{base}_{i}{ext}"):
                i += 1
            chemin_destination = f"{base}_{i}{ext}"

        # Afficher le contenu du répertoire archive
        print(f"📂 Contenu du dossier {dossier_destination} :")
        print(os.listdir(dossier_destination))
        shutil.move(fichier, chemin_destination)
        print(f"✅ Fichier déplacé vers : {chemin_destination}")

    except Exception as e:
        print(f"❌ Erreur lors du déplacement du fichier : {e}")


# Fonction pour vérifier et extraire les données du fichier Excel
def extraire_donnees(fichier_excel, champs_attendus):
    try:
        df = pd.read_excel(fichier_excel)
        df.columns = [col.strip() for col in df.columns]

        champs_manquants = [champ for champ in champs_attendus if champ not in df.columns]
        if champs_manquants:
            print(f"⚠️ Champs manquants : {', '.join(champs_manquants)}")
            return None

        return df[champs_attendus].to_dict(orient="records")
    except FileNotFoundError:
        print("❌ Fichier non trouvé.")
    except Exception as e:
        print(f"❌ Erreur : {e}")
    return None

# Fonction pour remplir un template Word
def remplir_template(fichier_template, dossier_sortie, donnees_liste, dateDuJour):
    try:
        os.makedirs(dossier_sortie, exist_ok=True)
        for donnees in donnees_liste:
            doc = Document(fichier_template)
            
            def traiter_paragraphe(paragraphe):
                for cle, valeur in donnees.items():
                    placeholder = f"{{{{ {cle} }}}}"
                    if pd.isna(valeur):
                        valeur = ""
                    if isinstance(valeur, (datetime, pd.Timestamp)):
                        valeur = valeur.strftime("%d/%m/%Y")
                    paragraphe.text = paragraphe.text.replace(placeholder, str(valeur))

            for paragraphe in doc.paragraphs:
                traiter_paragraphe(paragraphe)

            matricule = str(donnees.get("Matricule", "inconnu")).strip()
            fichier_sortie = os.path.join(dossier_sortie, f"accuseReception_{matricule}_{dateDuJour.replace('/', '-')}.docx")
            doc.save(fichier_sortie)
            print(f"✅ Document généré : {fichier_sortie}")
    except Exception as e:
        print(f"❌ Erreur : {e}")

# --- Exécution du programme ---
verifier_et_creer_repertoires() # Vérifie et crée les répertoires nécessaires
fichier_excel = choisir_fichier() # Sélectionner un fichier Excel

if fichier_excel:
    dateDuJour = date.today().strftime("%d/%m/%Y") # Date du jour au format JJ/MM/AAAA
    template_word = definir_chemin("template", "13. Accusé de réception déclaration d'impayés.docx") # Chemin du template Word
    dossier_sortie = definir_chemin("accuse_recep", f"accuses_reception_{dateDuJour.replace('/', '-')}") # Dossier de sortie

    champs_attendus = ["Date Liq", "Matricule", "Identité Allocataire", "Identité Destinataire bailleur",
                       "Adresse Ligne 2", "Adresse Ligne 3", "Adresse Ligne 4",
                       "Adresse Ligne 5", "Adresse Ligne 6", "Adresse Ligne 7"]

    donnees_liste = extraire_donnees(fichier_excel, champs_attendus) # Extraire les données du fichier Excel
    if not os.path.exists(fichier_excel):  # 🔍 Vérification après extraction
        print(f"❌ Problème : le fichier {fichier_excel} a été supprimé après l'extraction !")
    else:
        if donnees_liste:
            remplir_template(template_word, dossier_sortie, donnees_liste, dateDuJour)
            deplacer_fichier(fichier_excel)
else:
    print("❌ Aucun fichier sélectionné. Opération annulée.")
