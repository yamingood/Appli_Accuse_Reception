import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from datetime import date
from traitement import deplacer_fichier, extraire_donnees, remplir_template

# Fonction pour définir les chemins dynamiquement
def definir_chemins():
    base_dir = os.path.join(os.path.expanduser("~"), "OneDrive - Cafdoc", "Documents", "DEVS", "Appli_Accuse_Reception")
    return {
        "dossier_a_surveiller": os.path.join(base_dir, "a traiter"),
        "template_word": os.path.join(base_dir, "template", "13. Accusé de réception déclaration d'impayés.docx"),
        "dossier_sortie": os.path.join(base_dir, "accuse_recep", f"accuses_reception_{date.today().strftime('%d-%m-%Y')}"),
        "dossier_traite": os.path.join(base_dir, "archive")
    }

class MoniteurDossier(FileSystemEventHandler):
    def __init__(self, dossier_a_surveiller, template_word, dossier_sortie, champs_attendus, dateDuJour, dossier_traite):
        self.dossier_a_surveiller = dossier_a_surveiller
        self.template_word = template_word
        self.dossier_sortie = dossier_sortie
        self.champs_attendus = champs_attendus
        self.dateDuJour = dateDuJour
        self.dossier_traite = dossier_traite  # Dossier où déplacer les fichiers traités
    
    def on_created(self, event):
        if event.is_directory:
            return
        
        if event.src_path.endswith(".xlsx") or event.src_path.endswith('.xls'):
            print(f"📂 Nouveau fichier détecté : {event.src_path}")
            
            # Extraire les données et remplir le template
            donnees_liste = extraire_donnees(event.src_path, self.champs_attendus)
            if donnees_liste:
                remplir_template(self.template_word, self.dossier_sortie, donnees_liste, self.dateDuJour)
                # Déplacer le fichier traité vers le répertoire 'dossier_traite'
                deplacer_fichier(event.src_path)
            else:
                print(f"❌ Erreur dans l'extraction des données du fichier : {event.src_path}")

# Fonction pour surveiller un répertoire
def surveiller_repertoire(dossier_a_surveiller, template_word, dossier_sortie, champs_attendus, dateDuJour, dossier_traite):
    event_handler = MoniteurDossier(dossier_a_surveiller, template_word, dossier_sortie, champs_attendus, dateDuJour, dossier_traite)
    observer = Observer()
    observer.schedule(event_handler, dossier_a_surveiller, recursive=False)
    observer.start()

    print(f"👀 Surveillance du répertoire {dossier_a_surveiller} pour de nouveaux fichiers...")

    try:
        while True:
            time.sleep(1)  # Attendre 1 seconde avant de vérifier à nouveau
    except KeyboardInterrupt:
        observer.stop()
        print("❌ Surveillance arrêtée.")
    observer.join()

# --- Partie principale du programme ---
if __name__ == "__main__":
    # Récupération des chemins dynamiques
    chemins = definir_chemins()

    dateDuJour = date.today().strftime("%d/%m/%Y")

    # Champs attendus dans le fichier Excel
    champs_attendus = ["Date Liq", "Matricule", "Identité Allocataire", "Identité Destinataire bailleur",
                       "Adresse Ligne 2", "Adresse Ligne 3", "Adresse Ligne 4",
                       "Adresse Ligne 5", "Adresse Ligne 6", "Adresse Ligne 7"]
    
    # Lancer la surveillance du répertoire
    surveiller_repertoire(
        chemins["dossier_a_surveiller"], 
        chemins["template_word"], 
        chemins["dossier_sortie"], 
        champs_attendus, 
        dateDuJour, 
        chemins["dossier_traite"]
    )
