# 📄 Appli Accusés de Réception

Application Streamlit pour générer automatiquement des accusés de réception personnalisés à partir d'un fichier Excel et d’un template Word.

## Fonctionnalités
- Interface Web avec Streamlit
- Remplissage dynamique de documents Word avec `docxtpl`
- Archivage automatique des fichiers source
- Prévisualisation texte du premier document généré

## Lancement local

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run code/app.py
