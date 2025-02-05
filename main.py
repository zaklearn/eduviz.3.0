import streamlit as st
import pandas as pd
from pathlib import Path

# Import des analyses (tous les fichiers dans le même dossier)
from analyse1 import show_statistics
from analyse2 import show_zero_scores
from analyse3 import show_school_comparison
from analyse4 import show_language_effect
from analyse5 import show_correlation
from analyse6 import show_cronbach
from analyse7 import show_performance_school
from analyse8 import show_contextual_factors
from analyse9 import show_classroom_practices
from analyse10 import show_gender_effect
from analyse11 import show_ses_home_support
from analyse12 import show_international_comparison
from analyse13 import show_language_comparison

# Configuration de la page
st.set_page_config(
    page_title="Tableau de Bord Analyses",
    page_icon="📊",
    layout="wide"
)

# Dictionnaire des analyses disponibles
ANALYSES = {
    "📊 1. Statistiques Descriptives": show_statistics,
    "⚠️ 2. Scores Zéro": show_zero_scores,
    "🏫 3. Comparaison entre Écoles": show_school_comparison,
    "🗣 4. Effet de la Langue": show_language_effect,
    "📉 5. Corrélations": show_correlation,
    "📈 6. Fiabilité des Tests (Cronbach)": show_cronbach,
    "🏫 7. Performance par École": show_performance_school,
    "🏡 8. Facteurs Contextuels": show_contextual_factors,
    "👩‍🏫 9. Pratiques en Classe": show_classroom_practices,
    "🚻 10. Effet du Genre": show_gender_effect,
    "💰 11. Impact du SES & Soutien": show_ses_home_support,
    "🌍 12. Comparaison Internationale": show_international_comparison,
    "🇳🇱🇬🇧 13. Anglophones vs Néerlandophones": show_language_comparison,
}

def load_data():
    """Charge les données depuis le fichier Excel."""
    try:
        # Chargement du fichier de données
        uploaded_file = st.file_uploader(
            "Choisissez votre fichier Excel",
            type=["xlsx", "xls"],
            help="Sélectionnez le fichier contenant les données EGRA/EGMA"
        )
        
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
            return df
        else:
            st.warning("Veuillez uploader un fichier Excel pour commencer l'analyse.")
            return None
            
    except Exception as e:
        st.error(f"Erreur lors du chargement du fichier : {str(e)}")
        return None

def main():
    """Fonction principale de l'application."""
    
    # Sidebar pour la navigation
    st.sidebar.title("Tableau de Bord")
    st.sidebar.divider()
    
    # Chargement des données
    df = load_data()
    
    # Sélection de l'analyse
    analysis = st.sidebar.radio(
        "Sélectionnez une analyse :",
        options=list(ANALYSES.keys()),
        key="analysis_selector"
    )
    
    if df is not None:
        # Affichage du contenu principal
        st.title(analysis)
        st.divider()
        
        # Appel de la fonction d'analyse correspondante
        selected_analysis = ANALYSES[analysis]
        selected_analysis(df)

if __name__ == "__main__":
    main()