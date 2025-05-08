import streamlit as st
import os
import base64
import tempfile
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from docx import Document
from pdf2docx import Converter
from docx2pdf import convert
import io
import zipfile
import pdfplumber

# Configuration de la page
st.set_page_config(
    page_title="📁 Super Convertisseur de Fichiers",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Style CSS personnalisé
st.markdown("""
<style>
    .stApp {
        background-color: #060e2b;
    }
    .header {
        color: white;
        text-align: center;
        padding: 1rem;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
        padding: 10px 24px;
        font-weight: bold;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #45a049;
        transform: scale(1.02);
    }
    .stDownloadButton>button {
        background-color: #2196F3;
        color: white;
        border-radius: 5px;
    }
    .stDownloadButton>button:hover {
        background-color: #0b7dda;
    }
    .stSuccess {
        background-color: #e8f5e9;
        border-left: 5px solid #4CAF50;
        padding: 1rem;
    }
    .stInfo {
        background-color: #e3f2fd;
        border-left: 5px solid #2196F3;
        padding: 1rem;
    }
    .stWarning {
        background-color: #fff8e1;
        border-left: 5px solid #FFC107;
        padding: 1rem;
    }
    .stError {
        background-color: #ffebee;
        border-left: 5px solid #F44336;
        padding: 1rem;
    }
    .robot-img {
        display: block;
        margin-left: auto;
        margin-right: auto;
        width: 200px;
        margin-bottom: 1rem;
    }
    .tab-title {
        font-size: 1.2rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .footer {
        text-align: center;
        color: #7f8c8d;
        margin-top: 2rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
</style>
""", unsafe_allow_html=True)

def clean_column_names(columns):
    """Nettoie les noms de colonnes et gère les doublons"""
    cleaned = []
    seen = {}
    
    for i, col in enumerate(columns):
        # Nettoyage de base
        col = str(col).strip()
        if not col:
            col = f"Colonne_{i}"
        
        # Gestion des doublons
        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 1
        
        cleaned.append(col)
    
    return cleaned

def detect_best_separator(lines, sample_size=5):
    """Détecte le meilleur séparateur de colonnes"""
    separators = ['\t', '  ', '|', ';', ',']
    best_sep = None
    best_score = 0
    
    for sep in separators:
        col_counts = []
        for line in lines[:min(sample_size, len(lines))]:
            parts = line.split(sep)
            col_counts.append(len(parts))
        
        if len(set(col_counts)) == 1 and col_counts[0] > best_score:
            best_score = col_counts[0]
            best_sep = sep
    
    return best_sep

def looks_like_header(header_row, data_row):
    """Détermine si la première ligne ressemble à des en-têtes"""
    if len(header_row) != len(data_row):
        return False
    
    # Vérifie si les éléments de header_row semblent être des titres
    header_indicators = sum(
        1 for h, d in zip(header_row, data_row) 
        if (h.isupper() or ' ' not in h) and not d.replace('.','').isdigit()
    )
    
    return header_indicators / len(header_row) > 0.5

def main():
    # En-tête avec image centrale
    st.markdown("<h1 class='header'>🔄 Super Convertisseur de Fichiers</h1>", unsafe_allow_html=True)
    st.image("robot3.jpg", width=150, caption="Convertisseur Intelligent")
    st.markdown("""
    <div style='text-align: center; margin-bottom: 30px; font-size: 1.1rem;'>
        Transformez vos fichiers en un clin d'œil!<br>PDF ↔ Word ↔ Excel • Fusion • Fractionnement
    </div>
    """, unsafe_allow_html=True)

    # Création des onglets
    tabs = st.tabs([
        "📄 PDF → Word", 
        "📝 Word → PDF", 
        "📊 PDF → Excel", 
        "🔗 Fusion PDF", 
        "✂️ Fractionnement PDF"
    ])

    # Onglet 1: PDF vers Word
    with tabs[0]:
        st.markdown("<div class='tab-title'>📄 Conversion PDF vers Word</div>", unsafe_allow_html=True)
        st.markdown("Transformez vos fichiers PDF en documents Word modifiables")
        
        pdf_file = st.file_uploader("Choisissez un fichier PDF", type=["pdf"], 
                                  key="pdf_to_word", help="Format PDF uniquement")
        
        if pdf_file is not None:
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"📂 Fichier chargé: {pdf_file.name}")
            with col2:
                convert_button = st.button("✨ Convertir en Word", key="convert_pdf_to_word")
            
            if convert_button:
                with st.spinner("🔍 Conversion en cours... Un instant!"):
                    try:
                        # Sauvegarde temporaire du fichier PDF
                        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                        temp_pdf.write(pdf_file.read())
                        temp_pdf.close()
                        
                        # Chemin du fichier Word de sortie
                        output_docx = os.path.splitext(temp_pdf.name)[0] + '.docx'
                        
                        # Conversion PDF en Word
                        cv = Converter(temp_pdf.name)
                        cv.convert(output_docx)
                        cv.close()
                        
                        # Téléchargement du fichier converti
                        with open(output_docx, "rb") as file:
                            output_filename = os.path.splitext(pdf_file.name)[0] + '.docx'
                            docx_bytes = file.read()
                            st.download_button(
                                label="📥 Télécharger le fichier Word",
                                data=docx_bytes,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        
                        # Nettoyage des fichiers temporaires
                        os.unlink(temp_pdf.name)
                        os.unlink(output_docx)
                        
                        st.success("✅ Conversion terminée avec succès!")
                    except Exception as e:
                        st.error(f"❌ Une erreur est survenue: {str(e)}")

    # Onglet 2: Word vers PDF
    with tabs[1]:
        st.markdown("<div class='tab-title'>📝 Conversion Word vers PDF</div>", unsafe_allow_html=True)
        st.markdown("Convertissez vos documents Word en fichiers PDF professionnels")
        
        docx_file = st.file_uploader("Choisissez un fichier Word", type=["docx"], 
                                   key="word_to_pdf", help="Format .docx uniquement")
        
        if docx_file is not None:
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"📂 Fichier chargé: {docx_file.name}")
            with col2:
                convert_button = st.button("✨ Convertir en PDF", key="convert_word_to_pdf")
            
            if convert_button:
                with st.spinner("🔍 Conversion en cours... Un instant!"):
                    try:
                        # Sauvegarde temporaire du fichier Word
                        temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
                        temp_docx.write(docx_file.read())
                        temp_docx.close()
                        
                        # Chemin du fichier PDF de sortie
                        output_pdf = os.path.splitext(temp_docx.name)[0] + '.pdf'
                        
                        # Conversion Word en PDF
                        convert(temp_docx.name, output_pdf)
                        
                        # Téléchargement du fichier converti
                        with open(output_pdf, "rb") as file:
                            output_filename = os.path.splitext(docx_file.name)[0] + '.pdf'
                            pdf_bytes = file.read()
                            st.download_button(
                                label="📥 Télécharger le fichier PDF",
                                data=pdf_bytes,
                                file_name=output_filename,
                                mime="application/pdf"
                            )
                        
                        # Nettoyage des fichiers temporaires
                        os.unlink(temp_docx.name)
                        os.unlink(output_pdf)
                        
                        st.success("✅ Conversion terminée avec succès!")
                    except Exception as e:
                        st.error(f"❌ Une erreur est survenue: {str(e)}")

    # Onglet 3: PDF vers Excel
    with tabs[2]:
        st.markdown("<div class='tab-title'>📊 Conversion PDF vers Excel</div>", unsafe_allow_html=True)
        st.markdown("Extrayez les tableaux de vos PDF vers Excel en un clic")
        
        pdf_file = st.file_uploader("Choisissez un fichier PDF", type=["pdf"], 
                                  key="pdf_to_excel", help="Format PDF uniquement")
        
        if pdf_file is not None:
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"📂 Fichier chargé: {pdf_file.name}")
                page_number = st.number_input("Numéro de page à extraire", min_value=1, value=1)
                all_pages = st.checkbox("Convertir toutes les pages", value=False)
            with col2:
                # Définir une valeur par défaut pour space_threshold
                space_threshold = 2  # Valeur par défaut
                
                advanced_options = st.checkbox("⚙️ Options avancées")
                if advanced_options:
                    table_strategy = st.selectbox(
                        "Stratégie d'extraction",
                        ["Auto (recommandé)", "Tableaux uniquement", "Texte brut"]
                    )
                    space_threshold = st.slider("Seuil d'espacement pour les colonnes", 1, 10, space_threshold)
                
                convert_button = st.button("✨ Convertir en Excel", key="convert_pdf_to_excel")
            
            if convert_button:
                with st.spinner("🔍 Analyse du PDF en cours..."):
                    try:
                        # Sauvegarde temporaire du fichier PDF
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                            temp_pdf.write(pdf_file.read())
                            temp_path = temp_pdf.name
                        
                        # Initialisation
                        all_dfs = []
                        
                        with pdfplumber.open(temp_path) as pdf:
                            pages_to_process = pdf.pages if all_pages else [pdf.pages[page_number-1]]
                            
                            for page in pages_to_process:
                                try:
                                    # Paramètres d'extraction par défaut
                                    table_settings = {
                                        "vertical_strategy": "text", 
                                        "horizontal_strategy": "text",
                                        "explicit_vertical_lines": [],
                                        "explicit_horizontal_lines": [],
                                        "snap_tolerance": 3,
                                        "join_tolerance": 3,
                                        "edge_min_length": 3,
                                        "min_words_vertical": space_threshold
                                    }
                                    
                                    # Stratégie d'extraction
                                    if not advanced_options or table_strategy in ["Auto (recommandé)", "Tableaux uniquement"]:
                                        tables = page.extract_tables(table_settings)
                                        
                                        for table in tables:
                                            if len(table) > 1:  # Au moins une ligne d'en-tête + données
                                                # Nettoyage des noms de colonnes
                                                headers = clean_column_names(table[0])
                                                df = pd.DataFrame(table[1:], columns=headers)
                                                all_dfs.append(df)
                                    
                                    # Si mode auto ou texte brut et pas de tableaux trouvés
                                    if (not advanced_options or table_strategy in ["Auto (recommandé)", "Texte brut"]) and len(all_dfs) == 0:
                                        text = page.extract_text()
                                        if text:
                                            # Traitement intelligent du texte
                                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                                            
                                            # Détection des colonnes
                                            separator = detect_best_separator(lines)
                                            data = []
                                            
                                            for line in lines:
                                                if separator:
                                                    row = [cell.strip() for cell in line.split(separator)]
                                                else:
                                                    row = line.split(maxsplit=3)  # Limite le split pour les données simples
                                                data.append(row)
                                            
                                            if len(data) > 1:
                                                # Détection automatique des en-têtes
                                                if looks_like_header(data[0], data[1]):
                                                    headers = clean_column_names(data[0])
                                                    df = pd.DataFrame(data[1:], columns=headers)
                                                else:
                                                    df = pd.DataFrame(data)
                                                    df.columns = clean_column_names(df.columns)
                                                all_dfs.append(df)
                                
                                except Exception as e:
                                    st.warning(f"⚠️ Erreur sur la page {pages_to_process.index(page)+1}: {str(e)}")
                                    continue
                        
                        # Création du fichier Excel
                        if all_dfs:
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                for i, df in enumerate(all_dfs):
                                    sheet_name = f"Page_{(i//3)+1}_Table{(i%3)+1}" if all_pages else f"Données_{i+1}"
                                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                            
                            # Téléchargement
                            output_filename = f"{os.path.splitext(pdf_file.name)[0]}_converted.xlsx"
                            st.success(f"✅ Conversion réussie! {len(all_dfs)} tableaux extraits.")
                            
                            st.download_button(
                                label="📥 Télécharger le fichier Excel",
                                data=output.getvalue(),
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            # Aperçu
                            st.subheader("👀 Aperçu des données")
                            st.dataframe(all_dfs[0].head())
                        else:
                            st.error("❌ Aucune donnée exploitable n'a été trouvée dans le PDF.")
                    
                    except Exception as e:
                        st.error(f"❌ Erreur majeure lors de la conversion : {str(e)}")
                    finally:
                        # Nettoyage
                        if os.path.exists(temp_path):
                            os.unlink(temp_path)

    # Onglet 4: Fusion PDF
    with tabs[3]:
        st.markdown("<div class='tab-title'>🔗 Fusion de fichiers PDF</div>", unsafe_allow_html=True)
        st.markdown("Combinez plusieurs fichiers PDF en un seul document")
        
        uploaded_pdfs = st.file_uploader("Choisissez plusieurs fichiers PDF", 
                                       type=["pdf"], 
                                       accept_multiple_files=True, 
                                       key="merge_pdfs",
                                       help="Sélectionnez plusieurs fichiers PDF à fusionner")
        
        if uploaded_pdfs:
            pdf_names = [pdf.name for pdf in uploaded_pdfs]
            st.info(f"📂 Fichiers chargés ({len(uploaded_pdfs)}): {', '.join(pdf_names)}")
            
            if st.button("✨ Fusionner les PDF", key="merge_pdf_button"):
                with st.spinner("🔗 Fusion en cours..."):
                    try:
                        merger = PdfWriter()
                        
                        # Ajout de chaque PDF au merger
                        for pdf_file in uploaded_pdfs:
                            pdf_file.seek(0)
                            pdf = PdfReader(pdf_file)
                            for page in pdf.pages:
                                merger.add_page(page)
                        
                        # Création du PDF fusionné en mémoire
                        merged_pdf = io.BytesIO()
                        merger.write(merged_pdf)
                        merged_pdf.seek(0)
                        
                        # Téléchargement du fichier fusionné
                        st.download_button(
                            label="📥 Télécharger le PDF fusionné",
                            data=merged_pdf,
                            file_name="fichiers_fusionnes.pdf",
                            mime="application/pdf"
                        )
                        
                        st.success("✅ Fusion terminée avec succès!")
                    except Exception as e:
                        st.error(f"❌ Une erreur est survenue lors de la fusion: {str(e)}")

    # Onglet 5: Fractionnement PDF
    with tabs[4]:
        st.markdown("<div class='tab-title'>✂️ Fractionnement de fichier PDF</div>", unsafe_allow_html=True)
        st.markdown("Extrayez des pages spécifiques de vos fichiers PDF")
        
        pdf_file = st.file_uploader("Choisissez un fichier PDF", 
                                  type=["pdf"], 
                                  key="split_pdf",
                                  help="Sélectionnez un PDF à fractionner")
        
        if pdf_file is not None:
            # Lecture du PDF pour obtenir le nombre de pages
            pdf_file.seek(0)
            pdf = PdfReader(pdf_file)
            num_pages = len(pdf.pages)
            
            st.info(f"📂 Fichier chargé: {pdf_file.name} ({num_pages} pages)")
            
            col1, col2 = st.columns(2)
            with col1:
                split_option = st.radio(
                    "Option de fractionnement",
                    ["Toutes les pages individuellement", "Plage de pages spécifique"],
                    index=1
                )
            
            if split_option == "Plage de pages spécifique":
                with col2:
                    page_range = st.text_input(
                        "Entrez la plage de pages (ex: 1-3,5,7-9)",
                        placeholder="1-3,5,7-9",
                        help="Exemples: '1-3' (pages 1 à 3), '1,3,5' (pages 1,3,5), '1-3,5-7' (pages 1-3 et 5-7)"
                    )
            
            if st.button("✨ Fractionner le PDF", key="split_pdf_button"):
                with st.spinner("✂️ Fractionnement en cours..."):
                    try:
                        pdf_file.seek(0)
                        pdf = PdfReader(pdf_file)
                        
                        if split_option == "Toutes les pages individuellement":
                            # Création d'un ZIP pour contenir tous les PDF individuels
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                                for i in range(len(pdf.pages)):
                                    writer = PdfWriter()
                                    writer.add_page(pdf.pages[i])
                                    
                                    # Création du PDF individuel en mémoire
                                    output = io.BytesIO()
                                    writer.write(output)
                                    output.seek(0)
                                    
                                    # Ajout au ZIP
                                    zf.writestr(f"page_{i+1}.pdf", output.getvalue())
                            
                            # Téléchargement du ZIP
                            zip_buffer.seek(0)
                            st.download_button(
                                label="📥 Télécharger toutes les pages (ZIP)",
                                data=zip_buffer,
                                file_name="pages_individuelles.zip",
                                mime="application/zip"
                            )
                            
                        else:  # Plage de pages spécifique
                            if not page_range:
                                st.error("❌ Veuillez entrer une plage de pages valide.")
                                return
                            
                            # Parsing de la plage de pages
                            pages_to_extract = []
                            ranges = page_range.split(',')
                            for r in ranges:
                                if '-' in r:
                                    start, end = map(int, r.split('-'))
                                    pages_to_extract.extend(range(start, end + 1))
                                else:
                                    pages_to_extract.append(int(r))
                            
                            # Vérification que les pages sont dans la plage valide
                            pages_to_extract = [p for p in pages_to_extract if 1 <= p <= len(pdf.pages)]
                            
                            if not pages_to_extract:
                                st.error("❌ Aucune page valide spécifiée")
                                return
                            
                            # Création du PDF extrait
                            writer = PdfWriter()
                            for page_num in pages_to_extract:
                                writer.add_page(pdf.pages[page_num - 1])
                            
                            # Création du PDF en mémoire
                            output = io.BytesIO()
                            writer.write(output)
                            output.seek(0)
                            
                            # Téléchargement du PDF extrait
                            st.download_button(
                                label="📥 Télécharger les pages extraites",
                                data=output,
                                file_name=f"pages_extraites_{page_range}.pdf",
                                mime="application/pdf"
                            )
                        
                        st.success("✅ Fractionnement terminé avec succès!")
                    except Exception as e:
                        st.error(f"❌ Une erreur est survenue lors du fractionnement: {str(e)}")

    # Pied de page
    st.markdown("---")
    st.markdown("""
    <div class='footer'>
        🚀 Application développée avec Streamlit • © 2025 • Version 2.0 • PIAC
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()