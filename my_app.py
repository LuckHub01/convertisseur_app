import streamlit as st
import os
import tempfile
import io
import zipfile
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from docx import Document
from pdf2docx import Converter
import pdfplumber
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

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
        background-color: #f0f2f6;
        color: #333;
    }
    .header {
        color: #2c3e50;
        text-align: center;
        padding: 1rem;
        font-size: 2.5rem;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 10px 24px;
        font-weight: bold;
        transition: all 0.3s;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #45a049;
        transform: scale(1.02);
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .stDownloadButton>button {
        background-color: #2196F3;
        color: white;
        border-radius: 5px;
        width: 100%;
    }
    .stDownloadButton>button:hover {
        background-color: #0b7dda;
    }
    .stFileUploader>div>div>div>div {
        color: #2c3e50;
    }
    .tab-content {
        padding: 1.5rem;
        background-color: white;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        margin-bottom: 1.5rem;
    }
    .footer {
        text-align: center;
        color: #7f8c8d;
        margin-top: 2rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
    .conversion-info {
        background-color: #e8f5e9;
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

def create_pdf_from_docx(docx_path, output_pdf_path):
    """Convertit un DOCX en PDF en utilisant reportlab"""
    doc = Document(docx_path)
    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    width, height = letter
    
    y_position = height - 40
    for para in doc.paragraphs:
        if y_position < 40:
            c.showPage()
            y_position = height - 40
        
        text = para.text
        lines = text.split('\n')
        for line in lines:
            c.setFont("Helvetica", 12)
            c.drawString(40, y_position, line)
            y_position -= 15
    
    c.save()

def clean_column_names(columns):
    """Nettoie les noms de colonnes"""
    cleaned = []
    seen = {}
    for i, col in enumerate(columns):
        col = str(col).strip()
        if not col:
            col = f"Colonne_{i+1}"
        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 1
        cleaned.append(col)
    return cleaned

def extract_tables_from_pdf(pdf_path, all_pages=False, page_number=1, strategy="auto", space_threshold=2):
    """Extrait les tables d'un PDF"""
    all_dfs = []
    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages if all_pages else [pdf.pages[page_number-1]]
        
        for page in pages:
            try:
                if strategy in ["auto", "tables"]:
                    tables = page.extract_tables({
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "explicit_vertical_lines": [],
                        "explicit_horizontal_lines": [],
                        "snap_tolerance": 3,
                        "join_tolerance": 3,
                        "edge_min_length": 3,
                        "min_words_vertical": space_threshold
                    })
                    
                    for table in tables:
                        if len(table) > 1:
                            headers = clean_column_names(table[0])
                            df = pd.DataFrame(table[1:], columns=headers)
                            all_dfs.append(df)
                
                if (strategy in ["auto", "text"] and not all_dfs):
                    text = page.extract_text()
                    if text:
                        lines = [line.strip() for line in text.split('\n') if line.strip()]
                        if len(lines) > 1:
                            headers = clean_column_names(lines[0].split())
                            data = [line.split() for line in lines[1:]]
                            df = pd.DataFrame(data, columns=headers)
                            all_dfs.append(df)
            
            except Exception as e:
                st.warning(f"Erreur sur la page {pages.index(page)+1}: {str(e)}")
    
    return all_dfs

def main():
    st.markdown("<h1 class='header'>🔄 Super Convertisseur de Fichiers</h1>", unsafe_allow_html=True)
    st.markdown("""
    <div style='text-align: center; margin-bottom: 2rem; font-size: 1.1rem; color: #555;'>
        Transformez vos fichiers entre différents formats en quelques clics
    </div>
    """, unsafe_allow_html=True)

    tabs = st.tabs([
        "📄 PDF → Word", 
        "📝 Word → PDF", 
        "📊 PDF → Excel", 
        "🔗 Fusion PDF", 
        "✂️ Fractionnement PDF"
    ])

    # PDF to Word
    with tabs[0]:
        st.markdown("<div class='tab-content'>", unsafe_allow_html=True)
        st.markdown("### Conversion PDF vers Word")
        st.markdown("Transformez vos fichiers PDF en documents Word modifiables")
        
        pdf_file = st.file_uploader("Téléversez un fichier PDF", type=["pdf"], key="pdf_to_word")
        
        if pdf_file:
            st.markdown(f"<div class='conversion-info'>Fichier sélectionné : <strong>{pdf_file.name}</strong></div>", unsafe_allow_html=True)
            
            if st.button("Convertir en Word", key="btn_pdf_to_word"):
                with st.spinner("Conversion en cours..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                            tmp_pdf.write(pdf_file.read())
                            tmp_pdf_path = tmp_pdf.name
                        
                        output_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx').name
                        
                        cv = Converter(tmp_pdf_path)
                        cv.convert(output_docx)
                        cv.close()
                        
                        with open(output_docx, "rb") as f:
                            st.download_button(
                                "Télécharger le document Word",
                                f,
                                file_name=f"{os.path.splitext(pdf_file.name)[0]}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        
                        st.success("Conversion réussie !")
                    except Exception as e:
                        st.error(f"Erreur : {str(e)}")
                    finally:
                        for path in [tmp_pdf_path, output_docx]:
                            if os.path.exists(path):
                                os.unlink(path)
        st.markdown("</div>", unsafe_allow_html=True)

    # Word to PDF
    with tabs[1]:
        st.markdown("<div class='tab-content'>", unsafe_allow_html=True)
        st.markdown("### Conversion Word vers PDF")
        st.markdown("Convertissez vos documents Word en fichiers PDF")
        
        docx_file = st.file_uploader("Téléversez un fichier Word", type=["docx"], key="word_to_pdf")
        
        if docx_file:
            st.markdown(f"<div class='conversion-info'>Fichier sélectionné : <strong>{docx_file.name}</strong></div>", unsafe_allow_html=True)
            
            if st.button("Convertir en PDF", key="btn_word_to_pdf"):
                with st.spinner("Conversion en cours..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_docx:
                            tmp_docx.write(docx_file.read())
                            tmp_docx_path = tmp_docx.name
                        
                        output_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf').name
                        
                        create_pdf_from_docx(tmp_docx_path, output_pdf)
                        
                        with open(output_pdf, "rb") as f:
                            st.download_button(
                                "Télécharger le PDF",
                                f,
                                file_name=f"{os.path.splitext(docx_file.name)[0]}.pdf",
                                mime="application/pdf"
                            )
                        
                        st.success("Conversion réussie !")
                    except Exception as e:
                        st.error(f"Erreur : {str(e)}")
                    finally:
                        for path in [tmp_docx_path, output_pdf]:
                            if os.path.exists(path):
                                os.unlink(path)
        st.markdown("</div>", unsafe_allow_html=True)

    # PDF to Excel
    with tabs[2]:
        st.markdown("<div class='tab-content'>", unsafe_allow_html=True)
        st.markdown("### Conversion PDF vers Excel")
        st.markdown("Extrayez les tableaux de vos PDF vers Excel")
        
        pdf_file = st.file_uploader("Téléversez un fichier PDF", type=["pdf"], key="pdf_to_excel")
        
        if pdf_file:
            st.markdown(f"<div class='conversion-info'>Fichier sélectionné : <strong>{pdf_file.name}</strong></div>", unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                page_num = st.number_input("Numéro de page", min_value=1, value=1)
                all_pages = st.checkbox("Toutes les pages")
            with col2:
                strategy = st.selectbox("Stratégie", ["auto", "tables", "text"])
                space_thresh = st.slider("Seuil d'espacement", 1, 10, 2)
            
            if st.button("Convertir en Excel", key="btn_pdf_to_excel"):
                with st.spinner("Extraction des données..."):
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
                            tmp_pdf.write(pdf_file.read())
                            tmp_pdf_path = tmp_pdf.name
                        
                        tables = extract_tables_from_pdf(
                            tmp_pdf_path, 
                            all_pages=all_pages,
                            page_number=page_num,
                            strategy=strategy,
                            space_threshold=space_thresh
                        )
                        
                        if tables:
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                for i, df in enumerate(tables):
                                    df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
                            
                            st.download_button(
                                "Télécharger le fichier Excel",
                                output.getvalue(),
                                file_name=f"{os.path.splitext(pdf_file.name)[0]}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            st.dataframe(tables[0].head())
                            st.success(f"{len(tables)} tableaux extraits avec succès !")
                        else:
                            st.warning("Aucun tableau détecté dans le PDF")
                    except Exception as e:
                        st.error(f"Erreur : {str(e)}")
                    finally:
                        if os.path.exists(tmp_pdf_path):
                            os.unlink(tmp_pdf_path)
        st.markdown("</div>", unsafe_allow_html=True)

    # Fusion PDF
    with tabs[3]:
        st.markdown("<div class='tab-content'>", unsafe_allow_html=True)
        st.markdown("### Fusion de fichiers PDF")
        st.markdown("Combinez plusieurs PDF en un seul document")
        
        pdf_files = st.file_uploader("Téléversez plusieurs fichiers PDF", 
                                   type=["pdf"], 
                                   accept_multiple_files=True,
                                   key="merge_pdfs")
        
        if pdf_files and len(pdf_files) > 1:
            st.markdown(f"<div class='conversion-info'>{len(pdf_files)} fichiers sélectionnés</div>", unsafe_allow_html=True)
            
            if st.button("Fusionner les PDF", key="btn_merge_pdfs"):
                with st.spinner("Fusion en cours..."):
                    try:
                        merger = PdfWriter()
                        for pdf_file in pdf_files:
                            pdf_file.seek(0)
                            pdf = PdfReader(pdf_file)
                            for page in pdf.pages:
                                merger.add_page(page)
                        
                        output = io.BytesIO()
                        merger.write(output)
                        output.seek(0)
                        
                        st.download_button(
                            "Télécharger le PDF fusionné",
                            output,
                            file_name="fusion.pdf",
                            mime="application/pdf"
                        )
                        st.success("Fusion réussie !")
                    except Exception as e:
                        st.error(f"Erreur : {str(e)}")
        st.markdown("</div>", unsafe_allow_html=True)

    # Fractionnement PDF
    with tabs[4]:
        st.markdown("<div class='tab-content'>", unsafe_allow_html=True)
        st.markdown("### Fractionnement de PDF")
        st.markdown("Extrayez des pages spécifiques de vos PDF")
        
        pdf_file = st.file_uploader("Téléversez un fichier PDF", type=["pdf"], key="split_pdf")
        
        if pdf_file:
            pdf = PdfReader(pdf_file)
            num_pages = len(pdf.pages)
            st.markdown(f"<div class='conversion-info'>Fichier sélectionné : <strong>{pdf_file.name}</strong> ({num_pages} pages)</div>", unsafe_allow_html=True)
            
            split_option = st.radio("Options", ["Toutes les pages", "Plage de pages"])
            
            if split_option == "Plage de pages":
                page_range = st.text_input("Entrez la plage (ex: 1-3,5,7-9)", help="Ex: '1-3' pour les pages 1 à 3, '1,3,5' pour les pages 1,3 et 5")
            
            if st.button("Fractionner le PDF", key="btn_split_pdf"):
                with st.spinner("Traitement en cours..."):
                    try:
                        if split_option == "Toutes les pages":
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                                for i in range(num_pages):
                                    writer = PdfWriter()
                                    writer.add_page(pdf.pages[i])
                                    page_bytes = io.BytesIO()
                                    writer.write(page_bytes)
                                    page_bytes.seek(0)
                                    zf.writestr(f"page_{i+1}.pdf", page_bytes.getvalue())
                            
                            st.download_button(
                                "Télécharger toutes les pages (ZIP)",
                                zip_buffer.getvalue(),
                                file_name="pages_separees.zip",
                                mime="application/zip"
                            )
                        else:
                            if not page_range:
                                st.error("Veuillez entrer une plage valide")
                            else:
                                pages = []
                                for part in page_range.split(','):
                                    if '-' in part:
                                        start, end = map(int, part.split('-'))
                                        pages.extend(range(start-1, end))
                                    else:
                                        pages.append(int(part)-1)
                                
                                writer = PdfWriter()
                                for page_num in sorted(set(pages)):
                                    if 0 <= page_num < num_pages:
                                        writer.add_page(pdf.pages[page_num])
                                
                                output = io.BytesIO()
                                writer.write(output)
                                output.seek(0)
                                
                                st.download_button(
                                    "Télécharger les pages sélectionnées",
                                    output,
                                    file_name="pages_selectionnees.pdf",
                                    mime="application/pdf"
                                )
                        
                        st.success("Fractionnement réussi !")
                    except Exception as e:
                        st.error(f"Erreur : {str(e)}")
        st.markdown("</div>", unsafe_allow_html=True)

    # Pied de page
    st.markdown("""
    <div class='footer'>
        Application développée avec Streamlit • © 2023 • Version 2.0
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()