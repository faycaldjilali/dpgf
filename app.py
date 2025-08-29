import streamlit as st
import pandas as pd
from openai import OpenAI
import tempfile
import os
#####
##
import time


# Set page configuration
st.set_page_config(
    page_title="Construction Cost Analyzer",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .result-box {
        background-color: #1f77b4;
        padding: 20px;
        border-radius: 10px;
        margin-top: 20px;
    }
    .stButton>button {
        background-color: #1f77b4;
        color: white;
        font-weight: bold;
        width: 100%;
    }
    .info-box {
        background-color: #1f77b4;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# App title and description
st.markdown('<h1 class="main-header">🏗️ Analyseur de Coûts de Construction </h1>', unsafe_allow_html=True)
st.markdown("""
<div class="info-box">
    <p>Cet outil analyse vos fichiers Excel de construction afin de fournir un détail complet des coûts des matériaux, incluant les quantités, les prix
    unitaires et les coûts totaux. Pour les prix manquants, il utilise les tarifs du marché français..</p>
    <p><strong>Instructions:</strong> Upload your Excel file, then click 'Analyze Costs'.</p>
</div>
""", unsafe_allow_html=True)

# Initialize session state
if 'analysis_result' not in st.session_state:
    st.session_state.analysis_result = None
if 'extracted_text' not in st.session_state:
    st.session_state.extracted_text = None

# Sidebar for API key input
with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input("OpenAI API Key", type="password", 
                           help="Enter your OpenAI API key to enable cost analysis")
    
    st.markdown("---")
    st.info("""
    Cet outil permet de :

    Extraire les données de votre fichier Excel

    Identifier les quantités de matériaux

    Appliquer les prix du marché lorsque nécessaire

    Calculer les coûts totaux (quantité × prix)

    Fournir un détail complet des coûts
    """)

# File upload section
st.markdown('<div class="sub-header">Upload Your Construction Excel File</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'], label_visibility="collapsed")

if uploaded_file is not None:
    # Process the uploaded file
    try:
        # Read the Excel file
        xls = pd.ExcelFile(uploaded_file)
        all_data = {}
        all_text = ""

        # Display file info
        st.success(f"File uploaded successfully: {uploaded_file.name}")
        st.write(f"Number of sheets: {len(xls.sheet_names)}")
        
        # Create tabs for different views
        tab1, tab2, tab3 = st.tabs(["Sheets Overview", "Data Preview", "Raw Text"])
        
        with tab1:
            st.subheader("Sheets in Excel File")
            for i, sheet_name in enumerate(xls.sheet_names):
                with st.expander(f"Sheet {i+1}: {sheet_name}"):
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    all_data[sheet_name] = df
                    st.dataframe(df.head(), use_container_width=True)
                    
                    # Add to text representation
                    all_text += f"\n\n--- Sheet: {sheet_name} ---\n\n"
                    for index, row in df.iterrows():
                        row_text = " | ".join([str(cell) for cell in row if pd.notna(cell)])
                        all_text += row_text + "\n"
        
        with tab2:
            st.subheader("Data Preview")
            sheet_to_preview = st.selectbox("Select sheet to preview", xls.sheet_names)
            if sheet_to_preview:
                df_preview = pd.read_excel(xls, sheet_name=sheet_to_preview)
                st.dataframe(df_preview, use_container_width=True)
        
        with tab3:
            st.subheader("Extracted Text")
            st.text_area("Text representation of Excel data", all_text[:5000] + "..." if len(all_text) > 5000 else all_text, 
                        height=300, disabled=True)
        
        # Store the extracted text in session state
        st.session_state.extracted_text = all_text
        
        # Analysis button
        st.markdown("---")
        st.markdown('<div class="sub-header">Cost Analysis</div>', unsafe_allow_html=True)
        
        if not api_key:
            st.warning("Please enter your OpenAI API key in the sidebar to enable cost analysis")
        else:
            if st.button("Analyze Costs", type="primary"):
                with st.spinner("Analyzing construction costs... This may take a moment."):
                    try:
                        # Set up OpenAI client
                        OpenAI.api_key = api_key
                        client = OpenAI(api_key=api_key)

                        # Create the prompt
                        detailed_prompt = """
                        Analyze the provided construction data and create a detailed material breakdown with the following:

                        1. For each material/component found in the data:
                           - Material description
                           - Quantity (extract from data if available)
                           - Unit of measurement
                           - Unit price (use French market prices if not provided in data)
                           - Total price (quantity × unit price)

                        2. Organize the analysis by categories:
                           - Structural materials
                           - Finishes materials
                           - Technical installation materials
                           - External works materials

                        3. Include:
                           - Summary table with quantities and prices
                           - Total project cost estimate
                           - Cost per square meter (if area data available)
                           - Notes on any assumptions made for missing prices

                        Excel Data:
                        """ + all_text[:12000]  # Limit text to avoid token limits

                        # Call OpenAI API
                        response = client.chat.completions.create(
                            model="gpt-4",
                            messages=[
                                {"role": "system", "content": "You are a construction cost expert with knowledge of French market prices for building materials. Calculate total prices using quantity × unit price."},
                                {"role": "user", "content": detailed_prompt}
                            ],
                            temperature=0.3,
                            max_tokens=2500
                        )

                        # Store the result
                        st.session_state.analysis_result = response.choices[0].message.content
                        
                    except Exception as e:
                        st.error(f"Error during analysis: {str(e)}")
        
        # Display results if available
        if st.session_state.analysis_result:
            st.markdown("---")
            st.markdown('<div class="sub-header">Analysis Results</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="result-box">{st.session_state.analysis_result}</div>', unsafe_allow_html=True)
            
            # Download button for results
            st.download_button(
                label="Download Analysis Results",
                data=st.session_state.analysis_result,
                file_name="construction_cost_analysis.txt",
                mime="text/plain"
            )
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")

else:
    st.info("Please upload an Excel file to get started.")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center'>
        <p>Construction Cost Analyzer Tool • Built with Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)
