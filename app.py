import streamlit as st
import pandas as pd
import pdfkit
import io
import zipfile
import base64
import os

# Mobile-first page config
st.set_page_config(
    page_title="SGI Agent Monthly Report Generator",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for a beautiful, modern, minimalist UI (Dark Yellow & Dark Slate)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');

    /* Base Typography & Background */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif !important;
    }
    
    .stApp {
        background-color: #F8F9FA;
    }

    /* Hide default streamlit elements for minimalist look */
    header[data-testid="stHeader"] { background: transparent; }
    footer { visibility: hidden; }

    /* Layout Spacing */
    .main .block-container {
        padding-top: 3rem;
        padding-bottom: 2rem;
        max-width: 750px;
    }

    /* Headings & Text */
    h1 {
        color: #1A202C !important;
        font-weight: 700 !important;
        font-size: 2.2rem !important;
        text-align: center;
        letter-spacing: -0.5px;
        margin-bottom: 0.25rem !important;
    }
    
    p, label {
        color: #4A5568 !important;
        font-weight: 400 !important;
        font-size: 1.05rem !important;
        text-align: center;
    }

    /* File Uploader Card Styling */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF;
        border: 2px dashed #CBD5E0;
        border-radius: 16px;
        padding: 2rem 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #D4A017;
        box-shadow: 0 10px 15px -3px rgba(212, 160, 23, 0.15);
    }

    /* Primary Button (Dark Yellow/Gold) */
    .stButton>button {
        width: 100%;
        height: 3.5rem;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600;
        font-size: 1.1rem;
        background: linear-gradient(135deg, #E6AD00 0%, #C29100 100%);
        color: #FFFFFF !important;
        border: none;
        border-radius: 12px;
        box-shadow: 0 4px 14px 0 rgba(230, 173, 0, 0.35);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(230, 173, 0, 0.5);
        background: linear-gradient(135deg, #F0B400 0%, #CC9900 100%);
        color: #FFFFFF !important;
        border: none;
    }
    .stButton>button:active {
        transform: translateY(1px);
    }

    /* Download Button (Dark Slate to contrast the Gold) */
    [data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #1A202C 0%, #2D3748 100%) !important;
        box-shadow: 0 4px 14px 0 rgba(26, 32, 44, 0.35) !important;
    }
    [data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, #2D3748 0%, #4A5568 100%) !important;
        box-shadow: 0 6px 20px rgba(26, 32, 44, 0.5) !important;
    }

    /* Elegant Spinner */
    .stSpinner > div > div {
        border-color: #D4A017 !important;
        border-bottom-color: transparent !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("SGI Agent Monthly Report")
st.markdown("<p>Upload your Excel file to securely generate and download agent PDFs.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload .xlsx file", type=["xlsx", "xls"])

if uploaded_file is not None:
    if st.button("Generate Reports"):
        with st.spinner("Extracting Data & Generating PDFs..."):
            try:
                # 1. Read the Excel file
                xls = pd.ExcelFile(uploaded_file)
                
                # 2. Automatically target sheet containing "MUZAFFARPUR"
                target_sheet = None
                for sheet in xls.sheet_names:
                    if 'MUZAFFARPUR' in sheet.upper():
                        target_sheet = sheet
                        break
                
                if not target_sheet:
                    st.error("❌ No sheet containing 'MUZAFFARPUR' (case-insensitive) found.")
                    st.stop()
                    
                df = pd.read_excel(xls, sheet_name=target_sheet)
                
                # 3. Filter data
                if 'FINAL EXECUTIVE NAME' not in df.columns:
                    st.error("❌ Column 'FINAL EXECUTIVE NAME' not found in the sheet.")
                    st.stop()
                    
                df = df[df['FINAL EXECUTIVE NAME'] == 'KAUSHAL KISHORE']
                
                if df.empty:
                    st.warning("⚠️ No records found for 'KAUSHAL KISHORE'.")
                    st.stop()
                
                # 4. Map columns
                # Rename columns based on requirements
                rename_map = {}
                if 'S_MANF_DESC' in df.columns:
                    rename_map['S_MANF_DESC'] = 'S_MAKE_DESC'
                    
                df = df.rename(columns=rename_map)
                
                # Deduplicate columns by merging their data if they have the same name
                if df.columns.duplicated().any():
                    new_df = pd.DataFrame()
                    for col_name in df.columns.unique():
                        col_data = df[col_name]
                        if isinstance(col_data, pd.DataFrame):
                            # Column is duplicated, merge the text from all matching columns
                            new_df[col_name] = col_data.apply(
                                lambda row: ' '.join([str(x).strip() for x in row if pd.notna(x) and str(x).strip() not in ('nan', 'NaT', '')]), 
                                axis=1
                            )
                        else:
                            new_df[col_name] = col_data
                    df = new_df
                
                required_cols = [
                    'FINAL AGENT CODE', 'FINAL AGENT NAME', 'S_POLICYNO', 'SUB PRODUCT', 
                    'S_INSUREDNAME', 'ADDRESS', 'S_DOE', 'S_MAKE_DESC', 'S_VEH_REG_NO', 
                    'Previous Year IDV', 'Renewal IDV', 'TOAL DUE PREMIUM ON RN'
                ]
                
                # Strip whitespace from column names to prevent mismatch due to hidden spaces
                df.columns = df.columns.str.strip()
                
                # Ensure all required columns exist, fill with empty string if missing
                missing_cols = [col for col in required_cols if col not in df.columns]
                if missing_cols:
                    st.warning(f"⚠️ The following expected columns were NOT found in your Excel file and will be blank: {', '.join(missing_cols)}")
                    with st.expander("See all available columns in your Excel file"):
                        st.write(df.columns.tolist())
                
                for col in required_cols:
                    if col not in df.columns:
                        df[col] = ""
                        
                df = df[required_cols]
                df = df.fillna("")
                
                # Remove time from S_DOE
                if 'S_DOE' in df.columns:
                    df['S_DOE'] = df['S_DOE'].astype(str).str.split(' ').str[0]
                    # Also clean up string artifacts like "NaT" or "nan" if empty
                    df['S_DOE'] = df['S_DOE'].replace({'NaT': '', 'nan': ''})
                
                # 5. Get unique agents
                agents = df['FINAL AGENT NAME'].unique()
                
                # 6. Generate PDFs
                zip_buffer = io.BytesIO()
                
                header_colors = {
                    'FINAL AGENT CODE': '#FFC000',
                    'FINAL AGENT NAME': '#FFC000',
                    'S_POLICYNO': '#FFC000',
                    'SUB PRODUCT': '#FFC000',
                    'S_DOE': '#FFC000',
                    'S_MAKE_DESC': '#FFC000',
                    'S_VEH_REG_NO': '#FFC000',
                    'S_INSUREDNAME': '#FFFF00',
                    'ADDRESS': '#FFFF00',
                    'Previous Year IDV': '#FFFF00',
                    'Renewal IDV': '#FFFF00',
                    'TOAL DUE PREMIUM ON RN': '#FFFF00'
                }
                
                pdf_options = {
                    'orientation': 'Landscape',
                    'page-size': 'A4',
                    'margin-top': '5mm',
                    'margin-right': '5mm',
                    'margin-bottom': '5mm',
                    'margin-left': '5mm',
                    'encoding': "UTF-8",
                    'quiet': '',
                    'enable-local-file-access': ''
                }
                
                # PDFKit setup - works out of the box if wkhtmltopdf is in PATH (like on Streamlit Cloud with packages.txt)
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for agent in agents:
                        agent_df = df[df['FINAL AGENT NAME'] == agent]
                        
                        html_content = f"""
                        <!DOCTYPE html>
                        <html>
                        <head>
                        <meta charset="utf-8">
                        <style>
                            @page {{
                                size: A4 landscape;
                                margin: 5mm;
                            }}
                            body {{
                                font-family: Arial, Helvetica, sans-serif;
                                font-size: 14px;
                                font-weight: bold;
                                letter-spacing: 0.5px;
                            }}
                            table {{
                                width: 100%;
                                border-collapse: collapse;
                                table-layout: fixed;
                                word-wrap: break-word;
                            }}
                            th, td {{
                                border: 1px solid black;
                                height: 45px;
                                vertical-align: middle;
                                text-align: center;
                                overflow: hidden;
                                padding: 6px;
                                line-height: 1.4;
                            }}
                            th {{
                                font-weight: bold;
                            }}
                            .align-left {{
                                text-align: left !important;
                            }}
                            .italic-header {{
                                font-style: italic;
                            }}
                        </style>
                        </head>
                        <body>
                            <table>
                                <thead>
                                    <tr>
                        """
                        for col in required_cols:
                            color = header_colors.get(col, '#ffffff')
                            extra_class = "italic-header" if col in ['Previous Year IDV', 'Renewal IDV'] else ""
                            align_class = "align-left" if col in ['ADDRESS'] else ""
                            html_content += f'<th style="background-color: {color};" class="{extra_class} {align_class}">{col}</th>\n'
                            
                        html_content += """
                                    </tr>
                                </thead>
                                <tbody>
                        """
                        
                        for _, row in agent_df.iterrows():
                            html_content += "<tr>\n"
                            for col in required_cols:
                                val = str(row[col])
                                align_class = "align-left" if col in ['ADDRESS'] else ""
                                html_content += f'<td class="{align_class}">{val}</td>\n'
                            html_content += "</tr>\n"
                            
                        html_content += """
                                </tbody>
                            </table>
                        </body>
                        </html>
                        """
                        
                        # Generate PDF bytes
                        try:
                            # Automatically find wkhtmltopdf on Windows local test environment
                            pdfkit_config = None
                            local_path = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
                            if os.path.exists(local_path):
                                pdfkit_config = pdfkit.configuration(wkhtmltopdf=local_path)
                                
                            pdf_bytes = pdfkit.from_string(html_content, False, options=pdf_options, configuration=pdfkit_config)
                            safe_agent_name = str(agent).replace("/", "_").replace("\\", "_").strip()
                            if not safe_agent_name:
                                safe_agent_name = "UNKNOWN_AGENT"
                            zip_file.writestr(f"Report_{safe_agent_name}.pdf", pdf_bytes)
                        except Exception as e:
                            st.error(f"Error generating PDF for {agent}: {e}")
                            st.stop()
                            
                zip_buffer.seek(0)
                
                st.success("✅ Reports Generated Successfully!")
                st.download_button(
                    label="📥 Download All Reports (ZIP)",
                    data=zip_buffer,
                    file_name="Agent_Reports.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
