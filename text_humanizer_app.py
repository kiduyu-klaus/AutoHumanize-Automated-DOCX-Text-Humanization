import streamlit as st
import time
from finaltexttohuman import (
    get_huminizer_chrome_driver,
    get_texttohuman_humanizer_final,
    read_docx_with_spacing,
    split_text_preserve_paragraphs_and_newlines
)
import tempfile
import os
from docx import Document
from docx.shared import Pt
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="AI Text Humanizer",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern UI
st.markdown("""
    <style>
    /* Main container styling */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Content container */
    .main .block-container {
        padding: 2rem;
        max-width: 1200px;
    }
    
    /* Header styling */
    .header-container {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .header-title {
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .header-subtitle {
        color: #666;
        font-size: 1.1rem;
    }
    
    /* Card styling */
    .card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin-bottom: 1.5rem;
    }
    
    /* Button styling */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1rem;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
    }
    
    /* Text area styling */
    .stTextArea textarea {
        border-radius: 8px;
        border: 2px solid #e0e0e0;
        font-size: 1rem;
    }
    
    .stTextArea textarea:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 2px rgba(102, 126, 234, 0.2);
    }
    
    /* File uploader styling */
    .stFileUploader {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border: 2px dashed #667eea;
    }
    
    /* Success message */
    .success-message {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    
    /* Info box */
    .info-box {
        background: #e7f3ff;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2196F3;
        margin: 1rem 0;
    }
    
    /* Stats container */
    .stats-container {
        display: flex;
        justify-content: space-around;
        margin: 1rem 0;
    }
    
    .stat-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        flex: 1;
        margin: 0 0.5rem;
    }
    
    .stat-number {
        font-size: 2rem;
        font-weight: 700;
    }
    
    .stat-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'humanized_text' not in st.session_state:
    st.session_state.humanized_text = ""
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'driver' not in st.session_state:
    st.session_state.driver = None

# Header
st.markdown("""
    <div class="header-container">
        <h1 class="header-title">✨ AI Text Humanizer</h1>
        <p class="header-subtitle">Transform AI-generated text into natural, human-like content</p>
    </div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    
    chunk_size = st.slider(
        "Chunk Size (words)",
        min_value=500,
        max_value=3000,
        value=2000,
        step=100,
        help="Split long texts into chunks of this size"
    )
    
    st.markdown("---")
    st.markdown("### 📊 Statistics")
    if st.session_state.humanized_text:
        word_count = len(st.session_state.humanized_text.split())
        char_count = len(st.session_state.humanized_text)
        st.metric("Words", word_count)
        st.metric("Characters", char_count)
    else:
        st.info("Process text to see statistics")
    
    st.markdown("---")
    st.markdown("### ℹ️ About")
    st.markdown("""
    This tool uses advanced AI detection and humanization 
    to make your text sound more natural and human-written.
    
    **Features:**
    - Text & DOCX support
    - Preserves formatting
    - Smart chunk processing
    - Copy to clipboard
    """)

# Main content area
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📝 Input")
    
    # Tab selection for input method
    input_method = st.radio(
        "Choose input method:",
        ["Type/Paste Text", "Upload DOCX"],
        horizontal=True
    )
    
    input_text = ""
    
    if input_method == "Type/Paste Text":
        input_text = st.text_area(
            "Enter your text here:",
            height=400,
            placeholder="Paste or type your AI-generated text here...",
            key="text_input"
        )
    else:
        uploaded_file = st.file_uploader(
            "Upload a DOCX file",
            type=['docx'],
            help="Upload a Word document to humanize"
        )
        
        if uploaded_file is not None:
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # Read the DOCX file
            input_text = read_docx_with_spacing(tmp_file_path)
            
            # Clean up temp file
            os.unlink(tmp_file_path)
            
            if input_text:
                st.success(f"✅ File loaded successfully! ({len(input_text.split())} words)")
                with st.expander("Preview uploaded text"):
                    st.text_area("", value=input_text, height=200, disabled=True)
            else:
                st.error("❌ Failed to read the file. Please check the file format.")

with col2:
    st.markdown("### ✨ Humanized Output")
    
    if st.session_state.humanized_text:
        output_container = st.container()
        with output_container:
            st.text_area(
                "",
                value=st.session_state.humanized_text,
                height=400,
                key="output_text",
                disabled=True
            )
            
            # Copy button
            if st.button("📋 Copy to Clipboard", key="copy_btn"):
                st.code(st.session_state.humanized_text, language=None)
                st.success("✅ Text copied! Use Ctrl+C / Cmd+C to copy from the box above.")
    else:
        st.info("👈 Enter text or upload a document and click 'Humanize Text' to see results here.")

# Humanize button
st.markdown("---")
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])

with col_btn2:
    if st.button("🚀 Humanize Text", disabled=st.session_state.processing, key="humanize_btn"):
        if not input_text or not input_text.strip():
            st.error("⚠️ Please enter some text or upload a document first!")
        else:
            st.session_state.processing = True
            st.session_state.humanized_text = ""
            
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Initialize driver
                status_text.text("🔄 Initializing browser...")
                progress_bar.progress(10)
                driver = get_huminizer_chrome_driver()
                st.session_state.driver = driver
                
                # Split text into chunks
                status_text.text("📄 Processing text chunks...")
                progress_bar.progress(20)
                chunks = split_text_preserve_paragraphs_and_newlines(input_text, chunk_size)
                
                total_chunks = len(chunks)
                st.info(f"📊 Processing {total_chunks} chunk(s)...")
                
                humanized_chunks = []
                
                for i, chunk in enumerate(chunks):
                    status_text.text(f"🔄 Humanizing chunk {i+1}/{total_chunks}...")
                    progress = 20 + (70 * (i / total_chunks))
                    progress_bar.progress(int(progress))
                    
                    # Humanize chunk
                    humanized_chunk = get_texttohuman_humanizer_final(
                        chunk,
                        driver,
                        timeout=30
                    )
                    
                    if humanized_chunk:
                        humanized_chunks.append(humanized_chunk)
                    else:
                        st.warning(f"⚠️ Chunk {i+1} failed to process")
                
                # Combine results
                status_text.text("✅ Finalizing...")
                progress_bar.progress(90)
                
                st.session_state.humanized_text = "\n\n".join(humanized_chunks)
                
                progress_bar.progress(100)
                status_text.text("✅ Complete!")
                
                time.sleep(1)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                if st.session_state.driver:
                    st.session_state.driver.quit()
                    st.session_state.driver = None
            
            finally:
                st.session_state.processing = False

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: white; padding: 1rem;">
        <p>Made with ❤️ using Streamlit | © 2024 AI Text Humanizer</p>
    </div>
""", unsafe_allow_html=True)