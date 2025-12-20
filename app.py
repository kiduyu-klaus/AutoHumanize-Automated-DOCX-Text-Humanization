import streamlit as st
import time
import subprocess
import os
import sys
from pathlib import Path

# Function to check and install Playwright browsers
def ensure_playwright_installed():
    """Check if Playwright browsers are installed, install if not"""
    cache_dir = Path.home() / ".cache" / "ms-playwright"
    
    # Check if chromium is installed
    chromium_installed = False
    if cache_dir.exists():
        for item in cache_dir.iterdir():
            if item.is_dir() and "chromium" in item.name.lower():
                chromium_installed = True
                break
    
    if not chromium_installed:
        st.info("🎭 Installing Playwright browsers for the first time... This may take a few minutes.")
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            status_text.text("Installing Playwright Chromium browser...")
            progress_bar.progress(30)
            
            # Install chromium browser only
            result = subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                capture_output=True,
                text=True,
                timeout=300  # 5 minutes timeout
            )
            
            progress_bar.progress(70)
            
            if result.returncode != 0:
                st.error(f"❌ Failed to install Playwright browsers: {result.stderr}")
                st.info("💡 You may need to manually run: `playwright install chromium`")
                return False
            
            status_text.text("Installing system dependencies...")
            progress_bar.progress(90)
            
            # Try to install dependencies (may fail without sudo, but browser will still work in most cases)
            subprocess.run(
                [sys.executable, "-m", "playwright", "install-deps", "chromium"],
                capture_output=True,
                text=True,
                timeout=120
            )
            
            progress_bar.progress(100)
            status_text.text("✅ Playwright installation complete!")
            time.sleep(2)
            progress_bar.empty()
            status_text.empty()
            
            return True
            
        except subprocess.TimeoutExpired:
            st.error("❌ Installation timed out. Please try again.")
            return False
        except Exception as e:
            st.error(f"❌ Error during installation: {str(e)}")
            st.info("💡 You may need to manually run: `playwright install chromium`")
            return False
    
    return True

# Check and install Playwright before importing the module
if 'playwright_checked' not in st.session_state:
    st.session_state.playwright_checked = False

if not st.session_state.playwright_checked:
    ensure_playwright_installed()
    st.session_state.playwright_checked = True

# Now import the texttohuman module
from texttohuman import (
    get_huminizer_chrome_driver,
    get_texttohuman_humanizer_final,
    read_docx_with_spacing,
    split_text_preserve_paragraphs_and_newlines,
    read_docx_and_humanize
)
import tempfile
from docx import Document
from docx.shared import Pt
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="autohumanize-app : AI Text Humanizer",
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
    
    /* Download button styling */
    .stDownloadButton>button {
        width: 100%;
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1rem;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    
    .stDownloadButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
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
if 'output_filename' not in st.session_state:
    st.session_state.output_filename = ""
if 'input_filename' not in st.session_state:
    st.session_state.input_filename = ""
if 'docx_buffer' not in st.session_state:
    st.session_state.docx_buffer = None
if 'input_method' not in st.session_state:
    st.session_state.input_method = "Type/Paste Text"

def create_docx_from_text(text):
    """
    Create a DOCX document from text, preserving paragraph structure (newlines).
    
    Args:
        text: str - The text to convert to DOCX
        
    Returns:
        BytesIO: Buffer containing the DOCX file
    """
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # Split text by paragraphs and add them to document
    paragraphs = text.split('\n')
    
    for para_text in paragraphs:
        doc.add_paragraph(para_text)
    
    # Save to BytesIO buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

def save_docx_to_output(buffer, filename):
    """
    Save DOCX buffer to output folder.
    
    Args:
        buffer: BytesIO - The DOCX file buffer
        filename: str - Base filename (without extension)
    """
    try:
        # Create output directory if it doesn't exist
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(output_dir, f"{filename}_humanized_{timestamp}.docx")
        
        # Save document
        with open(output_path, "wb") as f:
            f.write(buffer.getbuffer())
        
        return output_path
    except Exception as e:
        st.error(f"❌ Error saving DOCX: {str(e)}")
        return None

def process_text_chunks(text, driver, chunk_size):
    """
    Splits text into chunks, humanizes each chunk, and returns the combined result.
    """
    final_humanized_text = ""
    chunks = split_text_preserve_paragraphs_and_newlines(text, chunk_size)
    
    st.info(f"Text split into {len(chunks)} chunks for processing.")
    
    for i, chunk in enumerate(chunks, 1):
        st.info(f"Processing Chunk {i}/{len(chunks)}...")
        
        try:
            result = get_texttohuman_humanizer_final(chunk, driver, save_debug=False)
            if result:
                final_humanized_text += result + "\n"
            else:
                st.warning(f"Chunk {i} returned no result. Skipping.")
                
        except Exception as e:
            st.error(f"Error processing chunk {i}: {e}")
            continue
            
    return final_humanized_text.strip()

def handle_process_click():
    """
    Handles the main processing logic when the button is clicked.
    """
    st.session_state.processing = True
    st.session_state.humanized_text = ""
    st.session_state.docx_buffer = None
    
    if st.session_state.input_method == "Upload DOCX":
        if 'uploaded_file_path' not in st.session_state or not st.session_state.uploaded_file_path:
            st.error("Please upload a DOCX file first.")
            st.session_state.processing = False
            return
        
        with st.spinner("Initializing browser and humanizing DOCX... This may take a moment."):
            try:
                from texttohuman import PlaywrightHumanizer
                with PlaywrightHumanizer(headless=True, debug=False) as driver:
                    st.session_state.docx_buffer = read_docx_and_humanize(
                        st.session_state.uploaded_file_path, 
                        driver, 
                        chunk_size=st.session_state.chunk_size
                    )
                
                if st.session_state.docx_buffer:
                    st.success("✅ DOCX Humanization Complete!")
                    
                    humanized_doc = Document(st.session_state.docx_buffer)
                    humanized_text = "\n".join([p.text for p in humanized_doc.paragraphs])
                    st.session_state.humanized_text = humanized_text
                    
                    st.session_state.output_filename = save_docx_to_output(
                        st.session_state.docx_buffer, 
                        st.session_state.input_filename
                    )
                else:
                    st.error("❌ DOCX Humanization Failed. Check logs for details.")
                    
            except Exception as e:
                st.error(f"An unexpected error occurred during DOCX processing: {e}")
            finally:
                if 'uploaded_file_path' in st.session_state and os.path.exists(st.session_state.uploaded_file_path):
                    os.remove(st.session_state.uploaded_file_path)
                    del st.session_state.uploaded_file_path
                st.session_state.processing = False
                st.rerun()
                
    else:
        input_text = st.session_state.text_input
        if not input_text.strip():
            st.error("Please enter some text to humanize.")
            st.session_state.processing = False
            return
            
        with st.spinner("Initializing browser and humanizing text... This may take a moment."):
            try:
                from texttohuman import PlaywrightHumanizer
                with PlaywrightHumanizer(headless=True, debug=False) as driver:
                    humanized_text = process_text_chunks(
                        input_text, 
                        driver, 
                        st.session_state.chunk_size
                    )
                
                if humanized_text:
                    st.session_state.humanized_text = humanized_text
                    st.success("✅ Text Humanization Complete!")
                    
                    st.session_state.docx_buffer = create_docx_from_text(humanized_text)
                    
                    st.session_state.output_filename = save_docx_to_output(
                        st.session_state.docx_buffer, 
                        st.session_state.input_filename
                    )
                else:
                    st.error("❌ Text Humanization Failed. Check logs for details.")
                    
            except Exception as e:
                st.error(f"An unexpected error occurred during text processing: {e}")
            finally:
                st.session_state.processing = False
                st.rerun()

# Header
st.markdown("""
    <div class="header-container">
        <h1 class="header-title">✨ autohumanize-app : AI Text Humanizer</h1>
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
        key="chunk_size",
        help="Split long texts into chunks of this size"
    )
    
    st.markdown("---")
    st.markdown("### 📊 Statistics")
    if st.session_state.humanized_text:
        word_count = len(st.session_state.humanized_text.split())
        char_count = len(st.session_state.humanized_text)
        st.metric("Words", word_count)
        st.metric("Characters", char_count)
        
        if st.session_state.output_filename:
            st.markdown("---")
            st.markdown("### 💾 Saved File")
            st.success(f"📄 {os.path.basename(st.session_state.output_filename)}")
    else:
        st.info("Process text to see statistics")
    
    st.markdown("---")
    st.markdown("### ℹ️ About")
    st.markdown("""
    This tool uses advanced AI detection and humanization 
    to make your text sound more natural and human-written.
    
    **Features:**
    - Text & DOCX support
    - Auto-save to DOCX
    - Preserves formatting
    - Smart chunk processing
    - Copy to clipboard
    - Download as TXT/DOCX
    """)

# Main content area
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📝 Input")
    
    input_method = st.radio(
        "Choose input method:",
        ["Type/Paste Text", "Upload DOCX"],
        horizontal=True,
        key="input_method"
    )
    
    input_text = ""
    
    if st.session_state.input_method == "Type/Paste Text":
        input_text = st.text_area(
            "Enter your text here:",
            height=400,
            placeholder="Paste or type your AI-generated text here...",
            key="text_input"
        )
        st.session_state.input_filename = "manual_input"
        if 'uploaded_file_path' in st.session_state:
            del st.session_state.uploaded_file_path
    else:
        uploaded_file = st.file_uploader(
            "Upload a DOCX file",
            type=['docx'],
            help="Upload a Word document to humanize"
        )
        
        if uploaded_file is not None:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                st.session_state.uploaded_file_path = tmp_file.name
            
            st.session_state.input_filename = os.path.splitext(uploaded_file.name)[0]
            
            st.success(f"File uploaded: {uploaded_file.name}")
        else:
            if 'uploaded_file_path' in st.session_state:
                os.remove(st.session_state.uploaded_file_path)
                del st.session_state.uploaded_file_path
            st.session_state.input_filename = ""
            
    st.button(
        "✨ Humanize Text",
        on_click=handle_process_click,
        disabled=st.session_state.processing,
        use_container_width=True
    )

with col2:
    st.markdown("### 📄 Output")
    
    st.text_area(
        "Humanized Text",
        st.session_state.humanized_text,
        height=400,
        key="output_text",
        disabled=True
    )
    
    col_dl1, col_dl2 = st.columns(2)
    
    if st.session_state.humanized_text:
        if st.session_state.docx_buffer:
            docx_filename = f"{st.session_state.input_filename}_humanized.docx"
            col_dl1.download_button(
                label="⬇️ Download DOCX",
                data=st.session_state.docx_buffer,
                file_name=docx_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            col_dl1.info("DOCX not available for download.")
            
        txt_filename = f"{st.session_state.input_filename}_humanized.txt"
        col_dl2.download_button(
            label="⬇️ Download TXT",
            data=st.session_state.humanized_text.encode('utf-8'),
            file_name=txt_filename,
            mime="text/plain",
            use_container_width=True
        )
    else:
        col_dl1.info("Output will appear here.")
        col_dl2.info("Output will appear here.")
        
st.markdown("---")
st.markdown("""
    <div class="info-box">
        **Note on Formatting:** For DOCX uploads, the original document structure (headings, tables, bold text) is preserved by editing the document in place. For text input, newlines are treated as paragraph breaks to maintain structure.
    </div>
""", unsafe_allow_html=True)