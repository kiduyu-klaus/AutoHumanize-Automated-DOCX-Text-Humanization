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

def create_docx_from_text(text, preserve_formatting=True):
    """
    Create a DOCX document from text while preserving formatting.
    
    Args:
        text: str - The text to convert to DOCX
        preserve_formatting: bool - Whether to preserve paragraph structure
        
    Returns:
        BytesIO: Buffer containing the DOCX file
    """
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    if preserve_formatting:
        # Split text by paragraphs and add them to document
        paragraphs = text.split('\n')
        
        for para_text in paragraphs:
            if para_text.strip():  # Add non-empty paragraphs
                doc.add_paragraph(para_text)
            else:  # Preserve empty lines
                doc.add_paragraph()
    else:
        # Add as single paragraph
        doc.add_paragraph(text)
    
    # Save to BytesIO buffer
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    return buffer

def save_docx_to_output(text, filename):
    """
    Save humanized text to output folder as DOCX file.
    
    Args:
        text: str - The humanized text
        filename: str - Base filename (without extension)
    """
    try:
        # Create output directory if it doesn't exist
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        
        # Create DOCX
        doc = Document()
        
        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        
        # Add paragraphs
        paragraphs = text.split('\n')
        for para_text in paragraphs:
            if para_text.strip():
                doc.add_paragraph(para_text)
            else:
                doc.add_paragraph()
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(output_dir, f"{filename}_{timestamp}.docx")
        
        # Save document
        doc.save(output_path)
        
        return output_path
    except Exception as e:
        st.error(f"❌ Error saving DOCX: {str(e)}")
        return None

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
        help="Split long texts into chunks of this size"
    )
    
    st.markdown("---")
    st.markdown("### 📊 Statistics")
    if st.session_state.humanized_text:
        word_count = len(st.session_state.humanized_text.split())
        char_count = len(st.session_state.humanized_text)
        st.metric("Words", word_count)
        st.metric("Characters", char_count)
        
        # Show saved file info
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
        st.session_state.input_filename = "manual_input"
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
            
            # Store original filename (without extension)
            st.session_state.input_filename = os.path.splitext(uploaded_file.name)[0]
            
            # Clean up temp file
            os.unlink(tmp_file_path)
            
            if input_text:
                st.success(f"✅ File loaded successfully! ({len(input_text.split())} words)")
                with st.expander("Preview uploaded text"):
                    st.text_area(label="Preview", value=input_text, height=200, disabled=True)
            else:
                st.error("❌ Failed to read the file. Please check the file format.")

with col2:
    st.markdown("### ✨ Humanized Output")
    
    if st.session_state.humanized_text:
        output_container = st.container()
        with output_container:
            st.text_area(
                label="Humanized Text",
                value=st.session_state.humanized_text,
                height=400,
                key="output_text",
                disabled=True
            )
            
            # Download buttons row
            col_download1, col_download2 = st.columns(2)
            
            with col_download1:
                # Download as TXT
                txt_data = st.session_state.humanized_text.encode('utf-8')
                txt_filename = f"{st.session_state.input_filename}_humanized_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                st.download_button(
                    label="💾 Download as TXT",
                    data=txt_data,
                    file_name=txt_filename,
                    mime="text/plain",
                    use_container_width=True
                )
            
            with col_download2:
                # Download as DOCX
                docx_buffer = create_docx_from_text(st.session_state.humanized_text)
                docx_filename = f"{st.session_state.input_filename}_humanized_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
                st.download_button(
                    label="📄 Download as DOCX",
                    data=docx_buffer,
                    file_name=docx_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            
            # Copy button
            if st.button("📋 Copy to Clipboard", key="copy_btn", use_container_width=True):
                st.code(st.session_state.humanized_text, language=None)
                st.success("✅ Text copied! Use Ctrl+C / Cmd+C to copy from the box above.")
            
            # Show auto-saved file location
            if st.session_state.output_filename:
                st.info(f"💾 Auto-saved to: {st.session_state.output_filename}")
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
            st.session_state.output_filename = ""
            
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
                status_text.text("🔄 Processing text chunks...")
                progress_bar.progress(20)
                chunks = split_text_preserve_paragraphs_and_newlines(input_text, chunk_size)
                
                total_chunks = len(chunks)
                st.info(f"📊 Processing {total_chunks} chunk(s)...")
                
                humanized_chunks = []
                
                for i, chunk in enumerate(chunks):
                    status_text.text(f"🔄 Humanizing chunk {i+1}/{total_chunks}...")
                    progress = 20 + (60 * (i / total_chunks))
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
                progress_bar.progress(85)
                
                st.session_state.humanized_text = "\n\n".join(humanized_chunks)
                
                # Auto-save to DOCX in output folder
                status_text.text("💾 Saving to DOCX...")
                progress_bar.progress(95)
                
                base_filename = f"{st.session_state.input_filename}_humanized"
                output_path = save_docx_to_output(
                    st.session_state.humanized_text,
                    base_filename
                )
                
                if output_path:
                    st.session_state.output_filename = output_path
                
                progress_bar.progress(100)
                status_text.text("✅ Complete!")
                
                # Show success message
                st.success(f"🎉 Text humanized successfully!")
                if output_path:
                    st.success(f"💾 Saved to: {output_path}")
                
                time.sleep(1)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
            
            finally:
                # Quit driver after all processing is complete
                if st.session_state.driver:
                    st.session_state.driver.quit()
                    st.session_state.driver = None
                st.session_state.processing = False

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: white; padding: 1rem;">
        <p>Made with ❤️ using Streamlit | © 2024 AI Text Humanizer</p>
    </div>
""", unsafe_allow_html=True)