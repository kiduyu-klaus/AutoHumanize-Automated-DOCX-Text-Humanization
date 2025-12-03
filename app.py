import streamlit as st
from texttohuman import process_docx_file, get_texttohuman_humanizer, split_text_by_words
import os
import tempfile
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="TextToHuman - AI Text Humanizer",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern UI
st.markdown("""
    <style>
    /* Main theme colors */
    :root {
        --primary-color: #10b981;
        --secondary-color: #3b82f6;
        --background-color: #0f172a;
        --card-background: #1e293b;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Main container styling */
    .main {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
    }
    
    /* Custom card styling */
    .custom-card {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        padding: 2rem;
        border-radius: 1rem;
        border: 1px solid rgba(59, 130, 246, 0.2);
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
        margin-bottom: 1.5rem;
    }
    
    /* Header styling */
    .header-container {
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 2rem;
    }
    
    .main-title {
        font-size: 3.5rem;
        font-weight: 800;
        background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        font-size: 1.3rem;
        color: #94a3b8;
        font-weight: 400;
    }
    
    /* Stats card */
    .stat-card {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        padding: 1.5rem;
        border-radius: 0.75rem;
        border: 1px solid rgba(16, 185, 129, 0.2);
        text-align: center;
        transition: transform 0.3s ease;
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
    }
    
    .stat-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #10b981;
        margin-bottom: 0.5rem;
    }
    
    .stat-label {
        font-size: 0.9rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    /* Progress styling */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #10b981 0%, #3b82f6 100%);
    }
    
    /* Button styling */
    .stButton>button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: 600;
        border-radius: 0.5rem;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 25px rgba(16, 185, 129, 0.4);
    }
    
    /* File uploader styling */
    .uploadedFile {
        border: 2px dashed #10b981;
        border-radius: 0.75rem;
        padding: 1rem;
    }
    
    /* Success/Error messages */
    .stSuccess {
        background: rgba(16, 185, 129, 0.1);
        border: 1px solid #10b981;
        border-radius: 0.5rem;
    }
    
    .stError {
        background: rgba(239, 68, 68, 0.1);
        border: 1px solid #ef4444;
        border-radius: 0.5rem;
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
    }
    
    /* Info box */
    .info-box {
        background: rgba(59, 130, 246, 0.1);
        border-left: 4px solid #3b82f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processing_complete' not in st.session_state:
    st.session_state.processing_complete = False
if 'humanized_text' not in st.session_state:
    st.session_state.humanized_text = None
if 'stats' not in st.session_state:
    st.session_state.stats = {
        'total_chunks': 0,
        'successful': 0,
        'failed': 0,
        'total_words': 0
    }

# Header
st.markdown("""
    <div class="header-container">
        <h1 class="main-title">✨ TextToHuman</h1>
        <p class="subtitle">Transform AI-Generated Text into Natural, Human-Like Content</p>
    </div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    
    mode = st.radio(
        "Select Mode",
        ["📄 DOCX Upload", "✍️ Text Input"],
        help="Choose how you want to input your text"
    )
    
    st.markdown("---")
    
    max_words = st.slider(
        "Words per Chunk",
        min_value=500,
        max_value=2000,
        value=1200,
        step=100,
        help="Split text into chunks of this size"
    )
    
    processing_timeout = st.slider(
        "Processing Timeout (seconds)",
        min_value=30,
        max_value=120,
        value=60,
        step=10,
        help="Maximum time to wait for each chunk"
    )
    
    st.markdown("---")
    
    st.markdown("""
        <div class="info-box">
            <strong>💡 Tips:</strong><br>
            • Larger chunks = fewer API calls<br>
            • Increase timeout for long texts<br>
            • Results are saved automatically
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### 📊 Session Stats")
    if st.session_state.stats['total_chunks'] > 0:
        st.metric("Total Chunks", st.session_state.stats['total_chunks'])
        st.metric("Successful", st.session_state.stats['successful'])
        st.metric("Failed", st.session_state.stats['failed'])
        st.metric("Total Words", st.session_state.stats['total_words'])

# Main content area
if mode == "📄 DOCX Upload":
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("### 📤 Upload Your Document")
    
    uploaded_file = st.file_uploader(
        "Choose a DOCX file",
        type=['docx'],
        help="Upload your AI-generated document to humanize"
    )
    
    if uploaded_file is not None:
        # Display file info
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-value">📄</div>
                    <div class="stat-label">{uploaded_file.name}</div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            file_size = uploaded_file.size / 1024
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-value">{file_size:.1f}</div>
                    <div class="stat-label">KB</div>
                </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-value">{max_words}</div>
                    <div class="stat-label">Words/Chunk</div>
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Process button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Humanize Document", use_container_width=True):
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name
                
                try:
                    with st.spinner('🔄 Processing your document...'):
                        # Create progress container
                        progress_container = st.container()
                        
                        with progress_container:
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            # Import and read file
                            from docx import Document
                            doc = Document(tmp_path)
                            full_text = '\n'.join([p.text for p in doc.paragraphs if p.text.strip()])
                            chunks = split_text_by_words(full_text, max_words)
                            
                            total_chunks = len(chunks)
                            humanized_chunks = []
                            failed_chunks = []
                            
                            for i, chunk in enumerate(chunks):
                                status_text.markdown(f"**Processing chunk {i+1}/{total_chunks}** ({len(chunk.split())} words)")
                                progress_bar.progress((i) / total_chunks)
                                
                                humanized = get_texttohuman_humanizer(chunk, processing_timeout=processing_timeout)
                                
                                if humanized:
                                    humanized_chunks.append(humanized)
                                else:
                                    humanized_chunks.append(chunk)
                                    failed_chunks.append(i+1)
                            
                            progress_bar.progress(1.0)
                            status_text.markdown("**✅ Processing complete!**")
                            
                            # Store results
                            st.session_state.humanized_text = '\n\n'.join(humanized_chunks)
                            st.session_state.processing_complete = True
                            st.session_state.stats = {
                                'total_chunks': total_chunks,
                                'successful': total_chunks - len(failed_chunks),
                                'failed': len(failed_chunks),
                                'total_words': len(full_text.split())
                            }
                            
                            # Show success message
                            if len(failed_chunks) == 0:
                                st.success(f"🎉 Successfully humanized all {total_chunks} chunks!")
                            else:
                                st.warning(f"⚠️ Completed with {len(failed_chunks)} failed chunks: {failed_chunks}")
                    
                except Exception as e:
                    st.error(f"❌ Error processing file: {str(e)}")
                finally:
                    # Clean up temp file
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)

else:  # Text Input Mode
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("### ✍️ Enter Your Text")
    
    input_text = st.text_area(
        "Paste your AI-generated text here",
        height=300,
        placeholder="Enter the text you want to humanize...",
        help="Paste your AI-generated content here"
    )
    
    if input_text:
        word_count = len(input_text.split())
        st.info(f"📊 Word count: **{word_count}** words | Will be split into **{(word_count // max_words) + 1}** chunks")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Process button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Humanize Text", disabled=not input_text, use_container_width=True):
            try:
                with st.spinner('🔄 Processing your text...'):
                    chunks = split_text_by_words(input_text, max_words)
                    total_chunks = len(chunks)
                    
                    progress_container = st.container()
                    
                    with progress_container:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        humanized_chunks = []
                        failed_chunks = []
                        
                        for i, chunk in enumerate(chunks):
                            status_text.markdown(f"**Processing chunk {i+1}/{total_chunks}** ({len(chunk.split())} words)")
                            progress_bar.progress((i) / total_chunks)
                            
                            humanized = get_texttohuman_humanizer(chunk, processing_timeout=processing_timeout)
                            
                            if humanized:
                                humanized_chunks.append(humanized)
                            else:
                                humanized_chunks.append(chunk)
                                failed_chunks.append(i+1)
                        
                        progress_bar.progress(1.0)
                        status_text.markdown("**✅ Processing complete!**")
                        
                        # Store results
                        st.session_state.humanized_text = '\n\n'.join(humanized_chunks)
                        st.session_state.processing_complete = True
                        st.session_state.stats = {
                            'total_chunks': total_chunks,
                            'successful': total_chunks - len(failed_chunks),
                            'failed': len(failed_chunks),
                            'total_words': len(input_text.split())
                        }
                        
                        # Show success message
                        if len(failed_chunks) == 0:
                            st.success(f"🎉 Successfully humanized all {total_chunks} chunks!")
                        else:
                            st.warning(f"⚠️ Completed with {len(failed_chunks)} failed chunks: {failed_chunks}")
                
            except Exception as e:
                st.error(f"❌ Error processing text: {str(e)}")

# Results section
if st.session_state.processing_complete and st.session_state.humanized_text:
    st.markdown("---")
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("### 📝 Humanized Results")
    
    # Display results in tabs
    tab1, tab2 = st.tabs(["📄 Preview", "💾 Download"])
    
    with tab1:
        st.text_area(
            "Humanized Text",
            value=st.session_state.humanized_text,
            height=400,
            help="Your humanized text appears here"
        )
        
        # Statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Chunks", st.session_state.stats['total_chunks'])
        with col2:
            st.metric("Successful", st.session_state.stats['successful'])
        with col3:
            st.metric("Failed", st.session_state.stats['failed'])
        with col4:
            st.metric("Total Words", st.session_state.stats['total_words'])
    
    with tab2:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"humanized_text_{timestamp}.txt"
        
        st.download_button(
            label="⬇️ Download Humanized Text",
            data=st.session_state.humanized_text,
            file_name=filename,
            mime="text/plain",
            use_container_width=True
        )
        
        st.info(f"💡 File will be saved as: **{filename}**")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Reset button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🔄 Process Another Document", use_container_width=True):
            st.session_state.processing_complete = False
            st.session_state.humanized_text = None
            st.rerun()

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #94a3b8; padding: 2rem 0;">
        <p>✨ <strong>TextToHuman</strong> - Powered by Advanced AI Humanization Technology</p>
        <p style="font-size: 0.9rem; margin-top: 0.5rem;">
            Made with ❤️ | Transform AI text into natural, human-like content
        </p>
    </div>
""", unsafe_allow_html=True)