# ✨ AutoHumanize App

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-2EAD33?style=for-the-badge&logo=playwright&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

**Transform AI-generated text into natural, human-like content with a single click.**

[🚀 Live Demo](https://autohumanizeapp.streamlit.app/) • [📖 Documentation](#documentation) • [🐛 Report Bug](https://github.com/kiduyu-klaus/AutoHumanize-Automated-DOCX-Text-Humanization/issues) • [✨ Request Feature](https://github.com/kiduyu-klaus/AutoHumanize-Automated-DOCX-Text-Humanization/issues)

</div>

---

## 🌟 Features

<table>
<tr>
<td width="50%">

### 📝 **Multi-Format Support**
- Direct text input via textarea
- DOCX file upload support
- Preserves document formatting
- Maintains paragraph structure

</td>
<td width="50%">

### 🎯 **Smart Processing**
- Intelligent text chunking
- Parallel processing capability
- Auto-save to DOCX format
- Real-time progress tracking

</td>
</tr>
<tr>
<td width="50%">

### 🎨 **Modern UI/UX**
- Beautiful gradient design
- Responsive layout
- Interactive progress indicators
- Dark mode optimized

</td>
<td width="50%">

### 💾 **Export Options**
- Download as TXT
- Download as DOCX
- Copy to clipboard
- Auto-save to output folder

</td>
</tr>
</table>

---

## 🚀 Quick Start

### 🌐 Online Version (Recommended)

Visit the live app at **[autohumanizeapp.streamlit.app](https://autohumanizeapp.streamlit.app/)** - no installation required!

### 💻 Local Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/kiduyu-klaus/AutoHumanize-Automated-DOCX-Text-Humanization.git
   cd autohumanize-app
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Playwright browsers**
   ```bash
   playwright install chromium
   ```

4. **Run the app**
   ```bash
   streamlit run text_humanizer_app.py
   ```

5. **Open your browser**
   ```
   Navigate to http://localhost:8501
   ```

---

## 📦 Requirements

```txt
streamlit>=1.28.0
playwright>=1.40.0
python-docx>=1.1.0
pyperclip>=1.8.2
undetected-chromedriver
selenium
lxml
Pillow
```

---

## 🎮 How to Use

### Method 1: Text Input

1. **Select "Type/Paste Text"** in the input section
2. **Paste or type** your AI-generated text
3. **Click "🚀 Humanize Text"** button
4. **Wait for processing** (progress bar shows status)
5. **Download or copy** the humanized text

### Method 2: DOCX Upload

1. **Select "Upload DOCX"** in the input section
2. **Upload your Word document**
3. **Preview the extracted text** (optional)
4. **Click "🚀 Humanize Text"** button
5. **Download** the humanized version as TXT or DOCX

### Advanced Settings

- **Chunk Size**: Adjust the word count per processing chunk (500-3000 words)
- **Playwright Installation**: Use the sidebar button if browser initialization fails

---

## 🏗️ Architecture

```
autohumanize-app/
├── 📄 text_humanizer_app.py      # Main Streamlit application
├── 📄 texttohuman.py              # Core humanization logic
├── 📄 requirements.txt            # Python dependencies
├── 📁 output/                     # Auto-saved DOCX files
└── 📄 README.md                   # This file
```

### Core Components

#### **text_humanizer_app.py**
- Streamlit UI interface
- File upload handling
- Progress tracking
- Export functionality

#### **texttohuman.py**
- Playwright browser automation
- Text chunking algorithm
- AI detection bypass
- DOCX processing utilities

---

## 🔧 Technical Details

### How It Works

1. **Text Input**: Accepts text or DOCX files
2. **Chunking**: Splits large texts into manageable chunks (preserves paragraphs)
3. **Processing**: Uses Playwright to interact with TextToHuman.com
4. **Optimization**: Automatically selects best alternatives with lowest AI detection
5. **Output**: Combines chunks and exports in multiple formats

### Browser Automation

The app uses **Playwright** for reliable, headless browser automation:
- Stealth mode to avoid detection
- Random user agent rotation
- Clipboard permission handling
- Automatic retry mechanisms

### Text Processing

- **Smart Chunking**: Preserves paragraph boundaries and line breaks
- **Context Preservation**: Maintains document structure
- **Parallel Processing**: Handles multiple chunks efficiently
- **Error Recovery**: Continues processing even if individual chunks fail

---

## 🎨 UI Features

### Design Highlights

- **Gradient Background**: Eye-catching purple gradient
- **Card-Based Layout**: Clean, organized interface
- **Real-Time Stats**: Word count, character count
- **Progress Indicators**: Visual feedback during processing
- **Responsive Design**: Works on desktop and mobile

### User Experience

- **One-Click Operation**: Simple humanization process
- **Multiple Export Options**: TXT, DOCX, clipboard
- **Auto-Save**: Automatic backup to output folder
- **Error Handling**: Clear error messages and recovery options

---

## 🐛 Troubleshooting

### Browser Installation Issues

If you see "Playwright browsers are not installed" error:

1. **Use the Manual Install Button**
   - Open the sidebar
   - Click "🎭 Install Playwright Browsers"
   - Wait for installation to complete

2. **Or install via terminal**
   ```bash
   playwright install chromium
   playwright install-deps chromium
   ```

### Common Issues

| Issue | Solution |
|-------|----------|
| **Text not processing** | Check internet connection, try smaller chunk size |
| **Upload fails** | Ensure DOCX file is valid and not corrupted |
| **Slow processing** | Large texts take time; consider splitting document |
| **Browser timeout** | Increase timeout in settings or try again |

---

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/AmazingFeature`)
3. **Commit your changes** (`git commit -m 'Add some AmazingFeature'`)
4. **Push to the branch** (`git push origin feature/AmazingFeature`)
5. **Open a Pull Request**

### Development Setup

```bash
# Clone your fork
git clone https://github.com/kiduyu-klaus/AutoHumanize-Automated-DOCX-Text-Humanization.git

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install Playwright
playwright install chromium

# Run in development mode
streamlit run text_humanizer_app.py --server.runOnSave true
```

---

## 📊 Performance

| Metric | Value |
|--------|-------|
| **Processing Speed** | ~2000 words/minute |
| **Max File Size** | Unlimited (chunked processing) |
| **Supported Formats** | TXT, DOCX |
| **Browser Support** | Chromium-based |
| **Concurrent Users** | Scalable on Streamlit Cloud |

---

## 🔒 Privacy & Security

- **No Data Storage**: Text is processed in real-time and not stored
- **Local Processing**: Browser automation runs in isolated environment
- **No Tracking**: No analytics or user tracking implemented
- **Open Source**: Full transparency of code

---

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- **[TextToHuman.com](https://texttohuman.com)** - AI humanization service
- **[Streamlit](https://streamlit.io)** - Web app framework
- **[Playwright](https://playwright.dev)** - Browser automation
- **[python-docx](https://python-docx.readthedocs.io)** - DOCX processing

---

## 📞 Support

- **Live App**: [autohumanizeapp.streamlit.app](https://autohumanizeapp.streamlit.app/)
- **Issues**: [GitHub Issues](https://github.com/kiduyu-klaus/AutoHumanize-Automated-DOCX-Text-Humanization/issues)
- **Email**: support@example.com

---

<div align="center">

**Made with ❤️ using Streamlit & Playwright**

⭐ Star this repo if you find it helpful!

[🔝 Back to Top](#-autohumanize-app)

</div>