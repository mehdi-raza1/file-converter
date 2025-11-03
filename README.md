# 📄 Universal File Converter

A comprehensive web-based file conversion tool built with Streamlit that supports multiple file formats including PDF, Word, Excel, PowerPoint, and various image formats.

## 🚀 Features

### PDF Conversions
- **To PDF**: Word, Excel, PowerPoint, Images, Text
- **From PDF**: Word, Excel, PowerPoint, Text, Images
- **PDF Tools**: Merge, Split, Compress, Rotate, Extract/Remove Pages

### Image Conversions
- Format conversion (PNG, JPG, WebP, BMP, etc.)
- Resize and rotate images
- PDF to images and images to PDF

### Office Document Conversions
- Word ↔ Excel
- CSV ↔ Excel ↔ JSON
- Enhanced formatting preservation

## 🛠️ Installation

### Prerequisites
- Python 3.8 or higher
- Virtual environment (recommended)

### Quick Setup

1. **Clone or download the project**
   ```bash
   cd file-converter
   ```

2. **Create and activate virtual environment**
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Linux/macOS
   source venv/bin/activate
   ```

3. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Install system dependencies**

   **Windows:**
   - Download and install [LibreOffice](https://www.libreoffice.org/download/download/)
   - Download [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/)
   - Add Poppler to your PATH environment variable

   **Linux (Ubuntu/Debian):**
   ```bash
   sudo apt-get update
   sudo apt-get install poppler-utils libreoffice
   ```

   **macOS:**
   ```bash
   brew install poppler
   brew install --cask libreoffice
   ```

5. **Run the application**
   ```bash
   streamlit run app.py
   ```

### Alternative Installation Script

Run the automated installation script:
```bash
python install_dependencies.py
```

## 🎯 Usage

1. Open your web browser and navigate to the provided URL (usually `http://localhost:8501`)
2. Select the conversion type from the sidebar
3. Upload your file(s)
4. Click the conversion button
5. Download the converted file

## 📋 Supported Formats

### Input Formats
- **Documents**: PDF, DOCX, XLSX, PPTX, TXT
- **Images**: PNG, JPG, JPEG, WebP, BMP, TIFF
- **Data**: CSV, JSON, Excel

### Output Formats
- **Documents**: PDF, DOCX, XLSX, PPTX, TXT
- **Images**: PNG, JPG, WebP, BMP
- **Archives**: ZIP (for multiple files)

## 🔧 Configuration

### File Size Limits
- Maximum file size: 50MB per file
- PDF page limits: 100 pages for processing-intensive operations
- Image batch limit: 20 pages for PDF to images

### Logging
- Application logs are saved to `file_converter.log`
- Console output shows real-time status

## 🚀 Production Deployment

### Option 1: Streamlit Community Cloud (Recommended)
1. Push your code to GitHub
2. Connect your GitHub repo to [Streamlit Community Cloud](https://share.streamlit.io/)
3. The `packages.txt` file will automatically install system dependencies
4. Deploy with one click!

### Option 2: Docker Deployment
```bash
# Build and run with Docker
docker build -t file-converter .
docker run -p 8501:8501 file-converter

# Or use docker-compose
docker-compose up -d
```

### Option 3: Cloud Platforms
- **Heroku**: Use the included `Dockerfile`
- **Railway**: Direct GitHub integration
- **Google Cloud Run**: Container deployment
- **AWS ECS**: Enterprise container hosting

### Environment Variables
```bash
STREAMLIT_SERVER_HEADLESS=true
STREAMLIT_SERVER_ENABLE_CORS=false
STREAMLIT_SERVER_PORT=8501
```

## 🐛 Troubleshooting

### Local Development Issues

1. **PDF to Image conversion fails**
   - Ensure Poppler is installed and in PATH
   - On Windows, restart command prompt after PATH changes

2. **PowerPoint conversion issues**
   - Install LibreOffice for better compatibility
   - Some complex layouts may not convert perfectly

3. **Memory errors with large files**
   - Use smaller files (< 50MB)
   - Split large PDFs before processing

4. **Import errors**
   - Ensure all requirements are installed: `pip install -r requirements.txt`
   - Check Python version compatibility (3.8+)

### Performance Tips
- Use compressed images for better performance
- Split large documents before conversion
- Close other applications to free up memory

## 🔒 Security Notes

- Files are processed locally and not stored permanently
- Temporary files are automatically cleaned up
- No data is sent to external servers

## 📝 Development

### Project Structure