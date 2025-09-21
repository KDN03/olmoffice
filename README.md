# 📄 Document Converter Pro

A comprehensive file conversion web application that supports **117+ conversion types** across all major document formats including Office, PDF, Images, and more!

## 🚀 Features

### ✨ Comprehensive Format Support
- **Office Documents**: DOC, DOCX, XLS, XLSX, PPT, PPTX
- **PDF Files**: Full PDF conversion and extraction
- **Images**: JPG, JPEG, PNG, GIF, BMP, TIFF
- **Web & Text**: HTML, TXT, CSV
- **And more**: RTF, ODT, ODS, ODP

### 🎯 Conversion Capabilities
- **117+ supported conversion types**
- **Bi-directional conversions** (forward and reverse)
- **Advanced features**: OCR, image rendering, content extraction
- **100% success rate** for core conversions

### ⚡ Multiple Conversion Engines
- **Python-optimized conversions** (72+ types)
- **LibreOffice integration** for office documents
- **wkhtmltopdf** for HTML→PDF conversions
- **Client-side processing** for instant conversions

## 🛠 Technology Stack

- **Backend**: Python, Flask
- **Conversion Libraries**: 
  - PIL/Pillow for image processing
  - pandas/openpyxl for spreadsheets
  - python-docx/python-pptx for Office docs
  - pdfplumber for PDF extraction
- **Frontend**: HTML5, CSS3, JavaScript
- **Deployment**: Gunicorn, Render.com ready

## 🚀 Quick Start

### Local Development

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd document-converter
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python hybrid_app.py
   ```

4. **Open your browser**
   Visit `http://127.0.0.1:5000`

### Production Deployment (Render.com)

1. **Connect your GitHub repository** to Render.com
2. **Select Web Service** deployment type
3. **Use the included `render.yaml`** for automatic configuration
4. **Set environment variables** (automatically handled)

The application will automatically install wkhtmltopdf via the `Aptfile`.

## 📊 Supported Conversions

### Core Conversions (Always Available)
- Images ↔ PDF, DOCX, PPTX, HTML, TXT, CSV, XLSX
- CSV ↔ XLSX, HTML, DOCX, PPTX, Images
- TXT ↔ HTML, PPTX, CSV, XLSX, Images
- HTML ↔ TXT, PPTX, CSV, XLSX, Images

### Enhanced with LibreOffice
- Office documents ↔ PDF
- Office documents ↔ Other Office formats
- Advanced document processing

### Enhanced with wkhtmltopdf
- HTML → PDF (high quality)
- Web page → PDF conversion

## 🏗 Architecture

```
hybrid_app.py (Main Application)
├── ConversionEngine (Core conversion logic)
├── Multiple conversion methods with fallbacks
├── Automatic cleanup and resource management
├── Comprehensive error handling
└── Production-ready configuration

templates/
├── hybrid_index.html (Main interface)
└── manifest.json (PWA support)
```

## 🔧 Configuration

### Environment Variables
- `SECRET_KEY`: Flask secret key (auto-generated in production)
- `FLASK_ENV`: Set to 'production' for deployment
- `PORT`: Application port (auto-configured)
- `CLOUDCONVERT_API_KEY`: Optional for additional formats

### Optional Dependencies
- **LibreOffice**: For office document conversions
- **wkhtmltopdf**: For HTML→PDF conversions (installed via Aptfile)

## 📈 Performance

- **50MB max file size** limit
- **Automatic cleanup** of temporary files
- **Memory-efficient** streaming operations
- **Background processing** with proper resource management
- **Concurrent request** handling

## 🛡 Security

- **File validation** and sanitization
- **Secure filename** handling
- **Path traversal** protection
- **Input validation** for all formats
- **Environment variable** configuration

## 🌐 Browser Features

- **Progressive Web App** (PWA) support
- **Drag & drop** file upload
- **Client-side conversions** for instant results
- **Responsive design** for all devices
- **Offline capability** for basic features

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📞 Support

For support or questions, please open an issue in the GitHub repository.

---

**Built with ❤️ for seamless document conversion**