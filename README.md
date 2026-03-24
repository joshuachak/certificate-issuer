# CertStudio - Local Certificate Issuer

CertStudio is a 100% private, local web application designed to automate the generation of certificates. All processing happens entirely in your browser—no data is ever sent to a server.

## 🌟 Features

- **Privacy First**: Process sensitive student data locally. No external APIs or cloud storage.
- **Template Support**: Import certificate templates as PDF or Image (PNG/JPG).
- **Batch Processing**: Import recipient lists via Quick Paste (newline separated) or TSV/CSV files.
- **Interactive Editor**: Drag-and-drop placeholders for Names, Dates, and IDs.
- **PowerPoint-like Controls**: Full control over font size, colors, alignment, and typography.
- **Advanced Zoom**: Google Maps style zoom (Ctrl + Scroll / Pinch) for precise placement.
- **Dual Export Modes**:
    - **Protected (Image PDF)**: Flattened PDF where text is baked into an image (non-selectable).
    - **Interactive (Selectable PDF)**: PDF with actual text layers for searchability.
- **ZIP Download**: Generates and packages all individual PDFs into a single ZIP file.

## 🚀 How to Run

Because modern browsers have security restrictions on `file://` URLs, you must run this app using a simple local server to enable PDF processing.

### Option A: Using Node.js (Recommended)
1. Open your terminal in this folder.
2. Run:
   ```bash
   npx serve .
   ```
3. Open the link provided (usually `http://localhost:3000`).

### Option B: Using Python
1. Open your terminal in this folder.
2. Run:
   ```bash
   python3 -m http.server 8000
   ```
3. Open `http://localhost:8000` in your browser.

## 🛠 Usage Guide

1. **Upload Template**: Select your blank certificate (PDF or Image).
2. **Input Data**: 
   - Paste names in the "Quick Paste" tab.
   - Or upload a TSV/CSV file in the "Batch Upload" tab. You can download a sample template from there.
3. **Place Fields**: Click "+ Name", "+ Date", or "+ ID" to add placeholders.
4. **Style**: Select a placeholder on the canvas to change its font, size, and color.
5. **Position**: Drag the placeholders to the exact spot on the certificate.
6. **Export**: Choose your output format and click **"Export Certificates"**.

## 📦 Tech Stack

- **Fabric.js**: Interactive Canvas Engine.
- **PDF.js**: Client-side PDF rendering.
- **jsPDF**: PDF generation.
- **JSZip**: In-browser ZIP compression.
- **PapaParse**: CSV/TSV parsing.

---
*Developed with ❤️ for privacy and efficiency.*
