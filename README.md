# Image Drop - Electron Desktop Application

A powerful Electron desktop application for capturing, converting, and organizing images with AI-powered descriptions.

## Features

### Image Capture & Import
- **Drag & Drop**: Drop PDF, Word, Excel, and image files directly into the app
- **Clipboard Paste**: Press Ctrl+V to save clipboard images
- **Screen Capture**: Built-in screen capture functionality
- **Windows Snipping Tool**: Integrate with Windows Snipping Tool for custom captures

### File Conversion
- PDF → PNG (all pages)
- Word Documents (.doc, .docx) → PNG
- Excel Spreadsheets (.xls, .xlsx) → PNG
- Images (PNG, JPG, JPEG) → PNG (standardized)

### AI Integration (Optional)
- Use OpenAI API to generate meaningful folder names
- Automatic image descriptions saved as text files
- Intelligent organization based on content

### Organization
- Automatic folder creation with timestamp or AI-generated names
- All files saved with timestamp in filename
- Description files included when using OpenAI

### Viewer
- Browse recent folders with image counts
- View all images in a folder
- Full-screen image preview
- Delete entire folders
- See AI-generated descriptions

## Installation

1. Clone or download this repository
2. Install dependencies:
```bash
npm install
```

3. Run the application:
```bash
npm start
```

## Requirements

- Node.js 14 or higher
- Windows 10/11 (for Snipping Tool integration)
- Poppler (for PDF conversion) - Install from: https://github.com/oschwaldp/node-poppler#install

### Installing Poppler (Windows)

1. Download Poppler for Windows from: https://github.com/oschwaldp/node-poppler#install
2. Extract to a folder (e.g., `C:\Program Files\poppler`)
3. Add the `bin` folder to your system PATH

## Configuration

### Company-Wide Deployment (Recommended)

For company use with a shared API key:

1. Copy `.env.example` to `.env`:
```bash
cp .env.example .env
```

2. Edit `.env` and add your company OpenAI API key:
```
OPENAI_API_KEY=sk-your-company-key-here
```

3. The `.env` file is in `.gitignore` and will NOT be committed to git
4. When deployed, OpenAI features are automatically enabled
5. Users cannot disable or change the API key (company-controlled)

**Security Notes:**
- Never commit `.env` to version control
- Distribute `.env` separately via secure channels (password manager, secure file share)
- Each installation needs the `.env` file placed in the app's root directory

### Individual User Settings

If no company key is set, users can:
1. **Root Folder**: Select where images will be saved
2. **OpenAI Integration**:
   - Check "Use OpenAI for descriptions"
   - Enter their own OpenAI API key
   - Get API key from: https://platform.openai.com/api-keys

### Default Settings
- Root folder: `Documents/ImageDrop`
- OpenAI: Auto-enabled if company key exists, otherwise disabled

## Usage

### Capture Images
1. **Drag & Drop**: Drag files into the dropzone
2. **Paste**: Copy an image and press Ctrl+V anywhere in the app
3. **Screen Capture**: Click "Capture Screen" button
4. **Snipping Tool**: Click "Snipping Tool" button, create snippet, save to clipboard

### View Images
1. Click "Open Viewer" button
2. Browse folders sorted by date (newest first)
3. Click a folder to view its images
4. Click an image for full-screen view
5. Press ESC to close full-screen view

### Folder Structure
```
Root Folder/
├── 2025-01-15_14-30-45/          # Timestamp folder (no AI)
│   ├── screenshot_2025-01-15_14-30-45.png
│   └── clipboard_2025-01-15_14-30-46.png
└── Invoice_Receipt_Document/     # AI-generated folder name
    ├── pdf-page_2025-01-15_14-31-00.png
    └── description.txt            # AI-generated description
```

## Building

To build the application for distribution:

```bash
npm run build
```

This will create installers in the `dist` folder.

## Technologies Used

- **Electron**: Desktop application framework
- **Sharp**: Image processing
- **pdf-poppler**: PDF to image conversion
- **Mammoth**: Word document processing
- **ExcelJS**: Excel spreadsheet processing
- **OpenAI**: AI-powered descriptions and naming
- **electron-store**: Settings persistence
- **screenshot-desktop**: Screen capture

## Troubleshooting

### PDF Conversion Not Working
- Make sure Poppler is installed and in your PATH
- Restart the application after installing Poppler

### Snipping Tool Not Opening
- Only works on Windows
- Make sure Snipping Tool is available in your system

### OpenAI Features Not Working
- Verify your API key is correct
- Check your internet connection
- Ensure you have API credits available

### Images Not Saving
- Check that the root folder exists and is writable
- Verify you have disk space available

## License

ISC

## Support

For issues and feature requests, please create an issue in the repository.
