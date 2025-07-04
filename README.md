# Apple Mail Exporter

A beautiful Electron application that exports Apple Mail emails to PDF format. The app first exports emails as .eml files using AppleScript, then converts them to professionally formatted PDF documents.

## Features

- 🍎 **Apple Mail Integration**: Seamlessly connects to Apple Mail accounts and folders
- 📄 **PDF Export**: Converts emails to beautifully formatted PDF documents
- 🎨 **Modern UI**: Clean, responsive interface with progress tracking
- 📁 **Flexible Output**: Choose your own output directory
- 🔢 **Batch Processing**: Export multiple emails at once
- 📱 **Cross-Platform**: Works on macOS (requires Apple Mail)

## Prerequisites

- macOS (Apple Mail is required)
- Node.js 16+ and npm
- Apple Mail app installed and configured with at least one email account

## Installation

1. **Clone or download the project**
   ```bash
   git clone <repository-url>
   cd AppleMail
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start the application**
   ```bash
   npm start
   ```

## Usage

### Step 1: Select Mail Account & Folder
- The app will automatically load your Apple Mail accounts
- Select the account you want to export from
- Choose a specific folder (INBOX, Sent, etc.)

### Step 2: Configure Export Settings
- Set the number of emails to export (1-100)
- Choose your output directory where PDF files will be saved

### Step 3: Export
- Click "Export to PDF" to start the process
- Monitor progress in real-time
- View results and open the output folder when complete

## How It Works

1. **AppleScript Integration**: Uses AppleScript to access Apple Mail and export emails as .eml files
2. **Email Parsing**: Parses .eml files using the `mailparser` library to extract email content, metadata, and attachments
3. **PDF Generation**: Uses Puppeteer to convert email content into professionally formatted PDF documents
4. **Cleanup**: Removes temporary .eml files after successful PDF conversion

## Technical Details

### Dependencies
- **Electron**: Cross-platform desktop app framework
- **Puppeteer**: PDF generation from HTML content
- **mailparser**: Email parsing and content extraction
- **fs-extra**: Enhanced file system operations

### File Structure
```
AppleMail/
├── main.js              # Main Electron process
├── preload.js           # Preload script for secure IPC
├── index.html           # Main application UI
├── styles.css           # Application styling
├── renderer.js          # Frontend logic
├── utils/
│   ├── emailParser.js   # Email parsing utilities
│   └── pdfConverter.js  # PDF conversion utilities
└── package.json         # Project configuration
```

## Troubleshooting

### Common Issues

1. **"Error loading mail accounts"**
   - Make sure Apple Mail is running
   - Ensure you have at least one email account configured
   - Check that Apple Mail has necessary permissions

2. **"Export failed"**
   - Verify the output directory is writable
   - Check that you have sufficient disk space
   - Ensure the selected folder contains emails

3. **PDF generation issues**
   - Make sure you have a stable internet connection (for Puppeteer)
   - Check that the output path doesn't contain special characters

### Permissions
The app requires access to:
- Apple Mail (for reading emails)
- File system (for saving PDFs)
- Network (for Puppeteer PDF generation)

## Development

### Running in Development Mode
```bash
npm run dev
```

### Building for Distribution
```bash
npm run build
```

### Project Structure
- `main.js`: Handles AppleScript execution and file operations
- `renderer.js`: Manages UI interactions and user experience
- `utils/`: Contains email parsing and PDF conversion logic

## License

MIT License - feel free to use this project for personal or commercial purposes.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Support

If you encounter any issues or have questions, please:
1. Check the troubleshooting section above
2. Review the console output for error messages
3. Ensure all prerequisites are met
4. Create an issue with detailed information about your problem

---

**Note**: This application is designed specifically for macOS and requires Apple Mail to be installed and configured. 