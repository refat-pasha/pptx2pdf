## install this first

- import os
- import traceback

- from flask import Flask, request, send_from_directory, render_template, redirect, url_for, flash
- import comtypes.client

- import comtypes


# PPTX to PDF Converter

A simple Flask web application that converts PowerPoint (.pptx) files to PDF format using Python and COM automation.

## Features

- **Web Interface**: User-friendly web interface for file upload and conversion
- **Batch Processing**: Convert multiple PPTX files at once
- **Automatic Download**: Converted PDF files are automatically served for download
- **Error Handling**: Robust error handling with user feedback
- **Clean Interface**: Simple and intuitive design

## Prerequisites

- Windows operating system (required for COM automation)
- Microsoft PowerPoint installed on the system
- Python 3.6 or higher

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/refat-pasha/pptx2pdf.git
   cd pptx2pdf
   ```

2. **Install required dependencies**
   ```bash
   pip install flask comtypes
   ```

3. **Create necessary directories**
   The application will automatically create `uploads` and `downloads` directories if they don't exist.

## Usage

1. **Start the Flask application**
   ```bash
   python app.py
   ```

2. **Access the web interface**
   Open your web browser and navigate to `http://localhost:5000`

3. **Upload and Convert**
   - Select one or more PPTX files using the file upload interface
   - Click the convert button
   - Wait for the conversion process to complete
   - Download the converted PDF files

## Code Structure

```
pptx2pdf/
│
├── app.py              # Main Flask application
├── templates/          # HTML templates
│   └── index.html      # Main upload interface
├── uploads/           # Temporary storage for uploaded files
├── downloads/         # Storage for converted PDF files
└── README.md          # This file
```

## Key Components

### Main Application (app.py)
- Flask web server setup
- File upload handling
- PowerPoint to PDF conversion using COM automation
- Error handling and user feedback
- File serving for downloads

### Dependencies
- **Flask**: Web framework for the user interface
- **comtypes**: Python COM automation library for interacting with PowerPoint
- **os**: File system operations
- **traceback**: Error tracking and debugging

## How It Works

1. Users upload PPTX files through the web interface
2. Files are temporarily stored in the `uploads` directory
3. The application uses COM automation to open PowerPoint
4. Each PPTX file is opened and exported as PDF
5. Converted PDF files are saved to the `downloads` directory
6. Users can download the converted files

## Error Handling

The application includes comprehensive error handling for:
- Invalid file formats
- PowerPoint application errors
- File system permissions
- Network interruptions during upload/download

## Limitations

- **Windows Only**: Requires Windows OS with PowerPoint installed
- **PowerPoint Dependency**: Microsoft PowerPoint must be installed and accessible
- **COM Limitations**: Subject to COM automation limitations and potential stability issues

## Security Considerations

- File uploads are handled securely
- Temporary files are managed automatically
- Consider implementing file size limits for production use
- Add file type validation for enhanced security

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is open source. Please check the repository for license details.

## Troubleshooting

### Common Issues

**PowerPoint not found error**
- Ensure Microsoft PowerPoint is installed
- Check if PowerPoint is properly registered for COM automation

**Permission errors**
- Run the application with appropriate permissions
- Check file system permissions for uploads/downloads directories

**Conversion failures**
- Verify the PPTX file is not corrupted
- Ensure PowerPoint can open the file manually

## Support

For issues, questions, or contributions, please visit the [GitHub repository](https://github.com/refat-pasha/pptx2pdf) and create an issue.

---

**Author**: refat-pasha  
**Repository**: https://github.com/refat-pasha/pptx2pdf
