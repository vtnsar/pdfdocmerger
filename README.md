# PDF Page Modifier & Merger

## Overview

The **PDF Page Modifier & Merger** is a Python application built using Tkinter that allows users to upload, modify, and merge PDF files. The application supports various file formats, including Word documents, PowerPoint presentations, and image files, converting them to PDF as needed.

## Features

- Upload multiple files (PDF, DOC, DOCX, PPT, PPTX, JPG, JPEG, PNG, TIFF).
- Modify PDF pages by adding numeration, deleting specific pages, and merging multiple PDFs.
- User-friendly GUI built with Tkinter.
- Progress bar to indicate processing status.
- Settings window to customize options.

## Requirements

- Python 3.x
- Tkinter (comes pre-installed with Python)
- Pillow (for image processing)
- PyPDF2 (for PDF manipulation)
- comtypes and pywin32 (for Windows COM automation)

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/pdf-merger-app.git
   cd pdf-merger-app
   ```

2. **Install the required packages**:
   You can install the required packages using pip:
   ```bash
   pip install Pillow PyPDF2 comtypes pywin32
   ```

3. **Run the application**:
   ```bash
   python tkinterapp.py
   ```

## Usage

1. Launch the application.
2. Click on the "Upload Files" button to select files you want to modify or merge.
3. Adjust the settings as needed in the "Settings" window.
4. Use the "Process & Merge PDFs" button to start processing the uploaded files.
5. The merged PDF will be saved in the same directory as the first uploaded file.

## Screenshots

*(Add screenshots of the application here)*

## Contributing

Contributions are welcome! If you have suggestions for improvements or new features, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Tkinter for the GUI framework.
- Pillow for image processing.
- PyPDF2 for PDF manipulation.
- comtypes and pywin32 for Windows automation.

## Contact

For any questions or feedback, please contact [your.email@example.com](mailto:your.email@example.com).
