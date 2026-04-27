# CourseExtractor

CourseExtractor is a standalone desktop application that extracts structured data from course PDF files and exports it into a professionally formatted Excel spreadsheet. It uses the same extraction logic as the main RA_IAA system but operates entirely locally without the need for a database or web server.

## Features

- **PDF & ZIP Support**: Upload individual PDF files or a ZIP archive containing multiple course PDFs.
- **Automatic Expansion**: ZIP files are automatically scanned and all contained PDFs are processed.
- **One Sheet Per Course**: Creates a separate Excel worksheet for every course, named by its NRC (e.g., `ING-101`).
- **Formatted Export**: Mirrors the app's visual hierarchy with color-coded sections for:
  - Course Information & Description
  - Pre-requisite Network
  - Learning Outcomes (RAs)
  - Graduation Profile Contributions (APEs)
  - Basic Bibliography
- **Native GUI**: Built with `CustomTkinter` for a modern, dark-themed desktop experience.
- **Single Executable**: Can be compiled into a standalone `.exe` or binary that runs on machines without Python installed.

## Setup for Development

If you want to run the app from source or make modifications:

1. **Navigate to the directory**:
   ```bash
   cd /home/diego/Codes/CourseExtractor
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python app.py
   ```

## Distribution (Building the Executable)

To create a single, self-contained executable for non-technical users:

1. **Run the build script**:
   ```bash
   bash build.sh
   ```

2. **Find the result**:
   The executable will be created in the `dist/` folder:
   - **Linux/Mac**: `dist/CourseExtractor`
   - **Windows**: `dist/CourseExtractor.exe`

## Dependencies

- `pdfplumber`: For high-precision PDF text and table extraction.
- `customtkinter`: For the modern native GUI.
- `openpyxl`: For creating and styling the Excel `.xlsx` files.
- `pyinstaller`: Used by the build script to bundle the app into a single binary.
