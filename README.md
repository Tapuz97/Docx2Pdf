# # [![Buy Me a Coffee]\(https\://i.imgur.com/rlatSuk.png)]\(https\://www\.buymeacoffee.com/galmitrani1)

# DOCX to PDF Converter

## Introduction

The DOCX to PDF Converter is a Python utility designed to automate the conversion of `.docx` files into `.pdf` format. This tool is especially useful for users who need to batch-process Word documents into PDFs efficiently without manual intervention.

## Features

- **Batch Conversion**: Converts all `.docx` files in the `docx/` folder and outputs PDFs to the `pdf/` folder.
- **Read-Only Mode**: Opens Word documents in read-only mode to prevent modification and avoid file lock issues.
- **Auto Process Cleanup**: Ensures Microsoft Word (`WINWORD.EXE`) is fully closed after conversion, preventing locked files.
- **Error Handling**: Skips temporary Word files (`~$filename.docx`) and logs conversion failures.

## Usage

### 1. **Prepare Files**

- Place all `.docx` files inside a folder named `docx/`.

### 2. **Directory Structure**

Ensure your project directory is set up as follows:

```
main_folder/
├── docx/        # Folder containing DOCX files
├── pdf/         # Folder where PDFs will be saved (created automatically)
├── convert.py   # Python script file
```

### 3. **Run the Script**

Execute the script to convert all `.docx` files:

```bash
python convert.py
```

### 4. **Review Converted Files**

All PDFs will be saved in the `pdf/` folder.

## Error Handling

The script ensures smooth operation by:

- Ignoring temp files (`~$filename.docx`)
- Logging errors when a file fails to convert
- Force-closing Microsoft Word (`WINWORD.EXE`) to prevent locked files

## Requirements

- **Windows OS** (with Microsoft Word installed)
- **Python 3.x**
- **Dependencies**:
  - `comtypes` (for Word automation)

Install dependencies using:

```bash
pip install comtypes
```

## How to Use

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/docx-to-pdf.git
   ```
2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Place ****`.docx`**** Files in the ****`docx/`**** Folder**.
4. **Run the Script**:
   ```bash
   python convert.py
   ```

## Local Development

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/docx-to-pdf.git
   ```
2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Modify the Code** as needed.
4. **Run the Script**:
   ```bash
   python convert.py
   ```

## Contribution

Contributions are welcome! To contribute:

1. **Fork the Repository**.
2. **Create a Branch**:
   ```bash
   git checkout -b feature/new-feature
   ```
3. **Commit Your Changes**:
   ```bash
   git commit -m 'Added a new feature'
   ```
4. **Push to the Branch**:
   ```bash
   git push origin feature/new-feature
   ```
5. **Open a Pull Request**.

## License

This project is licensed under the MIT License.

## Acknowledgments

- Thanks to all contributors for improving this tool.
- Special appreciation to the open-source community for supporting automation tools.

---

Thank you for using the DOCX to PDF Converter! 🚀

