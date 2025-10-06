# 📄 Word to PDF Converter GUI

A robust, standalone Python application built with **Tkinter** for the GUI and **`pywin32`** for reliable, formatting-preserving conversion of Microsoft Word documents (`.docx`, `.doc`) to PDF files.

---

## ✨ Features

* **GUI Interface:** Simple, user-friendly interface using Tkinter.
* **Accurate Conversion:** Leverages the native Microsoft Word COM automation to ensure **perfect preservation of document formatting** in the resulting PDF.
* **File Browsing:** Dedicated browse functions for selecting the input Word file and the output PDF save location.
* **Robustness:** Includes checks for directory write permissions and uses stable COM practices (`DispatchEx`, `CoInitialize`) to prevent common conversion failures and zombie Word processes.
* **Resizable Window:** The application window is fully resizable and responsive.

---

## ⚠️ Prerequisites (Windows Only)

This application relies on the **Microsoft Word COM interface** to perform the conversion. Therefore, it has two hard requirements:

1.  **Operating System:** Must be run on a **Windows** machine.
2.  **Microsoft Word:** **Microsoft Word must be installed** on the machine running the script or executable.

---

## 🚀 Installation & Setup

### 1. Install Dependencies

You'll need the following Python libraries. We specifically need `pywin32` for interacting with Word.

```bash
pip install pywin32 tk
```  

### 2. Post-Installation Check  

After installing `pywin32`, it's a good practice to run the post-install script to ensure the COM components are registered correctly, especially if you encounter errors:

```bash
# Run this command in your terminal
python -m win32com.client -install
```  

## 📝 Usage  

### Running the Script  

- Save the entire code block into a file named, for example, word_to_pdf_converter.py.  
- Run the application from your terminal:
`python word_to_pdf_converter.py`

### Using the GUI  

- Click **"Browse..."** next to "Select Word Document (Input)" to choose your `.docx` or `.doc` file.  

- Click **"Browse..."** next to "Select Save Location (Output)" to choose the output directory and filename for your `.pdf` file.  

- Click the **"Convert to PDF"** button. The status bar will update to "Converting... Please wait."  

- A success or failure message will appear, and the status bar will show the final result.  
