import tkinter as tk
from tkinter import filedialog, messagebox
import os
import win32com.client 
import pythoncom # Essential for stable COM connection

# --- Core Conversion Logic (Final Robust Version) ---

def check_write_permission(filepath):
    """
    Checks if the script has permission to write a file to the directory 
    specified by the filepath. Returns True if OK, False otherwise.
    """
    directory = os.path.dirname(filepath)
    # Use os.path.expanduser to handle paths that start with ~ or environment variables
    directory = os.path.expanduser(directory)
    
    if not directory or not os.path.exists(directory):
        # If the directory doesn't exist, we assume it's valid for SaveAs operation
        return True
    
    # Try creating a temporary file to confirm write access
    temp_file = os.path.join(directory, 'temp_write_test_converter.tmp')
    try:
        with open(temp_file, 'w') as f:
            f.write('test')
        os.remove(temp_file)
        return True
    except Exception:
        return False


def convert_word_to_pdf(docx_path, pdf_path):
    """
    Converts a Word document (DOCX/DOC) to a PDF file using the MS Word COM interface.
    Uses DispatchEx for robust launching and includes COM cleanup.
    """
    word = None
    
    # Pre-conversion write check
    if not check_write_permission(pdf_path):
        print(f"ERROR: Cannot write to output directory. Check permissions for: {os.path.dirname(pdf_path)}")
        return False
        
    try:
        # 1. Initialize the COM library
        pythoncom.CoInitialize() 
        wdFormatPDF = 17
        
        # 2. Use DispatchEx for robust launch, especially in automated environments
        word = win32com.client.DispatchEx('Word.Application') 
        word.Visible = False  
        word.DisplayAlerts = False 
        
        # 3. Ensure the paths are absolute for COM
        absolute_docx_path = os.path.abspath(docx_path)
        absolute_pdf_path = os.path.abspath(pdf_path)

        # 4. Open the document
        # Note: The 'False' is for ConfirmConversions, which is usually False for DOCX/DOC
        doc = word.Documents.Open(absolute_docx_path, ReadOnly=True, ConfirmConversions=False) 
        
        # 5. Save the document as a PDF
        doc.SaveAs(absolute_pdf_path, FileFormat=wdFormatPDF)
        
        # 6. Close the document
        doc.Close(False) 
        
        return True
    
    except Exception as e:
        # This will capture and print the specific COM error code (-2147352567, etc.)
        print(f"An error occurred during conversion: {e}")
        return False
    
    finally:
        # 7. Ensure Word is quit and COM is cleaned up
        if word:
            try:
                # Attempt to kill the Word process safely
                word.Quit()
            except:
                pass
        
        pythoncom.CoUninitialize()

# --- GUI Application Class (Layout Fixes Included) ---

class WordToPDFConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Word to PDF Converter")
        
        # Set initial size wide enough for all elements and enable resizing
        master.geometry("650x250") 
        master.resizable(True, True) 
        
        # Variables to hold file paths
        self.word_path = tk.StringVar()
        self.pdf_path = tk.StringVar()

        # Configure column 1 (Entry fields) to expand when resized
        master.grid_columnconfigure(1, weight=1)
        
        # --- Layout Widgets ---
        
        # 1. Input Word File Selection
        tk.Label(master, text="Select Word Document (Input):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(master, textvariable=self.word_path).grid(row=0, column=1, padx=5, pady=10, sticky="ew") 
        tk.Button(master, text="Browse...", command=self.browse_word_file).grid(row=0, column=2, padx=10, pady=10)
        
        # 2. Output PDF Location Selection
        tk.Label(master, text="Select Save Location (Output):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(master, textvariable=self.pdf_path).grid(row=1, column=1, padx=5, pady=10, sticky="ew") 
        tk.Button(master, text="Browse...", command=self.browse_pdf_save_location).grid(row=1, column=2, padx=10, pady=10)

        # 3. Conversion Button
        tk.Button(master, text="Convert to PDF", command=self.start_conversion, bg="#305CDE", fg="white", font=('Arial', 10, 'bold')).grid(row=2, column=0, columnspan=3, pady=20, ipadx=10, ipady=5)

        # 4. Status Label
        self.status_text = tk.StringVar()
        self.status_text.set("Ready. Select files to begin.")
        self.status_label = tk.Label(master, textvariable=self.status_text, bd=1, relief=tk.SUNKEN, anchor="w")
        self.status_label.grid(row=3, column=0, columnspan=3, sticky="ew") 


    # --- Widget Command Functions (File Handling) ---
    
    def browse_word_file(self):
        """Opens a file dialog to select the Word document."""
        filetypes = (("Word Documents", "*.docx;*.doc"), ("All files", "*.*"))
        path = filedialog.askopenfilename(
            title="Select Word Document",
            initialdir=os.getcwd(),
            filetypes=filetypes
        )
        if path:
            self.word_path.set(path)
            # Suggest a default output path
            if not self.pdf_path.get():
                base_name = os.path.splitext(os.path.basename(path))[0]
                # Default output file is in the same directory as the input file
                default_pdf_path = os.path.join(os.path.dirname(path), f"{base_name}.pdf")
                self.pdf_path.set(default_pdf_path)
            self.status_text.set(f"Word file selected: {os.path.basename(path)}")
            

    def browse_pdf_save_location(self):
        """Opens a file dialog to choose the output location and filename for the PDF."""
        filetypes = (("PDF files", "*.pdf"), ("All files", "*.*"))
        default_filename = ""
        if self.word_path.get():
            default_filename = os.path.splitext(os.path.basename(self.word_path.get()))[0] + ".pdf"
            
        path = filedialog.asksaveasfilename(
            title="Save PDF As...",
            initialdir=os.path.dirname(self.word_path.get()) if self.word_path.get() else os.getcwd(),
            initialfile=default_filename,
            defaultextension=".pdf",
            filetypes=filetypes
        )
        if path:
            self.pdf_path.set(path)
            self.status_text.set(f"Save location set: {os.path.basename(path)}")


    def start_conversion(self):
        """Validates paths and initiates the conversion process."""
        docx_path = self.word_path.get()
        pdf_path = self.pdf_path.get()
        
        # 1. Validation Checks
        if not docx_path or not pdf_path:
            messagebox.showerror("Error", "Please select both a Word document and a save location for the PDF.")
            return
            
        if not os.path.exists(docx_path):
            messagebox.showerror("Error", "Selected Word file does not exist.")
            return

        # 2. Start Conversion
        self.status_text.set("Converting... Please wait. Do not close the window.")
        self.master.update() # Force GUI update to show "Converting..." message
        
        # Run the conversion
        success = convert_word_to_pdf(docx_path, pdf_path)
        
        # 3. Final Feedback
        if success:
            messagebox.showinfo("Success", f"Successfully converted to PDF:\n{pdf_path}")
            self.status_text.set("Conversion complete! Find your PDF in the specified location.")
        else:
            messagebox.showerror("Conversion Failed", 
                                 "The conversion failed. This is often a permission or architecture issue. Check the console for the technical error code.")
            self.status_text.set("Conversion failed. See error pop-up and console.")


if __name__ == "__main__":
    root = tk.Tk()
    app = WordToPDFConverterApp(root)
    root.title("Word to PDF Converter")
    root.iconbitmap("convert.ico")  
    root.mainloop()