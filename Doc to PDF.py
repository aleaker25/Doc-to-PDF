import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import win32com.client
import pythoncom  # Essential for stable COM connection
import logging
import threading
import configparser  # For storing settings
import darkdetect  # For detecting system theme
import time


# Configure logging
logging.basicConfig(filename='converter.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')


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


def convert_word_to_pdf(docx_path, pdf_path, quality='Standard'):
    """
    Converts a Word document (DOCX/DOC) to a PDF file using the MS Word COM interface.
    Uses DispatchEx for robust launching and includes COM cleanup.
    """
    word = None

    # Pre-conversion write check
    if not check_write_permission(pdf_path):
        print(
            f"ERROR: Cannot write to output directory. Check permissions for: {os.path.dirname(pdf_path)}")
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
        doc = word.Documents.Open(absolute_docx_path, ReadOnly=True,
                                   ConfirmConversions=False)

        # 5. Save the document as a PDF
        # Adjust PDF quality based on the selected option
        if quality == 'Minimum':
            doc.SaveAs(absolute_pdf_path, FileFormat=wdFormatPDF, OptimizeFor=2)  # Minimize size
        elif quality == 'Standard':
            doc.SaveAs(absolute_pdf_path, FileFormat=wdFormatPDF, OptimizeFor=0)  # Standard
        else:  # Maximum
            doc.SaveAs(absolute_pdf_path, FileFormat=wdFormatPDF)  # No optimization

        # 6. Close the document
        doc.Close(False)

        return True

    except Exception as e:
        # This will capture and print the specific COM error code (-2147352567, etc.)
        print(f"An error occurred during conversion: {e}")
        logging.error(f"Conversion error: {e}")
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

class WordToPDFConverterApp:  # Don't inherit from TkinterDnD.Tk
    def __init__(self, master):
        self.master = master
        self.master.title("Word to PDF Converter")

        # Set initial size wide enough for all elements and enable resizing
        self.master.geometry("650x300")
        self.master.resizable(True, True)

        # Variables to hold file paths
        self.word_path = tk.StringVar(master)
        self.pdf_path = tk.StringVar(master)
        self.quality = tk.StringVar(master, value='Standard')  # Default quality
        self.theme_mode = tk.StringVar(master, value='System')  # Default theme mode

        # Configure column 1 (Entry fields) to expand when resized
        self.master.grid_columnconfigure(1, weight=1)

        # --- Layout Widgets ---

        # 1. Input Word File Selection
        tk.Label(master, text="Select Word Document (Input):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.word_entry = tk.Entry(master, textvariable=self.word_path)
        self.word_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        tk.Button(master, text="Browse...", command=self.browse_word_file).grid(row=0, column=2, padx=10, pady=10)

        # 2. Output PDF Location Selection
        tk.Label(master, text="Select Save Location (Output):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.pdf_entry = tk.Entry(master, textvariable=self.pdf_path)
        self.pdf_entry.grid(row=1, column=1, padx=5, pady=10, sticky="ew")
        tk.Button(master, text="Browse...", command=self.browse_pdf_save_location).grid(row=1, column=2, padx=10,
                                                                                           pady=10)

        # 3. Quality Selection
        tk.Label(master, text="Select PDF Quality:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        quality_choices = ['Minimum', 'Standard', 'Maximum']
        quality_menu = tk.OptionMenu(master, self.quality, *quality_choices)
        quality_menu.grid(row=2, column=1, padx=5, pady=10, sticky="ew")

        # 4. Conversion Button
        tk.Button(master, text="Convert to PDF", command=self.start_conversion, bg="#305CDE", fg="white",
                  font=('Arial', 10, 'bold')).grid(row=3, column=0, columnspan=3, pady=20, ipadx=10, ipady=5)

        # 5. Progress Bar
        self.progress = ttk.Progressbar(master, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky="ew", padx=10, pady=5)

        # 6. Status Label
        self.status_text = tk.StringVar()
        self.status_text.set("Ready. Select files to begin.")
        self.status_label = tk.Label(master, textvariable=self.status_text, bd=1, relief=tk.SUNKEN, anchor="w")
        self.status_label.grid(row=5, column=0, columnspan=3, sticky="ew")

        # --- Menu Bar ---
        self.menu_bar = tk.Menu(master)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Exit", command=master.quit)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)

        self.theme_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.theme_menu.add_command(label="Light", command=lambda: self.set_theme("Light"))
        self.theme_menu.add_command(label="Dark", command=lambda: self.set_theme("Dark"))
        self.theme_menu.add_command(label="System", command=lambda: self.set_theme("System"))
        self.menu_bar.add_cascade(label="Theme", menu=self.theme_menu)

        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.help_menu.add_command(label="About", command=self.show_about)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)

        master.config(menu=self.menu_bar)

        # --- Initial Theme Setup ---
        self.set_theme(self.theme_mode.get())

    def update_progress(self, value):
        """Updates the progress bar value."""
        self.progress['value'] = value
        self.master.update_idletasks()  # Force GUI update

    # --- Theme Handling ---
    def set_theme(self, mode):
        """Sets the theme of the application."""
        self.theme_mode.set(mode)
        if mode == "System":
            system_theme = darkdetect.theme()
            if system_theme == "Dark":
                self.apply_dark_theme()
            else:
                self.apply_light_theme()
        elif mode == "Dark":
            self.apply_dark_theme()
        else:
            self.apply_light_theme()

    def apply_light_theme(self):
        """Applies the light theme."""
        self.master.config(bg="white")
        fg_color = "black"
        bg_color = "white"
        button_bg = "#305CDE"
        button_fg = "white"

        for widget in self.master.winfo_children():
            widget_class = widget.winfo_class()
            if widget_class in ("TLabel", "Label"):
                widget.config(foreground=fg_color, background=bg_color)
            elif widget_class in ("TButton", "Button"):
                widget.config(foreground=button_fg, background=button_bg, relief=tk.RAISED)
            elif widget_class in ("TEntry", "Entry"):
                widget.config(foreground=fg_color, background="white", highlightbackground=bg_color)
            elif widget_class == "Frame":
                widget.config(background=bg_color)

        self.status_label.config(foreground=fg_color, background=bg_color)

    def apply_dark_theme(self):
        """Applies the dark theme."""
        self.master.config(bg="#333333")
        fg_color = "white"
        bg_color = "#333333"
        button_bg = "#555555"
        button_fg = "white"

        for widget in self.master.winfo_children():
            widget_class = widget.winfo_class()
            if widget_class in ("TLabel", "Label"):
                widget.config(foreground=fg_color, background=bg_color)
            elif widget_class in ("TButton", "Button"):
                widget.config(foreground=button_fg, background=button_bg, relief=tk.RAISED)
            elif widget_class in ("TEntry", "Entry"):
                widget.config(foreground=fg_color, background="#555555", highlightbackground=bg_color,
                              insertbackground=fg_color)
            elif widget_class == "Frame":
                widget.config(background=bg_color)

        self.status_label.config(foreground=fg_color, background=bg_color)

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
        quality = self.quality.get()

        # 1. Validation Checks
        if not docx_path or not pdf_path:
            messagebox.showerror("Error", "Please select both a Word document and a save location for the PDF.")
            return

        if not os.path.exists(docx_path):
            messagebox.showerror("Error", "Selected Word file does not exist.")
            return

        # 2. Start Conversion
        self.status_text.set("Converting... Please wait. Do not close the window.")
        self.master.update()  # Force GUI update to show "Converting..." message
        self.progress['value'] = 0  # Reset progress bar
        self.perform_conversion(docx_path, pdf_path, quality)  # Call perform_conversion directly

    def perform_conversion(self, docx_path, pdf_path, quality):
        """Performs the conversion and updates the GUI."""
        # Use a thread to run the conversion
        conversion_thread = threading.Thread(target=self.run_conversion, args=(docx_path, pdf_path, quality))
        conversion_thread.start()

    def run_conversion(self, docx_path, pdf_path, quality):
        """Executes the conversion and updates the GUI."""
        try:
            # Run the conversion
            success = convert_word_to_pdf(docx_path, pdf_path, quality)

            # Update GUI after conversion
            self.master.after(10, self.update_after_conversion, success, pdf_path)

        finally:
            # After conversion is complete, set progress bar to 100
            self.master.after(10, self.update_progress, 100)

    def update_after_conversion(self, success, pdf_path):
        """Updates the GUI elements after the conversion is complete."""
        # 3. Final Feedback
        if success:
            messagebox.showinfo("Success", f"Successfully converted to PDF:\n{pdf_path}")
            self.status_text.set("Conversion complete! Find your PDF in the specified location.")
        else:
            messagebox.showerror("Conversion Failed",
                                 "The conversion failed. This is often a permission or architecture issue. Check the console for the technical error code.")
            self.status_text.set("Conversion failed. See error pop-up and console.")

    def show_about(self):
        """Displays an About dialog box."""
        about_message = "Word to PDF Converter\nVersion 1.0\nCreated by Alec"
        messagebox.showinfo("About", about_message)


if __name__ == "__main__":
    try:
        root = tk.Tk()  # Create a Tk instance

        app = WordToPDFConverterApp(root)  # Pass the root to the app
        root.title("Word to PDF Converter")
        root.iconbitmap("convert.ico")
        root.mainloop()
    except Exception as e:
        print(f"An error occurred during GUI initialization: {e}")