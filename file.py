import os
import tkinter as tk
from tkinter import ttk, filedialog
import PyPDF2
from docx import Document


class FileConverterApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("File Converter")
        self.window.geometry("500x400")
        self.window.resizable(False, False)

        # File selection frame
        self.file_frame = ttk.LabelFrame(self.window, text="Select File", padding=(20, 10))
        self.file_frame.pack(fill="x", padx=20, pady=10)

        self.file_path = tk.StringVar()
        self.selected_extension = tk.StringVar()

        # File selection label
        ttk.Label(self.file_frame, text="File Path: ").grid(column=0, row=0, sticky="W")

        # File selection entry
        ttk.Entry(self.file_frame, textvariable=self.file_path, width=50, state="readonly").grid(column=1, row=0, padx=10)

        # File selection button
        ttk.Button(self.file_frame, text="Select File", command=self.select_file).grid(column=2, row=0, padx=10)

        # Conversion frame
        self.conversion_frame = ttk.LabelFrame(self.window, text="Conversion Options", padding=(20, 10))
        self.conversion_frame.pack(fill="x", padx=20, pady=10)

        # Extension selection label
        ttk.Label(self.conversion_frame, text="Select Extension: ").grid(column=0, row=0, sticky="W")

        # Extension selection combobox
        self.extension_combobox = ttk.Combobox(self.conversion_frame, textvariable=self.selected_extension, state="readonly")
        self.extension_combobox["values"] = [".txt", ".pdf", ".docx"]
        self.extension_combobox.current(0)
        self.extension_combobox.grid(column=1, row=0, padx=10)

        # Convert file button
        ttk.Button(self.conversion_frame, text="Convert", command=self.convert_file).grid(column=2, row=0, padx=10)

        self.window.mainloop()

    def select_file(self):
        file_path = filedialog.askopenfilename()
        self.file_path.set(file_path)

    def convert_file(self):
        # Get the file path and extension
        file_path = self.file_path.get()
        extension = self.selected_extension.get()

        # Convert PDF to DOCX
        if extension == ".docx" and os.path.splitext(file_path)[1] == ".pdf":
            # Extract text from PDF
            with open(file_path, "rb") as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                text = "\n".join([pdf_reader.pages[page_num].extract_text() for page_num in range(len(pdf_reader.pages))])

            # Create a new document and add the text to it
            doc = Document()
            doc.add_paragraph(text)

            # Save the document
            save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
            doc.save(save_path)

        # Change the extension of the file
        else:
            basename = os.path.splitext(os.path.basename(file_path))[0]
            new_file_path = os.path.join(os.path.dirname(file_path), basename + extension)
            os.rename(file_path, new_file_path)

            # Save file dialog
            save_path = filedialog.asksaveasfilename(defaultextension=extension, filetypes=[(f"{extension} files", f"*{extension}")])
            os.rename(new_file_path, save_path)



if __name__ == "__main__":
    converter = FileConverterApp()
