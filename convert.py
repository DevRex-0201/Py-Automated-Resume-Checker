import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Button
import os
import shutil
import comtypes.client

def doc_to_docx_and_move(doc_path, doc_folder):
    # Initialize the COM library
    comtypes.CoInitialize()

    try:
        # Load Word
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False  # Make Word run in the background
        
        # Convert the doc_path to an absolute path (to avoid COM object path issues)
        doc_path_absolute = os.path.abspath(doc_path)

        # Open the document
        doc = word.Documents.Open(doc_path_absolute)
        doc.Activate()

        # Ensure the doc folder exists
        if not os.path.exists(doc_folder):
            os.makedirs(doc_folder)

        # Construct new file path for .docx
        new_file_path = os.path.splitext(doc_path)[0] + ".docx"
        doc.SaveAs2(new_file_path, FileFormat=16)  # FileFormat=16 for .docx
        
        # Close the document and quit Word
        doc.Close(False)
        word.Quit()
    except Exception as e:
        raise Exception(f"Error converting file '{doc_path}': {e}")
    finally:
        # Uninitialize the COM library
        comtypes.CoUninitialize()

    # Move the original .doc file to the doc folder
    new_doc_path = os.path.join(doc_folder, os.path.basename(doc_path))
    shutil.move(doc_path, new_doc_path)
    
    return new_file_path

def convert_folder_docs_to_docx(folder_path):
    doc_folder = os.path.join(folder_path, "doc")
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".doc"):
                doc_path = os.path.join(root, file)
                try:
                    new_file_path = doc_to_docx_and_move(doc_path, doc_folder)
                    print(f"Converted and moved '{doc_path}' to '{new_file_path}'")
                except Exception as e:
                    print(f"Failed to convert and move '{doc_path}': {e}")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        try:
            convert_folder_docs_to_docx(folder_selected)
            messagebox.showinfo("Conversion Complete", "All DOC files have been converted to DOCX and moved to 'doc' folder.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# UI Setup
root = tk.Tk()
root.title("DOC to DOCX Converter and Mover")

browse_button = Button(root, text="Browse Folder", command=browse_folder)
browse_button.pack(pady=20)

root.mainloop()
