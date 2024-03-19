import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from docx import Document
import fitz  # PyMuPDF for PDF processing
import win32com.client as win32  # Required for .doc files
import logging
import math
import re


# Set up logging
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function to extract text from a Word document (.docx)
def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        text = [para.text for para in doc.paragraphs]
        return '\n'.join(text)
    except Exception as e:
        logging.error(f"Error processing .docx file {file_path}: {e}")
        return None

# Function to extract text from an older Word document (.doc)
def extract_text_from_doc(file_path):
    try:
        word = win32.Dispatch("Word.Application")
        word.visible = False
        doc = word.Documents.Open(file_path)
        text = doc.Range().Text
        doc.Close()
        word.Quit()
        return text
    except Exception as e:
        logging.error(f"Error processing .doc file {file_path}: {e}")
        return None

# Function to extract text from a PDF document
def extract_text_from_pdf(file_path):
    try:
        doc = fitz.open(file_path)
        text = ''
        for page in doc:
            extracted_text = page.get_text()
            text += extracted_text.encode('utf-8').decode('utf-8')
        doc.close()
        return text
    except Exception as e:
        logging.error(f"Error processing PDF file {file_path}: {e}")
        return None

# Function to try reading a CSV with different encodings
def try_read_csv(file_path):
    encodings = ['utf-8', 'utf-16', 'iso-8859-1', 'cp1252']  # common encodings
    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding)
        except UnicodeDecodeError:
            continue
    logging.error(f"Failed to read CSV file {file_path} with any known encoding.")
    return None

# Function to read and process the database file
def read_database(file_path):
    try:
        if file_path.endswith('.xlsx'):
            return pd.read_excel(file_path)
        else:
            return try_read_csv(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read database file: {e}")
        logging.error(f"Failed to read database file {file_path}: {e}")
        return None

# Function to process the files and update the database
def process_files(folder, database_file):
    db = read_database(database_file)
    if db is None:
        messagebox.showerror("Error", "Could not read the database file with any known encoding.")
        return

    # Add new columns for matched filenames and COS
    db['Resume'] = None
    db['COS'] = None
    db["SCORE"] = None

    for filename in os.listdir(folder):
        file_text = None
        file_path = os.path.join(folder, filename)
        if filename.endswith('.docx'):
            file_text = extract_text_from_docx(file_path)
        elif filename.endswith('.pdf'):
            file_text = extract_text_from_pdf(file_path)
            # print(filename)
            # print(file_text)
        elif filename.endswith('.doc'):
            file_text = extract_text_from_doc(file_path)
        # print(file_text)
        if file_text is not None:
            file_text_updated = re.sub(r"[\s()\-+]", "", file_text.lower())
            cos_key_word = re.sub(r"[\s()\-+]", "", "Terms and Conditions for Contract of Service – Version 3.0".lower())
            for index, row in db.iterrows():
                con_email = any(str(row[key]).lower() in file_text.lower() for key in ['Email']) and str(row['Email']) != "nan"
                con_phone = any((str(row[key]).lower().replace(' ', '').replace('.0', '').replace('(', '').replace(')', '').replace('-', '').replace('+', '') in file_text_updated for key in ['Phone'])) and str(row['Phone']) != "nan"
                con_mobile = any((str(row[key]).lower().replace(' ', '').replace('.0', '').replace('(', '').replace(')', '').replace('-', '').replace('+', '') in file_text_updated for key in ['Mobile'])) and str(row['Mobile']) != "nan"
                con_name = any(str(row[key]).lower() in file_text.lower() for key in ['Name']) and str(row['Name']) != "nan" and " " in str(row['Name'])
                con_100 = con_email and con_phone and con_mobile and con_name
                con_80 = con_email and con_phone and con_name
                con_70 = con_email and con_phone
                con_60 = con_email and con_mobile
                con_50 = con_email and con_name
                con_40 = con_phone and con_name
                con_30 = con_mobile and con_name

                if con_100: 
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 100
                    break

                elif con_80:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 80
                    break

                elif con_70:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 70
                    break

                elif con_60:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 60
                    break

                elif con_50:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 50
                    break

                elif con_40 or con_email:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 40
                    break

                elif con_30 or con_phone:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 30
                    break

                elif con_mobile:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 20
                    break

                elif con_name:                    
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    db.at[index, 'SCORE'] = 10
                    break

                elif any(str(row[key]).lower() in file_text.lower() for key in ['Name']) and str(row['Name']) != "nan" and " " in str(row['Name']).strip(): 
                    if cos_key_word in file_text_updated:
                        db.at[index, 'COS'] = filename
                    else:
                        db.at[index, 'Resume'] = filename
                    break


    # Save the updated database back to the Excel file
    try:
        db.to_excel(database_file, index=False)
        messagebox.showinfo("Completed", f"Processing complete. Database updated at {database_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save the updated database: {e}")
        logging.error(f"Failed to save the updated database {database_file}: {e}")

# Tkinter GUI functions
def select_folder():
    folder_path = filedialog.askdirectory()
    folder_entry.delete(0, tk.END)
    folder_entry.insert(0, folder_path)

def select_database():
    file_path = filedialog.askopenfilename()
    database_entry.delete(0, tk.END)
    database_entry.insert(0, file_path)

def start_processing():
    folder = folder_entry.get()
    database_file = database_entry.get()
    process_files(folder, database_file)

# Build the Tkinter GUI
root = tk.Tk()
root.title("Resume Matcher")

tk.Label(root, text="Select Folder:").grid(row=0)
folder_entry = tk.Entry(root, width=50)
folder_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=select_folder).grid(row=0, column=2)

tk.Label(root, text="Select Database File:").grid(row=1)
database_entry = tk.Entry(root, width=50)
database_entry.grid(row=1, column=1)
tk.Button(root, text="Browse", command=select_database).grid(row=1, column=2)

tk.Button(root, text="Start Processing", command=start_processing).grid(row=2, column=1)

root.mainloop()
