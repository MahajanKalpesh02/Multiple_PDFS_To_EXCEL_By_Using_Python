import pdfplumber
import openpyxl
import re
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading  # Import threading for background task

# Function to check duplicates in the text file
def check_duplicate_in_text_file(data, text_file_path):
    if not os.path.exists(text_file_path):
        return False
    with open(text_file_path, 'r') as file:
        existing_data = file.read()
        if data in existing_data:
            return True
    return False

# Function to append new data to the text file
def append_to_text_file(data, text_file_path):
    with open(text_file_path, 'a') as file:
        file.write(data + '\n')

# Function to process the PDFs and extract data
def process_pdfs():
    pdf_paths = pdf_textbox.get("1.0", "end").strip().split("\n")
    output_folder = output_folder_var.get()

    if not pdf_paths or not output_folder:
        messagebox.showerror("Error", "Please select PDF files and an output folder.")
        return

    progress_bar["value"] = 0
    progress_label.configure(text="Processing...", fg_color="orange")

    columns = [
			
           # Give Your Columns names here...... 
           #eg. "Invoice Number"
    ]

    details_patterns = {
         # Give Your Regex Pattern here.....
         #eg."Invoice Number": r"Invoice(?: No)?[:\s]*([A-Za-z0-9\-/]+)"
    }

    excel_file_path = os.path.join(output_folder, 'extracted_data.xlsx')
    text_file_path = os.path.join(output_folder, 'extracted_data.txt')

    total_files = len(pdf_paths)
    try:
        if os.path.exists(excel_file_path):
            wb = openpyxl.load_workbook(excel_file_path)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(columns)

        for i, pdf_path in enumerate(pdf_paths):
            if not os.path.isfile(pdf_path):
                continue

            with pdfplumber.open(pdf_path) as pdf:
                text = ''
                for page in pdf.pages:
                    text += page.extract_text()

            data = []
            for column in columns:
                if column in details_patterns:
                    match = re.search(details_patterns[column], text, re.IGNORECASE)
                    if match:
                        value = match.group(1).strip()
                        data.append(value)
                    else:
                        data.append("Not Found")
                else:
                    start_index = text.find(column)
                    if start_index != -1:
                        start_value = text[start_index + len(column):].strip()
                        end_index = start_value.find("\n")
                        value = start_value[:end_index] if end_index != -1 else start_value
                        value = value.replace(":", "").strip()
                        data.append(value)
                    else:
                        data.append("Not Found")

            data_string = " | ".join(data)

            if not check_duplicate_in_text_file(data_string, text_file_path):
                ws.append(data)
                append_to_text_file(data_string, text_file_path)

            progress_bar["value"] = ((i + 1) / total_files) * 100
            progress_label.configure(text=f"Processing {i+1}/{total_files} files...", fg_color="orange")
            root.update_idletasks()

        wb.save(excel_file_path)
        progress_label.configure(text="Process Completed!", fg_color="green")
        messagebox.showinfo("Success", "Data successfully processed and saved.")
    except Exception as e:
        progress_label.configure(text="Error!", fg_color="red")
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to select multiple PDF files
def select_pdfs():
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if file_paths:
        pdf_textbox.delete("1.0", "end")
        pdf_textbox.insert("1.0", "\n".join(file_paths))

# Function to select an output folder
def select_output_folder():
    folder_path = filedialog.askdirectory()
    output_folder_var.set(folder_path)

# Function to start processing in a separate thread
def start_processing():
    processing_thread = threading.Thread(target=process_pdfs)
    processing_thread.daemon = True
    processing_thread.start()

# Initialize the main window with custom theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("PDF_TO_EXCEL Data Extractor")
root.geometry("750x600")
root.resizable(True, True)  # Make the window resizable

# Variables for file and folder paths
output_folder_var = ctk.StringVar()

# Title label
title_label = ctk.CTkLabel(root, text="PDF_To_Excel Data Extractor", font=("Arial", 20, "bold"), text_color="white")
title_label.pack(pady=10)

# PDF file selection
ctk.CTkLabel(root, text="Select PDF Files:", font=("Arial", 14), text_color="#f0f0f0").pack(pady=5)
pdf_textbox = ctk.CTkTextbox(root, width=700, height=150, corner_radius=5, border_color="#4A90E2", border_width=2)
pdf_textbox.pack(pady=5)
ctk.CTkButton(root, text="Browse PDF Files", command=select_pdfs, fg_color="#4A90E2").pack(pady=5)

# Output folder selection
ctk.CTkLabel(root, text="Select Output Folder:", font=("Arial", 14), text_color="#f0f0f0").pack(pady=5)
output_entry = ctk.CTkEntry(root, textvariable=output_folder_var, width=700, corner_radius=5)
output_entry.pack(pady=5)
ctk.CTkButton(root, text="Browse Output Folder", command=select_output_folder, fg_color="#4A90E2").pack(pady=5)

# Progress bar and label
progress_label = ctk.CTkLabel(root, text="Progress:", font=("Arial", 14), text_color="#f0f0f0")
progress_label.pack(pady=5)
progress_bar = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
progress_bar.pack(pady=5)

# Start button
ctk.CTkButton(root, text="Start Process", command=start_processing, fg_color="#28a745", hover_color="#218838", font=("Arial", 14)).pack(pady=20)

# Run the main loop
root.mainloop()
