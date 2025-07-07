import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from doctr.io import DocumentFile
from doctr.models import ocr_predictor

# Load OCR model
model = ocr_predictor(pretrained=True)

# Store extracted results
extracted_results = []

def extract_text_from_file(file_path):
    try:
        doc = DocumentFile.from_images(file_path) if file_path.lower().endswith(('.png', '.jpg', '.jpeg')) else DocumentFile.from_pdf(file_path)
        result = model(doc)
        extracted_text = result.render()
        return extracted_text
    except Exception as e:
        return f"Error processing {file_path}: {str(e)}"

def browse_files():
    global extracted_results
    file_paths = filedialog.askopenfilenames(title="Select Files (Image or PDF)", 
                                             filetypes=[("Image and PDF files", "*.jpg *.jpeg *.png *.pdf")])
    text_box.delete("1.0", tk.END)
    extracted_results = []

    for path in file_paths:
        text = extract_text_from_file(path)
        extracted_results.append({'filename': os.path.basename(path), 'text': text})
        text_box.insert(tk.END, f"\n--- Extracted from: {os.path.basename(path)} ---\n{text}\n")

def save_to_csv():
    if not extracted_results:
        messagebox.showwarning("No Data", "Please extract text first.")
        return
    df = pd.DataFrame(extracted_results)
    df.to_csv("ocr_results.csv", index=False)
    messagebox.showinfo("Saved", "Data saved to ocr_results.csv")

def save_to_excel():
    if not extracted_results:
        messagebox.showwarning("No Data", "Please extract text first.")
        return
    df = pd.DataFrame(extracted_results)
    df.to_excel("ocr_results.xlsx", index=False)
    messagebox.showinfo("Saved", "Data saved to ocr_results.xlsx")

def calculate_field_accuracy():
    """
    A basic field accuracy check for Invoice, PO, Amount, Date.
    """
    if not extracted_results:
        messagebox.showwarning("No Data", "Extract text first to evaluate accuracy.")
        return

    key_fields = ['Invoice', 'PO', 'Amount', 'Date']
    score = 0
    total = 0

    for result in extracted_results:
        text = result['text'].lower()
        for field in key_fields:
            total += 1
            if field.lower() in text:
                score += 1

    accuracy = (score / total) * 100 if total > 0 else 0
    messagebox.showinfo("Field Accuracy", f"Approx. Field Detection Accuracy: {accuracy:.2f}%")

# GUI setup
root = tk.Tk()
root.title("Advanced OCR Extractor (with Save + Accuracy)")
root.geometry("900x700")

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Select Files", command=browse_files, width=15).grid(row=0, column=0, padx=5)
tk.Button(btn_frame, text="Save to CSV", command=save_to_csv, width=15).grid(row=0, column=1, padx=5)
tk.Button(btn_frame, text="Save to Excel", command=save_to_excel, width=15).grid(row=0, column=2, padx=5)
tk.Button(btn_frame, text="Check Accuracy", command=calculate_field_accuracy, width=15).grid(row=0, column=3, padx=5)

text_box = tk.Text(root, wrap="word", font=("Courier", 10))
text_box.pack(expand=True, fill="both")

root.mainloop()
