import os
import tkinter as tk
from tkinter import filedialog, messagebox
from doctr.io import DocumentFile
from doctr.models import ocr_predictor

# Load OCR model
model = ocr_predictor(pretrained=True)

def extract_text_from_file(file_path):
    try:
        # Load document (image or pdf)
        doc = DocumentFile.from_images(file_path) if file_path.lower().endswith(('.png', '.jpg', '.jpeg')) else DocumentFile.from_pdf(file_path)
        result = model(doc)
        extracted_text = result.render()
        return extracted_text
    except Exception as e:
        return f"Error processing {file_path}: {str(e)}"

def browse_files():
    file_paths = filedialog.askopenfilenames(title="Select Files (Image or PDF)", 
                                             filetypes=[("Image and PDF files", "*.jpg *.jpeg *.png *.pdf")])
    text_box.delete("1.0", tk.END)  # Clear existing text
    for path in file_paths:
        text_box.insert(tk.END, f"\n--- Extracted from: {os.path.basename(path)} ---\n")
        extracted = extract_text_from_file(path)
        text_box.insert(tk.END, extracted + "\n")

# GUI setup
root = tk.Tk()
root.title("Basic OCR Extractor using Doctr")
root.geometry("800x600")

btn = tk.Button(root, text="Select Files (IMG, PDF)", command=browse_files)
btn.pack(pady=10)

text_box = tk.Text(root, wrap="word", font=("Courier", 10))
text_box.pack(expand=True, fill="both")

root.mainloop()
