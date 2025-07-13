import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import requests
import json
import base64
from io import BytesIO
from PIL import Image, ImageTk
import os
import threading
import re
import csv
import openpyxl
import pandas as pd
from datetime import datetime

class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR.space Text & Table Extractor")
        self.root.geometry("900x700")
        
        # API Key
        self.api_key = "K83294946888957"  # Your existing API key
        
        # Variables
        self.image_path = None
        self.image_url = None
        self.current_image = None
        self.ocr_results = None
        self.table_data = []
        self.extracted_text = ""
        self.formatted_tables = ""
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="OCR Text & Table Extractor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="Input Options", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        # File upload
        ttk.Label(input_frame, text="Upload Image:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.file_path_var, state="readonly").grid(
            row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_file).grid(
            row=0, column=2, pady=5)
        
        # URL input
        ttk.Label(input_frame, text="Image URL:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.url_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.url_var).grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        ttk.Button(input_frame, text="Load URL", command=self.load_url).grid(
            row=1, column=2, pady=5)
        
        # OCR Options
        options_frame = ttk.LabelFrame(main_frame, text="OCR Options", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Language selection
        ttk.Label(options_frame, text="Language:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.language_var = tk.StringVar(value="eng")
        language_combo = ttk.Combobox(options_frame, textvariable=self.language_var, 
                                     values=["eng", "spa", "fre", "ger", "ita", "por", "rus", "jpn", "chi_sim"], 
                                     state="readonly", width=10)
        language_combo.grid(row=0, column=1, sticky=tk.W, padx=(10, 20), pady=5)
        
        # OCR Engine
        ttk.Label(options_frame, text="OCR Engine:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.engine_var = tk.StringVar(value="2")
        engine_combo = ttk.Combobox(options_frame, textvariable=self.engine_var, 
                                   values=["1", "2", "3"], state="readonly", width=5)
        engine_combo.grid(row=0, column=3, sticky=tk.W, padx=(10, 20), pady=5)
        
        # Table detection option
        self.table_var = tk.BooleanVar(value=True)
        table_check = ttk.Checkbutton(options_frame, text="Enable Table Detection", 
                                     variable=self.table_var)
        table_check.grid(row=0, column=4, sticky=tk.W, padx=(10, 0), pady=5)
        
        # Process and Export buttons
        process_frame = ttk.Frame(main_frame)
        process_frame.grid(row=3, column=0, columnspan=3, pady=10)

        self.process_btn = ttk.Button(process_frame, text="Extract Text & Tables", 
                             command=self.process_ocr, style="Accent.TButton")
        self.process_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Export button with dropdown menu
        self.export_btn = ttk.Button(process_frame, text="Export Results", 
                                    command=self.show_export_menu, state="disabled")
        self.export_btn.grid(row=0, column=1, padx=(0, 10))
        
        # Create export menu
        self.export_menu = tk.Menu(self.root, tearoff=0)
        self.export_menu.add_command(label="Export as Text File", command=self.export_text)
        self.export_menu.add_command(label="Export as CSV", command=self.export_csv)
        self.export_menu.add_command(label="Export as Excel", command=self.export_excel)
        self.export_menu.add_separator()
        self.export_menu.add_command(label="Copy Raw Text", command=self.copy_text)
        self.export_menu.add_command(label="Copy Formatted Tables", command=self.copy_table)
        
        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        results_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)
        
        # Image preview and text output in notebook
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Image preview tab
        self.preview_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.preview_frame, text="Image Preview")
        
        # Create scrollable frame for image
        self.preview_canvas = tk.Canvas(self.preview_frame, bg="white")
        preview_scrollbar_v = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.preview_canvas.yview)
        preview_scrollbar_h = ttk.Scrollbar(self.preview_frame, orient="horizontal", command=self.preview_canvas.xview)
        self.preview_canvas.configure(yscrollcommand=preview_scrollbar_v.set, xscrollcommand=preview_scrollbar_h.set)
        
        self.preview_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_scrollbar_v.grid(row=0, column=1, sticky=(tk.N, tk.S))
        preview_scrollbar_h.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_frame.rowconfigure(0, weight=1)
        
        # Raw text output tab
        self.text_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.text_frame, text="Raw Text")
        
        self.text_output = scrolledtext.ScrolledText(self.text_frame, wrap=tk.WORD, height=15)
        self.text_output.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.text_frame.columnconfigure(0, weight=1)
        self.text_frame.rowconfigure(0, weight=1)
        
        # Formatted table output tab
        self.table_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.table_frame, text="Formatted Tables")
        
        self.table_output = scrolledtext.ScrolledText(self.table_frame, wrap=tk.NONE, height=15, font=("Courier", 10))
        self.table_output.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.table_frame.columnconfigure(0, weight=1)
        self.table_frame.rowconfigure(0, weight=1)
        
        # Action buttons frame - simplified and more accessible
        button_frame = ttk.Frame(results_frame)
        button_frame.grid(row=1, column=0, pady=10)

        # Copy buttons
        self.copy_text_btn = ttk.Button(button_frame, text="Copy Raw Text", 
                               command=self.copy_text, state="disabled")
        self.copy_text_btn.grid(row=0, column=0, padx=5, pady=5)

        self.copy_table_btn = ttk.Button(button_frame, text="Copy Formatted Tables", 
                                command=self.copy_table, state="disabled")
        self.copy_table_btn.grid(row=0, column=1, padx=5, pady=5)

        # Export buttons - more prominent and accessible
        self.export_csv_btn = ttk.Button(button_frame, text="Export as CSV", 
                                command=self.export_csv, state="disabled")
        self.export_csv_btn.grid(row=0, column=2, padx=5, pady=5)

        self.export_excel_btn = ttk.Button(button_frame, text="Export as Excel", 
                                  command=self.export_excel, state="disabled")
        self.export_excel_btn.grid(row=0, column=3, padx=5, pady=5)

        # Save button
        self.save_btn = ttk.Button(button_frame, text="Save to File", 
                          command=self.save_results, state="disabled")
        self.save_btn.grid(row=0, column=4, padx=5, pady=5)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.image_path = file_path
            self.image_url = None
            self.url_var.set("")
            self.load_image_preview(file_path)
            
    def load_url(self):
        url = self.url_var.get().strip()
        if url:
            self.image_url = url
            self.image_path = None
            self.file_path_var.set("")
            self.load_image_preview(url, is_url=True)
            
    def load_image_preview(self, source, is_url=False):
        try:
            if is_url:
                self.status_var.set("Loading image from URL...")
                self.progress.start()
                response = requests.get(source, timeout=10)
                response.raise_for_status()
                image = Image.open(BytesIO(response.content))
            else:
                image = Image.open(source)
            
            # Resize image for preview
            display_size = (500, 400)
            image.thumbnail(display_size, Image.Resampling.LANCZOS)
            
            # Convert to PhotoImage
            self.current_image = ImageTk.PhotoImage(image)
            
            # Clear canvas and add image
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=self.current_image)
            
            # Update scroll region
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
            
            # Switch to preview tab
            self.notebook.select(self.preview_frame)
            
            self.status_var.set("Image loaded successfully")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load image: {str(e)}")
            self.status_var.set("Error loading image")
        finally:
            self.progress.stop()
            
    def process_ocr(self):
        if not self.image_path and not self.image_url:
            messagebox.showwarning("No Image", "Please select an image file or enter an image URL")
            return
            
        # Run OCR in separate thread to prevent UI blocking
        threading.Thread(target=self.run_ocr, daemon=True).start()
        
    def run_ocr(self):
        try:
            # Update UI on main thread
            self.root.after(0, self.start_processing)
            
            # OCR.space API endpoint
            url = "https://api.ocr.space/parse/image"
            
            # Headers
            headers = {
                'apikey': self.api_key
            }
            
            # Prepare data
            data = {
                'language': self.language_var.get(),
                'isOverlayRequired': 'false',
                'OCREngine': self.engine_var.get(),
                'detectOrientation': 'true',
                'scale': 'true',
                'isTable': str(self.table_var.get()).lower()  # Enable table detection
            }
            
            if self.image_url:
                # Process URL
                data['url'] = self.image_url
                response = requests.post(url, headers=headers, data=data, timeout=30)
            else:
                # Process local file
                with open(self.image_path, 'rb') as file:
                    files = {'file': file}
                    response = requests.post(url, headers=headers, data=data, files=files, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                
                if result.get('IsErroredOnProcessing', False):
                    error_msg = "OCR processing failed:\n"
                    error_messages = result.get('ErrorMessage', [])
                    if isinstance(error_messages, list):
                        error_msg += "\n".join(error_messages)
                    else:
                        error_msg += str(error_messages)
                    
                    self.root.after(0, self.show_error, error_msg)
                else:
                    # Extract text and tables
                    self.ocr_results = result  # Store full results
                    extracted_text = ""
                    formatted_tables = ""
                    self.table_data = []
                    
                    parsed_results = result.get('ParsedResults', [])
                    
                    for i, parsed_result in enumerate(parsed_results):
                        if len(parsed_results) > 1:
                            extracted_text += f"--- Result {i+1} ---\n"
                            formatted_tables += f"--- Result {i+1} ---\n"
                        
                        parsed_text = parsed_result.get('ParsedText', '')
                        extracted_text += parsed_text.strip() + "\n"
                        
                        # Format tables if detected
                        if self.table_var.get():
                            table_text, table_rows = self.format_table_text(parsed_text)
                            formatted_tables += table_text + "\n"
                            self.table_data.extend(table_rows)
                    
                    if not extracted_text.strip():
                        extracted_text = "No text found in the image."
                    
                    if not formatted_tables.strip():
                        formatted_tables = "No tables detected in the image."
                    
                    # Store results for export
                    self.extracted_text = extracted_text
                    self.formatted_tables = formatted_tables
                    
                    # Update UI on main thread
                    self.root.after(0, self.show_results, extracted_text, formatted_tables)
            else:
                self.root.after(0, self.show_error, f"HTTP Error {response.status_code}: {response.text}")
                
        except requests.exceptions.RequestException as e:
            self.root.after(0, self.show_error, f"Network error: {str(e)}")
        except Exception as e:
            self.root.after(0, self.show_error, f"Unexpected error: {str(e)}")
        finally:
            self.root.after(0, self.finish_processing)
    
    def format_table_text(self, text):
        """Format text to better display tables and extract table data"""
        if not text.strip():
            return "No table content detected.", []
        
        lines = text.split('\n')
        formatted_lines = []
        table_rows = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # Check if line contains table-like data
            if self.is_table_like_line(line):
                # Split by multiple spaces or tabs and format as table
                parts = re.split(r'\s{2,}|\t+', line)
                if len(parts) > 1:
                    # Clean and format parts
                    clean_parts = [part.strip() for part in parts if part.strip()]
                    if clean_parts:
                        # Format as table row with proper spacing
                        formatted_line = " | ".join(clean_parts)
                        formatted_lines.append(formatted_line)
                        table_rows.append(clean_parts)
                else:
                    formatted_lines.append(line)
            else:
                # Also check for simple space-separated values
                parts = line.split()
                if len(parts) > 1:
                    # Look for patterns that suggest columns (numbers, currencies, etc.)
                    has_numbers = any(re.search(r'\d', part) for part in parts)
                    if has_numbers and len(parts) >= 2:
                        formatted_line = " | ".join(parts)
                        formatted_lines.append(formatted_line)
                        table_rows.append(parts)
                    else:
                        formatted_lines.append(line)
                else:
                    formatted_lines.append(line)
        
        # Add table separators for better readability
        if formatted_lines:
            result = []
            for i, line in enumerate(formatted_lines):
                if '|' in line:
                    if i == 0 or '|' not in formatted_lines[i-1]:
                        # Add header separator
                        separator = "-" * len(line)
                        result.append(separator)
                    result.append(line)
                    if i == len(formatted_lines) - 1 or '|' not in formatted_lines[i+1]:
                        # Add footer separator
                        separator = "-" * len(line)
                        result.append(separator)
                else:
                    result.append(line)
            return '\n'.join(result), table_rows
        
        return text, table_rows
    
    def is_table_like_line(self, line):
        """Check if a line looks like it contains table data"""
        # Look for patterns that suggest tabular data
        parts = line.split()
        if len(parts) < 2:
            return False
        
        # Check for mixed numbers and text (common in tables)
        has_numbers = any(part.replace('.', '').replace(',', '').replace('$', '').replace('%', '').isdigit() for part in parts)
        has_text = any(not part.replace('.', '').replace(',', '').replace('$', '').replace('%', '').isdigit() for part in parts)
        
        # Check for common table separators or patterns
        has_separators = any(sep in line for sep in ['|', '\t', '  '])
        
        return (has_numbers and has_text) or has_separators
            
    def start_processing(self):
        self.status_var.set("Processing image...")
        self.progress.start()
        self.process_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.copy_text_btn.config(state="disabled")
        self.copy_table_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.export_csv_btn.config(state="disabled")
        self.export_excel_btn.config(state="disabled")
        
    def finish_processing(self):
        self.progress.stop()
        self.process_btn.config(state="normal")
        
    def show_results(self, text, table_text):
        # Show raw text
        self.text_output.delete(1.0, tk.END)
        self.text_output.insert(1.0, text)
        
        # Show formatted tables
        self.table_output.delete(1.0, tk.END)
        self.table_output.insert(1.0, table_text)
        
        # Switch to appropriate tab
        if self.table_var.get() and "No tables detected" not in table_text:
            self.notebook.select(self.table_frame)
        else:
            self.notebook.select(self.text_frame)
        
        self.status_var.set("Text and table extraction completed")
        self.export_btn.config(state="normal")
        self.copy_text_btn.config(state="normal")
        self.copy_table_btn.config(state="normal")
        self.save_btn.config(state="normal")
        self.export_csv_btn.config(state="normal")
        self.export_excel_btn.config(state="normal")
        
    def show_error(self, error_msg):
        messagebox.showerror("Error", error_msg)
        self.status_var.set("Error occurred")
        
    def show_export_menu(self):
        """Show export menu when export button is clicked"""
        if not self.extracted_text and not self.table_data:
            messagebox.showwarning("No Data", "No data available to export. Please process an image first.")
            return
            
        # Get button position to show menu
        x = self.export_btn.winfo_rootx()
        y = self.export_btn.winfo_rooty() + self.export_btn.winfo_height()
        self.export_menu.tk_popup(x, y)
        
    def export_text(self):
        """Export results as text file"""
        if not self.extracted_text:
            messagebox.showwarning("No Data", "No text data available to export.")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="Export as Text File",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("=== OCR RESULTS ===\n")
                    f.write(f"Processed on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                    f.write("=== RAW TEXT ===\n")
                    f.write(self.extracted_text)
                    f.write("\n\n=== FORMATTED TABLES ===\n")
                    f.write(self.formatted_tables)
                
                self.status_var.set("Text file exported successfully")
                messagebox.showinfo("Success", f"Results exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export text file: {str(e)}")
        
    def copy_text(self):
        if not self.extracted_text:
            messagebox.showwarning("No Data", "No text data available to copy.")
            return
            
        self.root.clipboard_clear()
        self.root.clipboard_append(self.extracted_text)
        self.status_var.set("Raw text copied to clipboard")
        messagebox.showinfo("Success", "Raw text copied to clipboard!")
    
    def copy_table(self):
        if not self.formatted_tables:
            messagebox.showwarning("No Data", "No table data available to copy.")
            return
            
        self.root.clipboard_clear()
        self.root.clipboard_append(self.formatted_tables)
        self.status_var.set("Formatted tables copied to clipboard")
        messagebox.showinfo("Success", "Formatted tables copied to clipboard!")
    
    def save_results(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Results",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("=== RAW TEXT ===\n")
                    f.write(self.text_output.get(1.0, tk.END))
                    f.write("\n\n=== FORMATTED TABLES ===\n")
                    f.write(self.table_output.get(1.0, tk.END))
                
                self.status_var.set("Results saved successfully")
                messagebox.showinfo("Success", f"Results saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")
    
    def export_csv(self):
        """Export table data as CSV"""
        if not self.table_data:
            messagebox.showwarning("No Table Data", "No table data available to export.")
            return
    
        file_path = filedialog.asksaveasfilename(
            title="Export as CSV",
            defaultextension=".csv",
            filetypes=[
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
    
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    
                    # Find max columns across all rows
                    max_cols = max(len(row) for row in self.table_data) if self.table_data else 0
                
                    # Write header
                    if max_cols > 0:
                        writer.writerow([f"Column {i+1}" for i in range(max_cols)])
                
                    # Write data rows - pad shorter rows with empty strings
                    for row in self.table_data:
                        padded_row = row + [''] * (max_cols - len(row))
                        writer.writerow(padded_row)
            
                self.status_var.set("CSV exported successfully")
                messagebox.showinfo("Success", f"Table data exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export CSV: {str(e)}")

    def export_excel(self):
        if not self.table_data:
            messagebox.showwarning("No Table Data", "No table data available to export.")
            return
    
        file_path = filedialog.asksaveasfilename(
            title="Export as Excel",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
    
        if file_path:
            try:
                # Create DataFrame
                max_cols = max(len(row) for row in self.table_data) if self.table_data else 0
                
                if max_cols == 0:
                    messagebox.showwarning("No Data", "No table data to export.")
                    return
            
                columns = [f"Column {i+1}" for i in range(max_cols)]
            
                # Pad rows to same length
                padded_data = []
                for row in self.table_data:
                    padded_row = row + [''] * (max_cols - len(row))
                    padded_data.append(padded_row)
            
                df = pd.DataFrame(padded_data, columns=columns)
            
                # Export to Excel
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='OCR_Results')
            
                self.status_var.set("Excel exported successfully")
                messagebox.showinfo("Success", f"Table data exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export Excel: {str(e)}")

def main():
    root = tk.Tk()
    app = OCRApp(root)
    root.mainloop()

if __name__ == "__main__":
    # Required packages: pip install tkinter pillow requests pandas openpyxl
    main()