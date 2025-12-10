"""
GUI Application for Azure AI Document Intelligence
Modern interface for document processing and table extraction
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
from pathlib import Path
import pandas as pd
from document_intelligence_service import DocumentIntelligenceService
import json
from datetime import datetime


class DocumentIntelligenceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Azure AI Document Intelligence - Table Extractor")
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)
        
        # Initialize service
        self.service = None
        self.current_file = None
        self.results = {}
        
        # Configure style
        self.setup_styles()
        
        # Create GUI
        self.create_widgets()
        
        # Try to initialize service
        self.initialize_service()
    
    def setup_styles(self):
        """Configure ttk styles for modern look"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Success.TLabel', foreground='green')
        style.configure('Error.TLabel', foreground='red')
        style.configure('Info.TLabel', foreground='blue')
    
    def create_widgets(self):
        """Create and layout all GUI widgets"""
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Azure AI Document Intelligence", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Left panel - Controls
        self.create_control_panel(main_frame)
        
        # Right panel - Results
        self.create_results_panel(main_frame)
        
        # Bottom panel - Status and progress
        self.create_status_panel(main_frame)
    
    def create_control_panel(self, parent):
        """Create the left control panel"""
        control_frame = ttk.LabelFrame(parent, text="Document Processing", padding="10")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        control_frame.columnconfigure(0, weight=1)
        
        # Service status
        self.service_status = ttk.Label(control_frame, text="Initializing service...", style='Info.TLabel')
        self.service_status.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # File selection
        ttk.Label(control_frame, text="Select Document:", style='Heading.TLabel').grid(row=1, column=0, sticky=tk.W, pady=(10, 5))
        
        file_frame = ttk.Frame(control_frame)
        file_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state='readonly')
        self.file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        self.browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_btn.grid(row=0, column=1)
        
        # Model selection
        ttk.Label(control_frame, text="Processing Model:", style='Heading.TLabel').grid(row=3, column=0, sticky=tk.W, pady=(10, 5))
        
        self.model_var = tk.StringVar(value="prebuilt-layout")
        model_frame = ttk.Frame(control_frame)
        model_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.layout_radio = ttk.Radiobutton(model_frame, text="Layout (Tables + Basic Text Analysis)", 
                                           variable=self.model_var, value="prebuilt-layout")
        self.layout_radio.grid(row=0, column=0, sticky=tk.W)
        
        self.document_radio = ttk.Radiobutton(model_frame, text="Document (Tables + Full Key-Value Extraction)", 
                                             variable=self.model_var, value="prebuilt-document")
        self.document_radio.grid(row=1, column=0, sticky=tk.W)
        
        # Model availability info
        self.model_info = ttk.Label(model_frame, text="", foreground="gray")
        self.model_info.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # Output directory
        ttk.Label(control_frame, text="Output Directory:", style='Heading.TLabel').grid(row=5, column=0, sticky=tk.W, pady=(10, 5))
        
        output_frame = ttk.Frame(control_frame)
        output_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_path_var = tk.StringVar(value="output")
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path_var)
        self.output_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(output_frame, text="Browse", command=self.browse_output_dir).grid(row=0, column=1)
        
        # Process button
        self.process_btn = ttk.Button(control_frame, text="Process Document", 
                                     command=self.process_document, state='disabled')
        self.process_btn.grid(row=7, column=0, pady=20, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress = ttk.Progressbar(control_frame, mode='indeterminate')
        self.progress.grid(row=8, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File format help
        help_frame = ttk.Frame(control_frame)
        help_frame.grid(row=9, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        help_text = "💡 Tips: Use high-quality scans (300 DPI), ensure files aren't corrupted"
        ttk.Label(help_frame, text=help_text, foreground="gray", font=('Arial', 8)).grid(row=0, column=0, sticky=tk.W)
        
        # Action buttons
        button_frame = ttk.Frame(control_frame)
        button_frame.grid(row=10, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        self.open_output_btn = ttk.Button(button_frame, text="Open Output Folder", 
                                         command=self.open_output_folder, state='disabled')
        self.open_output_btn.grid(row=0, column=0, padx=(0, 5), sticky=(tk.W, tk.E))
        
        self.clear_btn = ttk.Button(button_frame, text="Clear Results", command=self.clear_results)
        self.clear_btn.grid(row=0, column=1, padx=(5, 0), sticky=(tk.W, tk.E))
    
    def create_results_panel(self, parent):
        """Create the right results panel"""
        results_frame = ttk.LabelFrame(parent, text="Processing Results", padding="10")
        results_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)
        
        # Results summary
        self.results_summary = ttk.Label(results_frame, text="No document processed yet")
        self.results_summary.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Notebook for different result views
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Tables tab
        self.create_tables_tab()
        
        # Raw data tab
        self.create_raw_data_tab()
        
        # Key-value pairs tab
        self.create_kv_tab()
    
    def create_tables_tab(self):
        """Create the tables results tab"""
        tables_frame = ttk.Frame(self.notebook)
        self.notebook.add(tables_frame, text="Extracted Tables")
        
        tables_frame.columnconfigure(0, weight=1)
        tables_frame.rowconfigure(1, weight=1)
        
        # Table selection
        table_select_frame = ttk.Frame(tables_frame)
        table_select_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(table_select_frame, text="Select Table:").grid(row=0, column=0, padx=(0, 10))
        
        self.table_var = tk.StringVar()
        self.table_combo = ttk.Combobox(table_select_frame, textvariable=self.table_var, 
                                       state='readonly', width=20)
        self.table_combo.grid(row=0, column=1, padx=(0, 10))
        self.table_combo.bind('<<ComboboxSelected>>', self.on_table_selected)
        
        ttk.Button(table_select_frame, text="Export to Excel", 
                  command=self.export_table_excel).grid(row=0, column=2)
        
        # Table display
        self.create_table_display(tables_frame)
    
    def create_table_display(self, parent):
        """Create table display with treeview"""
        # Frame for treeview and scrollbars
        tree_frame = ttk.Frame(parent)
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Treeview for table data
        self.table_tree = ttk.Treeview(tree_frame)
        self.table_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.table_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.table_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.table_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.table_tree.configure(xscrollcommand=h_scrollbar.set)
    
    def create_raw_data_tab(self):
        """Create the raw data tab"""
        raw_frame = ttk.Frame(self.notebook)
        self.notebook.add(raw_frame, text="Raw JSON Data")
        
        raw_frame.columnconfigure(0, weight=1)
        raw_frame.rowconfigure(0, weight=1)
        
        self.raw_text = scrolledtext.ScrolledText(raw_frame, wrap=tk.WORD)
        self.raw_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    def create_kv_tab(self):
        """Create the key-value pairs tab"""
        kv_frame = ttk.Frame(self.notebook)
        self.notebook.add(kv_frame, text="Key-Value Pairs")
        
        kv_frame.columnconfigure(0, weight=1)
        kv_frame.rowconfigure(0, weight=1)
        
        # Treeview for key-value pairs
        self.kv_tree = ttk.Treeview(kv_frame, columns=('Key', 'Value'), show='headings')
        self.kv_tree.heading('Key', text='Key')
        self.kv_tree.heading('Value', text='Value')
        self.kv_tree.column('Key', width=200)
        self.kv_tree.column('Value', width=400)
        self.kv_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for KV tree
        kv_scrollbar = ttk.Scrollbar(kv_frame, orient=tk.VERTICAL, command=self.kv_tree.yview)
        kv_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.kv_tree.configure(yscrollcommand=kv_scrollbar.set)
    
    def create_status_panel(self, parent):
        """Create the bottom status panel"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, height=6, wrap=tk.WORD)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Add initial message
        self.log_message("Application started. Please configure your Azure credentials in .env file.")
    
    def initialize_service(self):
        """Initialize the Document Intelligence service"""
        try:
            self.service = DocumentIntelligenceService()
            self.service_status.config(text="✓ Service ready", style='Success.TLabel')
            self.process_btn.config(state='normal')
            self.log_message("Azure AI Document Intelligence service initialized successfully.")
            
            # Check model availability
            self.check_model_availability()
            
        except Exception as e:
            self.service_status.config(text="✗ Service error", style='Error.TLabel')
            self.log_message(f"Failed to initialize service: {str(e)}")
            self.log_message("Please check your .env file with AZURE_ENDPOINT and AZURE_KEY")
    
    def check_model_availability(self):
        """Check which models are available and update UI"""
        try:
            # Test prebuilt-document availability
            test_thread = threading.Thread(target=self._test_model_availability)
            test_thread.daemon = True
            test_thread.start()
        except Exception as e:
            self.log_message(f"Could not check model availability: {e}")
    
    def _test_model_availability(self):
        """Test model availability in background thread"""
        try:
            # Instead of testing with dummy content, we'll check available models differently
            # Skip the actual API test to avoid InvalidContent errors
            
            # For now, assume prebuilt-layout is always available and prebuilt-document might not be
            # We'll detect prebuilt-document availability during actual processing
            self.root.after(0, lambda: self.model_info.config(text="Models will be tested during processing"))
            self.root.after(0, lambda: self.log_message("Model availability will be checked during document processing"))
                    
        except Exception as e:
            self.root.after(0, lambda: self.log_message(f"Model availability check failed: {e}"))
    
    def _disable_document_model(self):
        """Disable the document model option"""
        self.document_radio.config(state='disabled')
        self.model_info.config(text="⚠ Document model not available in your region", foreground="orange")
        self.log_message("Note: prebuilt-document model not available. Using prebuilt-layout only.")
    
    def browse_file(self):
        """Browse for input file"""
        filetypes = [
            ("All supported", "*.pdf;*.jpg;*.jpeg;*.png;*.tiff;*.tif"),
            ("PDF files", "*.pdf"),
            ("Image files", "*.jpg;*.jpeg;*.png;*.tiff;*.tif"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select document to process",
            filetypes=filetypes
        )
        
        if filename:
            self.file_path_var.set(filename)
            self.current_file = filename
            self.log_message(f"Selected file: {os.path.basename(filename)}")
    
    def browse_output_dir(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_path_var.set(directory)
    
    def process_document(self):
        """Process the selected document"""
        if not self.current_file or not os.path.exists(self.current_file):
            messagebox.showerror("Error", "Please select a valid document file")
            return
        
        if not self.service:
            messagebox.showerror("Error", "Service not initialized. Check your Azure credentials.")
            return
        
        # Validate file before processing
        if not self.validate_file(self.current_file):
            return
        
        # Disable UI during processing
        self.process_btn.config(state='disabled')
        self.progress.start()
        
        # Run processing in separate thread
        thread = threading.Thread(target=self._process_document_thread)
        thread.daemon = True
        thread.start()
    
    def validate_file(self, file_path):
        """Validate file before processing"""
        try:
            # Check file size
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                messagebox.showerror("Error", "Selected file is empty")
                return False
            
            if file_size > 50 * 1024 * 1024:  # 50MB limit
                messagebox.showerror("Error", "File is too large (max 50MB)")
                return False
            
            # Check file extension
            valid_extensions = {'.pdf', '.jpg', '.jpeg', '.png', '.tiff', '.tif'}
            file_ext = Path(file_path).suffix.lower()
            
            if file_ext not in valid_extensions:
                messagebox.showerror("Error", f"Unsupported file format: {file_ext}\nSupported: PDF, JPG, PNG, TIFF")
                return False
            
            # Try to read file and validate format
            try:
                with open(file_path, 'rb') as f:
                    # Read first few bytes to ensure file is readable and check format
                    header = f.read(1024)
                    if len(header) == 0:
                        messagebox.showerror("Error", "File appears to be empty or corrupted")
                        return False
                    
                    # Basic file format validation based on magic numbers
                    if not self.validate_file_format(header, file_ext):
                        messagebox.showerror("Error", f"File format validation failed. The file may be corrupted or not a valid {file_ext.upper()} file.")
                        return False
                        
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read file: {str(e)}")
                return False
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"File validation failed: {str(e)}")
            return False
    
    def validate_file_format(self, header_bytes, file_ext):
        """Validate file format based on magic numbers"""
        try:
            # PDF files
            if file_ext == '.pdf':
                return header_bytes.startswith(b'%PDF-')
            
            # JPEG files
            elif file_ext in ['.jpg', '.jpeg']:
                return header_bytes.startswith(b'\xff\xd8\xff')
            
            # PNG files
            elif file_ext == '.png':
                return header_bytes.startswith(b'\x89PNG\r\n\x1a\n')
            
            # TIFF files
            elif file_ext in ['.tiff', '.tif']:
                return (header_bytes.startswith(b'II*\x00') or  # Little endian
                       header_bytes.startswith(b'MM\x00*'))     # Big endian
            
            # If we can't validate, assume it's okay
            return True
            
        except Exception:
            # If validation fails, assume file is okay to avoid false positives
            return True
    
    def _process_document_thread(self):
        """Process document in separate thread"""
        try:
            model_id = self.model_var.get()
            output_dir = self.output_path_var.get()
            
            self.root.after(0, lambda: self.log_message(f"Processing document with {model_id} model..."))
            
            # Process document with fallback
            csv_files, raw_result, actual_model = self._process_with_fallback(model_id, output_dir)
            
            self.results = {
                'csv_files': csv_files,
                'raw_result': raw_result,
                'output_dir': output_dir,
                'model_used': actual_model
            }
            
            # Update UI in main thread
            self.root.after(0, self._update_results_ui)
            
        except Exception as e:
            error_msg = f"Processing failed: {str(e)}"
            self.root.after(0, lambda: self.log_message(error_msg))
            self.root.after(0, lambda: messagebox.showerror("Processing Error", error_msg))
        finally:
            self.root.after(0, self._processing_complete)
    
    def _process_with_fallback(self, preferred_model, output_dir):
        """Process document with automatic fallback to working models"""
        try:
            # Try preferred model first
            csv_files, raw_result = self.service.analyze_tables(
                file_path=self.current_file,
                model_id=preferred_model,
                out_dir=output_dir
            )
            return csv_files, raw_result, preferred_model
            
        except Exception as e:
            error_str = str(e)
            
            if "ModelNotFound" in error_str and preferred_model == "prebuilt-document":
                # Fallback to prebuilt-layout
                self.root.after(0, lambda: self.log_message("prebuilt-document not available, falling back to prebuilt-layout"))
                
                csv_files, raw_result = self.service.analyze_tables(
                    file_path=self.current_file,
                    model_id="prebuilt-layout",
                    out_dir=output_dir
                )
                return csv_files, raw_result, "prebuilt-layout"
                
            elif "InvalidContent" in error_str:
                # Handle file format/corruption issues
                self.root.after(0, lambda: self.log_message("File format error detected"))
                raise Exception(f"File format issue: The file may be corrupted or in an unsupported format.\n\nDetails: {error_str}")
                
            elif "InvalidRequest" in error_str:
                # Handle request format issues
                raise Exception(f"Request format issue: Please check the file format and try again.\n\nDetails: {error_str}")
                
            else:
                # Re-raise other errors with more context
                raise Exception(f"Processing error: {error_str}")
    
    def _update_results_ui(self):
        """Update UI with processing results"""
        csv_files = self.results['csv_files']
        raw_result = self.results['raw_result']
        
        # Update summary
        tables_count = len(csv_files)
        pages_count = len(raw_result.get('pages', []))
        model_used = self.results['model_used']
        summary_text = f"✓ Processed: {tables_count} tables, {pages_count} pages (using {model_used})"
        self.results_summary.config(text=summary_text, style='Success.TLabel')
        
        # Update table combo
        table_options = [f"Table {i+1}" for i in range(tables_count)]
        self.table_combo['values'] = table_options
        if table_options:
            self.table_combo.set(table_options[0])
            self.on_table_selected(None)
        
        # Update raw data
        self.raw_text.delete(1.0, tk.END)
        self.raw_text.insert(1.0, json.dumps(raw_result, indent=2, ensure_ascii=False))
        
        # Update key-value pairs
        self.update_kv_display(raw_result)
        
        # Update KV tab title based on content
        self.update_kv_tab_title(raw_result)
        
        # Enable buttons
        self.open_output_btn.config(state='normal')
        
        self.log_message(f"Processing completed successfully!")
        self.log_message(f"Tables extracted: {tables_count}")
        self.log_message(f"Output saved to: {self.results['output_dir']}")
    
    def update_kv_display(self, raw_result):
        """Update key-value pairs display - show only extracted data"""
        # Clear existing items
        for item in self.kv_tree.get_children():
            self.kv_tree.delete(item)
        
        # Get formal key-value pairs from prebuilt-document model
        formal_kv_pairs = raw_result.get('keyValuePairs', [])
        
        # Get text-based key-value pairs from document content
        text_kv_pairs = self.extract_text_based_kv_pairs(raw_result)
        
        # Combine all extracted pairs
        all_pairs = {}
        
        # Add formal KV pairs first (higher priority)
        for kv in formal_kv_pairs:
            key = kv.get('key', {}).get('content', '').strip()
            value = kv.get('value', {}).get('content', '').strip() if kv.get('value') else ''
            if key and value:
                all_pairs[key] = value
        
        # Add text-based pairs (avoid duplicates)
        for key, value in text_kv_pairs.items():
            if key not in all_pairs and value:  # Only add if not already present
                all_pairs[key] = value
        
        # Display all extracted pairs (no headers, no messages, just data)
        for key, value in all_pairs.items():
            self.kv_tree.insert('', tk.END, values=(key, value))
        
        # Log extraction summary
        total_pairs = len(all_pairs)
        if total_pairs > 0:
            self.log_message(f"Extracted {total_pairs} key-value pairs from document")
        else:
            self.log_message("No key-value pairs found in document")
    
    def extract_text_based_kv_pairs(self, raw_result):
        """Extract only important key-value pairs for MTC certificates"""
        try:
            import re
            
            # Get all text content from pages
            all_text = ""
            pages = raw_result.get('pages', [])
            for page in pages:
                lines = page.get('lines', [])
                for line in lines:
                    content = line.get('content', '')
                    all_text += content + " "
            
            if not all_text.strip():
                return {}
            
            # Define only the important patterns for MTC certificates
            important_patterns = [
                # Heat Number variations
                (r'Heat\s*(?:No|Number|#)\s*:?\s*([A-Z0-9\-\.]+)', 'Heat Number'),
                
                # Work Order Number variations
                (r'(?:Work\s*Order|Works\s*Order)\s*(?:No|Number|#)\s*:?\s*([A-Z0-9\-\.]+)', 'Works Order No'),
                
                # Order Number variations
                (r'Order\s*(?:No|Number|#)\s*:?\s*([A-Z0-9\-\.]+)', 'Order No'),
                (r'PO[\-\s]*([A-Z0-9\-\.]+)', 'Order No'),
                
                # Customer Order Number
                (r'Customer\s*Order\s*(?:No|Number|#)\s*:?\s*([A-Z0-9\-\.]+)', 'Customer Order No'),
                
                # Test Certificate Number
                (r'(?:Test\s*)?Certificate\s*(?:No|Number|#)\s*:?\s*([A-Z0-9\-\.]+)', 'Test Certificate No'),
                (r'Cert\s*(?:No|Number|#)\s*:?\s*([A-Z0-9\-\.]+)', 'Test Certificate No'),
                
                # Date variations
                (r'(?:Date|Shipping\s*Date|Test\s*Date)\s*:?\s*([0-9]{1,2}[\.\/\-][0-9]{1,2}[\.\/\-][0-9]{2,4})', 'Date'),
                
                # Material
                (r'Material\s*:?\s*([A-Z0-9\-\./\s]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Material'),
                
                # Producer/Manufacturer
                (r'Producer\s*:?\s*([A-Za-z\s&\.]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Producer'),
                (r'Manufacturer\s*:?\s*([A-Za-z\s&\.]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Producer'),
                
                # Customer variations
                (r'Customer\s*:?\s*([A-Za-z\s&\.]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Customer'),
                (r'Customer\s*Name\s*:?\s*([A-Za-z\s&\.]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Customer Name'),
                
                # Address
                (r'Address\s*:?\s*([A-Za-z0-9\s,\.\-]+?)(?=\s+[A-Z][a-z]{2,}|\s*$)', 'Address'),
                (r'Customer\s*Address\s*:?\s*([A-Za-z0-9\s,\.\-]+?)(?=\s+[A-Z][a-z]{2,}|\s*$)', 'Address'),
                
                # Process of Manufacture
                (r'Process\s*(?:of\s*)?Manufacture\s*:?\s*([A-Za-z\s\-\.]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Process of Manufacture'),
                (r'Manufacturing\s*Process\s*:?\s*([A-Za-z\s\-\.]+?)(?=\s+[A-Z][a-z]|\s*$)', 'Process of Manufacture'),
            ]
            
            found_pairs = {}
            
            # Extract using specific patterns
            for pattern, key_name in important_patterns:
                matches = re.findall(pattern, all_text, re.MULTILINE | re.IGNORECASE)
                for match in matches:
                    value = match.strip()
                    # Clean up the value
                    value = re.sub(r'\s+', ' ', value)  # Normalize whitespace
                    value = re.sub(r'[,;\.]+$', '', value)  # Remove trailing punctuation
                    
                    if value and len(value) > 1 and len(value) < 150:
                        # Avoid duplicates - keep the first found value for each key
                        if key_name not in found_pairs:
                            found_pairs[key_name] = value
            
            # Additional generic pattern for any remaining important fields
            generic_important = [
                'heat number', 'works order', 'order no', 'customer order', 
                'certificate no', 'test certificate', 'material', 'producer', 
                'customer', 'address', 'process', 'manufacture'
            ]
            
            # Look for "Key: Value" patterns that match our important fields
            generic_pattern = r'([A-Za-z][A-Za-z\s\.\-#&]+?)\s*:\s*([^\n\r:]+?)(?=\s+[A-Za-z]|\s*$)'
            generic_matches = re.findall(generic_pattern, all_text, re.MULTILINE | re.IGNORECASE)
            
            for key, value in generic_matches:
                key = key.strip().lower()
                value = value.strip()
                
                # Check if this key contains any of our important terms
                for important_term in generic_important:
                    if important_term in key and len(value) > 1 and len(value) < 150:
                        # Create a clean key name
                        clean_key = key.title().replace('No ', 'No').replace('Number', 'No')
                        if clean_key not in found_pairs:
                            found_pairs[clean_key] = value
                        break
            
            return found_pairs
                    
        except Exception as e:
            self.log_message(f"Text extraction error: {e}")
            return {}
    
    def _is_important_field(self, key):
        """Check if a key is one of the important MTC fields"""
        key_lower = key.lower()
        important_keywords = [
            'heat', 'order', 'certificate', 'material', 'producer', 
            'customer', 'address', 'process', 'manufacture', 'date',
            'works', 'test'
        ]
        
        return any(keyword in key_lower for keyword in important_keywords)
    
    def update_kv_tab_title(self, raw_result):
        """Update the KV tab title based on available content"""
        try:
            # Count formal key-value pairs
            formal_kv_pairs = raw_result.get('keyValuePairs', [])
            
            # Count text-extracted pairs
            text_kv_pairs = self.extract_text_based_kv_pairs(raw_result)
            
            total_pairs = len(formal_kv_pairs) + len(text_kv_pairs)
            
            if total_pairs > 0:
                title = f"Key-Value Pairs ({total_pairs})"
            else:
                title = "Key-Value Pairs"
            
            # Update the tab title
            self.notebook.tab(2, text=title)  # KV tab is index 2
            
        except Exception as e:
            self.log_message(f"Failed to update KV tab title: {e}")
    
    def on_table_selected(self, event):
        """Handle table selection change"""
        if not self.results or not self.table_var.get():
            return
        
        try:
            table_index = int(self.table_var.get().split()[1]) - 1
            csv_file = self.results['csv_files'][table_index]
            
            # Load and display table
            df = pd.read_csv(csv_file, header=None)
            self.display_dataframe(df)
            
        except Exception as e:
            self.log_message(f"Error loading table: {str(e)}")
    
    def display_dataframe(self, df):
        """Display dataframe in treeview"""
        # Clear existing data
        for item in self.table_tree.get_children():
            self.table_tree.delete(item)
        
        # Configure columns
        columns = [f"Col_{i}" for i in range(len(df.columns))]
        self.table_tree['columns'] = columns
        self.table_tree['show'] = 'headings'
        
        # Set column headings and widths
        for col in columns:
            self.table_tree.heading(col, text=col)
            self.table_tree.column(col, width=100, minwidth=50)
        
        # Insert data
        for index, row in df.iterrows():
            values = [str(val) if pd.notna(val) else '' for val in row]
            self.table_tree.insert('', tk.END, values=values)
    
    def export_table_excel(self):
        """Export selected table to Excel"""
        if not self.results or not self.table_var.get():
            messagebox.showwarning("Warning", "No table selected")
            return
        
        try:
            table_index = int(self.table_var.get().split()[1]) - 1
            csv_file = self.results['csv_files'][table_index]
            
            # Ask for save location
            excel_file = filedialog.asksaveasfilename(
                title="Save table as Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if excel_file:
                df = pd.read_csv(csv_file, header=None)
                df.to_excel(excel_file, index=False, header=False)
                self.log_message(f"Table exported to: {excel_file}")
                messagebox.showinfo("Success", f"Table exported successfully!")
                
        except Exception as e:
            error_msg = f"Export failed: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Export Error", error_msg)
    
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        if self.results and 'output_dir' in self.results:
            output_dir = self.results['output_dir']
            if os.path.exists(output_dir):
                os.startfile(output_dir)  # Windows
            else:
                messagebox.showwarning("Warning", "Output directory not found")
        else:
            messagebox.showwarning("Warning", "No output directory available")
    
    def clear_results(self):
        """Clear all results and reset UI"""
        self.results = {}
        self.current_file = None
        self.file_path_var.set("")
        
        # Clear displays
        self.results_summary.config(text="No document processed yet", style='TLabel')
        self.table_combo['values'] = []
        self.table_combo.set("")
        
        # Clear table display
        for item in self.table_tree.get_children():
            self.table_tree.delete(item)
        self.table_tree['columns'] = []
        
        # Clear raw data
        self.raw_text.delete(1.0, tk.END)
        
        # Clear KV pairs
        for item in self.kv_tree.get_children():
            self.kv_tree.delete(item)
        
        # Disable buttons
        self.open_output_btn.config(state='disabled')
        
        self.log_message("Results cleared")
    
    def _processing_complete(self):
        """Re-enable UI after processing"""
        self.progress.stop()
        self.process_btn.config(state='normal')
    
    def log_message(self, message):
        """Add message to status log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.status_text.insert(tk.END, log_entry)
        self.status_text.see(tk.END)
        self.root.update_idletasks()


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = DocumentIntelligenceGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()