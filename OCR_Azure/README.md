# Azure AI Document Intelligence Table Extractor

A Python service for extracting tables and key-value pairs from documents using Azure AI Document Intelligence (formerly Form Recognizer).

## Features

- **Table Extraction**: Extract structured tables from PDFs, images, and other document formats
- **Key-Value Pair Extraction**: Extract form fields and document metadata
- **Document Validation**: Compare extracted data with manual inputs for quality control
- **Batch Processing**: Process multiple documents concurrently with error handling
- **Multiple Output Formats**: Export tables to CSV, save raw JSON results
- **Flexible Models**: Support for both `prebuilt-layout` (tables only) and `prebuilt-document` (tables + KV pairs)

## Setup

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure Azure credentials**:
   - Copy `.env.example` to `.env`
   - Add your Azure Document Intelligence endpoint and key:
   ```
   AZURE_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
   AZURE_KEY=your-api-key-here
   ```

3. **Get Azure Document Intelligence resource**:
   - Create a Document Intelligence resource in Azure Portal
   - Copy the endpoint URL and one of the keys

## Usage

### GUI Application (Recommended)

Launch the graphical interface for easy document processing:

```bash
python gui_app.py
```

Or on Windows, double-click `launch_gui.bat`

**GUI Features:**
- **File Selection**: Browse and select PDF/image files
- **Model Selection**: Choose between layout-only or document models
- **Real-time Processing**: Progress tracking with status updates
- **Table Viewer**: Interactive table display with column sorting
- **Export Options**: Export tables to Excel format
- **Raw Data View**: Inspect complete JSON responses
- **Key-Value Pairs**: View extracted form fields
- **Output Management**: Direct access to output folders

### Command Line Usage

### Basic Table Extraction

```python
from document_intelligence_service import DocumentIntelligenceService

# Initialize service
service = DocumentIntelligenceService()

# Extract tables only (fastest)
csv_files, raw_result = service.analyze_tables(
    file_path="invoice.pdf",
    model_id="prebuilt-layout",  # Tables only
    out_dir="output"
)

print(f"Extracted {len(csv_files)} tables")
```

### Extract Tables + Key-Value Pairs

```python
# Extract both tables and key-value pairs
csv_files, raw_result = service.analyze_tables(
    file_path="invoice.pdf", 
    model_id="prebuilt-document",  # Tables + KV pairs
    out_dir="output"
)

# Get key-value pairs
kv_pairs = service.extract_key_value_pairs("invoice.pdf")
print("Key-value pairs:", kv_pairs)
```

### Document Validation

```python
# Validate extracted data against manual input
required_keys = ["Invoice Number", "Date", "Total Amount"]
manual_input = {
    "Invoice Number": "INV-2024-001",
    "Date": "2024-01-15", 
    "Total Amount": "$1,250.00"
}

is_valid, mismatched_keys = service.validate_document(
    "invoice.pdf", required_keys, manual_input
)

if not is_valid:
    print(f"Validation failed for keys: {mismatched_keys}")
```

### Batch Processing

```python
from batch_processor import BatchProcessor

# Process multiple documents
processor = BatchProcessor(service, max_workers=3)
results = processor.process_batch(
    file_paths=["doc1.pdf", "doc2.pdf", "doc3.pdf"],
    output_base_dir="batch_output"
)

print(f"Success rate: {results['success_rate']:.1f}%")
```

## Supported File Formats

- **PDF**: Single and multi-page documents
- **Images**: JPG, PNG, TIFF
- **Quality recommendations**: 300 DPI, deskewed, good contrast

## Models

### prebuilt-layout
- **Use case**: Extract tables and document structure only
- **Speed**: Fastest option
- **Output**: Tables, text, document structure

### prebuilt-document  
- **Use case**: Extract tables AND key-value pairs in one call
- **Speed**: Slightly slower but more comprehensive
- **Output**: Tables, key-value pairs, selection marks, text

## Output Structure

### Tables (CSV)
Each table is exported as a separate CSV file with proper row/column structure.

### Raw JSON
Complete API response saved for audit and advanced processing:
```json
{
  "tables": [{
    "rowCount": 5,
    "columnCount": 4, 
    "cells": [
      {
        "rowIndex": 0,
        "columnIndex": 0,
        "content": "Item Code",
        "confidence": 0.99
      }
    ]
  }]
}
```

### Key-Value Pairs
```json
{
  "Invoice Number": "INV-2024-001",
  "Date": "January 15, 2024",
  "Total Amount": "$1,250.00"
}
```

## Error Handling

The service includes comprehensive error handling for:
- Invalid file formats
- Network connectivity issues
- API rate limits
- Malformed documents
- Missing credentials

## Best Practices

1. **Document Quality**: Use 300 DPI scans, ensure documents are deskewed
2. **Batch Processing**: Use concurrent processing for multiple documents
3. **Validation**: Always validate critical fields against manual input
4. **Confidence Filtering**: Check confidence scores for extracted data
5. **Cost Optimization**: Use `prebuilt-layout` if you only need tables

## API Costs

- Pricing is per page processed
- Multi-page PDFs count as multiple pages
- Consider filtering results client-side rather than multiple API calls

## Examples

Run the example scripts:

```bash
# Basic usage examples
python example_usage.py

# Batch processing example  
python batch_processor.py
```

## Troubleshooting

1. **Authentication errors**: Verify your endpoint URL and API key
2. **File format issues**: Ensure files are in supported formats (PDF, JPG, PNG, TIFF)
3. **Empty results**: Check document quality and try different models
4. **Rate limiting**: Implement retry logic with exponential backoff