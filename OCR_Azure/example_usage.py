"""
Example usage of the Document Intelligence Service
"""

from document_intelligence_service import DocumentIntelligenceService
import os

def main():
    # Initialize the service (reads from .env file)
    service = DocumentIntelligenceService()
    
    # List available models
    print("=== Available Models ===")
    try:
        models = service.list_available_models()
        print(f"Available models: {models}")
    except Exception as e:
        print(f"Could not list models: {e}")
    print()
    
    # Example 1: Extract tables only (fastest)
    print("=== Example 1: Extract Tables Only ===")
    try:
        csv_files, raw_result = service.analyze_tables(
            file_path="sample_document.pdf",  # Replace with your file
            model_id="prebuilt-layout",       # Tables only
            out_dir="tables_output"
        )
        print(f"Extracted {len(csv_files)} tables:")
        for csv_file in csv_files:
            print(f"  - {csv_file}")
    except Exception as e:
        print(f"Error: {e}")
    
    # Example 2: Extract tables with detailed analysis
    print("\n=== Example 2: Detailed Table Analysis ===")
    try:
        csv_files, raw_result = service.analyze_tables(
            file_path="sample_document.pdf",  # Replace with your file
            model_id="prebuilt-layout",       # Use working model
            out_dir="detailed_output"
        )
        
        print(f"Extracted {len(csv_files)} tables")
        
        # Show table details
        tables = raw_result.get("tables", [])
        for i, table in enumerate(tables, 1):
            print(f"  Table {i}: {table.get('rowCount', 0)} rows x {table.get('columnCount', 0)} columns")
            
        # Show some extracted text content
        pages = raw_result.get("pages", [])
        if pages:
            print(f"Document has {len(pages)} pages")
            
    except Exception as e:
        print(f"Error: {e}")
    
    # Example 3: Inspect extracted table content
    print("\n=== Example 3: Table Content Inspection ===")
    try:
        csv_files, raw_result = service.analyze_tables(
            file_path="sample_document.pdf",
            model_id="prebuilt-layout",
            out_dir="inspection_output"
        )
        
        print("Processing Results:")
        print(f"  Tables extracted: {len(csv_files)}")
        
        # Show content of first table if available
        if csv_files:
            import pandas as pd
            first_table = pd.read_csv(csv_files[0], header=None)
            print(f"  First table shape: {first_table.shape}")
            print("  First few rows:")
            print(first_table.head().to_string(index=False))
            
        # Show confidence scores for cells
        tables = raw_result.get("tables", [])
        if tables:
            first_table = tables[0]
            cells = first_table.get("cells", [])
            if cells:
                avg_confidence = sum(cell.get("confidence", 0) for cell in cells) / len(cells)
                print(f"  Average confidence: {avg_confidence:.2f}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()