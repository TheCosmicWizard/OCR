"""
Azure AI Document Intelligence Service for Table Extraction
Supports both prebuilt-layout (tables only) and prebuilt-document (tables + key-value pairs)
"""

from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
import os
import json
from typing import List, Dict, Optional, Tuple
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


class DocumentIntelligenceService:
    """Service for extracting tables and key-value pairs from documents using Azure AI Document Intelligence"""
    
    def __init__(self, endpoint: Optional[str] = None, key: Optional[str] = None):
        """
        Initialize the Document Intelligence service
        
        Args:
            endpoint: Azure Document Intelligence endpoint (defaults to env var AZURE_ENDPOINT)
            key: Azure Document Intelligence key (defaults to env var AZURE_KEY)
        """
        self.endpoint = endpoint or os.getenv("AZURE_ENDPOINT")
        self.key = key or os.getenv("AZURE_KEY")
        
        if not self.endpoint or not self.key:
            raise ValueError("Azure endpoint and key must be provided via parameters or environment variables")
        
        self.client = DocumentIntelligenceClient(self.endpoint, AzureKeyCredential(self.key))
    
    def list_available_models(self) -> List[str]:
        """
        List available models in your Document Intelligence resource
        
        Returns:
            List of available model IDs
        """
        try:
            # Try to get models using the correct method name
            models = self.client.get_models()
            return [model.model_id for model in models]
        except AttributeError:
            try:
                # Alternative method name
                models = self.client.list_models()
                return [model.model_id for model in models]
            except Exception:
                pass
        except Exception as e:
            print(f"Could not list models: {e}")
        
        # Return known working models
        return ["prebuilt-layout"]
    
    def analyze_tables(self, file_path: str, model_id: str = "prebuilt-layout", 
                      out_dir: str = "tables_out") -> Tuple[List[str], Dict]:
        """
        Analyze document and extract tables to CSV files
        
        Args:
            file_path: Path to the document file (PDF, TIFF, JPG, PNG)
            model_id: Model to use ("prebuilt-layout" for tables only, "prebuilt-document" for tables + KV)
            out_dir: Output directory for CSV files and raw JSON
            
        Returns:
            Tuple of (list of CSV file paths, raw analysis result dict)
        """
        os.makedirs(out_dir, exist_ok=True)
        
        # Analyze document
        with open(file_path, "rb") as f:
            document_content = f.read()
            
        poller = self.client.begin_analyze_document(
            model_id=model_id,
            body=document_content,
            content_type="application/octet-stream"
        )
        result = poller.result()
        
        # Save raw JSON result
        result_dict = result.as_dict()
        raw_json_path = os.path.join(out_dir, "raw_result.json")
        with open(raw_json_path, "w", encoding="utf-8") as jf:
            json.dump(result_dict, jf, ensure_ascii=False, indent=2)
        
        # Extract and export tables
        tables = result_dict.get("tables", [])
        csv_files = []
        
        for t_idx, table in enumerate(tables, start=1):
            csv_path = self._export_table_to_csv(table, out_dir, t_idx)
            csv_files.append(csv_path)
            
        return csv_files, result_dict
    
    def _export_table_to_csv(self, table: Dict, out_dir: str, table_index: int) -> str:
        """
        Export a single table to CSV format
        
        Args:
            table: Table data from Document Intelligence result
            out_dir: Output directory
            table_index: Index of the table for filename
            
        Returns:
            Path to the created CSV file
        """
        row_count = table.get("rowCount", 0)
        col_count = table.get("columnCount", 0)
        
        # Initialize empty grid
        grid = [["" for _ in range(col_count)] for _ in range(row_count)]
        
        # Fill grid with cell content
        cells = table.get("cells", [])
        for cell in cells:
            row_idx = cell.get("rowIndex", 0)
            col_idx = cell.get("columnIndex", 0)
            content = cell.get("content", "")
            
            # Ensure indices are within bounds
            if 0 <= row_idx < row_count and 0 <= col_idx < col_count:
                grid[row_idx][col_idx] = content
        
        # Create DataFrame and save to CSV
        df = pd.DataFrame(grid)
        csv_path = os.path.join(out_dir, f"table_{table_index}.csv")
        df.to_csv(csv_path, index=False, header=False)
        
        return csv_path
    
    def extract_key_value_pairs(self, file_path: str, out_dir: str = "kv_out", 
                               model_id: str = "prebuilt-document") -> Dict[str, str]:
        """
        Extract key-value pairs from document using specified model
        
        Args:
            file_path: Path to the document file
            out_dir: Output directory for results
            model_id: Model to use for KV extraction
            
        Returns:
            Dictionary of key-value pairs
        """
        os.makedirs(out_dir, exist_ok=True)
        
        # Use specified model for KV extraction
        with open(file_path, "rb") as f:
            document_content = f.read()
            
        try:
            poller = self.client.begin_analyze_document(
                model_id=model_id,
                body=document_content,
                content_type="application/octet-stream"
            )
            result = poller.result()
        except Exception as e:
            print(f"Model {model_id} failed, trying prebuilt-layout: {e}")
            # Fallback to prebuilt-layout
            poller = self.client.begin_analyze_document(
                model_id="prebuilt-layout",
                body=document_content,
                content_type="application/octet-stream"
            )
            result = poller.result()
        
        result_dict = result.as_dict()
        
        # Extract key-value pairs
        kv_pairs = {}
        for kv in result_dict.get("keyValuePairs", []):
            key = kv.get("key", {}).get("content", "")
            value = kv.get("value", {}).get("content", "") if kv.get("value") else ""
            if key:
                kv_pairs[key] = value
        
        # Save KV pairs to JSON
        kv_json_path = os.path.join(out_dir, "key_value_pairs.json")
        with open(kv_json_path, "w", encoding="utf-8") as jf:
            json.dump(kv_pairs, jf, ensure_ascii=False, indent=2)
        
        return kv_pairs
    
    def validate_document(self, file_path: str, required_keys: List[str], 
                         manual_input: Dict[str, str]) -> Tuple[bool, List[str]]:
        """
        Validate document by matching extracted key-value pairs with manual input
        
        Args:
            file_path: Path to the document file
            required_keys: List of required keys to validate
            manual_input: Dictionary of manually entered key-value pairs
            
        Returns:
            Tuple of (is_valid, list of mismatched keys)
        """
        extracted_kv = self.extract_key_value_pairs(file_path)
        
        mismatched_keys = []
        for key in required_keys:
            extracted_value = extracted_kv.get(key, "").strip().lower()
            manual_value = manual_input.get(key, "").strip().lower()
            
            if extracted_value != manual_value:
                mismatched_keys.append(key)
        
        is_valid = len(mismatched_keys) == 0
        return is_valid, mismatched_keys
    
    def process_document_complete(self, file_path: str, required_keys: Optional[List[str]] = None,
                                 manual_input: Optional[Dict[str, str]] = None,
                                 out_dir: str = "output") -> Dict:
        """
        Complete document processing: extract tables, KV pairs, and validate
        
        Args:
            file_path: Path to the document file
            required_keys: List of required keys for validation
            manual_input: Manual input for validation
            out_dir: Output directory
            
        Returns:
            Dictionary with processing results
        """
        results = {
            "file_path": file_path,
            "tables": [],
            "key_value_pairs": {},
            "validation": {"is_valid": True, "mismatched_keys": []}
        }
        
        try:
            # Extract tables and KV pairs using prebuilt-document model
            csv_files, raw_result = self.analyze_tables(
                file_path, 
                model_id="prebuilt-document", 
                out_dir=out_dir
            )
            results["tables"] = csv_files
            
            # Extract KV pairs from the same result
            kv_pairs = {}
            for kv in raw_result.get("keyValuePairs", []):
                key = kv.get("key", {}).get("content", "")
                value = kv.get("value", {}).get("content", "") if kv.get("value") else ""
                if key:
                    kv_pairs[key] = value
            results["key_value_pairs"] = kv_pairs
            
            # Validate if required
            if required_keys and manual_input:
                is_valid, mismatched_keys = self.validate_document(
                    file_path, required_keys, manual_input
                )
                results["validation"] = {
                    "is_valid": is_valid,
                    "mismatched_keys": mismatched_keys
                }
            
        except Exception as e:
            results["error"] = str(e)
            
        return results