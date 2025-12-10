"""
Batch processor for handling multiple documents with Azure AI Document Intelligence
Includes queuing, error handling, and progress tracking
"""

import os
import json
from typing import List, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import logging
from document_intelligence_service import DocumentIntelligenceService

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class BatchProcessor:
    """Batch processor for document analysis with queuing and error handling"""
    
    def __init__(self, service: DocumentIntelligenceService, max_workers: int = 3):
        """
        Initialize batch processor
        
        Args:
            service: DocumentIntelligenceService instance
            max_workers: Maximum number of concurrent processing threads
        """
        self.service = service
        self.max_workers = max_workers
        
    def process_batch(self, file_paths: List[str], output_base_dir: str = "batch_output",
                     model_id: str = "prebuilt-document") -> Dict:
        """
        Process multiple documents in batch
        
        Args:
            file_paths: List of file paths to process
            output_base_dir: Base directory for outputs
            model_id: Model to use for processing
            
        Returns:
            Dictionary with batch processing results
        """
        os.makedirs(output_base_dir, exist_ok=True)
        
        batch_results = {
            "started_at": datetime.now().isoformat(),
            "total_files": len(file_paths),
            "successful": [],
            "failed": [],
            "results": {}
        }
        
        # Process files concurrently
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all jobs
            future_to_file = {
                executor.submit(self._process_single_file, file_path, output_base_dir, model_id): file_path
                for file_path in file_paths
            }
            
            # Collect results as they complete
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                try:
                    result = future.result()
                    batch_results["successful"].append(file_path)
                    batch_results["results"][file_path] = result
                    logger.info(f"Successfully processed: {file_path}")
                except Exception as e:
                    batch_results["failed"].append({"file": file_path, "error": str(e)})
                    logger.error(f"Failed to process {file_path}: {e}")
        
        batch_results["completed_at"] = datetime.now().isoformat()
        batch_results["success_rate"] = len(batch_results["successful"]) / len(file_paths) * 100
        
        # Save batch summary
        summary_path = os.path.join(output_base_dir, "batch_summary.json")
        with open(summary_path, "w", encoding="utf-8") as f:
            json.dump(batch_results, f, ensure_ascii=False, indent=2)
        
        return batch_results
    
    def _process_single_file(self, file_path: str, output_base_dir: str, model_id: str) -> Dict:
        """
        Process a single file
        
        Args:
            file_path: Path to the file
            output_base_dir: Base output directory
            model_id: Model to use
            
        Returns:
            Processing result for the file
        """
        # Create file-specific output directory
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        file_output_dir = os.path.join(output_base_dir, file_name)
        
        # Process the document
        csv_files, raw_result = self.service.analyze_tables(
            file_path=file_path,
            model_id=model_id,
            out_dir=file_output_dir
        )
        
        # Extract key-value pairs if using prebuilt-document
        kv_pairs = {}
        if model_id == "prebuilt-document":
            for kv in raw_result.get("keyValuePairs", []):
                key = kv.get("key", {}).get("content", "")
                value = kv.get("value", {}).get("content", "") if kv.get("value") else ""
                if key:
                    kv_pairs[key] = value
        
        return {
            "file_path": file_path,
            "output_dir": file_output_dir,
            "tables_count": len(csv_files),
            "csv_files": csv_files,
            "kv_pairs_count": len(kv_pairs),
            "pages_processed": len(raw_result.get("pages", [])),
            "model_used": model_id
        }
    
    def validate_batch(self, file_paths: List[str], required_keys: List[str],
                      manual_inputs: Dict[str, Dict[str, str]]) -> Dict:
        """
        Validate multiple documents against manual inputs
        
        Args:
            file_paths: List of file paths to validate
            required_keys: List of required keys for validation
            manual_inputs: Dictionary mapping file paths to manual input dictionaries
            
        Returns:
            Batch validation results
        """
        validation_results = {
            "started_at": datetime.now().isoformat(),
            "total_files": len(file_paths),
            "passed": [],
            "failed": [],
            "results": {}
        }
        
        for file_path in file_paths:
            try:
                manual_input = manual_inputs.get(file_path, {})
                is_valid, mismatched_keys = self.service.validate_document(
                    file_path, required_keys, manual_input
                )
                
                result = {
                    "is_valid": is_valid,
                    "mismatched_keys": mismatched_keys,
                    "manual_input": manual_input
                }
                
                validation_results["results"][file_path] = result
                
                if is_valid:
                    validation_results["passed"].append(file_path)
                else:
                    validation_results["failed"].append(file_path)
                    
            except Exception as e:
                validation_results["failed"].append(file_path)
                validation_results["results"][file_path] = {"error": str(e)}
        
        validation_results["completed_at"] = datetime.now().isoformat()
        validation_results["pass_rate"] = len(validation_results["passed"]) / len(file_paths) * 100
        
        return validation_results


def main():
    """Example usage of batch processor"""
    
    # Initialize service and batch processor
    service = DocumentIntelligenceService()
    processor = BatchProcessor(service, max_workers=2)
    
    # Example file paths (replace with your actual files)
    file_paths = [
        "document1.pdf",
        "document2.pdf", 
        "document3.pdf"
    ]
    
    # Process batch
    print("Starting batch processing...")
    results = processor.process_batch(
        file_paths=file_paths,
        output_base_dir="batch_results",
        model_id="prebuilt-document"
    )
    
    print(f"Batch processing completed:")
    print(f"  Success rate: {results['success_rate']:.1f}%")
    print(f"  Successful: {len(results['successful'])}")
    print(f"  Failed: {len(results['failed'])}")
    
    # Example validation
    if results['successful']:
        print("\nRunning validation...")
        required_keys = ["Invoice Number", "Date", "Total"]
        manual_inputs = {
            "document1.pdf": {"Invoice Number": "INV-001", "Date": "2024-01-01", "Total": "$100"},
            "document2.pdf": {"Invoice Number": "INV-002", "Date": "2024-01-02", "Total": "$200"},
        }
        
        validation_results = processor.validate_batch(
            file_paths=results['successful'][:2],  # Validate first 2 successful files
            required_keys=required_keys,
            manual_inputs=manual_inputs
        )
        
        print(f"Validation completed:")
        print(f"  Pass rate: {validation_results['pass_rate']:.1f}%")


if __name__ == "__main__":
    main()