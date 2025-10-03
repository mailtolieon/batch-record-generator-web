import os
import json
from pathlib import Path
from typing import Dict, Any

class WebConfig:
    """Configuration management for web application"""
    
    def __init__(self, excel_file: str = None, template_file: str = None, sheet_name: str = "5_Arc_List"):
        self.excel_file = excel_file or "data/Batch_Local_2025_09_07.xlsx"
        self.template_file = template_file or "data/Batch Record Register_template.docx"
        self.sheet_name = sheet_name
        self.output_folder = "generated"
        
        # Default column mappings
        self.column_mappings = {
            "product": ["product_name", "productname", "product", "name", "product_description"],
            "batch_no": ["batch_no", "batch_no.", "batchno", "batch", "batch_number"],
            "mfg_date": ["mfg_date", "manufacturing_date", "mfgdate", "manufacture_date"],
            "expiry_date": ["expiry_date", "exp_date", "expirydate", "expiration_date"],
            "yield": ["total_batch_yield_%", "total_batch_yield", "batch_yield_%", "yield_%", "yield", "total_yield"],
            "accountability": ["total_batch_accountability_%", "total_batch_accountability", "batch_accountability_%", "accountability_%", "accountability"],
            "location": ["location_rack_shelf", "location", "rack_shelf", "storage_location", "location_rack", "rack"],
            "remarks": ["remarks", "remark", "comments", "note", "notes"],
            "sent_to_doc": ["sent_to_document_room_by_date", "sent_to_document_room", "document_room_date", "sent_to_doc", "doc_room_date"]
        }
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert configuration to dictionary"""
        return {
            'excel_file': self.excel_file,
            'template_file': self.template_file,
            'sheet_name': self.sheet_name,
            'output_folder': self.output_folder,
            'column_mappings': self.column_mappings
        }
    
    def save(self, filepath: str = "web_config.json"):
        """Save configuration to file"""
        with open(filepath, 'w') as f:
            json.dump(self.to_dict(), f, indent=4)
    
    @classmethod
    def load(cls, filepath: str = "web_config.json"):
        """Load configuration from file"""
        if os.path.exists(filepath):
            with open(filepath, 'r') as f:
                data = json.load(f)
            config = cls()
            config.excel_file = data.get('excel_file', config.excel_file)
            config.template_file = data.get('template_file', config.template_file)
            config.sheet_name = data.get('sheet_name', config.sheet_name)
            config.output_folder = data.get('output_folder', config.output_folder)
            config.column_mappings = data.get('column_mappings', config.column_mappings)
            return config
        return cls()