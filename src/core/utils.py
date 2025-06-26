import os
import sys
import re
import json
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


class Utils:
    """Utility functions shared across the application"""
    
    @staticmethod
    def normalize_scid(scid):
        if not scid:
            return scid
        
        scid_str = str(scid).strip()
        
        # Remove leading apostrophe that Excel sometimes adds to preserve text formatting
        if scid_str.startswith("'"):
            scid_str = scid_str[1:]
        
        # Handle simple numeric SCIDs with optional letters (like "001A" -> "1A")
        match = re.match(r'^0*(\d+)([A-Za-z]*)$', scid_str)
        if match:
            numeric_part = str(int(match.group(1)))
            letter_part = match.group(2).upper()
            return numeric_part + letter_part
        
        # Handle complex SCIDs with spaces and mixed alphanumeric (like "118 MISM013")
        # Split by spaces and normalize each numeric part while preserving structure
        parts = scid_str.split()
        normalized_parts = []
        
        for part in parts:
            # Try to normalize numeric parts with leading zeros
            if part.isdigit():
                normalized_parts.append(str(int(part)))
            else:
                # For mixed alphanumeric parts, normalize leading zeros in numeric portions
                # Handle patterns like "MISM013" -> "MISM13"
                part_match = re.match(r'^([A-Za-z]*)0*(\d+)([A-Za-z]*)$', part)
                if part_match:
                    prefix = part_match.group(1).upper()
                    numeric = str(int(part_match.group(2)))
                    suffix = part_match.group(3).upper()
                    normalized_parts.append(prefix + numeric + suffix)
                else:
                    normalized_parts.append(part.upper())
        
        return ' '.join(normalized_parts)
    
    @staticmethod
    def extract_numeric_part(scid):
        """Extract numeric part from SCID for sorting purposes"""
        match = re.match(r'(\d+)([A-Za-z]*)', str(scid))
        if match:
            num = int(match.group(1))
            alpha = match.group(2) or ''
            return (num, alpha)
        return (float('inf'), '')
    
    @staticmethod
    def filter_valid_nodes(nodes_df):
        """Filter nodes to include only valid poles and references (excluding underground)"""
        return nodes_df[
            (nodes_df['node_type'].str.strip().str.lower().isin(['pole', 'reference'])) &
            (~nodes_df['pole_status'].str.strip().str.lower().eq('underground'))
        ]
    
    @staticmethod
    def get_base_directory():
        """Get the base directory for the application"""
        # If we're running from a script, use the script's directory
        if hasattr(sys, '_MEIPASS'):
            # Running as a PyInstaller bundle
            return Path(sys._MEIPASS)
        else:
            # Running as a script - use the project root
            current_file = Path(__file__)
            # Go up from src/core/utils.py to the project root
            return current_file.parent.parent.parent
    
    @staticmethod
    def parse_height_format(height_str):
        s = str(height_str).strip()
        height_pattern = re.compile(r"(\d+)'-(\d+)\"")
        
        m = height_pattern.match(s)
        if m:
            return f"{int(m.group(1))}' {int(m.group(2))}\""
        
        m = re.match(r"(\d+)[^\d]+(\d+)", s)
        if m:
            return f"{int(m.group(1))}' {int(m.group(2))}\""
        
        m = re.match(r"(\d+)", s)
        if m:
            return f"{int(m.group(1))}' 0\""
        
        return ''
    
    @staticmethod
    def parse_height_decimal(height_str):
        try:
            s = str(height_str).strip()
            height_pattern = re.compile(r"(\d+)'-(\d+)\"")
            
            m = height_pattern.match(s)
            if m:
                feet = int(m.group(1))
                inches = int(m.group(2))
                return round(feet + inches / 12, 2)
            
            m = re.match(r"(\d+)'\s*(\d+)?\"?", s)
            if m:
                feet = int(m.group(1))
                inches = int(m.group(2)) if m.group(2) else 0
                return round(feet + inches / 12, 2)
        except:
            pass
        return None
    
    @staticmethod
    def inches_to_feet_format(inches):
        try:
            total_inches = int(float(inches))
            feet = total_inches // 12
            remaining_inches = total_inches % 12
            return f"{feet}' {remaining_inches}\""
        except:
            return ''