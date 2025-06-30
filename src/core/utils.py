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
    def normalize_scid(scid, ignore_keywords=None):
        """
        Normalize SCID format with optional ignore keywords
        
        Args:
            scid (str): Raw SCID string
            ignore_keywords (list, optional): Keywords to ignore when normalizing
            
        Returns:
            str: Normalized SCID
        """
        if not scid:
            return scid
        
        scid_str = str(scid).strip()
        
        # Remove leading apostrophe that Excel sometimes adds to preserve text formatting
        if scid_str.startswith("'"):
            scid_str = scid_str[1:]
        
        # Apply ignore keywords if provided
        if ignore_keywords:
            scid_cleaned = scid_str
            for keyword in ignore_keywords:
                if keyword and keyword.strip():  # Only process non-empty keywords
                    # Create a case-insensitive pattern to match the keyword
                    # Use word boundaries to avoid partial matches
                    pattern = r'\b' + re.escape(keyword.strip()) + r'\b'
                    scid_cleaned = re.sub(pattern, '', scid_cleaned, flags=re.IGNORECASE).strip()
            
            # Remove extra whitespace that might result from keyword removal
            scid_cleaned = re.sub(r'\s+', ' ', scid_cleaned).strip()
            scid_str = scid_cleaned
        
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
        """Get the base directory for the application (exe or script location)"""
        if getattr(sys, 'frozen', False):
            # Running as a PyInstaller bundle
            return Path(sys.executable).parent
        else:
            # Running as a script - use the main script's directory
            return Path(sys.argv[0]).resolve().parent
    
    @staticmethod
    def parse_height_format(height_str):
        if not height_str or str(height_str).strip() == '':
            return ''
            
        s = str(height_str).strip()
        
        # Handle various height formats
        # Pattern 1: 5'-10" or 5'10"
        height_pattern = re.compile(r"(\d+)'-?(\d+)\"")
        m = height_pattern.match(s)
        if m:
            return f"{int(m.group(1))}' {int(m.group(2))}\""
        
        # Pattern 2: 5' 10" (with space)
        m = re.match(r"(\d+)'\s+(\d+)\"?", s)
        if m:
            return f"{int(m.group(1))}' {int(m.group(2))}\""
        
        # Pattern 3: Just feet with apostrophe (5')
        m = re.match(r"(\d+)'", s)
        if m:
            return f"{int(m.group(1))}' 0\""
        
        # Pattern 4: Decimal feet (5.5 -> 5' 6")
        m = re.match(r"(\d+)\.(\d+)", s)
        if m:
            feet = int(m.group(1))
            decimal_part = float(f"0.{m.group(2)}")
            inches = round(decimal_part * 12)
            return f"{feet}' {inches}\""
        
        # Pattern 5: Just a number (assume feet)
        m = re.match(r"(\d+)$", s)
        if m:
            return f"{int(m.group(1))}' 0\""
        
        logging.debug(f"Could not parse height format: '{height_str}'")
        return ''
    
    @staticmethod
    def parse_height_decimal(height_str):
        if not height_str or str(height_str).strip() == '':
            return None
            
        try:
            s = str(height_str).strip()
            
            # Pattern 1: 5'-10" or 5'10"
            height_pattern = re.compile(r"(\d+)'-?(\d+)\"")
            m = height_pattern.match(s)
            if m:
                feet = int(m.group(1))
                inches = int(m.group(2))
                return round(feet + inches / 12, 2)
            
            # Pattern 2: 5' 10" (with space)
            m = re.match(r"(\d+)'\s*(\d+)?\"?", s)
            if m:
                feet = int(m.group(1))
                inches = int(m.group(2)) if m.group(2) else 0
                return round(feet + inches / 12, 2)
            
            # Pattern 3: Decimal number with explicit context
            # If it contains a decimal point and is reasonable for feet (< 50), treat as feet
            # Otherwise, treat as inches
            m = re.match(r"(\d+\.?\d*)", s)
            if m:
                value = float(m.group(1))
                # If it's a decimal and reasonably small, assume it's feet
                if '.' in s and value < 50:
                    return round(value, 2)
                else:
                    # Pure integers or large numbers are likely inches
                    return round(value / 12, 2)
                
        except (ValueError, TypeError) as e:
            logging.debug(f"Error parsing height decimal: '{height_str}' - {e}")
        
        return None
    
    @staticmethod
    def inches_to_feet_format(inches):
        try:
            # Handle both string and numeric inputs
            if isinstance(inches, str):
                inches_str = inches.strip()
                
                # If it's already in feet-inches format, parse and reformat
                if "'" in inches_str or "\"" in inches_str:
                    # Try to parse existing format first
                    parsed_decimal = Utils.parse_height_decimal(inches_str)
                    if parsed_decimal is not None:
                        # Convert back to total inches and then format
                        total_inches = round(parsed_decimal * 12)
                        feet = total_inches // 12
                        remaining_inches = total_inches % 12
                        return f"{int(feet)}' {int(remaining_inches)}\""
                    else:
                        return ''
                
                # Remove any quote marks that might be present and treat as inches
                inches_clean = inches_str.replace('"', '').replace('″', '')
                try:
                    # Check if this looks like decimal feet (small number with decimal)
                    if '.' in inches_clean:
                        value = float(inches_clean)
                        if value < 50:  # Reasonable range for feet
                            # Treat as decimal feet, convert to inches
                            total_inches = round(value * 12)
                        else:
                            # Large decimal, treat as inches
                            total_inches = round(value)
                    else:
                        # Pure integer, treat as inches
                        total_inches = float(inches_clean)
                except ValueError:
                    return ''
            else:
                total_inches = float(inches)
            
            # Handle negative values
            if total_inches < 0:
                return ''
            
            # Round to nearest inch for display purposes
            total_inches = round(total_inches)
            feet = total_inches // 12
            remaining_inches = total_inches % 12
            return f"{int(feet)}' {int(remaining_inches)}\""
        except (ValueError, TypeError) as e:
            logging.debug(f"Error converting inches to feet format: {inches} - {e}")
            return ''
    
    @staticmethod
    def decimal_feet_to_feet_format(decimal_feet):
        """Convert decimal feet to feet'inches" format"""
        try:
            # Convert to float and handle None/empty values
            if decimal_feet is None or str(decimal_feet).strip() == '':
                return None
                
            # Round to 2 decimal places
            decimal_feet = round(float(decimal_feet), 2)
            
            # Calculate feet and inches
            feet = int(decimal_feet)
            inches = round((decimal_feet - feet) * 12)
            
            # Handle case where inches rounds to 12
            if inches == 12:
                feet += 1
                inches = 0
            
            # Format as feet'inches"
            return f"{feet}'{inches}\""
            
        except (ValueError, TypeError) as e:
            logging.warning(f"Could not convert decimal feet to format: {decimal_feet}")
            return None