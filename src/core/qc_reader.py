import pandas as pd
import logging
from pathlib import Path


class QCReader:
    """Reads and processes QC (Quality Control) Excel files for pole connection filtering"""
    
    def __init__(self, qc_file_path=None):
        """
        Initialize QC reader with optional QC file path
        
        Args:
            qc_file_path (str, optional): Path to QC Excel file
        """
        self.qc_file_path = qc_file_path
        self.connections = set()  # Set of (from_scid, to_scid) tuples (normalized for matching)
        self.ordered_connections = []  # List preserving QC file order (normalized for matching)
        self.original_ordered_connections = []  # List preserving EXACT QC file format
        self.qc_data_rows = []  # Complete row data from QC file
        self.qc_scids = set()  # All SCIDs mentioned in QC file (normalized)
        self._active = False
        
        if qc_file_path:
            self.load_qc_file(qc_file_path)
    
    def load_qc_file(self, qc_file_path):
        """
        Load QC data from Excel file
        
        Args:
            qc_file_path (str): Path to QC Excel file
        """
        try:
            qc_path = Path(qc_file_path)
            if not qc_path.exists():
                logging.warning(f"QC file not found: {qc_file_path}")
                self._active = False
                return
            
            # Read QC Excel file - get all sheet names first
            import openpyxl
            wb = openpyxl.load_workbook(qc_file_path, data_only=True)
            sheet_names = wb.sheetnames
            logging.info(f"QC file has {len(sheet_names)} sheets: {sheet_names}")
            
            all_connections = []
            all_data_rows = []
            sheets_processed = 0
            
            # Process each sheet
            for sheet_name in sheet_names:
                try:
                    # Try different header row positions (common positions are 0, 1, 2)
                    df = None
                    for header_row in [2, 0, 1]:  # Try row 3 first (index 2), then row 1 (index 0), then row 2 (index 1)
                        try:
                            temp_df = pd.read_excel(qc_file_path, sheet_name=sheet_name, header=header_row, dtype=str, keep_default_na=False)
                            # Check if this header row has the required columns
                            if 'Pole' in temp_df.columns and 'To Pole' in temp_df.columns:
                                df = temp_df
                                logging.debug(f"Sheet '{sheet_name}': Found headers at row {header_row + 1}")
                                break
                        except:
                            continue
                    
                    if df is None:
                        # If no valid header row found, try reading without header and look for Pole/To Pole columns
                        df = pd.read_excel(qc_file_path, sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
                        # Look for rows containing 'Pole' and 'To Pole'
                        for idx, row in df.iterrows():
                            if 'Pole' in row.values and 'To Pole' in row.values:
                                # Use this row as header
                                df.columns = row
                                df = df.iloc[idx+1:].reset_index(drop=True)
                                logging.debug(f"Sheet '{sheet_name}': Found headers at row {idx + 1}")
                                break
                        else:
                            logging.info(f"Sheet '{sheet_name}' missing required columns 'Pole' and 'To Pole' - skipping")
                            continue

                    
                    # Process connections from this sheet
                    sheet_connections = []
                    sheet_data_rows = []
                    for _, row in df.iterrows():
                        from_pole_orig = str(row['Pole']).strip()
                        to_pole_orig = str(row['To Pole']).strip()
                        
                        # Skip empty rows
                        if not from_pole_orig or not to_pole_orig or from_pole_orig == 'nan' or to_pole_orig == 'nan':
                            continue
                        
                        sheet_connections.append((from_pole_orig, to_pole_orig))
                        
                        # Store complete row data as dictionary
                        row_data = {}
                        for col in df.columns:
                            value = str(row[col]).strip() if pd.notna(row[col]) and str(row[col]).strip() != 'nan' else ''
                            row_data[col] = value
                        sheet_data_rows.append(row_data)
                    
                    all_connections.extend(sheet_connections)
                    all_data_rows.extend(sheet_data_rows)
                    sheets_processed += 1
                    logging.info(f"Sheet '{sheet_name}': found {len(sheet_connections)} connections with {len(df.columns)} columns")
                    
                except Exception as e:
                    logging.warning(f"Error reading sheet '{sheet_name}': {e}")
                    continue
            
            if sheets_processed == 0:
                logging.warning(f"No valid sheets found in QC file: {qc_file_path}")
                self._active = False
                return
            
            # Process all connections from all sheets
            self.connections.clear()
            self.ordered_connections.clear()
            self.original_ordered_connections.clear()
            self.qc_data_rows.clear()
            self.qc_scids.clear()
            
            for i, (from_pole_orig, to_pole_orig) in enumerate(all_connections):
                # Store original format (EXACT from QC file)
                original_connection = (from_pole_orig, to_pole_orig)
                self.original_ordered_connections.append(original_connection)
                
                # Store complete row data
                if i < len(all_data_rows):
                    self.qc_data_rows.append(all_data_rows[i])
                
                # Normalize SCIDs for matching purposes
                from_pole_norm = self._normalize_scid(from_pole_orig)
                to_pole_norm = self._normalize_scid(to_pole_orig)
                
                # Add normalized versions for internal matching
                normalized_connection = (from_pole_norm, to_pole_norm)
                self.connections.add(normalized_connection)
                self.connections.add((to_pole_norm, from_pole_norm))  # Add reverse for bidirectional matching
                self.ordered_connections.append(normalized_connection)
                
                self.qc_scids.add(from_pole_norm)
                self.qc_scids.add(to_pole_norm)
            
            self._active = len(self.connections) > 0
            self.qc_file_path = qc_file_path
            
            logging.info(f"Loaded QC file: {len(self.ordered_connections)} connections from {sheets_processed} sheets, {len(self.qc_scids)} unique SCIDs")
            
        except Exception as e:
            logging.error(f"Error loading QC file {qc_file_path}: {e}")
            self._active = False
    
    def _normalize_scid(self, scid):
        """
        Normalize SCID format (remove leading zeros, etc.)
        
        Args:
            scid (str): Raw SCID string
            
        Returns:
            str: Normalized SCID
        """
        scid = str(scid).strip()
        
        # Handle numeric SCIDs with leading zeros
        if scid.isdigit():
            return str(int(scid))
        
        # Handle alphanumeric SCIDs (e.g., "001A" -> "1A")
        import re
        match = re.match(r'(\d+)([A-Za-z]*)', scid)
        if match:
            num_part = str(int(match.group(1)))
            alpha_part = match.group(2)
            return num_part + alpha_part
        
        return scid
    
    def is_active(self):
        """
        Check if QC reader is active (has loaded QC data)
        
        Returns:
            bool: True if QC data is loaded and active
        """
        return self._active
    
    def get_qc_scids(self):
        """
        Get set of all SCIDs mentioned in QC file
        
        Returns:
            set: Set of SCID strings
        """
        return self.qc_scids.copy()
    
    def get_ordered_connections(self):
        """
        Get connections in the order they appear in QC file (normalized for matching)
        
        Returns:
            list: List of (from_scid, to_scid) tuples in QC file order (normalized)
        """
        return self.ordered_connections.copy()
    
    def get_original_ordered_connections(self):
        """
        Get connections in EXACT format from QC file (preserving original format)
        
        Returns:
            list: List of (from_scid, to_scid) tuples in exact QC file format
        """
        return self.original_ordered_connections.copy()
    
    def get_qc_data_rows(self):
        """
        Get complete row data from QC file
        
        Returns:
            list: List of dictionaries containing complete row data from QC file
        """
        return self.qc_data_rows.copy()
    
    def has_connection(self, from_scid, to_scid):
        """
        Check if a specific connection exists in QC file
        
        Args:
            from_scid (str): Source SCID
            to_scid (str): Destination SCID
            
        Returns:
            bool: True if connection exists in QC file
        """
        if not self._active:
            return False
        
        # Normalize inputs
        from_scid = self._normalize_scid(str(from_scid))
        to_scid = self._normalize_scid(str(to_scid))
        
        return (from_scid, to_scid) in self.connections
    
    def get_connections_set(self):
        """
        Get set of all connections (bidirectional)
        
        Returns:
            set: Set of (from_scid, to_scid) tuples
        """
        return self.connections.copy()
    
    def get_all_connections(self):
        """
        Get all connections (alias for get_ordered_connections for compatibility)
        
        Returns:
            list: List of (from_scid, to_scid) tuples in QC file order (normalized)
        """
        return self.get_ordered_connections()
    
    def create_consolidated_qc_sheet(self, output_path=None):
        """
        Create a consolidated QC sheet that combines data from all sheets
        
        Args:
            output_path (str, optional): Path for the output file. If None, overwrites original file.
            
        Returns:
            str: Path to the created consolidated QC file
        """
        if not self._active or not self.qc_file_path:
            logging.warning("No QC file loaded - cannot create consolidated sheet")
            return None
            
        try:
            import openpyxl
            from pathlib import Path
            
            # Determine output path
            if output_path is None:
                original_path = Path(self.qc_file_path)
                output_path = original_path.parent / f"{original_path.stem}_Consolidated{original_path.suffix}"
            
            # Load the original workbook
            wb = openpyxl.load_workbook(self.qc_file_path, data_only=True)
            
            # Create new workbook for consolidated data
            new_wb = openpyxl.Workbook()
            
            # Remove default sheet and create QC sheet
            new_wb.remove(new_wb.active)
            qc_sheet = new_wb.create_sheet(title="QC")
            
            # Add header in row 3
            qc_sheet['A3'] = 'Pole'
            qc_sheet['B3'] = 'To Pole'
            
            # Style the header
            from openpyxl.styles import Font, Alignment
            header_font = Font(bold=True)
            qc_sheet['A3'].font = header_font
            qc_sheet['B3'].font = header_font
            qc_sheet['A3'].alignment = Alignment(horizontal='center')
            qc_sheet['B3'].alignment = Alignment(horizontal='center')
            
            # Use the connections that were already loaded and processed
            all_connections = self.get_original_ordered_connections()
            logging.info(f"Using {len(all_connections)} connections already loaded from QC file")
            
            # Write all connections to the QC sheet starting from row 4
            row_num = 4
            for from_pole, to_pole in all_connections:
                qc_sheet[f'A{row_num}'] = from_pole
                qc_sheet[f'B{row_num}'] = to_pole
                row_num += 1
            
            # Adjust column widths
            qc_sheet.column_dimensions['A'].width = 15
            qc_sheet.column_dimensions['B'].width = 15
            
            # Save the consolidated workbook
            new_wb.save(output_path)
            logging.info(f"Created consolidated QC sheet with {len(all_connections)} connections")
            logging.info(f"Consolidated QC file saved to: {output_path}")
            
            return str(output_path)
            
        except Exception as e:
            logging.error(f"Error creating consolidated QC sheet: {e}")
            return None
