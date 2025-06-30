from pathlib import Path
import pandas as pd
import json
import logging
import re
try:
    from .utils import Utils
except ImportError:
    from utils import Utils

class AttachmentDataReader:
    """Handles reading attachment data from the new Excel format"""
    
    def __init__(self, file_path, config=None, valid_scids=None):
        self.file_path = file_path
        self.attachment_data = {}
        self.config = config or {}
        self.valid_scids = set(valid_scids) if valid_scids else None
        self.load_attachment_data()
    
    def load_attachment_data(self):
        """Load attachment data from Excel file with SCID sheets.
           Sheet names are expected to be 'SCID <scid>' where <scid> is already filtered.
        """
        try:
            xls = pd.ExcelFile(self.file_path)
            scid_sheets = [sheet for sheet in xls.sheet_names if sheet.startswith("SCID ")]
            
            logging.info(f"AttachmentDataReader: Discovered {len(scid_sheets)} SCID sheet(s) in {self.file_path}")
            
            for sheet_name in scid_sheets:
                scid = sheet_name[5:].strip()
                ignore_keywords = self.config.get('ignore_scid_keywords', [])
                scid = Utils.normalize_scid(scid, ignore_keywords)
                
                if self.valid_scids is not None and scid not in self.valid_scids:
                    logging.debug(f"AttachmentDataReader: Skipping sheet '{sheet_name}' because SCID '{scid}' is not in the valid set")
                    continue
                
                try:
                    df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=1)
                    df = df.fillna("")
                    
                    logging.info(f"AttachmentDataReader: Loaded {len(df)} record(s) for SCID '{scid}' from sheet '{sheet_name}'")
                    
                    df.columns = df.columns.str.strip().str.lower()
                    required_cols = ['company', 'measured', 'height_in_inches']
                    missing_cols = [col for col in required_cols if col not in df.columns]
                    if missing_cols:
                        logging.warning(f"AttachmentDataReader: Sheet '{sheet_name}' missing columns: {missing_cols}. Available: {list(df.columns)}")
                        continue
                    
                    self.attachment_data[scid] = df
                except Exception as e:
                    logging.error(f"AttachmentDataReader: Error reading sheet '{sheet_name}': {e}")
            
            logging.info(f"AttachmentDataReader: Total valid SCIDs loaded: {len(self.attachment_data)}")
            if not self.attachment_data:
                logging.error("AttachmentDataReader: No valid SCID data loaded from the attachment file!")
        except Exception as e:
            logging.error(f"AttachmentDataReader: Failed to load attachment data from {self.file_path}: {e}")
    
    def get_scid_data(self, scid):
        """Get attachment data for a specific SCID"""
        ignore_keywords = self.config.get('ignore_scid_keywords', [])
        normalized_scid = Utils.normalize_scid(scid, ignore_keywords)
        data = self.attachment_data.get(normalized_scid, pd.DataFrame())
        if data.empty:
            logging.debug(f"No attachment data found for SCID {scid} (normalized: {normalized_scid})")
        return data

    def find_power_attachment(self, scid, power_keywords):
        """Find the lowest power attachment for a SCID"""
        df = self.get_scid_data(scid)
        if df.empty:
            return None
        try:
            power_company = self.config.get("power_company", "").strip().lower()
            if not power_company:
                return None
            
            df['company_stripped'] = df['company'].astype(str).str.strip().str.lower()
            
            power_company_pattern = r'\b' + re.escape(power_company) + r'\b'
            power_company_rows = df[df['company_stripped'].str.contains(power_company_pattern, na=False, regex=True)]
            
            if power_company_rows.empty:
                return None
            
            power_company_rows['measured_stripped'] = power_company_rows['measured'].astype(str).str.strip().str.lower()
            
            keyword_pattern = '|'.join([re.escape(kw.strip().lower()) for kw in power_keywords])
            power_rows = power_company_rows[
                power_company_rows['measured_stripped'].str.contains(keyword_pattern, na=False, regex=True)
            ]
            
            if power_rows.empty:
                return None
            
            power_rows['height_numeric'] = pd.to_numeric(
                power_rows['height_in_inches'].astype(str).str.replace('"', '').str.replace('″', ''), 
                errors='coerce'
            )
            power_rows = power_rows.dropna(subset=['height_numeric'])
            
            if not power_rows.empty:
                min_row = power_rows.loc[power_rows['height_numeric'].idxmin()]
                height_formatted = Utils.inches_to_feet_format(str(int(min_row['height_numeric'])))

                result = {
                    'height': height_formatted,
                    'height_decimal': float(min_row['height_numeric']) / 12,
                    'company': min_row['company'],
                    'measured': min_row['measured']
                }
                return result
        except Exception as e:
            logging.error(f"Error processing power attachment for SCID {scid}: {e}")
        return None
    
    def find_telecom_attachments(self, scid, telecom_keywords):
        """Find telecom attachments for a SCID and combine multiple heights in the same cell."""
        df = self.get_scid_data(scid)
        if df.empty:
            logging.warning(f"No data found for SCID {scid}")
            return {}
        
        attachments = {}
        try:
            logging.debug(f"Available columns for SCID {scid}: {list(df.columns)}")
            
            if 'height_in_inches' not in df.columns:
                logging.error(f"'height_in_inches' column missing for SCID {scid}")
                return {}
            
            for provider, keywords in telecom_keywords.items():
                clean_keywords = [kw.strip() for kw in keywords if kw.strip()]
                main_name = provider.strip()
                if main_name and main_name not in clean_keywords:
                    clean_keywords.append(main_name)
                
                company_regex = r'\b(?:' + '|'.join(re.escape(k.lower()) for k in clean_keywords) + r')\b'
                
                # Updated keywords for communication attachment selection
                # Include: 'CATV Com', 'Telco Com', 'Fiber Optic Com', 'insulator', 'Power Guy'
                if provider.lower() == 'proposed metronet':
                    measured_regex = r'(?i)(catv com|telco com|fiber optic com|insulator|power guy)'
                else:
                    measured_regex = r'(?i)(catv com|telco com|fiber optic com|insulator|power guy)'
                
                # For "Power Guy" keyword, company name must be in company column, not measured column
                if 'power guy' in measured_regex.lower():
                    # Check if measured column contains "power guy"
                    power_guy_rows = df[df['measured'].astype(str).str.lower().str.contains('power guy', na=False)]
                    if not power_guy_rows.empty:
                        # For power guy entries, company name should be in company column
                        provider_rows = power_guy_rows[
                            power_guy_rows['company'].astype(str).str.lower().str.contains(company_regex, na=False, regex=True)
                        ]
                    else:
                        # Regular telecom attachment matching
                        provider_rows = df[
                            (df['company'].astype(str).str.lower().str.contains(company_regex, na=False, regex=True)) &
                            (df['measured'].astype(str).str.contains(measured_regex, na=False, regex=True))
                        ]
                else:
                    # Regular telecom attachment matching
                    provider_rows = df[
                        (df['company'].astype(str).str.lower().str.contains(company_regex, na=False, regex=True)) &
                        (df['measured'].astype(str).str.contains(measured_regex, na=False, regex=True))
                    ]
                
                if not provider_rows.empty:
                    # Clean and convert height data with better error handling
                    def clean_height_value(height_val):
                        """Clean height value and convert to numeric"""
                        if pd.isna(height_val):
                            return None
                        height_str = str(height_val).replace('"', '').replace('″', '').strip()
                        if not height_str:
                            return None
                        try:
                            return pd.to_numeric(height_str)
                        except (ValueError, TypeError):
                            logging.warning(f"Could not convert height value '{height_val}' to numeric for SCID {scid}")
                            return None
                    
                    provider_rows['height_numeric'] = provider_rows['height_in_inches'].apply(clean_height_value)
                    
                    # Filter out rows with invalid height data
                    valid_rows = provider_rows.dropna(subset=['height_numeric'])
                    invalid_count = len(provider_rows) - len(valid_rows)
                    if invalid_count > 0:
                        logging.warning(f"Dropped {invalid_count} rows with invalid height data for provider {provider}, SCID {scid}")
                    
                    if not valid_rows.empty:
                        valid_rows = valid_rows.sort_values(by='height_numeric', ascending=False)
                        
                        heights = []
                        decimal_values = []
                        
                        for _, row in valid_rows.iterrows():
                            height_inches = row['height_numeric']
                            height_formatted = Utils.inches_to_feet_format(str(int(height_inches)))
                            if height_formatted:  # Only add if conversion was successful
                                heights.append(height_formatted)
                                decimal_values.append(float(height_inches) / 12)
                        
                        if heights:  # Only create attachment if we have valid heights
                            combined_heights = ', '.join(heights)
                            min_decimal = min(decimal_values) if decimal_values else None
                            
                            attachments[provider] = {
                                'heights': combined_heights,
                                'height_decimal': min_decimal,
                                'company': valid_rows.iloc[0]['company'],
                                'measured': valid_rows.iloc[0]['measured']
                            }
                            logging.debug(f"Created attachment for {provider}, SCID {scid}: {combined_heights}")
                        else:
                            logging.warning(f"No valid height data found for provider {provider}, SCID {scid}")
                    else:
                        logging.warning(f"No valid rows found for provider {provider}, SCID {scid} after height validation")
            return attachments
        except Exception as e:
            logging.error(f"Error processing telecom attachments for SCID {scid}: {e}")
            return {}
    
    def find_streetlight_attachment(self, scid):
        """Find the lowest street light attachment for a SCID (measured contains 'street light')"""
        df = self.get_scid_data(scid)
        if df.empty:
            return None
        try:
            df['measured_stripped'] = df['measured'].astype(str).str.strip().str.lower()
            streetlight_rows = df[df['measured_stripped'].str.contains('street light', na=False)]
            
            if streetlight_rows.empty:
                return None
            
            streetlight_rows['height_numeric'] = pd.to_numeric(
                streetlight_rows['height_in_inches'].astype(str).str.replace('"', '').str.replace('″', ''),
                errors='coerce'
            )
            streetlight_rows = streetlight_rows.dropna(subset=['height_numeric'])
            
            if not streetlight_rows.empty:
                min_row = streetlight_rows.loc[streetlight_rows['height_numeric'].idxmin()]
                height_formatted = Utils.inches_to_feet_format(str(int(min_row['height_numeric'])))
                return {
                    'height': height_formatted,
                    'height_decimal': float(min_row['height_numeric']) / 12,
                    'measured': min_row['measured']
                }
        except Exception as e:
            logging.error(f"Error processing streetlight attachment for SCID {scid}: {e}")
        return None