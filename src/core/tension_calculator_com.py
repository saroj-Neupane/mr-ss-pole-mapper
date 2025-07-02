import logging
import pandas as pd
from pathlib import Path
import shutil
import tempfile
import os
import time
from contextlib import contextmanager
import win32com.client as win32
from .utils import Utils
import pythoncom

try:
    import win32com.client as win32
    from win32com.client import constants
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    logging.warning("pywin32 not available - tension calculations will use openpyxl fallback")

class TensionCalculatorCOM:
    """Handles tension calculations using Excel COM automation"""
    
    def __init__(self, calculator_file_path="", worksheet_name="Calculations"):
        self.calculator_file_path = Path(calculator_file_path)
        self.worksheet_name = worksheet_name
        
        # Cell locations in the calculator
        self.cells = {
            'span_length': 'B2',
            'span_sag': 'E2', 
            'cable_installation': 'M4',
            'result_tension': 'R12',
            'calculate_button': 'P2'  # Add the cell reference for the Calculate button
        }
        
        # Excel COM objects - initialized on first use
        self._excel_app = None
        self._workbook = None
        self._worksheet = None
        self._temp_path = None
        self._is_initialized = False

    def _ensure_initialized(self):
        """Ensure Excel is initialized, initialize if needed"""
        if not self._is_initialized:
            if not self._initialize_excel():
                raise RuntimeError("Failed to initialize Excel")
    
    def _initialize_excel(self):
        """Initialize Excel instance and workbook if not already initialized"""
        if not COM_AVAILABLE:
            logging.error("COM automation not available - cannot initialize Excel")
            return False
        
        # Check if calculator file path is valid
        if not self.calculator_file_path or not self.calculator_file_path.exists():
            logging.warning(f"Tension calculator file not found or invalid: {self.calculator_file_path}")
            return False
            
        try:
            if self._excel_app is None:
                # Create temporary copy of calculator file
                with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as temp_file:
                    self._temp_path = Path(temp_file.name)
                
                # Copy calculator file to temp location
                shutil.copy2(self.calculator_file_path, self._temp_path)
                
                # Try to get existing Excel instance first
                try:
                    self._excel_app = win32.GetActiveObject("Excel.Application")
                    logging.info("Using existing Excel instance")
                except:
                    # If no existing instance, create a new one
                    self._excel_app = win32.Dispatch("Excel.Application")
                    logging.info("Created new Excel instance")
                
                # Test if we can access the Workbooks collection
                try:
                    test_workbooks = self._excel_app.Workbooks
                    logging.info("Excel Workbooks collection is accessible")
                except Exception as e:
                    logging.error(f"Excel Workbooks collection not accessible: {e}")
                    logging.error("Excel COM automation is not working properly. This may be due to:")
                    logging.error("1. Excel running in restricted mode")
                    logging.error("2. Permission issues with COM automation")
                    logging.error("3. Excel running as a different user")
                    self.cleanup()
                    return False
                
                # Configure the Excel instance to be hidden and non-interactive
                try:
                    self._excel_app.Visible = False
                except Exception as e:
                    logging.warning(f"Could not set Excel.Visible: {e}")
                try:
                    self._excel_app.DisplayAlerts = False
                except Exception as e:
                    logging.warning(f"Could not set Excel.DisplayAlerts: {e}")
                try:
                    self._excel_app.EnableEvents = False  # Disable events for speed
                except Exception as e:
                    logging.warning(f"Could not set Excel.EnableEvents: {e}")
                try:
                    self._excel_app.ScreenUpdating = False  # Disable screen updates
                except Exception as e:
                    logging.warning(f"Could not set Excel.ScreenUpdating: {e}")
                
                # Open workbook
                try:
                    self._workbook = self._excel_app.Workbooks.Open(str(self._temp_path.absolute()))
                    logging.info("Successfully opened calculator workbook")
                except Exception as e:
                    logging.error(f"Failed to open calculator workbook: {e}")
                    self.cleanup()
                    return False
                
                # Get worksheet
                try:
                    for ws in self._workbook.Worksheets:
                        if ws.Name == self.worksheet_name:
                            self._worksheet = ws
                            break
                            
                    if not self._worksheet:
                        raise ValueError(f"Worksheet '{self.worksheet_name}' not found")
                    
                    logging.info(f"Successfully accessed worksheet: {self.worksheet_name}")
                except Exception as e:
                    logging.error(f"Failed to access worksheet '{self.worksheet_name}': {e}")
                    self.cleanup()
                    return False
                
                self._is_initialized = True
                logging.info("Excel instance initialized successfully")
                
            return True
                
        except Exception as e:
            self.cleanup()
            logging.error(f"Failed to initialize Excel: {str(e)}")
            return False

    def calculate_tension(self, span_length, attachment_height, midspan_height):
        """Calculate tension for a single set of measurements"""
        try:
            self._ensure_initialized()
            return self._calculate_single_tension(span_length, attachment_height, midspan_height)
        except RuntimeError as e:
            logging.warning(f"Tension calculation skipped: {e}")
            return None

    def _calculate_single_tension(self, span_length, attachment_height, midspan_height):
        """Internal method to calculate a single tension value"""
        try:
            if not self._worksheet or not self._excel_app:
                logging.error("Excel worksheet or application not initialized")
                return None

            # Parse height values from feet'inches" format to decimal feet
            def parse_height(height_str):
                try:
                    if isinstance(height_str, (int, float)):
                        return round(float(height_str), 2)

                    decimal_value = Utils.parse_height_decimal(height_str)
                    if decimal_value is not None:
                        return round(decimal_value, 2)

                    # Fallback to stripping units and casting
                    clean_str = str(height_str).replace("'", "").replace('"', "").strip()
                    return round(float(clean_str), 2)
                except (ValueError, TypeError) as e:
                    logging.error(f"Error parsing height value '{height_str}': {e}")
                    return None

            # Convert heights to decimal feet (2 decimal places)
            attachment_decimal = parse_height(attachment_height)
            midspan_decimal = parse_height(midspan_height)
            
            if attachment_decimal is None or midspan_decimal is None:
                logging.error(f"Failed to parse height values: attachment={attachment_height}, midspan={midspan_height}")
                return None

            # Parse span length to ensure it's a number with 2 decimal places
            try:
                span_length = round(float(str(span_length).replace("'", "").strip()), 2)
                logging.info(f"Using span length: {span_length:.2f} ft")
            except (ValueError, TypeError) as e:
                logging.error(f"Failed to parse span length '{span_length}': {e}")
                return None

            # Calculate span sag as difference between attachment height and midspan height
            span_sag = round(attachment_decimal - midspan_decimal, 2)
            
            # Use minimum span sag of 0.8 to avoid division by zero
            if span_sag <= 0:
                logging.info("Span sag is 0 or negative, using minimum value of 0.8")
                span_sag = 0.8
            
            logging.info(f"CALCULATION INPUTS:")
            logging.info(f"1. Span Length (B2) = {span_length:.2f} ft")
            logging.info(f"2. Attachment Height (M4) = {attachment_decimal:.2f} ft (from {attachment_height})")
            logging.info(f"3. Midspan Height = {midspan_decimal:.2f} ft (from {midspan_height})")
            logging.info(f"4. Span Sag (E2) = {span_sag:.2f} ft (attachment - midspan)")
            
            try:
                # Directly set cell values (all rounded to 2 decimal places)
                self._worksheet.Range("B2").Value = span_length
                self._worksheet.Range("E2").Value = span_sag
                self._worksheet.Range("M4").Value = attachment_decimal  # Use attachment height for cable installation
                
                # Verify values were written correctly
                written_span = round(float(self._worksheet.Range("B2").Value), 2)
                written_sag = round(float(self._worksheet.Range("E2").Value), 2)
                written_install = round(float(self._worksheet.Range("M4").Value), 2)
                logging.info(f"EXCEL CELL VALUES:")
                logging.info(f"B2 (Span Length) = {written_span:.2f}")
                logging.info(f"E2 (Span Sag) = {written_sag:.2f}")
                logging.info(f"M4 (Cable Installation) = {written_install:.2f}")
                
                # Run the calculation macro
                logging.info("Running Calc_Sag_Data macro")
                self._excel_app.Run("Calc_Sag_Data")
                
                # Small delay to ensure calculation completes
                time.sleep(0.1)
                
                # Read and verify result
                tension_result = self._worksheet.Range(self.cells['result_tension']).Value
                logging.info(f"Raw tension result from Excel: {tension_result}")
                
                if tension_result is not None:
                    try:
                        # Round tension to whole number
                        tension_value = round(float(tension_result))
                        logging.info(f"Final tension value (rounded): {tension_value}")
                        return tension_value
                    except (ValueError, TypeError) as e:
                        logging.error(f"Invalid tension result: {tension_result}")
                else:
                    logging.error("No tension result read from cell")
                
                return None
                
            except Exception as e:
                logging.error(f"Error during Excel operations: {str(e)}")
                return None
            
        except Exception as e:
            logging.error(f"Error calculating tension: {str(e)}")
            return None

    def calculate_tensions_for_providers(self, provider_data_list):
        """
        Calculate tensions for multiple providers in one batch
        
        Args:
            provider_data_list: List of tuples (provider_data, span_length)
            
        Returns:
            List of calculated tension values
        """
        try:
            self._ensure_initialized()
        except RuntimeError as e:
            logging.warning(f"Tension calculation skipped: {e}")
            return [None] * len(provider_data_list)
        
        results = []
        
        for provider_data, span_length in provider_data_list:
            try:
                # Extract heights from provider data
                attachment_height = None
                midspan_height = None
                
                # Look for attachment height
                for key, value in provider_data.items():
                    if 'attachment' in key.lower() or key.endswith('Attachment Ht'):
                        if value and str(value).strip():
                            attachment_height = self._parse_height_value(value)
                            break
                
                # Look for midspan height  
                for key, value in provider_data.items():
                    if 'midspan' in key.lower() or key.endswith('Midspan Ht'):
                        if value and str(value).strip():
                            midspan_height = self._parse_height_value(value)
                            break
                
                if None in (attachment_height, midspan_height, span_length) or span_length <= 0:
                    results.append(None)
                    continue
                
                # Calculate tension
                tension = self._calculate_single_tension(span_length, attachment_height, midspan_height)
                results.append(tension)
                
            except Exception as e:
                logging.error(f"Error processing provider data: {str(e)}")
                results.append(None)
        
        return results

    def cleanup(self):
        """Clean up Excel resources - only close workbook, don't quit Excel"""
        try:
            # Close workbook only - don't quit Excel application
            if self._workbook:
                try:
                    self._workbook.Close(SaveChanges=False)
                    logging.info("Calculator workbook closed successfully")
                except Exception as e:
                    logging.warning(f"Error closing workbook: {str(e)}")
                finally:
                    self._workbook = None
                    self._worksheet = None
                
        except Exception as e:
            logging.error(f"Error during cleanup: {str(e)}")
        finally:
            self._worksheet = None
            self._is_initialized = False
            self._excel_app = None  # Clear reference but don't quit
            
            # Clean up temporary file
            if self._temp_path:
                try:
                    self._temp_path.unlink()
                except:
                    pass
                self._temp_path = None
    
    def __del__(self):
        """Ensure cleanup on object destruction"""
        try:
            self.cleanup()
        except:
            pass  # Ignore errors during destruction

    @contextmanager
    def excel_context(self):
        """Context manager for Excel operations with automatic cleanup"""
        try:
            self._ensure_initialized()
            yield self
        finally:
            self.cleanup()

    def _parse_height_value(self, value):
        """Parse height value from various formats including feet'inches" """
        try:
            if value is None or (isinstance(value, float) and pd.isna(value)):
                return None

            decimal_feet = Utils.parse_height_decimal(value)

            if decimal_feet is not None:
                return round(decimal_feet, 2)

            # Fallback: strip quotes / units and attempt float conversion
            value_str = str(value).replace("'", "").replace('"', "").replace("ft", "").replace("feet", "").strip()
            return round(float(value_str), 2)

        except (ValueError, TypeError) as e:
            logging.warning(f"Could not parse height value '{value}': {e}")
            return None 