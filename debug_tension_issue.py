from src.core.tension_calculator_com import TensionCalculatorCOM
import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO)

def convert_feet_inches(measurement):
    """Convert feet-inches format (e.g. '26' 4"') to decimal feet"""
    if isinstance(measurement, (int, float)):
        return float(measurement)
    
    parts = measurement.replace('"', '').split("'")
    feet = float(parts[0].strip())
    inches = float(parts[1].strip()) if len(parts) > 1 else 0
    return feet + (inches / 12.0)

def test_tension_calculation():
    # Check if the file exists
    calculator_path = Path("Test_Files/Metronet tension calculator.xlsm")
    logging.info(f"Checking if calculator file exists: {calculator_path}")
    logging.info(f"Absolute path: {calculator_path.absolute()}")
    logging.info(f"File exists: {calculator_path.exists()}")
    
    # List contents of Test_Files directory
    test_files_dir = Path("Test_Files")
    if test_files_dir.exists():
        logging.info(f"Contents of {test_files_dir}:")
        for file in test_files_dir.iterdir():
            logging.info(f"  {file.name}")
    
    # Use the correct path to the tension calculator file
    calculator = TensionCalculatorCOM("Test_Files/Metronet tension calculator.xlsm")
    
    # Test case from screenshot (pole 003)
    span_length = 100  # feet
    attachment_height = convert_feet_inches("26' 4")  # Proposed height of new attachment point
    final_midspan = convert_feet_inches("25' 0")     # Final Mid Span Ground Clearance of Proposed Attachment
    
    logging.info("\nTesting with exact values from screenshot (pole 003)")
    logging.info(f"Input values (after conversion):")
    logging.info(f"span_length = {span_length} feet")
    logging.info(f"attachment_height = {attachment_height} feet (from 26' 4\")")
    logging.info(f"final_midspan = {final_midspan} feet (from 25' 0\")")
    logging.info(f"calculated_sag = {attachment_height - final_midspan} feet")
    
    tension = calculator.calculate_tension(
        span_length,
        attachment_height,  # Proposed height of new attachment point
        final_midspan      # Final Mid Span Ground Clearance of Proposed Attachment
    )
    
    logging.info(f"Calculated tension: {tension}")
    logging.info(f"Expected tension from screenshot: 1541.2")
    logging.info(f"Difference: {abs(tension - 1541.2) if tension else 'N/A'}")
    
    calculator.cleanup()

if __name__ == '__main__':
    test_tension_calculation()
