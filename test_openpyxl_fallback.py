import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

def test_openpyxl_calculator():
    """Test the openpyxl-based tension calculator"""
    
    # Check if the calculator file exists
    calculator_path = Path("Test_Files/Metronet tension calculator.xlsm")
    print(f"Calculator file exists: {calculator_path.exists()}")
    print(f"Calculator file path: {calculator_path.absolute()}")
    
    if not calculator_path.exists():
        print("❌ Calculator file not found!")
        return
    
    try:
        from src.core.tension_calculator import TensionCalculator
        
        # Initialize calculator
        print("Initializing openpyxl-based tension calculator...")
        calculator = TensionCalculator(str(calculator_path))
        
        # Test calculation
        print("Testing tension calculation...")
        span_length = 100.0
        attachment_height = 26.33  # 26' 4" in decimal feet
        midspan_height = 25.0      # 25' 0" in decimal feet
        
        print(f"Input values:")
        print(f"  Span length: {span_length} ft")
        print(f"  Attachment height: {attachment_height} ft")
        print(f"  Midspan height: {midspan_height} ft")
        
        tension = calculator.calculate_tension(span_length, attachment_height, midspan_height)
        
        if tension is not None:
            print(f"✅ Tension calculation successful: {tension} lbs")
        else:
            print("❌ Tension calculation failed - returned None")
            
    except Exception as e:
        print(f"❌ Error testing openpyxl tension calculator: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_openpyxl_calculator() 