import pandas as pd
import sys
import os

# Add the src directory to the path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from core.pole_data_processor import PoleDataProcessor
from core.config_manager import ConfigManager

def test_lowest_overall_height_selection():
    """Test that the row with lowest overall attachment height is selected"""
    
    # Load configuration
    config_manager = ConfigManager()
    config = config_manager.load_config('configurations/MVEC.json')
    
    # Create processor
    processor = PoleDataProcessor(config)
    
    # Create test data with multiple sections for the same connection
    # This simulates multiple rows in sections_df for the same connection_id
    sections_df = pd.DataFrame({
        'connection_id': ['CONN1', 'CONN1', 'CONN1'],  # Same connection_id
        'pole': ['POLE1', 'POLE1', 'POLE1'],           # Same pole
        'to_pole': ['POLE2', 'POLE2', 'POLE2'],        # Same to_pole
        'POA_METRONET': ['MetroNet', 'MetroNet', 'MetroNet'],
        'POA_METRONETHT': ['30\'0"', '25\'6"', '35\'0"'],  # Row 2 has lowest MetroNet
        'POA_VERIZON': ['Verizon', 'Verizon', 'Verizon'],
        'POA_VERIZONHT': ['24\'0"', '29\'0"', '20\'0"'],   # Row 3 has lowest Verizon
        'POA_POWER': ['Power', 'Power', 'Power'],
        'POA_POWERHT': ['40\'0"', '35\'0"', '45\'0"']      # Row 2 has lowest Power
    })
    
    print("Test Data - Multiple sections for same connection:")
    print("Row 1: MetroNet=30'0\", Verizon=24'0\", Power=40'0\" (lowest: 24'0\")")
    print("Row 2: MetroNet=25'6\", Verizon=29'0\", Power=35'0\" (lowest: 25'6\")")
    print("Row 3: MetroNet=35'0\", Verizon=20'0\", Power=45'0\" (lowest: 20'0\")")
    print("\nExpected: Row 3 should be selected (has lowest overall height: 20'0\")")
    print()
    
    # Test the _find_section method
    section = processor._find_section('CONN1', sections_df, 'POLE1', 'POLE2')
    
    if section is not None:
        print("Selected section:")
        print(f"  MetroNet: {section['POA_METRONETHT']}")
        print(f"  Verizon: {section['POA_VERIZONHT']}")
        print(f"  Power: {section['POA_POWERHT']}")
        
        # Verify it selected the row with lowest overall height
        from src.core.utils import Utils
        
        metronet_decimal = Utils.parse_height_decimal(section['POA_METRONETHT'])
        verizon_decimal = Utils.parse_height_decimal(section['POA_VERIZONHT'])
        power_decimal = Utils.parse_height_decimal(section['POA_POWERHT'])
        
        print(f"Parsed heights: MetroNet={metronet_decimal}, Verizon={verizon_decimal}, Power={power_decimal}")
        
        lowest_height = min([h for h in [metronet_decimal, verizon_decimal, power_decimal] if h is not None])
        
        print(f"\nLowest height in selected row: {lowest_height}'")
        
        if abs(lowest_height - 20.0) < 0.01:  # Allow small floating point differences
            print("✓ SUCCESS: Row with lowest overall height (20'0\") was selected!")
        else:
            print("✗ FAILED: Wrong row was selected!")
            print(f"Expected: 20.0, Got: {lowest_height}")
    else:
        print("✗ FAILED: No section was found!")

if __name__ == "__main__":
    test_lowest_overall_height_selection() 