import pandas as pd
import sys
import os

# Add the src directory to the path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from core.pole_data_processor import PoleDataProcessor
from core.config_manager import ConfigManager

def test_midspan_multiple_connections():
    """Test midspan data extraction with multiple connections per node"""
    
    # Load configuration
    config_manager = ConfigManager()
    config = config_manager.load_config('configurations/MVEC.json')
    
    # Create processor
    processor = PoleDataProcessor(config)
    
    # Create test data
    nodes_df = pd.DataFrame({
        'node_id': ['NODE1', 'NODE2', 'NODE3'],
        'scid': ['POLE1', 'POLE2', 'POLE3'],
        'node_type': ['pole', 'pole', 'pole'],
        'mr_note': ['', '', '']
    })
    
    # Create connections with same connection_id but different midspan data
    connections_df = pd.DataFrame({
        'node_id_1': ['NODE1', 'NODE1', 'NODE2'],
        'node_id_2': ['NODE2', 'NODE3', 'NODE3'],
        'connection_id': ['CONN1', 'CONN1', 'CONN2'],  # Same connection_id for first two
        'span_distance': ['100', '150', '200']
    })
    
    # Create sections with different midspan data for the same connection_id
    # Row 1: MetroNet=25'6", Verizon=24'0", Power=35'0" (lowest overall: 24'0")
    # Row 2: MetroNet=30'0", Verizon=29'0", Power=40'0" (lowest overall: 29'0")
    # Row 3: MetroNet=28'0", Verizon=27'0", Power=38'0" (lowest overall: 27'0")
    sections_df = pd.DataFrame({
        'connection_id': ['CONN1', 'CONN1', 'CONN2'],
        'pole': ['POLE1', 'POLE1', 'POLE2'],
        'to_pole': ['POLE2', 'POLE3', 'POLE3'],
        'POA_METRONET': ['MetroNet', 'MetroNet', 'MetroNet'],
        'POA_METRONETHT': ['25\'6"', '30\'0"', '28\'0"'],
        'POA_VERIZON': ['Verizon', 'Verizon', 'Verizon'],
        'POA_VERIZONHT': ['24\'0"', '29\'0"', '27\'0"'],  # This should be the lowest for each row
        'POA_POWER': ['Power', 'Power', 'Power'],
        'POA_POWERHT': ['35\'0"', '40\'0"', '38\'0"']
    })
    
    print("Test Data:")
    print("Nodes:")
    print(nodes_df)
    print("\nConnections:")
    print(connections_df)
    print("\nSections:")
    print(sections_df)
    
    # Test the _find_section method directly
    print("\n=== Testing _find_section method ===")
    
    # Test finding section for CONN1, POLE1->POLE2
    section1 = processor._find_section('CONN1', sections_df, 'POLE1', 'POLE2')
    print(f"Section for CONN1 (POLE1->POLE2): {section1['POA_METRONETHT'] if section1 is not None else 'None'}")
    # Test finding section for CONN1, POLE1->POLE3
    section1b = processor._find_section('CONN1', sections_df, 'POLE1', 'POLE3')
    print(f"Section for CONN1 (POLE1->POLE3): {section1b['POA_METRONETHT'] if section1b is not None else 'None'}")
    # Test finding section for CONN2, POLE2->POLE3
    section2 = processor._find_section('CONN2', sections_df, 'POLE2', 'POLE3')
    print(f"Section for CONN2 (POLE2->POLE3): {section2['POA_METRONETHT'] if section2 is not None else 'None'}")
    
    # Test processing connections
    print("\n=== Testing connection processing ===")
    
    # Create mappings
    mappings = {
        'valid_poles': {'NODE1', 'NODE2', 'NODE3'},
        'node_id_to_scid': {'NODE1': 'POLE1', 'NODE2': 'POLE2', 'NODE3': 'POLE3'},
        'node_id_to_row': {
            'NODE1': {'node_id': 'NODE1', 'scid': 'POLE1', 'node_type': 'pole'},
            'NODE2': {'node_id': 'NODE2', 'scid': 'POLE2', 'node_type': 'pole'},
            'NODE3': {'node_id': 'NODE3', 'scid': 'POLE3', 'node_type': 'pole'}
        },
        'scid_to_row': {
            'POLE1': {'node_id': 'NODE1', 'scid': 'POLE1', 'node_type': 'pole'},
            'POLE2': {'node_id': 'NODE2', 'scid': 'POLE2', 'node_type': 'pole'},
            'POLE3': {'node_id': 'NODE3', 'scid': 'POLE3', 'node_type': 'pole'}
        }
    }
    
    # Process connections
    result_data = processor._process_standard_connections(connections_df, mappings, sections_df)
    
    print(f"Processed {len(result_data)} connections")
    
    # Check midspan data for each connection
    for i, row_data in enumerate(result_data):
        print(f"\nConnection {i+1}:")
        print(f"  Pole: {row_data.get('Pole', 'N/A')}")
        print(f"  To Pole: {row_data.get('To Pole', 'N/A')}")
        print(f"  Connection ID: {row_data.get('connection_id', 'N/A')}")
        print(f"  Proposed MetroNet_Midspan: {row_data.get('Proposed MetroNet_Midspan', 'N/A')}")
        print(f"  Verizon_Midspan: {row_data.get('Verizon_Midspan', 'N/A')}")

if __name__ == "__main__":
    test_midspan_multiple_connections() 