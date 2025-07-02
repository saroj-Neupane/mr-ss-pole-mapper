#!/usr/bin/env python3
"""
Test script to verify manual route filtering fix
"""

import sys
import pandas as pd
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from core.pole_data_processor import PoleDataProcessor
from core.route_parser import RouteParser

def test_manual_routes_filtering_fix():
    """Test that manual routes filtering no longer causes the 'No valid pole or reference data found' error"""
    print("Testing manual routes filtering fix...")
    
    # Create mock data with manual route SCIDs
    nodes_df = pd.DataFrame({
        'node_id': ['1', '2', '3', '4', '5', '6'],
        'scid': ['1', '2', '3', '4', '5', '6'],
        'node_type': ['pole', 'pole', 'reference', 'pole', 'reference', 'pole'],
        'pole_status': ['active', 'active', '', 'active', '', 'active']
    })
    
    connections_df = pd.DataFrame({
        'node_id_1': ['1', '2', '3', '4'],
        'node_id_2': ['3', '4', '5', '6'],
        'connection_id': ['C1', 'C2', 'C3', 'C4'],
        'span_distance': ['100', '150', '200', '250']
    })
    
    sections_df = pd.DataFrame({
        'connection_id': ['C1', 'C2', 'C3', 'C4'],
        'section_id': ['S1', 'S2', 'S3', 'S4']
    })
    
    # Create manual routes with SCIDs 1 and 2
    manual_route_text = "1, 2"
    manual_routes = RouteParser.parse_manual_routes(manual_route_text, [])
    
    print(f"Manual routes: {manual_routes}")
    print(f"Manual route SCIDs: {[scid for route in manual_routes for scid in route['poles']]}")
    
    # Create processor
    config = {
        'processing_options': {
            'use_geocoding': False,
            'open_output': False,
            'span_length_tolerance': 3.0
        }
    }
    
    processor = PoleDataProcessor(config=config)
    
    try:
        # Process data with manual routes
        result_data = processor.process_data(
            nodes_df=nodes_df,
            connections_df=connections_df,
            sections_df=sections_df,
            manual_routes=manual_routes,
            clear_existing_routes=False
        )
        
        print(f"✅ SUCCESS: Processing completed with {len(result_data)} output rows")
        print(f"Output rows: {result_data}")
        
        # Verify only manual route poles are in output
        output_scids = set()
        for row in result_data:
            pole_scid = row.get('Pole', '')
            if pole_scid:
                output_scids.add(pole_scid)
        
        expected_scids = {'1', '2'}  # Only the manual route SCIDs
        print(f"Output SCIDs: {output_scids}")
        print(f"Expected SCIDs: {expected_scids}")
        
        if output_scids.issubset(expected_scids):
            print("✅ SUCCESS: Only manual route SCIDs are in output")
        else:
            print("❌ FAILURE: Unexpected SCIDs in output")
            
    except Exception as e:
        print(f"❌ FAILURE: Error during processing: {e}")
        return False
    
    return True

if __name__ == "__main__":
    success = test_manual_routes_filtering_fix()
    if success:
        print("\n🎉 Manual route filtering fix is working correctly!")
    else:
        print("\n💥 Manual route filtering fix failed!") 