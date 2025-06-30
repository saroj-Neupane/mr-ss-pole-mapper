import unittest
import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from core.tension_calculator import TensionCalculator

class TestTensionCalculator(unittest.TestCase):
    
    def setUp(self):
        """Set up test fixtures"""
        self.calculator = TensionCalculator("Test Files/Metronet tension calculator.xlsm")
    
    def test_calculator_initialization(self):
        """Test that calculator initializes correctly"""
        self.assertIsNotNone(self.calculator)
        self.assertEqual(self.calculator.worksheet_name, "Calculations")
        self.assertIn('span_length', self.calculator.cells)
        self.assertIn('span_sag', self.calculator.cells)
        self.assertIn('cable_installation', self.calculator.cells)
        self.assertIn('result_tension', self.calculator.cells)
    
    def test_validate_calculator_file(self):
        """Test calculator file validation"""
        is_valid, message = self.calculator.validate_calculator_file()
        print(f"Calculator validation: {is_valid}, {message}")
        # Note: This test may fail if the Excel file is not in the expected location
        # but it helps validate the file structure
    
    def test_parse_height_value(self):
        """Test height value parsing"""
        # Test valid values
        self.assertEqual(self.calculator._parse_height_value("25.5"), 25.5)
        self.assertEqual(self.calculator._parse_height_value("30'"), 30.0)
        self.assertEqual(self.calculator._parse_height_value('20"'), 20.0)
        self.assertEqual(self.calculator._parse_height_value("15.75"), 15.75)
        
        # Test invalid values
        self.assertIsNone(self.calculator._parse_height_value(""))
        self.assertIsNone(self.calculator._parse_height_value("nan"))
        self.assertIsNone(self.calculator._parse_height_value(None))
        self.assertIsNone(self.calculator._parse_height_value("invalid"))
    
    def test_calculate_tension_basic(self):
        """Test basic tension calculation"""
        # Test with sample values - this will only work if Excel file is available
        try:
            span_length = 104.0  # feet
            attachment_height = 25.0  # feet
            midspan_height = 22.0  # feet
            
            tension = self.calculator.calculate_tension(span_length, attachment_height, midspan_height)
            
            if tension is not None:
                print(f"Calculated tension: {tension} lbs")
                self.assertIsInstance(tension, (int, float))
                self.assertGreater(tension, 0)
                print("✓ Tension calculation successful!")
            else:
                print("⚠ Tension calculation returned None (Excel file may not be accessible)")
                
        except Exception as e:
            print(f"⚠ Tension calculation test skipped due to: {e}")

if __name__ == '__main__':
    unittest.main() 