import unittest
from src.core.utils import Utils

class TestUtils(unittest.TestCase):

    def test_normalize_scid(self):
        # Test simple numeric SCIDs
        self.assertEqual(Utils.normalize_scid("001"), "1")
        self.assertEqual(Utils.normalize_scid("02A"), "2A")
        self.assertEqual(Utils.normalize_scid("  003 "), "3")
        self.assertEqual(Utils.normalize_scid("ABC"), "ABC")
        self.assertEqual(Utils.normalize_scid(""), "")
        
        # Test complex SCIDs with spaces and mixed alphanumeric
        self.assertEqual(Utils.normalize_scid("118 MISM013"), "118 MISM13")
        self.assertEqual(Utils.normalize_scid("5 TEST001"), "5 TEST1")
        self.assertEqual(Utils.normalize_scid("001 DEF002"), "1 DEF2")
        self.assertEqual(Utils.normalize_scid("MISM013"), "MISM13")
        self.assertEqual(Utils.normalize_scid("001A 002B"), "1A 2B")

    def test_parse_height_format(self):
        self.assertEqual(Utils.parse_height_format("5'-10\""), "5' 10\"")
        self.assertEqual(Utils.parse_height_format("6'"), "6' 0\"")
        self.assertEqual(Utils.parse_height_format("4' 6\""), "4' 6\"")
        self.assertEqual(Utils.parse_height_format(""), "")
        self.assertEqual(Utils.parse_height_format("5' - 10\""), "5' 10\"")

    def test_parse_height_decimal(self):
        self.assertEqual(Utils.parse_height_decimal("5'-10\""), 5.83)
        self.assertEqual(Utils.parse_height_decimal("6'"), 6.0)
        self.assertEqual(Utils.parse_height_decimal("4' 6\""), 4.5)
        self.assertEqual(Utils.parse_height_decimal(""), None)
        self.assertEqual(Utils.parse_height_decimal("5' - 10\""), 5.83)

    def test_inches_to_feet_format(self):
        self.assertEqual(Utils.inches_to_feet_format(60), "5' 0\"")
        self.assertEqual(Utils.inches_to_feet_format(72), "6' 0\"")
        self.assertEqual(Utils.inches_to_feet_format(65), "5' 5\"")
        self.assertEqual(Utils.inches_to_feet_format(0), "0' 0\"")
        self.assertEqual(Utils.inches_to_feet_format(-10), '')

if __name__ == '__main__':
    unittest.main()
