import sys
sys.path.append('src')
from core.pole_data_processor import PoleDataProcessor
from core.config_manager import ConfigManager

# Test the height text
p = PoleDataProcessor(ConfigManager().get_default_config())

height_text = """GROUND STREETLIGHT AND COVER FEED
AT HOA 24'0" LOWER ZAYO TO HOA 22'1"
AT HOA 22'1" LOWER COMCAST TO HOA 21'1" WITH DG
AT HOA 21'10" LOWER JACKSON ISD TO HOA 20'1"
AT HOA 21'8" LOWER US SIGNAL TO HOA 19'1"
AT HOA 21'5" LOWER AT&T GUY TO HOA 18'1"
AT HOA 20'7" LOWER AT&T TO HOA 18'1" """

print("Testing height text extraction:")
print("=" * 50)
print("Input text:")
print(height_text)
print()

# Test extraction
result = p._extract_guy_info(height_text)
print("Extracted result:")
print(f"  Leads: {result.get('leads', [])}")
print(f"  Directions: {result.get('directions', [])}")
print(f"  Sizes: {result.get('sizes', [])}")

print(f"\nExpected: Should extract NOTHING (no PL NEW patterns)")
print(f"Actual:   Extracted {len(result.get('leads', []))} items")

if len(result.get('leads', [])) == 0:
    print("✅ CORRECT: No extraction from height text!")
else:
    print("❌ INCORRECT: Extracting from height text when it shouldn't!")
    print("This is happening because the general pattern is matching height values.") 