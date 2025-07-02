import logging
from pathlib import Path
import win32com.client as win32

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

def test_excel_com():
    """Test Excel COM automation step by step"""
    
    print("=== Testing Excel COM Automation ===")
    
    try:
        # Test 1: Check if win32com is available
        print("1. Testing win32com availability...")
        print(f"   win32com available: {win32 is not None}")
        
        # Test 2: Try to get existing Excel instance
        print("2. Testing existing Excel instance...")
        try:
            excel_app = win32.GetActiveObject("Excel.Application")
            print("   ✅ Found existing Excel instance")
        except Exception as e:
            print(f"   ⚠️ No existing Excel instance: {e}")
            excel_app = None
        
        # Test 3: Try to create new Excel instance
        print("3. Testing new Excel instance creation...")
        try:
            if excel_app is None:
                excel_app = win32.Dispatch("Excel.Application")
                print("   ✅ Created new Excel instance")
            else:
                print("   ✅ Using existing Excel instance")
        except Exception as e:
            print(f"   ❌ Failed to create Excel instance: {e}")
            return False
        
        # Test 4: Try to set Excel properties
        print("4. Testing Excel property settings...")
        properties = [
            ("Visible", False),
            ("DisplayAlerts", False),
            ("EnableEvents", False),
            ("ScreenUpdating", False)
        ]
        
        for prop_name, prop_value in properties:
            try:
                setattr(excel_app, prop_name, prop_value)
                print(f"   ✅ Set {prop_name} = {prop_value}")
            except Exception as e:
                print(f"   ⚠️ Could not set {prop_name}: {e}")
        
        # Test 5: Check if calculator file exists
        calculator_path = Path("Test_Files/Metronet tension calculator.xlsm")
        print(f"5. Testing calculator file...")
        print(f"   File exists: {calculator_path.exists()}")
        print(f"   File path: {calculator_path.absolute()}")
        
        if not calculator_path.exists():
            print("   ❌ Calculator file not found!")
            return False
        
        # Test 6: Try to open the workbook
        print("6. Testing workbook opening...")
        try:
            workbook = excel_app.Workbooks.Open(str(calculator_path.absolute()))
            print("   ✅ Successfully opened workbook")
            
            # Test 7: Check worksheets
            print("7. Testing worksheet access...")
            worksheet_names = [ws.Name for ws in workbook.Worksheets]
            print(f"   Available worksheets: {worksheet_names}")
            
            if "Calculations" in worksheet_names:
                print("   ✅ 'Calculations' worksheet found")
                worksheet = workbook.Worksheets("Calculations")
                
                # Test 8: Try to read a cell
                print("8. Testing cell reading...")
                try:
                    cell_value = worksheet.Range("A1").Value
                    print(f"   ✅ Successfully read cell A1: {cell_value}")
                except Exception as e:
                    print(f"   ❌ Failed to read cell A1: {e}")
                
                # Test 9: Try to write to a cell
                print("9. Testing cell writing...")
                try:
                    worksheet.Range("B2").Value = 100.0
                    read_back = worksheet.Range("B2").Value
                    print(f"   ✅ Successfully wrote and read back cell B2: {read_back}")
                except Exception as e:
                    print(f"   ❌ Failed to write to cell B2: {e}")
                
            else:
                print("   ❌ 'Calculations' worksheet not found")
            
            # Clean up
            workbook.Close(SaveChanges=False)
            print("   ✅ Closed workbook")
            
        except Exception as e:
            print(f"   ❌ Failed to open workbook: {e}")
            return False
        
        # Test 10: Clean up Excel
        print("10. Testing Excel cleanup...")
        try:
            excel_app.Quit()
            print("   ✅ Successfully quit Excel")
        except Exception as e:
            print(f"   ⚠️ Could not quit Excel: {e}")
        
        return True
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_excel_com()
    if success:
        print("\n✅ Excel COM automation test completed successfully!")
    else:
        print("\n❌ Excel COM automation test failed!") 