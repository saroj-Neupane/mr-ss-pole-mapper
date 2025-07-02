import logging
from pathlib import Path
import win32com.client as win32

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

def test_force_new_excel():
    """Test with a forced new Excel instance"""
    
    print("=== Testing with Forced New Excel Instance ===")
    
    try:
        # Force a new Excel instance
        print("1. Creating new Excel instance...")
        excel_app = win32.Dispatch("Excel.Application")
        print("   ✅ Created new Excel instance")
        
        # Set properties (these should work with a new instance)
        print("2. Setting Excel properties...")
        try:
            excel_app.Visible = False
            print("   ✅ Set Visible = False")
        except Exception as e:
            print(f"   ⚠️ Could not set Visible: {e}")
        
        try:
            excel_app.DisplayAlerts = False
            print("   ✅ Set DisplayAlerts = False")
        except Exception as e:
            print(f"   ⚠️ Could not set DisplayAlerts: {e}")
        
        try:
            excel_app.EnableEvents = False
            print("   ✅ Set EnableEvents = False")
        except Exception as e:
            print(f"   ⚠️ Could not set EnableEvents: {e}")
        
        try:
            excel_app.ScreenUpdating = False
            print("   ✅ Set ScreenUpdating = False")
        except Exception as e:
            print(f"   ⚠️ Could not set ScreenUpdating: {e}")
        
        # Check calculator file
        calculator_path = Path("Test_Files/Metronet tension calculator.xlsm")
        print(f"3. Checking calculator file...")
        print(f"   File exists: {calculator_path.exists()}")
        print(f"   File path: {calculator_path.absolute()}")
        
        if not calculator_path.exists():
            print("   ❌ Calculator file not found!")
            excel_app.Quit()
            return False
        
        # Try to open the workbook
        print("4. Opening workbook...")
        try:
            workbook = excel_app.Workbooks.Open(str(calculator_path.absolute()))
            print("   ✅ Successfully opened workbook")
            
            # Check worksheets
            print("5. Checking worksheets...")
            worksheet_names = [ws.Name for ws in workbook.Worksheets]
            print(f"   Available worksheets: {worksheet_names}")
            
            if "Calculations" in worksheet_names:
                print("   ✅ 'Calculations' worksheet found")
                worksheet = workbook.Worksheets("Calculations")
                
                # Test cell operations
                print("6. Testing cell operations...")
                try:
                    # Write test value
                    worksheet.Range("B2").Value = 100.0
                    print("   ✅ Wrote to cell B2")
                    
                    # Read back
                    value = worksheet.Range("B2").Value
                    print(f"   ✅ Read back from B2: {value}")
                    
                    # Test other cells
                    worksheet.Range("E2").Value = 1.33
                    worksheet.Range("M4").Value = 26.33
                    print("   ✅ Wrote to cells E2 and M4")
                    
                    # Try to run macro (this might fail if macro doesn't exist)
                    try:
                        excel_app.Run("Calc_Sag_Data")
                        print("   ✅ Successfully ran Calc_Sag_Data macro")
                    except Exception as e:
                        print(f"   ⚠️ Could not run macro: {e}")
                    
                    # Read result
                    try:
                        result = worksheet.Range("R12").Value
                        print(f"   ✅ Read result from R12: {result}")
                    except Exception as e:
                        print(f"   ⚠️ Could not read result: {e}")
                    
                except Exception as e:
                    print(f"   ❌ Cell operation failed: {e}")
                
            else:
                print("   ❌ 'Calculations' worksheet not found")
            
            # Clean up
            workbook.Close(SaveChanges=False)
            print("   ✅ Closed workbook")
            
        except Exception as e:
            print(f"   ❌ Failed to open workbook: {e}")
            excel_app.Quit()
            return False
        
        # Clean up Excel
        print("7. Cleaning up...")
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
    success = test_force_new_excel()
    if success:
        print("\n✅ Force new Excel test completed successfully!")
    else:
        print("\n❌ Force new Excel test failed!") 