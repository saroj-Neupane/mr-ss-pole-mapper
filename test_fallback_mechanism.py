import logging
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

def test_fallback_mechanism():
    """Test the fallback mechanism in pole data processor"""
    
    # Create a mock config
    config = {
        "tension_calculator": {
            "file_path": "Test_Files/Metronet tension calculator.xlsm",
            "worksheet_name": "Calculations"
        }
    }
    
    print("Testing fallback mechanism in pole data processor...")
    
    try:
        from src.core.pole_data_processor import PoleDataProcessor
        
        # Initialize processor (this should trigger the fallback mechanism)
        print("Initializing PoleDataProcessor...")
        processor = PoleDataProcessor(config=config)
        
        # Test tension calculation
        print("Testing tension calculation through processor...")
        tension = processor.tension_calculator.calculate_tension(100.0, 26.33, 25.0)
        
        if tension is not None:
            print(f"✅ Tension calculation successful: {tension} lbs")
            print("✅ Fallback mechanism is working!")
        else:
            print("❌ Tension calculation failed - returned None")
            print("❌ Fallback mechanism may not be working")
            
    except Exception as e:
        print(f"❌ Error testing fallback mechanism: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_fallback_mechanism() 