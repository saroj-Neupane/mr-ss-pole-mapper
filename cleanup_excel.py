#!/usr/bin/env python3
"""
Simple script to clean up orphaned Excel processes
Run this if you notice hidden Excel instances after using the pole mapper
Note: The tension calculator now uses existing Excel instances, so this should rarely be needed.
"""

import subprocess
import sys

def cleanup_excel_processes():
    """Clean up any orphaned Excel processes"""
    print("Cleaning up orphaned Excel processes...")
    
    try:
        # Force kill all Excel processes
        result = subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], 
                              capture_output=True, timeout=10)
        
        if result.returncode == 0:
            print("✓ Successfully cleaned up Excel processes")
            print(f"Output: {result.stdout.decode()}")
        else:
            print("✓ No Excel processes found to clean up")
            print(f"Output: {result.stderr.decode()}")
            
    except subprocess.TimeoutExpired:
        print("⚠ Timeout while cleaning up Excel processes")
    except Exception as e:
        print(f"✗ Error cleaning up Excel processes: {e}")

if __name__ == "__main__":
    print("Excel Process Cleanup Tool")
    print("=" * 30)
    
    # Ask for confirmation
    response = input("This will close ALL Excel processes. Continue? (y/N): ")
    if response.lower() in ['y', 'yes']:
        cleanup_excel_processes()
    else:
        print("Cleanup cancelled.") 