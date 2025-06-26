import sys
import logging
from pathlib import Path
from tkinter import Tk

def setup_global_exception_handler():
    """Setup global exception handler to log uncaught exceptions"""
    def global_exception_handler(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        if issubclass(exc_type, RecursionError):
            logging.error("Recursion error detected. Application will exit.")
        else:
            logging.error(f"An unexpected error occurred: {exc_value}")
            
        # Continue execution without showing message box
    
    sys.excepthook = global_exception_handler
    
def handle_exception(exc_type, exc_value, exc_traceback):
    """Handle exceptions in tkinter applications"""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    
    if issubclass(exc_type, RecursionError):
        logging.error("Recursion error detected. Application will exit.")
    else:
        logging.error(f"An unexpected error occurred: {exc_value}")
        
    # Continue execution without showing message box

def main():
    """Main application entry point"""
    
    # Set up basic logging first
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[logging.StreamHandler()]
    )
    
    # Setup global exception handler
    setup_global_exception_handler()
    
    try:
        # Create root window
        root = Tk()
        root.withdraw()  # Hide the root window initially
        
        # Setup tkinter exception handler
        root.report_callback_exception = handle_exception
        
        # Import and start the GUI application
        from gui.main_window import PoleMapperApp
        app = PoleMapperApp(root)
        root.deiconify()  # Show the window
        root.mainloop()
            
    except Exception as e:
        logging.error(f"Failed to start application: {str(e)}")
        raise

if __name__ == "__main__":
    main()