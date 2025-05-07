import os
from excel_processor import ExcelProcessor
from queue import Queue
import logging
import sys
import time
import gc

def setup_basic_logging():
    """Set up basic logging for test."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        stream=sys.stdout
    )
    return logging.getLogger("ExcelProcessorTest")

def test_processor():
    """Test the ExcelProcessor directly."""
    logger = setup_basic_logging()
    
    # Get test file path from user
    file_path = input("Enter path to Excel file to test: ")
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        return
        
    password = input("Enter sheet password: ")
    
    # Create a queue to receive messages from processor
    test_queue = Queue()
    
    # Create processor
    processor = ExcelProcessor(test_queue)
    
    try:
        # Close any existing Excel instances
        print("Closing any existing Excel instances...")
        processor.close_excel_instances()
        
        # Force garbage collection before processing
        print("Running garbage collection...")
        gc.collect()
        time.sleep(1)
        
        # Process the file
        print(f"Processing file: {file_path}")
        processor.process_file(file_path, password)
        
        # Print messages from queue
        print("\nMessages from processor:")
        while not test_queue.empty():
            msg_type, data = test_queue.get()
            print(f"{msg_type}: {data}")
            test_queue.task_done()
            
        print("\nProcessing completed successfully.")
        
    except Exception as e:
        logger.error("Processing failed", exc_info=True)
        print(f"\nError: {str(e)}")
        
        # Print messages from queue even if there was an error
        print("\nMessages from processor (after error):")
        while not test_queue.empty():
            msg_type, data = test_queue.get()
            print(f"{msg_type}: {data}")
            test_queue.task_done()
    
if __name__ == "__main__":
    test_processor()
