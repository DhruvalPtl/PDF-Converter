import logging
import os

def log_system():
    try:
        log_dir = 'C:/PdfConverter/log'
        os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, 'app.log')
        print(f"Log file path: {log_file}")

        # Clear any existing handlers
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logging.basicConfig(
            filename=log_file,
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        logging.debug("Logging setup successful.")
        print("Logging setup successful.")
    except Exception as e:
        print(f"Logging setup failed: {e}")
        logging.exception("Exception occurred during logging setup.")
