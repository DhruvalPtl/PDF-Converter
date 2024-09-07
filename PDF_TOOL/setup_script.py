import logging
import shutil
import os

def install_resources():

    source_dir = os.path.join(os.path.dirname(__file__), 'resources')
    target_dir = 'C:/PdfConverter'  
    logging.info(f"Starting installation. Source: {source_dir}, Target: {target_dir}")

    try:

        os.makedirs(target_dir, exist_ok=True)
        os.makedirs("C:/PdfConverter/Temp", exist_ok=True)
        os.makedirs("C:/PdfConverter/Temp1", exist_ok=True)
        for item in os.listdir(source_dir):
            s = os.path.join(source_dir, item)
            d = os.path.join(target_dir, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, dirs_exist_ok=True)
            else:
                shutil.copy2(s, d)
        logging.info("Installation completed successfully.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")