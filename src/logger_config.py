# logger_config.py
import logging

def setup_logger():
    logger = logging.getLogger(__name__)
    if not logger.handlers:  # Avoid adding handlers multiple times
        logger.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler = logging.StreamHandler()
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger

# Create a logger instance that can be imported by other modules
logger = setup_logger()