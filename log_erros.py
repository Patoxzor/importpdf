import logging

def setup_logging():
    logger = logging.getLogger('Mixpdf')
    
    if not logger.hasHandlers():  
        logger.setLevel(logging.ERROR) 

        file_handler = logging.FileHandler('erros_log.txt')
        file_handler.setLevel(logging.ERROR)

        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
    
    return logger

logger = setup_logging()
