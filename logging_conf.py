import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
import traceback
import os
import win32com.client as win32


####################################################################################################################
# Setup logging
####################################################################################################################


def setup_logger(name="app_logger", log_dir="logs", level=logging.INFO):
    """
    Sets up and returns a configured logger.
    """

    # Ensure log directory exists
    os.makedirs(log_dir, exist_ok=True)

    log_filename = os.path.join(
        log_dir,
        f"{name}_{datetime.now().strftime('%Y%m%d')}.log"
    )

    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Prevent duplicate handlers
    if not logger.handlers:

        # File handler (rotating)
        file_handler = RotatingFileHandler(
            log_filename,
            maxBytes=10 * 1024 * 1024,  # 10 MB
            backupCount=3
        )

        # Formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )

        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Console handler (optional but useful)
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger, log_filename

def send_email_err_report(toemail, filepath, error_message, error_type, stack_trace):
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = toemail
    mail.Subject = 'PLEASE CHECK: Paycom Scraping Error Report'
    mail.Body = f"An error occurred during the Paycom scraping process.\n\nError Type: {error_type}\nError Message: {error_message}\nStack Trace:\n{stack_trace}"
    
    # Attach the log file if it exists
    if os.path.exists(filepath):
        mail.Attachments.Add(filepath)
    
    mail.Send()
