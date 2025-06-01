import logging
import os
import sys
from logging.handlers import RotatingFileHandler as RH      # A logging tool that creates new files when a file reaches a size limit.

app_dir: str = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys,'frozen', False) else os.path.dirname(os.path.abspath(__file__))

log_file: str = os.path.join(app_dir,"errors.log")

# Set up rotating file handler (max 100 KB per file, keep 3 backups)
handler = RH(log_file, maxBytes=100_000, backupCount=3, encoding='utf-8')
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)

logger = logging.getLogger(__name__)
logger.setLevel(logging.ERROR)
logger.addHandler(handler)
logger.propagate = False

def log_error(error_msg: str) -> None:
    logger.error(error_msg)

def show_error() -> str:
    """Returns the last non-empty line from the log file, or a fallback message."""
    try:
        with open(log_file, 'rb') as f:
            f.seek(0, 2) # Go to the end of file
            end = f.tell()
            # if end == 0:
            #     return "⚠️ Log file is empty."

            buffer = bytearray()
            pointer = end - 1

            while pointer >= 0:
                f.seek(pointer)
                byte = f.read(1)
                if byte == b'\n' and buffer:
                    break
                buffer.insert(0, byte[0])
                pointer -= 1

            return buffer.decode('utf-8').strip()

    except Exception as e:
        return f"⚠️ Could not read log: {e}"