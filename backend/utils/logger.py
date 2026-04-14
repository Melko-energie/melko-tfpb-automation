import logging
import sys
from pathlib import Path
from datetime import datetime

_logger = None
_log_buffer = []


def get_logger(name: str = "excel-to-dossier") -> logging.Logger:
    global _logger
    if _logger is not None:
        return _logger

    _logger = logging.getLogger(name)
    _logger.setLevel(logging.DEBUG)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter(
        "[%(asctime)s] %(levelname)-8s %(message)s",
        datefmt="%H:%M:%S"
    ))
    _logger.addHandler(console)

    log_dir = Path(__file__).resolve().parent.parent.parent / "logs"
    log_dir.mkdir(exist_ok=True)
    file_handler = logging.FileHandler(
        log_dir / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
        encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(
        "[%(asctime)s] %(levelname)-8s [%(module)s] %(message)s"
    ))
    _logger.addHandler(file_handler)

    buffer_handler = BufferHandler()
    buffer_handler.setLevel(logging.DEBUG)
    _logger.addHandler(buffer_handler)

    return _logger


class BufferHandler(logging.Handler):
    """Keeps logs in memory for the API to serve to the frontend."""

    def emit(self, record):
        entry = {
            "time": self.format(record) if self.formatter else record.getMessage(),
            "level": record.levelname,
            "message": record.getMessage(),
            "timestamp": datetime.now().isoformat(),
        }
        _log_buffer.append(entry)
        if len(_log_buffer) > 500:
            _log_buffer.pop(0)


def get_log_buffer() -> list:
    return list(_log_buffer)


def clear_log_buffer():
    _log_buffer.clear()
