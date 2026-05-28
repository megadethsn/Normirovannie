import logging

from app_paths import log_path


def setup_logging():
    root = logging.getLogger()
    if root.handlers:
        return
    handler = logging.FileHandler(log_path(), encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    root.addHandler(handler)
    root.setLevel(logging.INFO)


def get_logger(name):
    return logging.getLogger(name)
