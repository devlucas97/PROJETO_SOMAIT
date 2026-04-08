import logging
import os


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.hasHandlers():
        handler = logging.StreamHandler()
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
        )
        logger.addHandler(handler)
    debug = os.getenv("FLASK_DEBUG", "false").strip().lower() in {"1", "true", "yes", "on"}
    logger.setLevel(logging.DEBUG if debug else logging.INFO)
    return logger
