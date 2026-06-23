from __future__ import annotations

import logging
import logging.handlers
import os
import sys

from core.config import APP_DIR, LOG_FILE, MODEL_FILE
from ui.theme import APP_NAME, APP_VERSION


def setup_logging() -> logging.Logger:
    """Configure application logging handlers and return the named logger.

    Returns:
        logging.Logger: Configured logger instance for the application.
    """
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    log = logging.getLogger(APP_NAME)
    log.setLevel(logging.DEBUG)
    log.propagate = False

    target_path = os.fspath(LOG_FILE)
    has_file_handler = False
    for handler in log.handlers:
        if isinstance(handler, logging.handlers.RotatingFileHandler):
            base = getattr(handler, "baseFilename", "")
            if os.path.normcase(base) == os.path.normcase(target_path):
                has_file_handler = True
                break

    if not has_file_handler:
        file_handler = logging.handlers.RotatingFileHandler(
            LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
        file_handler.setFormatter(fmt)
        file_handler.setLevel(logging.DEBUG)
        log.addHandler(file_handler)

    log.info("=" * 60)
    log.info("Session started | version=%s | python=%s", APP_VERSION, sys.version.split()[0])
    log.info("app_dir=%s", APP_DIR)
    log.info("model=%s | exists=%s", MODEL_FILE, MODEL_FILE.exists())
    return log


logger = setup_logging()
