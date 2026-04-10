import os
import sys

from app.logging_config import get_logger
from app.database import criar

logger = get_logger(__name__)


def _project_root():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _get_app_icon_path():
    icon_path = os.path.join(_project_root(), "assets", "app.ico")
    if os.path.exists(icon_path):
        return icon_path
    return None


def run_flask():
    host = os.getenv("FLASK_HOST", "127.0.0.1")
    port = int(os.getenv("FALLBACK_FLASK_PORT", os.getenv("FLASK_PORT", "5001")))
    debug = os.getenv("FLASK_DEBUG", "true").strip().lower() in {"1", "true", "yes", "on"}
    logger.info("Iniciando servidor Flask (fallback)...")
    from app.web import app as flask_app
    flask_app.run(debug=debug, host=host, port=port)


def run_qt():
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication
    from app.ui_main import MainWindow

    logger.info("Tentando iniciar interface Qt...")
    app = QApplication(sys.argv)
    icon_path = _get_app_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))
    window = MainWindow()
    window.show()
    window.raise_()
    window.activateWindow()

    return app.exec()


if __name__ == "__main__":
    criar()

    # No container/headless sem DISPLAY, usar fallback web
    if sys.platform.startswith("linux") and not os.environ.get("DISPLAY"):
        logger.warning("DISPLAY não encontrado no Linux, ativando fallback web.")
        run_flask()
        sys.exit(0)

    try:
        exit_code = run_qt()
        sys.exit(exit_code)

    except Exception as err:
        logger.warning("Qt não pôde ser inicializado: %s", err)
        logger.info("Fallback para interface web Flask")
        run_flask()
        sys.exit(0)
