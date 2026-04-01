import os
import sys
import logging

from Projeto.database import criar

logger = logging.getLogger(__name__)
if not logger.hasHandlers():
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    logger.addHandler(handler)
logger.setLevel(logging.INFO)


def run_flask():
    logger.info("Iniciando servidor Flask (fallback)...")
    from Projeto.web import app as flask_app
    flask_app.run(debug=True, host="127.0.0.1", port=5001)


def run_qt():
    from PySide6.QtWidgets import QApplication
    from Projeto.ui_main import MainWindow

    logger.info("Tentando iniciar interface Qt...")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()

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
