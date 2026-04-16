"""
logger.py
=========
Configuração central de logging do pipeline de despesas.

Características:
  - Um único logger nomeado "despesas" usado por todos os módulos.
  - Saída simultânea para console (StreamHandler) e arquivo rotativo (logs/pipeline.log).
  - Nível configurável via variável de ambiente LOG_LEVEL (default: INFO).
  - Formato com timestamp, nível e nome do módulo chamador.
  - Rotação diária do arquivo de log (30 dias de retenção).

Uso em qualquer módulo:
    from logger import get_logger
    log = get_logger(__name__)
    log.info("Mensagem")
    log.warning("Aviso")
    log.error("Erro: %s", exc)
"""

import logging
import os
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

_LOG_DIR = Path(__file__).parent / "logs"
_LOG_FILE = _LOG_DIR / "pipeline.log"

_LEVEL_MAP = {
    "DEBUG":   logging.DEBUG,
    "INFO":    logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR":   logging.ERROR,
}

_FMT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
_DATE_FMT = "%Y-%m-%d %H:%M:%S"

_initialized = False


def _setup() -> None:
    global _initialized
    if _initialized:
        return

    _LOG_DIR.mkdir(exist_ok=True)

    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    level = _LEVEL_MAP.get(level_name, logging.INFO)

    root = logging.getLogger("despesas")
    root.setLevel(level)
    root.propagate = False

    formatter = logging.Formatter(_FMT, datefmt=_DATE_FMT)

    # Console
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    root.addHandler(console_handler)

    # Arquivo rotativo (diário, 30 dias)
    file_handler = TimedRotatingFileHandler(
        _LOG_FILE,
        when="midnight",
        backupCount=30,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    _initialized = True


def get_logger(name: str) -> logging.Logger:
    """
    Retorna um logger filho do logger raiz 'despesas'.
    Chame com get_logger(__name__) em cada módulo.
    """
    _setup()
    # Se name já começa com 'despesas', usa direto; senão, cria filho
    if name.startswith("despesas"):
        return logging.getLogger(name)
    return logging.getLogger(f"despesas.{name}")
