"""
io_excel.py
===========
Leitura e escrita robustas de arquivos Excel.

Estrategia de leitura (read_excel_robusto):
  1. Tentativas diretas com pandas (retries com back-off).
  2. Copia temporaria via shutil.
  3. Fallback COM (win32com) - abre e salva copia antes de ler.

Escrita (save_excel_com_retries):
  Tenta salvar com pandas/openpyxl, com retries em caso de PermissionError.

Formatacao de datas (format_excel_date_columns):
  Aplica formato dd/mm/yyyy nas celulas de colunas datetime via openpyxl.
"""

import os
import shutil
import tempfile
import time

import pandas as pd
from openpyxl import load_workbook

from logger import get_logger

log = get_logger(__name__)


def read_excel_robusto(
    path: str,
    sheet_name=0,
    engine: str = "openpyxl",
    retries: int = 5,
    wait: float = 0.8,
) -> pd.DataFrame:
    """
    Le um arquivo Excel com multiplas estrategias de fallback.
    Levanta RuntimeError se todas as tentativas falharem.
    """
    last_err = None

    # 1) Tentativas diretas
    for i in range(retries):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, engine=engine)
        except PermissionError as exc:
            last_err = exc
            log.warning("Arquivo em uso, tentativa %d/%d: %s", i + 1, retries, path)
            time.sleep(wait * (i + 1))
        except Exception:
            raise

    # 2) Copia temporaria via shutil
    try:
        tmpdir  = tempfile.mkdtemp(prefix="dbgeral_")
        tmpfile = os.path.join(tmpdir, os.path.basename(path))
        shutil.copy2(path, tmpfile)
        log.info("Lendo copia temporaria: %s", tmpfile)
        return pd.read_excel(tmpfile, sheet_name=sheet_name, engine=engine)
    except Exception:
        pass

    # 3) Fallback COM (somente Windows)
    try:
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible       = False
        excel.DisplayAlerts = False

        tmpdir  = tempfile.mkdtemp(prefix="dbgeral_com_")
        tmpfile = os.path.join(tmpdir, os.path.basename(path))

        wb = excel.Workbooks.Open(path, ReadOnly=True)
        wb.SaveCopyAs(tmpfile)
        wb.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()

        log.info("Lendo copia COM: %s", tmpfile)
        return pd.read_excel(tmpfile, sheet_name=sheet_name, engine=engine)

    except Exception as exc_com:
        raise RuntimeError(
            f"Falha ao ler arquivo. Ultimo erro direto: {last_err} | fallback COM: {exc_com}"
        )


def save_excel_com_retries(
    df: pd.DataFrame,
    out_path: str,
    retries: int = 5,
    wait: float = 0.6,
) -> None:
    """
    Salva um DataFrame em Excel com retries em caso de PermissionError.
    Levanta o ultimo PermissionError se todas as tentativas falharem.
    """
    last_err = None
    for i in range(retries):
        try:
            df.to_excel(out_path, index=False, engine="openpyxl")
            return
        except PermissionError as exc:
            last_err = exc
            log.warning("Arquivo em uso (escrita), tentativa %d/%d", i + 1, retries)
            time.sleep(wait * (i + 1))
    raise last_err


def format_excel_date_columns(file_path: str, df_ref: pd.DataFrame) -> None:
    """
    Aplica o formato de celula 'dd/mm/yyyy' em todas as colunas
    que contenham objetos date/datetime, usando openpyxl.
    """
    wb = load_workbook(file_path)
    ws = wb.active

    for col_idx, col_name in enumerate(df_ref.columns, start=1):
        series = df_ref[col_name].dropna()
        if not series.empty and hasattr(series.iloc[0], "year"):
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                row[0].number_format = "dd/mm/yyyy"

    wb.save(file_path)
    wb.close()
