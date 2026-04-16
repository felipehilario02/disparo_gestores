"""
comentarios.py
==============
Carrega e normaliza o mapa {conta -> comentario} a partir do arquivo de comentarios.

Funcao publica principal:
    carregar_base_comentarios(alt_file, sheet_name) -> dict[str, str]
"""

import re
import unicodedata
from typing import Optional

import pandas as pd

from logger import get_logger

log = get_logger(__name__)


def _norm_col(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("ASCII")
    return s.strip().lower()


def _norm_conta_key(val) -> Optional[str]:
    """
    Converte um valor de celula 'Conta' em chave normalizada para lookup.
    Retorna None para valores vazios/nulos.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    try:
        f = float(s.replace(",", "."))
        if abs(f - int(f)) < 1e-9:
            s = str(int(f))
    except Exception:
        pass
    s = s.replace("Total", "").replace("TOTAL", "").strip()
    s = re.sub(r"\s+", "", s)
    return s or None


def _carregar_conta_comments_excel(arquivo: str, sheet_name: str) -> dict[str, str]:
    """
    Le a aba de comentarios e retorna {conta_normalizada: comentario}.
    Retorna {} em caso de qualquer falha ou estrutura inesperada.
    """
    try:
        dfc = pd.read_excel(arquivo, sheet_name=sheet_name, engine="openpyxl")
    except Exception as exc:
        log.warning("Falha ao ler '%s' aba '%s': %s", arquivo, sheet_name, exc)
        return {}

    if dfc is None or dfc.empty:
        log.warning("Aba '%s' vazia em: %s", sheet_name, arquivo)
        return {}

    cols_map = {_norm_col(c): c for c in dfc.columns}

    def _find_col_name(candidates: list[str]) -> Optional[str]:
        for cand in candidates:
            key = _norm_col(cand)
            if key in cols_map:
                return cols_map[key]
        base = _norm_col(candidates[0]).split()[0]
        for k, orig in cols_map.items():
            if base in k:
                return orig
        return None

    col_conta  = _find_col_name(["Conta"])
    col_coment = _find_col_name(["Comentario", "Comentarios"])

    if not col_conta or not col_coment:
        log.warning(
            "Aba '%s' precisa das colunas 'Conta' e 'Comentario'. Encontradas: %s",
            sheet_name, list(dfc.columns),
        )
        return {}

    conta_comments: dict[str, str] = {}
    for _, row in dfc.iterrows():
        key   = _norm_conta_key(row[col_conta])
        value = row[col_coment]
        if key and isinstance(value, str) and value.strip():
            conta_comments[key] = value.strip()

    log.info(
        "Comentarios carregados: %d contas (arquivo='%s', aba='%s').",
        len(conta_comments), arquivo, sheet_name,
    )
    if conta_comments:
        sample = list(conta_comments.items())[:3]
        log.debug("Exemplos de contas carregadas: %s", sample)

    return conta_comments


def carregar_base_comentarios(
    alt_file: Optional[str],
    sheet_name: str = "Comentarios",
) -> dict[str, str]:
    """
    Carrega o mapa de comentarios somente do arquivo alternativo.
    Tenta variacoes de nome de aba. Retorna {} se nada for encontrado.
    """
    if not alt_file:
        log.info("Nenhum arquivo de comentarios configurado.")
        return {}

    abas = list(dict.fromkeys([sheet_name, "Comentarios", "Comentario"]))

    for aba in abas:
        result = _carregar_conta_comments_excel(alt_file, sheet_name=aba)
        if result:
            return result

    log.info("Comentarios nao encontrados no arquivo alternativo: %s", alt_file)
    return {}
