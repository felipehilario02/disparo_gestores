"""
helpers.py
==========
Funcoes utilitarias puras: manipulacao de datas, strings, nomes de arquivo
e localizacao de colunas em DataFrames.

Nenhuma dependencia de outros modulos do projeto.
"""

import os
import re
import unicodedata
from datetime import date
from typing import Optional, Tuple

from config import MESES_PT, FY_START_MONTH
from logger import get_logger

log = get_logger(__name__)


# ─────────────────────────────────────────────
# Calendario / Fiscal Year
# ─────────────────────────────────────────────

def prev_month_date(ref_date: Optional[date] = None) -> date:
    """Retorna um date apontando para o 1o dia do mes anterior."""
    if ref_date is None:
        ref_date = date.today()
    month = ref_date.month - 1
    year  = ref_date.year
    if month == 0:
        month = 12
        year -= 1
    return date(year, month, 1)


def mes_pt(ref_date: date) -> str:
    """Nome do mes em portugues (ex.: 'Marco')."""
    return MESES_PT[ref_date.month - 1]


def fiscal_year_2d(ref_date: date, fy_start_month: int = FY_START_MONTH) -> int:
    """
    Retorna o FY em 2 digitos.
    Exemplo: outubro/2025 -> FY26 -> retorna 26.
    """
    fy_year = ref_date.year + 1 if ref_date.month >= fy_start_month else ref_date.year
    return fy_year % 100


def mes_anterior_pt_e_fy(ref_date: Optional[date] = None) -> Tuple[str, int]:
    """Retorna (nome_mes_pt, fy_2d) com base no mes anterior ao ref_date."""
    pm = prev_month_date(ref_date)
    return mes_pt(pm), fiscal_year_2d(pm, FY_START_MONTH)


# ─────────────────────────────────────────────
# Strings / Normalizacao
# ─────────────────────────────────────────────

def normalize_ascii_lower(s: str) -> str:
    """Remove acentos e converte para lowercase."""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    return s.strip().lower()


def find_col(cols, candidates: list[str]) -> Optional[str]:
    """
    Localiza a primeira coluna cujo nome normalizado bata com algum candidato.
    Retorna o nome original da coluna ou None.
    """
    cols_map = {normalize_ascii_lower(c): c for c in cols}
    for cand in candidates:
        key = normalize_ascii_lower(cand)
        if key in cols_map:
            return cols_map[key]
    return None


# ─────────────────────────────────────────────
# Nomes de arquivo e diretorios
# ─────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    """Remove caracteres invalidos em nomes de pasta/arquivo."""
    s = str(name).strip()
    for ch in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
        s = s.replace(ch, '-')
    return s


def sanitize_filename_component(s: str) -> str:
    """Remove caracteres invalidos preservando espacos simples."""
    s = s.strip()
    s = re.sub(r'[\\/:*?"<>|]+', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s


def ensure_dir(path: str) -> None:
    """Cria o diretorio (e intermediarios) se nao existir."""
    os.makedirs(path, exist_ok=True)


# ─────────────────────────────────────────────
# Templates
# ─────────────────────────────────────────────

def render_template(template: str, context: dict) -> str:
    """Substitui {chave} no template pelos valores do context."""
    return template.format(**context)
