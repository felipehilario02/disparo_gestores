"""
processor.py
============
Transformacoes no DataFrame e orquestracao do fluxo por gestor.

Exporta a dataclass ExecutionResult, usada por main.py
para compor o relatorio de execucao final.

Modelo atual:
- Arquivo Excel salvo no SharePoint (pasta por gestor)
- E-mail enviado SEM anexo (apenas notificação + links no HTML)
"""

import os
import time
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
import pandas as pd

from config import (
    ALLOWED_COMMENTS_RECIPIENTS,
    BI_LINK,
    BODY_TEMPLATE,
    EMAIL_MAP,
    GESTORES_DIR,
    LOGO_CONTENT_ID,
    SUBJECT_TEMPLATE,
)
from email_sender import send_email_sendgrid
from helpers import (
    ensure_dir,
    find_col,
    render_template,
    sanitize_filename_component,
    sanitize_name,
)
from io_excel import format_excel_date_columns, save_excel_com_retries
from logger import get_logger
from pivot import create_pivot_win32

log = get_logger(__name__)


# ══════════════════════════════════════════════════════════════════
# Resultado de execucao
# ══════════════════════════════════════════════════════════════════

@dataclass
class ExecutionResult:
    """Acumula o resultado de cada gestor processado."""
    enviados:     list[str] = field(default_factory=list)
    sem_email:    list[str] = field(default_factory=list)
    falhas_pivot: list[str] = field(default_factory=list)
    falhas_email: list[str] = field(default_factory=list)
    ignorados:    list[str] = field(default_factory=list)

    def resumo(self) -> str:
        linhas = [
            "",
            "=" * 55,
            "  RESUMO DA EXECUCAO",
            "=" * 55,
            f"  Emails enviados    : {len(self.enviados)}",
            f"  Sem email mapeado  : {len(self.sem_email)}",
            f"  Falhas de pivot    : {len(self.falhas_pivot)}",
            f"  Falhas de email    : {len(self.falhas_email)}",
            f"  Ignorados (0 lin.) : {len(self.ignorados)}",
            "=" * 55,
        ]
        if self.enviados:
            linhas.append(f"  OK         : {', '.join(self.enviados)}")
        if self.sem_email:
            linhas.append(f"  SEM EMAIL  : {', '.join(self.sem_email)}")
        if self.falhas_pivot:
            linhas.append(f"  PIVOT FALHOU: {', '.join(self.falhas_pivot)}")
        if self.falhas_email:
            linhas.append(f"  EMAIL FALHOU: {', '.join(self.falhas_email)}")
        if self.ignorados:
            linhas.append(f"  IGNORADOS  : {', '.join(self.ignorados)}")
        linhas.append("=" * 55)
        return "\n".join(linhas)


# ══════════════════════════════════════════════════════════════════
# Transformacoes de DataFrame
# ══════════════════════════════════════════════════════════════════

def _normalize_mes_column(df: pd.DataFrame, label: str = "") -> pd.DataFrame:
    """Garante coluna MES no formato mm/YYYY e cria EXIBICAO_MES."""
    mes_col = find_col(df.columns, ["MES", "Mes"])

    df = df.copy()

    if not mes_col:
        base_col = find_col(
            df.columns,
            ["Competencia", "Mes Ref", "Data Competencia", "DATA", "DT BASE", "Data Base"],
        )
        if base_col:
            try:
                dt_tmp = pd.to_datetime(df[base_col], errors="coerce", dayfirst=True)
                df["MES"] = dt_tmp.dt.strftime("%m/%Y")
            except Exception as exc:
                log.warning("%s Falha ao derivar MES de %s: %s", label, base_col, exc)
                df["MES"] = np.nan
        else:
            log.warning("%s Nenhuma coluna de data encontrada para MES.", label)
            df["MES"] = np.nan
    else:
        try:
            dt_tmp = pd.to_datetime(df[mes_col], errors="coerce", dayfirst=True)
            df["MES"] = dt_tmp.dt.strftime("%m/%Y")
        except Exception:
            df.rename(columns={mes_col: "MES"}, inplace=True)

    try:
        exib_dt = pd.to_datetime(df["MES"], format="%m/%Y", errors="coerce")
        df["EXIBICAO_MES"] = exib_dt.dt.strftime("%d/%m/%Y")
    except Exception as exc:
        log.warning("%s Falha ao criar EXIBICAO_MES: %s", label, exc)
        df["EXIBICAO_MES"] = np.nan

    return df


def _normalize_cenario(df: pd.DataFrame) -> tuple[pd.DataFrame, Optional[str]]:
    scen_col = find_col(df.columns, ["DATA BASE", "Cenario", "Scenario"])
    if not scen_col:
        return df, None

    df = df.copy()
    df[scen_col] = (
        df[scen_col]
        .astype(str)
        .str.strip()
        .str.upper()
        .replace({
            "ORCADO": "BUDGET",
            "ATUAL": "ACTUAL",
            "REAL": "ACTUAL",
            "REALIZADO": "ACTUAL",
        })
    )
    return df, scen_col


def _allowed_months_from_actual(
    df: pd.DataFrame,
    scen_col: Optional[str],
    label: str = "",
) -> Optional[list[str]]:
    if not scen_col or "MES" not in df.columns:
        return None

    mask = df[scen_col].eq("ACTUAL")
    if not mask.any():
        return None

    meses = (
        df.loc[mask, "MES"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )
    try:
        meses.sort(key=lambda s: (int(s.split("/")[1]), int(s.split("/")[0])))
    except Exception:
        pass

    return meses


def _strip_datetime_time(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.select_dtypes(include=["datetime64[ns]", "datetime64[ns, UTC]"]).columns:
        df[col] = df[col].dt.date
    return df


def _prepare_df_for_export(
    df: pd.DataFrame,
    label: str = "",
) -> tuple[pd.DataFrame, Optional[str]]:
    df = _strip_datetime_time(df)
    df = _normalize_mes_column(df, label)
    df, scen_col = _normalize_cenario(df)
    df.columns = [str(c).strip() for c in df.columns]
    return df, scen_col


# ══════════════════════════════════════════════════════════════════
# Envio de e-mail (notificacao apenas)
# ══════════════════════════════════════════════════════════════════

def _send_report_email(
    gestor_nome: str,
    mes_anterior_pt: str,
    fy: int,
    today: str,
    result: ExecutionResult,
) -> None:
    email_dest = EMAIL_MAP.get(gestor_nome)
    if not email_dest:
        log.warning("Sem e-mail mapeado para: %s", gestor_nome)
        result.sem_email.append(gestor_nome)
        return

    context = {
        "responsavel": gestor_nome,
        "data": today,
        "link_bi": BI_LINK,
        "logo_cid": LOGO_CONTENT_ID,
        "mes_anterior_pt": mes_anterior_pt,
        "fy": fy,
    }

    try:
        send_email_sendgrid(
            email_dest,
            render_template(SUBJECT_TEMPLATE, context),
            render_template(BODY_TEMPLATE, context),
        )
        result.enviados.append(gestor_nome)
    except Exception as exc:
        log.error("Envio de e-mail falhou para %s: %s", gestor_nome, exc)
        result.falhas_email.append(gestor_nome)


# ══════════════════════════════════════════════════════════════════
# Fluxo: Base Completa
# ══════════════════════════════════════════════════════════════════

def process_base_completa(
    df: pd.DataFrame,
    gestor_nome: str,
    today: str,
    mes_anterior_pt: str,
    fy: int,
    result: ExecutionResult,
    conta_comment_map: Optional[dict] = None,
) -> None:
    log.info("[%s] Gerando BASE COMPLETA", gestor_nome)

    df_full, scen_col = _prepare_df_for_export(df, label=f"[{gestor_nome}]")
    allowed_months = _allowed_months_from_actual(df_full, scen_col)

    gestor_dir = os.path.join(GESTORES_DIR, sanitize_name(gestor_nome))
    ensure_dir(gestor_dir)

    out_file = (
        f"10.Despesas Gerais FY{fy:02d} ({mes_anterior_pt})"
        f" - {sanitize_filename_component(gestor_nome)}.xlsx"
    )
    out_path = os.path.join(gestor_dir, out_file)

    save_excel_com_retries(df_full, out_path)
    format_excel_date_columns(out_path, df_full)

    comment_map = conta_comment_map if gestor_nome in ALLOWED_COMMENTS_RECIPIENTS else None
    create_pivot_win32(
        out_path,
        sheet_pivot="Resumo",
        allowed_months=allowed_months,
        conta_comment_map=comment_map,
    )

    _send_report_email(gestor_nome, mes_anterior_pt, fy, today, result)


# ══════════════════════════════════════════════════════════════════
# Fluxo: Gestor Normal
# ══════════════════════════════════════════════════════════════════

def process_gestor_normal(
    df_resp: pd.DataFrame,
    gestor_nome: str,
    today: str,
    mes_anterior_pt: str,
    fy: int,
    result: ExecutionResult,
    conta_comment_map: Optional[dict] = None,
) -> None:
    if df_resp.empty:
        result.ignorados.append(gestor_nome)
        return

    df_resp, scen_col = _prepare_df_for_export(df_resp, label=f"[{gestor_nome}]")
    allowed_months = _allowed_months_from_actual(df_resp, scen_col)

    gestor_dir = os.path.join(GESTORES_DIR, sanitize_name(gestor_nome))
    ensure_dir(gestor_dir)

    out_file = (
        f"9.Despesas Gerais FY{fy:02d} ({mes_anterior_pt})"
        f" - {sanitize_filename_component(gestor_nome)}.xlsx"
    )
    out_path = os.path.join(gestor_dir, out_file)

    save_excel_com_retries(df_resp, out_path)
    time.sleep(0.3)
    format_excel_date_columns(out_path, df_resp)

    comment_map = conta_comment_map if gestor_nome in ALLOWED_COMMENTS_RECIPIENTS else None
    create_pivot_win32(
        out_path,
        sheet_pivot="Resumo",
        allowed_months=allowed_months,
        conta_comment_map=comment_map,
    )

    _send_report_email(gestor_nome, mes_anterior_pt, fy, today, result)