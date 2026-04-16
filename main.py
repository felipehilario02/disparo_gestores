"""
main.py
=======
Pipeline de distribuicao de despesas.

- Mantém o fluxo AUTOMÁTICO original.
- Permite definir um gestor PRIORITÁRIO (primeiro a receber).
- Respeita integralmente a regra de negócio FULL_BASE_RECIPIENTS
  (ex.: Izabella Barretti recebe base completa mesmo fora da base).
"""

import os
import sys
from datetime import datetime

from comentarios import carregar_base_comentarios
from config import (
    BASE_DIR,
    COMENTARIOS_FILE,
    COMENTARIOS_SHEET,
    DATE_FMT,
    DEBUG_SAMPLE_CHECK,
    FORCE_RESP_COL_NAME,
    FULL_BASE_RECIPIENTS,
    INPUT_FILE,
    SHEET_NAME,
)
from helpers import find_col, mes_anterior_pt_e_fy
from io_excel import read_excel_robusto
from logger import get_logger
from processor import ExecutionResult, process_base_completa, process_gestor_normal

log = get_logger(__name__)

# =====================================================
# CONFIGURAÇÃO DE PRIORIDADE (ORDEM, NÃO REGRA)
# =====================================================
PRIORITY_GESTOR = "Izabella Barretti"

# ─────────────────────────────────────────────
# Debug opcional
# ─────────────────────────────────────────────
def _debug_sample_check(df, conta_comments: dict) -> None:
    import re

    conta_col = find_col(df.columns, ["CONTA", "Conta"])
    if not conta_col:
        return

    contas = df[conta_col].astype(str).str.strip().dropna().unique().tolist()[:10]
    log.debug("Amostra de contas na base: %s", contas)

    hits = []
    for c in contas:
        key = re.sub(r"\s+", "", c).replace("Total", "").replace("TOTAL", "").strip()
        try:
            f = float(key.replace(",", "."))
            key = str(int(f)) if abs(f - int(f)) < 1e-9 else key
        except Exception:
            pass
        val = conta_comments.get(key)
        if not val and key.isdigit():
            for z in (8, 7, 6):
                val = conta_comments.get(key.zfill(z))
                if val:
                    break
        hits.append((c, bool(val)))

    log.debug("Comentario encontrado por conta (amostra): %s", hits)


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def main() -> None:
    log.info("=" * 55)
    log.info("  PIPELINE DE DESPESAS — INICIANDO")
    log.info("=" * 55)

    base_path = os.path.join(BASE_DIR, INPUT_FILE)
    if not os.path.exists(base_path):
        log.error("Arquivo nao encontrado: %s", base_path)
        sys.exit(1)

    # 1) Comentários
    conta_comments = carregar_base_comentarios(
        alt_file=COMENTARIOS_FILE,
        sheet_name=COMENTARIOS_SHEET,
    )

    # 2) Base principal
    df = read_excel_robusto(base_path, sheet_name=SHEET_NAME, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # 3) Coluna de responsável
    col_resp = (
        FORCE_RESP_COL_NAME
        if FORCE_RESP_COL_NAME in df.columns
        else df.columns[9]
    )

    # 4) Limpeza básica
    df = df[df[col_resp].notna()].copy()
    df[col_resp] = df[col_resp].astype(str).str.strip()
    df = df[df[col_resp].str.len() > 0]
    df = df[~df[col_resp].str.lower().isin(["nan", "none", "null"])]

    if DEBUG_SAMPLE_CHECK:
        _debug_sample_check(df, conta_comments)

    today = datetime.today().strftime(DATE_FMT)
    mes_anterior_pt, fy = mes_anterior_pt_e_fy()

    result = ExecutionResult()
    processed_full: set[str] = set()

    # =====================================================
    # ORDEM DE PROCESSAMENTO (COM PRIORIDADE)
    # =====================================================
    gestores_base = (
        df[col_resp]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
        .tolist()
    )

    gestores = sorted(set(gestores_base) | set(FULL_BASE_RECIPIENTS))

    if PRIORITY_GESTOR and PRIORITY_GESTOR in gestores:
        gestores.remove(PRIORITY_GESTOR)
        gestores.insert(0, PRIORITY_GESTOR)
        log.info("Gestor prioritário (primeiro envio): %s", PRIORITY_GESTOR)

    log.info("Ordem final de processamento: %s", gestores)

    # =====================================================
    # FLUXO AUTOMÁTICO (INALTERADO)
    # =====================================================
    for gestor_nome in gestores:

        # ✅ REGRA DE NEGÓCIO SEMPRE PRIMEIRO
        if gestor_nome in FULL_BASE_RECIPIENTS:
            process_base_completa(
                df,
                gestor_nome,
                today,
                mes_anterior_pt,
                fy,
                result=result,
                conta_comment_map=conta_comments,
            )
            processed_full.add(gestor_nome)

        # ✅ GESTORES NORMAIS DEPENDEM DA BASE
        else:
            df_resp = df[df[col_resp] == gestor_nome]

            if df_resp.empty:
                log.warning(
                    "[%s] Não encontrado na base e não pertence a FULL_BASE_RECIPIENTS.",
                    gestor_nome,
                )
                continue

            process_gestor_normal(
                df_resp,
                gestor_nome,
                today,
                mes_anterior_pt,
                fy,
                result=result,
                conta_comment_map=conta_comments,
            )

    # =====================================================
    # RESUMO FINAL
    # =====================================================
    log.info(result.resumo())
    log.info("PIPELINE CONCLUIDO.")


if __name__ == "__main__":
    main()