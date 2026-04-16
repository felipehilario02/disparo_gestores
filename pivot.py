"""
pivot.py
========
Criacao de Pivot Table via win32com (Excel COM automation).

Funcao publica:
    create_pivot_win32(excel_file_path, sheet_pivot, allowed_months, conta_comment_map)

Organizacao interna (funcoes privadas):
    _xl_helpers()           -> constantes e utilitarios COM reutilizaveis
    _build_pivot_structure  -> cria a pivot (campos, layout, refresh)
    _add_variance_column    -> coluna "Actual x Budget" fora da pivot
    _add_comments_column    -> coluna de comentarios por Conta
    _apply_formatting       -> alinhamento de labels e valores
"""

import re
import time
import unicodedata
from typing import Optional

from logger import get_logger

log = get_logger(__name__)


# ══════════════════════════════════════════════════════════════════
# Constantes Excel COM
# ══════════════════════════════════════════════════════════════════

class _XL:
    """Namespace para constantes da API COM do Excel."""
    Database    = 1
    RowField    = 1
    ColumnField = 2
    PageField   = 3
    Sum         = -4157
    TabularRow  = 1
    SrcRange    = 1
    Yes         = 1
    AtBottom    = 2
    Center      = -4108
    VCenter     = -4108
    Left        = -4131


# ══════════════════════════════════════════════════════════════════
# Helpers COM (independentes de instancia)
# ══════════════════════════════════════════════════════════════════

def _norm(s: str) -> str:
    return str(s).strip().lower()


def _strip_accents_lower(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return s.encode("ASCII", "ignore").decode("ASCII").strip().lower()


def _map_fields(pivot) -> dict:
    """Retorna {nome_normalizado: nome_original} de todos os PivotFields."""
    mapping: dict = {}
    try:
        for pf in pivot.PivotFields():
            try:
                name = str(pf.Name).strip()
                mapping[_norm(name)] = name
            except Exception:
                pass
    except Exception:
        try:
            for i in range(1, pivot.PivotFields().Count + 1):
                pf   = pivot.PivotFields(i)
                name = str(pf.Name).strip()
                mapping[_norm(name)] = name
        except Exception:
            pass
    return mapping


def _get_field(pivot, mapping: dict, candidates: list):
    """Retorna o primeiro PivotField que bata com algum candidato."""
    for cand in candidates:
        key = _norm(cand)
        if key in mapping:
            try:
                return pivot.PivotFields(mapping[key])
            except Exception:
                pass
    return None


def _find_header_col(ws, top_row, left_col, n_rows, n_cols, header_candidates):
    """Varre as primeiras linhas da pivot buscando um cabecalho especifico."""
    max_scan = min(top_row + 6, top_row + n_rows)
    for r in range(top_row, max_scan):
        for c in range(left_col, left_col + n_cols):
            v = ws.Cells(r, c).Value
            if not v:
                continue
            txt = str(v).strip()
            up  = txt.upper()
            for cand in header_candidates:
                if up == cand.upper():
                    return r, c
            for cand in header_candidates:
                tokens = [t for t in re.split(r"\s+", cand.upper()) if t]
                if tokens and all(t in up for t in tokens):
                    return r, c
    return None, None


def _a1_no_abs(cell) -> str:
    return str(cell.Address).replace("$", "")


def _parse_mm_yyyy(s: str):
    try:
        mm, yyyy = str(s).strip().split("/")
        return (int(yyyy), int(mm))
    except Exception:
        return None


def _order_pivot_items_by_chrono(pf) -> None:
    """Ordena PivotItems de um campo mes no formato 'mm/YYYY' cronologicamente."""
    if pf is None:
        return
    items = []
    try:
        for it in pf.PivotItems():
            nm  = str(it.Name).strip()
            key = _parse_mm_yyyy(nm)
            if key:
                items.append((key, nm))
    except Exception:
        pass
    if not items:
        return
    items.sort(key=lambda x: x[0])
    for pos, (_, nm) in enumerate(items, start=1):
        try:
            pf.PivotItems(nm).Position = pos
        except Exception:
            pass


def _norm_conta_key_com(val) -> Optional[str]:
    if val in (None, ""):
        return None
    s = str(val).strip()
    try:
        f = float(s.replace(",", "."))
        if abs(f - int(f)) < 1e-9:
            s = str(int(f))
    except Exception:
        pass
    s = s.replace("Total", "").replace("TOTAL", "").strip()
    s = re.sub(r"\s+", "", s)
    return s or None


def _is_total_or_subtotal_cell(cell_obj) -> bool:
    try:
        if cell_obj.PivotCell.PivotCellType in (3, 9, 10, 11, 12):
            return True
    except Exception:
        pass
    try:
        txt = str(cell_obj.Value)
        if txt and "total" in _strip_accents_lower(txt):
            return True
    except Exception:
        pass
    return False


def _is_valid_conta_text(val) -> bool:
    if val in (None, ""):
        return False
    s = _norm_conta_key_com(val)
    return bool(s and s.isdigit())


# ══════════════════════════════════════════════════════════════════
# Sub-rotinas da pivot (operam sobre objetos COM ja abertos)
# ══════════════════════════════════════════════════════════════════

def _build_pivot_structure(wb, ws_pivot, ws_data, allowed_months: Optional[list]):
    """
    Cria a pivot table completa: campos de linha, coluna, pagina, valor,
    layout, estilo e filtro de meses.

    Retorna o objeto pivot.
    """
    used      = ws_data.UsedRange
    rows      = used.Rows.Count
    cols      = used.Columns.Count
    src_range = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(rows, cols))

    # Recria ListObject "TblDados"
    try:
        for lo in ws_data.ListObjects:
            if lo.Name == "TblDados":
                lo.Unlist()
    except Exception:
        pass

    lo = ws_data.ListObjects.Add(
        SourceType=_XL.SrcRange,
        Source=src_range,
        XlListObjectHasHeaders=_XL.Yes,
    )
    lo.Name = "TblDados"

    cache = wb.PivotCaches().Create(
        SourceType=_XL.Database,
        SourceData="TblDados",
    )
    pivot = cache.CreatePivotTable(
        TableDestination=ws_pivot.Range("A3"),
        TableName="ResumoPivot",
    )
    pivot.ManualUpdate = True
    mapping = _map_fields(pivot)

    # ── Row fields ────────────────────────────────────────────────
    for group in [["GRUPO DESPESA"], ["CONTA"], ["DESCRICAO CONTA", "DESCRICAO CONTA"]]:
        pf = _get_field(pivot, mapping, group)
        if not pf:
            log.warning("Campo nao encontrado (linha): %s", group)
            continue
        pf.Orientation = _XL.RowField
        if _norm(group[0]) == "grupo despesa":
            pf.Subtotals = (False, True, False, False, False, False, False, False, False, False, False, False)
            try:
                pf.LayoutSubtotalLocation = _XL.AtBottom
            except Exception:
                pass
        else:
            pf.Subtotals = (False,) * 12

    mapping = _map_fields(pivot)

    # ── Column fields ─────────────────────────────────────────────
    pf_scen = _get_field(pivot, mapping, ["DATA BASE", "Cenario", "Scenario"])
    if pf_scen:
        pf_scen.Orientation = _XL.ColumnField
        pf_scen.Subtotals   = (True,) * 12
    else:
        log.warning("DATA BASE nao encontrado (coluna).")

    mapping = _map_fields(pivot)

    pf_mes = _get_field(pivot, mapping, ["MES", "Mes"])
    if pf_mes:
        pf_mes.Orientation = _XL.ColumnField
        pf_mes.Subtotals   = (False,) * 12
        try:
            _order_pivot_items_by_chrono(pf_mes)
        except Exception as exc:
            log.warning("Falha ao ordenar meses: %s", exc)
    else:
        log.warning("MES nao encontrado (coluna).")

    # ── Primeiro refresh ──────────────────────────────────────────
    pivot.ManualUpdate = False
    pivot.RefreshTable()
    time.sleep(0.25)
    pivot.ManualUpdate = True

    try:
        _order_pivot_items_by_chrono(_get_field(pivot, _map_fields(pivot), ["MES", "Mes"]))
    except Exception as exc:
        log.warning("Reordenacao pos-refresh falhou: %s", exc)

    # ── Ordenar cenarios BUDGET -> ACTUAL ─────────────────────────
    mapping = _map_fields(pivot)
    pf_scen = _get_field(pivot, mapping, ["DATA BASE", "Cenario", "Scenario"])
    if pf_scen:
        try:
            pf_scen.PivotItems("BUDGET").Position = 1
            pf_scen.PivotItems("ACTUAL").Position = 2
        except Exception as exc:
            log.warning("Nao foi possivel ordenar BUDGET->ACTUAL: %s", exc)

    # ── Page fields (filtros) ─────────────────────────────────────
    for fl in [["RESPONSAVEL", "Responsavel"], ["CCUSTO", "CENTRO DE CUSTO", "C.CUSTO"]]:
        pf = _get_field(pivot, mapping, fl)
        if pf:
            pf.Orientation = _XL.PageField

    pf_div = _get_field(pivot, mapping, ["DIVULGACAO"])
    if pf_div:
        try:
            pf_div.Orientation  = _XL.PageField
            pf_div.CurrentPage  = "SIM"
        except Exception:
            pass

    # ── Data field ────────────────────────────────────────────────
    pf_val = _get_field(pivot, mapping, ["BRL"])
    if pf_val:
        pivot.AddDataField(pf_val, "Soma BRL", _XL.Sum)
    else:
        log.warning("Campo BRL nao encontrado (valores).")

    try:
        for df_ in pivot.DataFields:
            df_.NumberFormat = "#.##0,00"
    except Exception as exc:
        log.warning("Falha ao formatar valores: %s", exc)

    # ── Layout e estilo ───────────────────────────────────────────
    pivot.RowAxisLayout(_XL.TabularRow)
    pivot.RowGrand    = False
    pivot.ColumnGrand = True

    try:
        pivot.TableStyle2 = "PivotStyleMedium9"
    except Exception:
        pass

    try:
        pivot.DisplayFieldCaptions = True
    except Exception:
        pass

    # ── Filtrar meses (opcional) ──────────────────────────────────
    if allowed_months:
        pf_mes2 = _get_field(pivot, _map_fields(pivot), ["MES", "Mes"])
        if pf_mes2:
            allowed = set(str(x).strip() for x in allowed_months)
            try:
                for it in pf_mes2.PivotItems():
                    try:
                        it.Visible = True
                    except Exception:
                        pass
            except Exception:
                pass

            any_visible  = False
            last_visible = None
            for it in pf_mes2.PivotItems():
                nm = str(it.Name).strip()
                ok = nm in allowed
                try:
                    it.Visible = ok
                    if ok:
                        any_visible  = True
                        last_visible = it
                except Exception:
                    pass

            if not any_visible and last_visible:
                try:
                    last_visible.Visible = True
                except Exception:
                    pass

            try:
                _order_pivot_items_by_chrono(pf_mes2)
            except Exception as exc:
                log.warning("Falha ao reordenar meses apos filtro: %s", exc)

    # ── Refresh final ─────────────────────────────────────────────
    pivot.ManualUpdate = False
    pivot.RefreshTable()
    time.sleep(0.25)
    pivot.ManualUpdate = True

    try:
        _order_pivot_items_by_chrono(_get_field(pivot, _map_fields(pivot), ["MES", "Mes"]))
    except Exception as exc:
        log.warning("Reordenacao final falhou: %s", exc)

    return pivot


def _add_variance_column(ws_pivot, pivot) -> tuple[Optional[int], Optional[int]]:
    """
    Adiciona a coluna 'Actual x Budget' apos a coluna 'ACTUAL Total'.
    Retorna (col_var_total, header_row) ou (None, None) se nao for possivel.
    """
    try:
        tbl      = pivot.TableRange2
        top_row  = tbl.Row
        left_col = tbl.Column
        n_rows   = tbl.Rows.Count
        n_cols   = tbl.Columns.Count

        hr_b, col_budget = _find_header_col(
            ws_pivot, top_row, left_col, n_rows, n_cols,
            ["BUDGET Total", "TOTAL BUDGET", "BUDGET TOTAL"],
        )
        hr_a, col_actual = _find_header_col(
            ws_pivot, top_row, left_col, n_rows, n_cols,
            ["ACTUAL Total", "TOTAL ACTUAL", "ACTUAL TOTAL"],
        )

        if not col_budget or not col_actual:
            log.warning("'BUDGET Total' e/ou 'ACTUAL Total' nao encontrados.")
            return None, None

        header_row    = max(hr_b, hr_a)
        col_var_total = col_actual + 1

        ws_pivot.Cells(header_row, col_var_total).Value     = "Actual x Budget"
        ws_pivot.Cells(header_row, col_var_total).Font.Bold = True

        first_data = header_row + 1
        last_row   = top_row + n_rows - 1

        for r in range(first_data, last_row + 1):
            b = ws_pivot.Cells(r, col_budget).Value
            a = ws_pivot.Cells(r, col_actual).Value
            if b is None and a is None:
                continue
            addr_a = _a1_no_abs(ws_pivot.Cells(r, col_actual))
            addr_b = _a1_no_abs(ws_pivot.Cells(r, col_budget))
            ws_pivot.Cells(r, col_var_total).Formula      = f"={addr_a}-{addr_b}"
            ws_pivot.Cells(r, col_var_total).NumberFormat = "#.##0,00"

        try:
            ws_pivot.Columns(col_var_total).AutoFit()
        except Exception:
            pass

        log.info("Coluna 'Actual x Budget' criada.")
        return col_var_total, header_row

    except Exception as exc:
        log.warning("Falha ao criar coluna Actual x Budget: %s", exc)
        return None, None


def _add_comments_column(
    ws_pivot,
    pivot,
    conta_comment_map: dict,
    col_var_total: int,
    header_row: int,
) -> None:
    """
    Adiciona a coluna 'Comentarios Bud. vs Act.' apos a coluna de variancia.
    Opera linha a linha sobre o DataBodyRange da pivot.
    """
    try:
        pf_conta = _get_field(pivot, _map_fields(pivot), ["CONTA", "Conta"])
        if not pf_conta:
            log.warning("Campo 'CONTA' nao esta como RowField.")
            return

        try:
            pivot.DisplayFieldCaptions = True
        except Exception:
            pass

        try:
            col_conta = int(pf_conta.LabelRange.Column)
        except Exception:
            log.warning("LabelRange para 'CONTA' nao disponivel.")
            return

        col_comment = col_var_total + 1
        ws_pivot.Cells(header_row, col_comment).Value     = "Comentarios Bud. vs Act."
        ws_pivot.Cells(header_row, col_comment).Font.Bold = True

        try:
            dbr            = pivot.DataBodyRange
            first_data_row = int(dbr.Row)
            last_row       = first_data_row + int(dbr.Rows.Count) - 1
        except Exception:
            tbl            = pivot.TableRange2
            first_data_row = header_row + 1
            last_row       = tbl.Row + tbl.Rows.Count - 1

        for r in range(first_data_row, last_row + 1):
            cell_conta = ws_pivot.Cells(r, col_conta)

            try:
                if _is_total_or_subtotal_cell(cell_conta):
                    continue
            except Exception:
                pass

            raw_val = cell_conta.Value
            if not _is_valid_conta_text(raw_val):
                continue

            conta_corrente = _norm_conta_key_com(raw_val)
            if not conta_corrente:
                continue

            comentario = conta_comment_map.get(conta_corrente)
            if not comentario and conta_corrente.isdigit():
                for z in (8, 7, 6):
                    comentario = conta_comment_map.get(conta_corrente.zfill(z))
                    if comentario:
                        break

            if comentario:
                ws_pivot.Cells(r, col_comment).Value = comentario

        try:
            ws_pivot.Columns(col_comment).AutoFit()
        except Exception:
            pass

        try:
            rng = ws_pivot.Range(
                ws_pivot.Cells(header_row, col_comment),
                ws_pivot.Cells(last_row, col_comment),
            )
            rng.HorizontalAlignment = _XL.Left
            rng.WrapText            = True
        except Exception:
            pass

        log.info("Coluna 'Comentarios Bud. vs Act.' criada.")

    except Exception as exc:
        log.warning("Falha ao criar coluna de comentarios: %s", exc)


def _apply_formatting(ws_pivot, pivot, col_var_total: Optional[int]) -> None:
    """
    Alinha labels (esquerda) e valores numericos (centro) na pivot.
    """
    try:
        tbl      = pivot.TableRange2
        top_row  = tbl.Row
        left_col = tbl.Column
        n_rows   = tbl.Rows.Count
        n_cols   = tbl.Columns.Count

        try:
            row_fields_count = pivot.RowFields().Count
        except Exception:
            row_fields_count = 3

        first_value_col = left_col + max(row_fields_count, 1)
        last_pivot_col  = left_col + n_cols - 1
        last_value_col  = last_pivot_col
        if col_var_total is not None:
            last_value_col = max(last_value_col, col_var_total)

        first_row = top_row
        last_row  = top_row + n_rows - 1
        label_end = first_value_col - 1

        if label_end >= left_col:
            rng = ws_pivot.Range(
                ws_pivot.Cells(first_row, left_col),
                ws_pivot.Cells(last_row, label_end),
            )
            rng.HorizontalAlignment = _XL.Left
            rng.VerticalAlignment   = _XL.VCenter

        if first_value_col <= last_value_col:
            rng = ws_pivot.Range(
                ws_pivot.Cells(first_row, first_value_col),
                ws_pivot.Cells(last_row, last_value_col),
            )
            rng.HorizontalAlignment = _XL.Center
            rng.VerticalAlignment   = _XL.VCenter

        log.info("Formatacao de alinhamento aplicada.")

    except Exception as exc:
        log.warning("Falha ao alinhar: %s", exc)


# ══════════════════════════════════════════════════════════════════
# Funcao publica
# ══════════════════════════════════════════════════════════════════

def create_pivot_win32(
    excel_file_path: str,
    sheet_pivot: str = "Resumo",
    allowed_months: Optional[list] = None,
    conta_comment_map: Optional[dict] = None,
) -> None:
    """
    Abre o arquivo Excel via COM, (re)cria a aba de Pivot Table
    e aplica colunas derivadas e formatacao.
    """
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible       = False
    excel.DisplayAlerts = False
    wb = None

    try:
        wb = excel.Workbooks.Open(excel_file_path)

        try:
            excel.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass

        # Remove aba pivot anterior se existir
        try:
            wb.Worksheets(sheet_pivot).Delete()
        except Exception:
            pass

        # Sheet de dados
        try:
            ws_data = wb.Worksheets("Sheet1")
        except Exception:
            ws_data = wb.Worksheets(1)

        # Nova aba pivot
        ws_pivot = wb.Worksheets.Add()
        ws_pivot.Name = sheet_pivot
        try:
            ws_pivot.Activate()
            excel.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass

        # 1) Estrutura da pivot
        pivot = _build_pivot_structure(wb, ws_pivot, ws_data, allowed_months)

        # 2) Coluna de variancia
        col_var_total, header_row = _add_variance_column(ws_pivot, pivot)

        # 3) Coluna de comentarios (somente se mapa fornecido e coluna de variancia criada)
        if conta_comment_map and col_var_total is not None and header_row is not None:
            _add_comments_column(ws_pivot, pivot, conta_comment_map, col_var_total, header_row)
        elif not conta_comment_map:
            log.info("Mapa de comentarios nao informado; coluna nao sera criada.")
        elif col_var_total is None:
            log.warning("'Actual x Budget' ausente; coluna de comentarios nao posicionada.")

        # 4) Formatacao
        _apply_formatting(ws_pivot, pivot, col_var_total)

        wb.Save()
        log.info("Pivot salva: %s", excel_file_path)

    except Exception as exc:
        log.error("create_pivot_win32 falhou: %s", exc)

    finally:
        try:
            if wb:
                wb.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
