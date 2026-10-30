"""Flexible reader for Saihu (赛狐) FBA shipment exports."""
from __future__ import annotations

from typing import Any

from openpyxl import load_workbook


# Heuristic aliases for auto-detect (Chinese / English)
FBA_ALIASES = [
    "货件号", "FBA货件号", "FBA单号", "Shipment ID", "ShipmentId", "shipment_id",
    "亚马逊货件号", "FBA shipment", "ShipmentID", "货件编号",
]
SKU_ALIASES = [
    "SKU", "卖家SKU", "MSKU", "Seller SKU", "商家SKU", "sku", "产品SKU",
]
QTY_ALIASES = [
    "数量", "发货数量", "Quantity", "Qty", "件数", "发货件数", "计划发货量", "实际发货量",
]
UNIT_WEIGHT_ALIASES = [
    "单件重量", "单重", "单位重量", "Unit Weight", "unit_weight", "产品重量", "毛重(单件)",
]
TOTAL_WEIGHT_ALIASES = [
    "总重量", "重量", "Weight", "合计重量", "发货重量", "计费重量", "毛重",
]


def _norm(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def peek_columns(path: str, max_header_scan: int = 15) -> dict[str, Any]:
    """Return sheet names, detected header row, column list, sample rows, suggested mapping."""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        sheets_info = []
        for name in wb.sheetnames:
            ws = wb[name]
            rows = []
            for i, row in enumerate(ws.iter_rows(max_row=30, max_col=40, values_only=True)):
                rows.append([_norm(v) for v in row])
                if i >= 29:
                    break
            header_row_idx = 0
            cols: list[str] = []
            # find row with most non-empty + looks like headers
            best = (-1, 0, [])
            for i, row in enumerate(rows[:max_header_scan]):
                nonempty = [c for c in row if c]
                score = len(nonempty)
                # bonus if known aliases present
                joined = set(nonempty)
                for alias_list in (FBA_ALIASES, SKU_ALIASES, QTY_ALIASES):
                    if any(a in joined for a in alias_list):
                        score += 5
                if score > best[1]:
                    best = (i, score, nonempty)
            header_row_idx, _, cols = best
            sample = []
            for row in rows[header_row_idx + 1 : header_row_idx + 6]:
                if any(row):
                    # align to cols length
                    sample.append(row[: len(cols)] if cols else row)
            suggested = suggest_mapping(cols)
            sheets_info.append(
                {
                    "name": name,
                    "header_row": header_row_idx + 1,  # 1-based for UI
                    "columns": cols,
                    "sample": sample,
                    "suggested_mapping": suggested,
                }
            )
        return {"ok": True, "sheets": sheets_info}
    finally:
        wb.close()


def suggest_mapping(columns: list[str]) -> dict[str, str | None]:
    def find(aliases: list[str], exclude: set[str] | None = None) -> str | None:
        exclude = exclude or set()
        candidates = [c for c in columns if c not in exclude]
        for c in candidates:
            if c in aliases:
                return c
        lower_map = {c.lower(): c for c in candidates}
        for a in aliases:
            if a.lower() in lower_map:
                return lower_map[a.lower()]
        for c in candidates:
            for a in aliases:
                if a.lower() in c.lower():
                    return c
        return None

    unit_w = find(UNIT_WEIGHT_ALIASES)
    # Prefer dedicated unit-weight; only suggest total_weight on a different column
    total_w = find(TOTAL_WEIGHT_ALIASES, exclude={unit_w} if unit_w else set())
    # Avoid matching bare "单件重量" as total via substring "重量"
    if total_w and unit_w and total_w == unit_w:
        total_w = None
    return {
        "fba_id": find(FBA_ALIASES),
        "sku": find(SKU_ALIASES),
        "qty": find(QTY_ALIASES),
        "unit_weight": unit_w,
        "total_weight": total_w,
    }


def parse_saihu(
    path: str,
    mapping: dict[str, str],
    sheet_name: str | None = None,
    header_row: int = 1,
    weight_mode: str = "unit",  # "unit" | "total"
) -> dict[str, Any]:
    """
    mapping keys: fba_id, sku, qty, and either unit_weight or total_weight column name.
    header_row is 1-based.
    """
    wb = load_workbook(path, data_only=True)
    try:
        name = sheet_name or wb.sheetnames[0]
        ws = wb[name]
        # Build header -> col index from the given header row
        headers = {}
        max_col = ws.max_column or 1
        for c in range(1, max_col + 1):
            val = _norm(ws.cell(header_row, c).value)
            if val:
                headers[val] = c

        required = ["fba_id", "sku", "qty"]
        for k in required:
            col_name = mapping.get(k)
            if not col_name or col_name not in headers:
                return {
                    "ok": False,
                    "error": f"缺少列映射: {k} -> {col_name}（表头中未找到）",
                    "available_columns": list(headers.keys()),
                }

        weight_col_name = mapping.get("unit_weight") or mapping.get("total_weight")
        if not weight_col_name or weight_col_name not in headers:
            return {
                "ok": False,
                "error": "需要映射单件重量或总重量列",
                "available_columns": list(headers.keys()),
            }

        weight_is_total = weight_mode == "total" or (
            mapping.get("total_weight") == weight_col_name and not mapping.get("unit_weight")
        )

        rows = []
        for r in range(header_row + 1, (ws.max_row or header_row) + 1):
            fba = _norm(ws.cell(r, headers[mapping["fba_id"]]).value)
            sku = _norm(ws.cell(r, headers[mapping["sku"]]).value)
            if not fba and not sku:
                continue
            qty_raw = ws.cell(r, headers[mapping["qty"]]).value
            w_raw = ws.cell(r, headers[weight_col_name]).value
            try:
                qty = float(qty_raw or 0)
            except (TypeError, ValueError):
                qty = 0
            try:
                w = float(str(w_raw).replace(",", "") if w_raw is not None else 0)
            except (TypeError, ValueError):
                w = 0
            rows.append(
                {
                    "fba_id": fba,
                    "sku": sku,
                    "qty": qty,
                    "unit_weight": w,
                    "weight_is_total": weight_is_total,
                }
            )

        return {
            "ok": True,
            "sheet": name,
            "row_count": len(rows),
            "fba_count": len({x["fba_id"] for x in rows if x["fba_id"]}),
            "sku_count": len({x["sku"] for x in rows if x["sku"]}),
            "weight_is_total": weight_is_total,
            "rows": rows,
        }
    finally:
        wb.close()
