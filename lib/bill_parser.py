"""Parse Meitong-style logistics freight bills (xlsx)."""
from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any
import re

from openpyxl import load_workbook


HEADER_KEYS = ("运单号", "费用类型")
# Allocation inputs (Meitong cleaned sheet letter positions):
#   G 扩展单号 → FBA list, K 收费重 → weight, R 金额 → CNY freight
WEIGHT_ALIASES = ("收费重(KG/CBM)", "收费重", "计费重", "计费重量", "收费重量")
AMOUNT_CNY_ALIASES = ("金额",)  # column R; do NOT use 折合人民币 for allocation
AMOUNT_ALIASES = ("金额",)
EXT_ALIASES = ("扩展单号",)
FBA_TOKEN_RE = re.compile(r"FBA[A-Z0-9]+", re.IGNORECASE)
# Footer / bank lines that appear under detail tables (not fee rows).
_JUNK_WAYBILL_MARKERS = (
    "收款",
    "账户",
    "账号",
    "公司名称",
    "开户银行",
    "备注",
    "合计",
)


@dataclass
class FreightRow:
    row_index: int
    waybill_no: str  # 运单号
    fee_type: str
    date: str
    customer_no: str  # 客户单号
    tracking_no: str  # 转单号
    ext_nos: str  # 扩展单号 (G列，货件号来源)
    service_type: str
    warehouse: str
    zipcode: str
    chargeable_weight: float
    unit_price: float | None
    unit: str
    quantity: float | None
    pieces: float | None
    country: str
    currency: str
    amount: float
    exchange_rate: float | None
    amount_cny: float
    status: str
    description: str
    customs_mode: str
    fba_ids: list[str]  # from 扩展单号 (prefer) / 客户单号 fallback
    shop: str = ""
    sheet: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _norm(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _to_float(v: Any) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _find_header_row(ws) -> tuple[int, dict[str, int]] | None:
    """Return 1-based header row and column name -> 1-based index map."""
    max_scan = min(ws.max_row or 1, 40)
    max_col = min(ws.max_column or 1, 40)
    for r in range(1, max_scan + 1):
        cells = {}
        for c in range(1, max_col + 1):
            val = _norm(ws.cell(r, c).value)
            if val:
                cells[val] = c
        if "运单号" in cells and "费用类型" in cells:
            return r, cells
    return None


def _pick_col(cols: dict[str, int], aliases: tuple[str, ...]) -> int | None:
    for a in aliases:
        if a in cols:
            return cols[a]
    for name, idx in cols.items():
        for a in aliases:
            if a in name:
                return idx
    return None


def _extract_fba_tokens(text: str) -> list[str]:
    """Extract FBA… tokens in order, de-duplicated."""
    ids: list[str] = []
    seen: set[str] = set()
    for m in FBA_TOKEN_RE.finditer(text or ""):
        tok = m.group(0).upper()
        # preserve common casing FBA…
        raw = m.group(0)
        if raw.upper().startswith("FBA"):
            tok = "FBA" + raw[3:]
        if tok not in seen:
            seen.add(tok)
            ids.append(tok)
    return ids


def _parse_fba_ids(ext_nos: str, customer_no: str = "") -> list[str]:
    """FBA list from G列扩展单号 only (allocation ignores 客户单号)."""
    from_ext = _extract_fba_tokens(ext_nos)
    if from_ext:
        return from_ext
    # last resort: whole 扩展单号 cell if it looks like a single id
    s = _norm(ext_nos)
    if s and " " not in s and "," not in s and "，" not in s:
        return [s]
    return []


def _is_junk_waybill(waybill: str) -> bool:
    compact = waybill.replace(" ", "").replace("\u3000", "")
    for m in _JUNK_WAYBILL_MARKERS:
        if m in waybill or m in compact:
            return True
    return False


def _first_row_vals(ws, max_col: int = 25) -> list[str]:
    """Return normalized values from the first worksheet row."""
    for row in ws.iter_rows(min_row=1, max_row=1, max_col=max_col, values_only=True):
        return [_norm(v) for v in row]
    return []


def _is_cleaned_header_row(vals: list[str]) -> bool:
    """True when row 1 already looks like column headers (user-cleaned sheet).

    Cleaned sheets start with 序号 or have 运单号 in the first row.
    Raw Meitong exports start with a company title (合肥美通…) instead.
    """
    if not vals:
        return False
    first = vals[0] if vals else ""
    if first == "序号":
        return True
    if "运单号" in vals:
        return True
    return False


def _sheet_has_detail_headers(ws) -> bool:
    """True if 运单号 + 费用类型 appear anywhere in the first 40 rows."""
    for row in ws.iter_rows(max_row=40, max_col=25, values_only=True):
        vals = [_norm(v) for v in row]
        if "运单号" in vals and "费用类型" in vals:
            return True
    return False


def detect_detail_sheets(path: str) -> list[str]:
    """Prefer cleaned sheets (row-1 headers); else all sheets with detail headers."""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        cleaned: list[str] = []
        detail: list[str] = []
        for name in wb.sheetnames:
            if "汇总" in name or "summary" in name.lower():
                continue
            ws = wb[name]
            if _is_cleaned_header_row(_first_row_vals(ws)):
                cleaned.append(name)
            elif _sheet_has_detail_headers(ws):
                detail.append(name)
        return cleaned if cleaned else detail
    finally:
        wb.close()


def _select_sheets(wb, sheet_name: str | None = None) -> tuple[list[str], str]:
    """Pick which sheets to parse.

    Preference:
      1. Explicit sheet_name if provided
      2. Sheets whose first row is already cleaned column headers (序号 / 运单号)
         — if any, use only those (single or merge multiple cleaned)
      3. Fallback: all non-汇总 sheets that contain detail headers anywhere
    """
    if sheet_name and sheet_name in wb.sheetnames:
        return [sheet_name], "explicit"

    candidates = [
        n
        for n in wb.sheetnames
        if "汇总" not in n and "summary" not in n.lower()
    ]
    if not candidates:
        candidates = list(wb.sheetnames)

    cleaned: list[str] = []
    detail: list[str] = []
    for name in candidates:
        ws = wb[name]
        if _is_cleaned_header_row(_first_row_vals(ws)):
            cleaned.append(name)
        elif _find_header_row(ws):
            detail.append(name)

    if cleaned:
        return cleaned, "cleaned"
    if detail:
        return detail, "detail"
    return candidates, "all"


def parse_bill(path: str, shop: str = "", sheet_name: str | None = None) -> dict[str, Any]:
    """Parse a Meitong-style bill. Prefer user-cleaned sheet(s); no fee-type filter.

    Allocation uses only three bill columns (plus 运单号 for ticket grouping):
      G 扩展单号 → FBA list, K 收费重 → weight, R 金额 → CNY freight.
    Other columns are ignored. User should delete 税金/清关费 rows before upload.
    """
    wb = load_workbook(path, data_only=True)
    try:
        sheets, sheet_select = _select_sheets(wb, sheet_name)

        all_rows: list[FreightRow] = []
        used_sheets: list[str] = []
        header_rows: dict[str, int] = {}

        for name in sheets:
            ws = wb[name]
            found = _find_header_row(ws)
            if not found:
                continue
            header_row, cols = found
            used_sheets.append(name)
            header_rows[name] = header_row
            # Structural: 运单号 groups tickets. Allocation inputs only:
            # G 扩展单号, K 收费重, R 金额 — ignore other bill columns.
            col_waybill = cols["运单号"]
            col_fee = cols.get("费用类型")  # only used to skip 合计 footer rows
            col_ext = _pick_col(cols, EXT_ALIASES) or cols.get("扩展单号")
            col_weight = _pick_col(cols, WEIGHT_ALIASES)
            col_amount = _pick_col(cols, AMOUNT_ALIASES)  # R 金额 (CNY)
            # Cleaned Meitong layout fallback by letter when headers missing aliases
            if col_ext is None and header_row == 1:
                col_ext = 7  # G
            if col_weight is None and header_row == 1:
                col_weight = 11  # K
            if col_amount is None and header_row == 1:
                col_amount = 18  # R

            for r in range(header_row + 1, (ws.max_row or header_row) + 1):
                waybill = _norm(ws.cell(r, col_waybill).value)
                fee = _norm(ws.cell(r, col_fee).value) if col_fee else ""

                if not waybill:
                    continue
                if _is_junk_waybill(waybill):
                    continue
                if fee and "合计" in fee:
                    continue

                weight = _to_float(ws.cell(r, col_weight).value) if col_weight else None
                amount_cny = _to_float(ws.cell(r, col_amount).value) if col_amount else None
                if amount_cny is None:
                    amount_cny = 0.0

                ext_nos = _norm(ws.cell(r, col_ext).value) if col_ext else ""

                row = FreightRow(
                    row_index=r,
                    waybill_no=waybill,
                    fee_type=fee or "运费",
                    date="",
                    customer_no="",
                    tracking_no="",
                    ext_nos=ext_nos,
                    service_type="",
                    warehouse="",
                    zipcode="",
                    chargeable_weight=weight or 0.0,
                    unit_price=None,
                    unit="",
                    quantity=None,
                    pieces=None,
                    country="",
                    currency="CNY",
                    amount=amount_cny,
                    exchange_rate=None,
                    amount_cny=amount_cny,
                    status="",
                    description="",
                    customs_mode="",
                    fba_ids=_parse_fba_ids(ext_nos),
                    shop=shop,
                    sheet=name,
                )
                all_rows.append(row)

        tickets: dict[str, dict[str, Any]] = {}
        for row in all_rows:
            t = tickets.setdefault(
                row.waybill_no,
                {
                    "waybill_no": row.waybill_no,
                    "freight_total": 0.0,
                    "weight_total": 0.0,
                    "line_count": 0,
                    "fba_ids": [],
                    "customer_nos": [],
                },
            )
            t["freight_total"] += row.amount_cny
            # Same waybill often repeats chargeable weight on 清关费/税金 lines;
            # take max so weight is not double-counted across fee rows.
            t["weight_total"] = max(t["weight_total"], row.chargeable_weight or 0.0)
            t["line_count"] += 1
            for fid in row.fba_ids:
                if fid not in t["fba_ids"]:
                    t["fba_ids"].append(fid)
            if row.customer_no and row.customer_no not in t["customer_nos"]:
                t["customer_nos"].append(row.customer_no)

        total_freight = sum(r.amount_cny for r in all_rows)
        # Prefer sum of per-ticket weights (avoids multi-fee double count)
        total_weight = sum(t["weight_total"] for t in tickets.values())

        sheet_label = " + ".join(used_sheets) if used_sheets else None
        header_row = header_rows.get(used_sheets[0]) if used_sheets else None

        return {
            "ok": True,
            "sheet": sheet_label,
            "sheets": used_sheets,
            "sheet_select": sheet_select,
            "header_row": header_row,
            "header_rows": header_rows,
            "shop": shop,
            "freight_row_count": len(all_rows),
            "ticket_count": len(tickets),
            "total_freight": round(total_freight, 4),
            "total_weight": round(total_weight, 4),
            "skipped_fee_types": {},
            "rows": [r.to_dict() for r in all_rows],
            "tickets": list(tickets.values()),
        }
    finally:
        wb.close()
