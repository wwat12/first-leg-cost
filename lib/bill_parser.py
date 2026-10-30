"""Parse Meitong-style logistics freight bills (xlsx)."""
from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any
import re

from openpyxl import load_workbook


HEADER_KEYS = ("运单号", "费用类型")
WEIGHT_ALIASES = ("收费重(KG/CBM)", "收费重", "计费重", "计费重量", "收费重量")
AMOUNT_CNY_ALIASES = ("折合人民币", "折算人民币", "本位币金额")
AMOUNT_ALIASES = ("金额",)
FBA_ALIASES = ("扩展单号", "客户单号")


@dataclass
class FreightRow:
    row_index: int
    waybill_no: str          # 运单号
    fee_type: str
    date: str
    customer_no: str         # 客户单号 (primary FBA)
    tracking_no: str         # 转单号
    ext_nos: str             # 扩展单号 (comma-separated FBA ids)
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
    fba_ids: list[str]       # parsed FBA shipment ids
    shop: str = ""

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        return d


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
    # fuzzy: contain
    for name, idx in cols.items():
        for a in aliases:
            if a in name:
                return idx
    return None


def _parse_fba_ids(customer_no: str, ext_nos: str) -> list[str]:
    ids: list[str] = []
    seen: set[str] = set()
    # Prefer 扩展单号 when present; also include 客户单号 if missing from list
    raw = ext_nos or customer_no
    parts = re.split(r"[,，;；\s]+", raw)
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if p not in seen:
            seen.add(p)
            ids.append(p)
    if customer_no and customer_no not in seen:
        ids.insert(0, customer_no)
    return ids


def detect_detail_sheets(path: str) -> list[str]:
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        names = []
        for name in wb.sheetnames:
            if "汇总" in name or "summary" in name.lower():
                continue
            ws = wb[name]
            # peek first 40 rows
            found = False
            for i, row in enumerate(ws.iter_rows(max_row=40, max_col=25, values_only=True)):
                vals = [_norm(v) for v in row]
                if "运单号" in vals and "费用类型" in vals:
                    found = True
                    break
            if found:
                names.append(name)
        return names
    finally:
        wb.close()


def parse_bill(path: str, shop: str = "", sheet_name: str | None = None) -> dict[str, Any]:
    """Parse a Meitong-style bill. Returns summary + freight rows."""
    wb = load_workbook(path, data_only=True)
    try:
        if sheet_name and sheet_name in wb.sheetnames:
            sheets = [sheet_name]
        else:
            sheets = [n for n in wb.sheetnames if "汇总" not in n]
            if not sheets:
                sheets = list(wb.sheetnames)

        all_rows: list[FreightRow] = []
        used_sheet = None
        header_row = None
        skipped_fee_types: dict[str, int] = {}
        skipped_other = 0

        for name in sheets:
            ws = wb[name]
            found = _find_header_row(ws)
            if not found:
                continue
            used_sheet = name
            header_row, cols = found
            col_waybill = cols["运单号"]
            col_fee = cols["费用类型"]
            col_date = cols.get("日期")
            col_customer = cols.get("客户单号")
            col_tracking = cols.get("转单号")
            col_ext = cols.get("扩展单号")
            col_service = cols.get("服务类型")
            col_wh = cols.get("仓库代码")
            col_zip = cols.get("邮编")
            col_weight = _pick_col(cols, WEIGHT_ALIASES)
            col_price = cols.get("单价")
            col_unit = cols.get("单价单位")
            col_qty = cols.get("数量")
            col_pieces = cols.get("件数")
            col_country = cols.get("发往国家")
            col_currency = cols.get("币种")
            col_amount = _pick_col(cols, AMOUNT_ALIASES)
            col_rate = cols.get("汇率")
            col_cny = _pick_col(cols, AMOUNT_CNY_ALIASES)
            col_status = cols.get("运单状态")
            col_desc = cols.get("描述")
            col_customs = cols.get("报关方式")

            for r in range(header_row + 1, (ws.max_row or header_row) + 1):
                waybill = _norm(ws.cell(r, col_waybill).value)
                fee = _norm(ws.cell(r, col_fee).value)

                # Skip blank / footer / bank info
                if not waybill and not fee:
                    # check if it's a total footer (金额 in amount cols without waybill)
                    continue
                if not waybill:
                    # footer like blank 运单号 with 合计 amounts
                    continue
                if waybill.startswith("收款") or "账户" in waybill or "公司名称" in waybill:
                    continue
                if "合计" in waybill or "合计" in fee:
                    continue

                if fee != "运费":
                    if fee:
                        skipped_fee_types[fee] = skipped_fee_types.get(fee, 0) + 1
                    else:
                        skipped_other += 1
                    continue

                weight = _to_float(ws.cell(r, col_weight).value) if col_weight else None
                amount = _to_float(ws.cell(r, col_amount).value) if col_amount else None
                amount_cny = _to_float(ws.cell(r, col_cny).value) if col_cny else None
                if amount_cny is None:
                    amount_cny = amount if amount is not None else 0.0
                if amount is None:
                    amount = amount_cny

                customer_no = _norm(ws.cell(r, col_customer).value) if col_customer else ""
                ext_nos = _norm(ws.cell(r, col_ext).value) if col_ext else ""

                row = FreightRow(
                    row_index=r,
                    waybill_no=waybill,
                    fee_type=fee,
                    date=_norm(ws.cell(r, col_date).value) if col_date else "",
                    customer_no=customer_no,
                    tracking_no=_norm(ws.cell(r, col_tracking).value) if col_tracking else "",
                    ext_nos=ext_nos,
                    service_type=_norm(ws.cell(r, col_service).value) if col_service else "",
                    warehouse=_norm(ws.cell(r, col_wh).value) if col_wh else "",
                    zipcode=_norm(ws.cell(r, col_zip).value) if col_zip else "",
                    chargeable_weight=weight or 0.0,
                    unit_price=_to_float(ws.cell(r, col_price).value) if col_price else None,
                    unit=_norm(ws.cell(r, col_unit).value) if col_unit else "",
                    quantity=_to_float(ws.cell(r, col_qty).value) if col_qty else None,
                    pieces=_to_float(ws.cell(r, col_pieces).value) if col_pieces else None,
                    country=_norm(ws.cell(r, col_country).value) if col_country else "",
                    currency=_norm(ws.cell(r, col_currency).value) if col_currency else "",
                    amount=amount or 0.0,
                    exchange_rate=_to_float(ws.cell(r, col_rate).value) if col_rate else None,
                    amount_cny=amount_cny or 0.0,
                    status=_norm(ws.cell(r, col_status).value) if col_status else "",
                    description=_norm(ws.cell(r, col_desc).value) if col_desc else "",
                    customs_mode=_norm(ws.cell(r, col_customs).value) if col_customs else "",
                    fba_ids=_parse_fba_ids(customer_no, ext_nos),
                    shop=shop,
                )
                all_rows.append(row)

            # Prefer first detail sheet that yields rows
            if all_rows:
                break

        # Ticket-level aggregation
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
            t["weight_total"] += row.chargeable_weight
            t["line_count"] += 1
            for fid in row.fba_ids:
                if fid not in t["fba_ids"]:
                    t["fba_ids"].append(fid)
            if row.customer_no and row.customer_no not in t["customer_nos"]:
                t["customer_nos"].append(row.customer_no)

        total_freight = sum(r.amount_cny for r in all_rows)
        total_weight = sum(r.chargeable_weight for r in all_rows)

        return {
            "ok": True,
            "sheet": used_sheet,
            "header_row": header_row,
            "shop": shop,
            "freight_row_count": len(all_rows),
            "ticket_count": len(tickets),
            "total_freight": round(total_freight, 4),
            "total_weight": round(total_weight, 4),
            "skipped_fee_types": skipped_fee_types,
            "rows": [r.to_dict() for r in all_rows],
            "tickets": list(tickets.values()),
        }
    finally:
        wb.close()
