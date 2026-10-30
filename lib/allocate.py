"""Three-step FBA first-leg freight allocation by chargeable weight."""
from __future__ import annotations

from collections import defaultdict
from typing import Any


def allocate_by_weight(
    freight_rows: list[dict[str, Any]],
    shipment_skus: list[dict[str, Any]],
    shop: str = "",
) -> dict[str, Any]:
    """
    shipment_skus items need:
      - fba_id: str
      - sku: str
      - qty: float          # 发货数量
      - unit_weight: float  # 单件重量 (kg); OR total_weight for that SKU line
      - weight_is_total: bool (optional, default False)

    Formula:
      1. 货件运费 = 整票总运费 × (该货件计费重 / 整票总计费重)
      2. SKU分摊费用 = 货件运费 × (SKU分摊重量 / 货件分摊总重量)
      3. 单件头程 = 该SKU所有货件分摊费用之和 / 该SKU总发货数量

    When Saihu provides per-shipment SKU weights, we can allocate a ticket
    across multiple FBA ids. Until then, each freight row is treated as one
    shipment-level cost (amount_cny already at line level).
    """
    # Index SKU lines by FBA id
    skus_by_fba: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for s in shipment_skus:
        sku = str(s.get("sku") or "").strip()
        fba = str(s.get("fba_id") or "").strip()
        if not fba or not sku:
            continue
        qty = float(s.get("qty") or 0)
        if qty <= 0:
            continue
        unit_w = float(s.get("unit_weight") or 0)
        if s.get("weight_is_total"):
            total_w = unit_w
            unit_w = total_w / qty if qty else 0
        else:
            total_w = unit_w * qty
        skus_by_fba[fba].append(
            {
                "fba_id": fba,
                "sku": sku,
                "qty": qty,
                "unit_weight": unit_w,
                "alloc_weight": total_w,
            }
        )

    # Group freight rows by waybill (票)
    by_ticket: dict[str, list[dict]] = defaultdict(list)
    for r in freight_rows:
        by_ticket[r["waybill_no"]].append(r)

    shipment_allocs: list[dict[str, Any]] = []
    sku_costs: dict[str, dict[str, Any]] = {}  # sku -> aggregate
    unmatched_freight: list[dict] = []
    unmatched_fba: list[str] = []
    matched_fba: set[str] = set()

    for waybill, lines in by_ticket.items():
        ticket_freight = sum(float(x["amount_cny"]) for x in lines)
        # Collect all FBA ids mentioned on this ticket
        ticket_fba_ids: list[str] = []
        for line in lines:
            for fid in line.get("fba_ids") or []:
                if fid not in ticket_fba_ids:
                    ticket_fba_ids.append(fid)

        # Which of those FBA ids have Saihu SKU data?
        present = [fid for fid in ticket_fba_ids if fid in skus_by_fba]

        if not present:
            unmatched_freight.append(
                {
                    "waybill_no": waybill,
                    "freight": ticket_freight,
                    "fba_ids": ticket_fba_ids,
                    "reason": "无赛狐货件匹配",
                }
            )
            continue

        # Shipment chargeable weights from Saihu (sum of SKU alloc weights)
        ship_weights: dict[str, float] = {}
        for fid in present:
            ship_weights[fid] = sum(x["alloc_weight"] for x in skus_by_fba[fid])
        total_ship_w = sum(ship_weights.values())

        # If bill lines already have per-line weights and each line maps to
        # distinct primary customer_no, prefer bill chargeable weights for step 1.
        # Fallback: allocate ticket freight by Saihu shipment total weights.
        use_bill_line_weights = False
        if len(lines) == len(present) and all(
            (line.get("customer_no") in present) for line in lines
        ):
            use_bill_line_weights = True

        if use_bill_line_weights:
            # Step 1 per bill line
            for line in lines:
                fid = line.get("customer_no")
                if fid not in skus_by_fba:
                    continue
                ship_freight = float(line["amount_cny"])
                _allocate_skus_in_shipment(
                    fid,
                    ship_freight,
                    skus_by_fba[fid],
                    waybill,
                    shop,
                    shipment_allocs,
                    sku_costs,
                )
                matched_fba.add(fid)
        else:
            # Step 1: allocate ticket freight by shipment weights
            if total_ship_w <= 0:
                unmatched_freight.append(
                    {
                        "waybill_no": waybill,
                        "freight": ticket_freight,
                        "fba_ids": present,
                        "reason": "赛狐货件分摊总重量为0",
                    }
                )
                continue
            for fid in present:
                ship_freight = ticket_freight * (ship_weights[fid] / total_ship_w)
                _allocate_skus_in_shipment(
                    fid,
                    ship_freight,
                    skus_by_fba[fid],
                    waybill,
                    shop,
                    shipment_allocs,
                    sku_costs,
                )
                matched_fba.add(fid)

        for fid in ticket_fba_ids:
            if fid not in present:
                unmatched_fba.append(fid)

    # Step 3 finalize unit cost
    sku_results = []
    for sku, agg in sorted(sku_costs.items()):
        unit = agg["total_freight"] / agg["total_qty"] if agg["total_qty"] else 0
        sku_results.append(
            {
                "sku": sku,
                "shop": shop,
                "total_freight": round(agg["total_freight"], 6),
                "total_qty": agg["total_qty"],
                "unit_first_leg": round(unit, 6),
                "shipment_count": agg["shipment_count"],
                "fba_ids": sorted(agg["fba_ids"]),
            }
        )

    return {
        "ok": True,
        "shop": shop,
        "sku_results": sku_results,
        "shipment_allocs": shipment_allocs,
        "unmatched_freight": unmatched_freight,
        "unmatched_fba_on_bill": sorted(set(unmatched_fba)),
        "matched_fba_count": len(matched_fba),
        "sku_count": len(sku_results),
        "total_allocated": round(sum(x["total_freight"] for x in sku_results), 4),
    }


def _allocate_skus_in_shipment(
    fba_id: str,
    ship_freight: float,
    sku_lines: list[dict],
    waybill: str,
    shop: str,
    shipment_allocs: list,
    sku_costs: dict,
) -> None:
    """Step 2: SKU分摊费用 = 货件运费 × (SKU分摊重量 / 货件分摊总重量)."""
    total_w = sum(x["alloc_weight"] for x in sku_lines)
    for s in sku_lines:
        if total_w > 0:
            share = ship_freight * (s["alloc_weight"] / total_w)
        else:
            # equal split fallback by qty
            total_qty = sum(x["qty"] for x in sku_lines) or 1
            share = ship_freight * (s["qty"] / total_qty)
        shipment_allocs.append(
            {
                "shop": shop,
                "waybill_no": waybill,
                "fba_id": fba_id,
                "sku": s["sku"],
                "qty": s["qty"],
                "alloc_weight": round(s["alloc_weight"], 6),
                "shipment_freight": round(ship_freight, 6),
                "sku_freight": round(share, 6),
                "unit_first_leg_partial": round(share / s["qty"], 6) if s["qty"] else 0,
            }
        )
        agg = sku_costs.setdefault(
            s["sku"],
            {
                "total_freight": 0.0,
                "total_qty": 0.0,
                "shipment_count": 0,
                "fba_ids": set(),
            },
        )
        agg["total_freight"] += share
        agg["total_qty"] += s["qty"]
        agg["shipment_count"] += 1
        agg["fba_ids"].add(fba_id)


def freight_by_shipment_export_rows(freight_rows: list[dict], shop: str = "") -> list[dict]:
    """Flatten freight rows for CSV/XLSX download (no Saihu needed)."""
    out = []
    for r in freight_rows:
        out.append(
            {
                "店铺": shop or r.get("shop") or "",
                "运单号": r.get("waybill_no"),
                "客户单号": r.get("customer_no"),
                "扩展单号": r.get("ext_nos"),
                "FBA货件号列表": ",".join(r.get("fba_ids") or []),
                "服务类型": r.get("service_type"),
                "仓库代码": r.get("warehouse"),
                "收费重": r.get("chargeable_weight"),
                "件数": r.get("pieces"),
                "发往国家": r.get("country"),
                "币种": r.get("currency"),
                "金额": r.get("amount"),
                "折合人民币": r.get("amount_cny"),
                "日期": r.get("date"),
                "运单状态": r.get("status"),
            }
        )
    return out
