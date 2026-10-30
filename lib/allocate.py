"""Three-step FBA first-leg freight allocation by actual gross weight."""
from __future__ import annotations

from collections import defaultdict
from typing import Any


def allocate_by_weight(
    freight_rows: list[dict[str, Any]],
    shipment_skus: list[dict[str, Any]],
    shop: str = "",
    shipment_weights: dict[str, float] | None = None,
) -> dict[str, Any]:
    """
    Strict formula:
      ① 货件运费＝整票实际运费 × 该货件实际毛重 ÷ 本票所有货件实际毛重之和
      ② SKU总重量＝发货数量×单件含包装重量；
         SKU分摊＝货件运费×SKU总重÷货件内SKU总重之和；
         单SKU则货件运费全归该SKU
      ③ 单件头程＝SKU分摊÷数量；同SKU多货件：合计分摊÷合计数量

    Multi-SKU shipments require unit packed weight; otherwise return Chinese error
    listing those shipments. Single-SKU works without unit weight.

    Output is keyed by (shop, sku). Shop comes from packing/shipment data when
    present; otherwise the optional shop argument / blank.
    """
    shipment_weights = dict(shipment_weights or {})

    # Index SKU lines by FBA id; also capture shop + unit weight
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
        line_shop = str(s.get("shop") or shop or "").strip()
        gw = float(s.get("shipment_gross_weight") or 0)
        if gw > 0 and fba not in shipment_weights:
            shipment_weights[fba] = gw
        skus_by_fba[fba].append(
            {
                "fba_id": fba,
                "sku": sku,
                "qty": qty,
                "unit_weight": unit_w,
                "shop": line_shop,
            }
        )

    # Merge duplicate SKU lines within same FBA
    for fba, lines in list(skus_by_fba.items()):
        merged: dict[str, dict] = {}
        for ln in lines:
            m = merged.setdefault(
                ln["sku"],
                {
                    "fba_id": fba,
                    "sku": ln["sku"],
                    "qty": 0.0,
                    "unit_weight": 0.0,
                    "shop": ln["shop"],
                },
            )
            m["qty"] += ln["qty"]
            if ln["unit_weight"] > 0 and m["unit_weight"] <= 0:
                m["unit_weight"] = ln["unit_weight"]
            if ln["shop"] and not m["shop"]:
                m["shop"] = ln["shop"]
        skus_by_fba[fba] = list(merged.values())
        if fba not in shipment_weights or shipment_weights[fba] <= 0:
            # derive from unit weights if possible
            uw_sum = sum(x["unit_weight"] * x["qty"] for x in skus_by_fba[fba])
            if uw_sum > 0:
                shipment_weights[fba] = uw_sum

    # Pre-check multi-SKU missing unit weight
    multi_sku_missing: list[dict[str, Any]] = []
    for fba, lines in skus_by_fba.items():
        if len(lines) <= 1:
            continue
        missing = [x["sku"] for x in lines if float(x.get("unit_weight") or 0) <= 0]
        if missing:
            multi_sku_missing.append(
                {
                    "fba_id": fba,
                    "skus": [x["sku"] for x in lines],
                    "missing_unit_weight_skus": missing,
                }
            )

    if multi_sku_missing:
        detail = "；".join(
            f"{m['fba_id']}（SKU: {', '.join(m['skus'])}，缺单件重量: {', '.join(m['missing_unit_weight_skus'])}）"
            for m in multi_sku_missing
        )
        return {
            "ok": False,
            "error": (
                "以下多SKU货件缺少「单件含包装重量」，无法按重量分摊。"
                "请上传单件重量表，或在装箱/货件数据中提供单件重量："
                f"{detail}"
            ),
            "multi_sku_missing_unit_weight": multi_sku_missing,
            "sku_results": [],
            "shipment_allocs": [],
            "unmatched_freight": [],
            "unmatched_fba_on_bill": [],
            "matched_fba_count": 0,
            "sku_count": 0,
            "total_allocated": 0.0,
            "included_bill_freight": 0.0,
        }

    # Each freight row is one ticket (keyed by ticket_id / 扩展单号+金额+收费重).
    # 运单号 is not used for grouping or matching.
    by_ticket: dict[str, list[dict]] = defaultdict(list)
    for r in freight_rows:
        tid = r.get("ticket_id") or (
            f"{r.get('ext_nos') or ''}|{float(r.get('amount_cny') or 0):.4f}|"
            f"{float(r.get('chargeable_weight') or 0):.4f}#row:{r.get('row_index')}"
        )
        by_ticket[tid].append(r)

    shipment_allocs: list[dict[str, Any]] = []
    # key: (shop, sku)
    sku_costs: dict[tuple[str, str], dict[str, Any]] = {}
    unmatched_freight: list[dict] = []
    unmatched_fba: list[str] = []
    matched_fba: set[str] = set()
    included_bill_freight = 0.0

    for ticket_id, lines in by_ticket.items():
        ticket_freight = sum(float(x["amount_cny"]) for x in lines)
        ticket_fba_ids: list[str] = []
        ext_nos = ""
        for line in lines:
            if not ext_nos:
                ext_nos = str(line.get("ext_nos") or "")
            for fid in line.get("fba_ids") or []:
                if fid not in ticket_fba_ids:
                    ticket_fba_ids.append(fid)
        ticket_label = ext_nos or ",".join(ticket_fba_ids) or ticket_id

        present = [fid for fid in ticket_fba_ids if fid in skus_by_fba]
        if not present:
            unmatched_freight.append(
                {
                    "ticket_id": ticket_id,
                    "waybill_no": ticket_label,  # UI/export label (ext / FBA)
                    "ext_nos": ext_nos,
                    "freight": ticket_freight,
                    "fba_ids": ticket_fba_ids,
                    "reason": "无赛狐装箱货件匹配（扩展单号 ↔ 货件编号）",
                }
            )
            continue

        # ① 货件实际毛重
        ship_w: dict[str, float] = {}
        for fid in present:
            w = float(shipment_weights.get(fid) or 0)
            if w <= 0:
                # single-SKU fallback: cannot allocate without weight
                w = 0.0
            ship_w[fid] = w
        total_ship_w = sum(ship_w.values())
        if total_ship_w <= 0:
            unmatched_freight.append(
                {
                    "ticket_id": ticket_id,
                    "waybill_no": ticket_label,
                    "ext_nos": ext_nos,
                    "freight": ticket_freight,
                    "fba_ids": present,
                    "reason": "匹配货件实际毛重之和为0（请检查装箱信息总重量/箱子重量）",
                }
            )
            continue

        included_bill_freight += ticket_freight

        for fid in present:
            ship_freight = ticket_freight * (ship_w[fid] / total_ship_w)
            _allocate_skus_in_shipment(
                fid,
                ship_freight,
                skus_by_fba[fid],
                ticket_id,
                ship_w[fid],
                shipment_allocs,
                sku_costs,
                ticket_label=ticket_label,
                ext_nos=ext_nos,
            )
            matched_fba.add(fid)

        for fid in ticket_fba_ids:
            if fid not in present:
                unmatched_fba.append(fid)

    # ③ finalize unit cost per (shop, sku)
    sku_results = []
    for (shop_key, sku), agg in sorted(sku_costs.items(), key=lambda x: (x[0][0], x[0][1])):
        unit = agg["total_freight"] / agg["total_qty"] if agg["total_qty"] else 0
        sku_results.append(
            {
                "shop": shop_key,
                "sku": sku,
                "total_freight": round(agg["total_freight"], 6),
                "total_qty": agg["total_qty"],
                "unit_first_leg": round(unit, 6),
                "shipment_count": agg["shipment_count"],
                "fba_ids": sorted(agg["fba_ids"]),
            }
        )

    total_allocated = round(sum(x["total_freight"] for x in sku_results), 4)
    return {
        "ok": True,
        "shop": shop,
        "sku_results": sku_results,
        "shipment_allocs": shipment_allocs,
        "unmatched_freight": unmatched_freight,
        "unmatched_fba_on_bill": sorted(set(unmatched_fba)),
        "matched_fba_count": len(matched_fba),
        "sku_count": len(sku_results),
        "total_allocated": total_allocated,
        "included_bill_freight": round(included_bill_freight, 4),
        "allocation_balance_ok": abs(total_allocated - round(included_bill_freight, 4)) < 0.05,
    }


def _allocate_skus_in_shipment(
    fba_id: str,
    ship_freight: float,
    sku_lines: list[dict],
    ticket_id: str,
    ship_gross_weight: float,
    shipment_allocs: list,
    sku_costs: dict,
    ticket_label: str = "",
    ext_nos: str = "",
) -> None:
    """
    ② SKU分摊：单SKU → 全归该SKU；
       多SKU → 按 (数量×单件含包装重量) 比例分摊。
    """
    if len(sku_lines) == 1:
        shares = [(sku_lines[0], ship_freight, sku_lines[0]["qty"] * float(sku_lines[0].get("unit_weight") or 0))]
    else:
        weights = []
        for s in sku_lines:
            sw = float(s["qty"]) * float(s.get("unit_weight") or 0)
            weights.append(sw)
        total_w = sum(weights)
        shares = []
        for s, sw in zip(sku_lines, weights):
            share = ship_freight * (sw / total_w) if total_w > 0 else 0.0
            shares.append((s, share, sw))

    label = ticket_label or ext_nos or ticket_id
    for s, share, sku_total_w in shares:
        line_shop = s.get("shop") or ""
        shipment_allocs.append(
            {
                "shop": line_shop,
                "ticket_id": ticket_id,
                "waybill_no": label,  # display/export label (ext / FBA), not 运单号
                "ext_nos": ext_nos,
                "fba_id": fba_id,
                "sku": s["sku"],
                "qty": s["qty"],
                "unit_weight": s.get("unit_weight") or 0,
                "sku_total_weight": round(sku_total_w, 6) if sku_total_w else None,
                "alloc_weight": round(ship_gross_weight, 6),
                "shipment_freight": round(ship_freight, 6),
                "sku_freight": round(share, 6),
                "unit_first_leg_partial": round(share / s["qty"], 6) if s["qty"] else 0,
            }
        )
        key = (line_shop, s["sku"])
        agg = sku_costs.setdefault(
            key,
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
    """Flatten freight rows for CSV/XLSX download (no Saihu needed). G/K/R only."""
    out = []
    for r in freight_rows:
        out.append(
            {
                "店铺": shop or r.get("shop") or "",
                "票标识": r.get("ticket_id") or "",
                "扩展单号": r.get("ext_nos"),
                "FBA货件号列表": ",".join(r.get("fba_ids") or []),
                "收费重": r.get("chargeable_weight"),
                "金额": r.get("amount_cny") if r.get("amount_cny") is not None else r.get("amount"),
            }
        )
    return out
