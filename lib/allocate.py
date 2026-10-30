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

    Multi-SKU shipments matched on the bill require unit packed weight (optional
    upload). Missing ones are skipped (listed in multi_sku_missing_unit_weight /
    unmatched_freight); other tickets still allocate. Single-SKU needs no unit weight.

    Output is keyed by sku only (no shop dimension). Same SKU across shops/tickets
    is aggregated together.
    """
    shipment_weights = dict(shipment_weights or {})

    # Index SKU lines by FBA id; capture unit weight
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
        gw = float(s.get("shipment_gross_weight") or 0)
        if gw > 0 and fba not in shipment_weights:
            shipment_weights[fba] = gw
        skus_by_fba[fba].append(
            {
                "fba_id": fba,
                "sku": sku,
                "qty": qty,
                "unit_weight": unit_w,
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
                },
            )
            m["qty"] += ln["qty"]
            if ln["unit_weight"] > 0 and m["unit_weight"] <= 0:
                m["unit_weight"] = ln["unit_weight"]
        skus_by_fba[fba] = list(merged.values())
        if fba not in shipment_weights or shipment_weights[fba] <= 0:
            # derive from unit weights if possible
            uw_sum = sum(x["unit_weight"] * x["qty"] for x in skus_by_fba[fba])
            if uw_sum > 0:
                shipment_weights[fba] = uw_sum

    # Each freight row is one ticket (keyed by ticket_id / 扩展单号+金额+收费重).
    # 运单号 is not used for grouping or matching.
    by_ticket: dict[str, list[dict]] = defaultdict(list)
    for r in freight_rows:
        tid = r.get("ticket_id") or (
            f"{r.get('ext_nos') or ''}|{float(r.get('amount_cny') or 0):.4f}|"
            f"{float(r.get('chargeable_weight') or 0):.4f}#row:{r.get('row_index')}"
        )
        by_ticket[tid].append(r)

    # Multi-SKU unit-weight check only for FBAs that appear on the bill tickets
    bill_fbas: set[str] = set()
    for lines in by_ticket.values():
        for line in lines:
            for fid in line.get("fba_ids") or []:
                if fid:
                    bill_fbas.add(str(fid).strip())

    multi_sku_missing: list[dict[str, Any]] = []
    for fba, lines in skus_by_fba.items():
        if fba not in bill_fbas:
            continue
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

    # Soft-block those FBAs: other tickets still allocate; UI can show missing list.
    blocked_multi_sku: set[str] = {m["fba_id"] for m in multi_sku_missing}

    shipment_allocs: list[dict[str, Any]] = []
    # key: sku
    sku_costs: dict[str, dict[str, Any]] = {}
    calc_steps: list[dict[str, Any]] = []
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

        present_all = [fid for fid in ticket_fba_ids if fid in skus_by_fba]
        blocked_here = [fid for fid in present_all if fid in blocked_multi_sku]
        present = [fid for fid in present_all if fid not in blocked_multi_sku]
        if not present_all:
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
        if blocked_here and not present:
            unmatched_freight.append(
                {
                    "ticket_id": ticket_id,
                    "waybill_no": ticket_label,
                    "ext_nos": ext_nos,
                    "freight": ticket_freight,
                    "fba_ids": ticket_fba_ids,
                    "reason": (
                        "多SKU货件缺少单件含包装重量，请上传单件重量表："
                        + ",".join(blocked_here)
                    ),
                }
            )
            continue
        # If ticket mixes OK + blocked FBAs, only allocate the OK ones' weight share
        # among present (blocked excluded from denominator).

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

        # Collect filled inputs used for this ticket (for UI核对)
        unit_inputs: list[dict[str, Any]] = []
        ship_inputs: list[dict[str, Any]] = []
        for fid in present:
            ship_inputs.append({"fba_id": fid, "gross_weight": round(ship_w[fid], 6)})
            sku_lines_here = skus_by_fba.get(fid) or []
            if len(sku_lines_here) > 1:
                for ln in sku_lines_here:
                    uw = float(ln.get("unit_weight") or 0)
                    unit_inputs.append(
                        {
                            "fba_id": fid,
                            "sku": ln["sku"],
                            "qty": ln["qty"],
                            "unit_weight": uw,
                        }
                    )
        chargeable_w = sum(float(x.get("chargeable_weight") or 0) for x in lines)
        step_ticket: dict[str, Any] = {
            "ticket_id": ticket_id,
            "ext_nos": ext_nos,
            "ticket_freight": round(ticket_freight, 6),
            "total_ship_weight": round(total_ship_w, 6),
            "formula": {
                "ship": "货件运费＝整票运费 × 该货件毛重 ÷ 本票货件毛重之和",
                "sku_multi": "SKU分摊＝货件运费 ×（数量×单件含包装重量）÷ 货件内SKU总重之和",
                "sku_single": "单SKU：货件运费全归该SKU",
                "unit": "单件头程＝SKU分摊 ÷ 数量",
            },
            "inputs": {
                "ticket_freight": round(ticket_freight, 6),
                "chargeable_weight": round(chargeable_w, 6),
                "shipments": ship_inputs,
                "unit_weights": unit_inputs,
            },
            "shipments": [],
            "sku_finalize": [],
        }

        for fid in present:
            ship_freight = ticket_freight * (ship_w[fid] / total_ship_w)
            ship_step = _allocate_skus_in_shipment(
                fid,
                ship_freight,
                skus_by_fba[fid],
                ticket_id,
                ship_w[fid],
                shipment_allocs,
                sku_costs,
                ticket_label=ticket_label,
                ext_nos=ext_nos,
                ticket_freight=ticket_freight,
                total_ship_weight=total_ship_w,
            )
            step_ticket["shipments"].append(ship_step)
            matched_fba.add(fid)

        for fid in ticket_fba_ids:
            if fid not in present:
                unmatched_fba.append(fid)

        calc_steps.append(step_ticket)

    # ③ finalize unit cost per sku
    sku_results = []
    for sku, agg in sorted(sku_costs.items(), key=lambda x: x[0]):
        unit = agg["total_freight"] / agg["total_qty"] if agg["total_qty"] else 0
        sku_results.append(
            {
                "sku": sku,
                "total_freight": round(agg["total_freight"], 6),
                "total_qty": agg["total_qty"],
                "unit_first_leg": round(unit, 6),
                "shipment_count": agg["shipment_count"],
                "fba_ids": sorted(agg["fba_ids"]),
            }
        )

    # Attach ③ finalize lines onto each ticket's calc_steps (SKUs touched by that ticket)
    for step in calc_steps:
        touched = {
            a["sku"]
            for a in shipment_allocs
            if a.get("ticket_id") == step["ticket_id"]
        }
        step["sku_finalize"] = [
            {
                "sku": r["sku"],
                "total_freight": r["total_freight"],
                "total_qty": r["total_qty"],
                "unit_first_leg": r["unit_first_leg"],
                "expr": (
                    f"{r['sku']}：单件头程＝{r['total_freight']:.4f} ÷ {r['total_qty']} "
                    f"＝ {r['unit_first_leg']:.4f}"
                ),
            }
            for r in sku_results
            if r["sku"] in touched
        ]

    total_allocated = round(sum(x["total_freight"] for x in sku_results), 4)
    return {
        "ok": True,
        "sku_results": sku_results,
        "shipment_allocs": shipment_allocs,
        "calc_steps": calc_steps,
        "unmatched_freight": unmatched_freight,
        "unmatched_fba_on_bill": sorted(set(unmatched_fba)),
        "matched_fba_count": len(matched_fba),
        "sku_count": len(sku_results),
        "total_allocated": total_allocated,
        "included_bill_freight": round(included_bill_freight, 4),
        "allocation_balance_ok": abs(total_allocated - round(included_bill_freight, 4)) < 0.05,
        "multi_sku_missing_unit_weight": multi_sku_missing,
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
    ticket_freight: float = 0.0,
    total_ship_weight: float = 0.0,
) -> dict[str, Any]:
    """
    ② SKU分摊：单SKU → 全归该SKU；
       多SKU → 按 (数量×单件含包装重量) 比例分摊。
    Returns a step dict for UI explanation.
    """
    single = len(sku_lines) == 1
    if single:
        shares = [(sku_lines[0], ship_freight, sku_lines[0]["qty"] * float(sku_lines[0].get("unit_weight") or 0))]
        sku_weight_total = shares[0][2]
    else:
        weights = []
        for s in sku_lines:
            sw = float(s["qty"]) * float(s.get("unit_weight") or 0)
            weights.append(sw)
        sku_weight_total = sum(weights)
        shares = []
        for s, sw in zip(sku_lines, weights):
            share = ship_freight * (sw / sku_weight_total) if sku_weight_total > 0 else 0.0
            shares.append((s, share, sw))

    label = ticket_label or ext_nos or ticket_id
    sku_steps = []
    for s, share, sku_total_w in shares:
        unit_partial = round(share / s["qty"], 6) if s["qty"] else 0
        shipment_allocs.append(
            {
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
                "unit_first_leg_partial": unit_partial,
            }
        )
        key = s["sku"]
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

        uw = float(s.get("unit_weight") or 0)
        if single:
            expr = (
                f"单SKU → 货件运费全归 {s['sku']}："
                f"{ship_freight:.4f}；单件＝{ship_freight:.4f} ÷ {s['qty']} ＝ {unit_partial:.4f}"
            )
        else:
            expr = (
                f"{s['sku']}：{ship_freight:.4f} × ({s['qty']}×{uw}＝{sku_total_w:.4f})"
                f" ÷ {sku_weight_total:.4f} ＝ {share:.4f}；"
                f"单件＝{share:.4f} ÷ {s['qty']} ＝ {unit_partial:.4f}"
            )
        sku_steps.append(
            {
                "sku": s["sku"],
                "qty": s["qty"],
                "unit_weight": uw,
                "sku_total_weight": round(sku_total_w, 6) if sku_total_w else 0,
                "sku_freight": round(share, 6),
                "unit_first_leg_partial": unit_partial,
                "expr": expr,
            }
        )

    ship_expr = (
        f"货件运费＝{ticket_freight:.4f} × {ship_gross_weight:.4f} ÷ {total_ship_weight:.4f}"
        f" ＝ {ship_freight:.4f}"
    )
    return {
        "fba_id": fba_id,
        "gross_weight": round(ship_gross_weight, 6),
        "ship_freight": round(ship_freight, 6),
        "single_sku": single,
        "sku_weight_total": round(sku_weight_total, 6) if sku_weight_total else 0,
        "ship_expr": ship_expr,
        "skus": sku_steps,
    }


def freight_by_shipment_export_rows(freight_rows: list[dict], shop: str = "") -> list[dict]:
    """Flatten freight rows for CSV/XLSX download (no Saihu needed). G/K/R only."""
    out = []
    for r in freight_rows:
        out.append(
            {
                "票标识": r.get("ticket_id") or "",
                "扩展单号": r.get("ext_nos"),
                "FBA货件号列表": ",".join(r.get("fba_ids") or []),
                "收费重": r.get("chargeable_weight"),
                "金额": r.get("amount_cny") if r.get("amount_cny") is not None else r.get("amount"),
            }
        )
    return out


def allocate_one_ticket(
    freight_rows: list[dict[str, Any]],
    shipment_skus: list[dict[str, Any]],
    ticket_id: str,
    shop: str = "",
    shipment_weights: dict[str, float] | None = None,
) -> dict[str, Any]:
    """Allocate a single bill ticket (one 票). Filters freight_rows to ticket_id."""
    tid = str(ticket_id or "").strip()
    if not tid:
        return {"ok": False, "error": "未指定票标识"}
    subset = [r for r in freight_rows if str(r.get("ticket_id") or "") == tid]
    if not subset:
        # allow match by ext_nos when UI passes extension string
        subset = [
            r
            for r in freight_rows
            if str(r.get("ext_nos") or "") == tid or str(r.get("ticket_id") or "") == tid
        ]
    if not subset:
        return {"ok": False, "error": f"账单中找不到票：{tid}"}
    result = allocate_by_weight(
        subset,
        shipment_skus,
        shop=shop,
        shipment_weights=shipment_weights,
    )
    result["ticket_id"] = subset[0].get("ticket_id") or tid
    result["ext_nos"] = subset[0].get("ext_nos") or ""
    result["ticket_freight"] = sum(float(x.get("amount_cny") or 0) for x in subset)
    result["ticket_weight"] = sum(float(x.get("chargeable_weight") or 0) for x in subset)
    result["fba_ids"] = []
    for x in subset:
        for fid in x.get("fba_ids") or []:
            if fid not in result["fba_ids"]:
                result["fba_ids"].append(fid)
    return result
