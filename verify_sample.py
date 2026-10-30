#!/usr/bin/env python3
"""Regression: row with FBAs FBA15MB02L89/… @ ¥562.50 → LR-BB223≈33.24 / LR-BB232≈20.86."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from lib.bill_parser import parse_bill
from lib.saihu import parse_packing
from lib.allocate import allocate_by_weight


def main() -> int:
    bill_path = ROOT / "sample_bill_sept.xlsx"
    if not bill_path.exists():
        bill_path = ROOT / "sample_bill.xlsx"
    saihu_path = ROOT / "sample_saihu_sept.xlsx"
    assert saihu_path.exists(), f"missing {saihu_path}"

    bill = parse_bill(str(bill_path))
    assert bill.get("ok") and bill.get("rows"), bill

    target_fbas = {"FBA15MB02L89", "FBA15MB030DD", "FBA15MB0BFP3"}
    trows = [
        r
        for r in bill["rows"]
        if abs(float(r["amount_cny"]) - 562.50) < 0.01
        and target_fbas.issubset(set(r.get("fba_ids") or []))
    ]
    if not trows:
        # fallback: match by extension containing all three FBAs
        trows = [
            r
            for r in bill["rows"]
            if target_fbas.issubset(set(r.get("fba_ids") or []))
        ]
    if not trows:
        print(f"FAIL: target FBA ticket not in bill {bill_path.name}")
        return 1

    freight = sum(r["amount_cny"] for r in trows)
    fbas = []
    for r in trows:
        for f in r["fba_ids"]:
            if f not in fbas:
                fbas.append(f)
    print(
        f"ticket ext={trows[0].get('ext_nos')!r} id={trows[0].get('ticket_id')!r}: "
        f"freight={freight}, fbas={fbas}"
    )
    assert abs(freight - 562.50) < 0.01, freight
    assert set(fbas) == target_fbas

    packing = parse_packing(str(saihu_path))
    assert packing.get("ok"), packing
    print(
        f"packing: rows={packing['row_count']} fbas={packing['fba_count']} "
        f"shops={packing.get('shops')}"
    )

    for fid in fbas:
        w = packing["shipment_weights"].get(fid)
        print(f"  {fid} gross={w}")
        lines = [r for r in packing["rows"] if r["fba_id"] == fid]
        for ln in lines:
            print(f"    {ln['sku']} qty={ln['qty']} shop={ln.get('shop')}")

    result = allocate_by_weight(
        trows,
        packing["rows"],
        shipment_weights=packing.get("shipment_weights"),
    )
    assert result.get("ok"), result
    assert result.get("allocation_balance_ok"), result

    by_sku = {r["sku"]: r for r in result["sku_results"]}
    print("SKU results:")
    for r in result["sku_results"]:
        print(
            f"  shop={r['shop']!r} sku={r['sku']} qty={r['total_qty']} "
            f"freight={r['total_freight']:.4f} unit={r['unit_first_leg']:.4f}"
        )

    u223 = by_sku["LR-BB223"]["unit_first_leg"]
    u232 = by_sku["LR-BB232"]["unit_first_leg"]
    ok223 = abs(u223 - 33.24) <= 0.02
    ok232 = abs(u232 - 20.86) <= 0.02
    print(f"LR-BB223 unit={u223:.4f} expect≈33.24 → {'PASS' if ok223 else 'FAIL'}")
    print(f"LR-BB232 unit={u232:.4f} expect≈20.86 → {'PASS' if ok232 else 'FAIL'}")
    print(
        f"allocated={result['total_allocated']} included={result['included_bill_freight']} "
        f"balance={result['allocation_balance_ok']}"
    )

    full = allocate_by_weight(
        bill["rows"],
        packing["rows"],
        shipment_weights=packing.get("shipment_weights"),
    )
    assert full.get("ok"), full
    print(
        f"full sept: matched_fba={full['matched_fba_count']} sku={full['sku_count']} "
        f"allocated={full['total_allocated']} included={full['included_bill_freight']} "
        f"balance={full['allocation_balance_ok']} unmatched_tickets={len(full['unmatched_freight'])}"
    )
    assert full["allocation_balance_ok"]

    shops = {r["shop"] for r in full["sku_results"]}
    print(f"shops in output: {shops}")
    assert any(s for s in shops), "expected shop dimension from 货件信息"

    # each data row is one ticket
    assert bill["ticket_count"] == bill["freight_row_count"], (
        bill["ticket_count"],
        bill["freight_row_count"],
    )
    print(f"tickets==rows: {bill['ticket_count']}")

    if not (ok223 and ok232):
        return 1
    print("ALL PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
