#!/usr/bin/env python3
"""Verify sample bill freight total ≈ 36095.4"""
from pathlib import Path
from lib.bill_parser import parse_bill

SAMPLE = Path(__file__).parent / "sample_bill.xlsx"
EXPECTED = 36095.4

def main():
    r = parse_bill(str(SAMPLE), shop="验证店铺")
    total = r["total_freight"]
    ok = abs(total - EXPECTED) < 0.05
    print("sheet:", r["sheet"])
    print("header_row:", r["header_row"])
    print("freight_rows:", r["freight_row_count"])
    print("tickets:", r["ticket_count"])
    print("total_freight:", total)
    print("skipped_fee_types:", r["skipped_fee_types"])
    print("expected:", EXPECTED)
    print("PASS" if ok else "FAIL")
    # show first 3
    for row in r["rows"][:3]:
        print(" ", row["waybill_no"], row["customer_no"], row["amount_cny"], row["fba_ids"])
    raise SystemExit(0 if ok else 1)

if __name__ == "__main__":
    main()
