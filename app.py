#!/usr/bin/env python3
"""Amazon FBA 头程运费分摊 — Flask app (实际毛重三步分摊)."""
from __future__ import annotations

import json
import os
import uuid
from datetime import datetime
from pathlib import Path

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
    jsonify,
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from lib.bill_parser import parse_bill
from lib.saihu import peek_columns, parse_saihu, parse_packing, parse_unit_weight_table
from lib.allocate import allocate_by_weight, freight_by_shipment_export_rows

BASE = Path(__file__).resolve().parent
UPLOADS = BASE / "uploads"
EXPORTS = BASE / "exports"
UPLOADS.mkdir(exist_ok=True)
EXPORTS.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = (
    os.environ.get("SECRET_KEY") or os.environ.get("FLASK_SECRET") or "first-leg-cost-dev-key"
)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64MB


def _sid() -> str:
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex
    return session["sid"]


def _state_path() -> Path:
    return UPLOADS / f"{_sid()}_state.json"


def _load_state() -> dict:
    p = _state_path()
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}


def _save_state(state: dict) -> None:
    _state_path().write_text(json.dumps(state, ensure_ascii=False, default=str), encoding="utf-8")


def _thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def _write_sheet(ws, rows: list[dict], headers: list[str] | None = None):
    if not rows:
        ws.append(["（无数据）"])
        return
    headers = headers or list(rows[0].keys())
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = _thin_border()
    for r_i, row in enumerate(rows, 2):
        for c_i, h in enumerate(headers, 1):
            val = row.get(h)
            if isinstance(val, (list, set)):
                val = ",".join(str(x) for x in val)
            cell = ws.cell(r_i, c_i, val)
            cell.border = _thin_border()
    for c in range(1, len(headers) + 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = 16


def _run_allocate(state: dict) -> dict | None:
    bill = state.get("bill")
    saihu = state.get("saihu")
    if not bill or not bill.get("rows") or not saihu or not saihu.get("rows"):
        return None
    unit_weights = state.get("unit_weights") or {}
    rows = saihu["rows"]
    if unit_weights:
        # inject missing unit weights
        patched = []
        for r in rows:
            rr = dict(r)
            if float(rr.get("unit_weight") or 0) <= 0 and rr.get("sku") in unit_weights:
                rr["unit_weight"] = unit_weights[rr["sku"]]
            patched.append(rr)
        rows = patched
    return allocate_by_weight(
        bill["rows"],
        rows,
        shop="",
        shipment_weights=saihu.get("shipment_weights"),
    )


@app.route("/")
def index():
    state = _load_state()
    return render_template(
        "index.html",
        bill=state.get("bill"),
        saihu=state.get("saihu"),
        alloc=state.get("alloc"),
        saihu_peek=state.get("saihu_peek"),
        saihu_mapping=state.get("saihu_mapping"),
        unit_weight_info=state.get("unit_weight_info"),
        alloc_error=state.get("alloc_error"),
    )


@app.route("/upload_bill", methods=["POST"])
def upload_bill():
    f = request.files.get("bill_file")
    if not f or not f.filename:
        flash("请选择物流账单 Excel 文件", "error")
        return redirect(url_for("index"))
    sid = _sid()
    dest = UPLOADS / f"{sid}_bill.xlsx"
    f.save(dest)
    try:
        result = parse_bill(str(dest), shop="")
    except Exception as e:
        flash(f"解析账单失败: {e}", "error")
        return redirect(url_for("index"))
    if not result.get("rows"):
        flash("未解析到运费行，请确认文件含明细表且费用类型含「运费」", "error")
    else:
        flash(
            f"账单解析成功：运费合计 ¥{result['total_freight']:.2f}，"
            f"{result['freight_row_count']} 行 / {result['ticket_count']} 票"
            "（货件号取自 G列扩展单号）",
            "ok",
        )
    state = _load_state()
    state["bill"] = result
    state["bill_path"] = str(dest)
    state.pop("alloc", None)
    state.pop("alloc_error", None)
    _save_state(state)
    return redirect(url_for("index"))


@app.route("/upload_saihu", methods=["POST"])
def upload_saihu():
    f = request.files.get("saihu_file")
    if not f or not f.filename:
        flash("请选择赛狐导出 Excel（含「装箱信息」表）", "error")
        return redirect(url_for("index"))
    sid = _sid()
    dest = UPLOADS / f"{sid}_saihu.xlsx"
    f.save(dest)
    state = _load_state()
    state["saihu_path"] = str(dest)
    state.pop("saihu", None)
    state.pop("alloc", None)
    state.pop("alloc_error", None)

    # Prefer auto-parse 装箱信息
    try:
        auto = parse_packing(str(dest), unit_weights=state.get("unit_weights"))
    except Exception as e:
        auto = {"ok": False, "error": str(e)}

    if auto.get("ok") and auto.get("rows"):
        state["saihu"] = auto
        state.pop("saihu_peek", None)
        _save_state(state)
        shops = "、".join(auto.get("shops") or []) or "（未识别店铺，输出仍保留店铺列）"
        flash(
            f"装箱信息解析成功：{auto['row_count']} 行，"
            f"{auto['fba_count']} 个货件，{auto['sku_count']} 个 SKU；店铺：{shops}",
            "ok",
        )
        return redirect(url_for("index") + "#allocate")

    # Fallback: column mapper
    try:
        peek = peek_columns(str(dest))
    except Exception as e:
        flash(f"读取赛狐文件失败: {e}", "error")
        return redirect(url_for("index"))
    state["saihu_peek"] = peek
    if peek.get("sheets"):
        state["saihu_mapping"] = peek["sheets"][0].get("suggested_mapping") or {}
        state["saihu_sheet"] = peek["sheets"][0]["name"]
        state["saihu_header_row"] = peek["sheets"][0]["header_row"]
    _save_state(state)
    msg = "未能自动识别装箱信息，请确认列映射后解析。"
    if auto.get("error"):
        msg = str(auto["error"]) + "。" + msg
    flash(msg, "error" if auto.get("error") else "ok")
    return redirect(url_for("index") + "#saihu")


@app.route("/parse_saihu", methods=["POST"])
def parse_saihu_route():
    state = _load_state()
    path = state.get("saihu_path")
    if not path or not Path(path).exists():
        flash("请先上传赛狐文件", "error")
        return redirect(url_for("index"))
    sheet = request.form.get("sheet") or state.get("saihu_sheet")
    header_row = int(request.form.get("header_row") or state.get("saihu_header_row") or 1)
    weight_mode = request.form.get("weight_mode") or "unit"
    mapping = {
        "fba_id": request.form.get("map_fba_id") or "",
        "sku": request.form.get("map_sku") or "",
        "qty": request.form.get("map_qty") or "",
        "unit_weight": request.form.get("map_unit_weight") or "",
        "total_weight": request.form.get("map_total_weight") or "",
        "shop": request.form.get("map_shop") or "",
    }
    result = parse_saihu(
        path,
        mapping=mapping,
        sheet_name=sheet,
        header_row=header_row,
        weight_mode=weight_mode,
        unit_weights=state.get("unit_weights"),
    )
    if not result.get("ok"):
        flash(result.get("error") or "解析失败", "error")
        state["saihu_mapping"] = mapping
        _save_state(state)
        return redirect(url_for("index") + "#saihu")
    state["saihu"] = result
    state["saihu_mapping"] = mapping
    state["saihu_sheet"] = sheet
    state["saihu_header_row"] = header_row
    state.pop("alloc", None)
    state.pop("alloc_error", None)
    _save_state(state)
    flash(
        f"赛狐解析成功：{result['row_count']} 行，"
        f"{result['fba_count']} 个货件，{result['sku_count']} 个 SKU",
        "ok",
    )
    return redirect(url_for("index") + "#allocate")


@app.route("/upload_unit_weights", methods=["POST"])
def upload_unit_weights():
    f = request.files.get("unit_weight_file")
    if not f or not f.filename:
        flash("请选择单件含包装重量表", "error")
        return redirect(url_for("index") + "#saihu")
    sid = _sid()
    dest = UPLOADS / f"{sid}_unit_weights.xlsx"
    f.save(dest)
    try:
        result = parse_unit_weight_table(str(dest))
    except Exception as e:
        flash(f"解析单件重量表失败: {e}", "error")
        return redirect(url_for("index") + "#saihu")
    if not result.get("ok"):
        flash(result.get("error") or "单件重量表解析失败", "error")
        return redirect(url_for("index") + "#saihu")
    state = _load_state()
    state["unit_weights"] = result["weights"]
    state["unit_weight_info"] = {"count": result["count"], "path": str(dest)}
    state.pop("alloc", None)
    state.pop("alloc_error", None)
    # Re-apply onto existing saihu rows if present
    if state.get("saihu_path") and Path(state["saihu_path"]).exists():
        try:
            auto = parse_packing(state["saihu_path"], unit_weights=result["weights"])
            if auto.get("ok"):
                state["saihu"] = auto
        except Exception:
            pass
    _save_state(state)
    flash(f"单件重量表已加载：{result['count']} 个 SKU（多SKU货件分摊需要）", "ok")
    return redirect(url_for("index") + "#allocate")


@app.route("/allocate", methods=["POST"])
def allocate_route():
    state = _load_state()
    bill = state.get("bill")
    saihu = state.get("saihu")
    if not bill or not bill.get("rows"):
        flash("请先上传并解析物流账单", "error")
        return redirect(url_for("index"))
    if not saihu or not saihu.get("rows"):
        flash("请先上传并解析赛狐装箱信息", "error")
        return redirect(url_for("index") + "#saihu")
    result = _run_allocate(state)
    if result is None:
        flash("分摊失败：数据不完整", "error")
        return redirect(url_for("index"))
    if not result.get("ok"):
        state["alloc_error"] = result.get("error")
        state.pop("alloc", None)
        _save_state(state)
        flash(result.get("error") or "分摊失败", "error")
        return redirect(url_for("index") + "#allocate")
    state["alloc"] = result
    state.pop("alloc_error", None)
    _save_state(state)
    bal = "平衡" if result.get("allocation_balance_ok") else "不平衡(请检查)"
    flash(
        f"分摊完成：{result['sku_count']} 个 SKU，"
        f"已分摊运费 ¥{result['total_allocated']:.2f} "
        f"（计入账单运费 ¥{result.get('included_bill_freight', 0):.2f}，{bal}），"
        f"匹配货件 {result['matched_fba_count']} 个",
        "ok",
    )
    return redirect(url_for("index") + "#result")


@app.route("/export/freight")
def export_freight():
    state = _load_state()
    bill = state.get("bill")
    if not bill or not bill.get("rows"):
        flash("无账单数据可导出", "error")
        return redirect(url_for("index"))
    rows = freight_by_shipment_export_rows(bill["rows"], shop="")
    wb = Workbook()
    ws = wb.active
    ws.title = "运费按货件"
    _write_sheet(ws, rows)
    ws2 = wb.create_sheet("按运单汇总")
    ticket_rows = []
    for t in bill.get("tickets") or []:
        ticket_rows.append(
            {
                "运单号": t["waybill_no"],
                "运费合计": round(t["freight_total"], 4),
                "收费重合计": round(t["weight_total"], 4),
                "运费行数": t["line_count"],
                "FBA货件号(扩展单号)": ",".join(t.get("fba_ids") or []),
                "客户单号": ",".join(t.get("customer_nos") or []),
            }
        )
    _write_sheet(ws2, ticket_rows)
    ws3 = wb.create_sheet("汇总")
    _write_sheet(
        ws3,
        [
            {
                "明细表": bill.get("sheet"),
                "运费行数": bill.get("freight_row_count"),
                "运单数": bill.get("ticket_count"),
                "运费合计_CNY": bill.get("total_freight"),
                "收费重合计": bill.get("total_weight"),
                "排除费用类型": json.dumps(bill.get("skipped_fee_types") or {}, ensure_ascii=False),
                "导出时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ],
    )
    out = EXPORTS / f"freight_{_sid()[:8]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(out)
    return send_file(out, as_attachment=True, download_name=out.name)


@app.route("/export/sku")
def export_sku():
    state = _load_state()
    alloc = state.get("alloc")
    if not alloc or not alloc.get("sku_results"):
        flash("请先完成分摊", "error")
        return redirect(url_for("index"))
    wb = Workbook()
    ws = wb.active
    ws.title = "SKU单件头程"
    sku_rows = []
    for r in alloc["sku_results"]:
        sku_rows.append(
            {
                "店铺": r.get("shop") or "",
                "SKU": r["sku"],
                "总发货数量": r["total_qty"],
                "分摊运费合计": r["total_freight"],
                "单件头程": r["unit_first_leg"],
                "涉及货件数": r["shipment_count"],
                "FBA货件号": ",".join(r.get("fba_ids") or []),
            }
        )
    _write_sheet(ws, sku_rows)

    ws2 = wb.create_sheet("货件SKU明细")
    detail_rows = []
    for r in alloc.get("shipment_allocs") or []:
        detail_rows.append(
            {
                "店铺": r.get("shop") or "",
                "运单号": r["waybill_no"],
                "FBA货件号": r["fba_id"],
                "SKU": r["sku"],
                "数量": r["qty"],
                "货件实际毛重": r["alloc_weight"],
                "货件运费": r["shipment_freight"],
                "SKU分摊费用": r["sku_freight"],
                "本行单件头程": r["unit_first_leg_partial"],
            }
        )
    _write_sheet(ws2, detail_rows)

    if alloc.get("unmatched_freight"):
        ws3 = wb.create_sheet("未匹配运费")
        _write_sheet(
            ws3,
            [
                {
                    "运单号": u["waybill_no"],
                    "运费": u["freight"],
                    "FBA货件号": ",".join(u.get("fba_ids") or []),
                    "原因": u.get("reason"),
                }
                for u in alloc["unmatched_freight"]
            ],
        )

    out = EXPORTS / f"sku_firstleg_{_sid()[:8]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(out)
    return send_file(out, as_attachment=True, download_name=out.name)


@app.route("/api/bill_summary")
def api_bill_summary():
    state = _load_state()
    bill = state.get("bill")
    if not bill:
        return jsonify({"ok": False, "error": "no bill"})
    return jsonify(
        {
            "ok": True,
            "total_freight": bill.get("total_freight"),
            "ticket_count": bill.get("ticket_count"),
            "freight_row_count": bill.get("freight_row_count"),
            "sheet": bill.get("sheet"),
        }
    )


@app.route("/reset", methods=["POST"])
def reset():
    sid = _sid()
    for p in UPLOADS.glob(f"{sid}*"):
        try:
            p.unlink()
        except OSError:
            pass
    session.clear()
    flash("已清空会话数据", "ok")
    return redirect(url_for("index"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8765"))
    print(f"头程运费分摊工具: http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG", "0") == "1")
