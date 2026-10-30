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
from lib.allocate import allocate_by_weight, allocate_one_ticket, freight_by_shipment_export_rows

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


def _saihu_rows_with_weights(state: dict, extra_weights: dict[str, float] | None = None) -> list[dict]:
    saihu = state.get("saihu") or {}
    rows = saihu.get("rows") or []
    weights = dict(state.get("unit_weights") or {})
    if extra_weights:
        weights.update({k: float(v) for k, v in extra_weights.items() if float(v) > 0})
    if not weights:
        return [dict(r) for r in rows]
    patched = []
    for r in rows:
        rr = dict(r)
        if float(rr.get("unit_weight") or 0) <= 0 and rr.get("sku") in weights:
            rr["unit_weight"] = float(weights[rr["sku"]])
        patched.append(rr)
    return patched


def _ticket_packing_preview(state: dict, ticket_id: str) -> dict:
    """Ticket + matched FBAs + packing lines together for verification / weight form."""
    bill = state.get("bill") or {}
    saihu = state.get("saihu") or {}
    rows = [r for r in (bill.get("rows") or []) if str(r.get("ticket_id") or "") == ticket_id]
    if not rows:
        return {"ok": False, "error": "票不存在"}
    fba_ids: list[str] = []
    for r in rows:
        for fid in r.get("fba_ids") or []:
            if fid and fid not in fba_ids:
                fba_ids.append(fid)
    unit_weights = state.get("unit_weights") or {}
    # Prefer detail_rows (full packing cols); fall back to aggregated rows
    detail = saihu.get("detail_rows") or []
    by_fba_detail: dict[str, list[dict]] = {}
    for d in detail:
        fid = d.get("fba_id")
        if fid in fba_ids:
            by_fba_detail.setdefault(fid, []).append(d)
    by_fba: dict[str, list[dict]] = {}
    for s in saihu.get("rows") or []:
        fid = s.get("fba_id")
        if fid in fba_ids:
            rr = dict(s)
            if float(rr.get("unit_weight") or 0) <= 0 and rr.get("sku") in unit_weights:
                rr["unit_weight"] = float(unit_weights[rr["sku"]])
            by_fba.setdefault(fid, []).append(rr)
    shipments = []
    needs_unit_weight = False
    skus_needing_weight: list[str] = []
    for fid in fba_ids:
        lines = by_fba.get(fid) or []
        packing_lines = by_fba_detail.get(fid) or lines
        multi = len(lines) > 1
        missing = [x["sku"] for x in lines if float(x.get("unit_weight") or 0) <= 0]
        if multi and missing:
            needs_unit_weight = True
            for sku in missing:
                if sku not in skus_needing_weight:
                    skus_needing_weight.append(sku)
        # Enrich agg lines with packing fields when detail available
        display_lines = []
        if by_fba_detail.get(fid):
            for d in packing_lines:
                rr = dict(d)
                if float(rr.get("unit_weight") or 0) <= 0 and rr.get("sku") in unit_weights:
                    rr["unit_weight"] = float(unit_weights[rr["sku"]])
                display_lines.append(rr)
        else:
            display_lines = lines
        shipments.append(
            {
                "fba_id": fid,
                "matched": bool(lines),
                "multi_sku": multi,
                "gross_weight": (saihu.get("shipment_weights") or {}).get(fid)
                or (lines[0].get("shipment_gross_weight") if lines else 0),
                "lines": display_lines,
                "agg_lines": lines,
                "missing_unit_weight_skus": missing if multi else [],
            }
        )
    return {
        "ok": True,
        "ticket_id": ticket_id,
        "ext_nos": rows[0].get("ext_nos") or "",
        "freight": sum(float(x.get("amount_cny") or 0) for x in rows),
        "chargeable_weight": sum(float(x.get("chargeable_weight") or 0) for x in rows),
        "fba_ids": fba_ids,
        "shipments": shipments,
        "needs_unit_weight": needs_unit_weight,
        "skus_needing_weight": skus_needing_weight,
    }


@app.route("/")
def index():
    state = _load_state()
    selected_ticket = state.get("selected_ticket_id") or ""
    ticket_preview = None
    if selected_ticket and state.get("bill") and state.get("saihu"):
        ticket_preview = _ticket_packing_preview(state, selected_ticket)
    ticket_allocs = state.get("ticket_allocs") or {}
    alloc = state.get("alloc")
    return render_template(
        "index.html",
        bill=state.get("bill"),
        saihu=state.get("saihu"),
        alloc=alloc,
        ticket_allocs=ticket_allocs,
        tickets_grouped=(alloc or {}).get("tickets_grouped") or [],
        selected_ticket_id=selected_ticket,
        ticket_preview=ticket_preview,
        current_ticket_result=state.get("current_ticket_result"),
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
        flash("未解析到账单明细行，请确认文件含已整理的明细表（已跳过汇总表）", "error")
    else:
        sheets = result.get("sheets") or []
        sheet_label = "、".join(sheets) if sheets else (result.get("sheet") or "未知")
        mode = result.get("sheet_select") or ""
        mode_hint = {
            "cleaned": "已整理表头",
            "detail": "明细表头",
            "explicit": "指定工作表",
            "all": "全部工作表",
        }.get(mode, mode)
        flash(
            f"账单解析成功：使用工作表「{sheet_label}」"
            + (f"（{mode_hint}）" if mode_hint else "")
            + f"；运费合计 ¥{result['total_freight']:.2f}，"
            f"{result['freight_row_count']} 行 / {result['ticket_count']} 票"
            "（分摊列：G扩展单号 / K收费重 / R金额）",
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
        flash(
            f"装箱信息解析成功：{auto['row_count']} 条明细"
            f"（聚合 {auto.get('agg_row_count', auto['row_count'])} 个货件×SKU），"
            f"{auto['fba_count']} 个货件，{auto['sku_count']} 个 SKU",
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


@app.route("/select_ticket", methods=["POST"])
def select_ticket():
    state = _load_state()
    if not state.get("bill") or not state.get("saihu"):
        flash("请先完成步骤 1（账单）与步骤 2（装箱）", "error")
        return redirect(url_for("index"))
    ticket_id = (request.form.get("ticket_id") or "").strip()
    if not ticket_id:
        flash("请选择一票", "error")
        return redirect(url_for("index") + "#allocate")
    state["selected_ticket_id"] = ticket_id
    state.pop("current_ticket_result", None)
    state.pop("alloc_error", None)
    _save_state(state)
    return redirect(url_for("index") + "#allocate")


@app.route("/allocate_ticket", methods=["POST"])
def allocate_ticket_route():
    """Compute ONE ticket: user supplies unit packed weights for multi-SKU, then clicks."""
    state = _load_state()
    bill = state.get("bill")
    saihu = state.get("saihu")
    if not bill or not bill.get("rows"):
        flash("请先上传并解析物流账单", "error")
        return redirect(url_for("index"))
    if not saihu or not saihu.get("rows"):
        flash("请先上传并解析赛狐装箱信息", "error")
        return redirect(url_for("index") + "#saihu")

    ticket_id = (request.form.get("ticket_id") or state.get("selected_ticket_id") or "").strip()
    if not ticket_id:
        flash("请先选择一票再计算", "error")
        return redirect(url_for("index") + "#allocate")

    # Collect per-SKU unit packed weights from the form (uw_<sku>)
    form_weights: dict[str, float] = {}
    for key, val in request.form.items():
        if not key.startswith("uw_"):
            continue
        sku = key[3:]
        try:
            w = float(str(val).strip().replace(",", "")) if str(val).strip() else 0.0
        except ValueError:
            w = 0.0
        if sku and w > 0:
            form_weights[sku] = w

    # Merge into session unit_weights (keep prior entries)
    unit_weights = dict(state.get("unit_weights") or {})
    unit_weights.update(form_weights)
    state["unit_weights"] = unit_weights
    state["selected_ticket_id"] = ticket_id

    preview = _ticket_packing_preview(state, ticket_id)
    if preview.get("needs_unit_weight"):
        still_missing = [
            s for s in preview.get("skus_needing_weight") or [] if s not in unit_weights
        ]
        # Re-check after form weights
        still_missing = [s for s in still_missing if float(unit_weights.get(s) or 0) <= 0]
        # Also check SKUs that appear in multi-SKU shipments for this ticket
        still_missing = []
        for ship in preview.get("shipments") or []:
            if not ship.get("multi_sku"):
                continue
            for ln in ship.get("lines") or []:
                sku = ln["sku"]
                uw = float(ln.get("unit_weight") or 0) or float(unit_weights.get(sku) or 0)
                if uw <= 0 and sku not in still_missing:
                    still_missing.append(sku)
        if still_missing:
            state["alloc_error"] = (
                "多SKU货件需先填写单件含包装重量(kg)再计算："
                + "、".join(still_missing)
            )
            state.pop("current_ticket_result", None)
            _save_state(state)
            flash(state["alloc_error"], "error")
            return redirect(url_for("index") + "#allocate")

    rows = _saihu_rows_with_weights(state, form_weights)
    result = allocate_one_ticket(
        bill["rows"],
        rows,
        ticket_id,
        shop="",
        shipment_weights=saihu.get("shipment_weights"),
    )
    if not result.get("ok"):
        state["alloc_error"] = result.get("error") or "本票分摊失败"
        state.pop("current_ticket_result", None)
        _save_state(state)
        flash(state["alloc_error"], "error")
        return redirect(url_for("index") + "#allocate")

    missing = result.get("multi_sku_missing_unit_weight") or []
    if missing and not result.get("sku_results"):
        ids = "、".join(m["fba_id"] for m in missing)
        state["alloc_error"] = f"本票多SKU货件仍缺单件含包装重量，未计算出结果：{ids}"
        state["current_ticket_result"] = result
        _save_state(state)
        flash(state["alloc_error"], "error")
        return redirect(url_for("index") + "#result")

    # Attach packing preview so UI can group ticket + FBAs + packing + calc together
    result["packing_preview"] = preview if preview.get("ok") else None

    # Persist this ticket's result
    ticket_allocs = dict(state.get("ticket_allocs") or {})
    ticket_allocs[ticket_id] = result
    state["ticket_allocs"] = ticket_allocs
    state["current_ticket_result"] = result
    state["alloc"] = _merge_ticket_allocs(ticket_allocs)  # cumulative for export
    state.pop("alloc_error", None)
    _save_state(state)

    flash(
        f"本票已计算：运费 ¥{result.get('ticket_freight', result.get('included_bill_freight', 0)):.2f}，"
        f"{result.get('sku_count', 0)} 个 SKU，"
        f"分摊 ¥{result.get('total_allocated', 0):.2f}",
        "ok",
    )
    if missing:
        flash(
            "部分多SKU货件因缺重量被跳过："
            + "、".join(m["fba_id"] for m in missing[:6]),
            "error",
        )
    return redirect(url_for("index") + "#result")


def _merge_ticket_allocs(ticket_allocs: dict) -> dict:
    """Merge per-ticket results into a cumulative alloc for export / summary (SKU only)."""
    sku_map: dict[str, dict] = {}
    shipment_allocs: list = []
    unmatched: list = []
    multi_missing: list = []
    matched_fba: set[str] = set()
    included = 0.0
    tickets_grouped: list = []
    for tid, res in ticket_allocs.items():
        included += float(res.get("included_bill_freight") or 0)
        shipment_allocs.extend(res.get("shipment_allocs") or [])
        unmatched.extend(res.get("unmatched_freight") or [])
        multi_missing.extend(res.get("multi_sku_missing_unit_weight") or [])
        matched_fba.update(
            fid
            for r in (res.get("sku_results") or [])
            for fid in (r.get("fba_ids") or [])
        )
        for r in res.get("sku_results") or []:
            key = r["sku"]
            agg = sku_map.setdefault(
                key,
                {
                    "sku": r["sku"],
                    "total_freight": 0.0,
                    "total_qty": 0.0,
                    "shipment_count": 0,
                    "fba_ids": set(),
                },
            )
            agg["total_freight"] += float(r["total_freight"])
            agg["total_qty"] += float(r["total_qty"])
            agg["shipment_count"] += int(r.get("shipment_count") or 0)
            agg["fba_ids"].update(r.get("fba_ids") or [])
        tickets_grouped.append(
            {
                "ticket_id": res.get("ticket_id") or tid,
                "ext_nos": res.get("ext_nos") or "",
                "ticket_freight": res.get("ticket_freight") or res.get("included_bill_freight"),
                "fba_ids": res.get("fba_ids") or [],
                "sku_results": res.get("sku_results") or [],
                "shipment_allocs": res.get("shipment_allocs") or [],
                "calc_steps": res.get("calc_steps") or [],
                "packing_preview": res.get("packing_preview"),
            }
        )
    sku_results = []
    for agg in sorted(sku_map.values(), key=lambda x: x["sku"]):
        qty = agg["total_qty"]
        sku_results.append(
            {
                "sku": agg["sku"],
                "total_freight": round(agg["total_freight"], 6),
                "total_qty": qty,
                "unit_first_leg": round(agg["total_freight"] / qty, 6) if qty else 0,
                "shipment_count": agg["shipment_count"],
                "fba_ids": sorted(agg["fba_ids"]),
            }
        )
    total_allocated = round(sum(x["total_freight"] for x in sku_results), 4)
    return {
        "ok": True,
        "sku_results": sku_results,
        "shipment_allocs": shipment_allocs,
        "unmatched_freight": unmatched,
        "matched_fba_count": len(matched_fba),
        "sku_count": len(sku_results),
        "total_allocated": total_allocated,
        "included_bill_freight": round(included, 4),
        "allocation_balance_ok": abs(total_allocated - round(included, 4)) < 0.05,
        "multi_sku_missing_unit_weight": multi_missing,
        "tickets_done": len(ticket_allocs),
        "tickets_grouped": tickets_grouped,
    }


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
    ws2 = wb.create_sheet("按票汇总")
    ticket_rows = []
    for t in bill.get("tickets") or []:
        ticket_rows.append(
            {
                "票标识": t.get("ticket_id") or "",
                "扩展单号": t.get("ext_nos") or "",
                "运费合计": round(t["freight_total"], 4),
                "收费重": round(t["weight_total"], 4),
                "FBA货件号(扩展单号)": ",".join(t.get("fba_ids") or []),
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
                "票数": bill.get("ticket_count"),
                "运费合计_CNY": bill.get("total_freight"),
                "收费重合计": bill.get("total_weight"),
                "明细表列表": ",".join(bill.get("sheets") or ([bill.get("sheet")] if bill.get("sheet") else [])),
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
                "票标识": r.get("ticket_id") or r.get("waybill_no") or "",
                "扩展单号": r.get("ext_nos") or "",
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
                    "票标识": u.get("ticket_id") or "",
                    "扩展单号": u.get("ext_nos") or u.get("waybill_no") or "",
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
