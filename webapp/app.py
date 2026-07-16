"""
Flask app cho DOffice Auto web dashboard.

Day la cong cu NOI BO chay tren may cua ban (hoac 1 may trong mang noi bo/
Tailscale), KHONG co lop dang nhap rieng - vi Playwright/Chromium can chay
ngay tren may nay va DOffice chi truy cap duoc tu mang noi bo. Xem run_web.py
o thu muc goc de biet cach khoi dong.
"""
from __future__ import annotations

import json
import queue
import time
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, render_template, request, send_file, url_for

import config as base_config
from do_auto import excel_log, history, settings_store
from webapp.run_manager import RunManager

app = Flask(__name__)
app.jinja_env.globals["zip"] = zip
run_manager = RunManager(base_config)


def _ordered_tasks(cfg):
    return sorted(cfg.TASKS.items(), key=lambda kv: kv[1].sheet_order)


def _auth_state_info(cfg) -> dict:
    p = Path(cfg.AUTH_STATE)
    if not p.exists():
        return {"exists": False, "age_days": None}
    age_days = (time.time() - p.stat().st_mtime) / 86400
    return {"exists": True, "age_days": round(age_days, 1)}


def _read_sheet_preview(excel_file: Path, sheet_name: str, limit: int = 300):
    from openpyxl import load_workbook

    excel_file = Path(excel_file)
    if not excel_file.exists():
        return excel_log.HEADERS, []

    wb = load_workbook(excel_file, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return excel_log.HEADERS, []
    ws = wb[sheet_name]

    all_rows = []
    for row in ws.iter_rows(min_row=excel_log.DATA_START_ROW, values_only=True):
        if not row or not any(v not in (None, "") for v in row):
            continue
        all_rows.append(row)
    wb.close()

    tail = all_rows[-limit:]
    tail.reverse()  # moi nhat len tren
    return excel_log.HEADERS, tail


@app.route("/")
def dashboard():
    effective_cfg = settings_store.build_effective_config(base_config)
    tasks = _ordered_tasks(effective_cfg)

    recent_by_task = {}
    for key, _task in tasks:
        rows = history.list_recent_runs_for_task(Path(effective_cfg.HISTORY_DB), key, limit=1)
        recent_by_task[key] = rows[0] if rows else None

    return render_template(
        "dashboard.html",
        tasks=tasks,
        recent_by_task=recent_by_task,
        auth=_auth_state_info(effective_cfg),
        status=run_manager.status(),
        excel_file=str(effective_cfg.EXCEL_FILE),
    )


@app.post("/api/run")
def api_run():
    payload = request.get_json(force=True, silent=True) or {}
    task_keys = payload.get("task_keys") or []
    headless = bool(payload.get("headless", True))
    test_mode = bool(payload.get("test_mode", False))

    effective_cfg = settings_store.build_effective_config(base_config)

    unknown = [k for k in task_keys if k not in effective_cfg.TASKS]
    if unknown:
        return jsonify({"started": False, "error": f"Tác vụ không hợp lệ: {', '.join(unknown)}"}), 400
    if not task_keys:
        return jsonify({"started": False, "error": "Chưa chọn tác vụ nào."}), 400
    if not Path(effective_cfg.AUTH_STATE).exists():
        return (
            jsonify(
                {
                    "started": False,
                    "error": "Chưa có phiên đăng nhập. Chạy 'python login_save_state.py' trên máy này trước.",
                }
            ),
            400,
        )

    started = run_manager.start(task_keys, headless=headless, test_mode=test_mode)
    if not started:
        return jsonify({"started": False, "error": "Đang có 1 lượt chạy khác, vui lòng đợi."}), 409
    return jsonify({"started": True})


@app.get("/api/status")
def api_status():
    return jsonify(run_manager.status())


@app.get("/api/logs/stream")
def api_logs_stream():
    def gen():
        q = run_manager.broadcaster.subscribe()
        try:
            for line in list(run_manager.broadcaster.history)[-200:]:
                yield f"data: {json.dumps(line)}\n\n"
            while True:
                try:
                    line = q.get(timeout=15)
                    yield f"data: {json.dumps(line)}\n\n"
                except queue.Empty:
                    yield ": ping\n\n"
        except GeneratorExit:
            pass
        finally:
            run_manager.broadcaster.unsubscribe(q)

    return Response(gen(), mimetype="text/event-stream")


@app.route("/history")
def history_page():
    effective_cfg = settings_store.build_effective_config(base_config)
    rows = history.list_recent_runs(Path(effective_cfg.HISTORY_DB), limit=150)
    return render_template("history.html", rows=rows)


@app.route("/excel")
def excel_page():
    effective_cfg = settings_store.build_effective_config(base_config)
    tasks = _ordered_tasks(effective_cfg)
    sheet_names = [t.sheet_name for _, t in tasks]

    sheet = request.args.get("sheet")
    if not sheet or sheet not in sheet_names:
        sheet = sheet_names[0] if sheet_names else None

    columns, rows, error = [], [], None
    if sheet:
        try:
            columns, rows = _read_sheet_preview(effective_cfg.EXCEL_FILE, sheet, limit=300)
        except Exception as e:
            error = str(e)

    return render_template(
        "excel_view.html",
        sheet_names=sheet_names,
        active_sheet=sheet,
        columns=columns,
        rows=rows,
        error=error,
        excel_file=str(effective_cfg.EXCEL_FILE),
        row_count=len(rows),
    )


@app.get("/excel/download")
def excel_download():
    effective_cfg = settings_store.build_effective_config(base_config)
    path = Path(effective_cfg.EXCEL_FILE)
    if not path.exists():
        return "Chưa có file Excel nào được tạo.", 404
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    if request.method == "POST":
        overrides = {"common": {}, "tasks": {}}

        overrides["common"]["stop_when_duplicate_found"] = request.form.get("stop_when_duplicate_found") == "on"
        overrides["common"]["duplicate_check_mode"] = request.form.get(
            "duplicate_check_mode", base_config.DUPLICATE_CHECK_MODE
        )
        try:
            overrides["common"]["slow_mo_ms"] = int(request.form.get("slow_mo_ms", base_config.SLOW_MO_MS))
        except (TypeError, ValueError):
            overrides["common"]["slow_mo_ms"] = base_config.SLOW_MO_MS

        for key, base_task in base_config.TASKS.items():
            try:
                max_docs = int(request.form.get(f"{key}_max_documents") or base_task.max_documents)
            except (TypeError, ValueError):
                max_docs = base_task.max_documents

            overrides["tasks"][key] = {
                "enabled": request.form.get(f"{key}_enabled") == "on",
                "role_pattern": request.form.get(f"{key}_role_pattern", base_task.role_pattern).strip(),
                "max_documents": max_docs,
                "enable_finish": request.form.get(f"{key}_enable_finish") == "on",
                "enable_download_pdf": request.form.get(f"{key}_enable_download_pdf") == "on",
                "ask_confirm_before_finish": request.form.get(f"{key}_ask_confirm_before_finish") == "on",
            }

        settings_store.save_overrides(overrides)
        return redirect(url_for("settings_page", saved=1))

    effective_cfg = settings_store.build_effective_config(base_config)
    tasks = _ordered_tasks(effective_cfg)
    common = {
        "STOP_WHEN_DUPLICATE_FOUND": effective_cfg.STOP_WHEN_DUPLICATE_FOUND,
        "DUPLICATE_CHECK_MODE": effective_cfg.DUPLICATE_CHECK_MODE,
        "SLOW_MO_MS": effective_cfg.SLOW_MO_MS,
    }
    saved = request.args.get("saved") == "1"
    return render_template("settings.html", tasks=tasks, common=common, saved=saved)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8877, threaded=True, debug=False)
