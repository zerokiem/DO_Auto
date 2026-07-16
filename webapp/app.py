"""
Flask app cho DOffice web dashboard.

Day la cong cu NOI BO chay tren may cua ban (hoac 1 may trong mang noi bo/
Tailscale), KHONG co lop dang nhap rieng - vi Playwright/Chromium can chay
ngay tren may nay va DOffice chi truy cap duoc tu mang noi bo. Xem run_web.py
o thu muc goc de biet cach khoi dong.
"""
from __future__ import annotations

import json
import queue
import time
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, abort, jsonify, redirect, render_template, request, send_file, url_for

import config as base_config
from do_auto import excel_log, history, scheduler as scheduler_mod, settings_store
from webapp.login_manager import LoginManager
from webapp.run_manager import RunManager

PROJECT_DIR = Path(__file__).resolve().parent.parent

app = Flask(__name__)
app.jinja_env.globals["zip"] = zip
run_manager = RunManager(base_config)
login_manager = LoginManager(base_config)


@app.context_processor
def inject_globals():
    return {"current_year": datetime.now().year}


def _ordered_tasks(cfg):
    return sorted(cfg.TASKS.items(), key=lambda kv: kv[1].sheet_order)


def _auth_state_info(cfg) -> dict:
    p = Path(cfg.AUTH_STATE)
    if not p.exists():
        return {"exists": False, "age_days": None}
    age_days = (time.time() - p.stat().st_mtime) / 86400
    return {"exists": True, "age_days": round(age_days, 1)}


def _read_sheet_preview(excel_file: Path, sheet_name: str, limit=300):
    """Doc sheet, tra ve (headers, rows_moi_nhat_len_tren, tong_so_dong).
    limit=None nghia la lay tat ca."""
    from openpyxl import load_workbook

    excel_file = Path(excel_file)
    if not excel_file.exists():
        return excel_log.HEADERS, [], 0

    wb = load_workbook(excel_file, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return excel_log.HEADERS, [], 0
    ws = wb[sheet_name]

    all_rows = []
    for row in ws.iter_rows(min_row=excel_log.DATA_START_ROW, values_only=True):
        if not row or not any(v not in (None, "") for v in row):
            continue
        all_rows.append(row)
    wb.close()

    total = len(all_rows)
    tail = all_rows if limit is None else all_rows[-limit:]
    tail = list(tail)
    tail.reverse()  # moi nhat len tren
    return excel_log.HEADERS, tail, total


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
        login_status=login_manager.status(),
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
                    "error": "Chưa có phiên đăng nhập. Bấm 'Đăng nhập lại' ở trên hoặc chạy "
                    "'python login_save_state.py' trên máy này trước.",
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


@app.post("/api/login/start")
def api_login_start():
    started = login_manager.start()
    if not started:
        return jsonify({"started": False, "error": "Đang có 1 lượt đăng nhập khác đang mở."}), 409
    return jsonify({"started": True})


@app.post("/api/login/confirm")
def api_login_confirm():
    ok = login_manager.confirm()
    if not ok:
        return jsonify({"ok": False, "error": "Chưa có lượt đăng nhập nào đang chờ xác nhận."}), 400
    return jsonify({"ok": True})


@app.get("/api/login/status")
def api_login_status():
    return jsonify(login_manager.status())


@app.post("/api/telegram/test")
def api_telegram_test():
    payload = request.get_json(force=True, silent=True) or {}
    bot_token = (payload.get("bot_token") or "").strip()
    chat_id = (payload.get("chat_id") or "").strip()

    from do_auto import notify

    test_message = "✅ DOffice - đây là tin nhắn thử. Nếu bạn thấy tin này, cấu hình Telegram đã đúng."
    ok, error = notify.send_telegram_message(bot_token, chat_id, test_message)
    if ok:
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": error}), 400


@app.route("/history")
def history_page():
    effective_cfg = settings_store.build_effective_config(base_config)
    rows = history.list_recent_runs(Path(effective_cfg.HISTORY_DB), limit=150)
    return render_template("history.html", rows=rows)


@app.get("/logs/<path:filename>")
def view_log(filename):
    effective_cfg = settings_store.build_effective_config(base_config)
    logs_dir = Path(effective_cfg.LOGS_DIR).resolve()
    target = (logs_dir / filename).resolve()
    # Chan doc file ngoai LOGS_DIR (vd truyen "../../config.py" qua filename).
    if logs_dir not in target.parents and target != logs_dir:
        abort(403)
    if not target.exists() or not target.is_file():
        abort(404)
    content = target.read_text(encoding="utf-8", errors="replace")
    return render_template("log_view.html", filename=filename, content=content)


@app.route("/excel")
def excel_page():
    effective_cfg = settings_store.build_effective_config(base_config)
    tasks = _ordered_tasks(effective_cfg)
    sheet_names = [t.sheet_name for _, t in tasks]

    sheet = request.args.get("sheet")
    if not sheet or sheet not in sheet_names:
        sheet = sheet_names[0] if sheet_names else None

    # So dong hien thi: tham so ?limit= (so nguyen hoac "all"). Mac dinh 300.
    limit_options = [50, 100, 200, 300, 500, 1000]
    limit_arg = (request.args.get("limit") or "300").strip().lower()
    if limit_arg in ("all", "0", "tatca", "tat_ca"):
        limit_arg = "all"
        limit = None
    else:
        try:
            limit = max(1, int(limit_arg))
            limit_arg = str(limit)
        except ValueError:
            limit, limit_arg = 300, "300"

    columns, rows, error, total = [], [], None, 0
    if sheet:
        try:
            columns, rows, total = _read_sheet_preview(effective_cfg.EXCEL_FILE, sheet, limit=limit)
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
        total_count=total,
        limit_options=limit_options,
        current_limit=limit_arg,
        column_widths=excel_log.COLUMN_WIDTHS,
    )


@app.get("/vb/<path:relpath>")
def serve_vb(relpath):
    """Phuc vu file PDF trong thu muc du lieu (DOWNLOAD_BASE_DIR = /data) qua web,
    de hyperlink trong Excel mo duoc tren dien thoai/Tailscale. Chi cho phep file
    NAM TRONG thu muc du lieu (chan '../' thoat ra ngoai)."""
    effective_cfg = settings_store.build_effective_config(base_config)
    base = Path(effective_cfg.DOWNLOAD_BASE_DIR).resolve()
    target = (base / relpath).resolve()
    if base != target and base not in target.parents:
        abort(403)
    if not target.exists() or not target.is_file():
        abort(404)
    return send_file(target)


@app.get("/excel/download")
def excel_download():
    effective_cfg = settings_store.build_effective_config(base_config)
    path = Path(effective_cfg.EXCEL_FILE)
    if not path.exists():
        return "Chưa có file Excel nào được tạo.", 404
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/settings", methods=["GET", "POST"])
def settings_page():
    # Dung effective_cfg (qua settings_store, co getattr + fallback mac dinh)
    # thay vi doc truc tiep base_config.X lam gia tri fallback ben duoi - de
    # KHONG BAO GIO crash du config.py cua nguoi dung cu hon code hien tai.
    effective_cfg = settings_store.build_effective_config(base_config)

    if request.method == "POST":
        doffice_url = request.form.get("doffice_url", "").strip()
        common_updates = {
            "STOP_WHEN_DUPLICATE_FOUND": request.form.get("stop_when_duplicate_found") == "on",
            "DUPLICATE_CHECK_MODE": request.form.get("duplicate_check_mode", effective_cfg.DUPLICATE_CHECK_MODE),
            "ENABLE_TELEGRAM_NOTIFY": request.form.get("enable_telegram_notify") == "on",
            "TELEGRAM_BOT_TOKEN": request.form.get("telegram_bot_token", effective_cfg.TELEGRAM_BOT_TOKEN).strip(),
            "TELEGRAM_CHAT_ID": request.form.get("telegram_chat_id", effective_cfg.TELEGRAM_CHAT_ID).strip(),
            "TELEGRAM_NOTIFY_ONLY_IF_NEW": request.form.get("telegram_notify_only_if_new") == "on",
        }
        if doffice_url:
            common_updates["DOFFICE_URL"] = doffice_url
        try:
            common_updates["SLOW_MO_MS"] = int(request.form.get("slow_mo_ms", effective_cfg.SLOW_MO_MS))
        except (TypeError, ValueError):
            common_updates["SLOW_MO_MS"] = effective_cfg.SLOW_MO_MS

        error = None
        try:
            settings_store.update_common_fields(common_updates)
            for key, eff_task in effective_cfg.TASKS.items():
                try:
                    max_docs = int(request.form.get(f"{key}_max_documents") or eff_task.max_documents)
                except (TypeError, ValueError):
                    max_docs = eff_task.max_documents

                settings_store.update_task_fields(
                    key,
                    {
                        "enabled": request.form.get(f"{key}_enabled") == "on",
                        "role_pattern": request.form.get(f"{key}_role_pattern", eff_task.role_pattern).strip(),
                        "max_documents": max_docs,
                        "enable_finish": request.form.get(f"{key}_enable_finish") == "on",
                        "enable_download_pdf": request.form.get(f"{key}_enable_download_pdf") == "on",
                        "ask_confirm_before_finish": request.form.get(f"{key}_ask_confirm_before_finish") == "on",
                    },
                )
            settings_store.reload_config(base_config)
        except Exception as e:
            error = str(e)

        if error:
            return redirect(url_for("settings_page", error=error))
        return redirect(url_for("settings_page", saved=1))

    tasks = _ordered_tasks(effective_cfg)
    common = {
        "DOFFICE_URL": effective_cfg.DOFFICE_URL,
        "STOP_WHEN_DUPLICATE_FOUND": effective_cfg.STOP_WHEN_DUPLICATE_FOUND,
        "DUPLICATE_CHECK_MODE": effective_cfg.DUPLICATE_CHECK_MODE,
        "SLOW_MO_MS": effective_cfg.SLOW_MO_MS,
        "ENABLE_TELEGRAM_NOTIFY": effective_cfg.ENABLE_TELEGRAM_NOTIFY,
        "TELEGRAM_BOT_TOKEN": effective_cfg.TELEGRAM_BOT_TOKEN,
        "TELEGRAM_CHAT_ID": effective_cfg.TELEGRAM_CHAT_ID,
        "TELEGRAM_NOTIFY_ONLY_IF_NEW": effective_cfg.TELEGRAM_NOTIFY_ONLY_IF_NEW,
    }
    saved = request.args.get("saved") == "1"
    error = request.args.get("error")
    return render_template(
        "settings.html",
        tasks=tasks,
        common=common,
        saved=saved,
        error=error,
        auth=_auth_state_info(effective_cfg),
        login_status=login_manager.status(),
    )


@app.route("/scheduler", methods=["GET", "POST"])
def scheduler_page():
    if request.method == "POST":
        t1 = request.form.get("time1", "").strip()
        t2 = request.form.get("time2", "").strip()
        times = [t for t in (t1, t2) if t]

        if not times:
            ok, message = scheduler_mod.remove_schedule()
        else:
            ok, message = scheduler_mod.apply_schedule(PROJECT_DIR, times)

        return render_template("scheduler.html", times=times, saved=True, ok=ok, message=message)

    current_times = scheduler_mod.get_current_times()
    return render_template("scheduler.html", times=current_times, saved=False, ok=None, message="")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8877, threaded=True, debug=False)
