// DOffice Auto dashboard - kich hoat chay tac vu, xem nhat ky truc tiep qua
// Server-Sent Events, theo doi trang thai dinh ky cho toi khi chay xong.

(function () {
  "use strict";

  const state = { eventSource: null };

  function selectedTaskKeys() {
    return Array.from(document.querySelectorAll(".task-checkbox:checked")).map((el) => el.value);
  }

  function showRunError(msg) {
    const el = document.getElementById("run-error");
    if (!el) return;
    el.textContent = msg;
    el.hidden = false;
  }

  function hideRunError() {
    const el = document.getElementById("run-error");
    if (el) el.hidden = true;
  }

  function setRunButtonBusy(busy) {
    const btn = document.getElementById("run-btn");
    if (!btn) return;
    btn.disabled = busy;
    btn.textContent = busy ? "Đang chạy..." : "Chạy các tác vụ đã chọn";
  }

  function openConsole() {
    const panel = document.getElementById("console-panel");
    const out = document.getElementById("console-output");
    if (panel) panel.hidden = false;
    if (out) out.textContent = "";
  }

  function appendLine(line) {
    const out = document.getElementById("console-output");
    if (!out) return;
    out.textContent += line + "\n";
    out.scrollTop = out.scrollHeight;
  }

  function connectLogStream() {
    if (state.eventSource) {
      state.eventSource.close();
    }
    const es = new EventSource("/api/logs/stream");
    es.onmessage = (evt) => {
      try {
        appendLine(JSON.parse(evt.data));
      } catch (e) {
        appendLine(evt.data);
      }
    };
    es.onerror = () => {
      // Trinh duyet se tu ket noi lai; khong can xu ly gi them o day.
    };
    state.eventSource = es;
  }

  function closeLogStreamSoon() {
    setTimeout(() => {
      if (state.eventSource) {
        state.eventSource.close();
        state.eventSource = null;
      }
    }, 1500);
  }

  async function pollStatus() {
    let data;
    try {
      const res = await fetch("/api/status");
      data = await res.json();
    } catch (e) {
      setTimeout(pollStatus, 3000);
      return;
    }

    if (data.is_running) {
      setTimeout(pollStatus, 2000);
      return;
    }

    setRunButtonBusy(false);
    if (data.last_error) {
      showRunError(data.last_error);
    }
    closeLogStreamSoon();
    // Tai lai trang sau vai giay de cap nhat "lan chay gan nhat" tren the tac vu.
    setTimeout(() => window.location.reload(), 1800);
  }

  async function startRun() {
    const taskKeys = selectedTaskKeys();
    if (taskKeys.length === 0) {
      showRunError("Chọn ít nhất 1 tác vụ trước khi chạy.");
      return;
    }

    const headless = document.getElementById("headless-toggle").checked;
    const testMode = document.getElementById("test-mode-toggle").checked;

    hideRunError();
    setRunButtonBusy(true);

    let res, data;
    try {
      res = await fetch("/api/run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ task_keys: taskKeys, headless, test_mode: testMode }),
      });
      data = await res.json();
    } catch (e) {
      showRunError("Không gọi được máy chủ. Kiểm tra lại server có đang chạy không.");
      setRunButtonBusy(false);
      return;
    }

    if (!data.started) {
      showRunError(data.error || "Không chạy được.");
      setRunButtonBusy(false);
      return;
    }

    openConsole();
    connectLogStream();
    pollStatus();
  }

  document.addEventListener("DOMContentLoaded", () => {
    const runBtn = document.getElementById("run-btn");
    if (runBtn) {
      runBtn.addEventListener("click", startRun);
    }

    // Neu mo lai trang trong luc dang co 1 lan chay dien ra (vd nguoi dung
    // reload), tu ket noi lai console va tiep tuc theo doi thay vi mat log.
    const initial = window.__INITIAL_STATUS__;
    if (initial && initial.is_running) {
      setRunButtonBusy(true);
      openConsole();
      connectLogStream();
      pollStatus();
    }
  });
})();
