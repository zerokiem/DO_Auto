// DOffice dashboard - kich hoat chay tac vu, xem nhat ky truc tiep qua
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

  // ---------- Dang nhap lai ----------

  function setLoginBoxVisible(visible, message) {
    const box = document.getElementById("login-status-box");
    const msgEl = document.getElementById("login-status-message");
    if (!box) return;
    box.hidden = !visible;
    if (msgEl && message) msgEl.textContent = message;
  }

  function setLoginStartButtonBusy(busy) {
    const btn = document.getElementById("login-start-btn");
    if (btn) btn.disabled = busy;
  }

  async function pollLoginStatus() {
    let data;
    try {
      const res = await fetch("/api/login/status");
      data = await res.json();
    } catch (e) {
      setTimeout(pollLoginStatus, 3000);
      return;
    }

    if (data.is_active) {
      setLoginBoxVisible(true, data.message);
      setLoginStartButtonBusy(true);
      setTimeout(pollLoginStatus, 2000);
      return;
    }

    setLoginStartButtonBusy(false);
    if (data.message) {
      setLoginBoxVisible(true, data.message);
      setTimeout(() => window.location.reload(), 2000);
    } else {
      setLoginBoxVisible(false, "");
    }
  }

  async function startLogin() {
    // May khong man hinh (NAS/Docker) hien them 2 o nay - xem settings.html.
    // Neu co mat, day la dang nhap headless: bat buoc phai co username/password.
    const usernameEl = document.getElementById("doffice_username");
    const passwordEl = document.getElementById("doffice_password");
    const isHeadlessForm = !!(usernameEl && passwordEl);

    let fetchOptions = { method: "POST" };
    if (isHeadlessForm) {
      const username = usernameEl.value.trim();
      const password = passwordEl.value;
      if (!username || !password) {
        alert("Điền tài khoản và mật khẩu trước khi bấm Đăng nhập.");
        return;
      }
      fetchOptions.headers = { "Content-Type": "application/json" };
      fetchOptions.body = JSON.stringify({ username, password });
    }

    setLoginStartButtonBusy(true);
    let res, data;
    try {
      res = await fetch("/api/login/start", fetchOptions);
      data = await res.json();
    } catch (e) {
      setLoginStartButtonBusy(false);
      return;
    } finally {
      // Xoa mat khau khoi form ngay sau khi gui, du thanh cong hay khong -
      // khong de nam lai tren man hinh/DOM lau hon can thiet.
      if (passwordEl) passwordEl.value = "";
    }
    if (!data.started) {
      setLoginStartButtonBusy(false);
      alert(data.error || "Không mở được cửa sổ đăng nhập.");
      return;
    }
    setLoginBoxVisible(true, isHeadlessForm ? "Đang đăng nhập..." : "Đang mở cửa sổ đăng nhập trên máy chủ...");
    pollLoginStatus();
  }

  async function confirmLogin() {
    const btn = document.getElementById("login-confirm-btn");
    if (btn) btn.disabled = true;
    try {
      await fetch("/api/login/confirm", { method: "POST" });
    } catch (e) {
      // bo qua, pollLoginStatus se tiep tuc cap nhat trang thai
    }
    if (btn) btn.disabled = false;
  }

  // ---------- Gui thu Telegram ----------

  async function testTelegram() {
    const btn = document.getElementById("telegram-test-btn");
    const resultEl = document.getElementById("telegram-test-result");
    const tokenEl = document.getElementById("telegram_bot_token");
    const chatIdEl = document.getElementById("telegram_chat_id");
    if (!btn || !resultEl || !tokenEl || !chatIdEl) return;

    const botToken = tokenEl.value.trim();
    const chatId = chatIdEl.value.trim();
    if (!botToken || !chatId) {
      resultEl.textContent = "Điền Bot Token và Chat ID trước khi gửi thử.";
      resultEl.hidden = false;
      return;
    }

    btn.disabled = true;
    resultEl.hidden = true;
    try {
      const res = await fetch("/api/telegram/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ bot_token: botToken, chat_id: chatId }),
      });
      const data = await res.json();
      resultEl.textContent = data.ok
        ? "✅ Đã gửi tin nhắn thử, kiểm tra Telegram."
        : "❌ Gửi thất bại: " + (data.error || "không rõ lỗi.");
    } catch (e) {
      resultEl.textContent = "❌ Không gọi được máy chủ.";
    }
    resultEl.hidden = false;
    btn.disabled = false;
  }

  document.addEventListener("DOMContentLoaded", () => {
    const runBtn = document.getElementById("run-btn");
    if (runBtn) {
      runBtn.addEventListener("click", startRun);
    }

    const loginStartBtn = document.getElementById("login-start-btn");
    if (loginStartBtn) {
      loginStartBtn.addEventListener("click", startLogin);
    }
    const loginConfirmBtn = document.getElementById("login-confirm-btn");
    if (loginConfirmBtn) {
      loginConfirmBtn.addEventListener("click", confirmLogin);
    }
    const telegramTestBtn = document.getElementById("telegram-test-btn");
    if (telegramTestBtn) {
      telegramTestBtn.addEventListener("click", testTelegram);
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

    // Tuong tu, neu dang co 1 luot dang nhap dang cho xac nhan.
    const initialLogin = window.__INITIAL_LOGIN_STATUS__;
    if (initialLogin && initialLogin.is_active) {
      setLoginBoxVisible(true, initialLogin.message);
      setLoginStartButtonBusy(true);
      pollLoginStatus();
    }
  });
})();
