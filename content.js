(() => {
  const PANEL_ID = "table-tools-panel";
  const EXPORT_BUTTON_ID = "table-export-btn";
  const SYNC_BUTTON_ID = "table-sync-feishu-btn";
  const TOAST_ID = "table-tools-toast";

  if (window.__tableToolsInjected) {
    return;
  }
  window.__tableToolsInjected = true;

  let exporting = false;
  let syncing = false;

  injectButtons();

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message?.type === "TRIGGER_EXPORT" || message?.type === "TRIGGER_SYNC") {
      runExport()
        .then((result) => sendResponse({ ok: true, ...result }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    if (message?.type === "TRIGGER_SYNC_TO_FEISHU") {
      runSyncToFeishu()
        .then((result) => sendResponse({ ok: true, ...result }))
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    return false;
  });

  function injectButtons() {
    if (document.getElementById(PANEL_ID)) {
      return;
    }

    const panel = document.createElement("div");
    panel.id = PANEL_ID;
    panel.style.cssText = [
      "position:fixed",
      "right:20px",
      "bottom:20px",
      "z-index:2147483647",
      "display:flex",
      "flex-direction:column",
      "gap:8px",
    ].join(";");

    const syncBtn = document.createElement("button");
    syncBtn.id = SYNC_BUTTON_ID;
    syncBtn.textContent = "同步飞书";
    syncBtn.title = "将当前页面表格同步到飞书多维表格";
    syncBtn.style.cssText = commonBtnStyle("#1677ff");
    syncBtn.addEventListener("click", () => {
      runSyncToFeishu().catch((error) => showToast(error.message, true));
    });

    const exportBtn = document.createElement("button");
    exportBtn.id = EXPORT_BUTTON_ID;
    exportBtn.textContent = "导出Excel";
    exportBtn.title = "将当前页面表格导出为Excel文件";
    exportBtn.style.cssText = commonBtnStyle("#16a34a");
    exportBtn.addEventListener("click", () => {
      runExport().catch((error) => showToast(error.message, true));
    });

    panel.appendChild(syncBtn);
    panel.appendChild(exportBtn);
    document.documentElement.appendChild(panel);
  }

  async function runExport() {
    if (exporting) {
      showToast("正在导出中，请稍候");
      return { exported: 0 };
    }

    const btn = document.getElementById(EXPORT_BUTTON_ID);
    exporting = true;
    setButtonState(btn, "导出中...", true);

    try {
      const payload = extractTableData();
      if (!payload.rows.length) {
        throw new Error("当前页面没有可导出的表格行数据");
      }

      const fileName = downloadExcel(payload);
      showToast(`已下载Excel：${fileName}`);
      return {
        exported: payload.rows.length,
        fileName,
      };
    } finally {
      exporting = false;
      setButtonState(btn, "导出Excel", false);
    }
  }

  async function runSyncToFeishu() {
    if (syncing) {
      showToast("正在同步中，请稍候");
      return { synced: 0 };
    }

    const btn = document.getElementById(SYNC_BUTTON_ID);
    syncing = true;
    setButtonState(btn, "同步中...", true);

    try {
      const payload = extractTableData();
      if (!payload.rows.length) {
        throw new Error("当前页面没有可同步的表格行数据");
      }

      const response = await chrome.runtime.sendMessage({
        type: "SYNC_TABLE_DATA",
        payload,
      });

      if (!response?.ok) {
        const message = response?.error || "同步失败";
        if (isFeishuConfigMissingError(message)) {
          await tryOpenOptionsPage();
          throw new Error("飞书参数未配置，已为你打开设置页。请填写后再试。");
        }
        throw new Error(message);
      }

      const duplicateSkipped = Number(response.duplicateSkipped || 0);
      const dedupeHint = duplicateSkipped > 0 ? `，已跳过重复 ${duplicateSkipped} 条（已同步过）` : "";
      showToast(`已同步到飞书：${response.created} 条${dedupeHint}`);
      return {
        synced: response.created || 0,
        duplicateSkipped,
        tableName: response.tableName || payload.tableName || "",
      };
    } finally {
      syncing = false;
      setButtonState(btn, "同步飞书", false);
    }
  }

  function extractTableData() {
    const table = findBestTable();
    if (!table) {
      throw new Error("页面中未找到可识别的表格");
    }

    const headerCells = Array.from(table.querySelectorAll("thead tr th"));
    let headers = headerCells.map((cell, index) => cleanHeaderName(readText(cell), index));

    let rowElements = Array.from(table.querySelectorAll("tbody tr"));

    if (headers.length === 0) {
      const allRows = Array.from(table.querySelectorAll("tr"));
      if (allRows.length === 0) {
        throw new Error("表格中没有行");
      }

      const firstRow = allRows[0];
      headers = Array.from(firstRow.querySelectorAll("th,td")).map((cell, index) =>
        cleanHeaderName(readText(cell), index)
      );
      rowElements = allRows.slice(1);
    }

    headers = makeUnique(headers);

    const matrix = rowElements
      .map((row) => {
        const cells = Array.from(row.querySelectorAll("td"));
        if (cells.length === 0) {
          return null;
        }
        const values = headers.map((_, index) => readText(cells[index]));
        return values;
      })
      .filter(Boolean);

    const keepIndexes = headers
      .map((name, index) => ({
        index,
        name,
        hasValue: matrix.some((r) => (r[index] || "").trim() !== ""),
      }))
      .filter((item) => item.hasValue || !/^列\d+$/.test(item.name));

    const columns = keepIndexes.map((item, idx) => ({
      key: `col_${idx + 1}`,
      name: item.name,
      sourceIndex: item.index,
    }));

    const rows = matrix
      .map((values) => {
        const row = {};
        for (const col of columns) {
          row[col.key] = values[col.sourceIndex] || "";
        }
        return row;
      })
      .filter((row) => Object.values(row).some((v) => (v || "").trim() !== ""));

    return {
      sourceUrl: location.href,
      sourceTitle: document.title,
      capturedAt: new Date().toISOString(),
      tableName: document.title,
      columns,
      rows,
    };
  }

  function findBestTable() {
    const tables = Array.from(document.querySelectorAll("table"));
    if (!tables.length) {
      return null;
    }

    const scored = tables
      .map((table) => {
        const rowCount = table.querySelectorAll("tbody tr").length || table.querySelectorAll("tr").length;
        const colCount = table.querySelectorAll("thead th").length || table.querySelector("tr")?.children.length || 0;
        return {
          table,
          score: rowCount * Math.max(colCount, 1),
          rowCount,
          colCount,
        };
      })
      .filter((item) => item.rowCount > 0 && item.colCount > 0)
      .sort((a, b) => b.score - a.score);

    return scored[0]?.table || null;
  }

  function readText(node) {
    if (!node) {
      return "";
    }

    return String(node.innerText || node.textContent || "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function cleanHeaderName(text, index) {
    const value = String(text || "")
      .replace(/[\r\n\t]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (!value) {
      return `列${index + 1}`;
    }

    return value.slice(0, 90);
  }

  function makeUnique(headers) {
    const seen = new Map();
    return headers.map((header) => {
      const count = seen.get(header) || 0;
      seen.set(header, count + 1);
      if (count === 0) {
        return header;
      }
      return `${header}_${count + 1}`;
    });
  }

  function commonBtnStyle(backgroundColor) {
    return [
      "padding:10px 14px",
      "border:none",
      "border-radius:10px",
      `background:${backgroundColor}`,
      "color:#fff",
      "font-size:14px",
      "font-weight:600",
      "cursor:pointer",
      "box-shadow:0 6px 16px rgba(0,0,0,.18)",
    ].join(";");
  }

  function setButtonState(btn, text, disabled) {
    if (!btn) {
      return;
    }
    btn.textContent = text;
    btn.disabled = disabled;
    btn.style.opacity = disabled ? "0.75" : "1";
    btn.style.cursor = disabled ? "not-allowed" : "pointer";
  }

  function downloadExcel(payload) {
    const headers = (payload.columns || []).map((col) => col.name);
    const metaHeaders = ["来源页面", "抓取时间"];
    const allHeaders = [...headers, ...metaHeaders];

    const rows = (payload.rows || []).map((row) => {
      const values = headers.map((name, index) => {
        const col = payload.columns[index];
        return col ? sanitizeForExcel(row[col.key]) : "";
      });
      values.push(
        sanitizeForExcel(payload.sourceUrl || location.href),
        sanitizeForExcel(payload.capturedAt || new Date().toISOString())
      );
      return values;
    });

    const tableHeaderHtml = allHeaders.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
    const tableBodyHtml = rows
      .map((cells) => `<tr>${cells.map((v) => `<td>${escapeHtml(v)}</td>`).join("")}</tr>`)
      .join("");

    const html = `<!doctype html>
<html>
  <head>
    <meta charset="UTF-8">
  </head>
  <body>
    <table border="1">
      <thead><tr>${tableHeaderHtml}</tr></thead>
      <tbody>${tableBodyHtml}</tbody>
    </table>
  </body>
</html>`;

    const blob = new Blob(["\uFEFF", html], { type: "application/vnd.ms-excel;charset=utf-8;" });
    const url = URL.createObjectURL(blob);

    const filename = buildExportFileName(payload.tableName);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    a.remove();

    window.setTimeout(() => URL.revokeObjectURL(url), 800);
    return filename;
  }

  function buildExportFileName(tableName) {
    const now = new Date();
    const pad = (num) => String(num).padStart(2, "0");
    const stamp = [
      now.getFullYear(),
      pad(now.getMonth() + 1),
      pad(now.getDate()),
      "_",
      pad(now.getHours()),
      pad(now.getMinutes()),
      pad(now.getSeconds()),
    ].join("");

    const safeName = String(tableName || "table")
      .replace(/[\\/:*?"<>|]+/g, "_")
      .replace(/\s+/g, "_")
      .slice(0, 40);

    return `${safeName || "table"}_export_${stamp}.xls`;
  }

  function sanitizeForExcel(value) {
    if (value === undefined || value === null) {
      return "";
    }
    return String(value).replace(/\s+/g, " ").trim();
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function showToast(message, isError = false) {
    let toast = document.getElementById(TOAST_ID);
    if (!toast) {
      toast = document.createElement("div");
      toast.id = TOAST_ID;
      toast.style.cssText = [
        "position:fixed",
        "right:20px",
        "bottom:112px",
        "z-index:2147483647",
        "padding:8px 12px",
        "border-radius:8px",
        "background:rgba(0,0,0,.78)",
        "color:#fff",
        "font-size:13px",
        "max-width:360px",
        "line-height:1.4",
        "word-break:break-all",
        "box-shadow:0 6px 12px rgba(0,0,0,.16)",
      ].join(";");
      document.documentElement.appendChild(toast);
    }

    toast.style.background = isError ? "#d94848" : "rgba(0,0,0,.78)";
    toast.textContent = message;
    toast.style.display = "block";

    window.clearTimeout(window.__tableToolsToastTimer);
    window.__tableToolsToastTimer = window.setTimeout(() => {
      toast.style.display = "none";
    }, 3200);
  }

  function isFeishuConfigMissingError(message) {
    return String(message || "").includes("请先在插件设置中填写");
  }

  async function tryOpenOptionsPage() {
    try {
      await chrome.runtime.sendMessage({ type: "OPEN_OPTIONS" });
    } catch (error) {
      // ignore
    }
  }
})();
