const syncBtn = document.getElementById("syncBtn");
const exportBtn = document.getElementById("exportBtn");
const settingsBtn = document.getElementById("settingsBtn");
const statusEl = document.getElementById("status");

syncBtn.addEventListener("click", async () => {
  await runAction({
    button: syncBtn,
    startText: "开始同步到飞书...",
    messageType: "TRIGGER_SYNC_TO_FEISHU",
    successText: (response) => {
      const duplicateSkipped = Number(response.duplicateSkipped || 0);
      const dedupeHint = duplicateSkipped > 0 ? `，重复跳过 ${duplicateSkipped} 条（已同步过）` : "";
      return `飞书同步完成：${response.synced || 0} 条${dedupeHint}`;
    },
    defaultError: "飞书同步失败",
  });
});

exportBtn.addEventListener("click", async () => {
  await runAction({
    button: exportBtn,
    startText: "开始导出Excel...",
    messageType: "TRIGGER_EXPORT",
    successText: (response) => {
      const fileName = response.fileName ? `，文件：${response.fileName}` : "";
      return `导出完成：${response.exported || 0} 条${fileName}`;
    },
    defaultError: "导出失败",
  });
});

settingsBtn.addEventListener("click", () => {
  chrome.runtime.openOptionsPage();
});

async function runAction({ button, startText, messageType, successText, defaultError }) {
  setStatus(startText);
  button.disabled = true;

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id) {
      throw new Error("找不到当前标签页");
    }

    const response = await chrome.tabs.sendMessage(tab.id, { type: messageType });
    if (!response?.ok) {
      const errorMessage = response?.error || defaultError;
      if (messageType === "TRIGGER_SYNC_TO_FEISHU" && isFeishuConfigMissingError(errorMessage)) {
        await chrome.runtime.openOptionsPage();
        throw new Error("飞书参数未配置，已打开设置页，请先填写 App ID / App Secret / App Token / Table ID。");
      }
      throw new Error(errorMessage);
    }

    setStatus(successText(response), false);
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? "#dc2626" : "#16a34a";
}

function isFeishuConfigMissingError(message) {
  return String(message || "").includes("请先在插件设置中填写");
}
