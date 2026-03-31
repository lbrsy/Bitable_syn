const CONFIG_KEY = "feishuConfig";

const elements = {
  tableUrl: document.getElementById("tableUrl"),
  robotSecret: document.getElementById("robotSecret"),
  appId: document.getElementById("appId"),
  userIdType: document.getElementById("userIdType"),
  autoCreateFields: document.getElementById("autoCreateFields"),
  includeMetaFields: document.getElementById("includeMetaFields"),
  skipOperationColumn: document.getElementById("skipOperationColumn"),
  ignoreConsistencyCheck: document.getElementById("ignoreConsistencyCheck"),
  enableDedup: document.getElementById("enableDedup"),
  dedupeFields: document.getElementById("dedupeFields"),
  previewAppToken: document.getElementById("previewAppToken"),
  previewTableId: document.getElementById("previewTableId"),
  saveBtn: document.getElementById("saveBtn"),
  openFeishuBtn: document.getElementById("openFeishuBtn"),
  status: document.getElementById("status"),
};

init().catch((error) => {
  setStatus(error.message, true);
});

elements.tableUrl.addEventListener("input", () => {
  updatePreview(elements.tableUrl.value.trim());
});

elements.saveBtn.addEventListener("click", async () => {
  const local = await chrome.storage.local.get(CONFIG_KEY);
  const existing = local[CONFIG_KEY] || {};

  const tableUrl = elements.tableUrl.value.trim();
  const robotSecret = elements.robotSecret.value.trim();
  const appIdInput = elements.appId.value.trim();

  if (!tableUrl || !robotSecret) {
    setStatus("请至少填写：飞书表格完整链接 + 应用密钥(App Secret)", true);
    return;
  }

  if (/^sec[-_a-zA-Z0-9]+$/i.test(robotSecret)) {
    setStatus("你填写的像是机器人 Webhook 密钥（SEC...），请改填开放平台应用的 App Secret", true);
    return;
  }

  const parsed = parseBitableInfoFromUrl(tableUrl);
  if (!parsed.tableId || (!parsed.appToken && !parsed.wikiToken)) {
    setStatus("链接无法解析出有效 App Token / Table ID，请检查链接是否完整", true);
    return;
  }

  const appId = appIdInput || existing.appId || "";
  if (!appId) {
    setStatus("首次请在高级设置填写 App ID（只需一次）", true);
    return;
  }

  const config = {
    appId,
    appSecret: robotSecret,
    appToken: parsed.appToken || parsed.wikiToken,
    tableId: parsed.tableId,
    tableUrl,
    userIdType: elements.userIdType.value || "open_id",
    autoCreateFields: elements.autoCreateFields.checked,
    includeMetaFields: elements.includeMetaFields.checked,
    skipOperationColumn: elements.skipOperationColumn.checked,
    ignoreConsistencyCheck: elements.ignoreConsistencyCheck.checked,
    enableDedup: elements.enableDedup.checked,
    dedupeFields: elements.dedupeFields.value.trim(),
  };

  await chrome.storage.local.set({
    [CONFIG_KEY]: config,
  });

  setStatus("保存成功", false);
});

elements.openFeishuBtn.addEventListener("click", () => {
  chrome.tabs.create({ url: "https://open.feishu.cn" });
});

async function init() {
  const local = await chrome.storage.local.get(CONFIG_KEY);
  const config = local[CONFIG_KEY] || {};

  elements.tableUrl.value = config.tableUrl || "";
  elements.robotSecret.value = config.appSecret || "";
  elements.appId.value = config.appId || "";
  elements.userIdType.value = config.userIdType || "open_id";

  elements.autoCreateFields.checked = config.autoCreateFields !== false;
  elements.includeMetaFields.checked = config.includeMetaFields !== false;
  elements.skipOperationColumn.checked = config.skipOperationColumn !== false;
  elements.ignoreConsistencyCheck.checked = config.ignoreConsistencyCheck !== false;
  elements.enableDedup.checked = config.enableDedup !== false;
  elements.dedupeFields.value = config.dedupeFields || "";

  updatePreview(elements.tableUrl.value.trim());
}

function parseBitableInfoFromUrl(urlText) {
  try {
    const url = new URL(urlText);
    const tableId = String(url.searchParams.get("table") || "");
    const appMatch = url.pathname.match(/\/base\/([a-zA-Z0-9]+)/);
    const wikiMatch = url.pathname.match(/\/wiki\/([a-zA-Z0-9]+)/);

    return {
      appToken: appMatch ? appMatch[1] : "",
      wikiToken: wikiMatch ? wikiMatch[1] : "",
      tableId,
    };
  } catch (error) {
    return {
      appToken: "",
      wikiToken: "",
      tableId: "",
    };
  }
}

function updatePreview(urlText) {
  if (!urlText) {
    elements.previewAppToken.textContent = "-";
    elements.previewTableId.textContent = "-";
    return;
  }

  const parsed = parseBitableInfoFromUrl(urlText);
  elements.previewAppToken.textContent = parsed.appToken || parsed.wikiToken || "-";
  elements.previewTableId.textContent = parsed.tableId || "-";
}

function setStatus(message, isError = false) {
  elements.status.textContent = message;
  elements.status.style.color = isError ? "#d94848" : "#16a34a";
}
