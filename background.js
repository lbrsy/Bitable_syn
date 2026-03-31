const FEISHU_API_BASE = "https://open.feishu.cn";
const CONFIG_KEY = "feishuConfig";
const TOKEN_CACHE_KEY = "feishuTokenCache";
const DOUYIN_MIN_MATCH_COUNT = 2;
const DOUYIN_DETECTION_FIELDS = ["姓名", "线索创建时间", "电话", "跟进员工", "最新跟进记录", "线索阶段", "意向线索"];
const DOUYIN_FIELD_MAPPINGS = [
  { targetName: "微信备注名称", sourceNames: ["姓名", "微信备注名称"] },
  { targetName: "线索时间", sourceNames: ["线索创建时间", "线索时间"] },
  { targetName: "来源", defaultValue: "付费流", alwaysUseDefault: true },
  { targetName: "手机号", sourceNames: ["电话", "手机号", "手机", "联系电话"] },
  { targetName: "销售", sourceNames: ["跟进员工", "销售"] },
  { targetName: "首次触达情况", sourceNames: ["最新跟进记录", "首次触达情况"] },
  { targetName: "跟进阶段", sourceNames: ["线索阶段", "跟进阶段"] },
];

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type === "SYNC_TABLE_DATA") {
    handleSync(message.payload)
      .then((result) => sendResponse({ ok: true, ...result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message?.type === "OPEN_OPTIONS") {
    chrome.runtime.openOptionsPage();
    sendResponse({ ok: true });
  }

  return false;
});

async function handleSync(payload) {
  if (!payload || !Array.isArray(payload.columns) || !Array.isArray(payload.rows)) {
    throw new Error("页面数据格式错误");
  }

  if (payload.rows.length === 0) {
    throw new Error("未检测到可同步的行数据");
  }

  const config = await getConfig();
  validateConfig(config);

  const tenantToken = await getTenantAccessToken(config);
  const target = await resolveBitableTarget(config, tenantToken);
  const runtimeConfig = {
    ...config,
    appToken: target.appToken,
    tableId: target.tableId,
  };

  const selectedColumns = pickColumns(payload.columns, runtimeConfig);
  if (selectedColumns.length === 0) {
    throw new Error("没有可同步的字段，请检查是否被过滤");
  }

  const metadataColumns = [];
  if (runtimeConfig.includeMetaFields !== false) {
    metadataColumns.push({ key: "__meta_source_url", name: "来源页面" });
    metadataColumns.push({ key: "__meta_capture_time", name: "抓取时间" });
  }

  const allColumns = [...selectedColumns, ...metadataColumns];

  const existingFieldNames = await getExistingFieldNames(runtimeConfig, tenantToken);
  if (runtimeConfig.autoCreateFields !== false) {
    await createMissingFields(runtimeConfig, tenantToken, allColumns, existingFieldNames);
  }

  const records = buildRecords(payload, selectedColumns, metadataColumns);
  if (records.length === 0) {
    throw new Error("所有行都为空，未写入飞书");
  }

  let recordsToCreate = records;
  let duplicateSkipped = 0;
  let dedupeFieldsUsed = [];

  if (runtimeConfig.enableDedup !== false) {
    const dedupeResult = await dedupeRecordsBeforeCreate(runtimeConfig, tenantToken, records);
    recordsToCreate = dedupeResult.recordsToCreate;
    duplicateSkipped = dedupeResult.duplicateSkipped;
    dedupeFieldsUsed = dedupeResult.dedupeFieldsUsed;
  }

  if (recordsToCreate.length === 0) {
    return {
      created: 0,
      duplicateSkipped,
      dedupeFieldsUsed,
      totalRows: payload.rows.length,
      tableName: payload.tableName || "",
    };
  }

  const created = await batchCreateRecords(runtimeConfig, tenantToken, recordsToCreate);

  return {
    created,
    duplicateSkipped,
    dedupeFieldsUsed,
    totalRows: payload.rows.length,
    tableName: payload.tableName || "",
  };
}

function pickColumns(columns, config) {
  const skipOperationColumn = config.skipOperationColumn !== false;

  const availableColumns = columns.filter((col) => {
    if (!col || !col.key || !col.name) {
      return false;
    }

    if (skipOperationColumn && /^(操作|action)$/i.test(col.name.trim())) {
      return false;
    }

    return true;
  });

  const mappedColumns = buildDouyinMappedColumns(availableColumns);
  if (mappedColumns.length > 0) {
    return mappedColumns;
  }

  return availableColumns;
}

function buildRecords(payload, selectedColumns, metadataColumns) {
  const records = [];

  for (const row of payload.rows) {
    const fields = {};
    let hasBusinessValue = false;

    for (const col of selectedColumns) {
      let value = normalizeCellValue(row[col.sourceKey || col.key]);
      if (col.alwaysUseDefault) {
        value = normalizeCellValue(col.defaultValue);
      } else if (value === "" && col.defaultValue !== undefined) {
        value = normalizeCellValue(col.defaultValue);
      }
      fields[col.name] = value;
      if (value !== "") {
        hasBusinessValue = true;
      }
    }

    if (!hasBusinessValue) {
      continue;
    }

    for (const col of metadataColumns) {
      if (col.key === "__meta_source_url") {
        fields[col.name] = payload.sourceUrl || "";
      }
      if (col.key === "__meta_capture_time") {
        fields[col.name] = payload.capturedAt || new Date().toISOString();
      }
    }

    records.push({ fields });
  }

  return records;
}

function buildDouyinMappedColumns(columns) {
  const matchedFieldCount = countDouyinSourceFieldMatches(columns);
  if (matchedFieldCount < DOUYIN_MIN_MATCH_COUNT) {
    return [];
  }

  return DOUYIN_FIELD_MAPPINGS.map((mapping, index) => {
    const sourceColumn = findColumnByNames(columns, mapping.sourceNames || []);
    if (!sourceColumn && mapping.defaultValue === undefined) {
      return null;
    }

    return {
      key: sourceColumn?.key || `__mapped_${index + 1}`,
      sourceKey: sourceColumn?.key || "",
      name: mapping.targetName,
      defaultValue: mapping.defaultValue,
      alwaysUseDefault: mapping.alwaysUseDefault === true,
    };
  }).filter(Boolean);
}

function countDouyinSourceFieldMatches(columns) {
  const normalizedColumnNames = new Set(
    columns
      .map((col) => normalizeColumnName(col?.name))
      .filter(Boolean)
  );

  return DOUYIN_DETECTION_FIELDS.reduce((count, fieldName) => {
    return count + (normalizedColumnNames.has(normalizeColumnName(fieldName)) ? 1 : 0);
  }, 0);
}

function findColumnByNames(columns, names) {
  const normalizedNames = new Set((names || []).map((name) => normalizeColumnName(name)).filter(Boolean));
  if (normalizedNames.size === 0) {
    return null;
  }

  for (const col of columns) {
    if (normalizedNames.has(normalizeColumnName(col?.name))) {
      return col;
    }
  }

  return null;
}

function normalizeColumnName(value) {
  return String(value || "")
    .replace(/[\s:：]+/g, "")
    .trim()
    .toLowerCase();
}

async function dedupeRecordsBeforeCreate(config, token, records) {
  const dedupeFieldsUsed = resolveDedupeFields(config, records);

  const uniqueInBatch = [];
  const seenInBatch = new Set();
  let duplicateInBatch = 0;

  for (const record of records) {
    const key = buildRecordDedupeKey(record.fields, dedupeFieldsUsed);
    if (!key) {
      uniqueInBatch.push(record);
      continue;
    }

    if (seenInBatch.has(key)) {
      duplicateInBatch += 1;
      continue;
    }

    seenInBatch.add(key);
    uniqueInBatch.push(record);
  }

  if (uniqueInBatch.length === 0) {
    return {
      recordsToCreate: [],
      duplicateSkipped: duplicateInBatch,
      dedupeFieldsUsed,
    };
  }

  const existingKeySet = await getExistingRecordDedupeKeys(config, token, dedupeFieldsUsed);
  const recordsToCreate = [];
  let duplicateInFeishu = 0;

  for (const record of uniqueInBatch) {
    const key = buildRecordDedupeKey(record.fields, dedupeFieldsUsed);
    if (!key) {
      recordsToCreate.push(record);
      continue;
    }

    if (existingKeySet.has(key)) {
      duplicateInFeishu += 1;
      continue;
    }

    existingKeySet.add(key);
    recordsToCreate.push(record);
  }

  return {
    recordsToCreate,
    duplicateSkipped: duplicateInBatch + duplicateInFeishu,
    dedupeFieldsUsed,
  };
}

async function getExistingRecordDedupeKeys(config, token, dedupeFieldsUsed) {
  const keySet = new Set();
  let pageToken = "";
  let scanned = 0;
  const pageSize = 500;
  const maxScan = 20000;

  while (true) {
    const query = pageToken
      ? `?page_size=${pageSize}&page_token=${encodeURIComponent(pageToken)}`
      : `?page_size=${pageSize}`;

    const response = await feishuRequest({
      path: `/open-apis/bitable/v1/apps/${encodeURIComponent(config.appToken)}/tables/${encodeURIComponent(config.tableId)}/records${query}`,
      method: "GET",
      token,
    });

    const items = Array.isArray(response?.items) ? response.items : [];
    for (const item of items) {
      const key = buildRecordDedupeKey(item?.fields || {}, dedupeFieldsUsed);
      if (key) {
        keySet.add(key);
      }
    }

    scanned += items.length;
    if (scanned >= maxScan) {
      break;
    }

    if (!response?.has_more || !response?.page_token) {
      break;
    }

    pageToken = response.page_token;
  }

  return keySet;
}

function resolveDedupeFields(config, records) {
  const availableFieldSet = new Set();
  for (const record of records) {
    Object.keys(record?.fields || {}).forEach((field) => availableFieldSet.add(field));
  }

  const configured = parseDedupeFields(config.dedupeFields).filter((field) => availableFieldSet.has(field));
  if (configured.length > 0) {
    return configured;
  }

  const defaultCandidates = [
    "电话",
    "手机号",
    "手机",
    "联系电话",
    "订单编号",
    "商品ID",
  ];
  const auto = defaultCandidates.filter((field) => availableFieldSet.has(field));
  if (auto.length > 0) {
    return auto;
  }

  return [];
}

function buildRecordDedupeKey(fields, dedupeFieldsUsed) {
  const values = fields || {};

  if (Array.isArray(dedupeFieldsUsed) && dedupeFieldsUsed.length > 0) {
    const parts = dedupeFieldsUsed.map((field) => normalizeDedupeValue(values[field]));
    if (parts.every((part) => part === "")) {
      return "";
    }
    return `k:${parts.join("||")}`;
  }

  const keys = Object.keys(values)
    .filter((key) => key !== "抓取时间")
    .sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));

  if (keys.length === 0) {
    return "";
  }

  const payload = keys.map((key) => `${key}=${normalizeDedupeValue(values[key])}`).join("||");
  return `r:${payload}`;
}

function normalizeDedupeValue(value) {
  if (value === undefined || value === null) {
    return "";
  }
  if (typeof value === "string") {
    return value.replace(/\s+/g, " ").trim();
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (Array.isArray(value)) {
    return value.map((item) => normalizeDedupeValue(item)).join(",");
  }
  if (typeof value === "object") {
    const sortedKeys = Object.keys(value).sort((a, b) => a.localeCompare(b));
    return sortedKeys.map((key) => `${key}:${normalizeDedupeValue(value[key])}`).join("|");
  }
  return String(value).trim();
}

function parseDedupeFields(value) {
  const text = String(value || "").trim();
  if (!text) {
    return [];
  }

  const parts = text
    .split(/[,，]/)
    .map((item) => item.trim())
    .filter(Boolean);

  const seen = new Set();
  const deduped = [];
  for (const item of parts) {
    if (seen.has(item)) {
      continue;
    }
    seen.add(item);
    deduped.push(item);
  }
  return deduped;
}

async function batchCreateRecords(config, token, records) {
  const chunkSize = 1000;
  const ignoreConsistencyCheck = config.ignoreConsistencyCheck !== false;
  const userIdType = config.userIdType || "open_id";
  let createdCount = 0;

  for (let i = 0; i < records.length; i += chunkSize) {
    const chunk = records.slice(i, i + chunkSize);
    const query = new URLSearchParams({
      client_token: generateUuidV4(),
      ignore_consistency_check: String(ignoreConsistencyCheck),
      user_id_type: userIdType,
    }).toString();

    const response = await feishuRequest({
      path: `/open-apis/bitable/v1/apps/${encodeURIComponent(config.appToken)}/tables/${encodeURIComponent(config.tableId)}/records/batch_create?${query}`,
      method: "POST",
      token,
      body: { records: chunk },
    });

    const items = response?.records || [];
    createdCount += items.length || chunk.length;
  }

  return createdCount;
}

async function getExistingFieldNames(config, token) {
  const names = new Set();
  let pageToken = "";

  while (true) {
    const query = pageToken
      ? `?page_size=200&page_token=${encodeURIComponent(pageToken)}`
      : "?page_size=200";

    const response = await feishuRequest({
      path: `/open-apis/bitable/v1/apps/${encodeURIComponent(config.appToken)}/tables/${encodeURIComponent(config.tableId)}/fields${query}`,
      method: "GET",
      token,
    });

    const items = Array.isArray(response?.items) ? response.items : [];
    for (const item of items) {
      if (item?.field_name) {
        names.add(item.field_name);
      }
    }

    if (!response?.has_more || !response?.page_token) {
      break;
    }

    pageToken = response.page_token;
  }

  return names;
}

async function createMissingFields(config, token, columns, existingFieldNames) {
  for (const col of columns) {
    if (existingFieldNames.has(col.name)) {
      continue;
    }

    const query = new URLSearchParams({
      client_token: generateUuidV4(),
    }).toString();

    try {
      await feishuRequest({
        path: `/open-apis/bitable/v1/apps/${encodeURIComponent(config.appToken)}/tables/${encodeURIComponent(config.tableId)}/fields?${query}`,
        method: "POST",
        token,
        body: {
          field_name: col.name,
          type: 1,
        },
      });
    } catch (error) {
      if (isFieldNameDuplicatedError(error?.message)) {
        existingFieldNames.add(col.name);
        await delay(120);
        continue;
      }
      throw error;
    }

    existingFieldNames.add(col.name);
    await delay(120);
  }
}

async function getTenantAccessToken(config) {
  const now = Date.now();
  const local = await chrome.storage.local.get(TOKEN_CACHE_KEY);
  const cache = local[TOKEN_CACHE_KEY];

  if (cache?.token && cache?.expireAt && cache.expireAt > now + 60 * 1000) {
    return cache.token;
  }

  const response = await feishuRequest({
    path: "/open-apis/auth/v3/tenant_access_token/internal",
    method: "POST",
    body: {
      app_id: config.appId,
      app_secret: config.appSecret,
    },
  });

  const token = response?.tenant_access_token;
  const expire = Number(response?.expire) || 7200;

  if (!token) {
    throw new Error("获取 tenant_access_token 失败");
  }

  await chrome.storage.local.set({
    [TOKEN_CACHE_KEY]: {
      token,
      expireAt: now + expire * 1000,
    },
  });

  return token;
}

async function feishuRequest({ path, method, token, body }) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
  };

  if (token) {
    headers.Authorization = `Bearer ${token}`;
  }

  const response = await fetch(`${FEISHU_API_BASE}${path}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  let json;
  try {
    json = await response.json();
  } catch (error) {
    throw new Error(`飞书接口响应解析失败: ${error.message}`);
  }

  if (!response.ok) {
    const detail = json?.msg || response.statusText;
    if (json?.code && json.code !== 0) {
      throw new Error(buildFeishuCodeError(json.code, json.msg));
    }
    if (response.status === 403) {
      const hint = [
        "可能原因：",
        "1) 飞书应用没有发布包含多维表格权限的最新版本",
        "2) 目标多维表格未把该应用加入协作者（可编辑）",
        "3) App Token 填错（不要填 wiki 链接里的文档 token）",
        "4) 应用与目标表格不在同一租户",
      ].join(" ");
      throw new Error(`飞书接口请求失败(403): ${detail}。${hint}`);
    }
    throw new Error(`飞书接口请求失败(${response.status}): ${detail}`);
  }

  if (json?.code !== 0) {
    throw new Error(buildFeishuCodeError(json.code, json.msg));
  }

  return json?.data || json;
}

async function getConfig() {
  const local = await chrome.storage.local.get(CONFIG_KEY);
  return local[CONFIG_KEY] || {};
}

function validateConfig(config) {
  if (!config.appId) {
    throw new Error("请先在插件设置中填写 App ID");
  }
  if (!config.appSecret) {
    throw new Error("请先在插件设置中填写 App Secret");
  }
  if (!config.appToken) {
    throw new Error("请先在插件设置中填写 App Token");
  }
  if (!config.tableId) {
    throw new Error("请先在插件设置中填写 Table ID");
  }
  if (!isValidAppTokenInput(config.appToken.trim())) {
    throw new Error("App Token 格式不正确：支持 base token、wikxxxx，或完整 wiki/base 链接");
  }
  if (!isValidTableIdInput(config.tableId.trim())) {
    throw new Error("Table ID 格式不正确：支持 tblxxxx，或带 table= 参数的完整链接");
  }
}

async function resolveBitableTarget(config, token) {
  const appToken = await normalizeAppToken(config.appToken, token);
  const tableId = normalizeTableId(config.tableId);

  if (!isLikelyBitableAppToken(appToken)) {
    throw new Error("无法解析有效 App Token。请填写 base token，或可解析为多维表格的 wiki/base 链接");
  }

  if (!/^tbl[a-zA-Z0-9]+$/.test(tableId)) {
    throw new Error("无法解析有效 Table ID。请填写 tblxxxx，或提供含 table= 的链接");
  }

  return { appToken, tableId };
}

async function normalizeAppToken(rawValue, token) {
  const value = String(rawValue || "").trim();

  if (/^wik[a-zA-Z0-9]+$/.test(value)) {
    return await resolveAppTokenByWikiToken(value, token);
  }

  if (isLikelyBitableAppToken(value)) {
    return value;
  }

  if (isHttpUrl(value)) {
    const parsed = parseBitableInfoFromUrl(value);
    if (parsed.appToken) {
      return parsed.appToken;
    }
    if (parsed.wikiToken) {
      return await resolveAppTokenByWikiToken(parsed.wikiToken, token);
    }
  }

  return value;
}

function normalizeTableId(rawValue) {
  const value = String(rawValue || "").trim();
  if (/^tbl[a-zA-Z0-9]+$/.test(value)) {
    return value;
  }

  if (isHttpUrl(value)) {
    const parsed = parseBitableInfoFromUrl(value);
    if (parsed.tableId) {
      return parsed.tableId;
    }
  }

  return value;
}

async function resolveAppTokenByWikiToken(wikiToken, token) {
  const response = await feishuRequest({
    path: `/open-apis/wiki/v2/spaces/get_node?token=${encodeURIComponent(wikiToken)}`,
    method: "GET",
    token,
  });

  const node = response?.node || {};
  if (node.obj_type !== "bitable") {
    throw new Error(`wiki 节点类型为 ${node.obj_type || "未知"}，不是多维表格节点`);
  }

  const appToken = String(node.obj_token || "");
  if (!isLikelyBitableAppToken(appToken)) {
    throw new Error("从 wiki 节点解析 app_token 失败，请手动填写 base token");
  }

  return appToken;
}

function isValidAppTokenInput(value) {
  if (!value) {
    return false;
  }
  if (/^wik[a-zA-Z0-9]+$/.test(value)) {
    return true;
  }
  if (isLikelyBitableAppToken(value)) {
    return true;
  }
  if (isHttpUrl(value)) {
    const parsed = parseBitableInfoFromUrl(value);
    return Boolean(parsed.appToken || parsed.wikiToken);
  }
  return false;
}

function isValidTableIdInput(value) {
  if (!value) {
    return false;
  }
  if (/^tbl[a-zA-Z0-9]+$/.test(value)) {
    return true;
  }
  if (isHttpUrl(value)) {
    const parsed = parseBitableInfoFromUrl(value);
    return Boolean(parsed.tableId);
  }
  return false;
}

function isHttpUrl(value) {
  return /^https?:\/\//i.test(value || "");
}

function parseBitableInfoFromUrl(urlText) {
  try {
    const url = new URL(urlText);
    const tableId = String(url.searchParams.get("table") || "");
    const appMatch = url.pathname.match(/\/base\/([a-zA-Z0-9]+)/);
    const wikiMatch = url.pathname.match(/\/wiki\/([a-zA-Z0-9]+)/);

    return {
      appToken: appMatch ? appMatch[1] : "",
      tableId,
      wikiToken: wikiMatch ? wikiMatch[1] : "",
    };
  } catch (error) {
    return {
      appToken: "",
      tableId: "",
      wikiToken: "",
    };
  }
}

function buildFeishuCodeError(code, msg) {
  const errorMsg = msg || "未知错误";
  const commonPrefix = `飞书接口错误(${code}): ${errorMsg}`;

  if (code === 1254003 || code === 1254040) {
    return `${commonPrefix}。app_token 错误：请使用 base 链接 /base/ 后的 token，wiki 链接需先解析为 obj_token。`;
  }
  if (code === 1254004 || code === 1254041) {
    return `${commonPrefix}。table_id 错误：请确认 table= 后的 tblxxxx 是否正确。`;
  }
  if (code === 1254302 || code === 1254304) {
    return `${commonPrefix}。权限不足：请在多维表格高级权限中给应用“可管理/可编辑”权限，并确认应用权限已发布。`;
  }
  if (code === 1254301) {
    return `${commonPrefix}。多维表格未开启高级权限或当前表格不支持该操作。`;
  }
  if (code === 1254291) {
    return `${commonPrefix}。写入冲突：请降低并发、稍后重试。`;
  }
  if (code === 1254014) {
    return `${commonPrefix}。字段名重复（FieldNameDuplicated）：通常可忽略，或请更换字段名称。`;
  }
  if (code === 1254028 || code === 1254029) {
    return `${commonPrefix}。字段名无效：请检查字段名是否为空、包含非法字符或超长。`;
  }
  if (code === 1254015) {
    return `${commonPrefix}。字段类型和值不匹配，请检查写入字段的数据格式。`;
  }
  if (code >= 1254080 && code <= 1254096) {
    return `${commonPrefix}。字段 property 配置错误，请按字段类型检查 property 结构。`;
  }
  if (code === 1254608) {
    return `${commonPrefix}。重复提交（ReqRecommited）：请避免重复 client_token 或过快重复请求。`;
  }
  if (code === 1254607) {
    return `${commonPrefix}。数据未就绪（Data not ready），建议稍后重试。`;
  }

  return commonPrefix;
}

function generateUuidV4() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (char) => {
    const random = Math.floor(Math.random() * 16);
    const value = char === "x" ? random : (random & 0x3) | 0x8;
    return value.toString(16);
  });
}

function isLikelyBitableAppToken(value) {
  const token = String(value || "").trim();
  if (!/^[a-zA-Z0-9]{8,}$/.test(token)) {
    return false;
  }
  if (/^tbl[a-zA-Z0-9]+$/i.test(token)) {
    return false;
  }
  if (/^vew[a-zA-Z0-9]+$/i.test(token)) {
    return false;
  }
  return true;
}

function normalizeCellValue(value) {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value).replace(/\s+/g, " ").trim();
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function isFieldNameDuplicatedError(message) {
  const text = String(message || "");
  return text.includes("1254014") || text.includes("FieldNameDuplicated") || text.includes("字段名重复");
}
