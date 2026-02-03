/****************************************************
 * AUDITORIA NOTURNA — Omnibees -> Auditoria
 * Arquitetura:
 *  - OMNI_RAW   : snapshot bruto (última importação) + metadados
 *  - OMNI_NORM  : dataset normalizado (última importação)
 *  - AUDIT_<dt> : auditoria do dia (sempre nova aba)
 *  - DASHBOARD  : KPIs, quebras, links, status geral
 *  - AUDIT_LOG  : trilha de execução/importação
 *
 * Requisito:
 *  - Ativar Drive API Avançada:
 *    Apps Script -> Serviços avançados do Google -> Drive API
 ****************************************************/

/** =======================
 *  Configuração
 *  ======================= */
const UX = {
  trashTempFiles: true,
  maxBlocksSoftLimit: 250,
  timezone: Session.getScriptTimeZone() || "America/Sao_Paulo",

  omniRawSheetName: "OMNI_RAW",
  omniNormSheetName: "OMNI_NORM",
  dashboardSheetName: "DASHBOARD",
  auditLogSheetName: "AUDIT_LOG",

  auditSheetPrefix: "", // se quiser: "AUDIT "
  auditSheetDateFormatPrimary: "dd/MM/yyyy",
  auditSheetDateFormatFallback: "dd-MM-yyyy",
  auditSource: "RAW", // "RAW" (linha a linha) | "NORM" (agrupado por reserva)
  enableOmniRaw: false, // descontinuado
  auditTemplateSheetName: "AUDIT_TEMPLATE",
  importDedupeMinutes: 30,
  importDedupeMaxEntries: 50,
  niaraBaseUrl: "https://pousadaportofino.niara.tech/services/",
  pmsBaseUrl: "https://pousadaareiadoforte.hflow.com.br/ultraportofino/pms/reservations/reservation/",
  bee2payBaseUrl: "",
  bee2payReservationBaseUrl: "https://app.bee2pay.com/HOTEL_OMNI_19525/hotelReservations?locator=",

  // Robustez / limites
  maxSidebarUploadBytes: 8 * 1024 * 1024, // upload único (não-chunk) - recomendado
  rawMaxRows: 8000,                       // evita sheet explodir + timeouts
  rawMaxCols: 40,
  openRetryCount: 6,
  openRetrySleepMs: 450,                  // backoff simples
};

const UPLOAD = {
  // Upload em chunks (para upload local robusto)
  chunkSizeChars: 200 * 1024,       // legado (base64 por caracteres)
  chunkSizeBytes: 256 * 1024,       // resumable (bytes reais)
  cacheTtlSeconds: 60 * 15,
};

const LISTS = {
  auditors: ["Lucas", "Michael"],
  validators: ["Michael", "Karol", "Magno", "Yasmin"],
  auditStatus: ["Validado", "Não Validado"],
  origins: ["Central de Reservas", "BE Mobile", "Booking Engine", "Booking", "Iterpec"],
  payments: ["Sem Pagamento", "Pago", "Não cobrada", "Central de Reservas", "Parcial"],
  paymentsDropdown: ["Pago", "Não Pago"],
  resStatus: ["Confirmado", "Cancelada", "Alterada"],
  resStatusDropdown: ["Confirmado", "Alterada", "Cancelada"],
  checks: ["✅", "❌"],
};

const THEME = {
  bg: "#FFFFFF",
  headerBand: "#F2F2F2",
  tableHeader: "#B7B7B7",
  cardBg: "#FFFFFF",
  checkRow: "#FFFFFF",
  omniRow: "#F3F3F3",
  pmsRow: "#D9D9D9",
  indexBg: "#E6E6E6",
  obsCol: "#C9DAF8",
  border: "#C8C8C8",
  text: "#1C1C1E",
  subtext: "#6E6E73",
  tableText: "#0000FF",
  statusTextOk: "#0F5132",
  statusTextBad: "#842029",
  statusTextWarn: "#7A4B00",
  ok: "#E9F7EC",
  bad: "#FDEAEA",
  chipOk: "#DFF5E1",
  chipBad: "#FAD9D9",
  chipWarn: "#FFE5C2",
  accent: "#0A58CA",
  warn: "#B54708",
};

const LAYOUT = {
  cols: 13,              // A..M
  startRow: 7,           // linha inicial dos dados
  blockHeight: 5,        // Header + Omni + PMS + Checks + Spacer
  freezeRows: 5,         // topo
};

const SHEET_NAME_MAX = 100;

/** =======================
 *  Menu
 *  ======================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Auditoria")
    .addItem("Importar Omnibees (do computador)", "showUploadSidebar")
    .addItem("Importar Niara (do computador)", "showNiaraUploadSidebar")
    .addItem("Importar Bee2Pay (do computador)", "showBee2PayUploadSidebar")
    .addItem("Importar Omnibees (por ID/URL do arquivo)", "importOmnibeesByFileIdPrompt")
    .addItem("Importar Omnibees (último arquivo de uma pasta)", "importLatestOmnibeesFromFolderPrompt")
    .addSeparator()
    .addItem("Validar ambiente (Drive API)", "validateEnvironment")
    .addToUi();
}

function onEdit(e) {
  try {
    const range = e && e.range;
    if (!range) return;
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
    const col = range.getColumn();
    const row = range.getRow();
    if (col !== 12 || row < LAYOUT.startRow) return;
    const v = String(range.getValue() || "").trim();
    if (!v) return;
    const m = v.match(/^[🟢🟡🔴]\s*(.+)$/);
    if (m && m[1]) range.setValue(m[1].trim());
  } catch (_) {}
}

function validateEnvironment() {
  assertDriveAdvancedServiceEnabled_();
  SpreadsheetApp.getUi().alert("Ambiente OK. Drive API Avançada disponível.");
}

/** =======================
 *  Sidebar (Upload local)
 *  ======================= */
function showUploadSidebar() {
  showUploadSidebarWithSource_("OMNI");
}

function showNiaraUploadSidebar() {
  showUploadSidebarWithSource_("NIARA");
}

function showBee2PayUploadSidebar() {
  showUploadSidebarWithSource_("BEE2PAY");
}

function showUploadSidebarWithSource_(sourceType) {
  const tpl = HtmlService.createTemplateFromFile("Upload");
  tpl.sourcePreset = String(sourceType || "OMNI").toUpperCase();
  const html = tpl.evaluate().setTitle(tpl.sourcePreset === "NIARA" ? "Importar Niara" : "Importar Omnibees");
  SpreadsheetApp.getUi().showSidebar(html);
}

/** =======================
 *  Config (auditores)
 *  ======================= */
function ensureConfigSheet_(ss) {
  const name = "CONFIG";
  const sh = upsertSheet_(ss, name);
  sh.setHiddenGridlines(true);
  ensureSheetSize_(sh, 20, 2);

  const header = sh.getRange(1, 1, 1, 2).getValues()[0];
  if (!String(header[0] || "").trim()) {
    sh.getRange(1, 1, 1, 2).setValues([["Auditores", "Observações"]])
      .setFontWeight("bold")
      .setBackground(THEME.tableHeader)
      .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 240);
    sh.setColumnWidth(2, 320);
  }

  const last = sh.getLastRow();
  let hasAuditors = false;
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, 1).getValues();
    hasAuditors = vals.some(r => String(r[0] || "").trim());
  }

  if (!hasAuditors && LISTS.auditors && LISTS.auditors.length) {
    sh.getRange(2, 1, LISTS.auditors.length, 1).setValues(LISTS.auditors.map(a => [a]));
  }

  return sh;
}

function getConfigAuditors_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ensureConfigSheet_(ss);
  const last = sh.getLastRow();
  if (last < 2) return (LISTS.auditors || []).slice();

  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  const list = unique_(vals.map(r => String(r[0] || "").trim())).filter(Boolean);
  return list.length ? list : (LISTS.auditors || []).slice();
}

function getSidebarUploadLimits_() {
  return {
    maxBytes: UX.maxSidebarUploadBytes,
    maxMB: Math.round((UX.maxSidebarUploadBytes / (1024 * 1024)) * 10) / 10,
    chunkSizeChars: UPLOAD.chunkSizeChars,
    chunkSizeBytes: UPLOAD.chunkSizeBytes
  };
}

/**
 * Endpoint único para a sidebar (Upload.html).
 * Roteia operações para manter a UI simples e evitar erros de função inexistente.
 */
function uploadApi(op, payload) {
  const p = payload || {};
  const sourceType = String(p.sourceType || "OMNI").toUpperCase();
  const isNiara = sourceType === "NIARA";
  const isBee2Pay = sourceType === "BEE2PAY";
  switch (String(op || "").toLowerCase()) {
    case "limits":
      return getSidebarUploadLimits_();
    case "auditors":
      return getConfigAuditors_();
    case "simple":
      return isNiara
        ? uploadNiaraFileAndUpdate(p.fileName, p.mimeType, p.base64Data)
        : (isBee2Pay
          ? uploadBee2PayFileAndUpdate(p.fileName, p.mimeType, p.base64Data)
          : uploadOmniFileAndGenerateByReportDate(p.fileName, p.mimeType, p.base64Data, p.auditorName));
    case "rstart":
      return resumableStart_(p.fileName, p.mimeType, p.totalSizeBytes, sourceType);
    case "rchunk":
      return resumableChunk_(p.token, p.offset, p.chunkBase64, p.totalSizeBytes, p.uploadUrl, p.fileName, p.auditorName, sourceType);
    case "rfinish":
      return resumableFinish_(p.token);
    case "start":
      return uploadStart_(p.fileName, p.mimeType, p.totalSizeBytes, p.totalChunks, p.auditorName, sourceType);
    case "chunk":
      return uploadChunk_(p.token, p.index, p.chunkBase64);
    case "finish":
      return uploadFinish_(p.token);
    default:
      throw new Error("Operação inválida em uploadApi: " + op);
  }
}

/**
 * Upload "simples" (não-chunk). Mantido para arquivos pequenos.
 */
function uploadOmniFileAndGenerateByReportDate(fileName, mimeType, base64Data, auditorName) {
  return withImportLock_(() => {
    assertDriveAdvancedServiceEnabled_();

    if (!base64Data || String(base64Data).length < 100) {
      throw new Error("Upload inválido: arquivo vazio ou não lido corretamente.");
    }

    // estimativa conservadora de bytes do base64
    const approxBytes = Math.floor((String(base64Data).length * 3) / 4);
    if (approxBytes > UX.maxSidebarUploadBytes) {
      throw new Error(
        `Arquivo muito grande para upload direto (~${(approxBytes / (1024 * 1024)).toFixed(1)}MB).\n` +
        `Use upload em chunks (recomendado) ou importe via Drive (ID/URL ou pasta).`
      );
    }

    mimeType = normalizeExcelMime_(fileName, mimeType);

    const ss = SpreadsheetApp.getActive();
    const folder = resolveUploadFolder_(ss);

    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const uploaded = folder.createFile(blob).setName(`[UPLOAD] ${fileName}`);

    try {
      const msg = pipelineFromExcelFileId_(uploaded.getId(), uploaded.getName(), { source: "SIDEBAR_SIMPLE", auditorName });
      return msg || "Importação concluída. Auditoria atualizada.";
    } finally {
      if (UX.trashTempFiles) safeTrashFile_(uploaded.getId());
    }
  });
}

/**
 * Upload Niara (não cria aba; atualiza aba de auditoria existente)
 */
function uploadNiaraFileAndUpdate(fileName, mimeType, base64Data) {
  return withImportLock_(() => {
    assertDriveAdvancedServiceEnabled_();

    if (!base64Data || String(base64Data).length < 100) {
      throw new Error("Upload inválido: arquivo vazio ou não lido corretamente.");
    }

    const approxBytes = Math.floor((String(base64Data).length * 3) / 4);
    if (approxBytes > UX.maxSidebarUploadBytes) {
      throw new Error(
        `Arquivo muito grande para upload direto (~${(approxBytes / (1024 * 1024)).toFixed(1)}MB).\n` +
        `Use upload em chunks (recomendado) ou importe via Drive (ID/URL ou pasta).`
      );
    }

    mimeType = normalizeExcelMime_(fileName, mimeType);

    const ss = SpreadsheetApp.getActive();
    const folder = resolveUploadFolder_(ss);

    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const uploaded = folder.createFile(blob).setName(`[UPLOAD] ${fileName}`);

    try {
      const msg = pipelineFromNiaraFileId_(uploaded.getId(), uploaded.getName(), { source: "SIDEBAR_SIMPLE_NIARA" });
      return msg || "Importação Niara concluída.";
    } finally {
      if (UX.trashTempFiles) safeTrashFile_(uploaded.getId());
    }
  });
}

/**
 * Upload Bee2Pay (não cria aba; atualiza aba de auditoria existente)
 */
function uploadBee2PayFileAndUpdate(fileName, mimeType, base64Data) {
  return withImportLock_(() => {
    assertDriveAdvancedServiceEnabled_();

    if (!base64Data || String(base64Data).length < 100) {
      throw new Error("Upload inválido: arquivo vazio ou não lido corretamente.");
    }

    const approxBytes = Math.floor((String(base64Data).length * 3) / 4);
    if (approxBytes > UX.maxSidebarUploadBytes) {
      throw new Error(
        `Arquivo muito grande para upload direto (~${(approxBytes / (1024 * 1024)).toFixed(1)}MB).\n` +
        `Use upload em chunks (recomendado) ou importe via Drive (ID/URL ou pasta).`
      );
    }

    mimeType = normalizeExcelMime_(fileName, mimeType);

    const ss = SpreadsheetApp.getActive();
    const folder = resolveUploadFolder_(ss);

    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const uploaded = folder.createFile(blob).setName(`[UPLOAD] ${fileName}`);

    try {
      const msg = pipelineFromBee2PayFileId_(uploaded.getId(), uploaded.getName(), { source: "SIDEBAR_SIMPLE_BEE2PAY" });
      return msg || "Importação Bee2Pay concluída.";
    } finally {
      if (UX.trashTempFiles) safeTrashFile_(uploaded.getId());
    }
  });
}

/** =======================
 *  Upload robusto em chunks
 *  ======================= */
function uploadStart_(fileName, mimeType, totalSizeBytes, totalChunks, auditorName, sourceType) {
  assertDriveAdvancedServiceEnabled_();
  const token = Utilities.getUuid();
  const meta = {
    fileName: String(fileName || ""),
    mimeType: String(mimeType || ""),
    totalSizeBytes: Number(totalSizeBytes) || 0,
    totalChunks: Number(totalChunks) || 0,
    auditorName: String(auditorName || "").trim(),
    sourceType: String(sourceType || "OMNI").toUpperCase(),
    received: 0,
    createdAt: Date.now()
  };
  CacheService.getScriptCache().put(token + ":meta", JSON.stringify(meta), UPLOAD.cacheTtlSeconds);
  return token;
}

function uploadChunk_(token, index, chunkBase64) {
  const cache = CacheService.getScriptCache();
  const metaRaw = cache.get(token + ":meta");
  if (!metaRaw) throw new Error("Sessão de upload expirada. Reabra a sidebar e tente novamente.");

  const i = Number(index);
  if (!isFinite(i) || i < 0) throw new Error("Chunk index inválido.");

  cache.put(token + ":chunk:" + i, String(chunkBase64 || ""), UPLOAD.cacheTtlSeconds);

  const meta = JSON.parse(metaRaw);
  meta.received = Math.max(meta.received, i + 1);
  cache.put(token + ":meta", JSON.stringify(meta), UPLOAD.cacheTtlSeconds);

  return { received: meta.received, total: meta.totalChunks };
}

function uploadFinish_(token) {
  return withImportLock_(() => {
    const cache = CacheService.getScriptCache();
    const metaRaw = cache.get(token + ":meta");
    if (!metaRaw) throw new Error("Sessão de upload expirada antes de finalizar.");

    const meta = JSON.parse(metaRaw);
    if (!meta.totalChunks || meta.totalChunks < 1) throw new Error("Sessão inválida: totalChunks vazio.");

    const parts = [];
    for (let i = 0; i < meta.totalChunks; i++) {
      const p = cache.get(token + ":chunk:" + i);
      if (!p) throw new Error(`Faltou o chunk ${i + 1}/${meta.totalChunks}. Tente novamente.`);
      parts.push(p);
    }

    const base64Data = parts.join("");
    const bytes = Utilities.base64Decode(base64Data);

    const ss = SpreadsheetApp.getActive();
    const folder = resolveUploadFolder_(ss);

    const mimeType = normalizeExcelMime_(meta.fileName, meta.mimeType);
    const blob = Utilities.newBlob(bytes, mimeType, meta.fileName);
    const uploaded = folder.createFile(blob).setName(`[UPLOAD] ${meta.fileName}`);

    try {
      let msg = "";
      const src = String(meta.sourceType || "OMNI").toUpperCase();
      if (src === "NIARA") {
        msg = pipelineFromNiaraFileId_(uploaded.getId(), uploaded.getName(), { source: "SIDEBAR_CHUNKS_NIARA" });
      } else if (src === "BEE2PAY") {
        msg = pipelineFromBee2PayFileId_(uploaded.getId(), uploaded.getName(), { source: "SIDEBAR_CHUNKS_BEE2PAY" });
      } else {
        msg = pipelineFromExcelFileId_(uploaded.getId(), uploaded.getName(), { source: "SIDEBAR_CHUNKS", auditorName: meta.auditorName || "" });
      }
      return msg || "Importação concluída (upload em chunks).";
    } finally {
      if (UX.trashTempFiles) safeTrashFile_(uploaded.getId());
      // cleanup best-effort
      try { cache.remove(token + ":meta"); } catch (_) {}
      for (let i = 0; i < meta.totalChunks; i++) {
        try { cache.remove(token + ":chunk:" + i); } catch (_) {}
      }
    }
  });
}

/** =======================
 *  Upload resumável (Drive API v3 via UrlFetch)
 *  ======================= */
function resumableStart_(fileName, mimeType, totalSizeBytes, sourceType) {
  assertDriveAdvancedServiceEnabled_();

  const ss = SpreadsheetApp.getActive();
  const folder = resolveUploadFolder_(ss);

  const safeName = `[UPLOAD] ${String(fileName || "arquivo")}`;
  const mt = normalizeExcelMime_(fileName, mimeType);
  const total = Number(totalSizeBytes) || 0;
  if (!total || total < 1) throw new Error("Tamanho total inválido para upload resumável.");

  const meta = {
    name: safeName,
    mimeType: mt,
    parents: folder ? [folder.getId()] : undefined
  };

  const url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable";
  const headers = {
    Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    "X-Upload-Content-Type": mt,
    "X-Upload-Content-Length": String(total)
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json; charset=UTF-8",
    payload: JSON.stringify(meta),
    headers,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const loc = (res.getHeaders() || {}).Location || (res.getHeaders() || {}).location;
  if (code < 200 || code > 299 || !loc) {
    throw new Error("Falha ao iniciar upload resumável. Código: " + code + " / " + res.getContentText());
  }

  const token = Utilities.getUuid();
  const session = {
    token,
    uploadUrl: loc,
    totalSizeBytes: total,
    fileName: String(fileName || ""),
    mimeType: mt,
    sourceType: String(sourceType || "OMNI").toUpperCase(),
    received: 0,
    done: false
  };
  saveResumableSession_(token, session);
  return session;
}

function resumableChunk_(token, offset, chunkBase64, totalSizeBytes, uploadUrl, fileName, auditorName, sourceType) {
  let session = null;
  let url = uploadUrl;
  let name = String(fileName || "");

  if (token) {
    session = getResumableSession_(token);
    if (session && session.done) {
      return { done: true, message: session.doneMessage || "Importação concluída." };
    }
  }

  if (!url && session) {
    url = session.url || session.uploadUrl;
    if (!name) name = String(session.fileName || "");
  }

  if (!url) throw new Error("URL de upload resumável ausente.");

  const total = Number(totalSizeBytes) || (session && session.totalSizeBytes) || 0;
  if (!total || total < 1) throw new Error("Total inválido na sessão resumável.");

  const start = Number(offset) || 0;
  if (session && start !== Number(session.received || 0)) {
    throw new Error("Offset inesperado. Esperado: " + session.received + " / Recebido: " + start);
  }

  const bytes = Utilities.base64Decode(String(chunkBase64 || ""));
  if (!bytes || !bytes.length) throw new Error("Chunk vazio.");

  const end = start + bytes.length - 1;
  if (end >= total) {
    throw new Error("Chunk excede o tamanho total esperado.");
  }

  const headers = {
    Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    "Content-Range": "bytes " + start + "-" + end + "/" + total
  };

  const res = UrlFetchApp.fetch(url, {
    method: "put",
    contentType: "application/octet-stream",
    payload: bytes,
    headers,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code === 308) {
    // Incompleto, atualiza range recebido
    if (session && token) {
      const range = (res.getHeaders() || {}).Range || (res.getHeaders() || {}).range;
      if (range) {
        const m = String(range).match(/bytes=0-(\d+)/);
        if (m) session.received = Number(m[1]) + 1;
      } else {
        session.received = end + 1;
      }
      saveResumableSession_(token, session);
      return { done: false, received: session.received, total };
    }
    return { done: false, received: end + 1, total };
  }

  if (code === 200 || code === 201) {
    // Upload completo
    let file = null;
    try { file = JSON.parse(res.getContentText() || "{}"); } catch (_) {}
    const fileId = file && file.id ? file.id : "";
    if (!fileId) {
      throw new Error("Upload completo, mas não recebi o ID do arquivo.");
    }

    try {
      const src = String(sourceType || (session && session.sourceType) || "OMNI").toUpperCase();
      const msg = withImportLock_(() => (
        src === "NIARA"
          ? pipelineFromNiaraFileId_(fileId, name || file.name || "", { source: "SIDEBAR_RESUMABLE_NIARA" })
          : (src === "BEE2PAY"
            ? pipelineFromBee2PayFileId_(fileId, name || file.name || "", { source: "SIDEBAR_RESUMABLE_BEE2PAY" })
            : pipelineFromExcelFileId_(fileId, name || file.name || "", { source: "SIDEBAR_RESUMABLE", auditorName }))
      ));
      if (session && token) {
        session.done = true;
        session.doneMessage = msg || "Importação concluída.";
        saveResumableSession_(token, session);
      }
      return { done: true, message: msg || "Importação concluída." };
    } catch (e) {
      if (token) cleanupResumableSession_(token);
      throw e;
    }
  }

  throw new Error("Falha no upload resumável. Código: " + code + " / " + res.getContentText());
}

function resumableFinish_(token) {
  if (token) cleanupResumableSession_(token);
  return { done: true };
}

function saveResumableSession_(token, session) {
  const raw = JSON.stringify(session || {});
  const key = "ru:" + token;
  try { CacheService.getScriptCache().put(key, raw, UPLOAD.cacheTtlSeconds); } catch (_) {}
  try { CacheService.getUserCache().put(key, raw, UPLOAD.cacheTtlSeconds); } catch (_) {}
  try { PropertiesService.getDocumentProperties().setProperty(key, raw); } catch (_) {}
  try { PropertiesService.getUserProperties().setProperty(key, raw); } catch (_) {}
  try { PropertiesService.getScriptProperties().setProperty(key, raw); } catch (_) {}
}

function getResumableSession_(token) {
  const key = "ru:" + token;
  let raw = null;
  try { raw = CacheService.getScriptCache().get(key); } catch (_) {}
  if (!raw) { try { raw = CacheService.getUserCache().get(key); } catch (_) {} }
  if (!raw) { try { raw = PropertiesService.getDocumentProperties().getProperty(key); } catch (_) {} }
  if (!raw) { try { raw = PropertiesService.getUserProperties().getProperty(key); } catch (_) {} }
  if (!raw) { try { raw = PropertiesService.getScriptProperties().getProperty(key); } catch (_) {} }
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}

function cleanupResumableSession_(token) {
  const key = "ru:" + token;
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
  try { CacheService.getUserCache().remove(key); } catch (_) {}
  try { PropertiesService.getDocumentProperties().deleteProperty(key); } catch (_) {}
  try { PropertiesService.getUserProperties().deleteProperty(key); } catch (_) {}
  try { PropertiesService.getScriptProperties().deleteProperty(key); } catch (_) {}
}

/** =======================
 *  Import por ID/URL de arquivo
 *  ======================= */
function importOmnibeesByFileIdPrompt() {
  return withImportLock_(() => {
    assertDriveAdvancedServiceEnabled_();

    const ui = SpreadsheetApp.getUi();
    const auditorName = promptAuditorName_(ui);
    if (auditorName === null) return;
    const resp = ui.prompt(
      "Importar por ID/URL do arquivo",
      "Cole o ID ou URL do arquivo .xls/.xlsx no Drive:",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const raw = (resp.getResponseText() || "").trim();
    const fileId = extractDriveId_(raw);
    if (!fileId) throw new Error("Não consegui extrair um ID válido do arquivo.");

    const file = DriveApp.getFileById(fileId);
    pipelineFromExcelFileId_(file.getId(), file.getName(), { source: "DRIVE_FILE", auditorName });
  });
}

/** =======================
 *  Import último arquivo da pasta
 *  ======================= */
function importLatestOmnibeesFromFolderPrompt() {
  return withImportLock_(() => {
    assertDriveAdvancedServiceEnabled_();

    const ui = SpreadsheetApp.getUi();
    const auditorName = promptAuditorName_(ui);
    if (auditorName === null) return;
    const resp = ui.prompt(
      "Importar último arquivo da pasta",
      "Cole o ID ou URL da pasta do Drive onde você salva os relatórios:",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const raw = (resp.getResponseText() || "").trim();
    const folderId = extractDriveId_(raw);
    if (!folderId) throw new Error("Não consegui extrair um ID válido da pasta.");

    const folder = DriveApp.getFolderById(folderId);
    const latest = findLatestExcelFileInFolder_(folder);
    if (!latest) throw new Error("Não encontrei nenhum .xls/.xlsx nessa pasta.");

    pipelineFromExcelFileId_(latest.getId(), latest.getName(), { source: "DRIVE_FOLDER_LATEST", folderId, auditorName });
  });
}

/** =======================
 *  Pipeline principal
 *  ======================= */
function pipelineFromExcelFileId_(excelFileId, excelName, meta) {
  assertDriveAdvancedServiceEnabled_();
  meta = meta || {};

  const t0 = Date.now();
  const ss = SpreadsheetApp.getActive();

  const sigInfo = computeImportSignature_(excelFileId, excelName);
  const dedupe = beginImportDedupe_(sigInfo.signature, sigInfo.fileName);
  if (dedupe && dedupe.skip) return dedupe.message;

  let convertedId = "";
  let extracted, auditSheet, normRowsCount = 0, auditRowsCount = 0;
  try {
    convertedId = convertExcelToGoogleSheet_(excelFileId, excelName);
    extracted = extractOmniReport_(convertedId);

    // 1) OMNI_RAW (snapshot bruto) - descontinuado
    if (UX.enableOmniRaw) {
      writeOmniRaw_(ss, UX.omniRawSheetName, extracted.raw, {
        sourceFileName: excelName,
        sourceFileId: excelFileId,
        sourceSheetName: extracted.sourceSheetName,
        conversionId: convertedId,
        importSource: meta.source || "",
        truncated: extracted.rawTruncated
      });
    }

    // 2) AUDIT sheet
    const auditRows = (String(UX.auditSource || "").toUpperCase() === "RAW")
      ? extracted.detailRows
      : extracted.normRows;
    auditRowsCount = auditRows.length;
    auditSheet = createAuditSheetForDate_(ss, extracted.inferredAuditDate);
    renderAuditSheet_(auditSheet, auditRows, extracted.inferredAuditDate, meta.auditorName || "", extracted.reportMeta || {});

    // 3) AUDIT_LOG
    writeAuditLog_(ss, {
      ts: new Date(),
      user: safeActiveUserEmail_(),
      importSource: meta.source || "",
      sourceFileName: excelName,
      sourceFileId: excelFileId,
      createdAuditSheetName: auditSheet.getName(),
      auditorName: meta.auditorName || "",
      reservations: auditRowsCount || normRowsCount,
      durationMs: Date.now() - t0
    });

    ss.setActiveSheet(auditSheet);
    finalizeImportDedupe_(sigInfo.signature, auditSheet.getName());
    return `Importação concluída. Aba: ${auditSheet.getName()}.`;
  } catch (e) {
    finalizeImportDedupe_(sigInfo.signature, "", e);
    throw e;
  } finally {
    if (convertedId && UX.trashTempFiles) safeTrashFile_(convertedId);
  }
}

/** =======================
 *  Pipeline Niara (atualiza aba de auditoria existente)
 *  ======================= */
function pipelineFromNiaraFileId_(excelFileId, excelName, meta) {
  assertDriveAdvancedServiceEnabled_();
  meta = meta || {};

  const t0 = Date.now();
  const ss = SpreadsheetApp.getActive();

  let convertedId = "";
  try {
    convertedId = convertExcelToGoogleSheet_(excelFileId, excelName);
    const niara = extractNiaraReport_(convertedId);

    const auditSheet = findAuditSheetByDate_(ss, niara.inferredDate);
    if (!auditSheet) {
      const dateLabel = niara.inferredDate ? formatDate_(niara.inferredDate) : "(data indefinida)";
      throw new Error(
        "Nenhuma aba de auditoria encontrada para a data " + dateLabel + ".\n" +
        "O relatório Niara parece ser de " + dateLabel + "."
      );
    }

    const result = applyNiaraToAuditSheet_(auditSheet, niara);

    writeNiaraLog_(ss, {
      ts: new Date(),
      sourceFileName: excelName,
      sourceFileId: excelFileId,
      auditSheetName: auditSheet.getName(),
      inferredDate: niara.inferredDate,
      totalReservations: niara.reservations.length,
      matched: result.matched,
      missing: result.missing,
      missingList: result.missingList
    });

    const durationMs = Date.now() - t0;
    return `Niara concluído. Aba: ${auditSheet.getName()}. ` +
      `Encontradas: ${result.matched}/${niara.reservations.length}. ` +
      (result.missing ? `Faltantes: ${result.missing} (ver NIARA_LOG). ` : "") +
      `Tempo: ${durationMs}ms.`;
  } finally {
    if (convertedId && UX.trashTempFiles) safeTrashFile_(convertedId);
  }
}

/** =======================
 *  Pipeline Bee2Pay (atualiza aba de auditoria existente)
 *  ======================= */
function pipelineFromBee2PayFileId_(excelFileId, excelName, meta) {
  assertDriveAdvancedServiceEnabled_();
  meta = meta || {};

  const t0 = Date.now();
  const ss = SpreadsheetApp.getActive();

  let convertedId = "";
  try {
    convertedId = convertExcelToGoogleSheet_(excelFileId, excelName);
    const bee = extractBee2PayReport_(convertedId);

    const auditSheet = findAuditSheetByDate_(ss, bee.inferredDate);
    if (!auditSheet) {
      const dateLabel = bee.inferredDate ? formatDate_(bee.inferredDate) : "(data indefinida)";
      throw new Error(
        "Nenhuma aba de auditoria encontrada para a data " + dateLabel + ".\n" +
        "O relatório Bee2Pay parece ser de " + dateLabel + "."
      );
    }

    const result = applyBee2PayToAuditSheet_(auditSheet, bee);

    writeBee2PayLog_(ss, {
      ts: new Date(),
      sourceFileName: excelName,
      sourceFileId: excelFileId,
      auditSheetName: auditSheet.getName(),
      inferredDate: bee.inferredDate,
      totalTransactions: bee.transactions.length,
      matched: result.matched,
      updatedPayments: result.updatedPayments,
      markedNotPaid: result.markedNotPaid
    });

    const durationMs = Date.now() - t0;
    return `Bee2Pay concluído. Aba: ${auditSheet.getName()}. ` +
      `Transações: ${bee.transactions.length}. Atualizadas: ${result.updatedPayments}. ` +
      `Tempo: ${durationMs}ms.`;
  } finally {
    if (convertedId && UX.trashTempFiles) safeTrashFile_(convertedId);
  }
}

/** =======================
 *  Extração robusta (RAW + NORM)
 *  ======================= */
function extractOmniReport_(convertedSpreadsheetId) {
  const omniSS = openSpreadsheetWithRetry_(convertedSpreadsheetId);
  const sheets = omniSS.getSheets();

  let target = null;
  let values = null;
  let headerRowIndex = -1;

  // varre abas até achar header com "Res. Nº"
  for (const sh of sheets) {
    const v = sh.getDataRange().getValues();
    const hri = findHeaderRowIndex_(v);
    if (hri >= 0) {
      target = sh;
      values = v;
      headerRowIndex = hri;
      break;
    }
  }

  if (!target || !values || headerRowIndex < 0) {
    throw new Error("Cabeçalho não encontrado em nenhuma aba. Procurei por 'Res. Nº' nas primeiras 80 linhas de cada aba.");
  }

  // RAW: recorta (limites) a partir do header até o fim, mas mantendo o header como linha 1 do raw
  const raw = [];
  const maxR = Math.min(values.length, headerRowIndex + 1 + UX.rawMaxRows);
  const maxC = Math.min(values[headerRowIndex].length || 0, UX.rawMaxCols);

  // inclui header e linhas seguintes
  for (let r = headerRowIndex; r < maxR; r++) {
    raw.push(values[r].slice(0, maxC));
  }

  const rawTruncated = (values.length - headerRowIndex) > UX.rawMaxRows || (values[headerRowIndex].length || 0) > UX.rawMaxCols;

  // NORM
  const header = values[headerRowIndex].map(v => String(v || "").trim());
  const idx = {
    res:      findColIndex_(header, ["Res. Nº", "Res. N°", "Res. No", "Res. Nº "]),
    estado:   findColIndex_(header, ["Estado"]),
    data:     findColIndex_(header, ["Data", "Data Criação", "Data de Criação"]),
    checkin:  findColIndex_(header, ["Check In", "Check-in", "Check In "]),
    checkout: findColIndex_(header, ["Check Out", "Check-out", "Check Out "]),
    hospede:  findColIndex_(header, ["Hóspede", "Hospede"]),
    canal:    findColIndex_(header, ["Canal/A.V./Empresa", "Canal", "Empresa"]),
    apto:     findColIndex_(header, ["Apartamento", "Apto", "Ap.", "Nº Apto"]),
    tarifa:   findColIndex_(header, ["Tarifa", "Tarifário", "Tarifario"]),
    total:    findColIndex_(header, ["Total", "Total Geral", "Total (R$)"]),
  };

  const required = ["res","estado","data","checkin","checkout","hospede","canal","tarifa","total"];
  const missing = required.filter(k => idx[k] < 0);
  if (missing.length) {
    throw new Error(
      `Relatório Omnibees com colunas inesperadas. Faltando: ${missing.join(", ")}.\n` +
      `Aba detectada: ${target.getName()}\n` +
      `Cabeçalho detectado: ${header.join(" | ")}`
    );
  }

  const grouped = new Map();
  const detailRows = [];
  const auditDateFreq = {};

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const res = String(row[idx.res] || "").trim();
    if (!res) continue;

    const item = {
      res,
      estado: row[idx.estado],
      data: row[idx.data],
      checkin: row[idx.checkin],
      checkout: row[idx.checkout],
      hospede: row[idx.hospede],
      canal: row[idx.canal],
      apto: idx.apto >= 0 ? row[idx.apto] : "",
      tarifa: row[idx.tarifa],
      totalNum: parseBRL_(row[idx.total]),
    };

    if (!grouped.has(res)) grouped.set(res, []);
    grouped.get(res).push(item);

    // Linha a linha (RAW -> detalhado)
    const status = normalizeStatus_(item.estado);
    const titular = extractTitular_(item.hospede);
    const origem = normalizeOrigin_(item.canal);
    detailRows.push([
      res,                       // 0 Localizador
      status,                    // 1 Status
      asDateOnly_(item.data),    // 2 Data criação
      origem,                    // 3 Origem
      asDateOnly_(item.checkin), // 4 Check-in
      asDateOnly_(item.checkout),// 5 Check-out
      titular,                   // 6 Titular
      1,                         // 7 Quartos (linha = 1 quarto)
      item.tarifa,               // 8 Tarifário
      item.totalNum,             // 9 Total
      item.apto || "",           // 10 Aptos
      "",                        // 11 Flags
      ""                         // 12 Observação auto
    ]);
  }

  const normRows = [];
  grouped.forEach((items, res) => {
    const first = items[0];

    const status = normalizeStatus_(first.estado);
    const titular = extractTitular_(first.hospede);
    const origem = normalizeOrigin_(first.canal);

    const quartos = items.length;
    const tarifarios = unique_(items.map(i => i.tarifa)).join(" + ");
    const aptos = unique_(items.map(i => i.apto)).join(" + ");
    const total = items.reduce((acc, i) => acc + (Number(i.totalNum) || 0), 0);

    const d = first.data instanceof Date ? asDateOnly_(first.data) : null;
    if (d) {
      const key = Utilities.formatDate(d, UX.timezone, "yyyy-MM-dd");
      auditDateFreq[key] = (auditDateFreq[key] || 0) + 1;
    }

    const flags = [];
    if (tarifarios.includes(" + ")) flags.push("TARIFAS_MULTIPLAS");
    if (aptos.includes(" + ")) flags.push("APTOS_MULTIPLOS");

    normRows.push([
      res,                    // 0 Localizador
      status,                 // 1 Status
      asDateOnly_(first.data),// 2 Data criação
      origem,                 // 3 Origem
      asDateOnly_(first.checkin), // 4 Check-in
      asDateOnly_(first.checkout),// 5 Check-out
      titular,                // 6 Titular
      quartos,                // 7 Quartos
      tarifarios,             // 8 Tarifário(s)
      total,                  // 9 Total
      aptos,                  // 10 Aptos
      flags.join(","),        // 11 Flags
      ""                      // 12 Observação auto (reservado)
    ]);
  });

  normRows.sort((a, b) => {
    const da = a[4] instanceof Date ? a[4].getTime() : 0;
    const db = b[4] instanceof Date ? b[4].getTime() : 0;
    return da === db ? String(a[0]).localeCompare(String(b[0])) : da - db;
  });

  const inferredAuditDate = pickMostFrequentDate_(auditDateFreq) || new Date();
  return {
    raw,
    rawTruncated,
    normRows,
    detailRows,
    inferredAuditDate,
    reportMeta: extractReportMeta_(values, headerRowIndex),
    sourceSheetName: target.getName()
  };
}

/** =======================
 *  Extração Niara (Reservas + Pagamentos)
 *  ======================= */
function extractNiaraReport_(convertedSpreadsheetId) {
  const niaraSS = openSpreadsheetWithRetry_(convertedSpreadsheetId);
  const reservasSheet = niaraSS.getSheetByName("Reservas");
  const pagamentosSheet = niaraSS.getSheetByName("Pagamentos");
  if (!reservasSheet || !pagamentosSheet) {
    throw new Error("Relatório Niara inválido: abas 'Reservas' e/ou 'Pagamentos' não encontradas.");
  }

  const resValues = reservasSheet.getDataRange().getValues();
  const payValues = pagamentosSheet.getDataRange().getValues();
  if (!resValues.length || !payValues.length) {
    throw new Error("Relatório Niara inválido: abas vazias.");
  }

  const resHeader = resValues[0].map(v => String(v || "").trim());
  const payHeader = payValues[0].map(v => String(v || "").trim());

  const resIdx = {
    localizador: findColIndex_(resHeader, ["Localizador"]),
    idViagem: findColIndex_(resHeader, ["ID Viagem", "Id Viagem"]),
    idReserva: findColIndex_(resHeader, ["ID Reserva", "Id Reserva"]),
    codigoPms: findColIndex_(resHeader, ["Código PMS", "Codigo PMS"]),
    dataReserva: findColIndex_(resHeader, ["Data da reserva"]),
    situacao: findColIndex_(resHeader, ["Situação", "Situacao"]),
    valor: findColIndex_(resHeader, ["Valor"]),
    tipoPagamento: findColIndex_(resHeader, ["Tipo de pagamento"])
  };
  const missingRes = ["localizador","dataReserva"].filter(k => resIdx[k] < 0);
  if (missingRes.length) {
    throw new Error("Relatório Niara (Reservas) com colunas inesperadas. Faltando: " + missingRes.join(", "));
  }

  const payIdx = {
    localizador: findColIndex_(payHeader, ["Localizador da reserva"]),
    status: findColIndex_(payHeader, ["Status da transação", "Status da transacao"]),
    valor: findColIndex_(payHeader, ["Valor total apropriado à reserva", "Valor total", "Valor da reserva"]),
    tipo: findColIndex_(payHeader, ["Tipo da transação", "Tipo da transacao"]),
    data: findColIndex_(payHeader, ["Data da transação", "Data da transacao"])
  };
  const missingPay = ["localizador","status"].filter(k => payIdx[k] < 0);
  if (missingPay.length) {
    throw new Error("Relatório Niara (Pagamentos) com colunas inesperadas. Faltando: " + missingPay.join(", "));
  }

  const reservations = [];
  const dateFreq = {};
  for (let r = 1; r < resValues.length; r++) {
    const row = resValues[r];
    const locRaw = String(row[resIdx.localizador] || "").trim();
    if (!locRaw) continue;

    const loc = normalizeLocalizador_(locRaw);
    const idViagem = resIdx.idViagem >= 0 ? String(row[resIdx.idViagem] || "").trim() : "";
    const idReserva = resIdx.idReserva >= 0 ? String(row[resIdx.idReserva] || "").trim() : "";
    const codigoPms = resIdx.codigoPms >= 0 ? String(row[resIdx.codigoPms] || "").trim() : "";
    const dataReserva = coerceDate_(row[resIdx.dataReserva]);
    const situacao = resIdx.situacao >= 0 ? String(row[resIdx.situacao] || "").trim() : "";
    const tipoPagamento = resIdx.tipoPagamento >= 0 ? String(row[resIdx.tipoPagamento] || "").trim() : "";
    const valor = resIdx.valor >= 0 ? parseBRL_(row[resIdx.valor]) : 0;

    if (dataReserva) {
      const key = Utilities.formatDate(dataReserva, UX.timezone, "yyyy-MM-dd");
      dateFreq[key] = (dateFreq[key] || 0) + 1;
    }

    reservations.push({
      localizador: loc,
      localizadorRaw: locRaw,
      idViagem,
      idReserva,
      codigoPms,
      dataReserva,
      situacao,
      tipoPagamento,
      valor
    });
  }

  const paymentsMap = {};
  for (let r = 1; r < payValues.length; r++) {
    const row = payValues[r];
    const locRaw = String(row[payIdx.localizador] || "").trim();
    if (!locRaw) continue;
    const loc = normalizeLocalizador_(locRaw);

    const status = payIdx.status >= 0 ? String(row[payIdx.status] || "").trim() : "";
    const tipo = payIdx.tipo >= 0 ? String(row[payIdx.tipo] || "").trim() : "";
    const valor = payIdx.valor >= 0 ? parseBRL_(row[payIdx.valor]) : 0;
    const data = payIdx.data >= 0 ? coerceDate_(row[payIdx.data]) : null;

    if (!paymentsMap[loc]) {
      paymentsMap[loc] = { paid: false, totalPaid: 0, methods: {}, statuses: {}, lastDate: null };
    }

    const p = paymentsMap[loc];
    if (tipo) p.methods[tipo] = true;
    if (status) p.statuses[status] = true;

    if (isPaidStatus_(status)) {
      p.paid = true;
      p.totalPaid += Number(valor) || 0;
    }

    if (data instanceof Date) {
      if (!p.lastDate || data.getTime() > p.lastDate.getTime()) p.lastDate = data;
    }
  }

  const inferredDate = pickMostFrequentDate_(dateFreq) || null;

  return {
    reservations,
    paymentsMap,
    inferredDate
  };
}

/** =======================
 *  Extração Bee2Pay (Relatório de Transações)
 *  ======================= */
function extractBee2PayReport_(convertedSpreadsheetId) {
  const beeSS = openSpreadsheetWithRetry_(convertedSpreadsheetId);
  const sheet = beeSS.getSheetByName("Relatório de Transações") || beeSS.getSheets()[0];
  if (!sheet) throw new Error("Relatório Bee2Pay inválido: aba não encontrada.");

  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("Relatório Bee2Pay inválido: aba vazia.");

  // Período Listado (linha de cabeçalho superior)
  let periodLabel = "";
  for (let r = 0; r < Math.min(values.length, 10); r++) {
    for (let c = 0; c < values[r].length; c++) {
      const s = String(values[r][c] || "").trim();
      if (!s) continue;
      if (s.toLowerCase().includes("período listado") || s.toLowerCase().includes("periodo listado")) {
        for (let k = 1; k <= 6 && c + k < values[r].length; k++) {
          const s2 = String(values[r][c + k] || "").trim();
          if (s2 && /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(s2)) { periodLabel = s2; break; }
        }
      }
    }
  }

  const periodDates = [];
  if (periodLabel) {
    const parts = periodLabel.split("-").map(s => s.trim());
    parts.forEach(p => { const d = parseDateTimeFromText_(p); if (d) periodDates.push(d); });
  }

  // Encontra header (linha com LOCALIZADOR + DATA DA RESERVA)
  let headerRow = -1;
  for (let r = 0; r < Math.min(values.length, 30); r++) {
    const row = values[r].map(v => String(v || "").trim().toUpperCase());
    if (row.includes("LOCALIZADOR") && row.includes("DATA DA RESERVA")) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) throw new Error("Relatório Bee2Pay: cabeçalho não encontrado.");

  const header = values[headerRow].map(v => String(v || "").trim());
  const idx = {
    localizador: findColIndex_(header, ["LOCALIZADOR"]),
    dataReserva: findColIndex_(header, ["DATA DA RESERVA"]),
    statusCobranca: findColIndex_(header, ["STATUS COBRANÇA", "STATUS COBRANCA"]),
    retorno: findColIndex_(header, ["RETORNO"]),
    formaPagamento: findColIndex_(header, ["FORMA DE PAGAMENTO", "FORMA DE PAG."]),
    valorTransacionado: findColIndex_(header, ["VALOR TRANSACIONADO"]),
    idTransacao: findColIndex_(header, ["ID DA TRANSAÇÃO", "ID DA TRANSACAO"]),
    dataTransacao: findColIndex_(header, ["DATA DA TRANSAÇÃO", "DATA DA TRANSACAO"])
  };
  const missing = ["localizador","dataReserva","statusCobranca","valorTransacionado","idTransacao"].filter(k => idx[k] < 0);
  if (missing.length) {
    throw new Error("Relatório Bee2Pay com colunas inesperadas. Faltando: " + missing.join(", "));
  }

  const transactions = [];
  const dateFreq = {};
  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    const locRaw = String(row[idx.localizador] || "").trim();
    if (!locRaw) continue;

    const loc = normalizeLocalizador_(locRaw);
    const dataReserva = coerceDate_(row[idx.dataReserva]);
    const statusCobranca = idx.statusCobranca >= 0 ? String(row[idx.statusCobranca] || "").trim() : "";
    const retorno = idx.retorno >= 0 ? String(row[idx.retorno] || "").trim() : "";
    const formaPagamento = idx.formaPagamento >= 0 ? String(row[idx.formaPagamento] || "").trim() : "";
    const valorTransacionado = idx.valorTransacionado >= 0 ? parseBRL_(row[idx.valorTransacionado]) : 0;
    const idTransacao = idx.idTransacao >= 0 ? String(row[idx.idTransacao] || "").trim() : "";
    const dataTransacao = idx.dataTransacao >= 0 ? coerceDate_(row[idx.dataTransacao]) : null;

    if (dataReserva) {
      const key = Utilities.formatDate(dataReserva, UX.timezone, "yyyy-MM-dd");
      dateFreq[key] = (dateFreq[key] || 0) + 1;
    }

    const paid = isBee2PayPaid_(statusCobranca, retorno, valorTransacionado);

    transactions.push({
      localizador: loc,
      localizadorRaw: locRaw,
      dataReserva,
      statusCobranca,
      retorno,
      formaPagamento,
      valorTransacionado,
      idTransacao,
      dataTransacao,
      paid
    });
  }

  const inferredDate = periodDates.length ? periodDates[0] : (pickMostFrequentDate_(dateFreq) || null);
  return { transactions, inferredDate, periodLabel };
}

function pickMostFrequentDate_(freqMap) {
  let bestKey = null, bestN = 0;
  Object.keys(freqMap || {}).forEach(k => {
    if (freqMap[k] > bestN) { bestN = freqMap[k]; bestKey = k; }
  });
  if (!bestKey) return null;
  const [y,m,d] = bestKey.split("-").map(Number);
  return new Date(y, m - 1, d);
}

/** =======================
 *  OMNI_RAW (snapshot bruto com design)
 *  ======================= */
function writeOmniRaw_(ss, sheetName, rawValues, meta) {
  meta = meta || {};
  const sh = upsertSheet_(ss, sheetName);
  sh.clear();
  sh.setHiddenGridlines(true);

  // layout: 4 linhas de header + tabela a partir da linha 6
  const topRows = 5;
  const startRow = topRows + 1;

  const rows = (rawValues && rawValues.length) ? rawValues.length : 1;
  const cols = (rawValues && rawValues[0] && rawValues[0].length) ? rawValues[0].length : 1;

  ensureSheetSize_(sh, startRow + rows + 10, Math.max(cols, 10));

  // Col widths: meta mais largo nas primeiras colunas
  for (let c = 1; c <= Math.min(10, sh.getMaxColumns()); c++) sh.setColumnWidth(c, c <= 2 ? 180 : 140);

  // Header band
  sh.getRange(1, 1, topRows, Math.max(cols, 10)).setBackground(THEME.headerBand);
  sh.getRange(1, 1, topRows, Math.max(cols, 10)).setFontFamily("Roboto").setFontColor(THEME.text);

  sh.getRange(1, 1, 2, 6).merge()
    .setValue("OMNI_RAW — Snapshot do Relatório Importado")
    .setFontSize(16).setFontWeight("bold");

  sh.getRange(3, 1, 1, 6).merge()
    .setValue("Dados brutos (última importação). Use para rastreabilidade e reprocessamento.")
    .setFontSize(10).setFontColor(THEME.subtext);

  // Meta panel
  const panel = sh.getRange(1, 7, 4, 4);
  panel.setBackground(THEME.bg)
    .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID)
    .setFontSize(9).setFontColor(THEME.subtext);

  const now = new Date();
  const metaRows = [
    ["Arquivo:", meta.sourceFileName || ""],
    ["Fonte:", meta.importSource || ""],
    ["Aba:", meta.sourceSheetName || ""],
    ["Gerado em:", now]
  ];
  sh.getRange(1, 7, 4, 1).setValues(metaRows.map(r => [r[0]])).setFontWeight("bold");
  const metaRange = sh.getRange(1, 8, 4, 3);
  metaRange.mergeAcross();
  for (let i = 0; i < metaRows.length; i++) {
    sh.getRange(1 + i, 8, 1, 3).setValue(metaRows[i][1]);
  }
  sh.getRange(4, 8).setNumberFormat("dd/MM/yyyy HH:mm");

  if (meta.truncated) {
    sh.getRange(5, 1, 1, 10).merge()
      .setValue(`⚠️ RAW truncado por segurança (máx ${UX.rawMaxRows} linhas e ${UX.rawMaxCols} colunas).`)
      .setFontColor(THEME.warn).setFontSize(10).setFontWeight("bold");
  } else {
    sh.getRange(5, 1, 1, 10).merge()
      .setValue("RAW completo dentro dos limites configurados.")
      .setFontColor(THEME.subtext).setFontSize(10);
  }

  // Tabela
  const out = (rawValues && rawValues.length) ? rawValues : [["(Sem dados brutos)"]];
  sh.getRange(startRow, 1, out.length, out[0].length).setValues(out);

  // Styling header row (primeira linha do raw)
  sh.getRange(startRow, 1, 1, out[0].length)
    .setBackground(THEME.tableHeader)
    .setFontWeight("bold")
    .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  // Body
  const bodyRows = Math.max(1, out.length);
  sh.getRange(startRow, 1, bodyRows, out[0].length)
    .setFontFamily("Roboto")
    .setFontSize(9)
    .setFontColor(THEME.text);

  // Borders leve no corpo
  sh.getRange(startRow + 1, 1, Math.max(1, out.length - 1), out[0].length)
    .setBorder(true, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  // Freeze
  sh.setFrozenRows(topRows);
}

/** =======================
 *  OMNI_NORM (dataset persistido)
 *  ======================= */
function writeOmniNorm_(ss, sheetName, normRows) {
  const sh = upsertSheet_(ss, sheetName);
  sh.clear();
  sh.setHiddenGridlines(true);

  ensureSheetSize_(sh, Math.max(3, normRows.length + 2), 13);

  // Title band
  sh.getRange(1, 1, 2, 13).setBackground(THEME.headerBand);
  sh.getRange(1, 1, 2, 13).setFontFamily("Roboto").setFontColor(THEME.text);

  sh.getRange(1, 1, 1, 9).merge()
    .setValue("OMNI_NORM — Dataset Normalizado")
    .setFontSize(14).setFontWeight("bold");

  sh.getRange(2, 1, 1, 9).merge()
    .setValue("Base normalizada por Localizador (Res. Nº). Usada pelo DASHBOARD e como referência.")
    .setFontSize(10).setFontColor(THEME.subtext);

  const headers = [
    "Localizador (Res. Nº)","Status (Omni)","Data Criação","Origem","Check-in","Check-out",
    "Titular","Quartos","Tarifário(s)","Total (R$)","Aptos (auto)","Flags (auto)","Observação (auto)"
  ];

  // Header row (row 3)
  sh.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground(THEME.tableHeader)
    .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  if (normRows.length) {
    sh.getRange(4, 1, normRows.length, headers.length).setValues(normRows);
  } else {
    sh.getRange(4, 1).setValue("(Sem reservas normalizadas)").setFontColor(THEME.subtext);
  }

  // Formatting
  const dataRows = Math.max(normRows.length, 1);
  sh.getRange(4, 3, dataRows, 1).setNumberFormat("dd/MM/yyyy");
  sh.getRange(4, 5, dataRows, 2).setNumberFormat("dd/MM/yyyy");
  sh.getRange(4, 10, dataRows, 1).setNumberFormat('"R$" #,##0.00');
  sh.getRange(4, 1, dataRows, 1).setNumberFormat("@");

  sh.getRange(3, 1, dataRows + 1, headers.length)
    .setFontFamily("Roboto").setFontColor(THEME.text).setFontSize(9);

  // Column widths
  const widths = [170,120,110,150,105,105,220,70,200,110,220,140,260];
  for (let c = 1; c <= widths.length; c++) sh.setColumnWidth(c, widths[c - 1]);

  // Freeze
  sh.setFrozenRows(3);
}

/** =======================
 *  Criar aba de auditoria por data (sempre nova)
 *  ======================= */
function createAuditSheetForDate_(ss, auditDate) {
  const namePrimary = Utilities.formatDate(auditDate, UX.timezone, UX.auditSheetDateFormatPrimary);
  let baseName = sanitizeSheetName_(UX.auditSheetPrefix + namePrimary);

  if (!baseName) {
    const fallback = Utilities.formatDate(auditDate, UX.timezone, UX.auditSheetDateFormatFallback);
    baseName = sanitizeSheetName_(UX.auditSheetPrefix + fallback) || "AUDITORIA";
  }

  const finalName = uniqueSheetName_(ss, baseName);
  return ss.insertSheet(finalName);
}

function sanitizeSheetName_(name) {
  let n = String(name || "").trim();
  if (!n) return "";
  n = n.replace(/[\[\]\*\?:\\]/g, "");
  n = n.replace(/\s+/g, " ").trim();
  if (n.length > SHEET_NAME_MAX) n = n.slice(0, SHEET_NAME_MAX).trim();
  return n;
}

function uniqueSheetName_(ss, baseName) {
  let name = baseName.slice(0, SHEET_NAME_MAX).trim();
  if (!ss.getSheetByName(name)) return name;

  for (let i = 2; i < 200; i++) {
    const suffix = ` (${i})`;
    const cut = SHEET_NAME_MAX - suffix.length;
    const candidate = (name.length > cut ? name.slice(0, cut).trim() : name) + suffix;
    if (!ss.getSheetByName(candidate)) return candidate;
  }

  const ts = Utilities.formatDate(new Date(), UX.timezone, "yyyyMMdd_HHmmss");
  const suffix = ` ${ts}`;
  const cut = SHEET_NAME_MAX - suffix.length;
  return (baseName.slice(0, Math.max(1, cut)).trim() || "AUDITORIA") + suffix;
}

function findAuditSheetByDate_(ss, date) {
  if (!date) return null;
  const base = Utilities.formatDate(date, UX.timezone, UX.auditSheetDateFormatPrimary);
  let sh = ss.getSheetByName(base);
  if (sh) return sh;

  const all = ss.getSheets();
  for (const s of all) {
    const name = s.getName();
    if (name === base) return s;
    if (name.startsWith(base + " ")) return s;
    if (name.startsWith(base + " (")) return s;
  }
  return null;
}

function buildAuditIndex_(sheet) {
  const startRow = LAYOUT.startRow;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return {};

  const rows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 2, rows, 2).getValues(); // col B Sistema, col C Localizador
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const sys = String(data[i][0] || "").trim();
    const loc = String(data[i][1] || "").trim();
    if (!loc) continue;
    if (sys !== "Omnibees") continue;
    map[normalizeLocalizador_(loc)] = startRow + i;
  }
  return map;
}

function upsertObservation_(existing, addition) {
  const add = String(addition || "").trim();
  const cur = String(existing || "").trim();
  if (!add) return cur;
  if (!cur) return add;

  const parts = cur.split(" | ").map(s => s.trim()).filter(Boolean);
  const filtered = parts.filter(p => !/^niara:/i.test(p) && !/^bee2pay:/i.test(p));
  filtered.push(add);
  return filtered.join(" | ");
}

function applyNiaraToAuditSheet_(sheet, niara) {
  const index = buildAuditIndex_(sheet);
  const missingList = [];
  let matched = 0;
  let updatedPayments = 0;
  let updatedObs = 0;

  for (const r of niara.reservations) {
    const loc = normalizeLocalizador_(r.localizador);
    if (!loc) continue;
    const omniRow = index[loc];
    if (!omniRow) {
      missingList.push(r);
      continue;
    }

    matched++;
    const pay = niara.paymentsMap[loc];
    const isPaid = pay && pay.paid;
    const paymentLabel = isPaid ? "Pago" : "Não Pago";

    // Atualiza Pagamento SOMENTE na linha Omnibees
    const payCell = sheet.getRange(omniRow, 11);
    const url = buildNiaraServiceUrl_(r.idViagem, r.idReserva);
    if (url) {
      const style = SpreadsheetApp.newTextStyle()
        .setForegroundColor(THEME.tableText)
        .setUnderline(true)
        .build();
      const rich = SpreadsheetApp.newRichTextValue()
        .setText(paymentLabel)
        .setLinkUrl(url)
        .setTextStyle(0, paymentLabel.length, style)
        .build();
      payCell.setRichTextValue(rich);
    } else {
      payCell.setValue(paymentLabel);
    }
    updatedPayments += 1;

    // Link PMS no Hotelflow (coluna Sistema)
    const pmsUrl = buildPmsReservationUrl_(r.codigoPms);
    if (pmsUrl) {
      const pmsCell = sheet.getRange(omniRow + 1, 2);
      const sysLabel = "Hotelflow";
      const sysStyle = SpreadsheetApp.newTextStyle()
        .setForegroundColor(THEME.tableText)
        .setUnderline(true)
        .setBold(true)
        .build();
      const sysRich = SpreadsheetApp.newRichTextValue()
        .setText(sysLabel)
        .setLinkUrl(pmsUrl)
        .setTextStyle(0, sysLabel.length, sysStyle)
        .build();
      pmsCell.setRichTextValue(sysRich);
    }

    // Observações (mantém o que já existe)
    const methods = pay && pay.methods ? Object.keys(pay.methods).filter(Boolean) : [];
    const methodLabel = methods.length
      ? methods.join(" + ")
      : (r.tipoPagamento || "");
    const amountLabel = pay && pay.totalPaid ? formatBRL_(pay.totalPaid) : "";

    let niaraText = `Niara: ${paymentLabel}`;
    if (methodLabel) niaraText += ` (${methodLabel})`;
    if (amountLabel) niaraText += ` ${amountLabel}`;

    const obsCell = sheet.getRange(omniRow, 13);
    const existingObs = obsCell.getValue();
    const merged = upsertObservation_(existingObs, niaraText);
    if (merged !== existingObs) {
      obsCell.setValue(merged);
      updatedObs++;
    }
  }

  return {
    matched,
    missing: missingList.length,
    updatedPayments,
    updatedObs,
    missingList
  };
}

function applyBee2PayToAuditSheet_(sheet, bee) {
  const index = buildAuditIndex_(sheet);
  const map = {};
  for (const t of bee.transactions) {
    const loc = normalizeLocalizador_(t.localizador);
    if (!loc) continue;
    if (!map[loc]) map[loc] = { paid: false, idTransacao: "", totalPaid: 0, methods: {} };
    const m = map[loc];
    if (t.paid) {
      m.paid = true;
      m.idTransacao = t.idTransacao || m.idTransacao;
      m.totalPaid += Number(t.valorTransacionado) || 0;
      if (t.formaPagamento) m.methods[t.formaPagamento] = true;
    }
  }

  let matched = 0;
  let updatedPayments = 0;
  let markedNotPaid = 0;
  let updatedObs = 0;

  const skipTariff = (tariff) => {
    const t = normalizeText_(tariff);
    if (!t) return false;
    if (t.includes("niara") && t.includes("deposito")) return true;
    if (t.includes("tarifa a vista") && (t.includes("ted") || t.includes("pix"))) return true;
    return false;
  };

  Object.keys(index).forEach(loc => {
    const omniRow = index[loc];
    if (!omniRow) return;

    const tariff = sheet.getRange(omniRow, 9).getValue(); // Tarifário (col I)
    if (skipTariff(tariff)) return;

    const obsCell = sheet.getRange(omniRow, 13);
    const obs = String(obsCell.getValue() || "");
    if (/niara:/i.test(obs)) return; // preserva reservas Niara

    const payCell = sheet.getRange(omniRow, 11);
    const sysCell = sheet.getRange(omniRow, 2);
    const locCellValue = sheet.getRange(omniRow, 3).getValue();
    const locatorForUrl = normalizeLocalizador_(locCellValue || loc);
    const entry = map[loc];

    if (entry) {
      const sysLabel = String(sysCell.getValue() || "Omnibees").trim() || "Omnibees";
      const sysUrl = buildBee2PayReservationUrl_(locatorForUrl);
      if (sysUrl) {
        const sysStyle = SpreadsheetApp.newTextStyle()
          .setForegroundColor(THEME.tableText)
          .setUnderline(true)
          .setBold(true)
          .build();
        const sysRich = SpreadsheetApp.newRichTextValue()
          .setText(sysLabel)
          .setLinkUrl(sysUrl)
          .setTextStyle(0, sysLabel.length, sysStyle)
          .build();
        sysCell.setRichTextValue(sysRich);
      }
    }

    if (entry && entry.paid) {
      const label = "Pago";
      const url = buildBee2PayUrl_(entry.idTransacao) || buildBee2PayReservationUrl_(locatorForUrl);
      if (url) {
        const style = SpreadsheetApp.newTextStyle()
          .setForegroundColor(THEME.tableText)
          .setUnderline(true)
          .build();
        const rich = SpreadsheetApp.newRichTextValue()
          .setText(label)
          .setLinkUrl(url)
          .setTextStyle(0, label.length, style)
          .build();
        payCell.setRichTextValue(rich);
      } else {
        payCell.setValue(label);
      }
      matched++;
      updatedPayments++;

      const methods = entry && entry.methods ? Object.keys(entry.methods).filter(Boolean) : [];
      const methodLabel = methods.length ? methods.join(" + ") : "";
      const amountLabel = entry.totalPaid ? formatBRL_(entry.totalPaid) : "";

      let beeText = "Bee2Pay: Pago";
      if (methodLabel) beeText += ` (${methodLabel})`;
      if (amountLabel) beeText += ` ${amountLabel}`;

      const merged = upsertObservation_(obs, beeText);
      if (merged !== obs) {
        obsCell.setValue(merged);
        updatedObs++;
      }
    } else {
      payCell.setValue("Não Pago");
      markedNotPaid++;
      updatedPayments++;
    }
  });

  return { matched, updatedPayments, markedNotPaid, updatedObs };
}

function writeBee2PayLog_(ss, entry) {
  const name = "BEE2PAY_LOG";
  const sh = upsertSheet_(ss, name);
  sh.setHiddenGridlines(true);

  const headers = [
    "Timestamp",
    "Aba Auditoria",
    "Data Bee2Pay",
    "Arquivo",
    "FileId",
    "Transações",
    "Atualizadas",
    "Marcadas Não Pago"
  ];

  ensureSheetSize_(sh, Math.max(2, sh.getLastRow() || 2), headers.length);

  const needsHeader = (() => {
    if (sh.getLastRow() === 0) return true;
    const row = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    return headers.some((h, i) => String(row[i] || "").trim() !== h);
  })();

  if (needsHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(THEME.tableHeader)
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setFrozenRows(1);
    const widths = [160,160,120,240,260,120,120,140];
    for (let c = 1; c <= widths.length; c++) sh.setColumnWidth(c, widths[c - 1]);
  }

  const ts = entry.ts || new Date();
  const dateLabel = entry.inferredDate ? formatDate_(entry.inferredDate) : "";
  const row = [
    ts,
    entry.auditSheetName || "",
    dateLabel,
    entry.sourceFileName || "",
    entry.sourceFileId || "",
    entry.totalTransactions || 0,
    entry.updatedPayments || 0,
    entry.markedNotPaid || 0
  ];
  const next = sh.getLastRow() + 1;
  sh.getRange(next, 1, 1, headers.length).setValues([row]);
  sh.getRange(next, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange(next, 1, 1, headers.length).setFontFamily("Roboto").setFontSize(9).setFontColor(THEME.text);
}

function writeNiaraLog_(ss, entry) {
  const name = "NIARA_LOG";
  const sh = upsertSheet_(ss, name);
  sh.setHiddenGridlines(true);

  const headers = [
    "Timestamp",
    "Aba Auditoria",
    "Data Niara",
    "Arquivo",
    "FileId",
    "Reservas (Niara)",
    "Encontradas",
    "Faltantes"
  ];

  ensureSheetSize_(sh, Math.max(2, sh.getLastRow() || 2), headers.length);

  const needsHeader = (() => {
    if (sh.getLastRow() === 0) return true;
    const row = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    return headers.some((h, i) => String(row[i] || "").trim() !== h);
  })();

  if (needsHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(THEME.tableHeader)
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setFrozenRows(1);
    const widths = [160,160,120,240,260,120,120,120];
    for (let c = 1; c <= widths.length; c++) sh.setColumnWidth(c, widths[c - 1]);
  }

  const ts = entry.ts || new Date();
  const dateLabel = entry.inferredDate ? formatDate_(entry.inferredDate) : "";
  const row = [
    ts,
    entry.auditSheetName || "",
    dateLabel,
    entry.sourceFileName || "",
    entry.sourceFileId || "",
    entry.totalReservations || 0,
    entry.matched || 0,
    entry.missing || 0
  ];
  const next = sh.getLastRow() + 1;
  sh.getRange(next, 1, 1, headers.length).setValues([row]);
  sh.getRange(next, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange(next, 1, 1, headers.length).setFontFamily("Roboto").setFontSize(9).setFontColor(THEME.text);

  if (entry.missingList && entry.missingList.length) {
    const start = next + 1;
    sh.getRange(start, 1, 1, 6).setValues([["Faltantes (Localizador)", "", "", "", "", ""]])
      .setFontWeight("bold")
      .setBackground("#EFEFEF");
    const rows = entry.missingList.map(r => [
      r.localizador || r.localizadorRaw || "",
      r.situacao || "",
      r.tipoPagamento || "",
      r.valor || "",
      r.idViagem || "",
      r.codigoPms || ""
    ]);
    sh.getRange(start + 1, 1, rows.length, 6).setValues(rows);
  }
}

/** =======================
 *  Render auditoria (layout premium + consistência)
 *  ======================= */
function renderAuditSheet_(sheet, normRows, auditDate, auditorName, reportMeta) {
  const blockCount = Math.max(normRows.length, 1);
  const minRowsNeeded = Math.max(
    LAYOUT.freezeRows + 10,
    LAYOUT.startRow + blockCount * LAYOUT.blockHeight + 50
  );

  ensureSheetSize_(sheet, minRowsNeeded, LAYOUT.cols);
  sheet.clear();
  sheet.setHiddenGridlines(true);
  ensureSheetSize_(sheet, minRowsNeeded, LAYOUT.cols);

  sheet.getRange(1, 1, minRowsNeeded, LAYOUT.cols).setFontFamily("Roboto").setFontColor(THEME.text);

  const widths = [27,110,140,110,154,98,98,250,162,127,151,169,169];
  for (let c = 1; c <= widths.length; c++) sheet.setColumnWidth(c, widths[c - 1]);

  const summary = computeAuditSummary_(normRows);
  buildTopHeader_(sheet, auditDate, auditorName, summary, reportMeta);
  sheet.setFrozenRows(0);

  // Espaço antes das pilhas
  sheet.setRowHeight(LAYOUT.startRow - 1, 25);

  if (!normRows.length) {
    sheet.getRange(LAYOUT.startRow, 1).setValue("Sem reservas detectadas no relatório.").setFontColor(THEME.subtext);
    return;
  }

  const totalRows = normRows.length * LAYOUT.blockHeight;
  ensureSheetSize_(sheet, LAYOUT.startRow + totalRows + 20, LAYOUT.cols);

  const values = Array.from({ length: totalRows }, () => Array(LAYOUT.cols).fill(""));
  const headers = [
    "Sistema",
    "Nº Localizador",
    "Data da Criação",
    "Origem",
    "Check-in",
    "Check-out",
    "Hóspede Titular",
    "Tarifário",
    "Total (R$)",
    "Pagamento",
    "Status",
    "Observações"
  ];

  for (let i = 0; i < normRows.length; i++) {
    const base = i * LAYOUT.blockHeight;
    const r = normRows[i];

    // Header row (por reserva)
    values[base][0] = i + 1;
    for (let h = 0; h < headers.length; h++) values[base][h + 1] = headers[h];

    // Omni row
    values[base + 1][1]  = "Omnibees";
    values[base + 1][2]  = r[0];
    values[base + 1][3]  = r[2];
    values[base + 1][4]  = r[3];
    values[base + 1][5]  = r[4];
    values[base + 1][6]  = r[5];
    values[base + 1][7]  = r[6];
    values[base + 1][8]  = r[8];
    values[base + 1][9]  = r[9];
    values[base + 1][11] = r[1];

    const obsParts = [];
    if (r[10]) obsParts.push(`Aptos: ${r[10]}`);
    if (r[11]) obsParts.push(`Flags: ${r[11]}`);
    values[base + 1][12] = obsParts.join(" | ");

    // PMS row
    values[base + 2][1] = "Hotelflow";

    // Checks row
    values[base + 3][1] = "Auditoria";
  }

  sheet.getRange(LAYOUT.startRow, 1, totalRows, LAYOUT.cols).setValues(values);
  sheet.getRange(LAYOUT.startRow, 1, totalRows, LAYOUT.cols)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setFontColor(THEME.tableText)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  // Observações centralizado
  sheet.getRange(LAYOUT.startRow, 13, totalRows, 1)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  // Coluna A: branca, sem bordas, número alinhado à direita
  sheet.getRange(LAYOUT.startRow, 1, totalRows, 1)
    .setBackground("#FFFFFF")
    .setBorder(false, false, false, false, false, false)
    .setHorizontalAlignment("right")
    .setVerticalAlignment("middle")
    .setFontColor("#000000")
    .setFontWeight("bold");

  const heavy = normRows.length > UX.maxBlocksSoftLimit;
  applyBlocksStyling_(sheet, normRows.length, LAYOUT.startRow, { heavy });
  applyValidations_(sheet, normRows.length, LAYOUT.startRow);
  applyFormats_(sheet, normRows.length, LAYOUT.startRow);
  applyConditionalFormatting_(sheet, normRows.length, LAYOUT.startRow, { heavy });

  // Proteção leve (opcional): travar linha Omnibees e cabeçalhos do topo
  // (mantém editável PMS + checks). Ajuste conforme política do seu time.
  applyAuditProtections_(sheet, normRows.length, LAYOUT.startRow);
}

function buildTopHeader_(sheet, auditDate, auditorName, summary, reportMeta) {
  sheet.getRange(1, 1, 5, LAYOUT.cols).breakApart();
  // Header inicia na coluna B (linhas 1..3)
  sheet.getRange(1, 2, 3, LAYOUT.cols - 1).setBackground("#B7B7B7");
  sheet.setRowHeights(1, 1, 26);
  sheet.setRowHeights(2, 1, 26);
  sheet.setRowHeights(3, 1, 26);
  sheet.setRowHeights(4, 1, 8);
  sheet.setRowHeights(5, 1, 8);

  // Título mesclado (linhas 1..3)
  const title = sheet.getRange(1, 2, 3, 7);
  title.merge();
  title.setValue("Relatório de Auditoria Noturna")
    .setFontSize(18)
    .setFontWeight("bold")
    .setFontColor(THEME.text)
    .setFontFamily("Georgia")
    .setVerticalAlignment("middle");

  // Informações agrupadas à direita (linhas 1..3)
  const auditors = getConfigAuditors_();
  const auditorValue = String(auditorName || "").trim() || auditors[0] || "";
  const periodLabel = (reportMeta && reportMeta.reportPeriodLabel) ? reportMeta.reportPeriodLabel : "";
  const reportDateLabel = (reportMeta && reportMeta.reportGeneratedAt) ? formatDateTime_(reportMeta.reportGeneratedAt) : "";

  // Linha 1
  sheet.getRange(1, 9).setValue("Auditor Responsável:").setFontWeight("bold");
  sheet.getRange(1, 10).setValue(auditorValue);
  sheet.getRange(1, 11).setValue("Resp. Validação:").setFontWeight("bold");
  sheet.getRange(1, 12).setValue("");
  // Linha 2
  sheet.getRange(2, 9).setValue("Período:").setFontWeight("bold");
  sheet.getRange(2, 10).setValue(periodLabel);
  sheet.getRange(2, 11).setValue("Status:").setFontWeight("bold");
  sheet.getRange(2, 12).setValue("Não Validado");
  // Linha 3
  sheet.getRange(3, 9).setValue("Data do Relatório:").setFontWeight("bold");
  sheet.getRange(3, 10).setValue(reportDateLabel);
  sheet.getRange(3, 11).setValue("Data Validação:").setFontWeight("bold");
  sheet.getRange(3, 12).setValue("");

  sheet.getRange(1, 9, 3, 4)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  const auditorCell = sheet.getRange(1, 10);
  if (auditors && auditors.length) {
    auditorCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(auditors, true).build());
  }
  sheet.getRange(2, 12).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(LISTS.auditStatus, false).build()
  );
  sheet.getRange(3, 12).setNumberFormat("dd/MM/yyyy HH:mm");

  sheet.getRange(6, 1, Math.max(sheet.getMaxRows() - 5, 1), LAYOUT.cols).setBackground(THEME.bg);
}

function applyBlocksStyling_(sheet, blockCount, startRow, opts) {
  opts = opts || {};
  const heavy = !!opts.heavy;
  const bh = LAYOUT.blockHeight;
  const totalRows = blockCount * bh;

  sheet.getRange(startRow, 1, totalRows, 1).breakApart();
  sheet.getRange(startRow, 13, totalRows, 1).breakApart();

  for (let i = 0; i < blockCount; i++) {
    const top = startRow + i * bh;
    sheet.setRowHeight(top, 25);       // header
    sheet.setRowHeight(top + 1, 25);   // omni
    sheet.setRowHeight(top + 2, 25);   // pms
    sheet.setRowHeight(top + 3, 25);   // checks
    sheet.setRowHeight(top + 4, 25);   // spacer
  }

  for (let i = 0; i < blockCount; i++) {
    const top = startRow + i * bh;

    const headerRow = sheet.getRange(top, 2, 1, LAYOUT.cols - 1);
    const omniRow  = sheet.getRange(top + 1, 2, 1, LAYOUT.cols - 1);
    const pmsRow   = sheet.getRange(top + 2, 2, 1, LAYOUT.cols - 1);
    const checkRow = sheet.getRange(top + 3, 2, 1, LAYOUT.cols - 1);

    headerRow.setBackground(THEME.tableHeader)
      .setFontWeight("bold")
      .setFontSize(10)
      .setFontColor("#000000")
      .setFontFamily("Roboto")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    // borda superior grossa branca + inferior fina cinza escuro
    headerRow.setBorder(true, true, null, null, null, null, "#6E6E73", SpreadsheetApp.BorderStyle.SOLID);
    headerRow.setBorder(true, null, null, null, null, null, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_THICK);
    // separadores brancos internos (espessura 3) apenas no header (B..M)
    headerRow.setBorder(null, null, null, null, true, false, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_THICK);
    // Observações no header centralizado
    sheet.getRange(top, 13).setHorizontalAlignment("center").setVerticalAlignment("middle");

    omniRow.setBackground(THEME.omniRow).setFontColor(THEME.tableText).setFontFamily("Roboto");
    pmsRow.setBackground(THEME.pmsRow).setFontColor(THEME.tableText).setFontFamily("Roboto");
    checkRow.setBackground(THEME.checkRow).setFontSize(10).setFontColor(THEME.tableText).setFontFamily("Roboto");

    // Sem bordas laterais nas linhas abaixo do cabeçalho (B..M) - apenas Omni/PMS
    const dataRange = sheet.getRange(top + 1, 2, 2, LAYOUT.cols - 1);
    dataRange.setBorder(true, true, false, false, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

    // Observações unem Omni + PMS + Auditoria
    const obsRange = sheet.getRange(top + 1, 13, 3, 1);
    if (!heavy) obsRange.merge();
    obsRange.setBackground(THEME.obsCol)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setFontFamily("Georgia")
      .setFontColor("#000000");
    // Borda esquerda branca grossa na coluna de Observações
    obsRange.setBorder(null, true, null, null, null, null, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Sistema: Omnibees + Hotelflow em azul/underline, Auditoria em preto
    sheet.getRange(top + 1, 2, 2, 1)
      .setFontColor(THEME.tableText)
      .setFontWeight("bold")
      .setFontLine("underline")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    sheet.getRange(top + 3, 2, 1, 1)
      .setFontColor("#000000")
      .setFontWeight("bold")
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");

    // Status com aparência de pílula: centralizado (apenas coluna Status)
    sheet.getRange(top + 1, 12, 2, 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Linha de auditoria: borda superior branca grossa + inferior cinza clara fina
    checkRow.setBorder(false, false, false, false, false, false);
    checkRow.setBorder(true, null, null, null, null, null, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_THICK);
    checkRow.setBorder(null, true, null, null, null, null, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);
    // reforça a linha inferior usando a borda superior do espaçador
    const spacerRow = sheet.getRange(top + 4, 2, 1, LAYOUT.cols - 1);
    spacerRow.setBorder(true, null, null, null, null, null, "#B7B7B7", SpreadsheetApp.BorderStyle.SOLID);

    if (!heavy) {
      sheet.getRange(top + 1, 8, 2, 1).setWrap(true);
      sheet.getRange(top + 1, 10, 2, 1).setWrap(true);
    }
  }
}

function applyValidations_(sheet, blockCount, startRow) {
  const bh = LAYOUT.blockHeight;
  const dvStatus  = SpreadsheetApp.newDataValidation()
    .requireValueInList(LISTS.resStatusDropdown, true)
    .setAllowInvalid(false)
    .build();
  const dvPayment = SpreadsheetApp.newDataValidation()
    .requireValueInList(LISTS.paymentsDropdown, true)
    .setAllowInvalid(false)
    .build();
  const dvChecks  = SpreadsheetApp.newDataValidation().requireValueInList(LISTS.checks, true).build();

  for (let i = 0; i < blockCount; i++) {
    const top = startRow + i * bh;
    const omni = top + 1;
    const pms  = top + 2;
    const check = top + 3;

    const statusRange = sheet.getRange(omni, 12, 2, 1);
    statusRange.setDataValidation(dvStatus);

    const paymentRange = sheet.getRange(omni, 11, 2, 1);
    paymentRange.setDataValidation(dvPayment);

    sheet.getRange(check, 3, 1, 11).setDataValidation(dvChecks);
  }
}

function applyFormats_(sheet, blockCount, startRow) {
  const totalRows = blockCount * LAYOUT.blockHeight;
  sheet.getRange(startRow, 4, totalRows, 1).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(startRow, 6, totalRows, 1).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(startRow, 7, totalRows, 1).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(startRow, 10, totalRows, 1).setNumberFormat('"R$" #,##0.00');
  sheet.getRange(startRow, 3, totalRows, 1).setNumberFormat("@");
  sheet.getRange(startRow, 2, totalRows, 1).setHorizontalAlignment("left");
}

function applyConditionalFormatting_(sheet, blockCount, startRow, opts) {
  opts = opts || {};
  const heavy = !!opts.heavy;

  const endRow = startRow + blockCount * LAYOUT.blockHeight - 1;
  const rules = [];

  const checksRange = sheet.getRange(startRow, 3, endRow - startRow + 1, 11);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("✅")
      .setBackground(THEME.ok)
      .setRanges([checksRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("❌")
      .setBackground(THEME.bad)
      .setRanges([checksRange])
      .build()
  );

  const statusRange = sheet.getRange(startRow, 12, endRow - startRow + 1, 1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Confirmado")
      .setFontColor(THEME.statusTextOk)
      .setRanges([statusRange])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Cancelada")
      .setFontColor(THEME.statusTextBad)
      .setRanges([statusRange])
      .build()
  );
  if (!heavy) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Alterada")
        .setFontColor(THEME.statusTextWarn)
        .setRanges([statusRange])
        .build()
    );
  }

  sheet.setConditionalFormatRules(rules);
}

/**
 * Proteções: trava header topo e linha "Omnibees" (dados imutáveis).
 * Mantém PMS + checks editáveis.
 */
function applyAuditProtections_(sheet, blockCount, startRow) {
  // Remove proteções anteriores dessa sheet (best-effort)
  try {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(p => {
      try {
        if (String(p.getDescription() || "").startsWith("AUDIT_LOCK")) p.remove();
      } catch (_) {}
    });
  } catch (_) {}

  // Protege topo (1..freezeRows)
  try {
    const top = sheet.getRange(1, 1, LAYOUT.freezeRows, LAYOUT.cols).protect();
    top.setDescription("AUDIT_LOCK_TOP");
    top.setWarningOnly(true); // warning-only para não bloquear totalmente
  } catch (_) {}

  // Protege cabeçalho e linha Omnibees por bloco
  const bh = LAYOUT.blockHeight;
  for (let i = 0; i < blockCount; i++) {
    const topRow = startRow + i * bh;
    const headerRow = topRow;
    const omniRow = topRow + 1;

    try {
      const h = sheet.getRange(headerRow, 2, 1, 12).protect();
      h.setDescription("AUDIT_LOCK_HEADER_ROW");
      h.setWarningOnly(true);
    } catch (_) {}

    try {
      const r = sheet.getRange(omniRow, 2, 1, 12).protect();
      r.setDescription("AUDIT_LOCK_OMNI_ROW");
      r.setWarningOnly(true);
    } catch (_) {}
  }
}

/** =======================
 *  DASHBOARD (KPI + quebras + links)
 *  ======================= */
function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const norm = ss.getSheetByName(UX.omniNormSheetName);
  if (!norm) throw new Error("OMNI_NORM não encontrado. Importe um relatório primeiro.");

  const data = norm.getDataRange().getValues();
  // espera: rows >= 4 com header em 3
  if (data.length < 4) throw new Error("OMNI_NORM está vazio. Importe um relatório primeiro.");

  const rows = data.slice(3).filter(r => String(r[0] || "").trim());
  const normRows = rows.map(r => ([
    r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[11], r[12]
  ]));

  upsertDashboard_(ss, {
    inferredAuditDate: new Date(),
    normRows,
    sourceFileName: "(manual refresh)",
    sourceFileId: "",
    importSource: "MANUAL_REFRESH",
    createdAuditSheetName: "",
    createdAuditSheetId: ss.getId(),
    sourceSheetName: ""
  });

  SpreadsheetApp.getUi().alert("DASHBOARD atualizado.");
}

function upsertDashboard_(ss, ctx) {
  const sh = upsertSheet_(ss, UX.dashboardSheetName);
  sh.clear();
  sh.setHiddenGridlines(true);

  const normRows = ctx.normRows || [];
  const stats = computeDashboardStats_(normRows);

  ensureSheetSize_(sh, 80, 14);

  // Column widths (layout limpo)
  const widths = [40,160,160,160,160,30,160,160,160,160,30,200,200,200];
  for (let c = 1; c <= widths.length; c++) sh.setColumnWidth(c, widths[c - 1]);

  // Header band
  sh.getRange(1, 1, 6, 14).setBackground(THEME.headerBand);
  sh.getRange(1, 1, 6, 14).setFontFamily("Roboto").setFontColor(THEME.text);

  sh.getRange(1, 1, 2, 10).merge()
    .setValue("DASHBOARD — Auditoria Noturna")
    .setFontSize(18).setFontWeight("bold");

  sh.getRange(3, 1, 1, 10).merge()
    .setValue("KPIs e quebras para controle e gerenciamento. Fonte: OMNI_NORM (última importação).")
    .setFontSize(10).setFontColor(THEME.subtext);

  // Meta panel
  const panel = sh.getRange(1, 11, 4, 4);
  panel.setBackground(THEME.bg)
    .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID)
    .setFontSize(9)
    .setFontColor(THEME.subtext);

  sh.getRange(1, 11).setValue("Última fonte:").setFontWeight("bold");
  sh.getRange(2, 11).setValue("Arquivo:").setFontWeight("bold");
  sh.getRange(3, 11).setValue("Aba origem:").setFontWeight("bold");
  sh.getRange(4, 11).setValue("Gerado em:").setFontWeight("bold");

  sh.getRange(1, 12, 1, 3).merge().setValue(ctx.importSource || "");
  sh.getRange(2, 12, 1, 3).merge().setValue(ctx.sourceFileName || "");
  sh.getRange(3, 12, 1, 3).merge().setValue(ctx.sourceSheetName || "");
  sh.getRange(4, 12, 1, 3).merge().setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm");

  // Links rápidos (removido)

  // KPI cards
  const cardsTop = 8;

  buildKpiCard_(sh, cardsTop, 1, 4, "Reservas", stats.totalReservations, "Total no OMNI_NORM");
  buildKpiCard_(sh, cardsTop, 5, 4, "Total (R$)", stats.totalAmount, "Somatório do Total (R$)", '"R$" #,##0.00');
  buildKpiCard_(sh, cardsTop, 9, 4, "Check-ins (min..max)", stats.checkinRangeLabel, "Intervalo de check-in");
  buildKpiCard_(sh, cardsTop, 13, 2, "Canceladas", stats.statusCounts.Cancelada || 0, "Status Omni");

  buildKpiCard_(sh, cardsTop + 5, 1, 4, "Confirmadas", stats.statusCounts.Confirmado || 0, "Status Omni");
  buildKpiCard_(sh, cardsTop + 5, 5, 4, "Alteradas", stats.statusCounts.Alterada || 0, "Status Omni");
  buildKpiCard_(sh, cardsTop + 5, 9, 4, "Flags (reservas)", stats.flaggedReservations, "Com Flags (auto)");
  buildKpiCard_(sh, cardsTop + 5, 13, 2, "Aptos múltiplos", stats.multiAptoReservations, "APTOS_MULTIPLOS");

  // Quebras: Origem / Status / Top Tarifários
  const tTop = cardsTop + 11;
  buildSmallTableSection_(sh, tTop, 1, "Origem (Top)", stats.originTop, ["Origem", "Qtd"]);
  buildSmallTableSection_(sh, tTop, 6, "Status (Omni)", stats.statusTop, ["Status", "Qtd"]);
  buildSmallTableSection_(sh, tTop, 11, "Tarifário (Top)", stats.tariffTop, ["Tarifário", "Qtd"]);

  // Final polish
  sh.getRange(1, 1, 80, 14).setFontFamily("Roboto").setFontColor(THEME.text);

  // Freeze topo do dashboard
  sh.setFrozenRows(6);
}

function buildKpiCard_(sh, topRow, leftCol, widthCols, title, value, subtitle, numberFormat) {
  const height = 4;
  const rng = sh.getRange(topRow, leftCol, height, widthCols);
  rng.setBackground(THEME.bg)
    .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  const titleCell = sh.getRange(topRow, leftCol, 1, widthCols).merge();
  titleCell.setValue(title).setFontSize(10).setFontColor(THEME.subtext).setFontWeight("bold");

  const valueCell = sh.getRange(topRow + 1, leftCol, 2, widthCols).merge();
  valueCell.setValue(value).setFontSize(20).setFontWeight("bold");
  if (numberFormat) valueCell.setNumberFormat(numberFormat);

  const subCell = sh.getRange(topRow + 3, leftCol, 1, widthCols).merge();
  subCell.setValue(subtitle || "").setFontSize(9).setFontColor(THEME.subtext);
}

function buildSmallTableSection_(sh, topRow, leftCol, title, rows, headers) {
  rows = rows || [];
  headers = headers || ["Item", "Qtd"];

  sh.getRange(topRow, leftCol, 1, 4).merge()
    .setValue(title)
    .setFontWeight("bold")
    .setFontSize(10)
    .setFontColor(THEME.text);

  const head = sh.getRange(topRow + 1, leftCol, 1, 4);
  head.setBackground(THEME.tableHeader)
    .setFontWeight("bold")
    .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(topRow + 1, leftCol, 1, 2).setValue(headers[0]);
  sh.getRange(topRow + 1, leftCol + 2, 1, 2).setValue(headers[1]);

  const maxRows = 10;
  const bodyRows = Math.max(1, Math.min(rows.length, maxRows));
  const body = Array.from({ length: bodyRows }, (_, i) => {
    const r = rows[i] || ["", 0];
    return [r[0], "", r[1], ""];
  });

  const bodyRange = sh.getRange(topRow + 2, leftCol, bodyRows, 4);
  bodyRange.setValues(body)
    .setFontSize(9)
    .setBorder(true, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  // alinhamento
  sh.getRange(topRow + 2, leftCol + 2, bodyRows, 2).setHorizontalAlignment("right");
}

function computeDashboardStats_(normRows) {
  const totalReservations = normRows.length;
  let totalAmount = 0;

  let minCheckin = null, maxCheckin = null;

  const statusCounts = {};
  const originCounts = {};
  const tariffCounts = {};

  let flaggedReservations = 0;
  let multiAptoReservations = 0;

  for (const r of normRows) {
    const status = String(r[1] || "").trim() || "(vazio)";
    const origin = String(r[3] || "").trim() || "(vazio)";
    const tariff = String(r[8] || "").trim() || "(vazio)";
    const total = Number(r[9]) || 0;
    const flags = String(r[11] || "").trim();
    const aptos = String(r[10] || "").trim();

    totalAmount += total;

    statusCounts[status] = (statusCounts[status] || 0) + 1;
    originCounts[origin] = (originCounts[origin] || 0) + 1;
    tariffCounts[tariff] = (tariffCounts[tariff] || 0) + 1;

    if (flags) flaggedReservations++;
    if (aptos.includes(" + ") || flags.includes("APTOS_MULTIPLOS")) multiAptoReservations++;

    const ci = r[4] instanceof Date ? r[4] : null;
    if (ci) {
      if (!minCheckin || ci.getTime() < minCheckin.getTime()) minCheckin = ci;
      if (!maxCheckin || ci.getTime() > maxCheckin.getTime()) maxCheckin = ci;
    }
  }

  const originTop = topCounts_(originCounts, 10);
  const statusTop = topCounts_(statusCounts, 10);
  const tariffTop = topCounts_(tariffCounts, 10);

  const checkinRangeLabel = (minCheckin && maxCheckin)
    ? `${formatDate_(minCheckin)} → ${formatDate_(maxCheckin)}`
    : "-";

  // Normaliza nomes esperados (pra cards)
  const normalizedStatusCounts = {
    Confirmado: statusCounts["Confirmado"] || 0,
    Cancelada: statusCounts["Cancelada"] || 0,
    Alterada: statusCounts["Alterada"] || 0,
    ...statusCounts
  };

  return {
    totalReservations,
    totalAmount,
    checkinRangeLabel,
    statusCounts: normalizedStatusCounts,
    originTop,
    statusTop,
    tariffTop,
    flaggedReservations,
    multiAptoReservations
  };
}

function topCounts_(map, limit) {
  const arr = Object.keys(map || {}).map(k => [k, map[k]]);
  arr.sort((a, b) => b[1] - a[1]);
  return arr.slice(0, limit);
}

function formatDate_(d) {
  return Utilities.formatDate(d, UX.timezone, "dd/MM/yyyy");
}

function formatDateTime_(d) {
  return Utilities.formatDate(d, UX.timezone, "dd/MM/yyyy HH:mm");
}

function parseDateTimeFromText_(text) {
  const s = String(text || "");
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  const d = Number(m[1]);
  const mo = Number(m[2]);
  let y = Number(m[3]);
  if (y < 100) y = 2000 + y;
  const hh = Number(m[4] || 0);
  const mm = Number(m[5] || 0);
  const ss = Number(m[6] || 0);
  return new Date(y, mo - 1, d, hh, mm, ss);
}

function asDateTime_(v) {
  if (v instanceof Date) return v;
  const t = parseDateTimeFromText_(v);
  return t;
}

function extractReportMeta_(values, headerRowIndex) {
  const maxRow = Math.min(values.length, Math.max(headerRowIndex, 20));
  let reportGeneratedAt = null;
  let reportPeriodLabel = "";

  for (let r = 0; r < maxRow; r++) {
    const row = values[r] || [];
    for (let c = 0; c < row.length; c++) {
      const s = String(row[c] || "").trim();
      if (!s) continue;
      const low = s.toLowerCase();

      // Período (Check In / Check Out / Data Reserva)
      if (!reportPeriodLabel && (low.includes("data reserva de / até") || low.includes("check in de / até") || low.includes("check out de / até") || low.includes("período") || low.includes("periodo"))) {
        const dates = [];
        for (let k = 1; k <= 20 && c + k < row.length; k++) {
          const dt = asDateTime_(row[c + k]);
          if (dt) dates.push(dt);
        }
        if (dates.length < 2 && values[r + 1]) {
          const row2 = values[r + 1] || [];
          for (let k = 1; k <= 20 && c + k < row2.length; k++) {
            const dt = asDateTime_(row2[c + k]);
            if (dt) dates.push(dt);
          }
        }
        if (dates.length >= 2) {
          const d1 = dates[0];
          const d2 = dates[1];
          reportPeriodLabel = (formatDate_(d1) === formatDate_(d2))
            ? formatDate_(d1)
            : `${formatDate_(d1)} a ${formatDate_(d2)}`;
        }
      }

      // Relatório gerado
      if (!reportGeneratedAt && (low.includes("relatório gerado") || low.includes("relatorio gerado") || low.includes("gerado"))) {
        let dt = asDateTime_(s);
        if (!dt) {
          for (let k = 1; k <= 20 && c + k < row.length; k++) {
            dt = asDateTime_(row[c + k]);
            if (dt) break;
          }
        }
        if (dt) reportGeneratedAt = dt;
      }
    }
  }

  return { reportGeneratedAt, reportPeriodLabel };
}

function computeAuditSummary_(rows) {
  const totalRows = rows.length;
  let totalAmount = 0;
  let minCheckin = null;
  let maxCheckin = null;
  const statusCounts = { Confirmado: 0, Cancelada: 0, Alterada: 0 };

  for (const r of rows) {
    totalAmount += Number(r[9]) || 0;

    const st = String(r[1] || "").trim();
    if (st) {
      if (st === "🟢") statusCounts.Confirmado++;
      else if (st === "🔴") statusCounts.Cancelada++;
      else if (st === "🟡") statusCounts.Alterada++;
      else statusCounts[st] = (statusCounts[st] || 0) + 1;
    }

    const ci = r[4] instanceof Date ? r[4] : null;
    if (ci) {
      if (!minCheckin || ci.getTime() < minCheckin.getTime()) minCheckin = ci;
      if (!maxCheckin || ci.getTime() > maxCheckin.getTime()) maxCheckin = ci;
    }
  }

  const checkinRangeLabel = (minCheckin && maxCheckin)
    ? `${formatDate_(minCheckin)} a ${formatDate_(maxCheckin)}`
    : "-";

  const statusLabel = `C:${statusCounts.Confirmado || 0}  •  X:${statusCounts.Cancelada || 0}  •  A:${statusCounts.Alterada || 0}`;

  return {
    totalRows,
    totalAmount,
    checkinRangeLabel,
    statusLabel
  };
}

function formatBRL_(n) {
  const v = Number(n) || 0;
  const s = v.toFixed(2).replace(".", ",");
  return "R$ " + s;
}

/** =======================
 *  AUDIT_LOG
 *  ======================= */
function writeAuditLog_(ss, entry) {
  const sh = upsertSheet_(ss, UX.auditLogSheetName);
  sh.setHiddenGridlines(true);

  const headers = [
    "Timestamp",
    "Timestamp (America/Sao_Paulo)",
    "Auditor",
    "Usuário",
    "Fonte",
    "Arquivo",
    "FileId",
    "Aba Auditoria",
    "Reservas",
    "Duração (ms)"
  ];

  ensureSheetSize_(sh, Math.max(2, sh.getLastRow() || 2), headers.length);

  const needsHeader = (() => {
    if (sh.getLastRow() === 0) return true;
    const row = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    return headers.some((h, i) => String(row[i] || "").trim() !== h);
  })();

  if (needsHeader) {
    ensureSheetSize_(sh, 2, headers.length);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(THEME.tableHeader)
      .setFontWeight("bold")
      .setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    sh.setFrozenRows(1);

    const widths = [160,220,160,220,160,240,260,160,90,110];
    for (let c = 1; c <= widths.length; c++) sh.setColumnWidth(c, widths[c - 1]);
  }

  const ts = entry.ts || new Date();
  const tsSp = Utilities.formatDate(ts, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
  const row = [
    ts,
    `${tsSp} (${UX.timezone})`,
    entry.auditorName || "",
    entry.user || "",
    entry.importSource || "",
    entry.sourceFileName || "",
    entry.sourceFileId || "",
    entry.createdAuditSheetName || "",
    entry.reservations || 0,
    entry.durationMs || 0
  ];

  const next = sh.getLastRow() + 1;
  sh.getRange(next, 1, 1, headers.length).setValues([row]);
  sh.getRange(next, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange(next, 1, 1, headers.length).setFontFamily("Roboto").setFontSize(9).setFontColor(THEME.text);

  // sem link OMNI_NORM
}

/** =======================
 *  Utilitários: sizing
 *  ======================= */
function ensureSheetSize_(sheet, minRows, minCols) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < minRows) sheet.insertRowsAfter(maxRows, minRows - maxRows);

  const maxCols = sheet.getMaxColumns();
  if (maxCols < minCols) sheet.insertColumnsAfter(maxCols, minCols - maxCols);
}

/** =======================
 *  Lock (evita concorrência)
 *  ======================= */
function withImportLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Já existe uma importação em andamento. Aguarde finalizar e tente novamente.");
  }
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** =======================
 *  Dedupe de importação (anti-duplicidade)
 *  ======================= */
function computeImportSignature_(fileId, fileName) {
  let md5 = "";
  let size = "";
  let mod = "";
  let name = String(fileName || "");
  try {
    const meta = Drive.Files.get(fileId);
    md5 = meta.md5Checksum || "";
    size = meta.fileSize || meta.size || "";
    mod = meta.modifiedDate || meta.modifiedTime || "";
    name = meta.title || meta.name || name;
  } catch (_) {}

  const base = md5 ? ("md5:" + md5) : ("id:" + String(fileId || ""));
  const signature = [base, size, mod].filter(Boolean).join("|") || base;
  return { signature, fileName: name };
}

function loadImportHistory_() {
  try {
    const raw = PropertiesService.getDocumentProperties().getProperty("IMPORT_HISTORY");
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return obj && typeof obj === "object" ? obj : {};
  } catch (_) {
    return {};
  }
}

function saveImportHistory_(hist) {
  try {
    PropertiesService.getDocumentProperties().setProperty("IMPORT_HISTORY", JSON.stringify(hist || {}));
  } catch (_) {}
}

function pruneImportHistory_(hist, nowMs) {
  const windowMs = (UX.importDedupeMinutes || 30) * 60 * 1000;
  const maxEntries = UX.importDedupeMaxEntries || 50;
  const entries = Object.entries(hist || {});
  const fresh = entries.filter(([, v]) => v && v.ts && (nowMs - v.ts) <= (windowMs * 6));
  fresh.sort((a, b) => (b[1].ts || 0) - (a[1].ts || 0));
  const trimmed = fresh.slice(0, maxEntries);
  const out = {};
  trimmed.forEach(([k, v]) => { out[k] = v; });
  return out;
}

function beginImportDedupe_(signature, fileName) {
  if (!signature) return { skip: false };
  const now = Date.now();
  let hist = loadImportHistory_();
  hist = pruneImportHistory_(hist, now);

  const existing = hist[signature];
  const windowMs = (UX.importDedupeMinutes || 30) * 60 * 1000;
  if (existing && existing.ts && (now - existing.ts) < windowMs) {
    const when = formatDateTime_(new Date(existing.ts));
    const sheetName = existing.sheetName || "";
    const msg = `Importação ignorada: arquivo já processado em ${when}.` + (sheetName ? ` Aba: ${sheetName}.` : "");
    return { skip: true, message: msg };
  }

  hist[signature] = {
    ts: now,
    fileName: String(fileName || ""),
    status: "in_progress"
  };
  saveImportHistory_(hist);
  return { skip: false };
}

function finalizeImportDedupe_(signature, sheetName, error) {
  if (!signature) return;
  let hist = loadImportHistory_();
  if (!hist[signature]) return;

  if (error) {
    delete hist[signature];
  } else {
    hist[signature] = {
      ts: Date.now(),
      fileName: hist[signature].fileName || "",
      sheetName: String(sheetName || ""),
      status: "done"
    };
  }
  saveImportHistory_(hist);
}

/** =======================
 *  Drive / conversão
 *  ======================= */
function assertDriveAdvancedServiceEnabled_() {
  if (typeof Drive === "undefined" || !Drive || !Drive.Files) {
    throw new Error(
      "Drive API Avançada não habilitada.\n\n" +
      "Apps Script → Serviços avançados do Google → ative 'Drive API'.\n" +
      "E habilite 'Google Drive API' no Google Cloud do projeto, se solicitado."
    );
  }
}

function convertExcelToGoogleSheet_(fileId, originalName) {
  const resource = {
    title: `[CONVERTIDO] ${originalName}`,
    mimeType: MimeType.GOOGLE_SHEETS,
  };

  const converted = Drive.Files.copy(resource, fileId, { convert: true });
  if (!converted || !converted.id) {
    throw new Error("Falha ao converter o Excel para Google Sheets (Drive.Files.copy retornou vazio).");
  }
  return converted.id;
}

function openSpreadsheetWithRetry_(spreadsheetId) {
  let lastErr = null;
  for (let i = 0; i < UX.openRetryCount; i++) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      lastErr = e;
      Utilities.sleep(UX.openRetrySleepMs * (i + 1));
    }
  }
  throw new Error("Não consegui abrir a planilha convertida após tentativas. Erro: " + (lastErr && lastErr.message ? lastErr.message : lastErr));
}

function safeTrashFile_(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (_) {}
}

/** =======================
 *  Upload folder
 *  ======================= */
function resolveUploadFolder_(ss) {
  const parent = tryGetSpreadsheetParentFolder_(ss);
  if (parent) return parent;
  return ensureFolderInRoot_("_Auditoria_Omnibees_Uploads");
}

function tryGetSpreadsheetParentFolder_(ss) {
  try {
    const file = DriveApp.getFileById(ss.getId());
    const parents = file.getParents();
    if (parents && parents.hasNext()) return parents.next();
  } catch (e) {}
  return null;
}

function ensureFolderInRoot_(name) {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return root.createFolder(name);
}

function findLatestExcelFileInFolder_(folder) {
  const files = folder.getFiles();
  let latest = null;
  let latestTime = 0;

  while (files.hasNext()) {
    const f = files.next();
    const mt = f.getMimeType();
    const name = (f.getName() || "").toLowerCase();
    const extOk = name.endsWith(".xls") || name.endsWith(".xlsx");

    const isExcel =
      mt === MimeType.MICROSOFT_EXCEL ||
      mt === "application/vnd.ms-excel" ||
      mt === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    if (!isExcel && !extOk) continue;

    const t = f.getLastUpdated().getTime();
    if (t > latestTime) { latestTime = t; latest = f; }
  }
  return latest;
}

function extractDriveId_(text) {
  const s = String(text || "").trim();
  if (!s) return "";
  const m = s.match(/[-\w]{20,}/);
  return m ? m[0] : "";
}

function normalizeExcelMime_(fileName, mimeType) {
  const n = String(fileName || "").toLowerCase();
  if (n.endsWith(".xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (n.endsWith(".xls"))  return "application/vnd.ms-excel";
  return mimeType || "application/vnd.ms-excel";
}

function upsertSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** =======================
 *  Parsers / Normalização
 *  ======================= */
function findHeaderRowIndex_(values) {
  for (let r = 0; r < Math.min(values.length, 80); r++) {
    const row = values[r].map(v => String(v || "").trim());
    const hit = row.some(x => normalizeHeader_(x) === normalizeHeader_("Res. Nº"));
    if (hit) return r;
  }
  return -1;
}

function findColIndex_(headerRow, candidates) {
  const norm = headerRow.map(x => normalizeHeader_(x));
  for (const cand of candidates) {
    const c = normalizeHeader_(cand);
    const idx = norm.indexOf(c);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeader_(s) {
  return String(s || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function parseBRL_(v) {
  if (typeof v === "number") return v;
  const s0 = String(v || "").trim();
  if (!s0) return 0;

  // remove moeda e espaços
  let s = s0.replace(/brl/ig, "").replace(/r\$/ig, "").trim();
  // mantém dígitos e separadores
  s = s.replace(/[^\d.,-]/g, "");

  if (!s) return 0;

  // Heurística:
  // - se tem vírgula e ponto: assume ponto milhar e vírgula decimal
  // - se só vírgula: assume vírgula decimal
  // - se só ponto: assume ponto decimal
  if (s.includes(",") && s.includes(".")) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (s.includes(",") && !s.includes(".")) {
    s = s.replace(",", ".");
  }

  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function extractTitular_(hospede) {
  const s = String(hospede || "").trim();
  if (!s) return "";
  return s.split(";")[0].trim();
}

function unique_(arr) {
  const out = [];
  const seen = {};
  arr.forEach(v => {
    const s = String(v || "").trim();
    if (!s) return;
    if (!seen[s]) { seen[s] = true; out.push(s); }
  });
  return out;
}

function normalizeStatus_(s) {
  const t = String(s || "").trim().toLowerCase();
  if (t.startsWith("confirm")) return "Confirmado";
  if (t.startsWith("cancel"))  return "Cancelada";
  if (t.startsWith("alter"))   return "Alterada";
  return String(s || "").trim();
}

function normalizeOrigin_(s) {
  const t = String(s || "").trim().toLowerCase();
  if (!t) return "";
  if (t.includes("central")) return "Central de Reservas";
  if (t.includes("be mobile")) return "BE Mobile";
  if (t.includes("booking engine")) return "Booking Engine";
  if (t === "booking" || t.includes("booking.com")) return "Booking";
  if (t.includes("iterpec")) return "Iterpec";
  return String(s || "").trim();
}

function asDateOnly_(v) {
  if (!(v instanceof Date)) return v;
  return new Date(v.getFullYear(), v.getMonth(), v.getDate());
}

function coerceDate_(v) {
  if (v instanceof Date) return v;
  const t = parseDateTimeFromText_(v);
  if (t) return t;
  const n = Number(v);
  if (isFinite(n) && n > 20000 && n < 80000) {
    const ms = Math.round((n - 25569) * 86400 * 1000);
    return new Date(ms);
  }
  return null;
}

function normalizeLocalizador_(s) {
  const t = String(s || "").trim();
  if (!t) return "";
  return t.split("/")[0].trim();
}

function normalizeText_(s) {
  const t = String(s || "").trim().toLowerCase();
  if (!t) return "";
  return t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function isPaidStatus_(status) {
  const s = String(status || "").toLowerCase();
  if (!s) return false;
  if (s.includes("confirm")) return true;
  if (s.includes("aprov")) return true;
  if (s.includes("pago")) return true;
  if (s.includes("liquid")) return true;
  if (s.includes("baixa")) return true;
  if (s.includes("pend")) return false;
  if (s.includes("cancel")) return false;
  if (s.includes("estorn")) return false;
  if (s.includes("recus")) return false;
  return false;
}

function buildNiaraServiceUrl_(idViagem, idReserva) {
  const base = String(UX.niaraBaseUrl || "").trim();
  if (!base || !idViagem) return "";
  const idV = encodeURIComponent(String(idViagem || "").trim());
  const idR = String(idReserva || "").trim();
  if (idR) {
    return base + idV + "#serviceId=" + encodeURIComponent(idR);
  }
  return base + idV;
}

function buildPmsReservationUrl_(codigoPms) {
  const base = String(UX.pmsBaseUrl || "").trim();
  const code = String(codigoPms || "").trim();
  if (!base || !code) return "";
  return base + encodeURIComponent(code);
}

function buildBee2PayUrl_(idTransacao) {
  const base = String(UX.bee2payBaseUrl || "").trim();
  const id = String(idTransacao || "").trim();
  if (!base || !id) return "";
  return base + encodeURIComponent(id);
}

function buildBee2PayReservationUrl_(locator) {
  const base = String(UX.bee2payReservationBaseUrl || "").trim();
  const loc = String(locator || "").trim();
  if (!base || !loc) return "";
  return base + encodeURIComponent(loc);
}

function isBee2PayPaid_(statusCobranca, retorno, valor) {
  const st = String(statusCobranca || "").toLowerCase();
  const rt = String(retorno || "").toLowerCase();
  if (st.includes("autoriz")) return true;
  if (st.includes("aprov")) return true;
  if (st.includes("confirm")) return true;
  if (rt.includes("sucesso")) return true;
  const v = Number(valor) || 0;
  return v > 0 && !st.includes("estorn") && !st.includes("cancel");
}

function safeActiveUserEmail_() {
  try {
    const e = Session.getActiveUser().getEmail();
    return e || "";
  } catch (_) {
    return "";
  }
}

function promptAuditorName_(ui) {
  const list = getConfigAuditors_();
  const hint = list.length ? ("Auditores cadastrados: " + list.join(", ")) : "Nenhum auditor cadastrado.";
  const resp = ui.prompt(
    "Auditor responsável",
    "Informe o nome do auditor.\n" + hint,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  return (resp.getResponseText() || "").trim();
}
