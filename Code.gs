/*******************************
 * CONFIG
 *******************************/
const CONFIG = {
  DEFAULT_SOURCE: "ML Coleta",
  SHEET_NAMES: ["ML Coleta", "ML 1", "Shopee", "Magalu", "Essência do Brasil", "Amazon", "Flex/Vapt", "Próximos Dias"],

  HEADER_ROW: 1,
  DATA_START_ROW: 2,

  // Estrutura real (A..J)
  COLS: {
    ID: 1,
    CLIENTE: 2,
    PRODUTO: 3,
    QTD: 4,
    STATUS: 5,
    BIPADO: 6,
    ETIQUETAS: 7,
    ANDAMENTO: 8,
    PRODUTOS_PENDENTES: 9, // I
    QUANTIDADE: 10,        // J
    TODOS_NV_PROD: 13,     // M — Falta/Não Verificado
    TODOS_NV_QTD:  14,     // N
  },
};

/*******************************
 * WEB APP ENTRY
 *******************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Dashboard | Pedidos ML")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*******************************
 * API: GET DATA
 * source: "ML Coleta" | "ML 1" | "TODOS"
 *******************************/
function getPedidos(source) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const src = normalizeSource_(source);

    const sheetsToRead = (src === "TODOS")
      ? CONFIG.SHEET_NAMES
      : [src];

    const pedidos = [];
    const faltamItems = [];
    const todosNvItems = [];
    const perSheet = {};

    for (const name of sheetsToRead) {
      const out = readFromSheet_(ss, name);
      perSheet[name] = { pedidos: out.pedidos.length, faltam: out.faltamItems.length };

      pedidos.push(...out.pedidos);
      faltamItems.push(...out.faltamItems);
      todosNvItems.push(...out.todosNvItems);
    }

    // Ordena pedidos por ID (ajuda no visual e multi-item)
    pedidos.sort((a, b) => String(a.id).localeCompare(String(b.id)));

    // FaltamItems: mantém ordem da planilha (ou ordena por produto se quiser)
    // faltamItems.sort((a,b)=>String(a.produtoPendentes).localeCompare(String(b.produtoPendentes)));

    const tenhoItems = buildTenhoItems_();
    const kpis = buildKpis_(pedidos, faltamItems, todosNvItems);
    kpis.tenho = tenhoItems.reduce((s, it) => s + Number(it.quantidade || 0), 0);

    const _todayCheck = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const _storedDate = PropertiesService.getScriptProperties().getProperty("COLLECTION_TIMES_DATE") || "";
    if (_storedDate !== _todayCheck) {
      const _p = PropertiesService.getScriptProperties();
      _p.deleteProperty("MONITOR_PAUSED_ML_COLETA");
      _p.deleteProperty("MONITOR_PAUSED_ML_1");
    }

    const hasBlueCells = checkBlueCells_(ss, sheetsToRead);
    const hasReviewCells = checkReviewCells_(ss, sheetsToRead);
    const monitorSnapshot = {
      "ML Coleta": readMonitorSnapshot_("LAST_MONITOR_SNAPSHOT_ML_COLETA"),
      "ML 1": readMonitorSnapshot_("LAST_MONITOR_SNAPSHOT_ML_1")
    };
    const sProps = PropertiesService.getScriptProperties();
    const monitorPaused = {
      "ML Coleta": sProps.getProperty("MONITOR_PAUSED_ML_COLETA") === "true",
      "ML 1": sProps.getProperty("MONITOR_PAUSED_ML_1") === "true"
    };

    const _impRaw = PropertiesService.getScriptProperties().getProperty("IMPORT_RUNNING") || "";
    let _importRunning = null;
    if (_impRaw) {
      const _parts = _impRaw.split("|");
      const _ts    = Number(_parts[1] || 0);
      if (Date.now() - _ts < 30 * 60 * 1000) _importRunning = _parts[0];
    }

    return {
      ok: true,
      source: src,
      pedidos,
      faltamItems,
      todosNvItems,
      tenhoItems,
      kpis,
      collectionTimes: getCollectionTimes(),
      hasBlueCells,
      hasReviewCells,
      monitorSnapshot,
      monitorPaused,
      importRunning: _importRunning,
      meta: {
        updatedAt: new Date().toISOString(),
        pedidosRows: pedidos.length,
        faltamRows: faltamItems.length,
        perSheet,
      },
    };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

/*******************************
 * READ ONE SHEET
 * - pedidos: linhas com ID preenchido
 * - faltamItems: linhas com Produtos Pendentes preenchido (independente de ID)
 *******************************/
function readFromSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { pedidos: [], faltamItems: [], todosNvItems: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { pedidos: [], faltamItems: [], todosNvItems: [] };

  const isProximosDias = sheetName === "Próximos Dias";
  const lastCol = isProximosDias ? 15 : CONFIG.COLS.TODOS_NV_QTD; // 15 = col O (marketplace)

  const range = sh.getRange(
    CONFIG.DATA_START_ROW,
    1,
    lastRow - CONFIG.DATA_START_ROW + 1,
    lastCol
  );

  const values = range.getValues();

  const pedidos = [];
  const faltamItems = [];
  const todosNvItems = [];

  const _pendDates = (["Shopee","Magalu","Essência do Brasil","Amazon"].includes(sheetName))
    ? pendDates_load_()
    : null;

  for (const r of values) {
    // ====== FALTAM LIST (independente) ======
    const produtoPendentes = String(r[CONFIG.COLS.PRODUTOS_PENDENTES - 1] ?? "").trim();
    const quantidade = parseNumber_(r[CONFIG.COLS.QUANTIDADE - 1]);

    // Só entra se tiver produto pendente (o principal)
    const rowId = String(r[CONFIG.COLS.ID - 1] ?? "").trim();
    if (produtoPendentes !== "" && rowId !== "") {
      faltamItems.push({ id: rowId, produtoPendentes, quantidade, sheet: sheetName });
    }

    // ====== TODOS NV LIST (M/N) ======
    const todosNvProd = String(r[CONFIG.COLS.TODOS_NV_PROD - 1] ?? "").trim();
    const todosNvQtd  = parseNumber_(r[CONFIG.COLS.TODOS_NV_QTD - 1]);
    if (todosNvProd !== "") {
      todosNvItems.push({
        produtoPendentes: todosNvProd,
        quantidade: todosNvQtd,
        sheet: sheetName
      });
    }

    // ====== PEDIDOS (somente linhas com ID) ======
    const id = String(r[CONFIG.COLS.ID - 1] ?? "").trim();
    if (!id) continue;

    const cliente = String(r[CONFIG.COLS.CLIENTE - 1] ?? "").trim();
    const produto = String(r[CONFIG.COLS.PRODUTO - 1] ?? "").trim();

    const qtdRaw = r[CONFIG.COLS.QTD - 1];
    const qtd = (qtdRaw === "" || qtdRaw == null) ? "" : String(qtdRaw).trim();

    const status = normalizePick_(r[CONFIG.COLS.STATUS - 1], ["PENDENTE", "CONFIRMADO", "CANCELADO"]);
    const bipado = normalizePick_(r[CONFIG.COLS.BIPADO - 1], ["PRONTO", "NÃO FINALIZADO", "NAO FINALIZADO", "CANCELADO"]);
    const etiquetas = normalizePick_(r[CONFIG.COLS.ETIQUETAS - 1], ["IMPRESSA", "PARA IMPRIMIR", "CANCELADO"]);
    const andamento = String(r[CONFIG.COLS.ANDAMENTO - 1] ?? "").trim();

    pedidos.push({
      id, cliente, produto, qtd,
      status, bipado, etiquetas, andamento,
      sheet: sheetName,
      marketplace: isProximosDias ? String(r[14] ?? "").trim() : "",
      diasPendente: _pendDates ? pendDates_diasDesde_(_pendDates[id]) : null
    });
  }

  return { pedidos, faltamItems, todosNvItems };
}

/*******************************
 * KPI LOGIC
 * - Total/Pendentes/Confirmados/Cancelados: IDs únicos em pedidos
 * - Bipados: IDs com bipado PRONTO
 * - Etiquetas Impressas: IDs com etiquetas IMPRESSA
 * - Faltam: soma Quantidade da lista independente faltamItems
 *******************************/
function buildKpis_(pedidos, faltamItems, todosNvItems) {
  const idsTotal = new Set();
  const idsPendentes = new Set();
  const idsConfirmados = new Set();
  const idsCancelados = new Set();
  const idsBipadosPronto = new Set();
  const idsEtiquetasImpressa = new Set();

  for (const p of pedidos) {
    const id = String(p.id);
    idsTotal.add(id);

    if (p.status === "PENDENTE") idsPendentes.add(id);
    if (p.status === "CONFIRMADO") idsConfirmados.add(id);
    if (p.status === "CANCELADO") idsCancelados.add(id);

    if (p.bipado === "PRONTO") idsBipadosPronto.add(id);
    if (p.etiquetas === "IMPRESSA") idsEtiquetasImpressa.add(id);
  }

  let faltam = 0;
  for (const it of faltamItems) {
    faltam += Number(it.quantidade || 0);
  }

  let faltaNaoVerificado = 0;
  for (const it of (todosNvItems || [])) {
    faltaNaoVerificado += Number(it.quantidade || 0);
  }

  return {
    total: idsTotal.size,
    pendentes: idsPendentes.size,
    confirmados: idsConfirmados.size,
    cancelados: idsCancelados.size,
    bipadosPronto: idsBipadosPronto.size,
    etiquetasImpressas: idsEtiquetasImpressa.size,
    faltam: faltam,
    faltaNaoVerificado,
    parciais: (todosNvItems || []).length,
  };
}

/*******************************
 * HELPERS
 *******************************/
function normalizeSource_(source) {
  const s = String(source || "").trim();
  if (!s) return CONFIG.DEFAULT_SOURCE;

  const u = s.toUpperCase();
  if (u === "TODOS" || u === "ALL") return "TODOS";

  for (const name of CONFIG.SHEET_NAMES) {
    if (name.toUpperCase() === u) return name;
  }
  return CONFIG.DEFAULT_SOURCE;
}

function normalizePick_(value, allowedListUpper) {
  const s = String(value ?? "").trim();
  if (!s) return "—";

  const u = s.toUpperCase();
  const u2 = u.replace("NÃO", "NAO");

  for (const a of allowedListUpper) {
    const ax = a.toUpperCase().replace("NÃO", "NAO");
    if (u2 === ax) {
      return a.toUpperCase().replace("NAO", "NÃO");
    }
  }
  return u;
}

function parseNumber_(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;

  const s = String(v).trim().replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function checkBlueCells_(ss, sheetNames) {
  for (const name of sheetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    const bgs = sh.getRange(2, CONFIG.COLS.PRODUTO, lastRow - 1, 1).getBackgrounds();
    for (let i = 0; i < bgs.length; i++) {
      if (String(bgs[i][0] || "").toLowerCase() === "#0000ff") return true;
    }
  }
  return false;
}

function checkReviewCells_(ss, sheetNames) {
  for (const name of sheetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    const bgs = sh.getRange(2, CONFIG.COLS.CLIENTE, lastRow - 1, 1).getBackgrounds();
    for (let i = 0; i < bgs.length; i++) {
      const c = String(bgs[i][0] || "").toLowerCase();
      if (c === "#ff0000" || c === "#ffff00") return true;
    }
  }
  return false;
}

function readMonitorSnapshot_(key) {
  try {
    const v = PropertiesService.getScriptProperties().getProperty(key);
    return v ? JSON.parse(v) : null;
  } catch(e) { return null; }
}

/*******************************
 * COLLECTION TIMES
 *******************************/
function setUserSelecting(state) {
  const props = PropertiesService.getScriptProperties();
  if (state) {
    props.setProperty("USER_SELECTING", String(Date.now()));
  } else {
    props.deleteProperty("USER_SELECTING");
  }
}

function getImportStatus() {
  const raw = PropertiesService.getScriptProperties().getProperty("IMPORT_RUNNING") || "";
  if (!raw) return null;
  const parts = raw.split("|");
  const ts = Number(parts[1] || 0);
  if (Date.now() - ts > 30 * 60 * 1000) return null;
  return parts[0];
}

function getCollectionTimes() {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty("COLLECTION_TIMES");
    const storedDate = props.getProperty("COLLECTION_TIMES_DATE") || "";
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (storedDate !== today) return { "ML Coleta": "", "ML 1": "" };
    return stored ? JSON.parse(stored) : { "ML Coleta": "", "ML 1": "" };
  } catch(e) {
    return { "ML Coleta": "", "ML 1": "" };
  }
}

function saveCollectionTime(sheet, time) {
  try {
    const props = PropertiesService.getScriptProperties();
    const times = getCollectionTimes();
    times[String(sheet)] = String(time || "").trim();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    props.setProperty("COLLECTION_TIMES", JSON.stringify(times));
    props.setProperty("COLLECTION_TIMES_DATE", today);
    return { ok: true, times };
  } catch(err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/*******************************
 * SIMPLE TRIGGER: onEdit
 *******************************/
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (!CONFIG.SHEET_NAMES.includes(sheetName)) return;

  const col = range.getColumn();
  const row = range.getRow();

  if (row < CONFIG.DATA_START_ROW) return;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  const val = String(range.getValue() || "").trim().toUpperCase();

  // Status = "Confirmado" → limpa Andamento (col H)
  if (col === CONFIG.COLS.STATUS && val === "CONFIRMADO") {
    sheet.getRange(row, CONFIG.COLS.ANDAMENTO).clearContent();
  }

  // Etiquetas = "Cancelado" → define Status = "Cancelado" (col E)
  if (col === CONFIG.COLS.ETIQUETAS && val === "CANCELADO") {
    sheet.getRange(row, CONFIG.COLS.STATUS).setValue("Cancelado");
  }
}

/*******************************
 * GOOGLE CHAT API
 *******************************/
function callChatApi_(method, path, payloadObj) {
  const token = getServiceAccountToken_();

  const options = {
    method: method,
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  };

  if (payloadObj !== undefined && payloadObj !== null) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(payloadObj);
  }

  const url = "https://chat.googleapis.com/v1/" + path;
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("Google Chat API HTTP " + code + " - " + body);
  }

  return body ? JSON.parse(body) : {};
}

function enviarHorariosNoChat_() {
  const props = PropertiesService.getScriptProperties();

  const times = getCollectionTimes();

  const spacesColeta = String(props.getProperty("CHAT_SPACE_ML_COLETA") || "").trim();
  const spacesML1    = String(props.getProperty("CHAT_SPACE_ML_1")      || "").trim();

  const normalize = (s) => {
    if (!s) return s;
    return s.startsWith("spaces/") ? s : "spaces/" + s;
  };

  const splitSpaces = (raw) =>
    raw.split(/\r?\n/)
       .map(s => s.trim())
       .filter(Boolean);

  const sendTo = (spacesRaw, sheetName) => {
    if (!spacesRaw) return;
    const time = times[sheetName] || "—";
    const text = "🚚 *" + sheetName + "* — Hor\xe1rio da coleta: *" + time + "*";

    splitSpaces(spacesRaw).forEach(spaceId => {
      try {
        callChatApi_("post", normalize(spaceId) + "/messages", { text: text });
      } catch (err) {
        Logger.log("Erro ao enviar hor\xe1rio para " + spaceId + ": " + err.message);
      }
    });
  };

  sendTo(spacesColeta, "ML Coleta");
  sendTo(spacesML1,    "ML 1");
}

function enviarHorarioMLColeta() {
  try {
    const props = PropertiesService.getScriptProperties();
    const times = getCollectionTimes();
    const spacesRaw = String(props.getProperty("CHAT_SPACE_ML_COLETA") || "").trim();
    if (!spacesRaw) return { ok: false, error: "CHAT_SPACE_ML_COLETA não configurado." };
    const normalize = (s) => (!s ? s : s.startsWith("spaces/") ? s : "spaces/" + s);
    const time = times["ML Coleta"] || "—";

    const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const lastDate = props.getProperty("LAST_SENT_TIME_DATE_ML_COLETA") || "";
    const lastTime = props.getProperty("LAST_SENT_TIME_ML_COLETA") || "";
    const changed  = lastDate === today && lastTime && lastTime !== time;

    let text = "🚚 *ML Coleta* — Horário da coleta: *" + time + "*";
    if (changed) text = "⚠️ *Atenção: horário alterado de " + lastTime + " para " + time + "*\n" + text;

    spacesRaw.split(/[\r\n,]+/).map(s => s.trim()).filter(Boolean).forEach(spaceId => {
      callChatApi_("post", normalize(spaceId) + "/messages", { text });
    });

    props.setProperty("LAST_SENT_TIME_ML_COLETA", time);
    props.setProperty("LAST_SENT_TIME_DATE_ML_COLETA", today);
    return { ok: true };
  } catch(err) {
    return { ok: false, error: String(err.message || err) };
  }
}

function enviarHorarioML1() {
  try {
    const props = PropertiesService.getScriptProperties();
    const times = getCollectionTimes();
    const spacesRaw = String(props.getProperty("CHAT_SPACE_ML_1") || "").trim();
    if (!spacesRaw) return { ok: false, error: "CHAT_SPACE_ML_1 não configurado." };
    const normalize = (s) => (!s ? s : s.startsWith("spaces/") ? s : "spaces/" + s);
    const time = times["ML 1"] || "—";

    const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const lastDate = props.getProperty("LAST_SENT_TIME_DATE_ML_1") || "";
    const lastTime = props.getProperty("LAST_SENT_TIME_ML_1") || "";
    const changed  = lastDate === today && lastTime && lastTime !== time;

    let text = "🚚 *ML 1* — Horário da entrega: *" + time + "*";
    if (changed) text = "⚠️ *Atenção: horário alterado de " + lastTime + " para " + time + "*\n" + text;

    spacesRaw.split(/[\r\n,]+/).map(s => s.trim()).filter(Boolean).forEach(spaceId => {
      callChatApi_("post", normalize(spaceId) + "/messages", { text });
    });

    props.setProperty("LAST_SENT_TIME_ML_1", time);
    props.setProperty("LAST_SENT_TIME_DATE_ML_1", today);
    return { ok: true };
  } catch(err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/*******************************
 * MONITOR: status horário no Chat
 *******************************/

function enviarStatusMonitor_(sheetName) {
  try {
    const props = PropertiesService.getScriptProperties();
    const times = getCollectionTimes();
    const collectionTime = times[sheetName] || "";
    if (!collectionTime) return;

    const [h, m] = collectionTime.split(":").map(Number);
    const now = new Date();
    const collectionDate = new Date(now);
    collectionDate.setHours(h, m, 0, 0);

    const closeTime = new Date(collectionDate.getTime() - 3 * 60 * 60 * 1000);
    const isOpen = now < closeTime;

    const label = sheetName === "ML 1" ? "Horário da entrega" : "Horário da coleta";
    const closedLabel = sheetName === "ML 1" ? "ML 1 Fechado" : "ML Coleta Fechado";
    const statusLine = isOpen ? "🟢 Marketplace Ainda em Aberto" : "🔴 " + closedLabel;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    const seenIds      = new Set();
    const pendingIds   = new Set();
    const confirmedIds = new Set();
    const canceledIds  = new Set();

    if (lastRow >= 2) {
      const n = lastRow - 1;
      const ids        = sh.getRange(2, 1, n, 1).getValues();
      const statusVals = sh.getRange(2, 5, n, 1).getValues();
      for (let i = 0; i < n; i++) {
        const id = String(ids[i][0] ?? "").trim();
        if (!id) continue;
        seenIds.add(id);
        const st = String(statusVals[i][0] ?? "").trim().toLowerCase();
        if (st === "pendente")   pendingIds.add(id);
        if (st === "confirmado") confirmedIds.add(id);
        if (st === "cancelado")  canceledIds.add(id);
      }
    }

    const dashUrl = String(props.getProperty("DASHBOARD_URL") || "").trim();

    const text =
      "📊 *" + sheetName + "* — Atualização\n" +
      statusLine + "\n" +
      "• Total: " + seenIds.size +
      " | Pendentes: " + pendingIds.size +
      " | Confirmados: " + confirmedIds.size +
      " | Cancelados: " + canceledIds.size + "\n" +
      "⏰ " + label + ": *" + collectionTime + "*" +
      (dashUrl ? "\n🔗 " + dashUrl : "");

    const spaceKey = sheetName === "ML 1" ? "CHAT_SPACE_ML_1" : "CHAT_SPACE_ML_COLETA";
    const spacesRaw = String(props.getProperty(spaceKey) || "").trim();
    if (!spacesRaw) return;

    const normalize = (s) => (!s ? s : s.startsWith("spaces/") ? s : "spaces/" + s);
    spacesRaw.split(/[\r\n,]+/).map(s => s.trim()).filter(Boolean).forEach(spaceId => {
      try { callChatApi_("post", normalize(spaceId) + "/messages", { text }); }
      catch (e) { Logger.log("Monitor Chat err (" + sheetName + "): " + e.message); }
    });

    const snapshotKey = sheetName === "ML 1"
      ? "LAST_MONITOR_SNAPSHOT_ML_1"
      : "LAST_MONITOR_SNAPSHOT_ML_COLETA";
    props.setProperty(snapshotKey, JSON.stringify({
      total:      seenIds.size,
      pendentes:  pendingIds.size,
      confirmados: confirmedIds.size,
      cancelados:  canceledIds.size,
      sentAt: Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"),
      isOpen
    }));

  } catch (err) {
    Logger.log("Erro em enviarStatusMonitor_ (" + sheetName + "): " + err.message);
  }
}

function monitorMLColetaTick() {
  const sheetName = "ML Coleta";
  try {
    const now = new Date();
    if (now.getDay() === 0 || now.getDay() === 6) return;

    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (props.getProperty("LAST_IMPORT_DATE_ML_COLETA") !== today) return;
    if (props.getProperty("MONITOR_PAUSED_ML_COLETA") === "true") return;

    const times = getCollectionTimes();
    const ct = times[sheetName] || "";
    if (!ct) return;
    const [h, m] = ct.split(":").map(Number);
    const stop = new Date(now);
    stop.setHours(h + 2, m, 0, 0);
    if (now > stop) return;

    const snapshot = readMonitorSnapshot_(
      sheetName === "ML 1" ? "LAST_MONITOR_SNAPSHOT_ML_1" : "LAST_MONITOR_SNAPSHOT_ML_COLETA"
    );
    if (snapshot && snapshot.sentAt) {
      const [sh, sm] = snapshot.sentAt.split(":").map(Number);
      const sentDate = new Date(now);
      sentDate.setHours(sh, sm, 0, 0);
      const diffMin = (now - sentDate) / 60000;
      if (diffMin >= 0 && diffMin < 30) return;
    }

    enviarStatusMonitor_(sheetName);
  } catch (err) {
    Logger.log("Erro em monitorMLColeta_tick_: " + err.message);
  }
}

function monitorML1Tick() {
  const sheetName = "ML 1";
  try {
    const now = new Date();
    if (now.getDay() === 0 || now.getDay() === 6) return;

    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (props.getProperty("LAST_IMPORT_DATE_ML_1") !== today) return;
    if (props.getProperty("MONITOR_PAUSED_ML_1") === "true") return;

    const times = getCollectionTimes();
    const ct = times[sheetName] || "";
    if (!ct) return;
    const [h, m] = ct.split(":").map(Number);
    const stop = new Date(now);
    stop.setHours(h + 2, m, 0, 0);
    if (now > stop) return;

    const snapshot = readMonitorSnapshot_(
      sheetName === "ML 1" ? "LAST_MONITOR_SNAPSHOT_ML_1" : "LAST_MONITOR_SNAPSHOT_ML_COLETA"
    );
    if (snapshot && snapshot.sentAt) {
      const [sh, sm] = snapshot.sentAt.split(":").map(Number);
      const sentDate = new Date(now);
      sentDate.setHours(sh, sm, 0, 0);
      const diffMin = (now - sentDate) / 60000;
      if (diffMin >= 0 && diffMin < 30) return;
    }

    enviarStatusMonitor_(sheetName);
  } catch (err) {
    Logger.log("Erro em monitorML1_tick_: " + err.message);
  }
}

function pausarMonitorMLColeta() {
  try { PropertiesService.getScriptProperties().setProperty("MONITOR_PAUSED_ML_COLETA", "true"); return { ok: true }; }
  catch(e) { return { ok: false, error: e.message }; }
}
function resumirMonitorMLColeta() {
  try { PropertiesService.getScriptProperties().deleteProperty("MONITOR_PAUSED_ML_COLETA"); return { ok: true }; }
  catch(e) { return { ok: false, error: e.message }; }
}
function pausarMonitorML1() {
  try { PropertiesService.getScriptProperties().setProperty("MONITOR_PAUSED_ML_1", "true"); return { ok: true }; }
  catch(e) { return { ok: false, error: e.message }; }
}
function resumirMonitorML1() {
  try { PropertiesService.getScriptProperties().deleteProperty("MONITOR_PAUSED_ML_1"); return { ok: true }; }
  catch(e) { return { ok: false, error: e.message }; }
}

/*******************************
 * LIMPEZA DIÁRIA (gatilho meia-noite)
 *******************************/
function limparPlanilhaDiaria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of ["ML Coleta", "ML 1"]) {
    const sh = ss.getSheetByName(name);
    if (sh) limparAbaExcetoCabecalho_(sh);
  }
}

/*******************************
 * BULK MARKING
 *******************************/

function bm_getOrCreateParciais_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Parciais");
  if (!sh) {
    sh = ss.insertSheet("Parciais");
    sh.getRange(1,1,1,6).setValues([["Source","OrderID","KitProduct","ComponentName","ComponentStatus","Qty"]]);
  }
  return sh;
}

// Reverse lookup: normalized component -> [normalized kit keys that contain it]
function bm_buildReverseMap_(mapSubseq) {
  const rev = new Map();
  for (const [kitKey, str] of mapSubseq.entries()) {
    for (const part of str.split(";").map(s=>s.trim()).filter(Boolean)) {
      const k = pend_norm_(part);
      if (!rev.has(k)) rev.set(k, []);
      rev.get(k).push(kitKey);
    }
  }
  return rev;
}

// Lê todas as linhas da Conferencia; retorna { rows, ids, prods, statuses, sh }
function bm_readConferencia_() {
  const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.ANDAMENTO_SOURCE_SPREADSHEET_ID);
  const sh = ss.getSheetByName("Conferencia");
  if (!sh) throw new Error("Sheet 'Conferencia' não encontrada");
  const last = sh.getLastRow();
  if (last < 2) return { rows: 0, ids: [], prods: [], statuses: [], sh };
  const n = last - 1;
  return {
    rows: n,
    ids:      sh.getRange(2, 1, n, 1).getValues(),
    prods:    sh.getRange(2, 4, n, 1).getValues(),
    statuses: sh.getRange(2, 6, n, 1).getValues(),
    sh,
  };
}

function buildTenhoItems_() {
  try {
    const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.ANDAMENTO_SOURCE_SPREADSHEET_ID);
    const sh = ss.getSheetByName("Conferencia");
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 2) return [];
    const n = last - 1;
    const prods  = sh.getRange(2, 4, n, 1).getValues();
    const qtds   = sh.getRange(2, 5, n, 1).getValues();
    const stats  = sh.getRange(2, 6, n, 1).getValues();
    const counts = new Map();
    for (let i = 0; i < n; i++) {
      const st = String(stats[i][0] ?? "").trim().toUpperCase();
      if (!st.includes("TENHO")) continue;
      const prod = String(prods[i][0] ?? "").trim();
      if (!prod) continue;
      const qty = pend_parseQtd_(qtds[i][0]);
      counts.set(prod, (counts.get(prod) || 0) + qty);
    }
    return Array.from(counts.entries()).map(([produtoPendentes, quantidade]) => ({ produtoPendentes, quantidade }));
  } catch(e) { return []; }
}

// Remove linhas da sheet Parciais por { orderId, componentName }
function bm_removeParciais_(items, statusMap) {
  const sh = bm_getOrCreateParciais_();
  if (sh.getLastRow() < 2) return;

  const resolvedSet = new Set(items.map(r =>
    String(r.orderId).trim() + "||" + String(r.kitProduct).trim() + "||" +
    pend_norm_(String(r.componentName).trim())
  ));
  const pairSet = new Set(items.map(r =>
    String(r.orderId).trim() + "||" + String(r.kitProduct).trim()
  ));

  // Passo 1: atualizar status das linhas resolvidas (não deletar)
  let vals = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < vals.length; i++) {
    const oId  = String(vals[i][1] ?? "").trim();
    const kit  = String(vals[i][2] ?? "").trim();
    const comp = pend_norm_(String(vals[i][3] ?? "").trim());
    const key  = oId + "||" + kit + "||" + comp;
    if (resolvedSet.has(key)) {
      const newStatus = statusMap ? (statusMap.get(key) || "FALTA") : "FALTA";
      sh.getRange(i + 2, 5).setValue(newStatus);
    }
  }

  // Passo 2: se não restar PENDENTE no par → deletar todas as linhas do par
  const last2 = sh.getLastRow();
  if (last2 < 2) return;
  vals = sh.getRange(2, 1, last2 - 1, 6).getValues();
  const pairHasPending = {};
  for (const v of vals) {
    const oId = String(v[1] ?? "").trim();
    const kit = String(v[2] ?? "").trim();
    const st  = String(v[4] ?? "").trim();
    const key = oId + "||" + kit;
    if (!pairSet.has(key)) continue;
    if (st === "PENDENTE") pairHasPending[key] = true;
  }
  const del = [];
  for (let i = 0; i < vals.length; i++) {
    const oId = String(vals[i][1] ?? "").trim();
    const kit = String(vals[i][2] ?? "").trim();
    const key = oId + "||" + kit;
    if (pairSet.has(key) && !pairHasPending[key]) del.push(i + 2);
  }
  for (let k = del.length - 1; k >= 0; k--) sh.deleteRow(del[k]);
}

function getOrdersByFaltaItem(productName) {
  try {
    const norm = pend_norm_(productName);
    const dashSS = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["ML Coleta","ML 1","Shopee","Magalu","Essência do Brasil","Amazon","Flex/Vapt"];
    const results = [];
    for (const name of sheetNames) {
      const sh = dashSS.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) continue;
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
      for (const row of data) {
        const id = String(row[0] ?? "").trim();
        if (!id) continue;
        const faltaCol = String(row[8] ?? "").trim();
        if (pend_norm_(faltaCol) === norm) {
          results.push({ orderId: id, source: name });
        }
      }
    }
    return { ok: true, orders: results };
  } catch(err) { return { ok: false, error: String(err.message || err) }; }
}

// Retorna todos os itens da sheet Parciais
function getParciais() {
  try {
    const sh = bm_getOrCreateParciais_();
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, items: [] };
    const vals = sh.getRange(2, 1, last - 1, 6).getValues();
    const items = vals
      .map(r => ({
        source:          String(r[0] ?? "").trim(),
        orderId:         String(r[1] ?? "").trim(),
        kitProduct:      String(r[2] ?? "").trim(),
        componentName:   String(r[3] ?? "").trim(),
        componentStatus: String(r[4] ?? "").trim(),
        qty:             Number(r[5] || 0),
      }))
      .filter(it => it.orderId);
    return { ok: true, items };
  } catch(err) { return { ok: false, error: String(err.message||err) }; }
}

// Retorna { ok, orders: [{orderId, source}] } para um dado componentName (para o modal Falta Parcial).
// "source" = nome da aba do dashboard onde o pedido aparece.
function getOrdersByComponent(componentName, ignoreStatus) {
  try {
    const mapSubseq = pend_buildDadosMapSubseq_();
    const rev = bm_buildReverseMap_(mapSubseq);
    const normComp = pend_norm_(componentName);
    const kitKeys  = rev.get(normComp) || [];

    const conf = bm_readConferencia_();
    const matchIds = new Set();

    for (let i = 0; i < conf.rows; i++) {
      const status = String(conf.statuses[i][0] ?? "").trim();
      if (ignoreStatus) {
        if (!status.toUpperCase().includes("FALTA")) continue;
      } else {
        if (status !== "") continue;
      }
      const pk = pend_norm_(String(conf.prods[i][0] ?? "").trim());
      const match = kitKeys.length ? kitKeys.includes(pk) : pk === normComp;
      if (!match) continue;
      const id = String(conf.ids[i][0] ?? "").trim();
      if (id) matchIds.add(id);
    }

    // Resolve source (aba do dashboard) por ID
    const dashSS = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["ML Coleta","ML 1","Shopee","Magalu","Essência do Brasil","Amazon","Flex/Vapt","Próximos Dias"];
    const idSource = new Map();
    for (const name of sheetNames) {
      const sh = dashSS.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) continue;
      const ids = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
      for (const r of ids) {
        const id = String(r[0]??"").trim();
        if (matchIds.has(id) && !idSource.has(id)) idSource.set(id, name);
      }
    }

    const orders = [...matchIds].map(id => ({ orderId: id, source: idSource.get(id) || "—" }));
    return { ok: true, orders };
  } catch(err) { return { ok: false, error: String(err.message||err) }; }
}

// Marca produtos selecionados como /Falta.
// selections = [{ prodName, qty, sheet }]  — vêm dos itens M/N (FALTANV_VIEW) selecionados.
// Lógica:
//   - Produto não-kit (não está no reverse map): busca Conferencia onde col D (norm) = prodName e F vazio → F = "FALTA"
//   - Produto kit-component:
//       • todos os componentes do kit estão em selections → F = "FALTA" nas rows do kit (F vazio)
//       • seleção parcial → NÃO atualiza Conferencia; escreve na sheet Parciais
//         (componentes selecionados: FALTA, restantes: PENDENTE)
function marcarProdutosComoFalta(selections) {
  try {
    if (!selections?.length) return { ok: true, updated: 0, parciaisAdded: 0 };

    const mapSubseq = pend_buildDadosMapSubseq_();
    const rev = bm_buildReverseMap_(mapSubseq);
    const fwd = new Map([...mapSubseq.entries()].map(([k, v]) =>
      [k, v.split(";").map(s=>s.trim()).filter(Boolean)]));

    const selectedNorms = new Set(selections.map(s => pend_norm_(s.prodName)));
    const conf = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);

    let updated = 0;
    let parciaisAdded = 0;
    const processedKits = new Set();

    for (const sel of selections) {
      const normComp = pend_norm_(sel.prodName);
      const kitKeys  = rev.get(normComp);

      if (!kitKeys?.length) {
        // Produto direto — sem mapeamento kit
        for (let i = 0; i < conf.rows; i++) {
          const pk = pend_norm_(String(conf.prods[i][0]??"").trim());
          if (pk !== normComp) continue;
          if (String(conf.statuses[i][0]??"").trim() !== "") continue;
          newSt[i][0] = "FALTA";
          updated++;
        }
        continue;
      }

      for (const kitKey of kitKeys) {
        if (processedKits.has(kitKey)) continue;
        processedKits.add(kitKey);

        const comps     = fwd.get(kitKey) || [];
        const normComps = comps.map(c => pend_norm_(c));
        const allSel    = normComps.every(nc => selectedNorms.has(nc));

        if (allSel) {
          // Kit completo
          for (let i = 0; i < conf.rows; i++) {
            const pk = pend_norm_(String(conf.prods[i][0]??"").trim());
            if (pk !== kitKey) continue;
            if (String(conf.statuses[i][0]??"").trim() !== "") continue;
            newSt[i][0] = "FALTA";
            updated++;
          }
        } else {
          // Kit parcial — escreve Parciais, sem tocar na Conferencia
          const ordersSeen = new Set();
          const kitProdOriginal = (() => {
            for (let i = 0; i < conf.rows; i++) {
              if (pend_norm_(String(conf.prods[i][0]??"").trim()) === kitKey)
                return String(conf.prods[i][0]).trim();
            }
            return kitKey;
          })();

          for (let i = 0; i < conf.rows; i++) {
            const pk = pend_norm_(String(conf.prods[i][0]??"").trim());
            if (pk !== kitKey) continue;
            if (String(conf.statuses[i][0]??"").trim() !== "") continue;
            const id = String(conf.ids[i][0]??"").trim();
            if (id && !ordersSeen.has(id)) ordersSeen.add(id);
          }

          const parciaisSheet = bm_getOrCreateParciais_();
          for (const orderId of ordersSeen) {
            for (let ci = 0; ci < comps.length; ci++) {
              const cNorm   = pend_norm_(comps[ci]);
              const cStatus = selectedNorms.has(cNorm) ? "FALTA" : "PENDENTE";
              parciaisSheet.appendRow([sel.sheet || "", orderId, kitProdOriginal, comps[ci], cStatus, sel.qty || 1]);
              parciaisAdded++;
            }
          }
        }
      }
    }

    if (updated > 0) conf.sh.getRange(2, 6, conf.rows, 1).setValues(newSt);
    return { ok: true, updated, parciaisAdded };
  } catch(err) { return { ok: false, error: String(err.message||err) }; }
}

// Para o modal Falta Parcial: marca tenhoIds com F = "FALTA - [componentName]"
// e os demais (status vazio) com F = "FALTA".
function marcarFaltaParcial(componentName, tenhoIds, allModalIds) {
  try {
    const mapSubseq = pend_buildDadosMapSubseq_();
    const rev       = bm_buildReverseMap_(mapSubseq);
    const normComp  = pend_norm_(componentName);
    const kitKeys   = rev.get(normComp) || [];
    const tenhoSet  = new Set(tenhoIds.map(id => String(id).trim()));
    const allSet    = new Set((allModalIds || []).map(id => String(id).trim()));

    const fwd = new Map();
    for (const [k, v] of mapSubseq.entries()) {
      fwd.set(k, v.split(";").map(s => s.trim()).filter(Boolean));
    }

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let updated       = 0;
    let parciaisAdded = 0;

    // Resolve source (aba) por orderId — necessário para entradas no Parciais
    const dashSS     = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["ML Coleta","ML 1","Shopee","Magalu","Essência do Brasil","Amazon","Flex/Vapt"];
    const idSource   = new Map();
    for (const name of sheetNames) {
      const sh = dashSS.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) continue;
      const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (const r of ids) {
        const id = String(r[0] ?? "").trim();
        if (allSet.has(id) && !idSource.has(id)) idSource.set(id, name);
      }
    }

    const parcSh        = bm_getOrCreateParciais_();
    const processedKeys = new Set();

    for (let i = 0; i < conf.rows; i++) {
      if (String(conf.statuses[i][0] ?? "").trim() !== "") continue;
      const pk    = pend_norm_(String(conf.prods[i][0] ?? "").trim());
      const match = kitKeys.length ? kitKeys.includes(pk) : pk === normComp;
      if (!match) continue;
      const id = String(conf.ids[i][0] ?? "").trim();
      if (!id) continue;
      if (allSet.size > 0 && !allSet.has(id)) continue;

      // Decisão por pedido: kit real = mais de 1 componente no Dados
      const comps     = fwd.get(pk) || [];
      const isRealKit = comps.length > 1;

      if (!isRealKit) {
        // Standalone (0 ou 1 componente): toca Conferencia diretamente
        if (tenhoSet.has(id)) continue;   // TENHO: ignora
        newSt[i][0] = "FALTA";
        updated++;
      } else {
        // Kit real: escreve na Parciais, não toca Conferencia
        const orderKey = id + "||" + pk;
        if (processedKeys.has(orderKey)) continue;
        processedKeys.add(orderKey);

        const kitProdOriginal = String(conf.prods[i][0]).trim();
        const source          = idSource.get(id) || "";
        const isTenho         = tenhoSet.has(id);

        for (const comp of comps) {
          const compNorm = pend_norm_(comp);
          let cStatus;
          if (compNorm === normComp) {
            cStatus = isTenho ? "TENHO" : "FALTA";
          } else {
            cStatus = "PENDENTE";
          }
          parcSh.appendRow([source, id, kitProdOriginal, comp, cStatus, 1]);
          parciaisAdded++;
        }
      }
    }

    if (updated > 0) conf.sh.getRange(2, 6, conf.rows, 1).setValues(newSt);
    return { ok: true, updated, parciaisAdded };
  } catch(err) { return { ok: false, error: String(err.message || err) }; }
}

// Resolve itens de Parciais como /Falta: Conferencia F = "FALTA" para as ordens dos itens.
// items = [{ orderId, kitProduct, componentName }]
function resolverParcialComoFalta(items) {
  try {
    if (!items?.length) return { ok: true, updated: 0 };

    const parcSh   = bm_getOrCreateParciais_();
    const parcData = parcSh.getDataRange().getValues();
    const ordKitStatus = {};
    for (let r = 1; r < parcData.length; r++) {
      const oId    = String(parcData[r][1] || "").trim();
      const kit    = String(parcData[r][2] || "").trim();
      const comp   = String(parcData[r][3] || "").trim();
      const status = String(parcData[r][4] || "").trim();
      if (!oId || !kit || !comp) continue;
      const key = oId + "||" + kit;
      if (!ordKitStatus[key]) ordKitStatus[key] = {};
      ordKitStatus[key][comp] = status;
    }

    // Projeta o estado final: aplica FALTA em cada item sobre o estado atual da Parciais
    const projectedStatus = {};
    for (const item of items) {
      const pairKey = String(item.orderId).trim() + "||" + String(item.kitProduct).trim();
      if (!projectedStatus[pairKey]) {
        projectedStatus[pairKey] = { ...(ordKitStatus[pairKey] || {}) };
      }
      projectedStatus[pairKey][item.componentName] = "FALTA";
    }

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let updated = 0;

    const processedPairs = new Set();
    for (const item of items) {
      const orderId = String(item.orderId).trim();
      const pairKey = orderId + "||" + String(item.kitProduct).trim();
      if (processedPairs.has(pairKey)) continue;
      processedPairs.add(pairKey);

      const projected = projectedStatus[pairKey] || {};
      // Se ainda restar algum PENDENTE, não toca na Conferencia
      if (Object.values(projected).some(st => st === "PENDENTE")) continue;

      const normKit    = pend_norm_(item.kitProduct);
      const faltaComps = Object.entries(projected)
        .filter(([, st]) => st === "FALTA")
        .map(([c]) => c);
      const hasAnyTenho = Object.values(projected).some(st => st === "TENHO");
      const finalStatus = hasAnyTenho
        ? "FALTA - " + faltaComps.join(";")
        : "FALTA";

      for (let i = 0; i < conf.rows; i++) {
        if (String(conf.ids[i][0]  ?? "").trim() !== orderId) continue;
        if (pend_norm_(String(conf.prods[i][0] ?? "").trim()) !== normKit) continue;
        if (String(conf.statuses[i][0] ?? "").toUpperCase().includes("TENHO")) continue;
        newSt[i][0] = finalStatus;
        updated++;
      }
    }

    if (updated > 0) conf.sh.getRange(2, 6, conf.rows, 1).setValues(newSt);

    const statusMap = new Map(items.map(item => [
      String(item.orderId).trim() + "||" +
      String(item.kitProduct).trim() + "||" +
      pend_norm_(String(item.componentName).trim()),
      "FALTA"
    ]));
    bm_removeParciais_(items, statusMap);
    return { ok: true, updated };
  } catch(err) { return { ok: false, error: String(err.message || err) }; }
}

// Resolve itens de Parciais com Tenho (modal de seleção de IDs).
// tenhoIds = IDs onde o componente é Tenho → F = "FALTA - [componentName]"
// demais IDs dos mesmos itens → F = "FALTA"
function resolverParcialComTenho(items, tenhoIds) {
  try {
    if (!items?.length) return { ok: true, updated: 0 };
    const tenhoSet = new Set(tenhoIds.map(id => String(id).trim()));

    const parcSh   = bm_getOrCreateParciais_();
    const parcData = parcSh.getDataRange().getValues();
    const ordKitStatus = {};
    for (let r = 1; r < parcData.length; r++) {
      const oId    = String(parcData[r][1] || "").trim();
      const kit    = String(parcData[r][2] || "").trim();
      const comp   = String(parcData[r][3] || "").trim();
      const status = String(parcData[r][4] || "").trim();
      if (!oId || !kit || !comp) continue;
      const key = oId + "||" + kit;
      if (!ordKitStatus[key]) ordKitStatus[key] = {};
      ordKitStatus[key][comp] = status;
    }

    // Projeta o estado final: aplica TENHO ou FALTA conforme tenhoSet
    const projectedStatus = {};
    for (const item of items) {
      const pairKey = String(item.orderId).trim() + "||" + String(item.kitProduct).trim();
      if (!projectedStatus[pairKey]) {
        projectedStatus[pairKey] = { ...(ordKitStatus[pairKey] || {}) };
      }
      projectedStatus[pairKey][item.componentName] =
        tenhoSet.has(String(item.orderId).trim()) ? "TENHO" : "FALTA";
    }

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let updated = 0;

    const processedPairs = new Set();
    for (const item of items) {
      const orderId = String(item.orderId).trim();
      const pairKey = orderId + "||" + String(item.kitProduct).trim();
      if (processedPairs.has(pairKey)) continue;
      processedPairs.add(pairKey);

      const projected = projectedStatus[pairKey] || {};
      // Se ainda restar algum PENDENTE, não toca na Conferencia
      if (Object.values(projected).some(st => st === "PENDENTE")) continue;

      const normKit    = pend_norm_(item.kitProduct);
      const faltaComps = Object.entries(projected)
        .filter(([, st]) => st === "FALTA")
        .map(([c]) => c);
      // Todos TENHO → nada a marcar como FALTA
      if (faltaComps.length === 0) continue;

      const hasAnyTenho = Object.values(projected).some(st => st === "TENHO");
      const finalStatus = hasAnyTenho
        ? "FALTA - " + faltaComps.join(";")
        : "FALTA";

      for (let i = 0; i < conf.rows; i++) {
        if (String(conf.ids[i][0]  ?? "").trim() !== orderId) continue;
        if (pend_norm_(String(conf.prods[i][0] ?? "").trim()) !== normKit) continue;
        if (String(conf.statuses[i][0] ?? "").toUpperCase().includes("TENHO")) continue;
        newSt[i][0] = finalStatus;
        updated++;
      }
    }

    if (updated > 0) conf.sh.getRange(2, 6, conf.rows, 1).setValues(newSt);

    const statusMap = new Map(items.map(item => {
      const key = String(item.orderId).trim() + "||" +
                  String(item.kitProduct).trim() + "||" +
                  pend_norm_(String(item.componentName).trim());
      return [key, tenhoSet.has(String(item.orderId).trim()) ? "TENHO" : "FALTA"];
    }));
    bm_removeParciais_(items, statusMap);
    return { ok: true, updated };
  } catch(err) { return { ok: false, error: String(err.message || err) }; }
}

function processarFotosPedidos() {
  try {
    const FOLDER_ID = "12gA8sJPuQ2LnkUgwpI0Ix3Bw0AvjMnP5";
    const EMAIL     = "victor@gigaimports.com";

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files  = folder.getFiles();

    const success     = [];
    const notFound    = [];
    const failed      = [];
    const multipleIds = [];
    let   total       = 0;

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let anyUpdated = false;

    while (files.hasNext()) {
      const file     = files.next();
      const mimeType = file.getMimeType();
      if (!mimeType.startsWith("image/")) continue;

      total++;
      const fileName = file.getName();

      let copiedId = null;
      try {
        const copied = DriveApi.Files.copy(
          { title: "ocr_temp_" + file.getId(),
            mimeType: "application/vnd.google-apps.document" },
          file.getId(),
          { ocr: true, ocrLanguage: "pt" }
        );
        copiedId = copied.id;
        const doc  = DocumentApp.openById(copied.id);
        const text = doc.getBody().getText();
        Logger.log("OCR FILE: " + fileName + " | TEXT: " + text.substring(0, 300));

        const matches = [...text.matchAll(/\d{2}\/\d{2}\/\d{4}\s*[-–]\s*([^\n\r]+)/g)];

        if (matches.length === 0) {
          failed.push({ file: fileName, reason: "ID não encontrado na foto" });
          file.setTrashed(true);
          continue;
        }
        if (matches.length > 1) {
          multipleIds.push({ file: fileName, ids: matches.map(m => m[1].trim()) });
          continue;
        }

        const rawCapture = matches[0][1].trim().replace(/^[#\s]+/, "").trim();
        const orderId = rawCapture.split(/\s+/)[0];

        let found = false;
        for (let i = 0; i < conf.rows; i++) {
          const id = String(conf.ids[i][0] ?? "").trim();
          if (id === orderId) {
            newSt[i][0] = "TENHO";
            found = true;
            anyUpdated = true;
          }
        }

        if (!found) {
          notFound.push({ file: fileName, orderId });
          file.setTrashed(true);
        } else {
          success.push({ file: fileName, orderId });
          file.setTrashed(true);
        }

      } catch (err) {
        failed.push({ file: fileName, reason: "Erro: " + err.message });
        file.setTrashed(true);
      } finally {
        if (copiedId) {
          try { DriveApi.Files.remove(copiedId); } catch(e) { Logger.log("Cleanup err: " + e.message); }
        }
      }
    }

    if (anyUpdated) conf.sh.getRange(2, 6, conf.rows, 1).setValues(newSt);

    let body = "Processamento de fotos concluído.\n\n"
      + "Total de fotos: " + total + "\n"
      + "✅ Sucesso: " + success.length + "\n"
      + "⚠️ Não Encontrados: " + notFound.length + "\n"
      + "❌ Falhas: " + failed.length + "\n"
      + "⚠️ Múltiplos pedidos na foto (puladas): " + multipleIds.length + "\n";

    if (success.length) {
      body += "\n--- SUCESSO ---\n";
      success.forEach(r => { body += "• " + r.file + " → Pedido " + r.orderId + " marcado como TENHO\n"; });
    }
    if (notFound.length) {
      body += "\n--- NÃO ENCONTRADOS NA CONFERENCIA ---\n";
      notFound.forEach(r => { body += "• " + r.file + ": Pedido " + r.orderId + " não encontrado\n"; });
    }
    if (failed.length) {
      body += "\n--- FALHAS ---\n";
      failed.forEach(r => { body += "• " + r.file + ": " + r.reason + "\n"; });
    }
    if (multipleIds.length) {
      body += "\n--- MÚLTIPLOS PEDIDOS NA FOTO (puladas, não apagadas) ---\n";
      multipleIds.forEach(r => { body += "• " + r.file + ": " + r.ids.join(", ") + "\n"; });
    }

    MailApp.sendEmail({
      to: EMAIL,
      subject: "Processamento de Fotos — " + success.length + " sucesso(s), " + notFound.length + " não encontrado(s), " + failed.length + " falha(s)",
      body: body
    });
    Logger.log("Email enviado para " + EMAIL);

    return { ok: true, total, success: success.length, failed: failed.length, skipped: multipleIds.length };

  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

function atualizarMarcacoesRapido() {
  try {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const mapSubseq     = pend_buildDadosMapSubseq_();
    const confById      = pend_buildConferenciaById_();
    const confTodosById = pend_buildTodosExcTenho_();
    const sheetNames = ["ML Coleta","ML 1","Shopee","Magalu","Essência do Brasil","Amazon","Flex/Vapt","Próximos Dias"];
    for (const name of sheetNames) {
      const sh = ss.getSheetByName(name);
      if (!sh) continue;
      if (name === "ML Coleta") {
        pend_process_MLColeta_(sh, confById, mapSubseq);
      } else {
        pend_process_porAba_(sh, confById, mapSubseq);
      }
      pend_process_nv_porId_(sh, confTodosById, mapSubseq);
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: String(e.message || e) };
  }
}

function bm_getOrderDetails_(orderIds) {
  const result = new Map();
  const set = new Set(orderIds.map(id => String(id).trim()));
  const dashSS = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ["ML 1","Flex/Vapt","ML Coleta","Magalu","Shopee","Amazon","Essência do Brasil","Próximos Dias"];
  for (const name of sheetNames) {
    const sh = dashSS.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) continue;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (const row of data) {
      const id = String(row[0] ?? "").trim();
      if (set.has(id) && !result.has(id))
        result.set(id, { id, cliente: String(row[1] ?? "").trim(), source: name });
    }
  }
  return result;
}

function gerarEtiquetasPDF_(tenhoOrderDetails, folderId) {
  if (!tenhoOrderDetails.length) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Etiquetas");
  if (!sheet) sheet = ss.insertSheet("Etiquetas");
  sheet.clearContents();
  sheet.setColumnWidth(1, 160);
  for (let i = 0; i < tenhoOrderDetails.length; i++) {
    const d = tenhoOrderDetails[i];
    const text = "*** ETIQUETA PROVISÓRIA DE UM PEDIDO INCOMPLETO ***\n" +
                 "Pedido: " + d.id +
                 "\nCliente: " + d.cliente +
                 "\nMarketplace: " + d.source +
                 "\nProdutos:\n" + d.produtos.join("\n") +
                 "\nPedido Incompleto";
    sheet.getRange(i + 1, 1).setValue(text);
  }
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha === 0) return null;
  sheet.getRange(1, 1, ultimaLinha, 1)
    .setWrap(true)
    .setHorizontalAlignment("left")
    .setBorder(true, true, true, true, true, true);
  const url_base = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?";
  const params = {
    format: "pdf", size: "a5", portrait: true, fitw: false,
    scale: 1, top_margin: 0.1, bottom_margin: 0.1,
    left_margin: 0.05, right_margin: 0.05,
    sheetnames: false, printtitle: false, pagenumbers: false,
    gridlines: false, fzr: false,
    gid: sheet.getSheetId(), range: "A1:A" + ultimaLinha
  };
  const query = Object.entries(params).map(([k,v]) => k + "=" + v).join("&");
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url_base + query,
    { headers: { Authorization: "Bearer " + token } });
  const folder  = DriveApp.getFolderById(folderId);
  const nome    = "Etiqueta_" + new Date().toISOString().replace(/[:.]/g, "-") + ".pdf";
  const arquivo = folder.createFile(response.getBlob().setName(nome));
  sheet.clearContents();
  return arquivo.getUrl();
}

function enviarInformacoesNaoMarcado(selections) {
  try {
    if (!selections?.length) return { ok: true, updated: 0, labels: 0 };
    const PRIORITY = ["ML 1","Flex/Vapt","ML Coleta","Magalu","Shopee","Amazon","Essência do Brasil","Próximos Dias"];
    const PDF_FOLDER = "1uhmWR9BRt09Ycwagd6g-jyEsulgtZiru";

    const mapSubseq = pend_buildDadosMapSubseq_();
    const rev = bm_buildReverseMap_(mapSubseq);
    const fwd = new Map([...mapSubseq.entries()].map(([k,v]) =>
      [k, v.split(";").map(s=>s.trim()).filter(Boolean)]));

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let updated = 0;
    let parciaisAdded = 0;
    const tenhoOrderProds    = new Map();
    const kitOrderCompStatus = new Map();
    const parcSh = bm_getOrCreateParciais_();
    const dashSS = SpreadsheetApp.getActiveSpreadsheet();

    const _idSourcePreload = new Map();
    for (const _nm of PRIORITY) {
      const _sh = dashSS.getSheetByName(_nm);
      if (!_sh || _sh.getLastRow() < 2) continue;
      const _rows = _sh.getRange(2, 1, _sh.getLastRow() - 1, 1).getValues();
      for (const _r of _rows) {
        const _id = String(_r[0] ?? "").trim();
        if (_id && !_idSourcePreload.has(_id)) _idSourcePreload.set(_id, _nm);
      }
    }

    for (const sel of selections) {
      if (!sel.falta && !sel.tenho) continue;
      const normComp = pend_norm_(sel.prodName);
      const kitKeys  = rev.get(normComp) || [];

      // Varre Conferencia diretamente: match direto ou componente de kit
      const matchedConf = [];
      for (let i = 0; i < conf.rows; i++) {
        if (String(conf.statuses[i][0] ?? "").trim() !== "") continue;
        const pk = pend_norm_(String(conf.prods[i][0] ?? "").trim());
        const id = String(conf.ids[i][0] ?? "").trim();
        if (!id) continue;
        const isDirect = pk === normComp;
        const isKit    = kitKeys.length > 0 && kitKeys.includes(pk);
        if (!isDirect && !isKit) continue;
        matchedConf.push({ confIdx: i, orderId: id, pk, isDirect, isKit });
      }
      if (!matchedConf.length) continue;

      // Resolve source usando mapa pré-carregado
      const orderIdSet = new Set(matchedConf.map(r => r.orderId));
      const idSource = new Map();
      for (const id of orderIdSet) {
        if (_idSourcePreload.has(id)) idSource.set(id, _idSourcePreload.get(id));
      }

      // Ignora pedidos que não existem no Dashboard
      const orders = [...orderIdSet]
        .filter(id => idSource.has(id))
        .map(id => ({ orderId: id, source: idSource.get(id) }))
        .sort((a, b) => PRIORITY.indexOf(a.source) - PRIORITY.indexOf(b.source));
      if (!orders.length) continue;

      // Divide TENHO / FALTA por prioridade
      let tenhoOrders, faltaOrders;
      if (sel.falta && !sel.tenho) {
        tenhoOrders = []; faltaOrders = orders;
      } else {
        const tq = (sel.tenhoQty > 0 && sel.tenhoQty < orders.length)
          ? sel.tenhoQty : orders.length;
        tenhoOrders = orders.slice(0, tq);
        faltaOrders = orders.slice(tq);
      }
      const tenhoSet = new Set(tenhoOrders.map(o => o.orderId));
      const faltaSet  = new Set(faltaOrders.map(o  => o.orderId));

      // Processa linhas da Conferencia encontradas
      for (const { confIdx, orderId, isDirect, isKit, pk } of matchedConf) {
        if (!tenhoSet.has(orderId) && !faltaSet.has(orderId)) continue;
        const status = tenhoSet.has(orderId) ? "TENHO" : "FALTA";
        if (isDirect) {
          newSt[confIdx][0] = status;
          updated++;
          if (status === "TENHO") {
            if (!tenhoOrderProds.has(orderId)) tenhoOrderProds.set(orderId, []);
            tenhoOrderProds.get(orderId).push(sel.prodName);
          }
        } else if (isKit) {
          const mapKey = orderId + "||" + pk;
          if (!kitOrderCompStatus.has(mapKey)) kitOrderCompStatus.set(mapKey, new Map());
          kitOrderCompStatus.get(mapKey).set(normComp, status);
        }
      }
    }

    // Lê snapshot do Parciais antes de resolver kits (para merge com envios anteriores)
    const parcSnapshot = parcSh.getLastRow() >= 2
      ? parcSh.getRange(2, 1, parcSh.getLastRow() - 1, 5).getValues()
      : [];

    // Resolve kits acumulados
    for (const [mapKey, compMap] of kitOrderCompStatus.entries()) {
      const sepIdx  = mapKey.indexOf("||");
      const orderId = mapKey.substring(0, sepIdx);
      const kitKey  = mapKey.substring(sepIdx + 2);
      const comps   = fwd.get(kitKey) || [];

      const kitProdOriginal = (() => {
        for (let i = 0; i < conf.rows; i++)
          if (pend_norm_(String(conf.prods[i][0] ?? "").trim()) === kitKey)
            return String(conf.prods[i][0]).trim();
        return kitKey;
      })();

      const source = bm_getOrderDetails_([orderId]).get(orderId)?.source || "";

      // Lê linhas existentes desta combinação orderId+kit no Parciais
      const existingRows = [];
      for (let i = 0; i < parcSnapshot.length; i++) {
        const pid   = String(parcSnapshot[i][1] ?? "").trim();
        const pkitN = pend_norm_(String(parcSnapshot[i][2] ?? "").trim());
        if (pid !== orderId || pkitN !== kitKey) continue;
        const compN = pend_norm_(String(parcSnapshot[i][3] ?? "").trim());
        const stat  = String(parcSnapshot[i][4] ?? "").trim();
        existingRows.push({ rowIdx: i + 2, compNorm: compN, status: stat });
      }

      // Merge: existente (TENHO/FALTA) + atual, atual tem precedência
      const mergedMap = new Map();
      for (const r of existingRows) {
        if (r.status === "TENHO" || r.status === "FALTA") mergedMap.set(r.compNorm, r.status);
      }
      for (const [k, v] of compMap) mergedMap.set(k, v);

      const allResolved = comps.length > 0 && comps.every(c =>
        mergedMap.get(pend_norm_(c)) === "TENHO" || mergedMap.get(pend_norm_(c)) === "FALTA"
      );

      if (allResolved) {
        // Kit completo → escreve Conferencia
        const allTenho = comps.every(c => mergedMap.get(pend_norm_(c)) === "TENHO");
        let finalStatus;
        if (allTenho) {
          finalStatus = "TENHO";
        } else {
          const faltaComps  = comps.filter(c => mergedMap.get(pend_norm_(c)) === "FALTA");
          const hasAnyTenho = comps.some(c => mergedMap.get(pend_norm_(c)) === "TENHO");
          finalStatus = (hasAnyTenho && faltaComps.length > 0)
            ? "FALTA - " + faltaComps.join(";")
            : "FALTA";
        }
        for (let i = 0; i < conf.rows; i++) {
          if (pend_norm_(String(conf.prods[i][0] ?? "").trim()) !== kitKey) continue;
          if (String(conf.ids[i][0] ?? "").trim() !== orderId) continue;
          if (String(conf.statuses[i][0] ?? "").trim() !== "") continue;
          newSt[i][0] = finalStatus;
          updated++;
        }
        // Kit com qualquer TENHO → inclui na geração de etiqueta com apenas os componentes TENHO
        const hasAnyTenhoKit = comps.some(c => mergedMap.get(pend_norm_(c)) === "TENHO");
        if (hasAnyTenhoKit) {
          if (!tenhoOrderProds.has(orderId)) tenhoOrderProds.set(orderId, []);
          const tenhoComps = comps.filter(c => mergedMap.get(pend_norm_(c)) === "TENHO");
          for (const tc of tenhoComps) {
            if (!tenhoOrderProds.get(orderId).includes(tc)) {
              tenhoOrderProds.get(orderId).push(tc);
            }
          }
        }
        // Atualiza linhas PENDENTE no Parciais para que o cleanup as remova
        for (const r of existingRows) {
          if (r.status === "PENDENTE") {
            parcSh.getRange(r.rowIdx, 5).setValue(mergedMap.get(r.compNorm) || "TENHO");
          }
        }
      } else {
        // Kit incompleto → só insere/atualiza componentes desta seleção, sem duplicar
        for (const comp of comps) {
          const normC = pend_norm_(comp);
          const newStatus  = compMap.get(normC);
          const existingRow = existingRows.find(r => r.compNorm === normC);

          if (existingRow) {
            // Atualiza linha PENDENTE existente se chegou status novo
            if (existingRow.status === "PENDENTE" && newStatus) {
              parcSh.getRange(existingRow.rowIdx, 5).setValue(newStatus);
            }
          } else {
            // Insere nova linha apenas para componentes sem linha ainda
            parcSh.appendRow([source, orderId, kitProdOriginal, comp, newStatus || "PENDENTE", 1]);
            parciaisAdded++;
          }
        }
      }
    }

    if (updated > 0) conf.sh.getRange(2, 6, conf.rows, 1).setValues(newSt);

    // Monta índice orderId → índices em conf (para verificar se pedido está completo)
    const confByOrderId = new Map();
    for (let i = 0; i < conf.rows; i++) {
      const id = String(conf.ids[i][0] ?? "").trim();
      if (!id) continue;
      if (!confByOrderId.has(id)) confByOrderId.set(id, []);
      confByOrderId.get(id).push(i);
    }

    // Gera etiqueta apenas para pedidos com TODOS os itens da Conferencia resolvidos
    let pdfUrl = null;
    let labelsGenerated = 0;
    if (tenhoOrderProds.size > 0) {
      const ids    = [...tenhoOrderProds.keys()];
      const detMap = bm_getOrderDetails_(ids);
      const details = [];

      for (const id of ids) {
        const confIdxs = confByOrderId.get(id) || [];
        const allDone  = confIdxs.length > 0 &&
          confIdxs.every(i => String(newSt[i][0] ?? "").trim() !== "");
        if (!allDone) continue;

        const isComplete = confIdxs.every(i =>
          String(newSt[i][0] ?? "").trim() === "TENHO"
        );
        const allTenhoProds = [...new Set(tenhoOrderProds.get(id) || [])];
        details.push({
          id,
          cliente:    detMap.get(id)?.cliente || "—",
          source:     detMap.get(id)?.source  || "—",
          produtos:   allTenhoProds,
          isComplete,
        });
      }

      if (details.length > 0) {
        labelsGenerated = details.length;
        const incompletos = details.filter(d => !d.isComplete);
        const completos   = details.filter(d =>  d.isComplete);

        // PDF somente para pedidos incompletos
        if (incompletos.length > 0) {
          try { pdfUrl = gerarEtiquetasPDF_(incompletos, PDF_FOLDER); }
          catch(e) { Logger.log("PDF err: " + e.message); }
        }

        // E-mail para pedidos completos, agrupados por marketplace
        if (completos.length > 0) {
          const grupos = {};
          for (const d of completos) {
            const src = d.source || "—";
            if (!grupos[src]) grupos[src] = [];
            grupos[src].push(d);
          }
          const linhas = [];
          for (const src of Object.keys(grupos).sort()) {
            linhas.push(src + ":");
            for (const d of grupos[src]) {
              linhas.push(d.id + "   " + d.cliente);
            }
            linhas.push("");
          }
          try {
            MailApp.sendEmail({
              to: "victor@gigaimports.com",
              subject: "Pedidos Completos — " + completos.length + " pedido(s)",
              body: linhas.join("\n")
            });
          } catch(e) { Logger.log("Email err: " + e.message); }
        }
      }
    }

    // Remove linhas do Parciais de kits sem PENDENTE restante
    const parcLast = parcSh.getLastRow();
    if (parcLast >= 2) {
      const pv = parcSh.getRange(2, 1, parcLast - 1, 6).getValues();
      const pairPend = {};
      for (const v of pv) {
        const key = String(v[1]??"").trim() + "||" + String(v[2]??"").trim();
        if (String(v[4]??"").trim() === "PENDENTE") pairPend[key] = true;
      }
      const del = [];
      for (let i = 0; i < pv.length; i++) {
        const key = String(pv[i][1]??"").trim() + "||" + String(pv[i][2]??"").trim();
        if (!pairPend[key]) del.push(i + 2);
      }
      for (let k = del.length - 1; k >= 0; k--) parcSh.deleteRow(del[k]);
    }

    return { ok: true, updated, parciaisAdded, labels: labelsGenerated, pdfUrl };
  } catch(err) {
    Logger.log("enviarInformacoesNaoMarcado: ERRO — " + (err.message || String(err)) + "\nStack: " + (err.stack || "n/a"));
    return { ok: false, error: String(err.message || err) };
  }
}

/******************* PEND DATES — contador de dias pendente *******************/
function pendDates_load_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty("PENDING_DATES");
    return raw ? JSON.parse(raw) : {};
  } catch(e) { return {}; }
}

function pendDates_save_(map) {
  PropertiesService.getScriptProperties().setProperty("PENDING_DATES", JSON.stringify(map));
}

function pendDates_hoje_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function pendDates_diasDesde_(isoDate) {
  if (!isoDate) return null;
  const hoje = new Date();
  const entrada = new Date(isoDate);
  return Math.floor((hoje - entrada) / (1000 * 60 * 60 * 24));
}

/******************* RELATÓRIO DIÁRIO DE PENDENTES *******************/
function enviarRelatorioPendentes_() {
  try {
    const EMAIL = "victor@gigaimports.com";
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const dates = pendDates_load_();
    const hoje  = pendDates_hoje_();
    const abas  = ["Shopee","Magalu","Essência do Brasil","Amazon"];

    const grupos = {};
    for (const nome of abas) {
      const sh = ss.getSheetByName(nome);
      if (!sh || sh.getLastRow() < 2) continue;
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
      const linhas = [];
      const _seenRel = new Set();
      for (const r of data) {
        const id      = String(r[0] ?? "").trim();
        const cliente = String(r[1] ?? "").trim();
        const status  = String(r[4] ?? "").trim();
        if (!id || status === "Confirmado" || status === "Cancelado") continue;
        if (_seenRel.has(id)) continue;
        _seenRel.add(id);
        const dias = pendDates_diasDesde_(dates[id]);
        linhas.push({ id, cliente, dias: dias !== null ? dias : 0 });
      }
      if (linhas.length) {
        linhas.sort((a, b) => b.dias - a.dias);
        grupos[nome] = linhas;
      }
    }

    if (!Object.keys(grupos).length) {
      Logger.log("enviarRelatorioPendentes_: nenhum pedido pendente.");
      return;
    }

    const partes = [];
    for (const [nome, linhas] of Object.entries(grupos)) {
      partes.push(nome + " (" + linhas.length + " pedido(s)):");
      partes.push("Dias   ID                     Cliente");
      partes.push("-----  ---------------------  --------");
      for (const l of linhas) {
        const d = String(l.dias).padStart(5);
        partes.push(d + "  " + l.id.padEnd(22) + " " + l.cliente);
      }
      partes.push("");
    }

    MailApp.sendEmail({
      to: EMAIL,
      subject: "Relatório de Pedidos Pendentes — " + hoje,
      body: partes.join("\n")
    });
    Logger.log("Relatório de pendentes enviado.");
  } catch(e) {
    Logger.log("enviarRelatorioPendentes_: erro — " + e.message);
  }
}

function criarGatilhoRelatorio() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "enviarRelatorioPendentes_")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("enviarRelatorioPendentes_")
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();

  Logger.log("Gatilho criado: enviarRelatorioPendentes_ às 5h.");
}

function getAlertaBanner() {
  try {
    const props  = PropertiesService.getScriptProperties();
    const times  = getCollectionTimes();
    const now    = new Date();
    const alerts = [];

    const checks = [
      { name: "ML Coleta", tsKey: "LAST_IMPORT_TS_ML_COLETA", folderId: "1L1wISlKCaTgZ713TlVI16qN6u20KV2vv" },
      { name: "ML 1",      tsKey: "LAST_IMPORT_TS_ML_1",      folderId: "1TxvlZusR0ilCNjyUmMV_yZsomjDZ7y8d" }
    ];

    for (const { name, tsKey, folderId } of checks) {
      const ct = times[name] || "";
      if (!ct) continue;
      const [h, m] = ct.split(":").map(Number);
      const collectionToday = new Date(now);
      collectionToday.setHours(h, m, 0, 0);
      if (now < collectionToday) continue;
      const lastTs = Number(props.getProperty(tsKey) || 0);
      if (lastTs >= collectionToday.getTime()) continue;
      try { if (pastaTemXlsx_(folderId)) continue; } catch(e) {}
      alerts.push(name);
    }
    return alerts;
  } catch(e) {
    return [];
  }
}

function avisoSegurancaShopee() {
  try {
    const now = new Date();
    if (now.getDay() === 0 || now.getDay() === 6) return;
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (props.getProperty("AVISO_SHOPEE_DATE") === today) return;

    const spacesRaw = String(props.getProperty("CHAT_SPACE_SHOPEE") || "").trim();
    if (!spacesRaw) { Logger.log("avisoSegurancaShopee: CHAT_SPACE_SHOPEE não configurado."); return; }

    const ssSrc = SpreadsheetApp.openById(ML_IMPORT_CONFIG.STATUS_SOURCE_SPREADSHEET_ID);
    const pendentes = [];
    const seenIds   = new Set();

    for (const sheetName of ["Junho", "Maio"]) {
      const sh = ssSrc.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) continue;
      const n      = sh.getLastRow() - 1;
      const values = sh.getRange(2, 1, n, 7).getValues();
      const bgs    = sh.getRange(2, 7, n, 1).getBackgrounds();
      for (let i = 0; i < n; i++) {
        const colC = String(values[i][2] ?? "").toLowerCase();
        if (!colC.includes("shopee")) continue;
        const colG = String(values[i][6] ?? "").trim().toLowerCase();
        const bgG  = String(bgs[i][0]    ?? "").trim().toLowerCase();
        if (colG.includes("ok") || bgG === ML_IMPORT_CONFIG.OK_GREEN.toLowerCase()) continue;
        if (colG.includes("cancelado")) continue;
        if (colG.includes("full") || colG.includes("fulfil")) continue;
        const id      = String(values[i][1] ?? "").trim();
        const cliente = String(values[i][3] ?? "").trim();
        if (!id || seenIds.has(id)) continue;
        seenIds.add(id);
        pendentes.push({ id, cliente });
      }
    }

    props.setProperty("AVISO_SHOPEE_DATE", today);

    if (!pendentes.length) { Logger.log("avisoSegurancaShopee: nenhum pedido Shopee pendente."); return; }

    const text =
      "⚠ *Shopee* — Pedidos Pendentes às 14:00\n" +
      "Ainda há " + pendentes.length + " pedido(s) pendente(s). Verificar se já foram enviados ou preparados:\n" +
      pendentes.map(p => "• " + p.id + " — " + p.cliente).join("\n");

    const normalize = id => id.startsWith("spaces/") ? id : "spaces/" + id;
    for (const spaceId of spacesRaw.split(",").map(s => s.trim()).filter(Boolean)) {
      try { callChatApi_("post", normalize(spaceId) + "/messages", { text }); }
      catch(e) { Logger.log("avisoSegurancaShopee chat err: " + e.message); }
    }
    Logger.log("avisoSegurancaShopee: mensagem enviada. " + pendentes.length + " pedido(s).");
  } catch(e) {
    Logger.log("avisoSegurancaShopee: ERRO — " + e.message);
  }
}

function avisoSegurancaMagalu() {
  try {
    const now = new Date();
    if (now.getDay() === 0 || now.getDay() === 6) return;
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (props.getProperty("AVISO_MAGALU_DATE") === today) return;

    const spacesRaw = String(props.getProperty("CHAT_SPACE_MAGALU") || "").trim();
    if (!spacesRaw) { Logger.log("avisoSegurancaMagalu: CHAT_SPACE_MAGALU não configurado."); return; }

    const ssSrc = SpreadsheetApp.openById(ML_IMPORT_CONFIG.STATUS_SOURCE_SPREADSHEET_ID);
    const pendentes = [];
    const seenIds   = new Set();

    for (const sheetName of ["Junho", "Maio"]) {
      const sh = ssSrc.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) continue;
      const n      = sh.getLastRow() - 1;
      const values = sh.getRange(2, 1, n, 7).getValues();
      const bgs    = sh.getRange(2, 7, n, 1).getBackgrounds();
      for (let i = 0; i < n; i++) {
        const colC = String(values[i][2] ?? "").toLowerCase();
        if (!colC.includes("magalu")) continue;
        const colG = String(values[i][6] ?? "").trim().toLowerCase();
        const bgG  = String(bgs[i][0]    ?? "").trim().toLowerCase();
        if (colG.includes("ok") || bgG === ML_IMPORT_CONFIG.OK_GREEN.toLowerCase()) continue;
        if (colG.includes("cancelado")) continue;
        if (colG.includes("full") || colG.includes("fulfil")) continue;
        const id      = String(values[i][1] ?? "").trim();
        const cliente = String(values[i][3] ?? "").trim();
        if (!id || seenIds.has(id)) continue;
        seenIds.add(id);
        pendentes.push({ id, cliente });
      }
    }

    props.setProperty("AVISO_MAGALU_DATE", today);

    if (!pendentes.length) { Logger.log("avisoSegurancaMagalu: nenhum pedido Magalu pendente."); return; }

    const text =
      "⚠ *Magalu* — Pedidos Pendentes às 13:00\n" +
      "Ainda há " + pendentes.length + " pedido(s) pendente(s). Verificar se já foram enviados ou preparados:\n" +
      pendentes.map(p => "• " + p.id + " — " + p.cliente).join("\n");

    const normalize = id => id.startsWith("spaces/") ? id : "spaces/" + id;
    for (const spaceId of spacesRaw.split(",").map(s => s.trim()).filter(Boolean)) {
      try { callChatApi_("post", normalize(spaceId) + "/messages", { text }); }
      catch(e) { Logger.log("avisoSegurancaMagalu chat err: " + e.message); }
    }
    Logger.log("avisoSegurancaMagalu: mensagem enviada. " + pendentes.length + " pedido(s).");
  } catch(e) {
    Logger.log("avisoSegurancaMagalu: ERRO — " + e.message);
  }
}

function criarGatilhoAvisoShopee() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "avisoSegurancaShopee")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("avisoSegurancaShopee").timeBased().everyDays(1).atHour(14).create();
  Logger.log("Gatilho criado: avisoSegurancaShopee às 14h.");
}

function criarGatilhoAvisoMagalu() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "avisoSegurancaMagalu")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("avisoSegurancaMagalu").timeBased().everyDays(1).atHour(13).create();
  Logger.log("Gatilho criado: avisoSegurancaMagalu às 13h.");
}
