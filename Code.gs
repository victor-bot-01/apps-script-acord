/*******************************
 * CONFIG
 *******************************/
const CONFIG = {
  DEFAULT_SOURCE: "ML Coleta",
  SHEET_NAMES: ["ML Coleta", "ML 1", "Shopee", "Magalu", "Essência do Brasil", "Amazon", "Flex/Vapt"],

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
      ? CONFIG.SHEET_NAMES.slice()
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

    const kpis = buildKpis_(pedidos, faltamItems, todosNvItems);

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

    return {
      ok: true,
      source: src,
      pedidos,
      faltamItems,
      todosNvItems,
      kpis,
      collectionTimes: getCollectionTimes(),
      hasBlueCells,
      hasReviewCells,
      monitorSnapshot,
      monitorPaused,
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

  const lastCol = CONFIG.COLS.TODOS_NV_QTD; // 14

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
      sheet: sheetName
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
function getOrdersByComponent(componentName) {
  try {
    const mapSubseq = pend_buildDadosMapSubseq_();
    const rev = bm_buildReverseMap_(mapSubseq);
    const normComp = pend_norm_(componentName);
    const kitKeys  = rev.get(normComp) || [];

    const conf = bm_readConferencia_();
    const matchIds = new Set();

    for (let i = 0; i < conf.rows; i++) {
      const status = String(conf.statuses[i][0] ?? "").trim();
      if (status !== "") continue;
      const pk = pend_norm_(String(conf.prods[i][0] ?? "").trim());
      const match = kitKeys.length ? kitKeys.includes(pk) : pk === normComp;
      if (!match) continue;
      const id = String(conf.ids[i][0] ?? "").trim();
      if (id) matchIds.add(id);
    }

    // Resolve source (aba do dashboard) por ID
    const dashSS = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["ML Coleta","ML 1","Shopee","Magalu","Essência do Brasil","Amazon","Flex/Vapt"];
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

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let updated = 0;

    for (const item of items) {
      const normKit = pend_norm_(item.kitProduct);
      const orderId = String(item.orderId).trim();
      for (let i = 0; i < conf.rows; i++) {
        if (String(conf.ids[i][0] ?? "").trim() !== orderId) continue;
        if (pend_norm_(String(conf.prods[i][0] ?? "").trim()) !== normKit) continue;
        if (String(conf.statuses[i][0] ?? "").toUpperCase().includes("TENHO")) continue;

        const compStatus = ordKitStatus[orderId + "||" + item.kitProduct] || {};
        const hasTenhoSibling = Object.entries(compStatus)
          .some(([c, st]) => c !== item.componentName && st === "TENHO");

        if (hasTenhoSibling) {
          const nonTenhoComps = [
            item.componentName,
            ...Object.entries(compStatus)
              .filter(([c, st]) => c !== item.componentName && st !== "TENHO")
              .map(([c]) => c)
          ];
          newSt[i][0] = "FALTA - " + nonTenhoComps.join(";");
        } else {
          newSt[i][0] = "FALTA";
        }
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

    const conf  = bm_readConferencia_();
    const newSt = conf.statuses.map(r => [r[0]]);
    let updated = 0;

    for (const item of items) {
      const normKit = pend_norm_(item.kitProduct);
      const orderId = String(item.orderId).trim();
      for (let i = 0; i < conf.rows; i++) {
        if (String(conf.ids[i][0]  ?? "").trim() !== orderId)  continue;
        if (pend_norm_(String(conf.prods[i][0] ?? "").trim()) !== normKit) continue;
        if (String(conf.statuses[i][0] ?? "").toUpperCase().includes("TENHO")) continue;

        if (tenhoSet.has(orderId)) {
          const compStatus      = ordKitStatus[orderId + "||" + item.kitProduct] || {};
          const missingSiblings = Object.entries(compStatus)
            .filter(([c, st]) => c !== item.componentName && st !== "TENHO")
            .map(([c]) => c);
          if (missingSiblings.length === 0) continue;
          newSt[i][0] = "FALTA - " + missingSiblings.join(";");
        } else {
          newSt[i][0] = "FALTA";
        }
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
