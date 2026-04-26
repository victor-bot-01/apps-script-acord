/*******************************
 * CONFIG
 *******************************/
const CONFIG = {
  DEFAULT_SOURCE: "ML Coleta",
  SHEET_NAMES: ["ML Coleta", "ML 1"],

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
    const perSheet = {};

    for (const name of sheetsToRead) {
      const out = readFromSheet_(ss, name);
      perSheet[name] = { pedidos: out.pedidos.length, faltam: out.faltamItems.length };

      pedidos.push(...out.pedidos);
      faltamItems.push(...out.faltamItems);
    }

    // Ordena pedidos por ID (ajuda no visual e multi-item)
    pedidos.sort((a, b) => String(a.id).localeCompare(String(b.id)));

    // FaltamItems: mantém ordem da planilha (ou ordena por produto se quiser)
    // faltamItems.sort((a,b)=>String(a.produtoPendentes).localeCompare(String(b.produtoPendentes)));

    const kpis = buildKpis_(pedidos, faltamItems);

    return {
      ok: true,
      source: src,
      pedidos,
      faltamItems,
      kpis,
      collectionTimes: getCollectionTimes(),
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
  if (!sh) return { pedidos: [], faltamItems: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { pedidos: [], faltamItems: [] };

  const lastCol = CONFIG.COLS.QUANTIDADE; // 10

  const range = sh.getRange(
    CONFIG.DATA_START_ROW,
    1,
    lastRow - CONFIG.DATA_START_ROW + 1,
    lastCol
  );

  const values = range.getValues();

  const pedidos = [];
  const faltamItems = [];

  for (const r of values) {
    // ====== FALTAM LIST (independente) ======
    const produtoPendentes = String(r[CONFIG.COLS.PRODUTOS_PENDENTES - 1] ?? "").trim();
    const quantidade = parseNumber_(r[CONFIG.COLS.QUANTIDADE - 1]);

    // Só entra se tiver produto pendente (o principal)
    if (produtoPendentes !== "") {
      faltamItems.push({
        produtoPendentes,
        quantidade,
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

  return { pedidos, faltamItems };
}

/*******************************
 * KPI LOGIC
 * - Total/Pendentes/Confirmados/Cancelados: IDs únicos em pedidos
 * - Bipados: IDs com bipado PRONTO
 * - Etiquetas Impressas: IDs com etiquetas IMPRESSA
 * - Faltam: soma Quantidade da lista independente faltamItems
 *******************************/
function buildKpis_(pedidos, faltamItems) {
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

  return {
    total: idsTotal.size,
    pendentes: idsPendentes.size,
    confirmados: idsConfirmados.size,
    cancelados: idsCancelados.size,
    bipadosPronto: idsBipadosPronto.size,
    etiquetasImpressas: idsEtiquetasImpressa.size,
    faltam: faltam, // soma
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

/*******************************
 * COLLECTION TIMES
 *******************************/
function getCollectionTimes() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty("COLLECTION_TIMES");
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
    props.setProperty("COLLECTION_TIMES", JSON.stringify(times));
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
