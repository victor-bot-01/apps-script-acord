/******************* CONFIG *******************/
const ML_IMPORT_CONFIG = {
  // Pastas (XLSX mais recente)
  FOLDERS: [
    { folderId: "1L1wISlKCaTgZ713TlVI16qN6u20KV2vv", targetSheetName: "ML Coleta" },
    { folderId: "1TxvlZusR0ilCNjyUmMV_yZsomjDZ7y8d", targetSheetName: "ML 1" }
  ],

  // Planilha STATUS (Març/Fev)
  STATUS_SOURCE_SPREADSHEET_ID: "1RHTxClkkFamZUE4hM8romnZfWWTecNwvkvI4SO7rYmU",
  STATUS_SOURCE_SHEETS: ["Abril", "Março"],

  // Planilha ponte ML1 -> ID do STATUS / ANDAMENTO (via Código+Cliente)
  ML1_LOOKUP_SPREADSHEET_ID: "1KCTrbN6Jx9sg1H-ho6eIZ4we1rm5HL0mjd6OtlpWMMA",

  // Planilha ANDAMENTO
  ANDAMENTO_SOURCE_SPREADSHEET_ID: "1eXC8YL-3SXsMKc7ICIa-WqvX_yi3RARY8ouer3yQxrQ",

  // Cores
  OK_GREEN: "#00ff00",
  YELLOW: "#ffff00",
  RED: "#ff0000",
  MULTI_GRAY: "#cfcfcf",

  // Colunas destino (aba ML Coleta / ML 1)
  COL_ID: 1,        // A
  COL_CLIENTE: 2,   // B
  COL_PRODUTO: 3,   // C
  COL_QTD: 4,       // D
  COL_STATUS: 5,    // E
  COL_ETIQUETAS: 7, // G
  COL_ANDAMENTO: 8, // H
  COL_HORARIO: 11,  // K
  COL_CODIGO: 12    // L (Código = SKU)
};

/******************* NORMALIZAÇÃO (case-insensitive) *******************/
function normKeyPart_(v) {
  return (v ?? "").toString().trim().toLowerCase();
}
function makeKeyCodigoCliente_(codigo, cliente) {
  return `${normKeyPart_(codigo)}||${normKeyPart_(cliente)}`;
}

/******************* ENTRY *******************/
function importarML_Coleta_e_ML1_v3(props, runId, TTL_MS) {
  const hb = () => {
    if (props && runId && TTL_MS) heartbeat_(props, runId, TTL_MS);
  };

  hb();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ✅ Limpa AZUIS antigos na coluna C (Produto) antes de tudo (mesmo se não importar)
  clearBlueProdutoColC_(ss.getSheetByName("ML Coleta"));
  clearBlueProdutoColC_(ss.getSheetByName("ML 1"));

  // 0) Critério de limpeza/import:
  // - Se a pasta tiver XLSX: limpa e importa
  // - Se não tiver XLSX: NÃO limpa/NÃO importa
  // IMPORTANTE: mesmo que não tenha XLSX em uma ou nas duas pastas,
  // o script continua e ainda recalcula STATUS/ANDAMENTO/PENDENTES para as abas existentes.
  const didImport = {}; // { "ML Coleta": true/false, "ML 1": true/false }

  for (const cfg of ML_IMPORT_CONFIG.FOLDERS) {
    const sh = ss.getSheetByName(cfg.targetSheetName);
    if (!sh) throw new Error(`Aba destino não encontrada: ${cfg.targetSheetName}`);

    let hasXlsx = false;
    try {
      hasXlsx = pastaTemXlsx_(cfg.folderId);
    } catch (e) {
      Logger.log(`Erro ao checar XLSX na pasta ${cfg.folderId}: ${e.message}`);
      hasXlsx = false;
    }

    if (hasXlsx) {
      limparAbaExcetoCabecalho_(sh);
      hb();
      didImport[cfg.targetSheetName] = importarUltimoXlsxDaPastaParaAba_(cfg.folderId, cfg.targetSheetName, hb);
    } else {
      Logger.log(`Sem XLSX na pasta (${cfg.folderId}). NÃO limpar e NÃO importar em '${cfg.targetSheetName}'.`);
      didImport[cfg.targetSheetName] = false;
    }
  }

  hb();

  // 1) STATUS
  const statusById = construirStatusById_(hb);

  aplicarStatus_MLColeta_porId_(statusById, hb);
  aplicarStatus_porAba_(statusById, "ML 1");

  hb();

  // 2) ANDAMENTO — NOVA LÓGICA
  const andamentoByIdQueue = construirFilaAndamentoPorId_(hb);

  aplicarAndamento_MLColeta_porIdQueue_(andamentoByIdQueue, hb);
  aplicarAndamento_porAba_(andamentoByIdQueue, "ML 1");

  hb();

  // ✅ 3) PENDENTES (Produtos Pendentes / Quantidade em I/J)
  aplicarPendentes_MLColeta_e_ML1_(hb);

  hb();

  // ✅ 4) MONITOR: registrar data de importação e enviar status inicial
  const today_ = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (didImport["ML Coleta"]) {
    try {
      PropertiesService.getScriptProperties().setProperty("LAST_IMPORT_DATE_ML_COLETA", today_);
      enviarStatusMonitor_("ML Coleta");
    } catch (err) { Logger.log("Aviso: erro no monitor ML Coleta: " + err.message); }
  }
  if (didImport["ML 1"]) {
    try {
      PropertiesService.getScriptProperties().setProperty("LAST_IMPORT_DATE_ML_1", today_);
      enviarStatusMonitor_("ML 1");
    } catch (err) { Logger.log("Aviso: erro no monitor ML 1: " + err.message); }
  }
  hb();

  Logger.log("Processo completo: limpeza/import condicional + status + andamento (novo) + pendentes.");
}

/******************* WRAPPER MANUAL *******************/
function executar_importarML_Coleta_e_ML1_v3() {
  importarML_Coleta_e_ML1_v3(null, null, null);
}

/******************* CHECK: pasta tem XLSX? *******************/
function pastaTemXlsx_(folderId) {
  const pasta = DriveApp.getFolderById(folderId);
  const arquivos = pasta.getFilesByType(MimeType.MICROSOFT_EXCEL);
  return arquivos.hasNext();
}

/******************* LIMPAR ABA (exceto cabeçalho) *******************/
function limparAbaExcetoCabecalho_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getMaxColumns();
  if (lastRow <= 1) return;

  const range = sh.getRange(2, 1, lastRow - 1, lastCol);
  range.clearContent();
  range.clearFormat(); // inclui bordas e preenchimento
  range.setBackground(null);
  range.setBorder(false, false, false, false, false, false);
}

/******************* IMPORTAR XLSX MAIS RECENTE *******************/
/**
 * Retorna true se importou; false se não havia XLSX/linhas úteis.
 */
function importarUltimoXlsxDaPastaParaAba_(folderId, targetSheetName, hb) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName(targetSheetName);
  if (!abaDestino) throw new Error(`Aba destino '${targetSheetName}' não encontrada.`);

  const pasta = DriveApp.getFolderById(folderId);
  const arquivos = pasta.getFilesByType(MimeType.MICROSOFT_EXCEL);

  let arquivoMaisRecente = null;
  let dataMaisRecente = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const t = arquivo.getLastUpdated().getTime();
    if (t > dataMaisRecente) {
      dataMaisRecente = t;
      arquivoMaisRecente = arquivo;
    }
  }

  if (!arquivoMaisRecente) {
    Logger.log(`Nenhum XLSX encontrado na pasta ${folderId} para a aba ${targetSheetName}.`);
    return false;
  }

  hb();

  // Converter XLSX -> Google Sheets (Drive API avançada habilitada)
  const blob = arquivoMaisRecente.getBlob();
  const arquivoConvertido = DriveApi.Files.insert(
    { title: "TEMP_CONVERTIDO_" + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );
  const tempId = arquivoConvertido.id;

  hb();

  const planilhaTemp = SpreadsheetApp.openById(tempId);
  const abaOrigem = planilhaTemp.getSheets()[0];
  const dados = abaOrigem.getDataRange().getValues();

  hb();

  // Cabeçalho linha 6 (index 5), dados a partir da linha 7 (index 6)
  const cabecalho = dados[5] || [];

  const idxVenda = cabecalho.indexOf("N.º de venda");
  const idxComprador = cabecalho.indexOf("Comprador");
  const idxTitulo = cabecalho.indexOf("Título do anúncio");
  const idxUnidades = cabecalho.indexOf("Unidades");
  const idxEstado = cabecalho.indexOf("Estado");
  const idxDescricaoStatus = cabecalho.indexOf("Descrição do status");
  const idxSku = cabecalho.indexOf("SKU"); // Código

  if ([idxVenda, idxComprador, idxTitulo, idxUnidades, idxEstado, idxDescricaoStatus, idxSku].some(i => i === -1)) {
    DriveApp.getFileById(tempId).setTrashed(true);
    throw new Error(
      `Colunas obrigatórias não encontradas no XLSX (${targetSheetName}). ` +
      `Esperado: N.º de venda, Comprador, Título do anúncio, Unidades, Estado, Descrição do status, SKU.`
    );
  }

  // Etiquetas: normaliza baseado no texto (conter termos)
  const normalizarEtiqueta_ = (txt) => {
    const s = (txt ?? "").toString().trim().toLowerCase();
    if (!s) return "Para Imprimir";
    if (s.includes("etiqueta impressa") || s.includes("pronto para coleta")) return "Impressa";
    if (s.includes("cancelado") || s.includes("cancelada")) return "Cancelado";
    return "Para Imprimir";
  };

  const registros = [];
  const linhasMultiploItem = new Set();

  let ultimoNumeroVenda = "";
  let ultimoComprador = "";
  let linhaPrincipalIndex = null;

  for (let i = 6; i < dados.length; i++) {
    if (i % 300 === 0) hb();

    const linha = dados[i];

    const numeroVendaAtual = (linha[idxVenda] ?? "").toString().trim();
    const compradorAtual = (linha[idxComprador] ?? "").toString().trim();

    if (compradorAtual) {
      ultimoNumeroVenda = numeroVendaAtual;
      ultimoComprador = compradorAtual;
      linhaPrincipalIndex = registros.length;
    } else {
      if (linhaPrincipalIndex !== null) linhasMultiploItem.add(linhaPrincipalIndex);
    }

    const idFinal = compradorAtual ? numeroVendaAtual : ultimoNumeroVenda;
    const compradorFinal = compradorAtual || ultimoComprador;

    const titulo = (linha[idxTitulo] ?? "").toString().trim();
    if (!titulo) continue; // Produto vazio -> elimina

    const unidades = linha[idxUnidades];
    const etiqueta = normalizarEtiqueta_(linha[idxEstado]);
    const horario = (linha[idxDescricaoStatus] ?? "").toString().trim();
    const codigo = (linha[idxSku] ?? "").toString().trim();

    registros.push({ id: idFinal, cliente: compradorFinal, produto: titulo, qtd: unidades, etiqueta, horario, codigo });

    if (!compradorAtual && linhaPrincipalIndex !== null) linhasMultiploItem.add(registros.length - 1);
  }

  hb();

  if (registros.length === 0) {
    DriveApp.getFileById(tempId).setTrashed(true);
    DriveApp.getFileById(arquivoMaisRecente.getId()).setTrashed(true);
    Logger.log(`Nada para inserir em ${targetSheetName} (tudo filtrado).`);
    return false;
  }

  const inicio = 2;
  const n = registros.length;

  // A:D
  abaDestino.getRange(inicio, 1, n, 4).setValues(
    registros.map(r => [r.id, r.cliente, r.produto, r.qtd])
  );

  hb();

  // G (Etiquetas)
  abaDestino.getRange(inicio, ML_IMPORT_CONFIG.COL_ETIQUETAS, n, 1).setValues(
    registros.map(r => [r.etiqueta])
  );

  // K (Horário)
  abaDestino.getRange(inicio, ML_IMPORT_CONFIG.COL_HORARIO, n, 1).setValues(
    registros.map(r => [r.horario])
  );

  // L (Código / SKU)
  abaDestino.getRange(inicio, ML_IMPORT_CONFIG.COL_CODIGO, n, 1).setValues(
    registros.map(r => [r.codigo])
  );

  hb();

  // Remover bordas do bloco escrito
  const lastCol = abaDestino.getMaxColumns();
  abaDestino.getRange(inicio, 1, n, lastCol).setBorder(false, false, false, false, false, false);

  // Cinza para múltiplos itens (A..L)
  linhasMultiploItem.forEach(idx => {
    const row = inicio + idx;
    abaDestino.getRange(row, 1, 1, ML_IMPORT_CONFIG.COL_CODIGO).setBackground(ML_IMPORT_CONFIG.MULTI_GRAY);
  });

  hb();

  // Corrigir notação científica na coluna A
  const colA = abaDestino.getRange(inicio, 1, n, 1);
  const valsA = colA.getValues();
  valsA.forEach((r, i) => {
    if (i % 400 === 0) hb();
    const v = r[0];
    if (typeof v === "number" && v.toExponential().includes("e")) {
      abaDestino.getRange(inicio + i, 1).setValue("'" + v.toFixed(0));
    }
  });

  hb();

  // Limpar temporário + apagar XLSX original (intencional)
  DriveApp.getFileById(tempId).setTrashed(true);
  DriveApp.getFileById(arquivoMaisRecente.getId()).setTrashed(true);

  Logger.log(`Importado ${n} linhas em '${targetSheetName}' a partir do XLSX mais recente.`);
  return true;
}

/******************* STATUS: construir mapa por ID direto no STATUS *******************/
function construirStatusById_(hb) {
  const byId = new Map(); // id -> "Cancelado" | "Confirmado" | "Pendente"
  const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.STATUS_SOURCE_SPREADSHEET_ID);

  ML_IMPORT_CONFIG.STATUS_SOURCE_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const lastCol = sh.getLastColumn();
    const numRows = lastRow - 1;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const idxPedido = header.indexOf("Pedido");
    const idxId = header.indexOf("ID");
    const idxVenda = header.indexOf("N.º de venda");

    const idCol = (idxPedido !== -1 ? idxPedido + 1 :
                  (idxId !== -1 ? idxId + 1 :
                  (idxVenda !== -1 ? idxVenda + 1 : 2))); // default B

    const statusCol = 7; // G

    const ids = sh.getRange(2, idCol, numRows, 1).getValues();
    const statusRange = sh.getRange(2, statusCol, numRows, 1);
    const statusVals = statusRange.getValues();
    const statusBgs = statusRange.getBackgrounds();

    for (let i = 0; i < numRows; i++) {
      if (i % 800 === 0) hb();

      const id = (ids[i][0] ?? "").toString().trim();
      if (!id) continue;

      const txt = (statusVals[i][0] ?? "").toString().trim().toLowerCase();
      const bg = (statusBgs[i][0] ?? "").toString().trim().toLowerCase();

      const isCancel = txt.includes("cancelado") || txt.includes("cancelada");
      const isOk = (txt === "ok") || (bg === ML_IMPORT_CONFIG.OK_GREEN);

      const atual = byId.get(id);
      // prioridade: Cancelado > Confirmado > Pendente
      if (isCancel) byId.set(id, "Cancelado");
      else if (isOk) {
        if (atual !== "Cancelado") byId.set(id, "Confirmado");
      } else {
        if (!atual) byId.set(id, "Pendente");
      }
    }
  });

  return byId;
}

/******************* STATUS ML COLETA: por ID direto no STATUS *******************/
function aplicarStatus_MLColeta_porId_(statusById, hb) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("ML Coleta");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const ids = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID, n, 1).getValues();
  const etiquetas = sh.getRange(2, ML_IMPORT_CONFIG.COL_ETIQUETAS, n, 1).getValues();

  const out = ids.map((r, i) => {
    if (i % 900 === 0) hb();
    const id = (r[0] ?? "").toString().trim();
    const etiqueta = (etiquetas[i][0] ?? "").toString().trim();
    if (etiqueta === "Cancelado") return ["Cancelado"];
    return [statusById.get(id) || "Pendente"];
  });

  sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS, n, 1).setValues(out);
}

/******************* ANDAMENTO (NOVO) — construir fila por ID (A -> F) *******************/
function construirFilaAndamentoPorId_(hb) {
  // id -> [v1, v2, ...] (ordem de leitura; shift na aplicação)
  const byIdQueue = new Map();
  const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.ANDAMENTO_SOURCE_SPREADSHEET_ID);

  ss.getSheets().forEach(sh => {
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 6) return;

    const numRows = lastRow - 1;

    // A = ID, F = Valor
    const ids = sh.getRange(2, 1, numRows, 1).getValues();
    const valsF = sh.getRange(2, 6, numRows, 1).getValues();

    for (let i = 0; i < numRows; i++) {
      if (i % 1200 === 0) hb();

      const id = (ids[i][0] ?? "").toString().trim();
      if (!id) continue;

      const v = valsF[i][0];
      if (!byIdQueue.has(id)) byIdQueue.set(id, []);
      byIdQueue.get(id).push(v);
    }
  });

  return byIdQueue;
}

/******************* ANDAMENTO ML COLETA: por ID (fila) *******************/
function aplicarAndamento_MLColeta_porIdQueue_(byIdQueue, hb) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("ML Coleta");
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;

  const ids = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID, n, 1).getValues();
  const statusVals = sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS, n, 1).getValues();

  const out = ids.map((r, i) => {
    if (i % 900 === 0) hb();
    const status = (statusVals[i][0] ?? "").toString().trim();
    if (status === "Confirmado") return [""];
    const id = (r[0] ?? "").toString().trim();
    const fila = byIdQueue.get(id);
    return [(fila && fila.length) ? fila.shift() : ""];
  });

  sh.getRange(2, ML_IMPORT_CONFIG.COL_ANDAMENTO, n, 1).setValues(out);
}

/* =========================================================================
   PENDENTES (I/J) + AZUL (col C) — integrado ao script principal
   Regras:
   - Limpa azuis antigos na col C antes de rodar (feito no início do entry)
   - Azul na col C SOMENTE quando:
       ID teve FALTA na Conferencia E não gerou item elegível (não bateu Dados/termos)
   - Gera lista agregada em I/J (1 item por linha) para ML Coleta e ML 1
   ========================================================================= */

const PEND_BLUE = "#0000ff";
const PEND_OUT_COL_PROD = 9; // I
const PEND_OUT_COL_QTD  = 10; // J
const PEND_NV_COL_PROD  = 13; // M — Falta/Não Verificado
const PEND_NV_COL_QTD   = 14; // N
const PEND_DADOS_MAX_COL = 26; // Z

// ANDAMENTO > Conferencia (fixo)
const PEND_CONF_COL_ID     = 1; // A Pedido
const PEND_CONF_COL_PROD   = 4; // D Produto
const PEND_CONF_COL_QTD    = 5; // E Qtd
const PEND_CONF_COL_STATUS = 6; // F Status

function aplicarPendentes_MLColeta_e_ML1_(hb) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shColeta = ss.getSheetByName("ML Coleta");
  const shML1    = ss.getSheetByName("ML 1");
  if (!shColeta || !shML1) return;

  const mapSubseq = pend_buildDadosMapSubseq_();
  const confById  = pend_buildConferenciaById_();

  pend_process_MLColeta_(shColeta, confById, mapSubseq);
  pend_process_porAba_(shML1, confById, mapSubseq);

  // Falta/Não Verificado (M/N)
  const confTodosById = pend_buildTodosExcTenho_();
  pend_process_nv_porId_(shColeta, confTodosById, mapSubseq);
  pend_process_nv_porId_(shML1, confTodosById, mapSubseq);
}

/************** PEND: limpar azul na col C **************/
function clearBlueProdutoColC_(sh) {
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const r = sh.getRange(2, ML_IMPORT_CONFIG.COL_PRODUTO, n, 1);
  const bgs = r.getBackgrounds();

  let changed = false;
  for (let i = 0; i < n; i++) {
    if (String(bgs[i][0] || "").toLowerCase() === String(PEND_BLUE).toLowerCase()) {
      bgs[i][0] = "#ffffff";
      changed = true;
    }
  }
  if (changed) r.setBackgrounds(bgs);
}

/************** PEND: Dados (produto -> lista) **************/
function pend_buildDadosMapSubseq_() {
  const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.ANDAMENTO_SOURCE_SPREADSHEET_ID);
  const sh = ss.getSheetByName("Dados");
  if (!sh) throw new Error('Aba "Dados" não encontrada na planilha ANDAMENTO.');

  const lastRow = sh.getLastRow();
  const lastCol = Math.min(sh.getLastColumn(), PEND_DADOS_MAX_COL);

  const map = new Map();
  if (lastRow < 1 || lastCol < 2) return map;

  const vals = sh.getRange(1, 1, lastRow, lastCol).getValues();

  for (let r = 0; r < vals.length; r++) {
    const row = vals[r];
    for (let c = 0; c < row.length - 1; c++) {
      const left = String(row[c] ?? "").trim();
      if (!left) continue;

      const key = pend_norm_(left);
      if (map.has(key)) continue;

      const right = String(row[c + 1] ?? "").trim();
      if (!right) continue;

      map.set(key, right);
    }
  }

  return map;
}

/************** PEND: Conferencia (somente linhas com FALTA) por ID **************/
function pend_buildConferenciaById_() {
  const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.ANDAMENTO_SOURCE_SPREADSHEET_ID);
  const sh = ss.getSheetByName("Conferencia");
  if (!sh) throw new Error('Aba "Conferencia" não encontrada na planilha ANDAMENTO.');

  const lastRow = sh.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const n = lastRow - 1;

  const ids   = sh.getRange(2, PEND_CONF_COL_ID, n, 1).getValues();
  const prods = sh.getRange(2, PEND_CONF_COL_PROD, n, 1).getValues();
  const qtds  = sh.getRange(2, PEND_CONF_COL_QTD, n, 1).getValues();
  const stats = sh.getRange(2, PEND_CONF_COL_STATUS, n, 1).getValues();

  for (let i = 0; i < n; i++) {
    const id = String(ids[i][0] ?? "").trim();
    if (!id) continue;

    const stRaw = String(stats[i][0] ?? "").trim();
    if (!stRaw.toUpperCase().includes("FALTA")) continue;

    const prod = String(prods[i][0] ?? "").trim();
    if (!prod) continue;

    const qty = pend_parseQtd_(qtds[i][0]);
    const terms = pend_extractTerms_(stRaw);

    if (!map.has(id)) map.set(id, []);
    map.get(id).push({ prod, qty, terms });
  }

  return map;
}

/************** PEND: Conferencia (tudo exceto TENHO) por ID **************/
function pend_buildTodosExcTenho_() {
  const ss = SpreadsheetApp.openById(ML_IMPORT_CONFIG.ANDAMENTO_SOURCE_SPREADSHEET_ID);
  const sh = ss.getSheetByName("Conferencia");
  if (!sh) throw new Error('Aba "Conferencia" não encontrada na planilha ANDAMENTO.');

  const lastRow = sh.getLastRow();
  const map = new Map();
  if (lastRow < 2) return map;

  const n = lastRow - 1;

  const ids   = sh.getRange(2, PEND_CONF_COL_ID,     n, 1).getValues();
  const prods = sh.getRange(2, PEND_CONF_COL_PROD,   n, 1).getValues();
  const qtds  = sh.getRange(2, PEND_CONF_COL_QTD,    n, 1).getValues();
  const stats = sh.getRange(2, PEND_CONF_COL_STATUS,  n, 1).getValues();

  for (let i = 0; i < n; i++) {
    const id = String(ids[i][0] ?? "").trim();
    if (!id) continue;

    const stRaw = String(stats[i][0] ?? "").trim();
    const stUp = stRaw.toUpperCase();
    if (stUp.includes("TENHO") || stUp.includes("FALTA")) continue;

    const prod = String(prods[i][0] ?? "").trim();
    if (!prod) continue;

    const qty   = pend_parseQtd_(qtds[i][0]);
    const terms = pend_extractTerms_(stRaw);

    if (!map.has(id)) map.set(id, []);
    map.get(id).push({ prod, qty, terms });
  }

  return map;
}

/************** PEND: Process ML Coleta **************/
function pend_process_MLColeta_(sh, confById, mapSubseq) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const idVals = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID, n, 1).getValues();
  const statusVals = sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS, n, 1).getValues();

  // limpa I/J
  pend_clearIJ_(sh);

  // backgrounds coluna C
  const rangeC = sh.getRange(2, ML_IMPORT_CONFIG.COL_PRODUTO, n, 1);
  const bgC = rangeC.getBackgrounds();
  const newBgC = bgC.map(r => [r[0]]);
  let anyBg = false;

  // IDs únicos (ignora linhas Confirmadas)
  const idsUnique = [];
  const seen = new Set();
  for (let i = 0; i < n; i++) {
    const id = String(idVals[i][0] ?? "").trim();
    if (!id) continue;
    const status = String(statusVals[i][0] ?? "").trim();
    if (status === "Confirmado") continue;
    if (!seen.has(id)) { seen.add(id); idsUnique.push(id); }
  }

  const counts = new Map();
  const perId = new Map(); // id -> {hadFalta, generated}

  for (const id of idsUnique) {
    const linhas = confById.get(id) || [];
    if (!linhas.length) {
      perId.set(id, { hadFalta: false, generated: false });
      continue;
    }

    let gen = false;
    for (const it of linhas) {
      const r = pend_processFaltaLine_(it, mapSubseq, counts);
      if (r.added > 0) gen = true;
    }
    perId.set(id, { hadFalta: true, generated: gen });
  }

  // azul somente se teve FALTA e não gerou item elegível
  for (let i = 0; i < n; i++) {
    const id = String(idVals[i][0] ?? "").trim();
    if (!id) continue;

    const pr = perId.get(id);
    if (pr && pr.hadFalta === true && pr.generated === false) {
      newBgC[i][0] = PEND_BLUE;
      anyBg = true;
    }
  }

  // escreve I/J
  const rowsOut = pend_countsToRows_(counts);
  if (rowsOut.length) {
    pend_ensureRows_(sh, 1 + rowsOut.length);
    sh.getRange(2, PEND_OUT_COL_PROD, rowsOut.length, 2).setValues(rowsOut);
  }

  if (anyBg) rangeC.setBackgrounds(newBgC);
}

/************** PEND: Process por aba genérica (Shopee, Magalu, Essência do Brasil) **************/
function pend_process_porAba_(sh, confById, mapSubseq) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const idVals     = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID,     n, 1).getValues();
  const statusVals = sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS,  n, 1).getValues();

  pend_clearIJ_(sh);

  const rangeC = sh.getRange(2, ML_IMPORT_CONFIG.COL_PRODUTO, n, 1);
  const bgC    = rangeC.getBackgrounds();
  const newBgC = bgC.map(r => [r[0]]);
  let anyBg    = false;

  const idsUnique = [];
  const seen = new Set();
  for (let i = 0; i < n; i++) {
    const id     = String(idVals[i][0]     ?? "").trim();
    const status = String(statusVals[i][0] ?? "").trim();
    if (!id) continue;
    if (status === "Confirmado") continue;
    if (!seen.has(id)) { seen.add(id); idsUnique.push(id); }
  }

  const counts = new Map();
  const perId  = new Map();

  for (const id of idsUnique) {
    const linhas = confById.get(id) || [];
    if (!linhas.length) {
      perId.set(id, { hadFalta: false, generated: false });
      continue;
    }
    let gen = false;
    for (const it of linhas) {
      const r = pend_processFaltaLine_(it, mapSubseq, counts);
      if (r.added > 0) gen = true;
    }
    perId.set(id, { hadFalta: true, generated: gen });
  }

  for (let i = 0; i < n; i++) {
    const id = String(idVals[i][0] ?? "").trim();
    if (!id) continue;
    const pr = perId.get(id);
    if (pr && pr.hadFalta === true && pr.generated === false) {
      newBgC[i][0] = PEND_BLUE;
      anyBg = true;
    }
  }

  const rowsOut = pend_countsToRows_(counts);
  if (rowsOut.length) {
    pend_ensureRows_(sh, 1 + rowsOut.length);
    sh.getRange(2, PEND_OUT_COL_PROD, rowsOut.length, 2).setValues(rowsOut);
  }

  if (anyBg) rangeC.setBackgrounds(newBgC);
}

/************** PEND: core (linha com FALTA) **************/
function pend_processFaltaLine_(it, mapSubseq, counts) {
  const prod = String(it.prod ?? "").trim();
  const terms = Array.isArray(it.terms) ? it.terms : [];
  const mult = Number(it.qty || 1);

  if (!prod) return { added: 0 };

  const key = pend_norm_(prod);
  const subseq = mapSubseq.get(key);
  if (!subseq) return { added: 0 };

  const partsAll = subseq.split(";").map(s => s.trim()).filter(Boolean);
  if (!partsAll.length) return { added: 0 };

  let partsToCount = partsAll;

  if (terms.length) {
    const matched = [];
    const matchedSet = new Set();

    for (const p of partsAll) {
      const pNorm = pend_norm_(p);
      for (const term of terms) {
        const needle = pend_norm_(term);
        if (needle && pNorm.includes(needle)) {
          if (!matchedSet.has(p)) {
            matchedSet.add(p);
            matched.push(p);
          }
          break;
        }
      }
    }

    partsToCount = matched;
    if (!partsToCount.length) return { added: 0 };
  }

  let added = 0;
  for (const p of partsToCount) {
    counts.set(p, (counts.get(p) || 0) + mult);
    added++;
  }

  return { added };
}

/************** PEND: utils **************/
function pend_clearIJ_(sh) {
  const max = sh.getMaxRows();
  if (max >= 2) sh.getRange(2, PEND_OUT_COL_PROD, max - 1, 2).clearContent();
}
function pend_clearMN_(sh) {
  const max = sh.getMaxRows();
  if (max >= 2) sh.getRange(2, PEND_NV_COL_PROD, max - 1, 2).clearContent();
}
function pend_countsToRows_(counts) {
  return Array.from(counts.entries())
    .sort((a, b) => String(a[0]).localeCompare(String(b[0]), "pt-BR"))
    .map(([name, qty]) => [String(name).trim(), qty]);
}
function pend_ensureRows_(sh, neededLastRow) {
  const max = sh.getMaxRows();
  if (neededLastRow > max) sh.insertRowsAfter(max, neededLastRow - max);
}
function pend_norm_(s) {
  return String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}
function pend_extractTerms_(statusRaw) {
  const s = String(statusRaw || "").trim();
  const up = s.toUpperCase();
  if (up.indexOf("FALTA") < 0) return [];

  const pos = up.indexOf("FALTA -");
  if (pos < 0) return [];

  const after = s.slice(pos + "FALTA -".length).trim();
  if (!after) return [];

  return after.split(";").map(x => String(x || "").trim()).filter(Boolean);
}
function pend_parseQtd_(v) {
  if (v === null || v === undefined || v === "") return 1;
  const s = String(v).trim().replace(",", ".");
  const n = Number(s);
  if (!isFinite(n) || n <= 0) return 1;
  return Math.floor(n);
}

/************** PEND: Falta/Não Verificado por ID direto (M/N) **************/
function pend_process_nv_porId_(sh, confTodosById, mapSubseq) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const idVals     = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID,     n, 1).getValues();
  const statusVals = sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS,  n, 1).getValues();

  pend_clearMN_(sh);

  const idsUnique = [];
  const seen = new Set();
  for (let i = 0; i < n; i++) {
    const id     = String(idVals[i][0]     ?? "").trim();
    const status = String(statusVals[i][0] ?? "").trim();
    if (!id) continue;
    if (status === "Confirmado") continue;
    if (!seen.has(id)) { seen.add(id); idsUnique.push(id); }
  }

  const counts = new Map();
  for (const id of idsUnique) {
    const linhas = confTodosById.get(id) || [];
    for (const it of linhas) {
      pend_processFaltaLine_(it, mapSubseq, counts);
    }
  }

  const rowsOut = pend_countsToRows_(counts);
  if (rowsOut.length) {
    pend_ensureRows_(sh, 1 + rowsOut.length);
    sh.getRange(2, PEND_NV_COL_PROD, rowsOut.length, 2).setValues(rowsOut);
  }
}

/******************* SHOPEE / MAGALU — IMPORT (gatilho 5 min) *******************/
function importarShopeeMagalu() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("importarShopeeMagalu: lock não obtido — outra execução em andamento.");
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shShopee   = ss.getSheetByName("Shopee");
    const shMagalu   = ss.getSheetByName("Magalu");
    const shEssencia = ss.getSheetByName("Essência do Brasil");
    if (!shShopee || !shMagalu || !shEssencia) {
      Logger.log("importarShopeeMagalu: abas 'Shopee', 'Magalu' ou 'Essência do Brasil' não encontradas.");
      return;
    }
    const shAmazon   = ss.getSheetByName("Amazon");
    const shFlexVapt = ss.getSheetByName("Flex/Vapt");
    if (!shAmazon || !shFlexVapt) {
      Logger.log("importarShopeeMagalu: abas 'Amazon' ou 'Flex/Vapt' não encontradas.");
      return;
    }

    limparAbaExcetoCabecalho_(shShopee);
    limparAbaExcetoCabecalho_(shMagalu);
    limparAbaExcetoCabecalho_(shEssencia);
    limparAbaExcetoCabecalho_(shAmazon);
    limparAbaExcetoCabecalho_(shFlexVapt);

    const ssSrc = SpreadsheetApp.openById(ML_IMPORT_CONFIG.STATUS_SOURCE_SPREADSHEET_ID);
    const rowsShopee   = [];
    const rowsMagalu   = [];
    const rowsEssencia = [];
    const rowsAmazon   = [];
    const rowsFlexVapt = [];

    for (const sheetName of ["Maio", "Abril"]) {
      const sh = ssSrc.getSheetByName(sheetName);
      if (!sh) { Logger.log("importarShopeeMagalu: aba '" + sheetName + "' não encontrada na fonte."); continue; }

      const lastRow = sh.getLastRow();
      if (lastRow < 2) continue;

      const n = lastRow - 1;
      const values = sh.getRange(2, 1, n, 7).getValues();
      const bgs    = sh.getRange(2, 7, n, 1).getBackgrounds(); // col G

      for (let i = 0; i < n; i++) {
        // Filtro 1 — Coluna C (índice 2): destino
        const colC = String(values[i][2] ?? "").toLowerCase();
        let dest = null;
        if (colC.includes("shopee"))      dest = "Shopee";
        else if (colC.includes("magalu")) dest = "Magalu";
        else if (colC.includes("essência do brasil") || colC.includes("essencia do brasil")) dest = "Essência do Brasil";
        else if (colC.includes("amazon")) dest = "Amazon";

        // Coluna G (índice 6): usada para exclusão e para Flex/Vapt
        const colG = String(values[i][6] ?? "").trim().toLowerCase();
        const bgG  = String(bgs[i][0]    ?? "").trim().toLowerCase();

        // Flex/Vapt: independente de destino, verifica coluna G
        const hasFlexVapt = colG.includes("flex") || colG.includes("vapt");
        const isExcluded  = colG.includes("ok") || bgG === ML_IMPORT_CONFIG.OK_GREEN.toLowerCase() || colG.includes("cancelado");
        if (hasFlexVapt && !isExcluded) {
          rowsFlexVapt.push([values[i][1], values[i][3], values[i][5], values[i][4]]);
        }

        if (!dest) continue;
        if (colG.includes("ok"))                   continue;
        if (bgG === ML_IMPORT_CONFIG.OK_GREEN)     continue;
        if (colG.includes("cancelado"))            continue;
        if (colG.includes("full"))                 continue;

        // B→ID, D→CLIENTE, F→PRODUTO, E→QTD
        const row = [values[i][1], values[i][3], values[i][5], values[i][4]];
        if (dest === "Shopee")               rowsShopee.push(row);
        else if (dest === "Magalu")          rowsMagalu.push(row);
        else if (dest === "Amazon")          rowsAmazon.push(row);
        else                                 rowsEssencia.push(row);
      }
    }

    if (rowsShopee.length)   shShopee.getRange(2, 1, rowsShopee.length, 4).setValues(rowsShopee);
    if (rowsMagalu.length)   shMagalu.getRange(2, 1, rowsMagalu.length, 4).setValues(rowsMagalu);
    if (rowsEssencia.length) shEssencia.getRange(2, 1, rowsEssencia.length, 4).setValues(rowsEssencia);
    if (rowsAmazon.length)   shAmazon.getRange(2, 1, rowsAmazon.length, 4).setValues(rowsAmazon);
    if (rowsFlexVapt.length) shFlexVapt.getRange(2, 1, rowsFlexVapt.length, 4).setValues(rowsFlexVapt);

    const noop = () => {};
    const statusById = construirStatusById_(noop);
    aplicarStatus_porAba_(statusById, "Shopee");
    aplicarStatus_porAba_(statusById, "Magalu");
    aplicarStatus_porAba_(statusById, "Essência do Brasil");
    aplicarStatus_porAba_(statusById, "Amazon");
    aplicarStatus_porAba_(statusById, "Flex/Vapt");

    const andamentoByIdQueue = construirFilaAndamentoPorId_(noop);
    aplicarAndamento_porAba_(andamentoByIdQueue, "Shopee");
    aplicarAndamento_porAba_(andamentoByIdQueue, "Magalu");
    aplicarAndamento_porAba_(andamentoByIdQueue, "Essência do Brasil");
    aplicarAndamento_porAba_(andamentoByIdQueue, "Amazon");
    aplicarAndamento_porAba_(andamentoByIdQueue, "Flex/Vapt");

    // Pendentes (colunas I e J)
    const confById  = pend_buildConferenciaById_();
    const mapSubseq = pend_buildDadosMapSubseq_();
    pend_process_porAba_(shShopee,   confById, mapSubseq);
    pend_process_porAba_(shMagalu,   confById, mapSubseq);
    pend_process_porAba_(shEssencia, confById, mapSubseq);
    pend_process_porAba_(shAmazon,   confById, mapSubseq);
    pend_process_porAba_(shFlexVapt, confById, mapSubseq);

    // Falta/Não Verificado (M/N)
    const confTodosById = pend_buildTodosExcTenho_();
    pend_process_nv_porId_(shShopee,   confTodosById, mapSubseq);
    pend_process_nv_porId_(shMagalu,   confTodosById, mapSubseq);
    pend_process_nv_porId_(shEssencia, confTodosById, mapSubseq);
    pend_process_nv_porId_(shAmazon,   confTodosById, mapSubseq);
    pend_process_nv_porId_(shFlexVapt, confTodosById, mapSubseq);

    Logger.log("importarShopeeMagalu: concluído. Shopee=" + rowsShopee.length + " | Magalu=" + rowsMagalu.length + " | Essência do Brasil=" + rowsEssencia.length + " | Amazon=" + rowsAmazon.length + " | Flex/Vapt=" + rowsFlexVapt.length);
  } finally {
    lock.releaseLock();
  }
}

function aplicarStatus_porAba_(statusById, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const n = lastRow - 1;
  const ids       = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID,        n, 1).getValues();
  const etiquetas = sh.getRange(2, ML_IMPORT_CONFIG.COL_ETIQUETAS,  n, 1).getValues();
  const out = ids.map((r, i) => {
    const id  = (r[0]              ?? "").toString().trim();
    const etq = (etiquetas[i][0]   ?? "").toString().trim();
    if (etq === "Cancelado") return ["Cancelado"];
    return [statusById.get(id) || "Pendente"];
  });
  sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS, n, 1).setValues(out);
}

function aplicarAndamento_porAba_(byIdQueue, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const n = lastRow - 1;
  const ids        = sh.getRange(2, ML_IMPORT_CONFIG.COL_ID,     n, 1).getValues();
  const statusVals = sh.getRange(2, ML_IMPORT_CONFIG.COL_STATUS,  n, 1).getValues();
  const out = ids.map((r, i) => {
    const status = (statusVals[i][0] ?? "").toString().trim();
    if (status === "Confirmado") return [""];
    const id   = (r[0] ?? "").toString().trim();
    const fila = byIdQueue.get(id);
    return [(fila && fila.length) ? fila.shift() : ""];
  });
  sh.getRange(2, ML_IMPORT_CONFIG.COL_ANDAMENTO, n, 1).setValues(out);
}
