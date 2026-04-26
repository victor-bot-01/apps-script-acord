/*******************************
 * CONFIG
 *******************************/
const CFG = {
  SPREADSHEET_ID: "1Tb5HyC5Axt9jxYbHotKTVx2NCQ__Fxy7WfIapKUL5hQ",
  SHEET_NAME: "Tarefas",
  TIMEZONE: "America/Sao_Paulo",
  SHEET_URL: "https://docs.google.com/spreadsheets/d/1Tb5HyC5Axt9jxYbHotKTVx2NCQ__Fxy7WfIapKUL5hQ/edit"
};

const STATUSES = ["PENDENTE", "ANDAMENTO", "BLOQUEADA", "CONCLUÍDA"];

/*******************************
 * LOG
 *******************************/
function log_(msg, obj) {
  if (obj !== undefined) {
    console.log("ERP BOT >>> " + msg + "\n" + JSON.stringify(obj, null, 2));
  } else {
    console.log("ERP BOT >>> " + msg);
  }
}

/*******************************
 * AUTH - SERVICE ACCOUNT
 *******************************/
function getServiceAccountToken_() {
  const ps = PropertiesService.getScriptProperties();

  const jsonRaw = ps.getProperty("SERVICE_ACCOUNT_JSON");
  const email = ps.getProperty("SERVICE_ACCOUNT_EMAIL");

  if (!jsonRaw) {
    throw new Error("Propriedade SERVICE_ACCOUNT_JSON não encontrada.");
  }

  if (!email) {
    throw new Error("Propriedade SERVICE_ACCOUNT_EMAIL não encontrada.");
  }

  const json = JSON.parse(jsonRaw);
  const now = Math.floor(Date.now() / 1000);

  const header = {
    alg: "RS256",
    typ: "JWT"
  };

  const claim = {
    iss: email,
    scope: "https://www.googleapis.com/auth/chat.bot",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now
  };

  const encodedHeader = Utilities.base64EncodeWebSafe(
    JSON.stringify(header),
    Utilities.Charset.UTF_8
  ).replace(/=+$/, "");

  const encodedClaim = Utilities.base64EncodeWebSafe(
    JSON.stringify(claim),
    Utilities.Charset.UTF_8
  ).replace(/=+$/, "");

  const unsignedJwt = encodedHeader + "." + encodedClaim;

  const signatureBytes = Utilities.computeRsaSha256Signature(
    unsignedJwt,
    json.private_key
  );

  const encodedSignature = Utilities.base64EncodeWebSafe(signatureBytes).replace(/=+$/, "");
  const jwt = unsignedJwt + "." + encodedSignature;

  const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("Falha ao obter token da service account: HTTP " + code + " - " + body);
  }

  const tokenData = JSON.parse(body);
  if (!tokenData.access_token) {
    throw new Error("Token access_token não retornado: " + body);
  }

  return tokenData.access_token;
}

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

/*******************************
 * VISUAL / COPY HELPERS
 *******************************/
function getStatusMeta_(status) {
  const s = String(status || "").trim().toUpperCase();

  if (s === "PENDENTE") {
    return {
      emoji: "🟡",
      label: "Pendente",
      chip: "Aguardando ação"
    };
  }

  if (s === "ANDAMENTO") {
    return {
      emoji: "🔵",
      label: "Em andamento",
      chip: "Em execução"
    };
  }

  if (s === "BLOQUEADA") {
    return {
      emoji: "🔴",
      label: "Bloqueada",
      chip: "Com impedimento"
    };
  }

  if (s === "CONCLUÍDA") {
    return {
      emoji: "🟢",
      label: "Concluída",
      chip: "Finalizada"
    };
  }

  return {
    emoji: "⚪",
    label: "Sem status",
    chip: "Sem definição"
  };
}

function formatPriorityLabel_(value) {
  const s = String(value || "").trim();
  if (!s) return "Não definida";

  const upper = s.toUpperCase();
  if (upper === "ALTA") return "Alta";
  if (upper === "MÉDIA" || upper === "MEDIA") return "Média";
  if (upper === "BAIXA") return "Baixa";
  if (upper === "URGENTE") return "Urgente";

  return s;
}

function buildBotHelpText_() {
  return (
    "✦ *ERP Tarefas*\n\n" +
    "Comandos disponíveis:\n" +
    "• `ajuda`\n" +
    "• `postar pendentes`\n" +
    "• `concluir 1`\n" +
    "• `status 1 andamento`\n" +
    "• `status 1 bloqueada`\n" +
    "• `status 1 concluída`\n\n" +
    "No grupo, use com menção:\n" +
    "`@ERP Tarefas postar pendentes`"
  );
}

function buildPostSummaryText_(result) {
  return (
    "✅ *Processamento concluído*\n\n" +
    "• Postadas: " + result.posted + "\n" +
    "• Ignoradas: " + result.skipped + "\n" +
    "• Erros: " + result.errors
  );
}

function buildStatusUpdatedText_(taskId, newStatus, who) {
  const meta = getStatusMeta_(newStatus);
  return (
    "✅ *Status atualizado com sucesso*\n\n" +
    "• Tarefa: #" + taskId + "\n" +
    "• Novo status: " + meta.emoji + " " + meta.label + "\n" +
    "• Atualizado por: " + who
  );
}

function buildDeleteSummaryText_(result) {
  return (
    "🗑️ *Processamento de exclusões concluído*\n\n" +
    "• Tarefas processadas: " + result.deleted + "\n" +
    "• Ignoradas: " + result.skipped + "\n" +
    "• Erros: " + result.errors
  );
}

/*******************************
 * EXTRAÇÃO DE DADOS DO EVENTO
 *******************************/
function getEventText_(e) {
  if (e && e.message && e.message.text) return String(e.message.text);

  if (
    e &&
    e.chat &&
    e.chat.messagePayload &&
    e.chat.messagePayload.message &&
    e.chat.messagePayload.message.text
  ) {
    return String(e.chat.messagePayload.message.text);
  }

  return "";
}

function getEventUserEmail_(e) {
  if (e && e.user && e.user.email) return String(e.user.email);

  if (e && e.chat && e.chat.user && e.chat.user.email) {
    return String(e.chat.user.email);
  }

  if (
    e &&
    e.chat &&
    e.chat.buttonClickedPayload &&
    e.chat.user &&
    e.chat.user.email
  ) {
    return String(e.chat.user.email);
  }

  if (
    e &&
    e.chat &&
    e.chat.messagePayload &&
    e.chat.messagePayload.message &&
    e.chat.messagePayload.message.sender &&
    e.chat.messagePayload.message.sender.email
  ) {
    return String(e.chat.messagePayload.message.sender.email);
  }

  return "";
}

function getEventUserDisplayName_(e) {
  if (e && e.user && e.user.displayName) return String(e.user.displayName);

  if (e && e.chat && e.chat.user && e.chat.user.displayName) {
    return String(e.chat.user.displayName);
  }

  if (
    e &&
    e.chat &&
    e.chat.buttonClickedPayload &&
    e.chat.user &&
    e.chat.user.displayName
  ) {
    return String(e.chat.user.displayName);
  }

  if (
    e &&
    e.chat &&
    e.chat.messagePayload &&
    e.chat.messagePayload.message &&
    e.chat.messagePayload.message.sender &&
    e.chat.messagePayload.message.sender.displayName
  ) {
    return String(e.chat.messagePayload.message.sender.displayName);
  }

  return "Usuário";
}

function getEventSpaceName_(e) {
  if (e && e.space && e.space.name) return String(e.space.name);

  if (e && e.chat && e.chat.space && e.chat.space.name) {
    return String(e.chat.space.name);
  }

  if (
    e &&
    e.chat &&
    e.chat.buttonClickedPayload &&
    e.chat.buttonClickedPayload.space &&
    e.chat.buttonClickedPayload.space.name
  ) {
    return String(e.chat.buttonClickedPayload.space.name);
  }

  if (
    e &&
    e.chat &&
    e.chat.messagePayload &&
    e.chat.messagePayload.space &&
    e.chat.messagePayload.space.name
  ) {
    return String(e.chat.messagePayload.space.name);
  }

  if (
    e &&
    e.chat &&
    e.chat.messagePayload &&
    e.chat.messagePayload.message &&
    e.chat.messagePayload.message.space &&
    e.chat.messagePayload.message.space.name
  ) {
    return String(e.chat.messagePayload.message.space.name);
  }

  if (e && e.addedToSpacePayload && e.addedToSpacePayload.space && e.addedToSpacePayload.space.name) {
    return String(e.addedToSpacePayload.space.name);
  }

  if (
    e &&
    e.chat &&
    e.chat.addedToSpacePayload &&
    e.chat.addedToSpacePayload.space &&
    e.chat.addedToSpacePayload.space.name
  ) {
    return String(e.chat.addedToSpacePayload.space.name);
  }

  return "";
}

/*******************************
 * NORMALIZAÇÃO DE TEXTO
 *******************************/
function normalizeCommandText_(raw) {
  let text = String(raw || "").trim();

  text = text.replace(/^<users\/[^>]+>\s*/i, "");

  // nomes antigos
  text = text.replace(/^@erp\s*tarefas[\s,:-]*/i, "");
  text = text.replace(/^erp\s*tarefas[\s,:-]*/i, "");

  // nome novo
  text = text.replace(/^@novo\s*funcion[áa]rio[\s,:-]*/i, "");
  text = text.replace(/^novo\s*funcion[áa]rio[\s,:-]*/i, "");

  return text.trim().toLowerCase();
}

/*******************************
 * COMPATIBILIDADE COM GATILHOS
 *******************************/
function onAppCommand(e) {
  log_("onAppCommand recebido", {
    space: getEventSpaceName_(e),
    user: getEventUserEmail_(e),
    text: getEventText_(e)
  });
  return { text: "Use `ajuda` para ver os comandos disponíveis." };
}

function onAddedToSpace(e) {
  return onAddToSpace(e);
}

function onRemovedFromSpace(e) {
  return onRemoveFromSpace(e);
}

/*******************************
 * GOOGLE CHAT APP ENTRYPOINTS
 *******************************/
function onAddToSpace(e) {
  const spaceName = getEventSpaceName_(e);

  log_("onAddToSpace recebido", e);
  log_("onAddToSpace spaceName=" + spaceName);

  return {
    text:
      "✅ *ERP Tarefas instalado com sucesso*\n\n" +
      "Comandos principais:\n" +
      "• `ajuda`\n" +
      "• `postar pendentes`\n" +
      "• `concluir 1`\n" +
      "• `status 1 andamento`\n" +
      "• `status 1 bloqueada`\n" +
      "• `status 1 concluída`\n\n" +
      "Use este valor na coluna `SpaceId` da planilha:\n" +
      (spaceName || "(space.name não encontrado neste evento)")
  };
}

function onRemoveFromSpace(e) {
  log_("onRemoveFromSpace recebido", e);
}

function onMessage(e) {
  const ini = new Date();

  try {
    const raw = getEventText_(e);
    const text = normalizeCommandText_(raw);

    log_("onMessage recebido", {
      space: getEventSpaceName_(e),
      user: getEventUserEmail_(e),
      text: raw
    });

    log_("Mensagem parseada: [" + raw + "] -> [" + text + "]");

    if (text === "ajuda" || text === "help" || text === "/ajuda") {
      return { text: buildBotHelpText_() };
    }

    if (text === "postar pendentes" || text === "/postar_pendentes") {
      const result = dispatchPendingTasksToSpace_();
      log_("Resultado postar pendentes", result);
      return { text: buildPostSummaryText_(result) };
    }

    let m = text.match(/^concluir\s+(\S+)$/i);
    if (m) {
      return handleSetStatusByCommand_(e, m[1], "CONCLUÍDA");
    }

    m = text.match(/^status\s+(\S+)\s+(.+)$/i);
    if (m) {
      const taskId = m[1];
      const newStatus = normalizeStatus_(m[2]);

      if (STATUSES.indexOf(newStatus) === -1) {
        return {
          text:
            "⚠️ *Status inválido*\n\n" +
            "Use um destes valores:\n" +
            "• `PENDENTE`\n" +
            "• `ANDAMENTO`\n" +
            "• `BLOQUEADA`\n" +
            "• `CONCLUÍDA`"
        };
      }

      return handleSetStatusByCommand_(e, taskId, newStatus);
    }

    return {
      text:
        "Não reconheci esse comando.\n\n" +
        "Digite `ajuda` para ver as opções disponíveis."
    };

  } catch (err) {
    log_("Erro no onMessage", {
      error: String(err),
      stack: err && err.stack ? err.stack : ""
    });
    return { text: "❌ Erro no bot: " + err };
  } finally {
    log_("onMessage fim (" + (new Date() - ini) + "ms)");
  }
}

function onCardClick(e) {
  try {
    log_("onCardClick recebido", e);

    const action =
      (e && e.action && e.action.actionMethodName) ||
      (e && e.commonEventObject && e.commonEventObject.invokedFunction) ||
      "";

    const params = keyValueParams_(e);
    const taskId = params.taskId;

    if (!taskId) {
      return { text: "Erro: taskId ausente." };
    }

    if (action === "markPending") {
      return handleSetStatusByCommand_(e, taskId, "PENDENTE");
    }

    if (action === "markInProgress") {
      return handleSetStatusByCommand_(e, taskId, "ANDAMENTO");
    }

    if (action === "markBlocked") {
      return handleSetStatusByCommand_(e, taskId, "BLOQUEADA");
    }

    if (action === "markDone") {
      return handleSetStatusByCommand_(e, taskId, "CONCLUÍDA");
    }

    return { text: "Ação não reconhecida." };
  } catch (err) {
    log_("Erro no onCardClick", {
      error: String(err),
      stack: err && err.stack ? err.stack : ""
    });
    return { text: "Erro no clique: " + err };
  }
}

/*******************************
 * FUNÇÕES GLOBAIS DOS BOTÕES
 *******************************/
function markPending(e) {
  try {
    log_("markPending recebido", e);
    const params = keyValueParams_(e);
    const taskId = params.taskId;

    if (!taskId) return { text: "Erro: taskId ausente." };

    return handleSetStatusByCommand_(e, taskId, "PENDENTE");
  } catch (err) {
    log_("Erro no markPending", {
      error: String(err),
      stack: err && err.stack ? err.stack : ""
    });
    return { text: "Erro ao definir como pendente: " + err };
  }
}

function markInProgress(e) {
  try {
    log_("markInProgress recebido", e);
    const params = keyValueParams_(e);
    const taskId = params.taskId;

    if (!taskId) return { text: "Erro: taskId ausente." };

    return handleSetStatusByCommand_(e, taskId, "ANDAMENTO");
  } catch (err) {
    log_("Erro no markInProgress", {
      error: String(err),
      stack: err && err.stack ? err.stack : ""
    });
    return { text: "Erro ao definir como andamento: " + err };
  }
}

function markBlocked(e) {
  try {
    log_("markBlocked recebido", e);
    const params = keyValueParams_(e);
    const taskId = params.taskId;

    if (!taskId) return { text: "Erro: taskId ausente." };

    return handleSetStatusByCommand_(e, taskId, "BLOQUEADA");
  } catch (err) {
    log_("Erro no markBlocked", {
      error: String(err),
      stack: err && err.stack ? err.stack : ""
    });
    return { text: "Erro ao definir como bloqueada: " + err };
  }
}

function markDone(e) {
  try {
    log_("markDone recebido", e);
    const params = keyValueParams_(e);
    const taskId = params.taskId;

    if (!taskId) return { text: "Erro: taskId ausente." };

    return handleSetStatusByCommand_(e, taskId, "CONCLUÍDA");
  } catch (err) {
    log_("Erro no markDone", {
      error: String(err),
      stack: err && err.stack ? err.stack : ""
    });
    return { text: "Erro ao concluir: " + err };
  }
}

/*******************************
 * TRIGGER MANUAL / TEMPO
 *******************************/
function dispatchPendingTasksToSpace() {
  return dispatchPendingTasksToSpace_();
}

function autoPostScheduledTasks() {
  const deleteResult = processDeleteRequests_();
  const postResult = dispatchScheduledTasks_();

  return {
    deleted: deleteResult.deleted,
    deleteErrors: deleteResult.errors,
    posted: postResult.posted,
    skipped: postResult.skipped,
    postErrors: postResult.errors
  };
}

function processDeleteRequests() {
  return processDeleteRequests_();
}

/*******************************
 * PROCESSAR EXCLUSÕES
 *******************************/
function processDeleteRequests_() {
  log_("processDeleteRequests_ start");

  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CFG.SHEET_NAME);
    if (!sh) throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { deleted: 0, skipped: 0, errors: 0 };

    const col = headerMap_(data[0]);

    if (col.ApagarMensagem === undefined) {
      log_("Coluna ApagarMensagem não existe. Nada para processar.");
      return { deleted: 0, skipped: data.length - 1, errors: 0 };
    }

    let deleted = 0;
    let skipped = 0;
    let errors = 0;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const id = String(row[col.ID] || "").trim();
      const apagarFlag = String(row[col.ApagarMensagem] || "").trim().toUpperCase();

      if (apagarFlag !== "APAGAR") {
        skipped++;
        continue;
      }

      const chatMessageRaw = (col.ChatMessageName !== undefined)
        ? String(row[col.ChatMessageName] || "").trim()
        : "";

      const messageNames = splitMessageNames_(chatMessageRaw);

      if (messageNames.length === 0) {
        if (col.ChatMessageName !== undefined) {
          sh.getRange(r + 1, col.ChatMessageName + 1).clearContent();
        }
        sh.getRange(r + 1, col.ApagarMensagem + 1).setValue("Apagado");
        deleted++;
        log_("Linha ID=" + id + " marcada como Apagado sem mensagens para excluir");
        continue;
      }

      const remaining = [];

      for (let i = 0; i < messageNames.length; i++) {
        const messageName = messageNames[i];
        try {
          deleteChatMessage_(messageName);
          log_("✅ Mensagem apagada: " + messageName);
        } catch (err) {
          remaining.push(messageName);
          errors++;
          log_("❌ Falha ao apagar mensagem " + messageName, {
            error: String(err),
            stack: err && err.stack ? err.stack : ""
          });
        }
      }

      if (remaining.length === 0) {
        if (col.ChatMessageName !== undefined) {
          sh.getRange(r + 1, col.ChatMessageName + 1).clearContent();
        }
        sh.getRange(r + 1, col.ApagarMensagem + 1).setValue("Apagado");
        deleted++;
        log_("Linha ID=" + id + " apagada com sucesso");
      } else {
        if (col.ChatMessageName !== undefined) {
          sh.getRange(r + 1, col.ChatMessageName + 1).setValue(remaining.join("\n"));
        }
        log_("Linha ID=" + id + " ainda tem mensagens não apagadas", { remaining: remaining });
      }
    }

    log_("processDeleteRequests_ done deleted=" + deleted + " skipped=" + skipped + " errors=" + errors);
    return { deleted, skipped, errors };

  } finally {
    lock.releaseLock();
  }
}

function deleteChatMessage_(messageName) {
  return callChatApi_("delete", messageName);
}

/*******************************
 * CORE - ENVIO MANUAL
 *******************************/
function dispatchPendingTasksToSpace_() {
  log_("dispatchPendingTasksToSpace_ start");

  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CFG.SHEET_NAME);
    if (!sh) throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');

    const data = sh.getDataRange().getValues();
    log_("Linhas lidas: " + data.length);

    if (data.length < 2) return { posted: 0, skipped: 0, errors: 0 };

    const col = headerMap_(data[0]);
    let posted = 0;
    let skipped = 0;
    let errors = 0;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];

      const id = String(row[col.ID] || "").trim();
      const status = String(row[col.Status] || "").trim().toUpperCase();
      let spaceId = String(row[col.SpaceId] || "").trim();

      if (!id || !spaceId) {
        skipped++;
        continue;
      }

      spaceId = normalizeSpaceId_(spaceId);

      if (status === "PENDENTE") {
        try {
          const task = rowToTask_(row, col);
          const created = postTaskCardToSpace_(spaceId, task);

          appendMessageName_(sh, r + 1, col.ChatMessageName + 1, created.name || "");

          posted++;
          log_("✅ Postado ID=" + id + " messageName=" + (created.name || ""));
        } catch (err) {
          errors++;
          log_("❌ Falha ao postar ID=" + id, {
            error: String(err),
            stack: err && err.stack ? err.stack : ""
          });
        }
      } else {
        skipped++;
      }
    }

    return { posted, skipped, errors };

  } finally {
    lock.releaseLock();
  }
}

/*******************************
 * CORE - ENVIO AUTOMÁTICO POR HORA
 *******************************/
function dispatchScheduledTasks_() {
  log_("dispatchScheduledTasks_ start");

  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CFG.SHEET_NAME);
    if (!sh) throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { posted: 0, skipped: 0, errors: 0 };

    const col = headerMap_(data[0]);
    const now = new Date();

    let posted = 0;
    let skipped = 0;
    let errors = 0;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];

      const id = String(row[col.ID] || "").trim();
      let spaceId = String(row[col.SpaceId] || "").trim();

      if (!id || !spaceId) {
        skipped++;
        continue;
      }

      const readyInfo = getTaskReadyInfo_(row, col, now);
      if (!readyInfo.ready) {
        skipped++;
        continue;
      }

      spaceId = normalizeSpaceId_(spaceId);

      try {
        const task = rowToTask_(row, col);
        const created = postTaskCardToSpace_(spaceId, task);

        appendMessageName_(sh, r + 1, col.ChatMessageName + 1, created.name || "");

        const ultimoDisparoCell = sh.getRange(r + 1, col.UltimoDisparo + 1);
        ultimoDisparoCell.setNumberFormat("@");
        ultimoDisparoCell.setValue(readyInfo.disparoKey);

        posted++;
        log_("✅ Postado agendado ID=" + id + " messageName=" + (created.name || "") + " disparoKey=" + readyInfo.disparoKey);
      } catch (err) {
        errors++;
        log_("❌ Falha ao postar tarefa agendada ID=" + id, {
          error: String(err),
          stack: err && err.stack ? err.stack : ""
        });
      }
    }

    return { posted, skipped, errors };

  } finally {
    lock.releaseLock();
  }
}

/*******************************
 * ENVIO / ATUALIZAÇÃO DE CARDS
 *******************************/
function buildTaskText_(task) {
  const meta = getStatusMeta_(task.status);
  const prioridade = formatPriorityLabel_(task.prioridade);

  let texto =
    meta.emoji + " *Tarefa #" + task.id + "*\n" +
    meta.label.toUpperCase() + " • Prioridade " + prioridade + "\n\n" +
    "*Responsável*\n" +
    (task.responsavel || "-") + "\n\n" +
    "*Atividade*\n" +
    (task.tarefa || "-") + "\n\n" +
    "*Prazo*\n" +
    (task.prazo || "-") + "\n\n" +
    "*Hora programada*\n" +
    (task.hora || "-");

  if (task.link) {
    texto += "\n\n*Material relacionado*\n" + task.link;
  }

  if (task.audit) {
    texto += "\n\n*Última atualização*\n" + task.audit;
  }

  return texto.trim();
}

function postTaskCardToSpace_(spaceName, task) {
  const card = buildTaskCard_(task);
  return callChatApi_("post", spaceName + "/messages", card);
}

function updateTaskCardMessage_(messageName, task) {
  const card = buildTaskCard_(task);
  return callChatApi_("patch", messageName + "?updateMask=cardsV2", card);
}

/*******************************
 * ATUALIZAÇÃO DE STATUS
 *******************************/
function handleSetStatusByCommand_(e, taskId, newStatus) {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw new Error('Aba "' + CFG.SHEET_NAME + '" não encontrada.');

  const data = sh.getDataRange().getValues();
  const col = headerMap_(data[0]);

  const rIndex = findRowIndexById_(data, col, taskId);
  if (rIndex < 0) return { text: "Tarefa " + taskId + " não encontrada." };

  const sheetRow = rIndex + 1;
  const who = getEventUserDisplayName_(e);
  const now = Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");

  sh.getRange(sheetRow, col.Status + 1).setValue(newStatus);
  sh.getRange(sheetRow, col.AtualizadoPor + 1).setValue(who);
  sh.getRange(sheetRow, col.AtualizadoEm + 1).setValue(now);

  const updatedRow = sh.getRange(sheetRow, 1, 1, data[0].length).getValues()[0];
  const task = rowToTask_(updatedRow, col);

  const messageNamesRaw = String(updatedRow[col.ChatMessageName] || "").trim();
  const messageNames = splitMessageNames_(messageNamesRaw);

  if (messageNames.length > 0) {
    let updatedCount = 0;
    for (let i = 0; i < messageNames.length; i++) {
      const messageName = messageNames[i];
      try {
        updateTaskCardMessage_(messageName, task);
        updatedCount++;
      } catch (err) {
        log_("Falha ao atualizar mensagem " + messageName, {
          error: String(err),
          stack: err && err.stack ? err.stack : ""
        });
      }
    }
    log_("Mensagens atualizadas para tarefa ID=" + taskId + ": " + updatedCount);
  } else {
    log_("Sem ChatMessageName para atualizar tarefa ID=" + taskId);
  }

  return {
    actionResponse: {
      type: "UPDATE_MESSAGE"
    },
    cardsV2: buildTaskCard_(task).cardsV2
  };
}

/*******************************
 * CARD PRINCIPAL - CLEAN PRO
 *******************************/
function buildTaskCard_(task) {
  const meta = getStatusMeta_(task.status);
  const prioridade = formatPriorityLabel_(task.prioridade);

  const widgetsStatus = [
    {
      decoratedText: {
        topLabel: "Estado atual",
        text: meta.emoji + " <b>" + meta.label + "</b> • " + meta.chip,
        wrapText: true
      }
    },
    {
      decoratedText: {
        topLabel: "Prioridade",
        text: prioridade
      }
    }
  ];

  const widgetsDetalhes = [
    {
      decoratedText: {
        startIcon: { knownIcon: "PERSON" },
        topLabel: "Responsável",
        text: task.responsavel || "-"
      }
    },
    {
      decoratedText: {
        startIcon: { knownIcon: "DESCRIPTION" },
        topLabel: "Atividade",
        text: task.tarefa || "-",
        wrapText: true
      }
    },
    {
      decoratedText: {
        startIcon: { knownIcon: "INVITE" },
        topLabel: "Prazo",
        text: task.prazo || "-"
      }
    },
    {
      decoratedText: {
        startIcon: { knownIcon: "CLOCK" },
        topLabel: "Hora programada",
        text: task.hora || "-"
      }
    }
  ];

  if (task.link) {
    widgetsDetalhes.push({
      decoratedText: {
        startIcon: { knownIcon: "DESCRIPTION" },
        topLabel: "Material relacionado",
        text: task.link,
        wrapText: true
      }
    });
  }

  if (task.audit) {
    widgetsDetalhes.push({
      decoratedText: {
        startIcon: { knownIcon: "CLOCK" },
        topLabel: "Última atualização",
        text: task.audit,
        wrapText: true
      }
    });
  }

  const botoesLinha1 = [
    {
      text: "Pendente",
      onClick: {
        action: {
          function: "markPending",
          parameters: [{ key: "taskId", value: String(task.id) }]
        }
      }
    },
    {
      text: "Andamento",
      onClick: {
        action: {
          function: "markInProgress",
          parameters: [{ key: "taskId", value: String(task.id) }]
        }
      }
    }
  ];

  const botoesLinha2 = [
    {
      text: "Bloqueada",
      onClick: {
        action: {
          function: "markBlocked",
          parameters: [{ key: "taskId", value: String(task.id) }]
        }
      }
    },
    {
      text: "Concluir",
      onClick: {
        action: {
          function: "markDone",
          parameters: [{ key: "taskId", value: String(task.id) }]
        }
      }
    }
  ];

  if (task.link) {
    botoesLinha2.push({
      text: "Abrir material",
      onClick: {
        openLink: { url: task.link }
      }
    });
  }

  return {
    cardsV2: [{
      cardId: "task_" + task.id,
      card: {
        header: {
          title: "Miau! O sol nasce todo dia #" + task.id,
          subtitle: meta.emoji + " " + meta.label + " • Prioridade " + prioridade
        },
        sections: [
          {
            header: "Status e prioridade",
            widgets: widgetsStatus
          },
          {
            header: "Detalhes",
            widgets: widgetsDetalhes
          },
          {
            header: "Ações",
            widgets: [
              {
                buttonList: {
                  buttons: botoesLinha1
                }
              },
              {
                buttonList: {
                  buttons: botoesLinha2
                }
              }
            ]
          }
        ]
      }
    }]
  };
}

/*******************************
 * HELPERS DE AGENDAMENTO
 *******************************/
function parseHoraToParts_(horaValue) {
  if (horaValue instanceof Date) {
    return {
      hour: horaValue.getHours(),
      minute: horaValue.getMinutes()
    };
  }

  const s = String(horaValue || "").trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;

  const hour = Number(m[1]);
  const minute = Number(m[2]);

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) return null;

  return { hour: hour, minute: minute };
}

function buildTodayHoraDate_(now, horaValue) {
  const hora = parseHoraToParts_(horaValue);
  if (!hora) return null;

  return new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    hora.hour,
    hora.minute,
    0,
    0
  );
}

function buildDisparoKey_(dateObj) {
  return Utilities.formatDate(dateObj, CFG.TIMEZONE, "yyyy-MM-dd HH:mm");
}

function normalizeUltimoDisparoValue_(value) {
  if (!value) return "";

  if (value instanceof Date) {
    return Utilities.formatDate(value, CFG.TIMEZONE, "yyyy-MM-dd HH:mm");
  }

  const s = String(value).trim();
  if (!s) return "";

  return s;
}

function getTaskReadyInfo_(row, col, now) {
  const status = String(row[col.Status] || "").trim().toUpperCase();
  if (status !== "PENDENTE") return { ready: false };

  const horaValue = row[col.Hora];
  const ultimoDisparo = normalizeUltimoDisparoValue_(row[col.UltimoDisparo]);

  const scheduledAt = buildTodayHoraDate_(now, horaValue);
  if (!scheduledAt) return { ready: false };

  if (scheduledAt.getTime() > now.getTime()) {
    return { ready: false };
  }

  const disparoKey = buildDisparoKey_(scheduledAt);

  log_("Comparando disparo", {
    id: String(row[col.ID] || "").trim(),
    ultimoDisparoLido: ultimoDisparo,
    disparoKeyAtual: disparoKey
  });

  if (ultimoDisparo === disparoKey) {
    return { ready: false, alreadySent: true, disparoKey: disparoKey };
  }

  return {
    ready: true,
    alreadySent: false,
    scheduledAt: scheduledAt,
    disparoKey: disparoKey
  };
}

/*******************************
 * HELPERS GERAIS
 *******************************/
function normalizeSpaceId_(spaceId) {
  const s = String(spaceId || "").trim();
  if (!s) return s;
  if (s.startsWith("spaces/")) return s;
  return "spaces/" + s;
}

function normalizeStatus_(value) {
  const s = String(value || "").trim().toUpperCase();

  if (s === "PENDENTE") return "PENDENTE";
  if (s === "ANDAMENTO") return "ANDAMENTO";
  if (s === "BLOQUEADA") return "BLOQUEADA";
  if (s === "CONCLUÍDA") return "CONCLUÍDA";
  if (s === "CONCLUIDA") return "CONCLUÍDA";

  return s;
}

function appendMessageName_(sheet, rowIndex, colIndex, newMessageName) {
  if (!newMessageName) return;

  const cell = sheet.getRange(rowIndex, colIndex);
  const current = String(cell.getValue() || "").trim();

  if (!current) {
    cell.setValue(newMessageName);
    return;
  }

  const lines = splitMessageNames_(current);

  if (lines.indexOf(newMessageName) === -1) {
    cell.setValue(current + "\n" + newMessageName);
  }
}

function splitMessageNames_(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map(function(x) { return String(x).trim(); })
    .filter(function(x) { return !!x; });
}

function headerMap_(headerRow) {
  const map = {};
  headerRow.forEach(function(h, i) {
    const key = String(h || "").trim();
    if (key) map[key] = i;
  });

  const required = [
    "ID",
    "Responsável",
    "Tarefa",
    "Prazo",
    "Hora",
    "Prioridade",
    "Status",
    "SpaceId",
    "ChatMessageName",
    "UltimoDisparo",
    "AtualizadoPor",
    "AtualizadoEm"
  ];

  required.forEach(function(k) {
    if (map[k] === undefined) throw new Error("Coluna obrigatória ausente: " + k);
  });

  return map;
}

function findRowIndexById_(data, col, taskId) {
  const wanted = String(taskId).trim();
  for (let r = 1; r < data.length; r++) {
    const id = String(data[r][col.ID] || "").trim();
    if (id === wanted) return r;
  }
  return -1;
}

function rowToTask_(row, col) {
  const id = String(row[col.ID] || "").trim();
  const responsavel = String(row[col["Responsável"]] || "").trim();
  const tarefa = String(row[col.Tarefa] || "").trim();
  const prazo = formatPrazo_(row[col.Prazo]);
  const hora = formatHora_(row[col.Hora]);
  const prioridade = String(row[col.Prioridade] || "").trim();
  const status = String(row[col.Status] || "").trim().toUpperCase();
  const link = (col.Link !== undefined) ? String(row[col.Link] || "").trim() : "";

  const updBy = String(row[col.AtualizadoPor] || "").trim();
  const updAt = String(row[col.AtualizadoEm] || "").trim();
  const audit = (updBy || updAt) ? (updBy || "-") + " • " + (updAt || "-") : "";

  return { id, responsavel, tarefa, prazo, hora, prioridade, status, link, audit };
}

function formatPrazo_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, CFG.TIMEZONE, "yyyy-MM-dd");
  }
  return String(v || "").trim();
}

function formatHora_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, CFG.TIMEZONE, "HH:mm");
  }
  return String(v || "").trim();
}

function keyValueParams_(e) {
  const out = {};

  function addParams_(params) {
    if (!params) return;

    if (Array.isArray(params)) {
      params.forEach(function(p) {
        if (!p) return;
        if (p.key !== undefined) {
          out[p.key] = p.value;
        }
      });
      return;
    }

    if (typeof params === "object") {
      Object.keys(params).forEach(function(key) {
        const val = params[key];

        if (val && typeof val === "object" && "value" in val) {
          out[key] = val.value;
        } else {
          out[key] = val;
        }
      });
    }
  }

  addParams_(e && e.action && e.action.parameters);
  addParams_(e && e.commonEventObject && e.commonEventObject.parameters);
  addParams_(e && e.common && e.common.parameters);

  if (
    e &&
    e.common &&
    e.common.formInputs &&
    e.common.formInputs.newStatus &&
    e.common.formInputs.newStatus.stringInputs &&
    e.common.formInputs.newStatus.stringInputs.value &&
    e.common.formInputs.newStatus.stringInputs.value.length > 0
  ) {
    out.newStatus = e.common.formInputs.newStatus.stringInputs.value[0];
  }

  if (
    e &&
    e.commonEventObject &&
    e.commonEventObject.formInputs &&
    e.commonEventObject.formInputs.newStatus &&
    e.commonEventObject.formInputs.newStatus.stringInputs &&
    e.commonEventObject.formInputs.newStatus.stringInputs.value &&
    e.commonEventObject.formInputs.newStatus.stringInputs.value.length > 0
  ) {
    out.newStatus = e.commonEventObject.formInputs.newStatus.stringInputs.value[0];
  }

  return out;
}

/*******************************
 * TESTE MANUAL OPCIONAL
 *******************************/
function testPostNoSpace() {
  const task = {
    id: "TESTE",
    responsavel: "Victor",
    tarefa: "Teste manual premium",
    prazo: Utilities.formatDate(new Date(), CFG.TIMEZONE, "yyyy-MM-dd"),
    hora: Utilities.formatDate(new Date(), CFG.TIMEZONE, "HH:mm"),
    prioridade: "Alta",
    status: "PENDENTE",
    link: CFG.SHEET_URL,
    audit: ""
  };

  return postTaskCardToSpace_("spaces/AAQARNpKPzE", task);
}

/*******************************
 * doPost só para referência
 *******************************/
function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const event = JSON.parse(body);

    let response;

    if (event.type === "MESSAGE") {
      response = onMessage(event);
    } else if (event.type === "CARD_CLICKED") {
      response = onCardClick(event);
    } else if (event.type === "ADDED_TO_SPACE") {
      response = onAddedToSpace(event);
    } else if (event.type === "REMOVED_FROM_SPACE") {
      response = onRemoveFromSpace(event);
    } else {
      response = { text: "Evento não tratado: " + (event.type || "unknown") };
    }

    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ text: "Erro no bot: " + err }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
