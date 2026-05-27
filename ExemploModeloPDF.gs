function exportarEtiquetas58mmParaPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Etiquetas");
  if (!sheet) {
    Logger.log("❌ Aba 'Etiquetas' não encontrada.");
    return;
  }

  // --- Remover linhas com coluna A vazia ---
  var ultimaLinha = sheet.getLastRow();
  for (var i = ultimaLinha; i >= 1; i--) {
    var valor = sheet.getRange(i, 1).getValue();
    if (!valor || valor.toString().trim() === "") {
      sheet.deleteRow(i);
    }
  }

  // --- Atualiza a última linha após deletar ---
  ultimaLinha = sheet.getLastRow();
  if (ultimaLinha === 0) {
    Logger.log("⚠️ Nenhuma célula preenchida.");
    return;
  }

  // --- Aplica bordas em todas as células usadas na coluna A ---
  var rangeUsado = sheet.getRange(1, 1, ultimaLinha, 1);
  rangeUsado.setBorder(true, true, true, true, true, true);

  // --- Verifica ou cria pasta ---
  var pastaNome = "Etiquetas Vendas PDF";
  var pastas = DriveApp.getFoldersByName(pastaNome);
  var pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(pastaNome);

  // --- Monta a URL para exportação ---
  var url_base = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?";
  var parametros = {
    format: "pdf",
    size: "letter",
    portrait: true,
    fitw: false,
    scale: 1,
    top_margin: 0,
    bottom_margin: 0,
    left_margin: 0,
    right_margin: 0,
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: false,
    fzr: false,
    gid: sheet.getSheetId(),
    range: "A1:A" + ultimaLinha
  };

  var query = [];
  for (var p in parametros) query.push(p + "=" + parametros[p]);
  var url_final = url_base + query.join("&");

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url_final, { headers: { Authorization: "Bearer " + token } });

  // --- Salva PDF ---
  var nomeArquivo = "Etiqueta58mm_" + new Date().toISOString().replace(/[:.]/g, "-") + ".pdf";
  var arquivo = pasta.createFile(response.getBlob().setName(nomeArquivo));

  Logger.log("✅ PDF com etiquetas em tamanho real: " + arquivo.getUrl());

  // === NOVO CÓDIGO (limpar valores, manter bordas) ===
  if (ultimaLinha > 0) {
    var rangeLimpar = sheet.getRange(1, 1, ultimaLinha);
    rangeLimpar.clearContent();  // Apaga textos
    Logger.log("🧹 Valores da coluna A apagados (bordas mantidas).");
  } else {
    Logger.log("⚠️ Nenhum dado para limpar na coluna A.");
  }
}
