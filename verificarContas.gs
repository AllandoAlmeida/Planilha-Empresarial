/**
 * Abre o pop-up para verificação de contas.
 * O arquivo HTML deve ser nomeado exatamente "popupVerificarContas.html".
 */
function abrirPopupVerificarContas() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('popupVerificarContas')
      .setWidth(1000)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Verificar Contas");
  } catch (error) {
    console.error("Erro ao abrir pop-up:", error);
  }
}

/**
 * Busca os lançamentos na aba "Banco de Dados Lançamentos" com base no tipo (opcional), mês e ano.
 * Os dados são lidos a partir da linha 1 (sem cabeçalho) e a data (coluna A) é utilizada para extrair o mês e o ano.
 * Retorna um HTML com uma tabela contendo: Data, Descrição, Tipo de Pagamento, Categoria, Tipo, Valor e Status.
 * Se o status for "Em aberto", exibe um botão "Pagar" para atualizar o status para "Pago".
 *
 * @param {string} tipo Tipo de lançamento para filtrar ("Receita" ou "Despesa") – opcional.
 * @param {number|string} mes O mês para filtrar (ex.: 1 para Janeiro).
 * @param {number|string} ano O ano para filtrar (ex.: 2025).
 * @return {string} HTML com a tabela dos lançamentos ou mensagem informativa.
 */
function buscarContas(tipo, mes, ano) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
    if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return "Nenhum lançamento encontrado.";
    
    // Lê todos os dados a partir da linha 1 (sem cabeçalho)
    var range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    var values = range.getValues();
    
    var results = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var data = new Date(row[0]);
      if (isNaN(data.getTime())) continue; // ignora linhas sem data válida
      
      var rowMonth = data.getMonth() + 1;
      var rowYear = data.getFullYear();
      var desc = row[1] ? row[1].toString().trim() : "";
      var tipoPg = row[2] ? row[2].toString().trim() : "";  // Tipo de Pagamento (coluna C)
      var categoria = row[3] ? row[3].toString().trim() : ""; // Categoria (coluna D)
      var tipoLanc = row[5] ? row[5].toString().trim() : "";  // Tipo (coluna F)
      var valor = row[6] ? parseFloat(row[6]) : 0;           // Valor (coluna G)
      var status = row[9] ? row[9].toString().trim() : "";     // Status (coluna J)
      
      // Se for informado um tipo, filtra
      if (tipo && tipo.trim() !== "" && tipoLanc.toLowerCase() !== tipo.toLowerCase()) {
        continue;
      }
      if (rowMonth === Number(mes) && rowYear === Number(ano)) {
        results.push({
          rowIndex: i + 1, // Como os dados começam na linha 1
          data: data,
          descricao: desc,
          tipoPagamento: tipoPg,
          categoria: categoria,
          tipoLancamento: tipoLanc,
          valor: valor,
          status: status
        });
      }
    }
    
    if (results.length === 0) return "Nenhum lançamento encontrado para o período selecionado.";
    
    var html = "<table style='width:100%; border-collapse: collapse;'>";
    html += "<thead><tr style='background-color:#007bff; color:#fff;'>";
    html += "<th>Data</th><th>Descrição</th><th>Tipo de Pagamento</th><th>Categoria</th><th>Tipo</th><th>Valor (R$)</th><th>Status</th><th>Ação</th>";
    html += "</tr></thead><tbody>";
    
    results.forEach(function(item) {
      var dataStr = Utilities.formatDate(item.data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      html += "<tr style='border: 1px solid #ddd;'>";
      html += "<td>" + dataStr + "</td>";
      html += "<td>" + item.descricao + "</td>";
      html += "<td>" + item.tipoPagamento + "</td>";
      html += "<td>" + item.categoria + "</td>";
      html += "<td>" + item.tipoLancamento + "</td>";
      html += "<td>" + item.valor.toFixed(2) + "</td>";
      html += "<td>" + item.status + "</td>";
      html += "<td>";
      if (item.status.toLowerCase() === "em aberto") {
        html += "<button onclick=\"google.script.run.withSuccessHandler(function() { buscarContas('" + tipo + "', " + mes + ", " + ano + "); }).pagarRegistro(" + item.rowIndex + ", " + mes + ", " + ano + ")\">Pagar</button>";
      } else {
        html += "-";
      }
      html += "</td>";
      html += "</tr>";
    });
    
    html += "</tbody></table>";
    return html;
  } catch (error) {
    Logger.log("Erro em buscarContas: " + error.message);
    return "Erro ao buscar contas. Verifique os dados e tente novamente.";
  }
}

/**
 * Altera o status do lançamento na linha especificada para "Pago".
 * Verifica se o status atual é "Em aberto" antes de atualizar.
 *
 * @param {number} rowIndex Número da linha a ser atualizada.
 * @param {number|string} mes Mês selecionado.
 * @param {number|string} ano Ano selecionado.
 * @return {string} Mensagem de sucesso ou erro.
 */
function pagarRegistro(rowIndex, mes, ano) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
    if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
    
    // Lê toda a linha para verificar o status atual
    var rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (rowData[9].toString().toLowerCase() !== "em aberto") {
      return "Lançamento não está em aberto.";
    }
    
    // Atualiza a célula da coluna J (Status) para "Pago"
    sheet.getRange(rowIndex, 10).setValue("Pago");
    SpreadsheetApp.flush();
    
    var novoStatus = sheet.getRange(rowIndex, 10).getValue();
    if (novoStatus === "Pago") {
      return "Lançamento marcado como Pago com sucesso!";
    } else {
      return "Falha ao atualizar o status para 'Pago'.";
    }
  } catch (error) {
    Logger.log("Erro ao marcar registro como pago: " + error.message);
    return "Erro ao marcar registro como pago.";
  }
}