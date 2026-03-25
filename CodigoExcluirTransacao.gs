/**
 * Abre o pop-up unificado para exclusão de transações.
 * O arquivo HTML deve ser nomeado exatamente "popupExcluirTransacao.html".
 */
function abrirPopupExcluirTransacao() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('popupExcluirTransacao')
      .setWidth(600)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Excluir Transação");
  } catch (error) {
    Logger.log("Erro ao abrir pop-up de exclusão: " + error.message);
  }
}

/**
 * Busca os lançamentos para exclusão, de acordo com o tipo selecionado, mês, ano e (opcional) filtro de descrição.
 * Se o tipo for "Despesa", busca na aba "Banco de Dados Lançamentos" (dados começam na linha 1).
 * Se o tipo for "Receita", busca na aba "Vendas" (dados começam na linha 1).
 *
 * @param {string} tipo "Despesa" ou "Receita"
 * @param {number|string} mes Mês selecionado.
 * @param {number|string} ano Ano selecionado.
 * @param {string} filtro Texto (parcial) para filtrar a descrição.
 * @return {string} HTML com a tabela dos registros encontrados ou mensagem informativa.
 */
function buscarTransacoesExcluir(tipo, mes, ano, filtro) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var html = "";
    if (tipo.toLowerCase() === "despesa") {
      var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
      if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
      var lastRow = sheet.getLastRow();
      if (lastRow < 1) return "Nenhum lançamento encontrado.";
      var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
      var results = [];
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var dt = new Date(row[0]);
        if (isNaN(dt.getTime())) continue;
        var rowMes = dt.getMonth() + 1;
        var rowAno = dt.getFullYear();
        if (rowMes !== Number(mes) || rowAno !== Number(ano)) continue;
        var desc = row[1] ? row[1].toString().trim() : "";
        if (filtro && filtro.trim() !== "" && desc.toLowerCase().indexOf(filtro.toLowerCase()) === -1) continue;
        results.push({
          rowIndex: i + 1,
          data: Utilities.formatDate(dt, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"),
          descricao: desc,
          valor: row[6] ? parseFloat(row[6]).toFixed(2) : "0.00",
          status: row[9] ? row[9].toString().trim() : ""
        });
      }
      if (results.length === 0) return "Nenhum lançamento encontrado para exclusão.";
      html += "<table style='width:100%; border-collapse: collapse;'>";
      html += "<thead><tr style='background-color:#007bff; color:#fff;'>";
      html += "<th>Data</th><th>Descrição</th><th>Valor (R$)</th><th>Status</th><th>Ação</th>";
      html += "</tr></thead><tbody>";
      results.forEach(function(item) {
        html += "<tr style='border: 1px solid #ddd;'>";
        html += "<td>" + item.data + "</td>";
        html += "<td>" + item.descricao + "</td>";
        html += "<td>" + item.valor + "</td>";
        html += "<td>" + item.status + "</td>";
        html += "<td><button onclick=\"excluirTransacao('Despesa', " + item.rowIndex + ", " + mes + ", " + ano + ", '" + item.descricao.replace(/'/g, "\\'") + "')\">Excluir</button></td>";
        html += "</tr>";
      });
      html += "</tbody></table>";
    } else if (tipo.toLowerCase() === "receita") {
      var sheet = ss.getSheetByName("Vendas");
      if (!sheet) return "Erro: Aba 'Vendas' não encontrada.";
      var lastRow = sheet.getLastRow();
      if (lastRow < 1) return "Nenhuma venda encontrada.";
      var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
      var results = [];
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var dt = new Date(row[0]);
        if (isNaN(dt.getTime())) continue;
        var rowMes = dt.getMonth() + 1;
        var rowAno = dt.getFullYear();
        if (rowMes !== Number(mes) || rowAno !== Number(ano)) continue;
        var desc = row[1] ? row[1].toString().trim() : "";
        if (filtro && filtro.trim() !== "" && desc.toLowerCase().indexOf(filtro.toLowerCase()) === -1) continue;
        results.push({
          rowIndex: i + 1,
          data: Utilities.formatDate(dt, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"),
          descricao: desc,
          valor: row[5] ? parseFloat(row[5]).toFixed(2) : "0.00",
          status: "Pago"  // Supondo que vendas registradas já tenham status "Pago"
        });
      }
      if (results.length === 0) return "Nenhuma venda encontrada para exclusão.";
      html += "<table style='width:100%; border-collapse: collapse;'>";
      html += "<thead><tr style='background-color:#007bff; color:#fff;'>";
      html += "<th>Data</th><th>Descrição</th><th>Valor (R$)</th><th>Status</th><th>Ação</th>";
      html += "</tr></thead><tbody>";
      results.forEach(function(item) {
        html += "<tr style='border: 1px solid #ddd;'>";
        html += "<td>" + item.data + "</td>";
        html += "<td>" + item.descricao + "</td>";
        html += "<td>" + item.valor + "</td>";
        html += "<td>" + item.status + "</td>";
        html += "<td><button onclick=\"excluirTransacao('Receita', " + item.rowIndex + ", " + mes + ", " + ano + ", '" + item.descricao.replace(/'/g, "\\'") + "')\">Excluir</button></td>";
        html += "</tr>";
      });
      html += "</tbody></table>";
    } else {
      html = "Tipo de transação inválido.";
    }
    return html;
  } catch (error) {
    Logger.log("Erro em buscarTransacoesExcluir: " + error.message);
    return "Erro ao buscar transações para exclusão.";
  }
}

/**
 * Exclui uma transação com base no tipo.
 * Se tipo for "Despesa", exclui da aba "Banco de Dados Lançamentos".
 * Se tipo for "Receita", exclui da aba "Vendas".
 *
 * @param {string} tipo "Despesa" ou "Receita"
 * @param {number} rowIndex Número da linha a ser excluída.
 * @param {number|string} mes Mês selecionado.
 * @param {number|string} ano Ano selecionado.
 * @param {string} descricao (opcional) Descrição para identificação.
 * @return {string} Mensagem de sucesso ou erro.
 */
function excluirTransacao(tipo, rowIndex, mes, ano, descricao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (tipo.toLowerCase() === "despesa") {
      var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
      if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
      sheet.deleteRow(rowIndex);
      return "Lançamento de despesa excluído com sucesso!";
    } else if (tipo.toLowerCase() === "receita") {
      var sheet = ss.getSheetByName("Vendas");
      if (!sheet) return "Erro: Aba 'Vendas' não encontrada.";
      sheet.deleteRow(rowIndex);
      return "Venda excluída com sucesso!";
    } else {
      return "Tipo de transação inválido.";
    }
  } catch (error) {
    Logger.log("Erro ao excluir transação: " + error.message);
    return "Erro ao excluir transação.";
  }
}
