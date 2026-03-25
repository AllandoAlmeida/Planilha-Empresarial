// Constantes para as colunas da planilha
const COL_DATA = 0; // Coluna A
const COL_CARTAO = 2; // Coluna C
const COL_TIPO_PG = 3; // Coluna D
const COL_VALOR = 6; // Coluna G
const COL_STATUS = 9; // Coluna J

/**
 * Abre o pop-up para ver faturas de cartão de crédito.
 */
function abrirPopupFatura() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('popupFatura')
      .setWidth(600)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Ver Faturas de Cartão");
  } catch (error) {
    console.error("Erro ao abrir pop-up de fatura:", error);
  }
}

/**
 * Busca as faturas de cartão de crédito na planilha.
 * Filtra por mês, ano e tipo de pagamento "Crédito".
 * Retorna um HTML com a tabela de faturas.
 */
function buscarFaturas(mes, ano) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
    if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return "Nenhuma fatura encontrada.";
    
    var range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    var values = range.getValues();
    Logger.log("Dados lidos da planilha: " + JSON.stringify(values)); // Log dos dados
    
    var faturas = {}; // Objeto para agrupar faturas por cartão
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var data = new Date(row[COL_DATA]);
      if (isNaN(data.getTime())) {
        Logger.log("Linha " + (i + 1) + ": Data inválida - " + row[COL_DATA]);
        continue;
      }
      
      var rowMonth = data.getMonth() + 1;
      var rowYear = data.getFullYear();
      if (rowMonth !== Number(mes) || rowYear !== Number(ano)) {
        Logger.log("Linha " + (i + 1) + ": Fora do período - Mês " + rowMonth + ", Ano " + rowYear);
        continue;
      }
      
      var tipoPg = row[COL_TIPO_PG] ? row[COL_TIPO_PG].toString().trim().toLowerCase() : "";
      if (tipoPg !== "crédito") {
        Logger.log("Linha " + (i + 1) + ": Tipo de pagamento inválido - " + tipoPg);
        continue;
      }
      
      var cardName = row[COL_CARTAO] ? row[COL_CARTAO].toString().trim() : "";
      if (!cardName) {
        Logger.log("Linha " + (i + 1) + ": Nome do cartão ausente");
        continue;
      }
      
      var valor = row[COL_VALOR] ? parseFloat(row[COL_VALOR]) : 0;
      var status = row[COL_STATUS] ? row[COL_STATUS].toString().trim().toLowerCase() : "";
      
      if (!faturas[cardName]) {
        faturas[cardName] = { total: 0, status: "pago", linhas: [] };
      }
      faturas[cardName].total += valor;
      faturas[cardName].linhas.push(i + 1); // Armazena o número da linha
      if (status === "em aberto") {
        faturas[cardName].status = "em aberto";
      }
    }
    
    Logger.log("Faturas agrupadas: " + JSON.stringify(faturas)); // Log das faturas
    
    var cards = Object.keys(faturas);
    if (cards.length === 0) return "Nenhuma fatura encontrada para o período selecionado.";
    
    var html = "<table style='width:100%; border-collapse: collapse;'>";
    html += "<thead><tr style='background-color:#007bff; color:#fff;'>";
    html += "<th>Cartão</th><th>Total da Fatura (R$)</th><th>Status</th><th>Ação</th>";
    html += "</tr></thead><tbody>";
    
    cards.forEach(function(cardName) {
      var total = faturas[cardName].total;
      var status = faturas[cardName].status;
      html += "<tr style='border: 1px solid #ddd;'>";
      html += "<td>" + cardName + "</td>";
      html += "<td>" + total.toFixed(2) + "</td>";
      html += "<td>" + status.charAt(0).toUpperCase() + status.slice(1) + "</td>";
      html += "<td>";
      if (status === "em aberto") {
        html += "<button onclick=\"google.script.run.withSuccessHandler(function() { buscarFaturas(" + mes + ", " + ano + "); }).pagarFatura('" + cardName.replace(/'/g, "\\'") + "', " + mes + ", " + ano + ")\">Pagar Fatura</button>";
      } else {
        html += "-";
      }
      html += "</td>";
      html += "</tr>";
    });
    
    html += "</tbody></table>";
    return html;
  } catch (error) {
    Logger.log("Erro em buscarFaturas: " + error.message);
    return "<p style='color: red;'>Erro ao buscar faturas. Verifique os dados e tente novamente.</p>";
  }
}

/**
 * Paga a fatura de um cartão de crédito.
 * Atualiza o status dos lançamentos de "Em aberto" para "Pago".
 */
function pagarFatura(cardName, mes, ano) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
    if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return "Nenhum lançamento encontrado.";
    
    var range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    var values = range.getValues();
    var count = 0;
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var data = new Date(row[COL_DATA]);
      if (isNaN(data.getTime())) continue;
      
      var rowMonth = data.getMonth() + 1;
      var rowYear = data.getFullYear();
      if (rowMonth !== Number(mes) || rowYear !== Number(ano)) continue;
      
      var tipoPg = row[COL_TIPO_PG] ? row[COL_TIPO_PG].toString().trim().toLowerCase() : "";
      if (tipoPg !== "crédito") continue;
      
      var currentCard = row[COL_CARTAO] ? row[COL_CARTAO].toString().trim() : "";
      if (currentCard !== cardName) continue;
      
      var status = row[COL_STATUS] ? row[COL_STATUS].toString().trim().toLowerCase() : "";
      if (status === "em aberto") {
        sheet.getRange(i + 1, COL_STATUS + 1).setValue("Pago");
        count++;
      }
    }
    SpreadsheetApp.flush();
    
    if (count > 0) {
      return "Fatura do cartão '" + cardName + "' marcada como Pago (" + count + " transações atualizadas).";
    } else {
      return "Nenhum lançamento em aberto encontrado para esse cartão no período.";
    }
  } catch (error) {
    Logger.log("Erro ao pagar fatura: " + error.message);
    return "Erro ao pagar fatura.";
  }
}