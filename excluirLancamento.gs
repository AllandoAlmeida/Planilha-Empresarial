/**
 * Abre o pop-up para exclusão de lançamento.
 * O arquivo HTML deve ser nomeado exatamente "popupExcluirLancamento.html".
 */
function abrirPopUpExcluirLancamento() {
  try {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('popupExcluirLancamento')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Excluir Lançamento");
  } catch (error) {
    console.error("Erro ao abrir pop-up:", error);
  }
}

/**
 * Busca os lançamentos na aba "Banco de Dados Lançamentos" com base no mês e ano selecionados.
 * Os dados começam na linha 1 (não há cabeçalho). A data (coluna A) é utilizada para extrair o mês e o ano.
 * Opcionalmente, pode filtrar por parte da descrição.
 *
 * Retorna um HTML com uma tabela contendo:
 * Data, Descrição, Valor (coluna G), Status (coluna J) e um botão "Excluir" para cada lançamento.
 *
 * @param {number|string} mes O mês para filtrar (ex.: 1 para Janeiro)
 * @param {number|string} ano O ano para filtrar (ex.: 2025)
 * @param {string} descricao Parte da descrição para filtrar (se vazio, retorna todos)
 * @return {string} HTML com a tabela dos lançamentos ou mensagem informativa.
 */
function buscarDetalhesLancamento(mes, ano, descricao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
    if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return "Nenhum lançamento encontrado.";
    
    // Lê todos os dados a partir da linha 1 (pois não há cabeçalho)
    var dataRange = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
    var values = dataRange.getValues();
    
    var results = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      
      // 1) Verifica se a coluna A (data) está preenchida
      if (!row[0] || row[0].toString().trim() === "") {
        continue; // pula a linha se não houver data
      }
      
      // 2) Converte a data em objeto Date e verifica se é válida
      var dateObj = new Date(row[0]);
      if (isNaN(dateObj.getTime())) {
        continue; // pula a linha se não for uma data válida
      }
      
      var desc = row[1] ? row[1].toString().trim() : "";
      var valor = row[6] ? parseFloat(row[6]) : 0; // Coluna G: Valor
      var status = row[9] ? row[9].toString().trim() : "";
      
      var rowMonth = dateObj.getMonth() + 1;
      var rowYear = dateObj.getFullYear();
      
      // Aplica filtro de descrição, se fornecido
      if (descricao && descricao.trim() !== "") {
        if (desc.toLowerCase().indexOf(descricao.toLowerCase()) === -1) {
          continue;
        }
      }
      
      // Compara mês e ano
      if (rowMonth === Number(mes) && rowYear === Number(ano)) {
        // Armazena o índice real da linha na planilha (i + 1)
        results.push({ rowIndex: i + 1, data: dateObj, descricao: desc, valor: valor, status: status });
      }
    }
    
    if (results.length === 0) {
      return "Nenhum lançamento encontrado para o período selecionado.";
    }
    
    var html = "<table style='width:100%; border-collapse: collapse;'>";
    html += "<thead><tr style='background-color:#007bff; color:#fff;'>";
    html += "<th>Data</th><th>Descrição</th><th>Valor (R$)</th><th>Status</th><th>Ação</th>";
    html += "</tr></thead><tbody>";
    
    results.forEach(function(item) {
      var dataStr = Utilities.formatDate(item.data, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
      html += "<tr style='border: 1px solid #ddd;'>";
      html += "<td>" + dataStr + "</td>";
      html += "<td>" + item.descricao + "</td>";
      html += "<td>" + item.valor.toFixed(2) + "</td>";
      html += "<td>" + item.status + "</td>";
      html += "<td><button onclick=\"excluirRegistro(" + item.rowIndex + ", " + mes + ", " + ano + ")\">Excluir</button></td>";
      html += "</tr>";
    });
    
    html += "</tbody></table>";
    return html;
  } catch (error) {
    console.error("Erro ao buscar lançamentos:", error);
    return "Erro ao buscar lançamentos. Verifique o console para mais detalhes.";
  }
}

/**
 * Exclui o lançamento na linha especificada.
 *
 * @param {number} rowIndex O número da linha a ser excluída (na planilha).
 * @param {number|string} mes O mês selecionado.
 * @param {number|string} ano O ano selecionado.
 * @return {string} Mensagem de sucesso ou erro.
 */
function excluirRegistro(rowIndex, mes, ano) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Banco de Dados Lançamentos");
    if (!sheet) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
    
    sheet.deleteRow(rowIndex);
    return "Lançamento excluído com sucesso!";
  } catch (error) {
    console.error("Erro ao excluir lançamento:", error);
    return "Erro ao excluir lançamento. Verifique o console para mais detalhes.";
  }
}
