/**
 * Cria um menu personalizado para acessar a verificação de vendas.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Menu Personalizado')
    .addItem('Verificar Vendas', 'abrirPopupRelatorioVendas')
    .addToUi();
}

/**
 * Abre o popup para verificar vendas.
 */
function abrirPopupRelatorioVendas() {
  const html = HtmlService.createHtmlOutputFromFile("popupRelatorioVendas")
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "Verificar Vendas");
}

/**
 * Formata a data para exibição no formato "dd/MM/yyyy".
 */
function formatarData(data) {
  if (data instanceof Date) {
    return Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return data;
}

/**
 * Busca e agrupa vendas por código, filtrando pelo mês informado.
 * Estrutura da planilha:
 *   A: Código da Venda
 *   B: Data
 *   C: Produto
 *   D: Cliente
 *   E: Quantidade
 *   F: Valor Unitário
 *   G: Total
 */
function buscarVendas(mes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Vendas");
    
    if (!sheet) {
      throw new Error("Aba 'Vendas' não encontrada.");
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const registros = data.slice(1);
    const vendasAgrupadas = {};
    
    registros.forEach(row => {
      if (row.length < 7) return;
      
      const codigo = row[0];
      const dataVendaObj = new Date(row[1]);
      if (isNaN(dataVendaObj.getTime())) return;
      
      const mesVenda = dataVendaObj.getMonth() + 1;
      if (mesVenda !== mes) return;
      
      const dataVenda = formatarData(row[1]);
      const cliente = row[3]; // Coluna D - Cliente
      const total = parseFloat(row[6]) || 0; // Coluna G - Total
      
      if (!vendasAgrupadas[codigo]) {
        vendasAgrupadas[codigo] = {
          codigo: codigo,
          data: dataVenda,
          cliente: cliente,
          total: 0
        };
      }
      vendasAgrupadas[codigo].total += total;
    });
    
    return Object.values(vendasAgrupadas).map(venda => [
      venda.codigo,
      venda.data,
      venda.cliente,
      venda.total.toFixed(2)
    ]);
    
  } catch (erro) {
    Logger.log("Erro em buscarVendas: " + erro.message);
    return [];
  }
}

/**
 * Retorna os detalhes de uma venda a partir do código.
 */
function detalhesVenda(codigoVenda) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Vendas");
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    return data.slice(1)
      .filter(row => row[0] == codigoVenda)
      .map(row => [
        row[2] || "-", // Produto (Coluna C)
        row[4] || "-", // Quantidade (Coluna E)
        isNaN(parseFloat(row[5])) ? "-" : parseFloat(row[5]).toFixed(2), // Valor Unitário (Coluna F)
        isNaN(parseFloat(row[6])) ? "-" : parseFloat(row[6]).toFixed(2)  // Total (Coluna G)
      ]);
      
  } catch (erro) {
    Logger.log("Erro em detalhesVenda: " + erro.message);
    return [];
  }
}

/**
 * Exclui uma venda específica da planilha.
 */
function excluirVenda(codigoVenda) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Vendas");
    
    if (!sheet) {
      throw new Error("Aba 'Vendas' não encontrada.");
    }
    
    const data = sheet.getDataRange().getValues();
    let linhasParaExcluir = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == codigoVenda) {
        linhasParaExcluir.push(i + 1);
      }
    }
    
    linhasParaExcluir.sort((a, b) => b - a).forEach(linha => {
      sheet.deleteRow(linha);
    });
    
    return linhasParaExcluir.length > 0;
    
  } catch (erro) {
    Logger.log("Erro em excluirVenda: " + erro.message);
    throw erro;
  }
}