/**
 * Cria um menu personalizado na planilha para acessar a verificação de estoque.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Menu Personalizado')
    .addItem('Verificar Estoque', 'abrirPopupEstoque')
    .addToUi();
}

/**
 * Abre o popup para verificar/atualizar o estoque dos produtos.
 * O arquivo HTML deve ser nomeado "popupEstoque.html".
 */
function abrirPopupEstoque() {
  var html = HtmlService.createHtmlOutputFromFile("popupEstoque")
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Verificar Estoque");
}

/**
 * Lê os dados da aba "Produtos e Serviços" e retorna um array de objetos.
 * Estrutura esperada:
 *   Coluna B (índice 1): Descrição do Produto
 *   Coluna E (índice 4): Preço de Venda (Valor Unitário)
 *   Coluna F (índice 5): Estoque
 *
 * @return {Array} Array de objetos com { descricao, preco, estoque }.
 */
function obterEstoqueProdutos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produtos e Serviços");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var produtos = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // Verifica se a linha tem pelo menos 6 colunas (índices 0 a 5)
    if (row.length < 6) continue;
    
    var descricao = row[1] ? row[1].toString().trim() : "";
    // Se o preço estiver formatado como string, remova formatação
    var precoCell = row[4];
    var preco = 0;
    if (typeof precoCell === "string") {
      preco = parseFloat(precoCell.replace("R$", "").replace(/\./g, "").replace(",", ".").trim()) || 0;
    } else {
      preco = parseFloat(precoCell) || 0;
    }
    var estoque = parseFloat(row[5]) || 0;
    
    produtos.push({
      descricao: descricao,
      preco: preco,
      estoque: estoque
    });
  }
  
  return produtos;
}

/**
 * Atualiza o estoque de um produto na aba "Produtos e Serviços".
 * Procura todas as linhas onde a descrição (Coluna B) é igual a productDesc e atualiza a coluna F (Estoque).
 *
 * @param {string} productDesc - Descrição do produto.
 * @param {number} novoEstoque - Novo valor de estoque.
 * @return {string} Mensagem de sucesso.
 */
function updateEstoqueProduto(productDesc, novoEstoque) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produtos e Serviços");
  if (!sheet) return "Erro: aba 'Produtos e Serviços' não encontrada.";
  
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var updated = 0;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // Coluna B (índice 1) contém a descrição
    if (row[1] && row[1].toString().trim() === productDesc.trim()) {
      // Atualiza a coluna F (índice 5)
      sheet.getRange(i + 1, 6).setValue(novoEstoque);
      updated++;
    }
  }
  
  SpreadsheetApp.flush();
  return "Estoque atualizado para \"" + productDesc + "\". Linhas atualizadas: " + updated;
}
