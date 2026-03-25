/**
 * Cria um menu personalizado na planilha quando ela é aberta.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Menu Personalizado')
    .addItem('Abrir Venda', 'abrirPopupVenda')
    .addToUi();
}

/**
 * Abre o popup para realizar a venda.
 * O arquivo HTML deve ser nomeado exatamente "popupVenda.html".
 */
function abrirPopupVenda() {
  try {
    var html = HtmlService.createHtmlOutputFromFile("popupVenda")
      .setWidth(700)
      .setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(html, "Realizar Venda");
  } catch (error) {
    Logger.log("Erro ao abrir popup de venda: " + error.message);
  }
}

/**
 * Obtém os produtos da aba "Produtos e Serviços".
 * Dados começam na linha 1 (sem cabeçalho).
 * Espera:
 *  - Coluna B (índice 1): Descrição do Produto.
 *  - Coluna E (índice 4): Preço de Venda (Valor Unitário).
 *  - Coluna F (índice 5): Estoque.
 */
function obterProdutos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produtos e Serviços");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var produtos = [];
  for (var i = 0; i < data.length; i++) {
    var desc = data[i][1];
    var precoUnitario = data[i][4];
    var estoque = data[i][5];
    if (desc && desc.toString().trim() !== "") {
      produtos.push({
        descricao: desc.toString().trim(),
        precoUnitario: precoUnitario,
        estoque: estoque
      });
    }
  }
  return produtos;
}

/**
 * Obtém os clientes da aba "Clientes".
 * Dados começam na linha 1 (sem cabeçalho).
 * Supõe que a Coluna A (índice 0) contenha o Nome do Cliente.
 */
function obterClientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  var data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var clientes = [];
  for (var i = 0; i < data.length; i++) {
    var nome = data[i][0];
    if (nome && nome.toString().trim() !== "") {
      clientes.push({
        nome: nome.toString().trim()
      });
    }
  }
  return clientes;
}

/**
 * Obtém os cartões de crédito da aba "Cadastro" (intervalo E6:E20).
 */
function obterCartoesCredito() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cadastro");
  if (!sheet) return [];
  var data = sheet.getRange("E6:E20").getValues();
  var cartoes = [];
  for (var i = 0; i < data.length; i++) {
    var cartao = data[i][0];
    if (cartao && cartoes.indexOf(cartao.toString().trim()) === -1) {
      cartoes.push(cartao.toString().trim());
    }
  }
  return cartoes;
}

/**
 * Obtém os bancos da aba "Cadastro" (intervalo L6:L20).
 */
function obterBancos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cadastro");
  if (!sheet) return [];
  var data = sheet.getRange("L6:L20").getValues();
  var bancos = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== "") {
      bancos.push(data[i][0].toString().trim());
    }
  }
  return bancos;
}

/**
 * Gera um código curto e único para o pedido.
 * Exemplo de código gerado: "PEDABC123"
 */
function gerarCodigoPedido() {
  var code = Math.random().toString(36).substr(2, 6).toUpperCase();
  return "PED" + code;
}

/**
 * Registra uma venda com múltiplos itens.
 * O objeto "dados" deve ter a seguinte estrutura:
 * {
 *   cliente: string,
 *   dataVenda: string (formato YYYY-MM-DD),
 *   tipoPagamento: string, // Ex.: "Dinheiro", "Cartão", "Transferência", "Boleto"
 *   pagamentoDetalhe: string, // Se "Cartão", nome do cartão; se "Transferência" ou "Boleto", nome do banco; se Dinheiro, pode ser vazio.
 *   vencimento: string (opcional, formato YYYY-MM-DD),
 *   statusVenda: string, // "Pago" ou "Em Aberto"
 *   itens: [
 *      { produto: string, quantidade: number, valorUnitario: number, valorTotal: number },
 *      ...
 *   ]
 * }
 *
 * Ao registrar a venda, a função separa a data em duas novas colunas: "Mês" e "Ano".
 */
function registrarVendaMulti(dados) {
  if (!dados) return "Erro: Nenhum dado recebido.";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProd = ss.getSheetByName("Produtos e Serviços");
  if (!sheetProd) return "Erro: Aba 'Produtos e Serviços' não encontrada.";
  
  // Gera um código único para o pedido.
  var orderCode = gerarCodigoPedido();
  
  // Lê os dados dos produtos.
  var prodLastRow = sheetProd.getLastRow();
  var prodData = sheetProd.getRange(1, 1, prodLastRow, sheetProd.getLastColumn()).getValues();
  
  // Abre ou cria a aba de Vendas.
  var sheetVendas = ss.getSheetByName("Vendas");
  if (!sheetVendas) {
    sheetVendas = ss.insertSheet("Vendas");
    // Cabeçalho da aba Vendas, incluindo as novas colunas "Mês" e "Ano"
    sheetVendas.appendRow([
      "Código Pedido", "Data", "Produto", "Cliente", "Quantidade", 
      "Valor Unitário", "Total", "Tipo de Pagamento", "Detalhe", 
      "Vencimento", "Status", "Mês", "Ano"
    ]);
  }
  
  var msg = "";
  var updates = [];     // Para acumular atualizações de estoque
  var rowsToAppend = []; // Para acumular registros de vendas
  
  dados.itens.forEach(function(item) {
    // Procura o produto na aba "Produtos e Serviços" pela descrição (coluna B, índice 1)
    var found = null;
    var rowIndexProd = null;
    for (var i = 0; i < prodData.length; i++) {
      if (prodData[i][1].toString().trim() === item.produto.toString().trim()) {
        found = prodData[i];
        rowIndexProd = i + 1;
        break;
      }
    }
    if (!found) {
      msg += "Produto '" + item.produto + "' não encontrado. ";
      return; // pula este item
    }
    
    var estoqueAtual = parseFloat(found[5]) || 0;
    var quantidade = parseFloat(item.quantidade) || 0;
    if (quantidade > estoqueAtual) {
      msg += "Estoque insuficiente para '" + item.produto + "'. Estoque atual: " + estoqueAtual + ". ";
      return;
    }
    
    var novoEstoque = estoqueAtual - quantidade;
    updates.push({ row: rowIndexProd, newEstoque: novoEstoque });
    
    // Converte a data da venda para objeto Date
    var dataVenda = dados.dataVenda ? new Date(dados.dataVenda) : new Date();
    // Formata a data no formato "dd/MM/yyyy"
    var dataStr = Utilities.formatDate(dataVenda, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    // Extrai o mês como número (1 a 12) e o ano
    var mesValor = parseInt(Utilities.formatDate(dataVenda, ss.getSpreadsheetTimeZone(), "M"), 10);
    var anoValor = Utilities.formatDate(dataVenda, ss.getSpreadsheetTimeZone(), "yyyy");
    
    var precoUnit = parseFloat(item.valorUnitario) || 0;
    var totalVenda = parseFloat(item.valorTotal) || (precoUnit * quantidade);
    
    // Adiciona a linha com os dados e as novas colunas "Mês" e "Ano"
    rowsToAppend.push([
      orderCode,
      dataStr,
      item.produto,
      dados.cliente,
      quantidade,
      precoUnit,
      totalVenda,
      dados.tipoPagamento,
      dados.pagamentoDetalhe,
      dados.vencimento || "",
      dados.statusVenda,
      mesValor,
      anoValor
    ]);
  });
  
  // Atualiza o estoque na aba "Produtos e Serviços"
  updates.forEach(function(update) {
    sheetProd.getRange(update.row, 6).setValue(update.newEstoque);
  });
  
  // Registra as linhas de vendas na aba "Vendas"
  rowsToAppend.forEach(function(row) {
    sheetVendas.appendRow(row);
  });
  
  SpreadsheetApp.flush();
  return "Venda registrada com sucesso! " + (msg ? "Observações: " + msg : "");
}
