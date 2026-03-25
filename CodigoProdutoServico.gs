/**
 * Abre o pop-up para cadastro de produtos e serviços.
 * O arquivo HTML deve ser nomeado exatamente "popupProdutoServico.html".
 */
function abrirPopupProdutoServico() {
  try {
    var html = HtmlService.createHtmlOutputFromFile("popupProdutoServico")
      .setWidth(500)
      .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, "Cadastro de Produtos e Serviços");
  } catch (error) {
    Logger.log("Erro ao abrir pop-up: " + error.message);
  }
}

/**
 * Obtém as categorias de produtos a partir da aba "Cadastro".
 * Agora, as categorias são lidas do intervalo B6:B60.
 */
function obterCategoriasProdutos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Cadastro");
  if (!aba) return [];
  var dados = aba.getRange("B6:B60").getValues();
  var categorias = [];
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][0] && categorias.indexOf(dados[i][0].toString().trim()) === -1) {
      categorias.push(dados[i][0].toString().trim());
    }
  }
  return categorias;
}

/**
 * Obtém os bancos da aba "Cadastro" (L6:L20).
 */
function obterBancos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Cadastro");
  if (!aba) return [];
  var dados = aba.getRange("L6:L20").getValues();
  var bancos = [];
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][0] && dados[i][0].toString().trim() !== "") {
      bancos.push(dados[i][0].toString().trim());
    }
  }
  return bancos;
}

/**
 * Obtém os cartões de crédito do Cadastro (E6:E20).
 */
function obterCartoesCredito() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Cadastro");
  if (!aba) return [];
  var dados = aba.getRange("E6:E20").getValues();
  var cartoes = [];
  for (var i = 0; i < dados.length; i++) {
    var cartao = dados[i][0];
    if (cartao && cartoes.indexOf(cartao.toString().trim()) === -1) {
      cartoes.push(cartao.toString().trim());
    }
  }
  return cartoes;
}

/**
 * Obtém os tipos de pagamento da aba "Cadastro" (J6:J20).
 */
function obterTiposPagamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Cadastro");
  if (!aba) return [];
  var dados = aba.getRange("J6:J20").getValues();
  var tipos = [];
  for (var i = 0; i < dados.length; i++) {
    if (dados[i][0] && dados[i][0].toString().trim() !== "") {
      tipos.push(dados[i][0].toString().trim());
    }
  }
  return tipos;
}

/**
 * Salva os dados do produto/serviço na aba "Produtos e Serviços".
 *
 * Estrutura da aba "Produtos e Serviços":
 *   A: Código do Produto/Serviço  
 *   B: Descrição/Produto  
 *   C: Categoria  
 *   D: Preço de Custo  
 *   E: Preço de Venda  
 *   F: Estoque  
 *   G: Fornecedor  
 *   H: Data de Cadastro (data atual)  
 *   I: Observações
 *
 * @param {Object} dados Objeto com as propriedades: codigo, descricao, categoria, precoCusto, precoVenda, estoque, fornecedor, observacoes.
 * @return {string} Mensagem de sucesso ou erro.
 */
function salvarProdutoServico(dados) {
  Logger.log("Dados recebidos: " + JSON.stringify(dados));
  if (!dados) return "Erro: Nenhum dado recebido.";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Produtos e Serviços");
  if (!aba) return "Erro: Aba 'Produtos e Serviços' não encontrada.";
  
  // Usa a data atual para Data de Cadastro
  var dataCadastro = new Date();
  var dataFormatada = Utilities.formatDate(dataCadastro, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  
  // Estrutura: A: Código, B: Descrição, C: Categoria, D: Preço de Custo, E: Preço de Venda,
  // F: Estoque, G: Fornecedor, H: Data de Cadastro, I: Observações.
  var novaLinha = [
    dados.codigo || "",
    dados.descricao || "",
    dados.categoria || "",
    parseFloat(dados.precoCusto) || 0,
    parseFloat(dados.precoVenda) || 0,
    parseFloat(dados.estoque) || 0,
    dados.fornecedor || "",
    dataFormatada,
    dados.observacoes || ""
  ];
  
  aba.appendRow(novaLinha);
  return "Produto/Serviço cadastrado com sucesso!";
}
