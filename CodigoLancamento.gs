/**
 * Abre o pop-up para lançamentos.
 * O arquivo HTML deve ser "popupLancamento.html".
 */
function abrirPopupLancamento() {
  var html = HtmlService.createHtmlOutputFromFile("popupLancamento")
      .setWidth(500)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Lançamento Financeiro");
}

/**
 * Obtém as categorias do Cadastro (B6:B60).
 */
function obterCategorias() {
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
 * Salva um lançamento financeiro.
 * Se o tipo de pagamento for "Crédito", cria um lançamento para cada parcela, 
 * dividindo o valor total e ajustando a data (incrementando o mês para cada parcela).
 * Para outros tipos, cria um único lançamento.
 *
 * O objeto "dados" deve conter as seguintes propriedades:
 *   - data: (string) Data base do lançamento (ex.: "2025-03-10")
 *   - descricao: (string) Descrição
 *   - categoria: (string) Categoria
 *   - tipoPagamento: (string) Tipo de Pagamento (ex.: "Pix", "Dinheiro", "Crédito", etc.)
 *   - tipoLancamento: (string) "Receita" ou "Despesa"
 *   - status: (string) "Pago" ou "Em aberto"
 *   - valor: (string ou número) Valor (para lançamentos não-crédito)
 *   - banco: (string) (opcional) Para transações pagas (não crédito)
 *   - cartaoCredito: (string) Para transações de crédito
 *   - parcelas: (string ou número) Número de parcelas para crédito
 *   - valorTotal: (string ou número) Valor total para transações de crédito
 */
function salvarLancamento(dados) {
  Logger.log("Dados recebidos: " + JSON.stringify(dados));
  if (!dados) return "Erro: Nenhum dado recebido.";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Banco de Dados Lançamentos");
  if (!aba) return "Erro: Aba 'Banco de Dados Lançamentos' não encontrada.";
  
  // Converte a data base de string para objeto Date utilizando as partes para evitar problemas de fuso horário.
  var dataBaseStr = dados.data;  // ex.: "2025-03-10"
  var partes = dataBaseStr.split("-");
  var dataBase = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
  
  if (isNaN(dataBase.getTime())) {
    return "Erro: Data inválida!";
  }
  
  var descricao = dados.descricao;
  var categoria = dados.categoria;
  var tipoPagamento = dados.tipoPagamento;
  var tipoLancamento = dados.tipoLancamento; // "Receita" ou "Despesa"
  var status = dados.status; // "Pago" ou "Em aberto"
  
  var cartaoOuConta = "";
  var valor = 0;
  
  // Função auxiliar para inserir um lançamento na planilha.
  // Recebe a data, descrição, valor, etc.
  function adicionarLancamento(dataLanc, desc, val) {
    var mes = dataLanc.getMonth() + 1;
    var ano = dataLanc.getFullYear();
    // Salva a data formatada utilizando o fuso "GMT-3"
    var dataFormatada = Utilities.formatDate(dataLanc, "GMT-3", "yyyy-MM-dd");
    var linha = [
      dataFormatada, // Coluna A: Data
      desc,          // Coluna B: Descrição (pode incluir indicação de parcela)
      cartaoOuConta, // Coluna C: Cartão/Conta
      tipoPagamento, // Coluna D: Tipo de Pagamento
      categoria,     // Coluna E: Categoria
      tipoLancamento,// Coluna F: Tipo (Receita/Despesa)
      val,           // Coluna G: Valor
      mes,           // Coluna H: Mês
      ano,           // Coluna I: Ano
      status         // Coluna J: Status
    ];
    aba.appendRow(linha);
  }
  
  if (tipoPagamento === "Crédito") {
    cartaoOuConta = dados.cartaoCredito;
    var parcelas = parseInt(dados.parcelas) || 1;
    var valorTotal = parseFloat(dados.valorTotal) || 0;
    // Cria um lançamento para cada parcela, incrementando o mês para cada parcela
    for (var i = 0; i < parcelas; i++) {
      var installmentDate = new Date(dataBase);
      installmentDate.setMonth(installmentDate.getMonth() + i);
      var installmentValue = valorTotal / parcelas;
      var installmentDesc = descricao + " (" + (i + 1) + "/" + parcelas + ")";
      adicionarLancamento(installmentDate, installmentDesc, installmentValue);
    }
  } else {
    if (status.toLowerCase() === "pago" && dados.banco) {
      cartaoOuConta = dados.banco;
    } else {
      cartaoOuConta = dados.cartaoConta || "";
    }
    valor = parseFloat(dados.valor) || 0;
    adicionarLancamento(dataBase, descricao, valor);
  }
  
  // Se o lançamento foi salvo com status "Pago" e não for de Crédito, atualiza o saldo do banco
  if (status.toLowerCase() === "pago" && tipoPagamento.toLowerCase() !== "crédito" && dados.banco) {
    var resultado = atualizarSaldoBancoSelecionado(dados.banco, valor);
    Logger.log(resultado);
  }
  
  return "Lançamento salvo com sucesso!";
}

/**
 * Atualiza o saldo do banco na aba "Cadastro" para a conta selecionada.
 */
function atualizarSaldoBancoSelecionado(bankName, valor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Cadastro");
  if (!aba) return "Erro: Aba 'Cadastro' não encontrada.";
  
  var bancos = aba.getRange("L6:L20").getValues();
  var saldos = aba.getRange("M6:M20").getValues();
  for (var i = 0; i < bancos.length; i++) {
    if (bancos[i][0] && bancos[i][0].toString().trim().toLowerCase() === bankName.toLowerCase()) {
      var saldoAtual = parseFloat(saldos[i][0]) || 0;
      var novoSaldo = saldoAtual - valor;
      aba.getRange(i + 6, 13).setValue(novoSaldo);
      return "Saldo do banco '" + bankName + "' atualizado para: R$ " + novoSaldo.toFixed(2);
    }
  }
  return "Banco não encontrado para atualizar saldo.";
}
