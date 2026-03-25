/**
 * Abre o pop-up para cadastro de cliente.
 * O arquivo HTML deve ser nomeado exatamente "popupCliente.html".
 */
function abrirPopupCliente() {
  try {
    var html = HtmlService.createHtmlOutputFromFile("popupCliente")
      .setWidth(500)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, "Cadastro de Cliente");
  } catch (error) {
    Logger.log("Erro ao abrir pop-up: " + error.message);
  }
}

/**
 * Salva os dados do cliente na aba "Clientes".
 * Os dados são gravados nas colunas:
 *   A: Nome/Razão Social  
 *   B: CPF/CNPJ  
 *   C: Telefone  
 *   D: E-mail  
 *   E: Endereço  
 *   F: Data de Cadastro (data atual formatada)  
 *   G: Observações
 *
 * @param {Object} dados Objeto com as propriedades: nome, cpfCnpj, telefone, email, endereco, observacoes.
 * @return {string} Mensagem de sucesso ou erro.
 */
function salvarCliente(dados) {
  Logger.log("Dados recebidos: " + JSON.stringify(dados));
  if (!dados) return "Erro: Nenhum dado recebido.";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Clientes");
  if (!aba) return "Erro: Aba 'Clientes' não encontrada.";
  
  // Usa a data atual como Data de Cadastro
  var dataCadastro = new Date();
  var dataFormatada = Utilities.formatDate(dataCadastro, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  
  var novaLinha = [
    dados.nome || "",
    dados.cpfCnpj || "",
    dados.telefone || "",
    dados.email || "",
    dados.endereco || "",
    dataFormatada,
    dados.observacoes || ""
  ];
  
  aba.appendRow(novaLinha);
  return "Cliente cadastrado com sucesso!";
}
