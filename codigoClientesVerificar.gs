/**
 * Cria o menu personalizado ao abrir a planilha.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Menu Personalizado')
    .addItem('Consultar Clientes', 'abrirPopupClientes')
    .addToUi();
}

/**
 * Abre o popup HTML para consulta de clientes.
 */
function abrirPopupClientes() {
  const html = HtmlService.createHtmlOutputFromFile("popupClientes")
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Consulta de Clientes");
}

/**
 * Busca dados dos clientes na aba "Clientes".
 * @return {Array} Array de clientes no formato [Nome, CPF/CNPJ, Telefone, Email, Endereço].
 */
function buscarClientes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Clientes");

    if (!sheet) {
      throw new Error("A aba 'Clientes' não foi encontrada!");
    }

    // Lê dados a partir da linha 2 (ignora cabeçalho)
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange("A2:E" + lastRow).getValues();

    // Filtra linhas onde a coluna A (Nome) não está vazia
    const filteredData = data.filter(row => row[0].toString().trim() !== "");

    return filteredData;
  } catch (erro) {
    Logger.log("Erro em buscarClientes: " + erro.message);
    return [];
  }
}

/**
 * Salva as alterações do cliente na planilha.
 * @param {number} index Índice do cliente no array (linha na planilha = index + 2).
 * @param {Array} novosDados Dados do cliente no formato [Nome, CPF/CNPJ, Telefone, Email, Endereço].
 */
function salvarCliente(index, novosDados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Clientes");

    if (!sheet) {
      throw new Error("A aba 'Clientes' não foi encontrada!");
    }

    // Calcula a linha na planilha (linha 1 = cabeçalho)
    const linhaNaPlanilha = index + 2;

    // Atualiza os dados nas colunas A a E
    sheet.getRange(`A${linhaNaPlanilha}:E${linhaNaPlanilha}`)
      .setValues([novosDados]);

    Logger.log(`Dados atualizados na linha ${linhaNaPlanilha}: ${novosDados}`);
  } catch (erro) {
    Logger.log("Erro em salvarCliente: " + erro.message);
    throw erro; // Repassa o erro para ser exibido no HTML
  }
}