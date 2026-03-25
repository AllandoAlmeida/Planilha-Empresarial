function irParaBalancoMensal() {
    irParaAba("Balanço Mensal");
}

function irParaBalancoAnual() {
    irParaAba("Balanço Anual");
}

function irParaLancamentos() {
    irParaAba("Lançamentos");
}

function irParaBalancoFinanceiro() {
    irParaAba("Balanço Financeiro");
}

function irParaInvestimentos() {
    irParaAba("Investimentos");
}

function irParaCadastro() {
    irParaAba("Cadastro");
}

// Função genérica para mudar de aba
function irParaAba(nomeAba) {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var aba = planilha.getSheetByName(nomeAba);
    
    if (aba) {
        planilha.setActiveSheet(aba);
    } else {
        SpreadsheetApp.getUi().alert("Aba '" + nomeAba + "' não encontrada!");
    }
}
