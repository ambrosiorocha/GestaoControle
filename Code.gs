// Função 'include' para carregar arquivos CSS/JS no HTML
function include(filename) {
  // Divide o nome do ficheiro pela barra "/"
  var parts = filename.split('/');

  // Pega apenas a última parte (o nome do ficheiro real)
  var flattenedName = parts[parts.length - 1];

  // Carrega o ficheiro com o nome aplainado
  return HtmlService.createHtmlOutputFromFile(flattenedName)
      .getContent();
}

// --- Conteúdo do antigo 'Código.js' ---

function doPost(e) {
  try {
    var contents = e.postData ? e.postData.contents : '';
    if (!contents) return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Payload vazio' })).setMimeType(ContentService.MimeType.JSON);
    
    var requestData;
    try {
      requestData = JSON.parse(contents);
    } catch (parseErr) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Erro ao processar JSON: ' + parseErr.message })).setMimeType(ContentService.MimeType.JSON);
    }

    // Normalização: aceita 'action' ou 'acao'
    var action = requestData.action || requestData.acao;
    var data = requestData.data || requestData; // Se não houver 'data', tenta ler da raiz
    var result;

    if (!action) return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Ação (action) não informada' })).setMimeType(ContentService.MimeType.JSON);
    
    // Suporte específico para campos sugeridos pelo usuário no atualizarCredenciais
    if (action === 'atualizarCredenciais') {
      if (!data.usuario && requestData.novoUsuario) data.usuario = requestData.novoUsuario;
      if (data.senha === undefined && requestData.novaSenha !== undefined) data.senha = requestData.novaSenha;
    }

    var isWriteAction = action.startsWith('salvar') || action.startsWith('excluir') || action.startsWith('lancar') || action.startsWith('finalizar') || action.startsWith('estornar') || action.startsWith('baixar') || action === 'arquivarVendasAntigas';
    var lock = null;
    
    if (isWriteAction) {
      lock = LockService.getScriptLock();
      lock.waitLock(30000); // Aguarda até 30 segundos
    }

    try {
      switch (action) {
        case 'lancarVenda':
          result = lancarVenda(data);
          break;
        case 'salvarRascunho':
          result = salvarRascunho(data);
          break;
        case 'finalizarPendente':
          result = finalizarPendente(data);
          break;
        case 'estornarVenda':
          result = estornarVenda(data.id);
          break;
        case 'obterProdutos':
          result = { status: 'sucesso', dados: obterProdutos() };
          break;
        case 'salvarProduto':
          result = salvarProduto(data);
          break;
        case 'excluirProduto':
          result = excluirProduto(data.id);
          break;
        case 'obterProdutoPorId':
          result = { status: 'sucesso', dados: obterProdutoPorId(data.id) };
          break;
        case 'obterProdutosUnicos':
          result = { status: 'sucesso', dados: obterProdutosUnicos() };
          break;
        case 'obterVendas':
          result = { status: 'sucesso', dados: obterVendas(data) };
          break;
        case 'arquivarVendasAntigas':
          result = executarArquivamento();
          break;
        case 'obterClientes':
          result = { status: 'sucesso', dados: obterDadosGeral("Clientes") };
          break;
        case 'salvarCliente':
          result = salvarDadosGeral("Clientes", data);
          break;
        case 'excluirCliente':
          result = excluirDadosGeral("Clientes", data.id);
          break;
        case 'obterFornecedores':
          result = { status: 'sucesso', dados: obterDadosGeral("Fornecedores") };
          break;
        case 'salvarFornecedor':
          result = salvarDadosGeral("Fornecedores", data);
          break;
        case 'excluirFornecedor':
          result = excluirDadosGeral("Fornecedores", data.id);
          break;
        case 'obterFinanceiro':
          result = { status: 'sucesso', dados: obterDadosGeral("Financeiro") };
          break;
        case 'salvarFinanceiro':
          result = salvarDadosGeral("Financeiro", data);
          break;
        case 'excluirFinanceiro':
          result = excluirDadosGeral("Financeiro", data.id);
          break;
        case 'obterOperadores':
          result = { status: 'sucesso', dados: obterOperadores() };
          break;
        case 'salvarOperador':
          result = salvarOperador(data);
          break;
        case 'excluirOperador':
          result = excluirOperador(data.nome);
          break;
        case 'autenticarOperador':
          result = autenticarOperador(data);
          break;
        case 'verificarPrimeiroAcesso':
          result = verificarPrimeiroAcesso();
          break;
        case 'primeiroAcesso':
          result = realizarPrimeiroAcesso(data);
          break;
        case 'registrarMestra':
          result = registrarMestra(data);
          break;
        case 'atualizarCredenciais':
          result = atualizarCredenciais(data);
          break;
        case 'baixarLancamento':
          result = baixarLancamento(data.id);
          break;
        default:
          result = { status: 'erro', mensagem: 'Ação não reconhecida: ' + action };
      }
    } finally {
      if (lock) {
        lock.releaseLock();
      }
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'erro', mensagem: 'Erro no servidor: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================================================
// HELPER: Baixa de Estoque (multi-item)
// ==================================================
function baixarEstoqueItens(sheetProdutos, itensList) {
  if (!sheetProdutos || sheetProdutos.getLastRow() < 2 || !itensList || itensList.length === 0) return null;
  var dadosProd = sheetProdutos.getDataRange().getValues();
  var colNome = dadosProd[0].indexOf('Nome');
  var colQtd  = dadosProd[0].indexOf('Quantidade');
  if (colNome === -1 || colQtd === -1) return 'Colunas Nome/Quantidade não encontradas em Produtos.';
  // Valida primeiro
  for (var k = 0; k < itensList.length; k++) {
    var nm = String(itensList[k].nome).trim();
    var qt = parseFloat(itensList[k].quantidade) || 0;
    var found = false;
    for (var i = 1; i < dadosProd.length; i++) {
      if (String(dadosProd[i][colNome]).trim() === nm) {
        if ((parseFloat(dadosProd[i][colQtd]) || 0) < qt)
          return '❌ Estoque insuficiente para "' + nm + '"! Disponível: ' + dadosProd[i][colQtd];
        found = true; break;
      }
    }
    if (!found) return 'Produto "' + nm + '" não encontrado no estoque.';
  }
  // Subtrai
  for (var k = 0; k < itensList.length; k++) {
    var nm = String(itensList[k].nome).trim();
    var qt = parseFloat(itensList[k].quantidade) || 0;
    for (var i = 1; i < dadosProd.length; i++) {
      if (String(dadosProd[i][colNome]).trim() === nm) {
        var novo = (parseFloat(dadosProd[i][colQtd]) || 0) - qt;
        sheetProdutos.getRange(i + 1, colQtd + 1).setValue(novo);
        dadosProd[i][colQtd] = novo;
        break;
      }
    }
  }
  return null; // sem erro
}

// Helper: devolve estoque
function devolverEstoqueItens(sheetProdutos, itensList) {
  if (!sheetProdutos || !itensList || itensList.length === 0) return;
  var dadosProd = sheetProdutos.getDataRange().getValues();
  var colNome = dadosProd[0].indexOf('Nome');
  var colQtd  = dadosProd[0].indexOf('Quantidade');
  if (colNome === -1 || colQtd === -1) return;
  for (var k = 0; k < itensList.length; k++) {
    var nm = String(itensList[k].nome).trim();
    var qt = parseFloat(itensList[k].quantidade) || 0;
    for (var i = 1; i < dadosProd.length; i++) {
      if (String(dadosProd[i][colNome]).trim() === nm) {
        sheetProdutos.getRange(i + 1, colQtd + 1).setValue((parseFloat(dadosProd[i][colQtd]) || 0) + qt);
        break;
      }
    }
  }
}

// Helper: gera ID sequencial para Vendas
function proximoIdVendas(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return 1;
  var val = sheet.getRange(last, 1).getValue();
  return (parseInt(val) || 0) + 1;
}

// ==================================================
// ESTRUTURA DA ABA VENDAS (colunas por posicao)
// Col 1:ID  2:Data  3:Cliente  4:Itens  5:Qtd  6:Subtotal
// Col 7:Desc%  8:DescR$  9:Total  10:FormaPgto  11:Usuario
// Col 12:Status  13:Vencimento  14:ItensJSON
// ==================================================

// Salvar como Rascunho/Pendente — SEM estoque, SEM financeiro
function salvarRascunho(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Vendas');
  if (!sheet) return { status: 'erro', mensagem: 'Aba Vendas não encontrada.' };
  var novoId = proximoIdVendas(sheet);
  sheet.appendRow([
    novoId, dados.data,
    dados.cliente || 'Consumidor Interno',
    dados.itens, dados.quantidadeVendida,
    dados.subtotal, dados.descontoPercentual, dados.descontoReal,
    dados.totalComDesconto,
    dados.formaPagamento || '',
    dados.usuario || '',
    'Pendente',            // col 12: Status
    '',                    // col 13: Vencimento (definido ao finalizar)
    JSON.stringify(dados.itensList || [])  // col 14: ItensJSON
  ]);
  return { status: 'sucesso', mensagem: '💾 Rascunho #' + novoId + ' salvo! Finalize quando quiser.', id: novoId };
}

// Finalização direta (nova venda, sem rascunho prévio)
function lancarVenda(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVendas   = ss.getSheetByName('Vendas');
  var sheetProdutos = ss.getSheetByName('Produtos');
  var sheetFin      = ss.getSheetByName('Financeiro');
  if (!sheetVendas) return { status: 'erro', mensagem: 'Aba Vendas não encontrada.' };

  // Baixa de estoque
  var erro = baixarEstoqueItens(sheetProdutos, dados.itensList);
  if (erro) return { status: 'erro', mensagem: erro };

  var novoId = proximoIdVendas(sheetVendas);
  var vencimento  = dados.vencimento  || dados.data;
  var statusFin   = dados.statusFinanceiro || 'Pendente';

  sheetVendas.appendRow([
    novoId, dados.data,
    dados.cliente || 'Consumidor Interno',
    dados.itens, dados.quantidadeVendida,
    dados.subtotal, dados.descontoPercentual, dados.descontoReal,
    dados.totalComDesconto,
    dados.formaPagamento || '',
    dados.usuario || '',
    'Concluda',            // col 12: Status
    vencimento,            // col 13: Vencimento
    JSON.stringify(dados.itensList || [])   // col 14: ItensJSON
  ]);

  if (sheetFin) {
    var ultimIdFin = sheetFin.getLastRow() > 1 ? sheetFin.getRange(sheetFin.getLastRow(), 1).getValue() : 0;
    sheetFin.appendRow([
      ultimIdFin + 1,
      'Venda #' + novoId + ' - ' + (dados.cliente || 'Consumidor'),
      dados.totalComDesconto, 'Receber',
      vencimento, statusFin, 'Venda', novoId
    ]);
  }
  return { status: 'sucesso', mensagem: '✅ Venda #' + novoId + ' concluída!', id: novoId };
}

// Finaliza uma venda Pendente já existente
function finalizarPendente(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVendas   = ss.getSheetByName('Vendas');
  var sheetProdutos = ss.getSheetByName('Produtos');
  var sheetFin      = ss.getSheetByName('Financeiro');
  if (!sheetVendas) return { status: 'erro', mensagem: 'Aba Vendas não encontrada.' };

  // Localiza a linha da venda pelo ID
  var todosDados = sheetVendas.getDataRange().getValues();
  var linhaVenda = -1;
  for (var i = 1; i < todosDados.length; i++) {
    if (String(todosDados[i][0]) === String(dados.id)) { linhaVenda = i + 1; break; }
  }
  if (linhaVenda === -1) return { status: 'erro', mensagem: 'Venda #' + dados.id + ' não encontrada.' };

  // Lê itensList do JSON armazenado (col 14)
  var itensList = [];
  try { itensList = JSON.parse(todosDados[linhaVenda - 1][13] || '[]'); } catch(e) {}
  if (dados.itensList && dados.itensList.length > 0) itensList = dados.itensList;

  // Baixa de estoque
  var erro = baixarEstoqueItens(sheetProdutos, itensList);
  if (erro) return { status: 'erro', mensagem: erro };

  var vencimento = dados.vencimento || dados.data || todosDados[linhaVenda - 1][1];
  var statusFin  = dados.statusFinanceiro || 'Pendente';
  var total      = parseFloat(todosDados[linhaVenda - 1][8]) || 0;
  var cliente    = todosDados[linhaVenda - 1][2] || 'Consumidor';
  var pgto       = dados.formaPagamento || todosDados[linhaVenda - 1][9] || '';

  // Atualiza colunas Status (12), Vencimento (13), FormaPgto (10)
  sheetVendas.getRange(linhaVenda, 10).setValue(pgto);
  sheetVendas.getRange(linhaVenda, 12).setValue('Concluda');
  sheetVendas.getRange(linhaVenda, 13).setValue(vencimento);

  // Cria registro financeiro
  if (sheetFin) {
    var ultimIdFin = sheetFin.getLastRow() > 1 ? sheetFin.getRange(sheetFin.getLastRow(), 1).getValue() : 0;
    sheetFin.appendRow([
      ultimIdFin + 1,
      'Venda #' + dados.id + ' - ' + cliente,
      total, 'Receber', vencimento, statusFin, 'Venda', dados.id
    ]);
  }
  return { status: 'sucesso', mensagem: '✅ Venda #' + dados.id + ' finalizada!' };
}

// Estorna uma venda Concluda: devolve estoque e cancela financeiro
function estornarVenda(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVendas   = ss.getSheetByName('Vendas');
  var sheetProdutos = ss.getSheetByName('Produtos');
  var sheetFin      = ss.getSheetByName('Financeiro');
  if (!sheetVendas) return { status: 'erro', mensagem: 'Aba Vendas não encontrada.' };

  var todosDados = sheetVendas.getDataRange().getValues();
  var linhaVenda = -1;
  for (var i = 1; i < todosDados.length; i++) {
    if (String(todosDados[i][0]) === String(id)) { linhaVenda = i + 1; break; }
  }
  if (linhaVenda === -1) return { status: 'erro', mensagem: 'Venda #' + id + ' não encontrada.' };

  // Devolve estoque
  var itensList = [];
  try { itensList = JSON.parse(todosDados[linhaVenda - 1][13] || '[]'); } catch(e) {}
  devolverEstoqueItens(sheetProdutos, itensList);

  // Marca venda como Estornada
  sheetVendas.getRange(linhaVenda, 12).setValue('Estornada');

  // Cancela o registro financeiro ligado
  if (sheetFin && sheetFin.getLastRow() > 1) {
    var dadosFin = sheetFin.getDataRange().getValues();
    for (var i = 1; i < dadosFin.length; i++) {
      // Procura pelo campo referencia (col 8 = novoId)
      if (String(dadosFin[i][7]) === String(id) && dadosFin[i][3] === 'Receber') {
        sheetFin.getRange(i + 1, 6).setValue('Estornado');
        break;
      }
    }
  }
  return { status: 'sucesso', mensagem: '↩️ Venda #' + id + ' estornada. Estoque devolvido.' };
}

// Marca lançamento financeiro como Pago/Recebido
function baixarLancamento(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Financeiro');
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: 'erro', mensagem: 'Planilha Financeiro não encontrada ou vazia.' };
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colId = headers.indexOf('id');
  var colStatus = headers.indexOf('status');
  if (colId === -1 || colStatus === -1) {
    return { status: 'erro', mensagem: 'Colunas id/status não encontradas na aba Financeiro.' };
  }
  var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i][colId]) === String(id)) {
      sheet.getRange(i + 2, colStatus + 1).setValue('Pago');
      return { status: 'sucesso', mensagem: 'Lançamento #' + id + ' baixado com sucesso!' };
    }
  }
  return { status: 'erro', mensagem: 'Lançamento #' + id + ' não encontrado.' };
}

function obterProdutos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produtos");
  if (!sheet || sheet.getLastRow() < 2) {
    return { compact: true, headers: [], rows: [] };
  }
  var dados = sheet.getDataRange().getValues();
  var headers = dados.shift();
  return { compact: true, headers: headers, rows: dados };
}

function obterProdutosUnicos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProdutos = ss.getSheetByName("Produtos");
  if (!sheetProdutos || sheetProdutos.getLastRow() < 2) {
    return [];
  }
  var dadosProdutos = sheetProdutos.getRange(2, 2, sheetProdutos.getLastRow() - 1, 1).getValues();
  var produtosUnicos = Array.from(new Set(dadosProdutos.flat()));
  return produtosUnicos.sort();
}

function obterProdutoPorId(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Produtos");
  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }
  var dados = sheet.getDataRange().getValues();
  var headers = dados.shift();
  var produto = dados.find(function(row) {
    return row[0] == id;
  });
  if (produto) {
    var obj = {};
    headers.forEach(function(header, i) {
      obj[header] = produto[i];
    });
    return obj;
  }
  return null;
}

function obterVendas(filtros) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVendas = ss.getSheetByName('Vendas');
  var sheetHistorico = ss.getSheetByName('Historico_Vendas');
  var todosDados = [];
  
  if (sheetVendas && sheetVendas.getLastRow() > 1) {
    todosDados = sheetVendas.getRange(2, 1, sheetVendas.getLastRow() - 1, sheetVendas.getLastColumn()).getValues();
  }

  var limiteAtual = new Date();
  limiteAtual.setDate(limiteAtual.getDate() - 60);
  
  var buscarNoHistorico = false;
  if (filtros && filtros.dataInicio) {
    var dInicioBusca = new Date(filtros.dataInicio + 'T00:00:00');
    if (!isNaN(dInicioBusca) && dInicioBusca < limiteAtual) {
      buscarNoHistorico = true;
    }
  }

  if (buscarNoHistorico && sheetHistorico && sheetHistorico.getLastRow() > 1) {
    var dadosHistorico = sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, sheetHistorico.getLastColumn()).getValues();
    todosDados = todosDados.concat(dadosHistorico);
  }

  function parseDateBr(val) {
    if (val instanceof Date && !isNaN(val)) return val;
    var parts = String(val).split('/');
    if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
    return null;
  }

  var dadosFiltrados = todosDados;

  if (filtros && (filtros.dataInicio || filtros.dataFim)) {
    var reqInicio = filtros.dataInicio ? new Date(filtros.dataInicio + 'T00:00:00') : null;
    var reqFim    = filtros.dataFim    ? new Date(filtros.dataFim + 'T23:59:59') : null;
    
    dadosFiltrados = todosDados.filter(function(row) {
      var dRow = parseDateBr(row[1]);
      if (!dRow) return true;
      if (reqInicio && dRow < reqInicio) return false;
      if (reqFim && dRow > reqFim) return false;
      return true;
    });
  } else {
    // Default 60 dias
    dadosFiltrados = todosDados.filter(function(row) {
      var dRow = parseDateBr(row[1]);
      if (!dRow) return true;
      return dRow >= limiteAtual;
    });
  }

  function fmtDate(val) {
    var d = parseDateBr(val);
    if (d) {
      return d.getDate().toString().padStart(2,'0') + '/' +
             (d.getMonth()+1).toString().padStart(2,'0') + '/' +
             d.getFullYear();
    }
    return String(val || '');
  }

  var rows = dadosFiltrados.map(function(row) {
    return [
      row[0], fmtDate(row[1]), row[2] || '', row[3] || '',
      row[4] || 0, row[5] || 0, row[6] || 0, row[7] || 0, row[8] || 0,
      row[9] || '', row[10] || '', row[11] || '', fmtDate(row[12]), row[13] || '[]'
    ];
  });
  
  var headers = [
    'ID da Venda', 'Data', 'Cliente', 'Itens', 'Quantidade Vendida', 'Subtotal',
    'Desconto (%)', 'Desconto (R$)', 'Total com Desconto', 'Forma de Pagamento',
    'Usuario', 'Status', 'Vencimento', 'ItensJSON'
  ];

  return { compact: true, headers: headers, rows: rows };
}

// ==================================================
// ARQUIVAMENTO HISTÓRICO DE VENDAS
// ==================================================
function executarArquivamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetVendas = ss.getSheetByName('Vendas');
  if (!sheetVendas || sheetVendas.getLastRow() < 2) {
    return { status: 'sucesso', mensagem: 'Nenhuma venda para arquivar.' };
  }
  
  var sheetHistorico = ss.getSheetByName('Historico_Vendas');
  if (!sheetHistorico) {
    sheetHistorico = ss.insertSheet('Historico_Vendas');
    var headers = sheetVendas.getRange(1, 1, 1, sheetVendas.getLastColumn()).getValues();
    sheetHistorico.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
  }
  
  var limiteArquivamento = new Date();
  limiteArquivamento.setDate(limiteArquivamento.getDate() - 365);
  
  var dados = sheetVendas.getDataRange().getValues();
  var linhasParaMover = [];
  var indicesParaExcluir = [];
  
  for (var i = 1; i < dados.length; i++) {
    var val = dados[i][1];
    var rowDate;
    if (val instanceof Date && !isNaN(val)) {
        rowDate = val;
    } else {
        var parts = String(val).split('/');
        if (parts.length === 3) rowDate = new Date(parts[2], parts[1] - 1, parts[0]);
    }
    
    if (rowDate && rowDate < limiteArquivamento) {
      linhasParaMover.push(dados[i]);
      indicesParaExcluir.push(i + 1); // linha no sheets (1-based, onde header = row 1)
    }
  }
  
  if (linhasParaMover.length > 0) {
    var novaLinha = sheetHistorico.getLastRow() + 1;
    sheetHistorico.getRange(novaLinha, 1, linhasParaMover.length, linhasParaMover[0].length).setValues(linhasParaMover);
    
    // Deleta de trás para frente para evitar problemas de offset
    for (var j = indicesParaExcluir.length - 1; j >= 0; j--) {
      sheetVendas.deleteRow(indicesParaExcluir[j]);
    }
    return { status: 'sucesso', mensagem: linhasParaMover.length + ' vendas antigas (>365 dias) arquivadas c/ sucesso.' };
  }
  
  return { status: 'sucesso', mensagem: 'Nenhuma venda com mais de 365 dias para arquivar.' };
}



function salvarProduto(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produtos");
  if (!sheet) {
    return { status: 'erro', mensagem: 'A planilha "Produtos" não foi encontrada.' };
  }
  var idProduto = dados.idProduto;

  // Colunas: B=Nome, C=Unidade, D=Preço_de_custo, E=Margem_de_lucro(%),
  //          F=Margem_de_lucro(R$), G=Preço_de_venda, H=Quantidade, I=Descrição
  var valoresProduto = [
    dados.nome,
    dados.unidadeVenda,
    parseFloat(dados.precoCusto) || 0,
    parseFloat(dados.margemPct)  || 0,
    parseFloat(dados.margemRS)   || 0,
    parseFloat(dados.preco)      || 0,
    parseFloat(dados.quantidade) || 0,
    dados.descricao || ''
  ];

  if (idProduto) {
    var dadosSheet = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var linha = dadosSheet.findIndex(function(row) { return row[0] == idProduto; });
    if (linha > -1) {
      var rowNum = linha + 2;
      sheet.getRange(rowNum, 2, 1, valoresProduto.length).setValues([valoresProduto]);
      return { status: 'sucesso', mensagem: `Produto "${dados.nome}" atualizado com sucesso!` };
    } else {
      return { status: 'erro', mensagem: 'Produto não encontrado para atualização.' };
    }
  } else {
    var ultimaLinha = sheet.getLastRow();
    var novoId = (ultimaLinha > 1) ? sheet.getRange(ultimaLinha, 1).getValue() + 1 : 1;
    sheet.appendRow([novoId, ...valoresProduto]);
    return { status: 'sucesso', mensagem: `Produto "${dados.nome}" cadastrado com sucesso!` };
  }
}


function excluirProduto(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Produtos");
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: 'erro', mensagem: 'A planilha "Produtos" não foi encontrada ou está vazia.' };
  }
  var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var linha = dados.findIndex(function(row) { return row[0] == id; });
  if (linha > -1) {
    sheet.deleteRow(linha + 2);
    return { status: 'sucesso', mensagem: 'Produto excluído com sucesso!' };
  } else {
    return { status: 'erro', mensagem: 'Produto não encontrado para exclusão.' };
  }
}

function obterDadosGeral(nomePlanilha) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomePlanilha);
  if (!sheet || sheet.getLastRow() < 2) {
    return { compact: true, headers: [], rows: [] };
  }
  var dados = sheet.getDataRange().getValues();
  var headers = dados.shift();
  return { compact: true, headers: headers, rows: dados };
}

function salvarDadosGeral(nomePlanilha, dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomePlanilha);
  if (!sheet) {
    return { status: 'erro', mensagem: `A planilha "${nomePlanilha}" não foi encontrada.` };
  }
  
  var id = dados.id;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rowValues = headers.map(header => {
    var val = dados[header];
    return val !== undefined ? val : "";
  });

  if (id) {
    var dataIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var rowIndex = dataIds.findIndex(function(row) { return row[0] == id; });
    if (rowIndex > -1) {
      sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([rowValues]);
      return { status: 'sucesso', mensagem: `Registro atualizado com sucesso!` };
    }
    return { status: 'erro', mensagem: 'Registro não encontrado para atualização.' };
  } else {
    var lastRow = sheet.getLastRow();
    var nextId = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;
    rowValues[0] = nextId; // Assume primeira coluna é ID
    sheet.appendRow(rowValues);
    return { status: 'sucesso', mensagem: `Registro cadastrado com sucesso!` };
  }
}

function excluirDadosGeral(nomePlanilha, id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomePlanilha);
  if (!sheet || sheet.getLastRow() < 2) {
    return { status: 'erro', mensagem: `A planilha "${nomePlanilha}" não foi encontrada ou está vazia.` };
  }
  var dadosIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var linha = dadosIds.findIndex(function(row) { return row[0] == id; });
  if (linha > -1) {
    sheet.deleteRow(linha + 2);
    return { status: 'sucesso', mensagem: 'Registro excluído com sucesso!' };
  } else {
    return { status: 'erro', mensagem: 'Registro não encontrado para exclusão.' };
  }
}

// ==================================================
// CONFIGURAÇÕES — Operadores/Usuários Autorizados
// Aba: "Configurações" | Colunas: Nome | Nível | Senha | Plano | Permissões(JSON)
// ==================================================
function obterConfiguracoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Configurações');
  if (!sheet) {
    sheet = ss.insertSheet('Configurações');
    sheet.appendRow(['Nome', 'Nível', 'Senha', 'Plano', 'Permissões']);
    sheet.appendRow(['Administrador', 'Admin', 'admin123', 'Premium', '']);
    sheet.appendRow(['Operador 1', 'Operador', '1234', 'Pro', '']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  } else {
    var lr = sheet.getLastRow();
    var lc = sheet.getLastColumn();
    if (lc < 5) {
      if (lc < 4) {
        sheet.getRange(1, 4).setValue('Plano').setFontWeight('bold');
        if (lr > 1) {
          for (var i = 2; i <= lr; i++) {
            if (!sheet.getRange(i, 4).getValue()) sheet.getRange(i, 4).setValue('Pro');
          }
        }
      }
      sheet.getRange(1, 5).setValue('Permissões').setFontWeight('bold');
    }
    if (lc < 3) {
      if (lc === 1) {
        sheet.getRange('B1').setValue('Nível');
        sheet.getRange('C1').setValue('Senha');
        if (lr > 1) {
          var nomes = sheet.getRange(2, 1, lr - 1, 1).getValues();
          for (var j = 0; j < nomes.length; j++) {
            var isAdm = (String(nomes[j][0]).toLowerCase() === 'administrador');
            sheet.getRange(j + 2, 2).setValue(isAdm ? 'Admin' : 'Operador');
            sheet.getRange(j + 2, 3).setValue(isAdm ? 'admin123' : '1234');
          }
        }
      }
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
  }
  return sheet;
}

// Retorna lista de {nome, nivel, plano, permissoes} — SEM senhas
function FORCAR_AUTORIZACAO() {
  UrlFetchApp.fetch("https://google.com");
}

// ==================================================
// ATUALIZAR CREDENCIAIS E NOTIFICAR MESTRA
// ==================================================
function atualizarCredenciais(dados) {
  var sheet = obterConfiguracoes();
  if (!sheet) return { status: 'erro', mensagem: 'Aba Configurações não encontrada.' };
  
  if (!dados || !dados.usuario) {
    return { status: 'erro', mensagem: 'O usuário não pode ser vazio.' };
  }

  var novoUsuario = String(dados.usuario).trim();
  var novaSenha = dados.senha ? String(dados.senha) : '';

  // Procurar o Administrador
  var lastRow = sheet.getLastRow();
  var atualizado = false;

  if (lastRow > 1) {
    var valores = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < valores.length; i++) {
      var role = String(valores[i][1]).toLowerCase();
      if (role === 'admin' || role === 'administrador') {
        var rowNum = i + 2;
        sheet.getRange(rowNum, 1).setValue(novoUsuario); // Coluna A: Login
        if (novaSenha !== '') {
          sheet.getRange(rowNum, 3).setValue(novaSenha); // Coluna C: Senha
        }
        atualizado = true;
        break;
      }
    }
  }

  if (!atualizado) {
    return { status: 'erro', mensagem: 'Administrador não encontrado na planilha.' };
  }

  // Notificar Mestra sobre a mudança do NOME DE USUÁRIO
  try {
    var idPlanilha = SpreadsheetApp.getActiveSpreadsheet().getId();
    var sUrl = '';
    try { sUrl = ScriptApp.getService().getUrl() || ''; } catch(e){}
    
    var payloadMestra = {
      action: 'atualizarCredencialMestra',
      spreadsheetId: idPlanilha,
      scriptUrl: sUrl,
      usuario: novoUsuario
    };
    notificarMestra(payloadMestra);
  } catch(e) {
    console.error('Falha ao notificar a Mestra: ' + e.message);
  }

  return { status: 'sucesso', mensagem: 'As configurações de conta foram atualizadas.' };
}
function obterOperadores() {
  var sheet = obterConfiguracoes();
  if (sheet.getLastRow() < 2) return [{ nome: 'Administrador', nivel: 'Admin', plano: 'Pro', permissoes: {} }];
  var cols = Math.min(sheet.getLastColumn(), 5);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
  return rows
    .filter(function(r) { return String(r[0]).trim() !== '' && String(r[0]).trim() !== 'Nome'; })
    .map(function(r) {
      var perm = {};
      try { perm = JSON.parse(String(r[4] || '{}')); } catch(e) {}
      return {
        nome:       String(r[0]).trim(),
        nivel:      String(r[1] || 'Operador').trim(),
        plano:      String(r[3] || 'Pro').trim(),
        permissoes: perm
      };
    });
}

// Autentica operador com senha — retorna nome, nivel, plano e permissoes (sem senha)
function autenticarOperador(dados) {
  if (!dados || !dados.nome) return { status: 'erro', mensagem: 'Nome inválido.' };
  var nome  = String(dados.nome).trim();
  var senha = String(dados.senha || '');
  var sheet = obterConfiguracoes();
  if (sheet.getLastRow() < 2) return { status: 'erro', mensagem: 'Nenhum operador cadastrado.' };
  var cols = Math.min(sheet.getLastColumn(), 5);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === nome) {
      var senhaCad = String(rows[i][2]);
      if (senha === senhaCad) {
        var perm = {};
        try { perm = JSON.parse(String(rows[i][4] || '{}')); } catch(e) {}
        
        var licenca = verificarEObterLicenca();
        
        var emp = licenca.empresa;
        if (!emp) {
            try { emp = PropertiesService.getScriptProperties().getProperty('empresaNome') || 'Gestão&Controle'; } catch(e){}
        }
        var nivelUser = String(rows[i][1] || 'Operador').trim();
        var planoUser = String(rows[i][3] || 'Pro').trim();
        var planoFinal = planoUser;
        
        // Aplicação da Trava de Plano/Expiração Mestra
        if (licenca.plano) {
            var lp = licenca.plano.toLowerCase();
            if (lp === 'básico' || lp === 'basico') {
                planoFinal = 'Básico';
            } else if (lp === 'premium' && nivelUser === 'Admin') {
                planoFinal = 'Premium';
            } else if (lp === 'pro' && nivelUser === 'Admin') {
                planoFinal = 'Pro';
            }
        }
        
        return {
          status:     'sucesso',
          nome:       nome,
          nivel:      nivelUser,
          plano:      planoFinal,
          permissoes: perm,
          empresa:    emp
        };
      } else {
        return { status: 'erro', mensagem: 'Senha incorreta.' };
      }
    }
  }
  return { status: 'erro', mensagem: 'Operador não encontrado.' };
}

function salvarOperador(dados) {
  if (!dados || !dados.nome || String(dados.nome).trim() === '') {
    return { status: 'erro', mensagem: 'Nome do operador não pode ser vazio.' };
  }
  var sheet = obterConfiguracoes();
  var nome  = String(dados.nome).trim();
  var nivel = String(dados.nivel || 'Operador').trim();
  var senha = String(dados.senha || '1234');
  var plano = String(dados.plano || 'Pro').trim();
  var permissoes = JSON.stringify(dados.permissoes || { relatorios: true, fiado: true, visaoDono: false });
  if (sheet.getLastRow() > 1) {
    var existentes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r){ return r[0]; });
    if (existentes.indexOf(nome) > -1) return { status: 'erro', mensagem: 'Operador "' + nome + '" já existe.' };
  }
  sheet.appendRow([nome, nivel, senha, plano, permissoes]);
  return { status: 'sucesso', mensagem: 'Operador "' + nome + '" adicionado!' };
}

// ==================================================
// PRIMEIRO ACESSO — Configura admin inicial
// ==================================================

// Retorna true se a aba Configurações estiver vazia ou sem nenhum admin
function verificarPrimeiroAcesso() {
  var sheet = obterConfiguracoes();
  if (sheet.getLastRow() < 2) return { primeiroAcesso: true };
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var temAdmin = rows.some(function(r) {
    return String(r[0]).trim() !== '' &&
           (String(r[1]).toLowerCase() === 'admin' || String(r[1]).toLowerCase() === 'administrador');
  });
  return { primeiroAcesso: !temAdmin };
}

// Cria o primeiro administrador (só funciona se não houver admins)
function realizarPrimeiroAcesso(dados) {
  if (!dados || !dados.login || !dados.senha) {
    return { status: 'erro', mensagem: 'Login e senha são obrigatórios.' };
  }
  // Validação de segurança: só executa se não houver admins
  var check = verificarPrimeiroAcesso();
  if (!check.primeiroAcesso) {
    return { status: 'erro', mensagem: 'O sistema já possui um administrador cadastrado.' };
  }
  var sheet = obterConfiguracoes();
  var login  = String(dados.login).trim();
  var senha  = String(dados.senha);
  var nome   = String(dados.nomeCompleto || login).trim();
  var empresa = String(dados.empresa || '').trim();
  var telefone = String(dados.whatsapp || dados.telefone || '').trim();
  // Verificar login duplicado
  if (sheet.getLastRow() > 1) {
    var existentes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r){ return String(r[0]).trim(); });
    if (existentes.indexOf(login) > -1) return { status: 'erro', mensagem: 'Este login já está em uso.' };
  }
  
  // Auto-Setup de licença inicial (cria aba Licença com Básico por padrão)
  verificarEObterLicenca();

  // Verifica se a tabela de configurações tem a coluna Telefone
  var lc = sheet.getLastColumn();
  if (lc < 6 || sheet.getRange(1, 6).getValue() !== 'Telefone') {
      sheet.getRange(1, 6).setValue('Telefone').setFontWeight('bold');
  }

  // Grava o admin inicial com plano Básico
  var permissoesIniciais = JSON.stringify({relatorios:true,fiado:true,visaoDono:true});
  sheet.appendRow([login, 'Admin', senha, 'Básico', permissoesIniciais, telefone]);

  // Grava nome da empresa nas propriedades do script
  try {
    PropertiesService.getScriptProperties().setProperties({ empresaNome: empresa, adminNome: nome });
  } catch(e) { /* ignora */ }

  // Coleta IDs para enviar à Mestra via notificarMestra()
  var ss2 = SpreadsheetApp.getActiveSpreadsheet();
  var scriptUrlRet = '';
  try { scriptUrlRet = ScriptApp.getService().getUrl() || ''; } catch(e) {}

  // Repassa os dados para a Planilha Mestra (backend-to-backend)
  notificarMestra({
    nome:           empresa,
    usuario:        nome,
    whatsapp:       telefone,
    scriptUrl:      scriptUrlRet,
    spreadsheetId:  ss2.getId(),
    spreadsheetUrl: ss2.getUrl(),
    registro:       dados.registro || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    plano:          'Básico',
    ativacao:       dados.ativacao || '',
    expiracao:      dados.expiracao || ''
  });

  return { status: 'sucesso', mensagem: 'Administrador criado. Faça o login.' };
}

// ==================================================
// NOTIFICAR PLANILHA MESTRA (Backend-to-Backend)
// ==================================================
function notificarMestra(payload) {
  var MASTER_URL = 'https://script.google.com/macros/s/AKfycbzcgNx583XfUKPzJ1Odi28vQeothOb3ned1pSJeXx8Du11IXEoTmbebthTVJPrkkrk0Bg/exec';

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(MASTER_URL, options);
    console.log('[notificarMestra] Status: ' + response.getResponseCode());
    console.log('[notificarMestra] Resposta: ' + response.getContentText());
  } catch(e) {
    console.error('[notificarMestra] FALHA: ' + e.message);
  }
}

// Função de TESTE manual — Execute uma vez no editor do Apps Script para
// conceder a permissão de URL Fetch (OAuth scope: external requests)
function testarConexaoMestra() {
  notificarMestra({
    nome:           'EMPRESA TESTE',
    usuario:        'Teste Manual',
    whatsapp:       '11999999999',
    scriptUrl:      ScriptApp.getService().getUrl(),
    spreadsheetId:  SpreadsheetApp.getActiveSpreadsheet().getId(),
    spreadsheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    registro:       Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    plano:          'Básico',
    ativacao:       '18/03/2026',
    expiracao:      '17/04/2026'
  });
  console.log('[testarConexaoMestra] Concluído. Veja os logs acima.');
}

function excluirOperador(nome) {
  var sheet = obterConfiguracoes();
  if (sheet.getLastRow() < 2) return { status: 'erro', mensagem: 'Nenhum operador cadastrado.' };
  var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === String(nome).trim()) {
      sheet.deleteRow(i + 2);
      return { status: 'sucesso', mensagem: 'Operador removido.' };
    }
  }
  return { status: 'erro', mensagem: 'Operador não encontrado.' };
}



// ==================================================
// CONTROLE DE LICENÇA
// ==================================================
// NOTA: O registro na Planilha Mestra agora é feito
// pelo frontend (js/auth.js) via Dual Dispatch.
// A função registrarMestra() foi removida do backend.

function verificarEObterLicenca() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Usamos 'Licença' porque a 'Configurações' do sistema contém a tabela de usuários e chaves globais ali misturariam
  var sheet = ss.getSheetByName('Licença');
  if (!sheet) {
    sheet = ss.insertSheet('Licença');
    sheet.appendRow(['Propriedade', 'Valor']);
    sheet.appendRow(['Plano', 'Básico']);
    sheet.appendRow(['Expiração', '']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  
  var licenca = { plano: 'Básico', expiracao: '', empresa: '' };
  var dados = sheet.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] === 'Plano') licenca.plano = dados[i][1];
    if (dados[i][0] === 'Expiração') licenca.expiracao = dados[i][1];
    if (dados[i][0] === 'Empresa') licenca.empresa = dados[i][1];
  }
  
  // Trava de expiração
  if (licenca.expiracao) {
    var hoje = new Date();
    hoje.setHours(0,0,0,0);
    var dExp = new Date(licenca.expiracao);
    if (licenca.expiracao instanceof Date) dExp = licenca.expiracao;
    if (!isNaN(dExp) && dExp < hoje) {
      licenca.plano = 'Básico'; // Se passou da expiração, vira básico
    }
  }
  return licenca;
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index.html');
  var html = template.evaluate()
      .setTitle("Gestão&Controle")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}
