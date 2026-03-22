/**
 * ==========================================================
 * MASTER SCRIPT - CÉREBRO ÚNICO (API Gateway)
 * ==========================================================
 * Este script centraliza toda a lógica do sistema.
 * As planilhas dos clientes são "burras" e apenas enviam dados para este endpoint.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // Aguarda até 10 segundos
  
  try {
    // 1. Recepção e Parse do JSON
    var payload;
    try {
      if (!e || !e.postData || !e.postData.contents) {
        return responseErro("Payload vazio ou inválido.");
      }
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      return responseErro("JSON Inválido: " + err.message);
    }
    
    // 2. Roteador de Ações (Flexível para transição)
    var acao = payload.acao || payload.action;
    var data = payload.payload || payload.data || payload; 
    
    // 3. Roteador de Ações
    switch (acao) {
      case 'primeiroAcesso':
        return handlePrimeiroAcesso(data);
      
      case 'verificarPrimeiroAcesso':
        return handleVerificarPrimeiroAcesso(data);
      
      case 'atualizarCredenciais':
        return handleAtualizarCredenciais(data);

      case 'autenticarOperador':
        return handleAutenticarOperador(data);
      
      case 'lancarVenda':
        var resPlano = verificarPermissaoPlano(data.spreadsheetId, 'Vendas');
        if (resPlano.status === 'erro') return responseErro(resPlano.mensagem);
        return handleLancarVenda(data);

      case 'salvarRascunho':
        return handleSalvarRascunho(data);

      case 'estornarVenda':
        return handleEstornarVenda(data);
      
      case 'obterDadosRelatorios':
        var resPlanoRel = verificarPermissaoPlano(data.spreadsheetId, 'Relatórios');
        if (resPlanoRel.status === 'erro') return responseErro(resPlanoRel.mensagem);
        // handleObterDadosRelatorios será implementado na migração de Relatórios
        return responseErro("Ação 'obterDadosRelatorios' ainda não migrada para a Mestra.");
    }
    
  } catch (err) {
    return responseErro("Erro crítico no Servidor: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * AÇÃO: autenticarOperador
 * Valida o usuário e senha acessando a planilha do cliente.
 */
function handleAutenticarOperador(data) {
  var id = data.spreadsheetId;
  var user = data.nome; // Frontend envia 'nome' como o campo de login
  var pass = data.senha;

  if (!id) return responseErro("ID da planilha não fornecido.");

  // 1. TRAVA DE SEGURANÇA: Verificar status na Mestra antes de qualquer coisa
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("Clientes");
    if (masterSheet) {
      var mHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      var mColId = mHeaders.indexOf("Spreadsheet ID");
      var mColStatus = mHeaders.indexOf("Status");
      var mRow = findRowById(masterSheet, mColId, id);
      
      if (mRow > -1 && mColStatus > -1) {
        var status = masterSheet.getRange(mRow, mColStatus + 1).getValue();
        if (String(status).toLowerCase() === 'inativo') {
          return responseErro('Acesso bloqueado. Seu sistema está inativo. Por favor, fale com o gestor.');
        }
      }
    }
  } catch (e) {
    console.warn("Erro ao verificar trava de segurança: " + e.message);
  }

  // 2. Prosseguir com a autenticação na planilha do cliente
  try {
    var clientSS = SpreadsheetApp.openById(id);
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) return responseErro("Planilha do cliente não configurada corretamente.");

    // Layout Horizontal: Cabeçalhos na Linha 1, Dados na Linha 2
    var headers = configSheet.getRange(1, 1, 1, configSheet.getLastColumn()).getValues()[0];
    var dataRow = configSheet.getRange(2, 1, 1, configSheet.getLastColumn()).getValues()[0];
    
    var configs = {};
    for (var i = 0; i < headers.length; i++) {
        configs[headers[i]] = dataRow[i];
    }

    if (configs["Usuario"] === user && String(configs["Senha"]) === String(pass)) {
      return responseSucessoMsg("Login realizado com sucesso", {
        nome: user,
        nivel: configs["Nivel"] || "Admin",
        plano: configs["Plano"] || "Básico",
        empresa: configs["Empresa"] || "Minha Empresa",
        whatsapp: configs["Telefone"] || configs["WhatsApp"] || "",
        permissoes: configs["Permissoes"] ? JSON.parse(configs["Permissoes"]) : {}
      });
    } else {
      return responseErro("Usuário ou senha incorretos.");
    }
  } catch (e) {
    return responseErro("Erro ao autenticar: " + e.message);
  }
}

/**
 * Helper para resposta de sucesso com dados extras
 */
function responseSucessoMsg(msg, extraData) {
  var res = { status: "sucesso", mensagem: msg };
  if (extraData) {
    for (var key in extraData) res[key] = extraData[key];
  }
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * AÇÃO: verificarPrimeiroAcesso
 * Verifica se o cliente já está cadastrado na Mestra.
 */
function handleVerificarPrimeiroAcesso(data) {
  var spreadsheetId = data.spreadsheetId;
  if (!spreadsheetId) return responseJson({ primeiroAcesso: false });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  
  if (!sheet) return responseJson({ primeiroAcesso: true });

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colId = headers.indexOf("Spreadsheet ID");
  var colUser = headers.indexOf("Usuário Admin");
  
  var row = findRowById(sheet, colId, spreadsheetId);

  // 1. Se não encontrou a linha, é primeiro acesso (ID Novo)
  if (row === -1) return responseJson({ primeiroAcesso: true });

  // 2. Se encontrou a linha, verifica se o usuário admin está preenchido
  var usuarioExistente = sheet.getRange(row, colUser + 1).getValue();
  
  // Se estiver vazio, o cliente ainda não fez o setup inicial
  return responseJson({ primeiroAcesso: (!usuarioExistente || usuarioExistente === "") });
}

/**
 * Helper para retornar JSON puro (sem o wrapper de status sucesso/erro se necessário)
 */
function responseJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handlePrimeiroAcesso(data) {
  var id = data.spreadsheetId;
  if (!id) return responseErro("ID da Planilha não fornecido.");

  // TESTE DE ABERTURA (Integridade do Banco)
  var clientSS;
  try {
    clientSS = SpreadsheetApp.openById(id);
  } catch (e) {
    return responseErro("Falha crítica: Não foi possível abrir a planilha do cliente (ID: " + id + "). Operação cancelada para garantir integridade.");
  }

  try {
    // 1. Forçar valores padrão para a Mestra
    data.status = 'Ativo';
    data.plano = data.plano || 'Básico';

    // A. Salvar na Planilha Mestra (Upsert + Audit Trail + Dropdowns)
    upsertMasterClient(data, "Primeiro Acesso concluído");

    // B. Salvar na Planilha do Cliente (aba 'Configurações' - LAYOUT HORIZONTAL)
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var configSheet = clientSS.getSheetByName("Configurações") || clientSS.insertSheet("Configurações");
    
    // Mapeamento exato solicitado para o Banco de Dados do Cliente:
    var clientConfigsMap = {
      "Empresa": data.nome || "",
      "Nome": data.nomeCompleto || "",
      "Usuario": data.usuario || "",
      "Senha": data.senha || "",
      "Nível": "Admin",
      "Permissões": "{}",
      "Telefone": data.whatsapp || "",
      "Plano": data.plano,
      "Status": "Ativo",
      "UltimoAcesso": timestamp
    };
    
    updateHorizontalConfig(configSheet, clientConfigsMap);
    return responseSucesso("Primeiro acesso registrado com sucesso.");
    
  } catch (err) {
    return responseErro("Erro ao processar Primeiro Acesso: " + err.message);
  }
}

/**
 * AÇÃO: atualizarCredenciais
 * Tenta abrir o cliente PRIMEIRO. Se ok, atualiza cliente e Mestra.
 */
function handleAtualizarCredenciais(data) {
  var id = data.spreadsheetId;
  if (!id) return responseErro("ID da Planilha não fornecido.");

  // TESTE DE Abertura
  var clientSS;
  try {
    clientSS = SpreadsheetApp.openById(id);
  } catch (e) {
    return responseErro("Erro ao abrir planilha do cliente para atualizar credenciais: " + e.message);
  }

  try {
    // A. Atualizar na Planilha do Cliente (aba 'Configurações' - LAYOUT HORIZONTAL)
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) return responseErro("Aba 'Configurações' não encontrada no cliente.");
    
    var credentialsMap = {};
    var logParts = [];
    
    if (data.novoUsuario) {
        credentialsMap["Usuario"] = data.novoUsuario;
        logParts.push("Usuário alterado para " + data.novoUsuario);
    }
    
    if (data.whatsapp) {
        credentialsMap["Telefone"] = data.whatsapp;
        logParts.push("Telefone atualizado");
    }
    
    if (data.novaSenha && data.novaSenha.trim() !== "") {
        credentialsMap["Senha"] = data.novaSenha;
        logParts.push("Senha alterada");
    }
    
    if (Object.keys(credentialsMap).length === 0) {
        return responseSucesso("Nenhuma alteração detectada.");
    }
    
    updateHorizontalConfig(configSheet, credentialsMap);
    
    // B. Atualizar na Mestra (Upsert + Audit Trail)
    var upsertData = { spreadsheetId: id };
    if (data.novoUsuario) upsertData.usuario = data.novoUsuario;
    if (data.whatsapp) upsertData.whatsapp = data.whatsapp;
    
    upsertMasterClient(upsertData, "Atualização de Perfil: " + logParts.join(" | "));
    
    return responseSucesso("Credenciais atualizadas com sucesso.");
    
  } catch (err) {
    return responseErro("Erro ao atualizar credenciais: " + err.message);
  }
}

/**
 * CORE: O Coração do Registro (Upsert + Trilha de Auditoria)
 * Atualiza apenas as colunas fornecidas e concatena o histórico.
 */
function upsertMasterClient(data, actionDescription) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes") || setupMasterSheet(ss);
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  var colIdIndex = headers.indexOf("Spreadsheet ID");
  var colObsIndex = headers.indexOf("Observações");
  
  var id = data.spreadsheetId;
  var rowIndex = findRowById(sheet, colIdIndex, id);
  
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  var logEntry = "[" + timestamp + "] - " + actionDescription;
  
  // Mapeamento de chaves do payload para Headers da Mestra
  var mapping = {
    "nome": "Nome da Empresa / App",
    "usuario": "Usuário Admin",
    "whatsapp": "WhatsApp",
    "spreadsheetUrl": "Link da Planilha",
    "spreadsheetId": "Spreadsheet ID",
    "status": "Status",
    "plano": "Plano",
    "ativacao": "Ativação",
    "expiracao": "Expiração",
    "linkAcesso": "Link de Acesso"
  };

  var targetRow;
  if (rowIndex > -1) {
    targetRow = rowIndex;
    // ── ATUALIZAÇÃO INTELIGENTE (Somente colunas presentes no data)
    for (var key in data) {
      if (mapping[key]) {
        var hIdx = headers.indexOf(mapping[key]);
        if (hIdx > -1) {
          sheet.getRange(rowIndex, hIdx + 1).setValue(data[key]);
        }
      }
    }
    // Concatena na Trilha de Auditoria (Audit Trail)
    var currentObs = sheet.getRange(rowIndex, colObsIndex + 1).getValue();
    var newObs = (currentObs ? currentObs + "\n" : "") + logEntry;
    sheet.getRange(rowIndex, colObsIndex + 1).setValue(newObs);
    
  } else {
    // ── NOVO REGISTRO (AppendRow)
    var newRowData = new Array(headers.length).fill("");
    for (var key in data) {
      if (mapping[key]) {
        var hIdx = headers.indexOf(mapping[key]);
        if (hIdx > -1) newRowData[hIdx] = data[key];
      }
    }
    newRowData[colObsIndex] = logEntry; // Histórico inicial
    sheet.appendRow(newRowData);
    targetRow = sheet.getLastRow();
  }

  // 1. APLICAR DATA VALIDATION (Dropdowns)
  var colStatusIndex = headers.indexOf("Status");
  var colPlanoIndex = headers.indexOf("Plano");

  if (colStatusIndex > -1) {
    var ruleStatus = SpreadsheetApp.newDataValidation().requireValueInList(['Ativo', 'Inativo'], true).build();
    sheet.getRange(targetRow, colStatusIndex + 1).setDataValidation(ruleStatus);
  }
  if (colPlanoIndex > -1) {
    var rulePlano = SpreadsheetApp.newDataValidation().requireValueInList(['Básico', 'Pro', 'Premium'], true).build();
    sheet.getRange(targetRow, colPlanoIndex + 1).setDataValidation(rulePlano);
  }

  // 2. CONGELAR FÓRMULAS (Freeze to Values)
  // Converte toda a linha em valores estáticos para blindar os links gerados por fórmula
  var rangeToFreeze = sheet.getRange(targetRow, 1, 1, lastCol);
  rangeToFreeze.setValues(rangeToFreeze.getValues());
}

/**
 * UTILS
 */

function setupMasterSheet(ss) {
  var sheet = ss.insertSheet("Clientes");
  var headers = ["Nome da Empresa / App", "Usuário Admin", "WhatsApp", "Link da Planilha", "ScriptURL", "Spreadsheet ID", "Link de Acesso", "Status", "Plano", "Ativação", "Expiração", "Observações"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#dcfce7");
  sheet.setFrozenRows(1);
  return sheet;
}

function findRowById(sheet, colIndex, id) {
  if (colIndex === -1) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]).trim() === String(id).trim()) {
      return i + 1;
    }
  }
  return -1;
}

/**
 * UTILS: SINCRONIZAÇÃO HORIZONTAL (Configurações do Cliente)
 */
function updateHorizontalConfig(sheet, dataMap) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  
  for (var key in dataMap) {
    var colIdx = headers.indexOf(key);
    if (colIdx === -1) {
      // Cria a coluna se não existir
      colIdx = headers.length;
      headers.push(key);
      sheet.getRange(1, colIdx + 1).setValue(key).setFontWeight("bold");
    }
    // Grava o valor sempre na Linha 2
    sheet.getRange(2, colIdx + 1).setValue(dataMap[key]);
  }
}

function responseSucesso(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "sucesso", msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function responseErro(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "erro", mensagem: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * SEGURANÇA: Verificar Permissão por Plano
 */
function verificarPermissaoPlano(spreadsheetId, recurso) {
  if (!spreadsheetId) return { status: 'erro', mensagem: 'ID da planilha não informado.' };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Clientes");
    if (!sheet) return { status: 'sucesso' }; // Se não tem aba, deixa passar (primeiro acesso)
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colId = headers.indexOf("Spreadsheet ID");
    var colPlano = headers.indexOf("Plano");
    
    var row = findRowById(sheet, colId, spreadsheetId);
    if (row === -1) return { status: 'erro', mensagem: 'Cliente não localizado na Mestra.' };
    
    var plano = sheet.getRange(row, colPlano + 1).getValue();
    var p = String(plano).toLowerCase();
    
    if (p === 'básico' || p === 'basico') {
      if (recurso === 'Relatórios') {
        return { status: 'erro', mensagem: 'Relatórios Estratégicos não disponíveis no Plano Básico. Faça o upgrade!' };
      }
      // Outras restrições podem ser adicionadas aqui
    }
    
    return { status: 'sucesso' };
  } catch (e) {
    return { status: 'erro', mensagem: 'Erro ao validar plano: ' + e.message };
  }
}

/**
 * MOTOR DE VENDAS (MIGRADO DO CÓDIGO LEGADO)
 */

function handleLancarVenda(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheetVendas = ss.getSheetByName('Vendas');
    var sheetProdutos = ss.getSheetByName('Produtos');
    var sheetFin = ss.getSheetByName('Financeiro');
    
    if (!sheetVendas) return responseErro("Aba 'Vendas' não encontrada no cliente.");

    // 1. Baixa de estoque
    var erroEstoque = baixarEstoqueItens(sheetProdutos, data.itensList);
    if (erroEstoque) return responseErro(erroEstoque);

    // 2. Gravar Venda
    var novoId = proximoIdVendas(sheetVendas);
    var vencimento = data.vencimento || data.data;
    var statusFin = data.statusFinanceiro || 'Pendente';

    sheetVendas.appendRow([
      novoId, data.data,
      data.cliente || 'Consumidor Interno',
      data.itens, data.quantidadeVendida,
      data.subtotal, data.descontoPercentual, data.descontoReal,
      data.totalComDesconto,
      data.formaPagamento || '',
      data.usuario || '',
      'Concluda',
      vencimento,
      JSON.stringify(data.itensList || [])
    ]);

    // 3. Financeiro
    if (sheetFin) {
      var lastRowFin = sheetFin.getLastRow();
      var nextIdFin = lastRowFin > 1 ? (parseInt(sheetFin.getRange(lastRowFin, 1).getValue()) || 0) + 1 : 1;
      sheetFin.appendRow([
        nextIdFin,
        'Venda #' + novoId + ' - ' + (data.cliente || 'Consumidor'),
        data.totalComDesconto, 'Receber',
        vencimento, statusFin, 'Venda', novoId
      ]);
    }

    return responseSucesso("✅ Venda #" + novoId + " concluída com sucesso!");
  } catch (e) {
    return responseErro("Falha ao lançar venda: " + e.message);
  }
}

function handleSalvarRascunho(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName('Vendas');
    if (!sheet) return responseErro("Aba 'Vendas' não encontrada.");

    var novoId = proximoIdVendas(sheet);
    sheet.appendRow([
      novoId, data.data,
      data.cliente || 'Consumidor Interno',
      data.itens, data.quantidadeVendida,
      data.subtotal, data.descontoPercentual, data.descontoReal,
      data.totalComDesconto,
      data.formaPagamento || '',
      data.usuario || '',
      'Pendente',
      '', 
      JSON.stringify(data.itensList || [])
    ]);
    return responseSucesso("💾 Rascunho #" + novoId + " salvo com sucesso!");
  } catch (e) {
    return responseErro("Erro ao salvar rascunho: " + e.message);
  }
}

function handleEstornarVenda(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheetVendas = ss.getSheetByName('Vendas');
    var sheetProdutos = ss.getSheetByName('Produtos');
    var sheetFin = ss.getSheetByName('Financeiro');
    
    if (!sheetVendas) return responseErro("Aba 'Vendas' não encontrada.");

    var todosDados = sheetVendas.getDataRange().getValues();
    var linhaVenda = -1;
    for (var i = 1; i < todosDados.length; i++) {
        if (String(todosDados[i][0]) === String(data.id)) { linhaVenda = i + 1; break; }
    }
    if (linhaVenda === -1) return responseErro("Venda #" + data.id + " não encontrada.");

    // 1. Devolve estoque
    var itensList = [];
    try { itensList = JSON.parse(todosDados[linhaVenda - 1][13] || '[]'); } catch(e) {}
    devolverEstoqueItens(sheetProdutos, itensList);

    // 2. Status Venda
    sheetVendas.getRange(linhaVenda, 12).setValue('Estornada');

    // 3. Cancela Financeiro
    if (sheetFin && sheetFin.getLastRow() > 1) {
      var dadosFin = sheetFin.getDataRange().getValues();
      for (var i = 1; i < dadosFin.length; i++) {
        if (String(dadosFin[i][7]) === String(data.id) && dadosFin[i][3] === 'Receber') {
          sheetFin.getRange(i + 1, 6).setValue('Estornado');
          break;
        }
      }
    }
    return responseSucesso("↩️ Venda #" + data.id + " estornada. Estoque devolvido.");
  } catch (e) {
    return responseErro("Erro ao estornar venda: " + e.message);
  }
}

/**
 * HELPERS DE NEGÓCIO
 */

function baixarEstoqueItens(sheetProdutos, itensList) {
  if (!sheetProdutos || sheetProdutos.getLastRow() < 2 || !itensList || itensList.length === 0) return null;
  var dadosProd = sheetProdutos.getDataRange().getValues();
  var colNome = dadosProd[0].indexOf('Nome');
  var colQtd  = dadosProd[0].indexOf('Quantidade');
  if (colNome === -1 || colQtd === -1) return 'Colunas Nome/Quantidade não encontradas em Produtos.';

  // Valida
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
    if (!found) return 'Produto "' + nm + '" não encontrado.';
  }

  // Subtrai
  for (var k = 0; k < itensList.length; k++) {
    var nm = String(itensList[k].nome).trim();
    var qt = parseFloat(itensList[k].quantidade) || 0;
    for (var i = 1; i < dadosProd.length; i++) {
      if (String(dadosProd[i][colNome]).trim() === nm) {
        var novo = (parseFloat(dadosProd[i][colQtd]) || 0) - qt;
        sheetProdutos.getRange(i + 1, colQtd + 1).setValue(novo);
        break;
      }
    }
  }
  return null;
}

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
        var novo = (parseFloat(dadosProd[i][colQtd]) || 0) + qt;
        sheetProdutos.getRange(i + 1, colQtd + 1).setValue(novo);
        break;
      }
    }
  }
}

function proximoIdVendas(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return 1;
  var val = sheet.getRange(last, 1).getValue();
  return (parseInt(val) || 0) + 1;
}
