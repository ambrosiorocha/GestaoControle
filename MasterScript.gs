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
      
      default:
        return responseErro("Ação desconhecida: " + acao);
    }
    
  } catch (err) {
    return responseErro("Erro crítico no Servidor: " + err.message);
  } finally {
    lock.releaseLock();
  }
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

/**
 * AÇÃO: primeiroAcesso
 * Tenta abrir o cliente PRIMEIRO. Se falhar, aborta.
 * Se ok, salva na Mestra e na planilha do cliente (aba 'Configurações').
 */
function handlePrimeiroAcesso(data) {
  var spreadsheetId = data.spreadsheetId;
  if (!spreadsheetId) return responseErro("ID da Planilha não fornecido.");

  // TESTE DE Abertura (Integridade do Banco)
  var clientSS;
  try {
    clientSS = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return responseErro("Falha crítica: Não foi possível abrir a planilha do cliente (ID: " + spreadsheetId + "). Operação cancelada para garantir integridade.");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes") || setupMasterSheet(ss);
  
  try {
    // A. Salvar na Planilha Mestra (Clientes)
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colId = headers.indexOf("Spreadsheet ID");
    
    var rowToUpdate = findRowById(sheet, colId, spreadsheetId);
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    var newRowData = [
      data.nome || "Novo Cliente",
      data.usuario || "admin",
      data.whatsapp || "",
      data.spreadsheetUrl || "",
      "", // ScriptURL (obsoleto na nova arquitetura)
      spreadsheetId,
      "", // Link de Acesso
      "Ativo",
      data.plano || "Básico",
      timestamp,
      data.expiracao || "",
      "Primeiro acesso em " + timestamp
    ];
    
    if (rowToUpdate > -1) {
      sheet.getRange(rowToUpdate, 1, 1, newRowData.length).setValues([newRowData]);
    } else {
      sheet.appendRow(newRowData);
    }
    
    // B. Salvar na Planilha do Cliente (aba 'Configurações')
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) {
      configSheet = clientSS.insertSheet("Configurações");
      configSheet.getRange("A1:B1").setValues([["Chave", "Valor"]]).setFontWeight("bold");
    }
    
    // Mapeamento de dados para o cliente
    var clientConfigs = [
      ["Empresa", data.nome || ""],
      ["Usuario", data.usuario || ""],
      ["Senha", data.senha || ""],
      ["Plano", data.plano || "Básico"],
      ["Status", "Ativo"],
      ["UltimoAcesso", timestamp]
    ];
    
    // Atualiza ou insere chaves
    updateKeyValueSheet(configSheet, clientConfigs);
    
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
  var spreadsheetId = data.spreadsheetId;
  if (!spreadsheetId) return responseErro("ID da Planilha não fornecido.");

  // TESTE DE Abertura
  var clientSS;
  try {
    clientSS = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return responseErro("Erro ao abrir planilha do cliente para atualizar credenciais: " + e.message);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  
  try {
    // A. Atualizar na Planilha do Cliente (aba 'Configurações')
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) return responseErro("Aba 'Configurações' não encontrada no cliente.");
    
    var credentials = [
      ["Usuario", data.novoUsuario || data.usuario],
      ["Senha", data.novaSenha || data.senha]
    ];
    
    updateKeyValueSheet(configSheet, credentials);
    
    // B. Atualizar apenas o Usuário na Mestra
    if (sheet) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var colId = headers.indexOf("Spreadsheet ID");
      var colUser = headers.indexOf("Usuário Admin");
      
      var rowToUpdate = findRowById(sheet, colId, spreadsheetId);
      if (rowToUpdate > -1 && colUser > -1) {
        sheet.getRange(rowToUpdate, colUser + 1).setValue(data.novoUsuario || data.usuario);
      }
    }
    
    return responseSucesso("Credenciais atualizadas com sucesso.");
    
  } catch (err) {
    return responseErro("Erro ao atualizar credenciais: " + err.message);
  }
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

function updateKeyValueSheet(sheet, kvArray) {
  var data = sheet.getDataRange().getValues();
  kvArray.forEach(function(pair) {
    var key = pair[0];
    var val = pair[1];
    var found = false;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(val);
        found = true;
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow([key, val]);
    }
  });
}

function responseSucesso(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "sucesso", msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function responseErro(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "erro", msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
