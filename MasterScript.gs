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
 * AÇÃO: autenticarOperador
 * Valida o usuário e senha acessando a planilha do cliente.
 */
function handleAutenticarOperador(data) {
  var id = data.spreadsheetId;
  var user = data.nome; // Frontend envia 'nome' como o campo de login
  var pass = data.senha;

  if (!id) return responseErro("ID da planilha não fornecido.");

  try {
    var clientSS = SpreadsheetApp.openById(id);
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) return responseErro("Planilha do cliente não configurada corretamente.");

    var values = configSheet.getDataRange().getValues();
    var configs = {};
    for (var i = 1; i < values.length; i++) {
      configs[values[i][0]] = values[i][1];
    }

    if (configs["Usuario"] === user && String(configs["Senha"]) === String(pass)) {
      return responseSucessoMsg("Login realizado com sucesso", {
        nome: user,
        nivel: configs["Nivel"] || "Admin",
        plano: configs["Plano"] || "Básico",
        empresa: configs["Empresa"] || "Minha Empresa",
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

/**
 * AÇÃO: primeiroAcesso
 * Tenta abrir o cliente PRIMEIRO. Se falhar, aborta.
 * Se ok, salva na Mestra e na planilha do cliente (aba 'Configurações').
 */
/**
 * AÇÃO: primeiroAcesso
 * Tenta abrir o cliente PRIMEIRO. Se falhar, aborta.
 * Se ok, salva na Mestra e na planilha do cliente (aba 'Configurações').
 */
function handlePrimeiroAcesso(data) {
  var id = data.spreadsheetId;
  if (!id) return responseErro("ID da Planilha não fornecido.");

  // TESTE DE Abertura (Integridade do Banco)
  var clientSS;
  try {
    clientSS = SpreadsheetApp.openById(id);
  } catch (e) {
    return responseErro("Falha crítica: Não foi possível abrir a planilha do cliente (ID: " + id + "). Operação cancelada para garantir integridade.");
  }

  try {
    // A. Salvar na Planilha Mestra (Upsert + Audit Trail)
    upsertMasterClient(data, "Primeiro Acesso concluído");

    // B. Salvar na Planilha do Cliente (aba 'Configurações')
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) {
      configSheet = clientSS.insertSheet("Configurações");
      configSheet.getRange("A1:B1").setValues([["Chave", "Valor"]]).setFontWeight("bold");
    }
    
    var clientConfigs = [
      ["Empresa", data.nome || ""],
      ["Usuario", data.usuario || ""],
      ["Senha", data.senha || ""],
      ["Plano", data.plano || "Básico"],
      ["Status", "Ativo"],
      ["UltimoAcesso", timestamp]
    ];
    
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
    // A. Atualizar na Planilha do Cliente (aba 'Configurações')
    var configSheet = clientSS.getSheetByName("Configurações");
    if (!configSheet) return responseErro("Aba 'Configurações' não encontrada no cliente.");
    
    var credentials = [
      ["Usuario", data.novoUsuario || data.usuario],
      ["Senha", data.novaSenha || data.senha]
    ];
    updateKeyValueSheet(configSheet, credentials);
    
    // B. Atualizar na Mestra (Upsert + Audit Trail)
    var logMsg = "Credenciais atualizadas: ";
    if (data.novoUsuario) logMsg += "Usuário alterado (" + data.novoUsuario + ") ";
    if (data.novaSenha) logMsg += "Senha alterada";
    
    upsertMasterClient({
      spreadsheetId: id,
      usuario: data.novoUsuario || data.usuario
    }, logMsg);
    
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

  if (rowIndex > -1) {
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
