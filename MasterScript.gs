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
    
    // Garantir que o spreadsheetId esteja disponível no objeto data para os handlers
    if (!data.spreadsheetId && payload.spreadsheetId) {
      data.spreadsheetId = payload.spreadsheetId;
    }
    
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
      
      case 'registrarMestra':
        return handleRegistrarMestra(data);
      
      case 'lancarVenda':
        var resPlano = verificarPermissaoPlano(data.spreadsheetId, 'Vendas');
        if (resPlano.status === 'erro') return responseErro(resPlano.mensagem);
        return handleLancarVenda(data);

      case 'salvarRascunho':
        return handleSalvarRascunho(data);

      case 'finalizarPendente':
        return handleFinalizarPendente(data);

      case 'estornarVenda':
        return handleEstornarVenda(data);
      
      case 'listarProdutos':
      case 'obterProdutos':
        return handleObterProdutos(data);
      
      case 'obterProdutoPorId':
        return handleObterProdutoPorId(data);

      case 'obterOperadores':
        return handleObterOperadores(data);

      case 'obterClientes':
        return handleObterClientes(data);

      case 'obterVendas':
        return handleObterVendas(data);

      case 'obterLucro':
      case 'obterDadosRelatorios':
        var resPlanoRel = verificarPermissaoPlano(data.spreadsheetId, 'Relatórios');
        if (resPlanoRel.status === 'erro') return responseErro(resPlanoRel.mensagem);
        return handleObterDadosRelatorios(data);

      case 'obterRascunhos':
        return handleObterRascunhos(data);

      case 'obterFinanceiro':
        return handleObterFinanceiro(data);

      case 'excluirRascunho':
        return handleExcluirRascunho(data);
      
      case 'baixarLancamento':
        return handleBaixarLancamento(data);
      
      case 'salvarFinanceiro':
        return handleSalvarFinanceiro(data);
      
      case 'excluirFinanceiro':
        return handleExcluirFinanceiro(data);

      case 'salvarCliente':
        return handleSalvarCliente(data);
      
      case 'excluirCliente':
        return handleExcluirCliente(data);
      
      case 'salvarFornecedor':
        return handleSalvarFornecedor(data);
      
      case 'excluirFornecedor':
        return handleExcluirFornecedor(data);
      
      case 'obterFornecedores':
        return handleObterFornecedores(data);
      
      case 'salvarProduto':
        return handleSalvarProduto(data);
      
      case 'excluirProduto':
        return handleExcluirProduto(data);
      
      case 'salvarOperador':
        return handleSalvarOperador(data);
      
      case 'excluirOperador':
        return handleExcluirOperador(data);

      case 'obterProdutosUnicos':
        return handleObterProdutosUnicos(data);

      case 'obterDashboard':
        return handleObterDashboard(data);
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
        caixas: configs["Caixas"] || "Dinheiro",
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
function handleObterFinanceiro(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Financeiro");
    if (!sheet || sheet.getLastRow() < 2) {
      return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: [], rows: [] } });
    }
    var values = sheet.getDataRange().getValues();
    var headers = values.shift();
    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headers, rows: values } });
  } catch (e) { return responseErro(e.message); }
}

function handleObterRascunhos(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Vendas");
    if (!sheet || sheet.getLastRow() < 2) return responseSucessoMsg("Sucesso", { dados: [] });
    
    var values = sheet.getDataRange().getValues();
    var headers = values.shift();
    var rascunhos = values.filter(function(row) {
      return row[11] === 'Pendente'; 
    });
    
    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headers, rows: rascunhos } });
  } catch (e) { return responseErro(e.message); }
}

function handleExcluirRascunho(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Vendas");
    if (!sheet) return responseErro("Aba Vendas não encontrada.");
    
    // Blindagem de ID: Tenta capturar de várias chaves comuns no payload
    var id = data.id || data.idRascunho || data.idVenda || data.numero || (data.data ? (data.data.id || data.data.idRascunho) : null);
    if (!id) return responseErro("ID da venda/rascunho não fornecido no payload.");

    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        if (values[i][11] !== 'Pendente') {
          return responseErro("Apenas rascunhos pendentes podem ser excluídos diretamente.");
        }
        sheet.deleteRow(i + 1);
        return responseSucesso("🗑️ Rascunho #" + id + " excluído.");
      }
    }
    return responseErro("Rascunho não encontrado.");
  } catch (e) { return responseErro(e.message); }
}

function handleObterDadosRelatorios(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheetVendas = ss.getSheetByName('Vendas');
    var sheetProds  = ss.getSheetByName('Produtos');
    if (!sheetVendas || sheetVendas.getLastRow() < 2) {
      return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: [], rows: [] } });
    }

    var vendasRaw = sheetVendas.getDataRange().getValues();
    var headersVendas = vendasRaw.shift();
    var prodsRaw = sheetProds ? sheetProds.getDataRange().getValues() : [];
    
    // Mapeamento de custos para cálculo de lucro
    var custosMap = {};
    if (prodsRaw.length > 1) {
      var hP = prodsRaw.shift();
      var idxPNome = hP.indexOf('Nome');
      var idxPCusto = hP.indexOf('Preço_de_custo');
      if (idxPCusto === -1) idxPCusto = hP.indexOf('PrecoCusto');
      
      prodsRaw.forEach(function(p) {
        if (idxPNome > -1 && idxPCusto > -1) {
          custosMap[String(p[idxPNome]).trim()] = parseFloat(p[idxPCusto]) || 0;
        }
      });
    }

    var rowsCompleto = vendasRaw.map(function(v) {
      var itensJSON = [];
      try { itensJSON = JSON.parse(v[13] || '[]'); } catch(e) {}
      
      var custoTotalVenda = 0;
      itensJSON.forEach(function(it) {
        var unitCusto = custosMap[String(it.nome).trim()] || 0;
        custoTotalVenda += unitCusto * (parseFloat(it.quantidade) || 0);
      });

      var r = v.slice(); // Cópia
      r.push(custoTotalVenda); // Nova coluna: Custo Total
      return r;
    });

    var headersFinal = headersVendas.slice();
    headersFinal.push('Custo Total');

    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headersFinal, rows: rowsCompleto } });
  } catch (e) { return responseErro(e.message); }
}

function handleBaixarLancamento(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName('Financeiro');
    if (!sheet || sheet.getLastRow() < 2) return responseErro("Planilha Financeiro não encontrada ou vazia.");
    
    var idTarget = data.id || (data.data ? data.data.id : null);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colId = headers.indexOf('id');
    var colStatus = headers.indexOf('status');
    if (colId === -1 || colStatus === -1) return responseErro("Colunas id/status não encontradas.");
    
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][colId]) === String(idTarget)) {
        sheet.getRange(i + 2, colStatus + 1).setValue('Pago');
        return responseSucesso("Lançamento #" + idTarget + " baixado com sucesso!");
      }
    }
    return responseErro("Lançamento não encontrado.");
  } catch (e) { return responseErro(e.message); }
}

function handleSalvarFinanceiro(data) {
  return handleSalvarDadosGeral("Financeiro", data);
}

function handleExcluirFinanceiro(data) {
  var id = data.id || (data.data ? data.data.id : null);
  return handleExcluirDadosGeral("Financeiro", data.spreadsheetId, id);
}

function handleSalvarCliente(data) {
  return handleSalvarDadosGeral("Clientes", data);
}

function handleExcluirCliente(data) {
  var id = data.id || (data.data ? data.data.id : null);
  return handleExcluirDadosGeral("Clientes", data.spreadsheetId, id);
}

function handleSalvarFornecedor(data) {
  return handleSalvarDadosGeral("Fornecedores", data);
}

function handleExcluirFornecedor(data) {
  var id = data.id || (data.data ? data.data.id : null);
  return handleExcluirDadosGeral("Fornecedores", data.spreadsheetId, id);
}

function handleObterFornecedores(data) {
  return handleObterDadosGeral(data, "Fornecedores");
}

function handleSalvarOperador(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName('Configurações');
    if (!sheet) return responseErro("Aba Configurações não encontrada.");
    
    var vData = data.data || data;
    var nome = String(vData.nome).trim();
    var nivel = String(vData.nivel || 'Operador').trim();
    var senha = String(vData.senha || '1234');
    var plano = String(vData.plano || 'Pro').trim();
    var permissoes = JSON.stringify(vData.permissoes || { relatorios: true, fiado: true, visaoDono: false });
    
    if (sheet.getLastRow() > 1) {
      var existentes = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r){ return r[0]; });
      if (existentes.indexOf(nome) > -1) return responseErro("Operador \"" + nome + "\" já existe.");
    }
    sheet.appendRow([nome, nivel, senha, plano, permissoes]);
    return responseSucesso("Operador \"" + nome + "\" adicionado!");
  } catch (e) { return responseErro(e.message); }
}

function handleExcluirOperador(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName('Configurações');
    if (!sheet || sheet.getLastRow() < 2) return responseErro("Nenhum operador cadastrado.");
    
    var vData = data.data || data;
    var nome = String(vData.nome || vData.id).trim();
    var dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === nome) {
        sheet.deleteRow(i + 2);
        return responseSucesso("Operador removido.");
      }
    }
    return responseErro("Operador não encontrado.");
  } catch (e) { return responseErro(e.message); }
}

function handleSalvarProduto(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Produtos");
    if (!sheet) return responseErro("Aba Produtos não encontrada.");
    
    var vData = data.data || data;
    var idProduto = vData.idProduto || vData.id;
    
    var valoresProduto = [
      vData.nome,
      vData.unidadeVenda,
      parseFloat(vData.precoCusto) || 0,
      parseFloat(vData.margemPct)  || 0,
      parseFloat(vData.margemRS)   || 0,
      parseFloat(vData.preco)      || 0,
      parseFloat(vData.quantidade) || 0,
      vData.descricao || ''
    ];

    if (idProduto) {
      var dadosSheet = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      var linha = dadosSheet.findIndex(function(row) { return row[0] == idProduto; });
      if (linha > -1) {
        sheet.getRange(linha + 2, 2, 1, valoresProduto.length).setValues([valoresProduto]);
        return responseSucesso("Produto \"" + vData.nome + "\" atualizado!");
      }
      return responseErro("Produto não encontrado para atualização.");
    } else {
      var ultimaLinha = sheet.getLastRow();
      var novoId = (ultimaLinha > 1) ? (parseInt(sheet.getRange(ultimaLinha, 1).getValue()) || 0) + 1 : 1;
      sheet.appendRow([novoId, ...valoresProduto]);
      return responseSucesso("Produto \"" + vData.nome + "\" cadastrado!");
    }
  } catch (e) { return responseErro(e.message); }
}

function handleExcluirProduto(data) {
  var id = data.id || (data.data ? data.data.id : null);
  return handleExcluirDadosGeral("Produtos", data.spreadsheetId, id);
}

/**
 * HELPERS GENÉRICOS DE CRUD
 */
function handleObterDadosGeral(data, nomePlanilha) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName(nomePlanilha);
    if (!sheet || sheet.getLastRow() < 2) {
      return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: [], rows: [] } });
    }
    var values = sheet.getDataRange().getValues();
    var headers = values.shift();
    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headers, rows: values } });
  } catch (e) { return responseErro(e.message); }
}

function handleSalvarDadosGeral(nomePlanilha, data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName(nomePlanilha);
    if (!sheet) return responseErro("Aba " + nomePlanilha + " não encontrada.");
    
    var vData = data.data || data;
    var id = vData.id;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowValues = headers.map(header => {
      var val = vData[header];
      return val !== undefined ? val : "";
    });

    if (id) {
      var dataIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      var rowIndex = dataIds.findIndex(function(row) { return row[0] == id; });
      if (rowIndex > -1) {
        sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([rowValues]);
        return responseSucesso("Registro atualizado com sucesso!");
      }
      return responseErro("Registro não encontrado para atualização.");
    } else {
      var lastRow = sheet.getLastRow();
      var nextId = lastRow > 1 ? (parseInt(sheet.getRange(lastRow, 1).getValue()) || 0) + 1 : 1;
      rowValues[0] = nextId; 
      sheet.appendRow(rowValues);
      return responseSucesso("Registro cadastrado com sucesso!");
    }
  } catch (e) { return responseErro(e.message); }
}

function handleExcluirDadosGeral(nomePlanilha, spreadsheetId, id) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(nomePlanilha);
    if (!sheet || sheet.getLastRow() < 2) return responseErro("Aba " + nomePlanilha + " não encontrada ou vazia.");
    
    var dadosIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var linha = dadosIds.findIndex(function(row) { return row[0] == id; });
    if (linha > -1) {
      sheet.deleteRow(linha + 2);
      return responseSucesso("Registro excluído com sucesso!");
    }
    return responseErro("Registro não encontrado para exclusão.");
  } catch (e) { return responseErro(e.message); }
}

function handleObterProdutoPorId(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Produtos");
    if (!sheet) return responseErro("Aba Produtos não encontrada.");
    
    var id = data.id || (data.data ? data.data.id : null);
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        var obj = {};
        headers.forEach((h, idx) => obj[h] = values[i][idx]);
        return responseSucessoMsg("Sucesso", { dados: obj });
      }
    }
    return responseErro("Produto não encontrado.");
  } catch (e) { return responseErro(e.message); }
}

function responseSucessoMsg(msg, extra) {
  var res = { status: "sucesso", mensagem: msg };
  if (extra) {
    for (var k in extra) res[k] = extra[k];
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

function handleRegistrarMestra(data) {
  try {
    // 1. Atualizar a Planilha Mestra (Auditoria e Controle Central)
    var resultMestra = upsertMasterClient(data, "Atualização via Login/Acesso");

    // 2. Sincronizar com a Planilha do Cliente (Informação Local)
    var id = data.spreadsheetId;
    if (id) {
       try {
         var clientSS = SpreadsheetApp.openById(id);
         var configSheet = clientSS.getSheetByName("Configurações");
         if (configSheet) {
           var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
           // Agora sincronizamos também o Plano que veio da Mestra
           updateHorizontalConfig(configSheet, { 
             "UltimoAcesso": timestamp,
             "Plano": resultMestra.plano,
             "Caixas": data.caixas || undefined
           });
         }
       } catch (e) {
         console.warn("Erro ao atualizar login na planilha do cliente: " + e.message);
       }
    }

    // Retornamos o plano e status reais para que o frontend possa se auto-atualizar (ex: upgrade para Pro)
    return responseSucessoMsg("Registro atualizado na Mestra.", {
      plano: resultMestra.plano,
      status: resultMestra.status,
      caixas: data.caixas || undefined
    });
  } catch (e) { return responseErro(e.message); }
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
      "Caixas": "Dinheiro",
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

    if (data.caixas) {
        credentialsMap["Caixas"] = data.caixas;
        logParts.push("Caixas personalizadas: " + data.caixas);
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
    "empresa": "Nome da Empresa / App",
    "nome": "Nome da Empresa / App", // Retrocompatibilidade
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
    var dateStr = timestamp.split(' ')[0];
    var shouldAddLog = true;

    // Se for um log de login, verifica se já existe um para o dia de hoje
    if (actionDescription === "Atualização via Login/Acesso" && currentObs) {
      if (currentObs.indexOf("[" + dateStr) !== -1 && currentObs.indexOf("Atualização via Login/Acesso") !== -1) {
        shouldAddLog = false;
      }
    }

    if (shouldAddLog) {
      var newObs = (currentObs ? currentObs + "\n" : "") + logEntry;
      sheet.getRange(rowIndex, colObsIndex + 1).setValue(newObs);
    }
    
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

  // Retorna os dados atuais para sincronia do frontend
  return {
    plano: (colPlanoIndex > -1) ? sheet.getRange(targetRow, colPlanoIndex + 1).getValue() : (data.plano || 'Básico'),
    status: (colStatusIndex > -1) ? sheet.getRange(targetRow, colStatusIndex + 1).getValue() : 'Ativo'
  };
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
  return ContentService.createTextOutput(JSON.stringify({ status: "sucesso", mensagem: msg }))
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

    var vData = data.data || data;
    var idRascunho = vData.id || vData.idRascunho || data.id || data.idRascunho;
    var todosDados = sheet.getDataRange().getValues();
    var linhaVenda = -1;

    // Se tiver ID, tenta achar para ATUALIZAR
    if (idRascunho) {
      for (var i = 1; i < todosDados.length; i++) {
        if (String(todosDados[i][0]) === String(idRascunho)) { linhaVenda = i + 1; break; }
      }
    }

    var finalId = idRascunho && linhaVenda > -1 ? idRascunho : proximoIdVendas(sheet);
    var rowData = [
      finalId, vData.data,
      vData.cliente || 'Consumidor Interno',
      vData.itens, vData.quantidadeVendida,
      vData.subtotal, vData.descontoPercentual, vData.descontoReal,
      vData.totalComDesconto,
      vData.formaPagamento || '-',
      vData.usuario || '',
      'Pendente',
      '', 
      JSON.stringify(vData.itensList || [])
    ];

    if (linhaVenda > -1) {
      sheet.getRange(linhaVenda, 1, 1, rowData.length).setValues([rowData]);
      return responseSucesso("💾 Rascunho #" + finalId + " atualizado!");
    } else {
      sheet.appendRow(rowData);
      return responseSucesso("💾 Rascunho #" + finalId + " salvo!");
    }
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
    var idVenda = data.id || data.idVenda || data.idRascunho || data.numero || (data.data ? (data.data.id || data.data.idVenda) : null);
    if (!idVenda) return responseErro("ID da venda não fornecido no payload.");

    var linhaVenda = -1;
    for (var i = 1; i < todosDados.length; i++) {
        if (String(todosDados[i][0]) === String(idVenda)) { linhaVenda = i + 1; break; }
    }
    if (linhaVenda === -1) return responseErro("Venda #" + idVenda + " não encontrada.");

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
        if (String(dadosFin[i][7]) === String(idVenda) && dadosFin[i][3] === 'Receber') {
          sheetFin.getRange(i + 1, 6).setValue('Estornado');
          break;
        }
      }
    }
    return responseSucesso("↩️ Venda #" + idVenda + " estornada. Estoque devolvido.");
  } catch (e) {
    return responseErro("Erro ao estornar venda: " + e.message);
  }
}

function handleFinalizarPendente(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheetVendas = ss.getSheetByName('Vendas');
    var sheetProdutos = ss.getSheetByName('Produtos');
    var sheetFin = ss.getSheetByName('Financeiro');
    if (!sheetVendas) return responseErro("Aba 'Vendas' não encontrada.");

    var vData = data.data || data;
    // Blindagem de ID: Captura de múltiplas chaves
    var idVenda = vData.id || vData.idVenda || vData.idRascunho || data.id || data.idVenda || data.idRascunho || data.numero;
    
    var todosDados = sheetVendas.getDataRange().getValues();
    var linhaVenda = -1;
    for (var i = 1; i < todosDados.length; i++) {
      if (String(todosDados[i][0]) === String(idVenda)) { linhaVenda = i + 1; break; }
    }
    if (linhaVenda === -1) return responseErro("Venda/Rascunho #" + idVenda + " não encontrada para finalização.");

    var itensList = [];
    try { itensList = JSON.parse(todosDados[linhaVenda - 1][13] || '[]'); } catch(e) {}
    if (vData.itensList && vData.itensList.length > 0) itensList = vData.itensList;

    var erro = baixarEstoqueItens(sheetProdutos, itensList);
    if (erro) return responseErro(erro);

    var vencimento = vData.vencimento || vData.data || todosDados[linhaVenda - 1][1];
    var statusFin  = vData.statusFinanceiro || 'Pendente';
    var total      = parseFloat(todosDados[linhaVenda - 1][8]) || 0;
    var cliente    = todosDados[linhaVenda - 1][2] || 'Consumidor';
    var pgto       = vData.formaPagamento || todosDados[linhaVenda - 1][9] || '';

    sheetVendas.getRange(linhaVenda, 10).setValue(pgto);
    sheetVendas.getRange(linhaVenda, 12).setValue('Concluída');
    sheetVendas.getRange(linhaVenda, 13).setValue(vencimento);

    if (sheetFin) {
      var lastRowFin = sheetFin.getLastRow();
      var nextIdFin = lastRowFin > 1 ? (parseInt(sheetFin.getRange(lastRowFin, 1).getValue()) || 0) + 1 : 1;
      sheetFin.appendRow([
        nextIdFin,
        'Venda #' + idVenda + ' - ' + cliente,
        total, 'Receber', vencimento, statusFin, 'Venda', idVenda
      ]);
    }
    return responseSucesso("✅ Venda #" + idVenda + " finalizada!");
  } catch (e) {
    return responseErro("Erro ao finalizar pendente: " + e.message);
  }
}

/**
 * FUNÇÕES DE LEITURA (MIGRADO DO CÓDIGO LEGADO)
 */

function handleObterProdutos(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Produtos");
    if (!sheet || sheet.getLastRow() < 2) {
      return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: [], rows: [] } });
    }
    var values = sheet.getDataRange().getValues();
    var headers = values.shift();
    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headers, rows: values } });
  } catch (e) { return responseErro(e.message); }
}

function handleObterProdutosUnicos(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Produtos");
    if (!sheet || sheet.getLastRow() < 2) return responseSucessoMsg("Sucesso", { dados: [] });
    
    // Supondo que a coluna 2 seja o Nome do Produto (conforme Code.gs)
    var dados = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    var unicos = Array.from(new Set(dados.flat().map(v => String(v).trim()))).filter(v => v !== "");
    return responseSucessoMsg("Sucesso", { dados: unicos.sort() });
  } catch (e) { return responseErro("Erro ao obter produtos únicos: " + e.message); }
}

function handleObterOperadores(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName('Configurações');
    if (!sheet || sheet.getLastRow() < 2) return responseSucessoMsg("Sucesso", { dados: [] });
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var colNome = headers.indexOf('Nome');
    if (colNome === -1) colNome = headers.indexOf('Usuario'); 
    
    var nomes = [];
    for (var i = 1; i < values.length; i++) {
      if (values[i][colNome]) nomes.push(values[i][colNome]);
    }
    return responseSucessoMsg("Sucesso", { dados: nomes });
  } catch (e) { return responseErro(e.message); }
}

function handleObterClientes(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheet = ss.getSheetByName("Clientes");
    if (!sheet || sheet.getLastRow() < 2) {
      return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: [], rows: [] } });
    }
    var values = sheet.getDataRange().getValues();
    var headers = values.shift();
    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headers, rows: values } });
  } catch (e) { return responseErro(e.message); }
}

function handleObterVendas(data) {
  try {
    var ss = SpreadsheetApp.openById(data.spreadsheetId);
    var sheetVendas = ss.getSheetByName('Vendas');
    var sheetHistorico = ss.getSheetByName('Historico_Vendas');
    var todosDados = [];
    
    if (sheetVendas && sheetVendas.getLastRow() > 1) {
      todosDados = sheetVendas.getRange(2, 1, sheetVendas.getLastRow() - 1, sheetVendas.getLastColumn()).getValues();
    }

    var limite = new Date();
    limite.setDate(limite.getDate() - 60);
    
    var dIni = data.dataInicio ? new Date(data.dataInicio + 'T00:00:00') : null;
    if (dIni && !isNaN(dIni) && dIni < limite && sheetHistorico && sheetHistorico.getLastRow() > 1) {
      var dadosHist = sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, sheetHistorico.getLastColumn()).getValues();
      todosDados = todosDados.concat(dadosHist);
    }

    var rows = todosDados.map(function(row) {
      var dataFmt = row[1];
      if (dataFmt instanceof Date) {
        dataFmt = Utilities.formatDate(dataFmt, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      var vencFmt = row[12];
      if (vencFmt instanceof Date) {
        vencFmt = Utilities.formatDate(vencFmt, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return [
        row[0], dataFmt, row[2] || '', row[3] || '',
        row[4] || 0, row[5] || 0, row[6] || 0, row[7] || 0, row[8] || 0,
        row[9] || '', row[10] || '', row[11] || '', vencFmt, row[13] || '[]'
      ];
    });
    
    var headers = ['ID da Venda', 'Data', 'Cliente', 'Itens', 'Quantidade Vendida', 'Subtotal', 'Desconto (%)', 'Desconto (R$)', 'Total com Desconto', 'Forma de Pagamento', 'Usuario', 'Status', 'Vencimento', 'ItensJSON'];

    return responseSucessoMsg("Sucesso", { dados: { compact: true, headers: headers, rows: rows } });
  } catch (e) { return responseErro(e.message); }
}

function handleObterDashboard(data) {
  // O dashboard utiliza os mesmos dados das vendas (último ano)
  return handleObterVendas(data);
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
