/**
 * ==========================================================
 * PARTE 1: SCRIPT NA PLANILHA MESTRA (O Cérebro)
 * ==========================================================
 * Autor: Gestão&Controle
 * 
 * Instruções de Instalação:
 * 1. Abra sua Planilha Mestra.
 * 2. Acesse no menu: Extensões > Apps Script.
 * 3. Cole este código substituindo todo o conteúdo atual.
 * 4. Salve o projeto.
 * 5. Clique em "Implantar" > "Nova Implantação".
 * 6. Tipo: App da Web.
 *    - Executar como: Você
 *    - Quem tem acesso: Qualquer pessoa
 * 7. Copie a URL do App da Web gerada e atualize o código do cliente ('Code.gs') com essa URL.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Clientes");
    
    // Auto-setup
    if (!sheet) {
      sheet = ss.insertSheet("Clientes");
      var headers = ["Nome da Empresa / App", "Usuário Admin", "WhatsApp", "Link da Planilha", "ScriptURL", "Spreadsheet ID", "Link de Acesso", "Status", "Plano", "Ativação", "Expiração", "Observações"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#dcfce7");
      sheet.setFrozenRows(1);
    }
    
    // Recepção Protegida
    var json = {};
    if (e && e.postData && e.postData.contents) {
      try { json = JSON.parse(e.postData.contents); } catch(err) { throw new Error("JSON Inválido: " + err.message); }
    } else {
      throw new Error("Payload Vazio");
    }
    
    var reqData = json.data || json;
    
    var empresa = reqData.nome || "Novo Cliente";
    var usuario = reqData.usuario || "N/A";
    var whatsapp = reqData.whatsapp || reqData.telefone || "";
    var registro = reqData.registro || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var spreadsheetUrl = reqData.spreadsheetUrl || "";
    var scriptUrl = reqData.scriptUrl || "";
    var spreadsheetId = reqData.spreadsheetId || "";
    var planoPayload = reqData.plano || "Básico";
    var ativacaoPayload = reqData.ativacao || "";
    var expiracaoPayload = reqData.expiracao || "";
    
    if (!spreadsheetId) {
      throw new Error("Faltou enviar o ID da Planilha");
    }
    
    var dados = sheet.getDataRange().getValues();
    var headersCurrent = dados[0].map(function(h){ return String(h).trim().toLowerCase(); });
    
    // Função utilitária para match seguro
    function getColIdx(namesArray) {
       for(var i=0; i<headersCurrent.length; i++){
          var h = headersCurrent[i];
          for(var j=0; j<namesArray.length; j++){
             if(h === namesArray[j].toLowerCase().trim()) return i;
          }
       }
       return -1;
    }
    
    var colId = getColIdx(["Spreadsheet ID", "ID Planilha", "ID"]);
    if(colId === -1) colId = 5; // fallback bruto
    
    var colEmp = getColIdx(["Nome da Empresa / App", "Empresa"]);
    var colUser = getColIdx(["Usuário Admin", "Usuário"]);
    var colZap = getColIdx(["WhatsApp", "Telefone", "Contato"]);
    var colSpread = getColIdx(["Link da Planilha", "URL Planilha"]);
    var colScript = getColIdx(["ScriptURL", "URL WebApp"]);
    var colLink = getColIdx(["Link de Acesso", "Link Mágico"]);
    var colStatus = getColIdx(["Status"]);
    var colPlano = getColIdx(["Plano", "Plano Atual"]);
    var colAtivacao = getColIdx(["Ativação", "Data de Ativação"]);
    var colExpiracao = getColIdx(["Expiração", "Data de Expiração"]);
    var colObs = getColIdx(["Observações", "Obs"]);

    var rowToUpdate = -1;
    for (var i = 1; i < dados.length; i++) {
        var idMatch = (colId > -1 && dados[i][colId] === spreadsheetId);
        
        var scriptMatch = false;
        if (colScript > -1 && scriptUrl) {
            var urlPlanilha = String(dados[i][colScript]).trim();
            var id1 = scriptUrl.match(/\/s\/([^\/]+)/);
            var id2 = urlPlanilha.match(/\/s\/([^\/]+)/);
            if (id1 && id2 && id1[1] === id2[1]) {
                scriptMatch = true;
            } else if (urlPlanilha === scriptUrl.trim()) {
                scriptMatch = true;
            }
        }
        
        if (idMatch || scriptMatch) {
            rowToUpdate = i + 1;
            break;
        }
    }
    
    // Gerar Link Mágico
    var linkMagico = "";
    if (scriptUrl) {
      var scriptIdMatch = scriptUrl.match(/\/s\/([^\/]+)\/exec/);
      if (scriptIdMatch && scriptIdMatch[1]) {
        linkMagico = "https://ambrosiorocha.github.io/VS_Teste/?id=" + scriptIdMatch[1];
      }
    }

    if (rowToUpdate > -1) {
      // Atualiza Cliente Existente
      var sIdAtual = colId > -1 ? String(sheet.getRange(rowToUpdate, colId + 1).getValue()) : "";
      var statusAtual = colStatus > -1 ? String(sheet.getRange(rowToUpdate, colStatus + 1).getValue()) : "";
      var obsAtual = colObs > -1 ? String(sheet.getRange(rowToUpdate, colObs + 1).getValue()) : "";

      if(colEmp > -1 && empresa) sheet.getRange(rowToUpdate, colEmp + 1).setValue(empresa);
      if(colUser > -1 && usuario) sheet.getRange(rowToUpdate, colUser + 1).setValue(usuario);
      if(colZap > -1 && whatsapp) sheet.getRange(rowToUpdate, colZap + 1).setValue(whatsapp);
      if(colSpread > -1 && spreadsheetUrl) sheet.getRange(rowToUpdate, colSpread + 1).setValue(spreadsheetUrl);
      if(colScript > -1 && scriptUrl) sheet.getRange(rowToUpdate, colScript + 1).setValue(scriptUrl);
      if(colLink > -1 && linkMagico) sheet.getRange(rowToUpdate, colLink + 1).setValue(linkMagico);
      if(colPlano > -1 && planoPayload) sheet.getRange(rowToUpdate, colPlano + 1).setValue(planoPayload);
      if(colAtivacao > -1 && ativacaoPayload) sheet.getRange(rowToUpdate, colAtivacao + 1).setValue(ativacaoPayload);
      if(colExpiracao > -1 && expiracaoPayload) sheet.getRange(rowToUpdate, colExpiracao + 1).setValue(expiracaoPayload);
      
      // Os novos blocos gravam dados em campos vazios para não sobrescrever:
      if(colId > -1 && spreadsheetId && (!sIdAtual || sIdAtual === "undefined" || sIdAtual === "")) {
          sheet.getRange(rowToUpdate, colId + 1).setValue(spreadsheetId);
      }
      if(colStatus > -1 && (!statusAtual || statusAtual === "")) {
          sheet.getRange(rowToUpdate, colStatus + 1).setValue("Ativo");
      }
      if(colObs > -1) {
          var textObs = obsAtual ? obsAtual + " / Primeiro Acesso Realizado (" + registro + ")" : "Registro auto (Data: " + registro + ")";
          sheet.getRange(rowToUpdate, colObs + 1).setValue(textObs);
      }
    } else {
      // Inserir Novo Cliente Seguro (evita jogar para a linha 1000 se houver fórmulas arrastadas)
      var novaLinha = [];
      for(var c=0; c<headersCurrent.length; c++) novaLinha.push(""); // Inicializa com strings preenchidas
      
      if(colEmp > -1) novaLinha[colEmp] = empresa;
      if(colUser > -1) novaLinha[colUser] = usuario;
      if(colZap > -1) novaLinha[colZap] = whatsapp;
      if(colSpread > -1) novaLinha[colSpread] = spreadsheetUrl;
      if(colScript > -1) novaLinha[colScript] = scriptUrl;
      if(colId > -1) novaLinha[colId] = spreadsheetId;
      if(colLink > -1) novaLinha[colLink] = linkMagico;
      if(colStatus > -1) novaLinha[colStatus] = "Ativo";
      if(colPlano > -1) novaLinha[colPlano] = planoPayload;
      if(colAtivacao > -1) novaLinha[colAtivacao] = ativacaoPayload;
      if(colExpiracao > -1) novaLinha[colExpiracao] = expiracaoPayload;
      if(colObs > -1) novaLinha[colObs] = "Registro auto (Data: " + registro + ")";
      
      // Fallback
      if(colId === -1 || colEmp === -1){
         novaLinha = [empresa, usuario, whatsapp, spreadsheetUrl, scriptUrl, spreadsheetId, linkMagico, "Ativo", planoPayload, ativacaoPayload, expiracaoPayload, "Registro auto"];
      }
      
      // Procura a primeira linha realmente vazia na Coluna A (Empresa) e E (ScriptURL)
      var inseriu = false;
      for (var k = 1; k < dados.length; k++) {
         if (!dados[k][0] && (!colScript || colScript === -1 || !dados[k][colScript])) {
            sheet.getRange(k + 1, 1, 1, novaLinha.length).setValues([novaLinha]);
            inseriu = true;
            break;
         }
      }
      if (!inseriu) {
         sheet.appendRow(novaLinha);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "sucesso"})).setMimeType(ContentService.MimeType.JSON);
    
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "erro", msg: err.message, stack: err.stack})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Clientes") return;
  
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row <= 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colPlano = headers.indexOf("Plano") + 1;
  var colExp = headers.indexOf("Expiração") + 1;
  var colId = headers.indexOf("Spreadsheet ID") + 1;
  var colAtiv = headers.indexOf("Ativação") + 1;
  
  // Detetando alteração na Coluna Plano, Expiração ou Nome (dinâmico)
  var colEmpresa = headers.indexOf("Nome da Empresa / App") + 1;
  
  if (col === colPlano || col === colExp || (colEmpresa > 0 && col === colEmpresa)) {
    var ss = e.source;
    var spreadsheetId = sheet.getRange(row, colId).getValue();
    var plano = sheet.getRange(row, colPlano).getValue();
    var expiracao = sheet.getRange(row, colExp).getValue();
    var empresaNome = colEmpresa > 0 ? sheet.getRange(row, colEmpresa).getValue() : "";
    
    if (!spreadsheetId) return;
    
    try {
      var clientApp = SpreadsheetApp.openById(spreadsheetId);
      
      // Conforme as especificações, a trava acontece através de propriedades Chave-Valor.
      var configSheet = clientApp.getSheetByName("Licença");
      
      if (!configSheet) {
         configSheet = clientApp.insertSheet("Licença");
         configSheet.appendRow(["Propriedade", "Valor"]);
         configSheet.appendRow(["Plano", plano]);
         configSheet.appendRow(["Expiração", expiracao]);
         configSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f3f4f6");
      } else {
         var dados = configSheet.getDataRange().getValues();
         var atualizouPlano = false;
         var atualizouExp = false;
         var atualizouEmp = false;
         for (var i = 1; i < dados.length; i++) {
           if (dados[i][0] === "Plano") {
             configSheet.getRange(i + 1, 2).setValue(plano);
             atualizouPlano = true;
           }
           if (dados[i][0] === "Expiração") {
             configSheet.getRange(i + 1, 2).setValue(expiracao);
             atualizouExp = true;
           }
           if (dados[i][0] === "Empresa") {
             configSheet.getRange(i + 1, 2).setValue(empresaNome);
             atualizouEmp = true;
           }
         }
         if(!atualizouPlano) configSheet.appendRow(["Plano", plano]);
         if(!atualizouExp) configSheet.appendRow(["Expiração", expiracao]);
         if(!atualizouEmp && empresaNome) configSheet.appendRow(["Empresa", empresaNome]);
      }
      
      // Se a aba Configurações antiga existir (Lista de Operadores), também força a atualização visual de quem tem admin
      var oldConfig = clientApp.getSheetByName("Configurações");
      if (oldConfig) {
        var td = oldConfig.getDataRange().getValues();
        var headRow = td[0];
        var colPlanoIdx = headRow.indexOf("Plano"); // 0-indexed
        if (colPlanoIdx > -1) {
            for (var u = 1; u < td.length; u++) {
                if (td[u][1] === "Admin") { // admin level user
                  oldConfig.getRange(u + 1, colPlanoIdx + 1).setValue(plano);
                }
            }
        }
      }
      
      // Preenche a Coluna 'Ativação' com a data atual se estiver em branco.
      if (colAtiv > 0) {
          var ativacaoRange = sheet.getRange(row, colAtiv);
          if (!ativacaoRange.getValue()) {
            ativacaoRange.setValue(new Date());
          }
      }
      
      sheet.getRange(row, col).clearNote();
      
    } catch(err) {
      sheet.getRange(row, col).setNote("Erro ao acessar cliente: " + err.message + "\n(Verifique se você é Mestre/Dono da planilha do cliente)");
    }
  }
}
