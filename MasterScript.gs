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
    
    // Auto-setup da aba Clientes na Mestra se não existir
    if (!sheet) {
      sheet = ss.insertSheet("Clientes");
      var headers = ["Nome da Empresa / App", "Usuário Admin", "WhatsApp", "Link da Planilha", "ScriptURL", "Spreadsheet ID", "Link de Acesso", "Status", "Plano", "Ativação", "Expiração", "Observações"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#dcfce7");
      sheet.setFrozenRows(1);
    }
    
    var json = "{}";
    try { json = JSON.parse(e.postData.contents); } catch(err){}
    
    var reqData = json.data || json;
    
    var empresa = reqData.nome || "Novo Cliente";
    var usuario = reqData.usuario || "N/A";
    var whatsapp = reqData.whatsapp || reqData.telefone || "";
    var registro = reqData.registro || "";
    var spreadsheetUrl = reqData.spreadsheetUrl || "";
    var scriptUrl = reqData.scriptUrl || "";
    var spreadsheetId = reqData.spreadsheetId || "";
    var planoPayload = reqData.plano || "Básico";
    var ativacaoPayload = reqData.ativacao || "";
    var expiracaoPayload = reqData.expiracao || "";
    
    if (!spreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({status: "erro", msg: "Sem ID da Planilha"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var dados = sheet.getDataRange().getValues();
    var headersCurrent = dados[0];
    var colId = headersCurrent.indexOf("Spreadsheet ID");
    if(colId === -1) colId = 5; // fallback se o cabeçalho mudou (5 era o antigo ID ou 4 era o mais antigo)
    
    var rowToUpdate = -1;
    // Procura se o cliente já existe pela coluna de ID
    for (var i = 1; i < dados.length; i++) {
        // Checando na coluna encontrada, mas com fallback de leniência em algumas colunas
        if (dados[i][colId] === spreadsheetId || dados[i][4] === spreadsheetId || dados[i][5] === spreadsheetId) {
            rowToUpdate = i + 1;
            break;
        }
    }
    
    var colEmp = headersCurrent.indexOf("Nome da Empresa / App");
    var colUser = headersCurrent.indexOf("Usuário Admin");
    var colZap = headersCurrent.indexOf("WhatsApp");
    var colSpread = headersCurrent.indexOf("Link da Planilha");
    var colScript = headersCurrent.indexOf("ScriptURL");
    var colLink = headersCurrent.indexOf("Link de Acesso");

    if (rowToUpdate > -1) {
      // Atualiza os dados de link caso tenham mudado
      if(colEmp > -1) sheet.getRange(rowToUpdate, colEmp + 1).setValue(empresa);
      if(colUser > -1) sheet.getRange(rowToUpdate, colUser + 1).setValue(usuario);
      if(colZap > -1) sheet.getRange(rowToUpdate, colZap + 1).setValue(whatsapp);
      if(colSpread > -1) sheet.getRange(rowToUpdate, colSpread + 1).setValue(spreadsheetUrl);
      if(colScript > -1) sheet.getRange(rowToUpdate, colScript + 1).setValue(scriptUrl);
      var colPlano = headersCurrent.indexOf("Plano");
      if(colPlano > -1) sheet.getRange(rowToUpdate, colPlano + 1).setValue(planoPayload);
      var colAtivacao = headersCurrent.indexOf("Ativação");
      if(colAtivacao > -1 && ativacaoPayload) sheet.getRange(rowToUpdate, colAtivacao + 1).setValue(ativacaoPayload);
      var colExpiracao = headersCurrent.indexOf("Expiração");
      if(colExpiracao > -1 && expiracaoPayload) sheet.getRange(rowToUpdate, colExpiracao + 1).setValue(expiracaoPayload);

      // Gera Link Mágico
      if (scriptUrl && colLink > -1) {
        var scriptIdMatch = scriptUrl.match(/\/s\/([^\/]+)\/exec/);
        if (scriptIdMatch && scriptIdMatch[1]) {
          var magicLink = "https://ambrosiorocha.github.io/VS_Teste/?id=" + scriptIdMatch[1];
          sheet.getRange(rowToUpdate, colLink + 1).setValue(magicLink);
        }
      }
    } else {
      // Extrai o ID do scriptUrl se houver para formar o Link Mágico
      var linkMagico = "";
      if (scriptUrl) {
          var scriptIdMatch = scriptUrl.match(/\/s\/([^\/]+)\/exec/);
          if (scriptIdMatch && scriptIdMatch[1]) {
              linkMagico = "https://ambrosiorocha.github.io/VS_Teste/?id=" + scriptIdMatch[1];
          }
      }

      var novaLinha = [];
      for(var c=0; c<headersCurrent.length; c++) novaLinha.push("");
      
      if(colEmp > -1) novaLinha[colEmp] = empresa;
      if(colUser > -1) novaLinha[colUser] = usuario;
      if(colZap > -1) novaLinha[colZap] = whatsapp;
      if(colSpread > -1) novaLinha[colSpread] = spreadsheetUrl;
      if(colScript > -1) novaLinha[colScript] = scriptUrl;
      if(colId > -1) novaLinha[colId] = spreadsheetId;
      if(colLink > -1) novaLinha[colLink] = linkMagico;
      
      var colStatus = headersCurrent.indexOf("Status");
      if(colStatus > -1) novaLinha[colStatus] = "Ativo";
      var colPlano = headersCurrent.indexOf("Plano");
      if(colPlano > -1) novaLinha[colPlano] = planoPayload;
      var colAtivacao = headersCurrent.indexOf("Ativação");
      if(colAtivacao > -1) novaLinha[colAtivacao] = ativacaoPayload;
      var colExpiracao = headersCurrent.indexOf("Expiração");
      if(colExpiracao > -1) novaLinha[colExpiracao] = expiracaoPayload;
      var colObs = headersCurrent.indexOf("Observações");
      if(colObs > -1) novaLinha[colObs] = "Registro auto (Data: " + registro + ")";

      // fallback simple append row se cabecalho estiver totalmente desconfigurado
      if(colId === -1 || colEmp === -1){
         novaLinha = [empresa, usuario, whatsapp, spreadsheetUrl, scriptUrl, spreadsheetId, linkMagico, "Ativo", "Básico", "", "", "Registro automático via Login"];
      }
      sheet.appendRow(novaLinha);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "sucesso"})).setMimeType(ContentService.MimeType.JSON);
    
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status: "erro", msg: err.message})).setMimeType(ContentService.MimeType.JSON);
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
