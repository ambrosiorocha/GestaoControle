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
      var headers = ["Nome da Empresa / App", "Usuário Admin", "Link da Planilha", "ScriptURL", "Spreadsheet ID", "Link de Acesso", "Status", "Plano", "Ativação", "Expiração", "Observações"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#dcfce7");
      sheet.setFrozenRows(1);
    }
    
    var json = "{}";
    try { json = JSON.parse(e.postData.contents); } catch(err){}
    
    var reqData = json.data || json;
    
    var empresa = reqData.nome || "Novo Cliente";
    var usuario = reqData.usuario || "N/A";
    var spreadsheetUrl = reqData.spreadsheetUrl || "";
    var scriptUrl = reqData.scriptUrl || "";
    var spreadsheetId = reqData.spreadsheetId || "";
    
    if (!spreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({status: "erro", msg: "Sem ID da Planilha"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var dados = sheet.getDataRange().getValues();
    var rowToUpdate = -1;
    
    // Procura se o cliente já existe pela coluna E (índice 4: Spreadsheet ID)
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][4] === spreadsheetId) {
        rowToUpdate = i + 1;
        break;
      }
    }
    
    if (rowToUpdate > -1) {
      // Atualiza os dados de link caso tenham mudado
      sheet.getRange(rowToUpdate, 1).setValue(empresa);
      sheet.getRange(rowToUpdate, 2).setValue(usuario);
      sheet.getRange(rowToUpdate, 3).setValue(spreadsheetUrl);
      sheet.getRange(rowToUpdate, 4).setValue(scriptUrl);
      // Gera Link Mágico na Coluna F (Índice 6)
      if (scriptUrl) {
        var scriptIdMatch = scriptUrl.match(/\/s\/([^\/]+)\/exec/);
        if (scriptIdMatch && scriptIdMatch[1]) {
          var magicLink = "https://ambrosiorocha.github.io/VS_Teste/?id=" + scriptIdMatch[1];
          sheet.getRange(rowToUpdate, 6).setValue(magicLink);
        }
      }
    } else {
      // Cria nova linha com Plano Padrão = Básico
      // Extrai o ID do scriptUrl se houver para formar o Link Mágico
      var linkMagico = "";
      if (scriptUrl) {
          var scriptIdMatch = scriptUrl.match(/\/s\/([^\/]+)\/exec/);
          if (scriptIdMatch && scriptIdMatch[1]) {
              linkMagico = "https://ambrosiorocha.github.io/VS_Teste/?id=" + scriptIdMatch[1];
          }
      }

      var novaLinha = [
        empresa,
        usuario,
        spreadsheetUrl,
        scriptUrl,
        spreadsheetId,
        linkMagico,  // Coluna F: Link de Acesso
        "Ativo",     // Coluna G: Status
        "Básico",    // Coluna H: Plano Padrão
        "",          // Coluna I: Ativação
        "",          // Coluna J: Expiração
        "Registro automático via Login" // Coluna K: Obs
      ];
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
  
  // Detetando alteração na Coluna G (Plano - Índice 7) ou Expiração (Coluna I - Índice 9)
  if (row > 1 && (col === 7 || col === 9)) {
    var ss = e.source;
    var spreadsheetId = sheet.getRange(row, 5).getValue(); // Coluna E tem o ID
    var plano = sheet.getRange(row, 7).getValue();         // Coluna G tem o Plano
    var expiracao = sheet.getRange(row, 9).getValue();     // Coluna I tem a Expiração
    
    if (!spreadsheetId) return;
    
    try {
      var clientApp = SpreadsheetApp.openById(spreadsheetId);
      
      // Conforme as especificações, a trava acontece através de propriedades Chave-Valor.
      // E usamos a aba "Licença" para evitar colidir com a tabela padrão de "Configurações" de Operadores, 
      // embora estejamos criando a infraestrutura requisitada de plano master.
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
         for (var i = 1; i < dados.length; i++) {
           if (dados[i][0] === "Plano") {
             configSheet.getRange(i + 1, 2).setValue(plano);
             atualizouPlano = true;
           }
           if (dados[i][0] === "Expiração") {
             configSheet.getRange(i + 1, 2).setValue(expiracao);
             atualizouExp = true;
           }
         }
         if(!atualizouPlano) configSheet.appendRow(["Plano", plano]);
         if(!atualizouExp) configSheet.appendRow(["Expiração", expiracao]);
      }
      
      // Se a aba Configurações antiga existir (Lista de Operadores), também força a atualização visual de quem tem admin
      var oldConfig = clientApp.getSheetByName("Configurações");
      if (oldConfig) {
        var td = oldConfig.getDataRange().getValues();
        var headRow = td[0];
        // Encontra a coluna de plano
        var colPlanoIdx = headRow.indexOf("Plano"); // 0-indexed
        if (colPlanoIdx > -1) {
            for (var u = 1; u < td.length; u++) {
                if (td[u][1] === "Admin") { // admin level user
                  oldConfig.getRange(u + 1, colPlanoIdx + 1).setValue(plano);
                }
            }
        }
      }
      
      // Preenche a Coluna H ('Ativação') com a data atual se estiver em branco.
      var ativacaoRange = sheet.getRange(row, 8);
      if (!ativacaoRange.getValue()) {
        ativacaoRange.setValue(new Date());
      }
      
      sheet.getRange(row, col).clearNote();
      
    } catch(err) {
      sheet.getRange(row, col).setNote("Erro ao acessar cliente: " + err.message + "\n(Verifique se você é Mestre/Dono da planilha do cliente)");
    }
  }
}
