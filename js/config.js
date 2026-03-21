// config.js — Configurações Globais (Cérebro Único)

// 1. URL da Planilha Mestra (API Gateway)
window.MASTER_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbw1YlcOcSpNQA7FjWjqVVrl_z7aux0mGI2de2sQ6cfVPCfnWgV4QOuOKjezKesoOBLx/exec";

(function () {
    // 2. Tentar ler parâmetros da URL: ?id=XYZ&setup=true
    const urlParams = new URLSearchParams(window.location.search);
    let paramId = urlParams.get('id') || "";
    let paramSetup = urlParams.get('setup') === 'true';

    // Caso o setup=true esteja "dentro" do parâmetro ID (comum em links de WhatsApp)
    if (paramId.includes('setup=true')) {
        paramSetup = true;
    }

    if (paramId) {
        // Limpa o ID de possíveis sufixos /exec ou URLs completas
        let cleanId = paramId.trim().replace(/https:\/\/script\.google\.com\/macros\/s\/|\/exec\/?$/g, '');
        // Remove eventuais parâmetros de query que sobraram (ex: ID?setup=true)
        cleanId = cleanId.split('?')[0].split('&')[0];

        localStorage.setItem('sv_spreadsheet_id', cleanId.trim());
    }

    // 3. Captura o modo de configuração (?setup=true)
    window.IS_SETUP = paramSetup;

    if (paramId || paramSetup) {
        // Remove os parâmetros da URL para manter a interface limpa
        window.history.replaceState({}, document.title, window.location.pathname);
    }

    // 4. Define o ID Global
    window.SPREADSHEET_ID = localStorage.getItem('sv_spreadsheet_id') || "";

    // Se não houver ID, podemos redirecionar ou mostrar um aviso (opcional)
    if (!window.SPREADSHEET_ID && !window.location.pathname.includes('index.html')) {
        console.warn("Nenhum Spreadsheet ID localizado. O sistema pode não funcionar corretamente.");
    }
})();
