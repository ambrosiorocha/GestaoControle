// config.js — Configurações Globais (Cérebro Único)

// 1. URL da Planilha Mestra (API Gateway)
window.MASTER_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbw1YlcOcSpNQA7FjWjqVVrl_z7aux0mGI2de2sQ6cfVPCfnWgV4QOuOKjezKesoOBLx/exec";

(function () {
    // 2. Tentar ler o Spreadsheet ID da URL: ?id=XYZ
    const urlParams = new URLSearchParams(window.location.search);
    const paramId = urlParams.get('id');

    if (paramId) {
        // Limpa o ID de possíveis sufixos /exec ou URLs completas
        let cleanId = paramId.trim().replace(/https:\/\/script\.google\.com\/macros\/s\/|\/exec\/?$/g, '');

        localStorage.setItem('sv_spreadsheet_id', cleanId);

        // Remove o ID da URL para manter a interface limpa
        window.history.replaceState({}, document.title, window.location.pathname);
    }

    // 3. Define o ID Global
    window.SPREADSHEET_ID = localStorage.getItem('sv_spreadsheet_id') || "";

    // Se não houver ID, podemos redirecionar ou mostrar um aviso (opcional)
    if (!window.SPREADSHEET_ID && !window.location.pathname.includes('index.html')) {
        console.warn("Nenhum Spreadsheet ID localizado. O sistema pode não funcionar corretamente.");
    }
})();
