(function () {
    // 1. Tentar ler da URL: ?id=XYZ
    const urlParams = new URLSearchParams(window.location.search);
    const paramId = urlParams.get('id');

    // Se passarem um param ?id=NaURL, salvar no localStorage
    if (paramId) {
        let cleanId = paramId.trim().replace(/\/exec\/?$/, ''); // Evita duplicação do /exec
        let finalUrl = cleanId.startsWith('http') ? cleanId : `https://script.google.com/macros/s/${cleanId}/exec`;

        localStorage.setItem('script_url', finalUrl);

        // Força o reload limpo para igualar o ciclo de vida do salvamento manual
        window.location.href = window.location.pathname;
        return; // interrompe a execução até que recarregue limpo
    }

    // 2. Tentar ler do localStorage
    let storedUrl = localStorage.getItem('script_url');

    if (storedUrl) {
        window.SCRIPT_URL = storedUrl;
    } else {
        // 3. Se for a primeira vez e não tiver URL, definimos como vazia e mostramos campo configurador
        window.SCRIPT_URL = '';

        // Modal for initial configuration (Interrompe o fluxo e força o cadastro)
        document.addEventListener('DOMContentLoaded', () => {
            const overlay = document.createElement('div');
            overlay.id = 'setupScriptOverlay';
            overlay.style.cssText = 'position:fixed;inset:0;background:linear-gradient(135deg,#0f172a 0%,#16a34a 100%);z-index:99999;display:flex;align-items:center;justify-content:center;padding:1rem;font-family:sans-serif;';
            overlay.innerHTML = `
                <div style="background:white;padding:2rem;border-radius:1rem;max-width:400px;width:100%;box-shadow:0 10px 25px rgba(0,0,0,0.5);text-align:center;animation:fadeInUp 0.35s ease;">
                    <style>@keyframes fadeInUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}</style>
                    <div style="width:64px;height:64px;border-radius:18px;background:linear-gradient(135deg,#16a34a,#15803d);display:flex;align-items:center;justify-content:center;margin:0 auto 1.25rem;box-shadow:0 8px 24px rgba(22,163,74,0.35);">
                        <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <rect x="2" y="3" width="20" height="14" rx="2" ry="2"></rect><line x1="8" y1="21" x2="16" y2="21"></line><line x1="12" y1="17" x2="12" y2="21"></line>
                        </svg>
                    </div>
                    <h2 style="margin-top:0;color:#1e293b;font-size:1.4rem;">🔧 Configuração da Instância</h2>
                    <p style="color:#64748b;font-size:0.9rem;margin-bottom:1.5rem;line-height:1.4;">Para iniciar, cole a <strong>URL do App Script</strong> correspondente à sua planilha mestra (banco de dados).</p>
                    <input type="text" id="setupScriptUrl" placeholder="https://script.google.com/macros/s/.../exec" style="width:100%;padding:0.8rem;border:1.5px solid #cbd5e1;border-radius:0.5rem;margin-bottom:1rem;box-sizing:border-box;font-size:0.85rem;outline:none;" autocomplete="off">
                    <button id="btnSaveConfig" style="width:100%;padding:0.9rem;background:linear-gradient(135deg,#16a34a,#15803d);color:white;border:none;border-radius:0.5rem;font-weight:700;cursor:pointer;font-size:0.95rem;box-shadow:0 4px 12px rgba(22,163,74,0.35);transition:transform 0.1s;">Conectar Banco de Dados</button>
                    <p style="font-size:0.75rem;color:#94a3b8;margin-top:1.25rem;margin-bottom:0;">Ou acesse usando um link mágico:<br><strong>?id=SEU_SCRIPT_ID</strong></p>
                </div>
            `;
            document.body.appendChild(overlay);

            // Evitar que o resto do app inicialize mostrando erros na tela
            document.body.style.overflow = 'hidden';

            const input = document.getElementById('setupScriptUrl');
            input.addEventListener('focus', () => input.style.borderColor = '#16a34a');
            input.addEventListener('blur', () => input.style.borderColor = '#cbd5e1');
            input.addEventListener('keydown', (e) => { if (e.key === 'Enter') document.getElementById('btnSaveConfig').click(); });

            document.getElementById('btnSaveConfig').onclick = () => {
                const val = input.value.trim();
                // Basic validation
                if (val && val.includes('script.google.com/macros/s/')) {
                    localStorage.setItem('script_url', val);
                    window.location.reload();
                } else {
                    input.style.borderColor = '#ef4444'; // Red error outline
                    alert('Por favor, insira uma URL válida do Google Apps Script (App Web).');
                }
            };
        });
    }
})();
