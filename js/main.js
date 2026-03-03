// main.js — Sidebar responsiva + Auth
const APP_VERSION = '1.0.0';

document.addEventListener('DOMContentLoaded', function () {

    // ── Injetar CSS do menu mobile ─────────────────────────────
    const style = document.createElement('style');
    style.textContent = `
        #hamburgerBtn {
            display: none;
            position: fixed; top: 0.75rem; left: 0.75rem; z-index: 1100;
            background: #16a34a; color: white; border: none; border-radius: 0.5rem;
            padding: 0.55rem 0.75rem; font-size: 1.2rem; cursor: pointer;
            box-shadow: 0 2px 8px rgba(0,0,0,0.25);
        }
        #sidebarOverlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,0.45); z-index: 1050;
        }
        #sidebarOverlay.open { display: block; }
        @media (max-width: 767px) {
            #hamburgerBtn { display: flex; align-items: center; justify-content: center; }
            #desktop-sidebar {
                position: fixed !important; left: -280px; top: 0; height: 100vh;
                width: 260px !important; z-index: 1060; transition: left 0.25s ease;
                box-shadow: 4px 0 24px rgba(0,0,0,0.35);
            }
            #desktop-sidebar.sidebar-open { left: 0 !important; }
            .desktop-sidebar { display: flex !important; }
            .hidden.md\\:flex { display: flex !important; }
            .main-content { padding-top: 3.5rem !important; }
        }
        .user-badge-block {
            padding: 0.6rem 1rem; border-top: 1px solid rgba(255,255,255,0.12);
            margin-top: auto; font-size: 0.78rem; color: rgba(255,255,255,0.85);
            display: flex; flex-direction: column; gap: 0.35rem;
        }
        .logout-btn {
            background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2);
            color: white; border-radius: 0.375rem; padding: 0.35rem 0.6rem;
            font-size: 0.72rem; cursor: pointer; text-align: left; width: 100%;
            transition: background 0.15s;
        }
        .logout-btn:hover { background: rgba(239,68,68,0.4); }
    `;
    document.head.appendChild(style);

    // ── Hamburger e overlay mobile ─────────────────────────────
    const hamBtn = document.createElement('button');
    hamBtn.id = 'hamburgerBtn';
    hamBtn.innerHTML = '☰';
    hamBtn.setAttribute('aria-label', 'Menu');
    document.body.appendChild(hamBtn);

    const overlay = document.createElement('div');
    overlay.id = 'sidebarOverlay';
    document.body.appendChild(overlay);

    function openSidebar() { document.getElementById('desktop-sidebar')?.classList.add('sidebar-open'); overlay.classList.add('open'); }
    function closeSidebar() { document.getElementById('desktop-sidebar')?.classList.remove('sidebar-open'); overlay.classList.remove('open'); }
    hamBtn.addEventListener('click', openSidebar);
    overlay.addEventListener('click', closeSidebar);

    // ── Sidebar HTML ───────────────────────────────────────────
    const sidebar = `
        <aside id="desktop-sidebar" class="desktop-sidebar hidden md:flex flex-col">
            <div class="sidebar-header" style="display:flex;justify-content:center;align-items:center;gap:8px;">
                <img src="assets/logo.png" alt="Logo" style="height:28px; object-fit:contain; filter: brightness(0) invert(1);"> 
                <span>Gestão&Controle</span>
            </div>
            <nav id="desktop-nav" class="sidebar-nav flex-1">
                <a href="index.html" class="nav-link-menu">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>
                    <span>Início</span>
                </a>
                <a href="Vendas.html" class="nav-link-menu">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="9" cy="21" r="1"></circle><circle cx="20" cy="21" r="1"></circle><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"></path></svg>
                    <span>Vendas</span>
                </a>
                <a href="Produtos.html" class="nav-link-menu">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="4" width="22" height="6" rx="2" ry="2"></rect><rect x="1" y="14" width="22" height="6" rx="2" ry="2"></rect></svg>
                    <span>Produtos</span>
                </a>
                <a href="Clientes.html" class="nav-link-menu">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                    <span>Clientes</span>
                </a>
                <a href="Fornecedores.html" class="nav-link-menu">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="3" width="15" height="13"></rect><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"></polygon><circle cx="5.5" cy="18.5" r="2.5"></circle><circle cx="18.5" cy="18.5" r="2.5"></circle></svg>
                    <span>Fornecedores</span>
                </a>
                <a href="Financeiro.html" class="nav-link-menu" data-admin-only>
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"></line><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path></svg>
                    <span>Financeiro</span>
                </a>
                <a href="Relatorios.html" class="nav-link-menu" data-admin-only>
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3h7v9H3z"></path><path d="M14 3h7v5h-7z"></path><path d="M14 12h7v9h-7z"></path><path d="M3 16h7v5h-7z"></path></svg>
                    <span>Relatórios</span>
                </a>
            </nav>
            <!-- Badge do usuário logado -->
            <div class="user-badge-block">
                <div id="userBadge" style="line-height:1.4;">—</div>
                <button class="logout-btn" onclick="Auth.logout()">⎋ Sair do sistema</button>
            </div>
            
            <!-- Footer Version & Help -->
            <div style="padding: 0.75rem 1rem; border-top: 1px solid rgba(255,255,255,0.08); display:flex; justify-content:space-between; align-items:center;">
                <span class="text-xs text-gray-400">v${APP_VERSION}</span>
                <button onclick="openHelpModal()" class="w-6 h-6 rounded-full bg-slate-700 hover:bg-slate-600 text-gray-300 text-xs font-bold flex items-center justify-center transition-colors" title="Ajuda">?</button>
            </div>
        </aside>

        <!-- Help Modal -->
        <div id="helpModal" class="fixed inset-0 bg-slate-900/60 z-[1200] hidden items-center justify-center p-4">
            <div class="bg-white rounded-xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[90vh]">
                <div class="p-4 border-b border-gray-100 flex justify-between items-center bg-slate-50">
                    <h3 class="font-bold text-slate-800 text-lg flex items-center gap-2">
                        <span class="bg-blue-100 text-blue-700 w-7 h-7 rounded-full flex items-center justify-center text-sm">?</span>
                        Central de Ajuda
                    </h3>
                    <button onclick="closeHelpModal()" class="text-gray-400 hover:text-red-500 transition-colors">
                        <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
                    </button>
                </div>
                
                <div class="flex border-b border-gray-200">
                    <button onclick="switchHelpTab('vendas')" id="tab-vendas" class="flex-1 py-3 text-sm font-semibold text-blue-600 border-b-2 border-blue-600 bg-blue-50/30 transition-colors">Vendas</button>
                    <button onclick="switchHelpTab('financeiro')" id="tab-financeiro" class="flex-1 py-3 text-sm font-semibold text-gray-500 border-b-2 border-transparent hover:text-gray-700 hover:bg-gray-50 transition-colors">Financeiro</button>
                    <button onclick="switchHelpTab('suporte')" id="tab-suporte" class="flex-1 py-3 text-sm font-semibold text-gray-500 border-b-2 border-transparent hover:text-gray-700 hover:bg-gray-50 transition-colors">Suporte</button>
                </div>
                
                <div class="p-5 overflow-y-auto">
                    <div id="content-vendas" class="help-content block">
                        <h4 class="font-semibold text-slate-800 mb-2">Como lançar pedido e gerar cupom</h4>
                        <p class="text-sm text-gray-600 leading-relaxed mb-3">Selecione o cliente e adicione produtos ao carrinho listando quantidade e preço. O sistema calculará o total.</p>
                        <ul class="text-sm text-gray-600 list-disc pl-5 space-y-1">
                            <li><strong>Salvar Rascunho:</strong> Guarda o pedido em aberto na tela de Vendas. Não baixa estoque nem gera contas a receber.</li>
                            <li><strong>Finalizar Venda:</strong> Conclui a baixa no estoque e lança no Financeiro. O botão de imprimir cupom aparecerá no Recibo de Conclusão e no Histórico de Vendas.</li>
                        </ul>
                    </div>
                    
                    <div id="content-financeiro" class="help-content hidden">
                        <h4 class="font-semibold text-slate-800 mb-2">Pendente (Atrasado) vs Concluído</h4>
                        <p class="text-sm text-gray-600 leading-relaxed mb-3">Vendas finalizadas a prazo e lançamentos a pagar/receber entram no Financeiro como <span class="text-yellow-600 font-medium badge-status">Pendentes</span>.</p>
                        <ul class="text-sm text-gray-600 list-disc pl-5 space-y-2">
                            <li><strong>Atrasos:</strong> Vendas pendentes com data de vencimento ultrapassada alimentam os alertas em vermelho no Dashboard.</li>
                            <li><strong>Baixa:</strong> Marque a conta como <span class="text-green-600 font-medium badge-status">Pago/Concluído</span> na tela Financeiro ou Relatórios assim que receber o pagamento, para atualizar os caixas.</li>
                        </ul>
                    </div>
                    
                    <div id="content-suporte" class="help-content hidden text-center py-4">
                        <div class="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                            <svg class="w-8 h-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M18.36 6.64a9 9 0 1 1-12.73 0M12 2v2m0 16v2m-8.5-8.5h2m16 0h2"></path></svg>
                        </div>
                        <h4 class="font-semibold text-slate-800 mb-2">Precisa de ajuda técnica?</h4>
                        <p class="text-sm text-gray-600 mb-5">Fale diretamente com o desenvolvedor do sistema para relatar erros, pedir novos recursos ou tirar dúvidas avançadas.</p>
                        <a href="https://wa.me/5581981438242" target="_blank" class="inline-flex items-center gap-2 px-5 py-2.5 bg-green-600 hover:bg-green-700 text-white text-sm font-semibold rounded-lg transition-colors shadow-sm">
                            <svg class="w-5 h-5" fill="currentColor" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
                            Falar no WhatsApp
                        </a>
                    </div>
                </div>
            </div>
        </div>
    `;
    document.body.insertAdjacentHTML('afterbegin', sidebar);

    // ── Link ativo ─────────────────────────────────────────────
    const currentPath = window.location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll('.nav-link-menu').forEach(link => {
        if (link.getAttribute('href') === currentPath) link.classList.add('active');
    });

    // ── Inicializar autenticação ───────────────────────────────
    if (typeof Auth !== 'undefined') {
        Auth.init(() => {
            // Preenche operador na tela de vendas se disponível
            const opSel = document.getElementById('usuario');
            if (opSel && Auth.getUser()) {
                // Garante que a opção do usuário logado existe
                let found = Array.from(opSel.options).some(o => o.value === Auth.getUser());
                if (!found) {
                    const opt = new Option(Auth.getUser(), Auth.getUser());
                    opSel.insertBefore(opt, opSel.firstChild);
                }
                opSel.value = Auth.getUser();
                if (!Auth.isAdmin()) opSel.disabled = true; // Operador não pode trocar
            }

            // Atualiza saudação no index se existir
            const greetEl = document.getElementById('greetingMsg');
            if (greetEl && Auth.getUser()) {
                greetEl.textContent = `Bem-vindo(a) ao Gestão&Controle, ${Auth.getUser()}!`;
            }
        });
    }
});

// ── Funções Globais da Central de Ajuda ───────────────────────
window.openHelpModal = function () {
    const modal = document.getElementById('helpModal');
    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
        window.switchHelpTab('vendas'); // padrao
    }
};

window.closeHelpModal = function () {
    const modal = document.getElementById('helpModal');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

window.switchHelpTab = function (tabName) {
    document.querySelectorAll('.help-content').forEach(c => c.classList.add('hidden'));
    document.querySelectorAll('.help-content').forEach(c => c.classList.remove('block'));

    ['vendas', 'financeiro', 'suporte'].forEach(t => {
        const btn = document.getElementById('tab-' + t);
        if (btn) btn.className = "flex-1 py-3 text-sm font-semibold text-gray-500 border-b-2 border-transparent hover:text-gray-700 hover:bg-gray-50 transition-colors";
    });

    const content = document.getElementById('content-' + tabName);
    const tabBtn = document.getElementById('tab-' + tabName);

    if (content) {
        content.classList.remove('hidden');
        content.classList.add('block');
    }
    if (tabBtn) {
        tabBtn.className = "flex-1 py-3 text-sm font-semibold text-blue-600 border-b-2 border-blue-600 bg-blue-50/30 transition-colors";
    }
};

// ==========================================
// UTILITÁRIOS GLOBAIS DE OTIMIZAÇÃO
// ==========================================

// Parse do JSON Compacto
function parseCompactData(data) {
    if (data && data.compact) {
        return data.rows.map(row => {
            let obj = {};
            data.headers.forEach((h, i) => obj[h] = row[i]);
            return obj;
        });
    }
    return data; // Retrocompatibilidade
}

// Gerenciador de Cache (localStorage)
const CacheAPI = {
    get: function (key) {
        const stored = localStorage.getItem(key);
        if (stored) {
            try { return JSON.parse(stored); } catch (e) { return null; }
        }
        return null;
    },
    set: function (key, data) {
        localStorage.setItem(key, JSON.stringify(data));
    },
    clear: function (key) {
        localStorage.removeItem(key);
    }
};

// Prevenir Duplo Clique e Adicionar Spinner
async function execWithSpinner(btnElement, asyncFunc) {
    if (!btnElement) {
        await asyncFunc();
        return;
    }
    const originalText = btnElement.innerHTML;
    btnElement.disabled = true;
    btnElement.style.opacity = '0.7';
    btnElement.style.cursor = 'not-allowed';
    btnElement.innerHTML = `<svg style="animation: spin 1s linear infinite; height: 1.25rem; width: 1.25rem; margin-right: 0.5rem; display: inline-block;" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg> Aguarde...`;

    try {
        await asyncFunc();
    } finally {
        btnElement.disabled = false;
        btnElement.style.opacity = '1';
        btnElement.style.cursor = 'pointer';
        btnElement.innerHTML = originalText;
    }
}

// ── Registro do Service Worker (PWA) ────────────────────────
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then(reg => console.log('Service Worker registrado com sucesso!', reg.scope))
            .catch(err => console.error('Falha ao registrar Service Worker:', err));
    });
}
