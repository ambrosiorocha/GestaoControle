// ============================================================
// AUTH.JS — Autenticação, Primeiro Acesso e Controle de Acesso
// Depende de: config.js (window.SCRIPT_URL)
// ============================================================

// ⚙️ URL do webhook da Planilha Mestra (Dual Dispatch)
// Preencha com a URL do seu Apps Script da Mestra.
const MASTER_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbzcgNx583XfUKPzJ1Odi28vQeothOb3ned1pSJeXx8Du11IXEoTmbebthTVJPrkkrk0Bg/exec';

window.Auth = (function () {
    const K = { user: 'sv_user', nivel: 'sv_nivel', ts: 'sv_ts', plano: 'sv_plano', perm: 'sv_permissoes' };
    const SESSION_MS = 8 * 3600 * 1000;
    let _cb = null;

    function getUser() { return localStorage.getItem(K.user) || ''; }
    function getNivel() { return localStorage.getItem(K.nivel) || 'Operador'; }
    function isAdmin() {
        const n = getNivel().toLowerCase();
        return n === 'admin' || n === 'administrador';
    }
    function isLoggedIn() {
        if (!getUser()) return false;
        const ts = parseInt(localStorage.getItem(K.ts) || '0');
        return (Date.now() - ts) < SESSION_MS;
    }
    function saveSession(nome, nivel, plano, permissoes, empresa) {
        localStorage.setItem(K.user, nome);
        localStorage.setItem(K.nivel, nivel);
        localStorage.setItem(K.plano, plano || 'Básico');
        localStorage.setItem(K.perm, JSON.stringify(permissoes || {}));
        if (empresa) localStorage.setItem('sv_empresa', empresa);
        localStorage.setItem(K.ts, Date.now().toString());
    }

    function getEmpresa() { return localStorage.getItem('sv_empresa') || 'Gestão&Controle'; }

    function getPlan() { return localStorage.getItem(K.plano) || 'Pro'; }
    function isPlanBasico() {
        const p = getPlan().toLowerCase();
        return p === 'básico' || p === 'basico';
    }

    // Permissões de operador
    function getPermissoes() {
        if (isAdmin()) return { relatorios: true, fiado: true, visaoDono: true };
        try { return JSON.parse(localStorage.getItem(K.perm) || '{}'); } catch (e) { return {}; }
    }
    function podeVerRelatorios() { return isAdmin() || getPermissoes().relatorios !== false; }
    function podeVenderFiado() { return !isPlanBasico() && (isAdmin() || getPermissoes().fiado !== false); }
    function podeVerCusto() { return !isPlanBasico() && (isAdmin() || getPermissoes().visaoDono === true); }

    function logout() {
        [K.user, K.nivel, K.ts, K.plano, K.perm].forEach(k => localStorage.removeItem(k));
        window.location.href = 'index.html';
    }

    // Chame nas páginas restritas por nível (Financeiro)
    function requireAdmin() {
        if (!isLoggedIn() || !isAdmin()) window.location.href = 'index.html';
    }

    // Chame nas páginas restritas por plano
    function requirePlan(minPlan) {
        if (!isLoggedIn()) { window.location.href = 'index.html'; return; }
        // Relatórios exibe upgrade wall internamente
    }

    // Aplica restrições de UI baseadas no nível
    function applyUI() {
        if (!isAdmin()) {
            document.querySelectorAll('[data-admin-only]').forEach(el => { el.style.display = 'none'; });
            document.querySelectorAll('[data-admin-btn]').forEach(el => { el.style.display = 'none'; el.disabled = true; });
        } else {
            document.querySelectorAll('[data-admin-only]').forEach(el => { el.style.display = ''; });
            document.querySelectorAll('[data-admin-btn]').forEach(el => { el.style.display = ''; el.disabled = false; });
        }
    }

    function updateBadge() {
        const el = document.getElementById('userBadge');
        if (el) {
            const nivel = getNivel();
            const plan = getPlan();
            const planStyles = {
                'básico': { bg: '#dcfce7', color: '#15803d', label: '🟢 Básico' },
                'basico': { bg: '#dcfce7', color: '#15803d', label: '🟢 Básico' },
                'pro': { bg: '#ede9fe', color: '#7c3aed', label: '🚀 Pro' },
                'premium': { bg: '#fef3c7', color: '#92400e', label: '⭐ Premium' }
            };
            const ps = planStyles[plan.toLowerCase()] || planStyles['pro'];
            el.innerHTML = `
                <span style="font-weight:600;display:block;">${getUser()}</span>
                <span style="font-size:0.67rem;opacity:0.75;display:block;margin-bottom:0.3rem;">${nivel}</span>
                <span style="display:inline-block;font-size:0.65rem;font-weight:700;padding:0.18rem 0.55rem;
                    border-radius:999px;background:${ps.bg};color:${ps.color};letter-spacing:0.03em;line-height:1.4;">
                    ${ps.label}
                </span>`;
        }
    }

    // ── CSS compartilhado para os overlays ──────────────────────
    const _sharedInputStyle = `width:100%;padding:0.7rem 0.75rem;border:1.5px solid #d1d5db;border-radius:0.5rem;
        font-size:0.9rem;box-sizing:border-box;font-family:'Poppins',sans-serif;outline:none;
        transition:border-color 0.2s;`;
    const _sharedLabelStyle = `font-size:0.82rem;font-weight:600;display:block;margin-bottom:0.3rem;color:#374151;`;

    // ── Confirmação do login ─────────────────────────────────────
    function _doLogin() {
        const inp = document.getElementById('loginNome');
        const pwd = document.getElementById('loginSenha');
        const errEl = document.getElementById('loginError');
        const btn = document.getElementById('loginBtn');
        const nome = inp ? inp.value.trim() : '';
        const senha = pwd ? pwd.value : '';

        if (!nome) { errEl.textContent = 'Informe seu login.'; errEl.style.display = 'block'; return; }
        if (!senha) { errEl.textContent = 'Informe sua senha.'; errEl.style.display = 'block'; return; }

        btn.disabled = true;
        btn.innerHTML = `<svg style="animation:spin 1s linear infinite;height:1rem;width:1rem;display:inline-block;vertical-align:middle;margin-right:6px"
            xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
            <circle opacity=".25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
            <path opacity=".75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>Verificando...`;
        errEl.style.display = 'none';

        fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'autenticarOperador', data: { nome, senha } })
        })
            .then(r => r.text()).then(txt => {
                try { return JSON.parse(txt); }
                catch (e) { throw new Error('Resposta inválida do servidor.'); }
            })
            .then(data => {
                if (data.status === 'sucesso') {
                    saveSession(data.nome, data.nivel, data.plano, data.permissoes, data.empresa);
                    const ov = document.getElementById('loginOverlay');
                    if (ov) ov.remove();
                    if (_cb) _cb();
                    updateBadge();

                    // Dispara o auto-registro na mestra em background
                    fetch(window.SCRIPT_URL, {
                        method: 'POST',
                        body: JSON.stringify({ action: 'registrarMestra', data: { nome: data.nome, empresa: data.empresa, whatsapp: "Ler de Configurações" } })
                    }).catch(() => console.log('Registro background pendente.'));

                } else {
                    errEl.textContent = data.mensagem || 'Usuário ou senha incorretos.';
                    errEl.style.display = 'block';
                    btn.disabled = false;
                    btn.innerHTML = '🔑 Entrar';
                }
            })
            .catch(err => {
                errEl.textContent = 'Falha na conexão. Tente novamente.';
                errEl.style.display = 'block';
                btn.disabled = false;
                btn.innerHTML = '🔑 Entrar';
            });
    }

    // ── Verificar se é o primeiro acesso (sem admins cadastrados) ──
    function _verificarPrimeiroAcesso(callback) {
        if (!window.SCRIPT_URL) { callback(false); return; }
        fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'verificarPrimeiroAcesso' })
        })
            .then(r => r.text()).then(txt => {
                try { return JSON.parse(txt); }
                catch (e) { return { primeiroAcesso: false }; }
            })
            .then(data => callback(data.primeiroAcesso === true))
            .catch(() => callback(false));
    }

    // ── Salvar primeiro acesso (Dual Dispatch) ────────────────
    function _doFirstAccess() {
        const nomeCompleto = document.getElementById('faNomeCompleto')?.value.trim();
        const empresa = document.getElementById('faEmpresa')?.value.trim();
        const telefone = document.getElementById('faTelefone')?.value.trim();
        const login = document.getElementById('faLogin')?.value.trim();
        const senha = document.getElementById('faSenha')?.value;
        const errEl = document.getElementById('faError');
        const btn = document.getElementById('faBtnSalvar');

        if (!nomeCompleto) { errEl.textContent = 'Informe seu nome completo.'; errEl.style.display = 'block'; return; }
        if (!empresa) { errEl.textContent = 'Informe o nome da empresa.'; errEl.style.display = 'block'; return; }
        if (!telefone || telefone.length < 10) { errEl.textContent = 'Informe um WhatsApp válido (10 ou 11 números com DDD).'; errEl.style.display = 'block'; return; }
        if (!login) { errEl.textContent = 'Defina um login de acesso.'; errEl.style.display = 'block'; return; }
        if (!senha || senha.length < 4) { errEl.textContent = 'A senha deve ter ao menos 4 caracteres.'; errEl.style.display = 'block'; return; }

        btn.disabled = true;
        btn.textContent = 'Salvando...';
        errEl.style.display = 'none';

        // ── Payload compartilhado ──────────────────────────────
        const now = new Date();
        const exp = new Date(); exp.setDate(exp.getDate() + 30);
        const fmt = d => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;

        const sharedMeta = {
            nome: nomeCompleto,
            empresa: empresa,
            whatsapp: telefone,
            scriptUrl: window.SCRIPT_URL || '',
            registro: now.toLocaleString('pt-BR'),
            plano: 'Básico',
            ativacao: fmt(now),
            expiracao: fmt(exp)
        };

        // ── 1. Disparo Principal → SCRIPT_URL do cliente ───────
        fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'primeiroAcesso',
                data: { nomeCompleto, empresa, whatsapp: telefone, login, senha }
            })
        })
            .then(r => r.text())
            .then(txt => {
                try { return JSON.parse(txt); }
                catch (e) { throw new Error('Resposta inválida do servidor.'); }
            })
            .then(data => {
                if (data.status === 'sucesso') {
                    // ── 2. Disparo Paralelo → Planilha Mestra (fire-and-forget) ──
                    if (MASTER_WEBHOOK_URL) {
                        const masterPayload = {
                            nome: empresa,
                            usuario: nomeCompleto,
                            whatsapp: telefone,
                            scriptUrl: data.scriptUrl || window.SCRIPT_URL || '',
                            spreadsheetId: data.spreadsheetId || '',
                            spreadsheetUrl: data.spreadsheetUrl || '',
                            registro: sharedMeta.registro,
                            plano: sharedMeta.plano,
                            ativacao: sharedMeta.ativacao,
                            expiracao: sharedMeta.expiracao
                        };
                        fetch(MASTER_WEBHOOK_URL, {
                            method: 'POST',
                            body: JSON.stringify(masterPayload),
                            headers: { 'Content-Type': 'application/json' }
                        }).catch(() => { });
                    }

                    // Fecha o modal e abre o login
                    const fa = document.getElementById('firstAccessOverlay');
                    if (fa) fa.remove();
                    showModal(_cb);
                    setTimeout(() => {
                        const loginMsg = document.getElementById('loginSuccessMsg');
                        if (loginMsg) { loginMsg.textContent = '✅ Conta criada! Faça seu login abaixo.'; loginMsg.style.display = 'block'; }
                        const loginNome = document.getElementById('loginNome');
                        if (loginNome) loginNome.value = login;
                    }, 100);
                } else {
                    errEl.textContent = data.mensagem || 'Erro ao salvar. Tente novamente.';
                    errEl.style.display = 'block';
                    btn.disabled = false;
                    btn.textContent = '🚀 Criar minha conta';
                }
            })
            .catch(() => {
                errEl.textContent = 'Falha na conexão. Tente novamente.';
                errEl.style.display = 'block';
                btn.disabled = false;
                btn.textContent = '🚀 Criar minha conta';
            });
    }

    // ── Modal de Primeiro Acesso ──────────────────────────────────
    function showFirstAccessModal() {
        const ov = document.createElement('div');
        ov.id = 'firstAccessOverlay';
        ov.style.cssText = 'position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;padding:1rem;background:linear-gradient(135deg,#0f172a 0%,#16a34a 100%);';
        ov.innerHTML = `
        <div style="background:white;border-radius:1.5rem;padding:2rem 1.75rem;width:100%;max-width:400px;box-shadow:0 30px 80px rgba(0,0,0,0.5);animation:fadeInUp 0.35s ease;">
            <style>@keyframes fadeInUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
            @keyframes spin{to{transform:rotate(360deg)}}</style>

            <!-- Ícone de boas-vindas -->
            <div style="text-align:center;margin-bottom:1.25rem;">
                <div style="width:64px;height:64px;border-radius:18px;background:linear-gradient(135deg,#16a34a,#15803d);
                    display:flex;align-items:center;justify-content:center;margin:0 auto 0.75rem;box-shadow:0 8px 24px rgba(22,163,74,0.35);">
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                    </svg>
                </div>
                <h2 style="font-size:1.4rem;font-weight:800;color:#1e293b;margin:0 0 0.2rem;">Bem-vindo ao</h2>
                <p style="font-size:1.55rem;font-weight:900;color:#16a34a;margin:0;letter-spacing:-0.03em;">Gestão&amp;Controle</p>
                <p style="font-size:0.78rem;color:#64748b;margin-top:0.4rem;">Configure sua conta de administrador para começar</p>
            </div>

            <!-- Erro -->
            <div id="faError" style="display:none;background:#fee2e2;color:#991b1b;border-radius:0.5rem;padding:0.55rem 0.75rem;font-size:0.8rem;margin-bottom:0.9rem;"></div>

            <!-- Formulário -->
            <div style="display:flex;flex-direction:column;gap:0.75rem;">
                <div>
                    <label style="${_sharedLabelStyle}">👤 Seu nome completo</label>
                    <input id="faNomeCompleto" type="text" placeholder="Ex: João da Silva" autocomplete="name"
                        style="${_sharedInputStyle}margin-bottom:0;">
                </div>
                <div>
                    <label style="${_sharedLabelStyle}">🏢 Nome da empresa</label>
                    <input id="faEmpresa" type="text" placeholder="Ex: Silva Distribuidora Ltda"
                        style="${_sharedInputStyle}margin-bottom:0;">
                </div>
                <div>
                    <label style="${_sharedLabelStyle}">📱 WhatsApp</label>
                    <input id="faTelefone" type="text" placeholder="Ex: 81999999999" maxlength="11" oninput="this.value = this.value.replace(/[^0-9]/g, '')"
                        style="${_sharedInputStyle}margin-bottom:0;">
                </div>
                <div>
                    <label style="${_sharedLabelStyle}">🔐 Login de acesso</label>
                    <input id="faLogin" type="text" placeholder="Ex: joao.admin" autocomplete="username"
                        style="${_sharedInputStyle}margin-bottom:0;">
                </div>
                <div>
                    <label style="${_sharedLabelStyle}">🔑 Senha (mínimo 4 caracteres)</label>
                    <input id="faSenha" type="password" placeholder="••••••••" autocomplete="new-password"
                        style="${_sharedInputStyle}margin-bottom:0;">
                </div>
            </div>

            <!-- Aviso de plano -->
            <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:0.625rem;padding:0.65rem 0.85rem;
                margin:1rem 0;display:flex;gap:0.6rem;align-items:flex-start;">
                <span style="font-size:1rem;flex-shrink:0;">🟢</span>
                <p style="font-size:0.75rem;color:#166534;margin:0;line-height:1.5;">
                    Conta criada com <strong>Plano Básico</strong>. Para adicionar funcionários ou acessar
                    relatórios avançados, fale com nosso suporte para fazer o upgrade.
                </p>
            </div>

            <button id="faBtnSalvar" onclick="Auth._doFirstAccess()"
                style="width:100%;padding:0.9rem;background:linear-gradient(135deg,#16a34a,#15803d);color:white;
                border:none;border-radius:0.75rem;font-weight:700;font-size:0.95rem;cursor:pointer;
                font-family:'Poppins',sans-serif;box-shadow:0 4px 14px rgba(22,163,74,0.4);transition:transform 0.15s;">
                🚀 Criar minha conta
            </button>
            <p style="text-align:center;font-size:0.72rem;color:#94a3b8;margin-top:0.75rem;margin-bottom:0;">
                Já tem acesso? <a href="#" onclick="document.getElementById('firstAccessOverlay').remove();Auth.showModal(null);"
                    style="color:#16a34a;font-weight:600;text-decoration:none;">Fazer login</a>
            </p>
        </div>`;
        document.body.appendChild(ov);
        // Enter no último campo
        document.getElementById('faSenha').addEventListener('keydown', e => {
            if (e.key === 'Enter') Auth._doFirstAccess();
        });
        // Focus style nos inputs
        ov.querySelectorAll('input').forEach(inp => {
            inp.addEventListener('focus', () => inp.style.borderColor = '#16a34a');
            inp.addEventListener('blur', () => inp.style.borderColor = '#d1d5db');
        });
    }

    // ── Modal de Login ────────────────────────────────────────────
    function showModal(callback) {
        if (callback !== null) _cb = callback; // null = chamada interna do primeiro acesso
        const ov = document.createElement('div');
        ov.id = 'loginOverlay';
        ov.style.cssText = 'position:fixed;inset:0;z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:1rem;overflow-y:auto;background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);';
        ov.innerHTML = `
        <style>@keyframes fadeInUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
        @keyframes spin{to{transform:rotate(360deg)}}</style>
        <div style="background:white;border-radius:1.5rem;padding:2rem 1.75rem;width:100%;max-width:360px;
            box-shadow:0 25px 60px rgba(0,0,0,0.5);animation:fadeInUp 0.3s ease;margin:auto;">

            <!-- Logo + Brand -->
            <div style="display:flex;justify-content:center;align-items:center;gap:0.75rem;margin-bottom:0.4rem;">
                <img src="assets/logo.png" alt="Logo" style="height:52px;object-fit:contain;">
                <div style="text-align:left;line-height:1.1;">
                    <div style="font-size:1.7rem;font-weight:800;color:#16a34a;letter-spacing:-0.03em;">Gestão</div>
                    <div style="font-size:1.3rem;font-weight:600;color:#334155;letter-spacing:-0.02em;">&amp;Controle</div>
                </div>
            </div>
            <p style="text-align:center;font-size:0.78rem;color:#64748b;margin-bottom:1.5rem;">Identifique-se para continuar</p>

            <!-- Mensagem de sucesso (pós-cadastro) -->
            <div id="loginSuccessMsg" style="display:none;background:#f0fdf4;color:#166534;border:1px solid #bbf7d0;
                border-radius:0.5rem;padding:0.55rem 0.75rem;font-size:0.8rem;margin-bottom:0.9rem;font-weight:500;"></div>

            <!-- Erro -->
            <div id="loginError" style="display:none;background:#fee2e2;color:#991b1b;border-radius:0.5rem;
                padding:0.6rem 0.75rem;font-size:0.82rem;margin-bottom:0.9rem;"></div>

            <!-- Campos -->
            <label style="${_sharedLabelStyle}">Login</label>
            <input id="loginNome" type="text" placeholder="Seu login de acesso" autocomplete="username"
                style="${_sharedInputStyle}margin-bottom:0.85rem;">

            <label style="${_sharedLabelStyle}">Senha</label>
            <input id="loginSenha" type="password" placeholder="••••••••" autocomplete="current-password"
                style="${_sharedInputStyle}margin-bottom:1.25rem;">

            <!-- Botão entrar -->
            <button id="loginBtn" onclick="Auth._doLogin()"
                style="width:100%;padding:0.85rem;background:linear-gradient(135deg,#16a34a,#15803d);color:white;
                border:none;border-radius:0.625rem;font-weight:700;font-size:0.95rem;cursor:pointer;
                font-family:'Poppins',sans-serif;box-shadow:0 4px 12px rgba(22,163,74,0.35);">
                🔑 Entrar
            </button>

            <!-- Primeiro acesso -->
            <p style="text-align:center;font-size:0.74rem;color:#94a3b8;margin-top:1rem;margin-bottom:0;position:relative;z-index:10000;">
                Primeiro acesso?
                <a href="#" id="firstAccessLink"
                    style="color:#16a34a;font-weight:700;text-decoration:none;display:inline-block;padding:0.2rem 0.4rem;">
                    Clique aqui
                </a>
            </p>
        </div>`;
        document.body.appendChild(ov);

        // Eventos
        document.getElementById('loginSenha').addEventListener('keydown', e => {
            if (e.key === 'Enter') Auth._doLogin();
        });
        document.getElementById('firstAccessLink').addEventListener('click', e => {
            e.preventDefault();
            // Não remove o overlay ainda — mantém cobrindo o body durante o fetch
            _verificarPrimeiroAcesso(isPrimeiro => {
                if (isPrimeiro) {
                    // 1º adiciona o novo overlay, 2º remove o antigo
                    // → nunca há um frame sem cobertura
                    showFirstAccessModal();
                    ov.remove();
                } else {
                    // Já tem admins — mostra aviso
                    showModal(null);
                    ov.remove();
                    setTimeout(() => {
                        const err = document.getElementById('loginError');
                        if (err) {
                            err.textContent = 'O sistema já está configurado. Peça ao administrador para criar seu acesso.';
                            err.style.display = 'block';
                        }
                    }, 100);
                }
            });
        });
        // Focus style
        ov.querySelectorAll('input').forEach(inp => {
            inp.addEventListener('focus', () => inp.style.borderColor = '#16a34a');
            inp.addEventListener('blur', () => inp.style.borderColor = '#d1d5db');
        });
    }

    // ── Revelar o body (anti-flicker) ────────────────────────────
    function _revealBody() {
        // visibility:snap — sem transição de opacidade para evitar que o
        // dashboard "vaze" por baixo do overlay durante frames intermediários.
        document.body.style.visibility = 'visible';
    }

    // ── Ponto de entrada ──────────────────────────────────────────
    function init(callback) {
        if (isLoggedIn()) {
            applyUI();
            updateBadge();
            _revealBody();
            if (callback) callback();
        } else {
            showModal(callback);
            _revealBody();
        }
    }

    return {
        getUser, getNivel, getPlan, getEmpresa, isPlanBasico, isAdmin, isLoggedIn,
        getPermissoes, podeVerRelatorios, podeVenderFiado, podeVerCusto,
        logout, requireAdmin, requirePlan, applyUI, updateBadge,
        init, showModal, _doLogin, _doFirstAccess
    };
})();
