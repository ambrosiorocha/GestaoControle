// relatorios.js — Dashboard Estratégico (Pro/Premium)
// Requer: config.js, auth.js, main.js (parseCompactData)

let _privacyOn = false;

document.addEventListener('DOMContentLoaded', function () {
    // Aguarda Auth.init (disparado pelo main.js)
    // A chamada a carregarRelatorios é feita via Auth callback no main.js
    // Mas para páginas que carregam o script diretamente, ouvimos aqui também.
    if (!window.MASTER_WEBHOOK_URL) {
        console.error('URL da Mestra não configurada.');
        return;
    }

    // Checa plano após Auth estar pronto
    _verificarAcessoECarregar();
});

// ── Verificação de Plano ────────────────────────────────────────
let _authCheckAttempts = 0;
function _verificarAcessoECarregar() {
    _authCheckAttempts++;
    // Aguarda o Auth estar inicializado (pode já estar se main.js rodou antes)
    // Após 10s sem login, redireciona para home
    if (typeof Auth === 'undefined' || !Auth.isLoggedIn()) {
        if (_authCheckAttempts > 100) { window.location.href = 'index.html'; return; }
        setTimeout(_verificarAcessoECarregar, 100);
        return;
    }


    // Bloqueia APENAS se for plano Básico E não for Admin
    const isBasico = Auth.isPlanBasico();
    const isAdmin = typeof Auth.isAdmin === 'function' && Auth.isAdmin();
    const plan = Auth.getPlan();

    // Atualiza badge de plano
    const badgeWrap = document.getElementById('planBadgeWrap');
    if (badgeWrap) {
        if (plan.toLowerCase() === 'premium') {
            badgeWrap.innerHTML = '<span class="plan-badge premium">⭐ Premium</span>';
        } else if (!isBasico || isAdmin) {
            badgeWrap.innerHTML = '<span class="plan-badge pro">🚀 Pro</span>';
        } else {
            badgeWrap.innerHTML = '<span style="font-size:0.72rem;color:#64748b;">Plano Básico</span>';
        }
    }

    if (isBasico && !isAdmin) {
        // MATA-LOADING AGRESSIVO: Força o sumiço de qualquer indicador de carregamentos
        ['loading', 'planBadgeWrap'].forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.style.setProperty('display', 'none', 'important');
                el.classList.add('hidden');
            }
        });

        const wall = document.getElementById('upgradeWall');
        if (wall) {
            wall.style.setProperty('display', 'flex', 'important');
        }

        // Garante que o conteúdo real do relatório continue invisível
        const content = document.getElementById('report-content');
        if (content) content.classList.add('hidden');

        return;
    }

    // Pro ou Premium — garante que a Upgrade Wall está oculta
    const wallEl = document.getElementById('upgradeWall');
    if (wallEl) wallEl.style.setProperty('display', 'none', 'important');
    const wallContainer = document.querySelector('.upgrade-wall-container');
    if (wallContainer) wallContainer.style.setProperty('display', 'none', 'important');

    // Pro ou Premium — exibe botão de privacidade e carrega dados
    const privBtn = document.getElementById('privacyToggle');
    if (privBtn) privBtn.style.display = 'flex';

    // Premium: exibe card de lucro
    if (plan.toLowerCase() === 'premium') {
        const kpiLucro = document.getElementById('kpiLucroWrap');
        if (kpiLucro) kpiLucro.style.display = '';
    }

    carregarRelatorios();
}

// ── Status Message ──────────────────────────────────────────────
function exibirStatus(resposta) {
    var statusMessage = document.getElementById('statusMessage');
    statusMessage.textContent = resposta.mensagem;
    statusMessage.className = '';
    if (resposta.status) {
        statusMessage.classList.add(resposta.status);
    }
    statusMessage.style.display = 'block';
    setTimeout(function () {
        statusMessage.style.display = 'none';
    }, 5000);
}

// ── Carregar Dados ──────────────────────────────────────────────
async function carregarRelatorios() {
    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterDadosRelatorios', spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await response.json();

        document.getElementById('loading').classList.add('hidden');
        document.getElementById('report-content').classList.remove('hidden');

        if (data.status === 'sucesso' && data.dados) {
            const vendas = parseCompactData(data.dados);
            renderizarKPIs(vendas);
            renderizarGraficoVendasPorMes(vendas);
            renderizarGraficoVendasPorProduto(vendas);
        } else {
            exibirStatus({ status: 'error', mensagem: 'Nenhuma venda encontrada para gerar relatórios.' });
        }
    } catch (error) {
        document.getElementById('loading').classList.add('hidden');
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar dados dos relatórios: ' + error.message });
    }
}

// ── KPIs ────────────────────────────────────────────────────────
function renderizarKPIs(vendas) {
    const totalVendas = vendas.reduce((acc, venda) => {
        let val = venda['Total com Desconto'];
        if (typeof val === 'string') val = val.replace(/[R$\s\.]/g, '').replace(',', '.');
        return acc + (parseFloat(val) || 0);
    }, 0);
    const totalTransacoes = vendas.length;
    const ticketMedio = totalTransacoes > 0 ? totalVendas / totalTransacoes : 0;

    document.getElementById('totalVendas').textContent = formatCurrencyBRL(totalVendas);
    document.getElementById('totalTransacoes').textContent = totalTransacoes;
    document.getElementById('ticketMedio').textContent = formatCurrencyBRL(ticketMedio);

    // Lucro estimado (Premium apenas — usa campo Custo se disponível)
    const lucroEl = document.getElementById('lucroEstimado');
    if (lucroEl) {
        const totalCusto = vendas.reduce((acc, v) => {
            let c = v['Custo Total'] || v['Custo'] || 0;
            if (typeof c === 'string') c = c.replace(/[R$\s\.]/g, '').replace(',', '.');
            return acc + (parseFloat(c) || 0);
        }, 0);
        const lucro = totalVendas - totalCusto;
        lucroEl.textContent = formatCurrencyBRL(lucro);
    }
}

// ── Extrai "MM/AAAA" de qualquer formato de data ────────────────
function extrairMesAno(valorData) {
    if (!valorData) return null;
    if (valorData instanceof Date && !isNaN(valorData)) {
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        return `${mes}/${valorData.getFullYear()}`;
    }
    const str = String(valorData).trim();
    const matchBr = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (matchBr) return `${matchBr[2].padStart(2, '0')}/${matchBr[3]}`;
    const d = new Date(str);
    if (!isNaN(d)) return `${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
    return null;
}

// ── Gráfico: Vendas por Mês ─────────────────────────────────────
function renderizarGraficoVendasPorMes(vendas) {
    const vendasPorMes = vendas.reduce((acc, venda) => {
        const mesAno = extrairMesAno(venda.Data || venda['Data']);
        if (!mesAno) return acc;
        let val = venda['Total com Desconto'];
        if (typeof val === 'string') val = val.replace(/[R$\s\.]/g, '').replace(',', '.');
        acc[mesAno] = (acc[mesAno] || 0) + (parseFloat(val) || 0);
        return acc;
    }, {});

    const labels = Object.keys(vendasPorMes).sort((a, b) => {
        const [ma, ya] = a.split('/');
        const [mb, yb] = b.split('/');
        return new Date(`${ya}-${ma}-01`) - new Date(`${yb}-${mb}-01`);
    });
    const dados = labels.map(mes => vendasPorMes[mes]);

    const ctx = document.getElementById('vendasPorMesChart').getContext('2d');
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Receita (R$)',
                data: dados,
                backgroundColor: 'rgba(22, 163, 74, 0.7)',
                borderColor: 'rgba(21, 128, 61, 1)',
                borderWidth: 1,
                borderRadius: 5
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: v => 'R$ ' + v.toLocaleString('pt-BR')
                    },
                    grid: { color: '#f1f5f9' }
                },
                x: { grid: { display: false } }
            }
        }
    });
}

// ── Gráfico: Vendas por Produto ─────────────────────────────────
function renderizarGraficoVendasPorProduto(vendas) {
    const vendasPorProduto = vendas.reduce((acc, venda) => {
        const produto = venda.Itens || venda.itens || venda['Produto'] || 'Outros';
        let val = venda['Total com Desconto'];
        if (typeof val === 'string') val = val.replace(/[R$\s\.]/g, '').replace(',', '.');
        acc[produto] = (acc[produto] || 0) + (parseFloat(val) || 0);
        return acc;
    }, {});

    // Top 7 produtos para legibilidade
    const sorted = Object.entries(vendasPorProduto).sort((a, b) => b[1] - a[1]).slice(0, 7);
    const labels = sorted.map(e => e[0]);
    const dados = sorted.map(e => e[1]);

    const COLORS = ['#16a34a', '#2563eb', '#0d9488', '#d97706', '#9333ea', '#e11d48', '#0ea5e9'];

    const ctx = document.getElementById('vendasPorProdutoChart').getContext('2d');
    new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                data: dados,
                backgroundColor: COLORS,
                hoverOffset: 8,
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: { font: { size: 11 }, padding: 14 }
                }
            }
        }
    });
}

// ── Privacy Toggle ──────────────────────────────────────────────
window.togglePrivacy = function () {
    _privacyOn = !_privacyOn;
    const btn = document.getElementById('privacyToggle');
    const label = document.getElementById('privacyLabel');
    const eyeIcon = document.getElementById('eyeIcon');
    const kpiRow = document.getElementById('kpiRow');

    if (_privacyOn) {
        kpiRow.classList.add('privacy-mode');
        btn.classList.add('active');
        label.textContent = 'Exibir valores';
        eyeIcon.innerHTML = `<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line>`;
    } else {
        kpiRow.classList.remove('privacy-mode');
        btn.classList.remove('active');
        label.textContent = 'Ocultar valores';
        eyeIcon.innerHTML = `<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle>`;
    }
};
