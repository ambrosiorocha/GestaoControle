// ================================
// ESTADO GLOBAL
// ================================
let produtos = [];
let carrinho = [];
let formaPagamentoSelecionada = '';
let vendaEditandoId = null; // ID da venda Pendente em edição

// ================================
// INICIALIZAÇÃO
// ================================
document.addEventListener('DOMContentLoaded', function () {
    if (window.MASTER_WEBHOOK_URL === '' || window.MASTER_WEBHOOK_URL.includes('COLE_AQUI')) {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.MASTER_WEBHOOK_URL no config.js.' });
        return;
    }
    carregarProdutos();
    carregarClientes();
    carregarOperadores();
    carregarHistoricoVendas();

    document.getElementById('produto').addEventListener('change', atualizarPrecoUnitario);
    document.getElementById('quantidade').addEventListener('input', () => setTimeout(atualizarSubtotalItem, 10));
    document.getElementById('descontoItemPct').addEventListener('input', () => setTimeout(() => sincronizarDesconto('pct'), 10));
    document.getElementById('descontoItemReais').addEventListener('input', () => setTimeout(() => sincronizarDesconto('reais'), 10));

    document.querySelectorAll('.pgto-btn').forEach(btn => {
        btn.addEventListener('click', function () {
            if (this.dataset.planBlock) return; // Bloqueado por plano
            document.querySelectorAll('.pgto-btn').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            formaPagamentoSelecionada = this.dataset.pgto;
            aplicarPrazo(this);
            atualizarModalTotais();
        });
    });
    document.getElementById('descontoGeralModal').addEventListener('input', atualizarModalTotais);

    // Aplica restrições de plano após Auth estar pronto
    let _gateTries = 0;
    function tryApplyGate() {
        if (typeof Auth !== 'undefined' && Auth.isLoggedIn()) {
            aplicarGatePlanVendas();
        } else if (++_gateTries < 30) {
            setTimeout(tryApplyGate, 200);
        }
    }
    tryApplyGate();

    // Evento para fechar modal com tecla ESC
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') fecharSlideoverVenda();
    });
});

// ================================
// CONTROLE DO SLIDE-OVER (GAVETA)
// ================================
function abrirSlideoverVenda() {
    const overlay = document.getElementById('slideoverVenda');
    const container = document.getElementById('slideoverContainer');
    
    // Preparar estado limpo se não estiver editando
    if (!vendaEditandoId && carrinho.length === 0) {
        document.getElementById('produto').value = '';
        document.getElementById('quantidade').value = '';
        document.getElementById('descontoItemPct').value = '0';
        document.getElementById('descontoItemReais').value = '0,00';
        document.getElementById('precoUnitario').value = '';
        document.getElementById('subtotalItem').textContent = '—';
    }

    overlay.classList.add('open');
    container.classList.add('open');
    document.body.classList.add('slideover-open');
}

function fecharSlideoverVenda() {
    const overlay = document.getElementById('slideoverVenda');
    const container = document.getElementById('slideoverContainer');
    
    overlay.classList.remove('open');
    container.classList.remove('open');
    document.body.classList.remove('slideover-open');
}

function validarFechamentoSlideover(e) {
    if (e.target.id === 'slideoverVenda') {
        fecharSlideoverVenda();
    }
}

// ================================
// GATING DE PLANO — VENDAS

// ================================
function aplicarGatePlanVendas() {
    const isBasico = typeof Auth !== 'undefined' && Auth.isPlanBasico();
    if (!isBasico) return;
    const pgtosBloqueados = ['fiado', 'parcelado', 'boleto', 'crédito', 'credito'];
    document.querySelectorAll('.pgto-btn').forEach(btn => {
        const pgto = (btn.dataset.pgto || btn.textContent).trim().toLowerCase();
        if (pgtosBloqueados.some(b => pgto.includes(b))) {
            btn.dataset.planBlock = 'true';
            btn.style.opacity = '0.4';
            btn.style.cursor = 'not-allowed';
            btn.style.filter = 'grayscale(1)';
            btn.title = '🔒 Disponível no Plano Pro';
        }
    });
    document.querySelectorAll('[data-print-btn], .btn-imprimir, #btnImprimirCupom').forEach(el => {
        el.style.display = 'none';
    });
}

// ================================
// STATUS MESSAGE
// ================================
function exibirStatus(resposta) {
    const isSlideoverOpen = document.getElementById('slideoverVenda').classList.contains('open');
    const targetId = isSlideoverOpen ? 'statusMessageSlideover' : 'statusMessage';
    var el = document.getElementById(targetId);
    if (!el) el = document.getElementById('statusMessage');

    const isError = resposta.status === 'error' || resposta.status === 'erro';
    const icon = isError ? '🛑 ' : '✅ ';
    el.innerHTML = icon + resposta.mensagem;

    el.className = '';
    if (resposta.status) el.classList.add(resposta.status);

    if (isError && isSlideoverOpen) {
        el.classList.add('shake');
        setTimeout(() => el.classList.remove('shake'), 600);
    }

    el.style.display = 'block';
    setTimeout(() => el.style.display = 'none', 7000);
}

// ================================
// CARREGAR DADOS
// ================================
async function carregarProdutos() {
    const sel = document.getElementById('produto');
    sel.innerHTML = '<option value="">Carregando...</option>';

    const cached = CacheAPI.get('cache_produtos');
    if (cached) {
        produtos = cached;
        preencherSelectProdutos(sel, produtos);
        return;
    }

    try {
        const res = await fetch(window.MASTER_WEBHOOK_URL, { method: 'POST', body: JSON.stringify({ action: 'listarProdutos', spreadsheetId: window.SPREADSHEET_ID }) });
        const data = await res.json();
        produtos = parseCompactData(data.dados) || [];
        CacheAPI.set('cache_produtos', produtos);
        preencherSelectProdutos(sel, produtos);
    } catch (e) { sel.innerHTML = '<option value="">Erro ao carregar</option>'; }
}

function preencherSelectProdutos(sel, lista) {
    sel.innerHTML = '<option value="">Selecione um produto</option>';
    lista.forEach(p => {
        const estoque = parseFloat(p.Quantidade) || 0;
        const precoVal = parseCurrencyBRL(p['Preço_de_venda'] || p['Preço'] || p.Preco || 0);
        const opt = document.createElement('option');
        opt.value = p.Nome;
        opt.textContent = `${p.Nome}  (Estoque: ${estoque})`;
        opt.dataset.preco = precoVal;
        opt.dataset.estoque = estoque;
        sel.appendChild(opt);
    });
}

async function carregarOperadores() {
    const sel = document.getElementById('usuario');
    try {
        const res = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterOperadores', spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await res.json();
        if (data.status === 'sucesso' && Array.isArray(data.dados) && data.dados.length > 0) {
            sel.innerHTML = '';
            data.dados.forEach(item => {
                const opt = document.createElement('option');
                const val = typeof item === 'object' ? (item.nome || item.NOME || '') : item;
                opt.value = val;
                opt.textContent = val;
                sel.appendChild(opt);
            });
            // Reaplicar permissões do Auth após carregar os operadores
            if (typeof Auth !== 'undefined' && Auth.getUser()) {
                let found = Array.from(sel.options).some(o => o.value === Auth.getUser());
                if (!found) {
                    const opt = new Option(Auth.getUser(), Auth.getUser());
                    sel.insertBefore(opt, sel.firstChild);
                }
                sel.value = Auth.getUser();
                if (!Auth.isAdmin()) sel.disabled = true; // Operador não pode trocar
            }
        }
    } catch (e) { /* mantém opção padrão */ }
}

async function carregarClientes() {
    const sel = document.getElementById('cliente');
    sel.innerHTML = '<option value="Consumidor Interno">Consumidor Interno</option>';

    const cached = CacheAPI.get('cache_clientes');
    if (cached) {
        preencherSelectClientes(sel, cached);
        return;
    }

    try {
        const res = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterClientes', spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await res.json();
        if (data.status === 'sucesso' && data.dados) {
            const arr = parseCompactData(data.dados);
            CacheAPI.set('cache_clientes', arr);
            preencherSelectClientes(sel, arr);
        }
    } catch (e) { /* mantém Consumidor Interno */ }
}

function preencherSelectClientes(sel, lista) {
    if (!lista) return;
    lista.forEach(c => {
        const opt = document.createElement('option');
        const nome = c.nome || c.Nome || '';
        opt.value = nome;
        opt.textContent = nome;
        sel.appendChild(opt);
    });
}

// ================================
// DESCONTO E SUBTOTAL POR ITEM
// ================================
function getPrecoUnitario() {
    const sel = document.getElementById('produto');
    const opt = sel.options[sel.selectedIndex];
    return opt ? parseFloat(opt.dataset.preco) || 0 : 0;
}

function atualizarPrecoUnitario() {
    const preco = getPrecoUnitario();
    document.getElementById('precoUnitario').value = preco > 0 ? formatCurrencyBRL(preco) : 'R$ 0,00';
    document.getElementById('descontoItemPct').value = '0';
    document.getElementById('descontoItemReais').value = '0,00';
    atualizarSubtotalItem();
}

function sincronizarDesconto(origem) {
    const preco = getPrecoUnitario();
    const qtd = parseFloat(document.getElementById('quantidade').value) || 0;
    const bruto = preco * qtd;
    if (origem === 'pct') {
        const pct = parseFloat(document.getElementById('descontoItemPct').value) || 0;
        document.getElementById('descontoItemReais').value = bruto > 0 ? (bruto * pct / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00';
    } else {
        const reais = parseCurrencyBRL(document.getElementById('descontoItemReais').value) || 0;
        document.getElementById('descontoItemPct').value = bruto > 0 ? (reais / bruto * 100).toFixed(2) : '0';
    }
    atualizarSubtotalItem();
}

function atualizarSubtotalItem() {
    const preco = getPrecoUnitario();
    const qtd = parseFloat(document.getElementById('quantidade').value) || 0;
    const descPct = parseFloat(document.getElementById('descontoItemPct').value) || 0;
    const subtotal = Math.max(0, qtd * preco * (1 - descPct / 100));
    const el = document.getElementById('subtotalItem');
    if (qtd > 0 && preco > 0) {
        el.textContent = formatCurrencyBRL(subtotal);
        el.style.color = '#16a34a';
    } else {
        el.textContent = '—';
    }
}

// ================================
// CARRINHO
// ================================
function adicionarItemCarrinho() {
    const sel = document.getElementById('produto');
    const opt = sel.options[sel.selectedIndex];
    if (!sel.value) { exibirStatus({ status: 'error', mensagem: '⚠️ Selecione um produto.' }); return; }
    const quantidade = parseFloat(document.getElementById('quantidade').value);
    if (!quantidade || quantidade <= 0) { exibirStatus({ status: 'error', mensagem: '⚠️ Informe a quantidade.' }); return; }

    const descPct = parseFloat(document.getElementById('descontoItemPct').value) || 0;
    const estoqueDisp = parseFloat(opt.dataset.estoque) || 0;
    const jaNoCarrinho = carrinho.filter(i => i.nome === sel.value).reduce((s, i) => s + i.quantidade, 0);

    if (quantidade + jaNoCarrinho > estoqueDisp) {
        exibirStatus({ status: 'error', mensagem: `❌ Estoque insuficiente! Disponível: ${estoqueDisp - jaNoCarrinho}` });
        return;
    }

    const preco = parseFloat(opt.dataset.preco) || 0;
    const subtotal = Math.max(0, quantidade * preco * (1 - descPct / 100));
    const existente = carrinho.find(i => i.nome === sel.value && i.desconto === descPct);
    if (existente) {
        existente.quantidade += quantidade;
        existente.subtotal = Math.max(0, existente.quantidade * existente.preco * (1 - existente.desconto / 100));
    } else {
        carrinho.push({ nome: sel.value, preco, quantidade, desconto: descPct, subtotal });
    }

    sel.value = '';
    document.getElementById('quantidade').value = '';
    document.getElementById('descontoItemPct').value = '0';
    document.getElementById('descontoItemReais').value = '0';
    document.getElementById('precoUnitario').value = '';
    document.getElementById('subtotalItem').textContent = '—';
    renderizarCarrinho();
}

function renderizarCarrinho() {
    const tbody = document.getElementById('carrinhoItens');
    tbody.innerHTML = '';
    const editandoBadge = vendaEditandoId
        ? `<div style="background:#fef3c7;border:1px solid #fde68a;border-radius:0.5rem;padding:0.4rem 0.75rem;font-size:0.8rem;color:#92400e;margin-bottom:0.5rem;">✏️ Editando rascunho #${vendaEditandoId} &nbsp;<button onclick="cancelarEdicao()" style="border:none;background:none;color:#ef4444;cursor:pointer;font-size:0.8rem;">✕ cancelar edição</button></div>`
        : '';
    document.getElementById('editandoBadge').innerHTML = editandoBadge;

    const fmtBRL = v => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

    const setBtns = (enabled) => {
        ['btnFinalizar', 'btnRascunho'].forEach(id => {
            const b = document.getElementById(id);
            if (!b) return;
            b.disabled = !enabled;
            b.style.cursor = enabled ? 'pointer' : 'not-allowed';
            b.style.opacity = enabled ? '1' : '0.55';
        });
    };

    if (carrinho.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="carrinho-vazio">Nenhum item adicionado ainda.</td></tr>';
        document.getElementById('totalCarrinho').textContent = 'R$ 0,00';
        document.getElementById('qtdItensLabel').textContent = '0 item(ns)';
        setBtns(false);
        return;
    }

    let totalGeral = 0;
    carrinho.forEach((item, idx) => {
        totalGeral += item.subtotal;
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><a href="#" onclick="editarItemCarrinho(${idx}); return false;" style="color:#2563eb; text-decoration:underline;" title="Clique para editar esse item">${item.nome}</a></td>
            <td style="text-align:center;">${item.quantidade}</td>
            <td style="text-align:right;">${fmtBRL(item.subtotal)}</td>
            <td><button class="remove-item" onclick="removerItem(${idx})" title="Remover">✕</button></td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('totalCarrinho').textContent = fmtBRL(totalGeral);
    document.getElementById('qtdItensLabel').textContent = `${carrinho.length} item(ns)`;
    setBtns(true);
}


function removerItem(idx) { carrinho.splice(idx, 1); renderizarCarrinho(); }

function editarItemCarrinho(idx) {
    const item = carrinho[idx];
    document.getElementById('produto').value = item.nome;
    document.getElementById('quantidade').value = item.quantidade;
    document.getElementById('descontoItemPct').value = item.desconto;
    carrinho.splice(idx, 1);
    atualizarPrecoUnitario();
    sincronizarDesconto('pct');
    renderizarCarrinho();
}

function cancelarEdicao() {
    vendaEditandoId = null;
    carrinho = [];
    renderizarCarrinho();
    exibirStatus({ status: 'success', mensagem: 'Edição cancelada.' });
}

// ================================
// SALVAR RASCUNHO (Pendente — sem estoque, sem financeiro)
// ================================
async function salvarRascunho() {
    const btn = document.getElementById('btnRascunho');
    await execWithSpinner(btn, async () => {
        if (carrinho.length === 0) { exibirStatus({ status: 'error', mensagem: '⚠️ Adicione itens ao pedido.' }); return; }
        const payload = montarPayloadVenda();
        payload.formaPagamento = payload.formaPagamento || '-';
        payload.statusFinanceiro = 'Pendente';

        const action = vendaEditandoId ? 'finalizarPendente' : 'salvarRascunho';
        if (vendaEditandoId) payload.id = vendaEditandoId;

        try {
            const res = await fetch(window.MASTER_WEBHOOK_URL, {
                method: 'POST',
                body: JSON.stringify({ acao: action, action: action, data: payload, spreadsheetId: window.SPREADSHEET_ID })
            });
            const data = await res.json();
            exibirStatus(data);
            if (data.status === 'sucesso') {
                carrinho = [];
                vendaEditandoId = null;
                renderizarCarrinho();
                fecharSlideoverVenda();
                await carregarHistoricoVendas();
            }
        } catch (e) { exibirStatus({ status: 'error', mensagem: 'Erro: ' + e }); }
    });
}

// ================================
// MODAL FINALIZAR VENDA
// ================================
function abrirModal() {
    formaPagamentoSelecionada = '';
    document.querySelectorAll('.pgto-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('descontoGeralModal').value = '0';
    document.getElementById('prazoContainer').style.display = 'none';
    document.getElementById('prazoCustom').style.display = 'none';
    document.getElementById('vencimentoCustom').value = '';

    const caixaSelect = document.getElementById('caixaVenda');
    if (caixaSelect) {
        if (typeof Auth !== 'undefined' && Auth.isPlanBasico()) {
            caixaSelect.innerHTML = '<option value="Dinheiro">💵 Dinheiro</option>';
            caixaSelect.value = 'Dinheiro';
            caixaSelect.disabled = true;
        } else {
            caixaSelect.innerHTML = `
                <option value="Dinheiro">💵 Dinheiro</option>
                <option value="Conta Banco do Brasil">🏦 Conta Banco do Brasil</option>
                <option value="Conta Itaú">🏦 Conta Itaú</option>
                <option value="Conta Caixa">🏦 Conta Bradesco</option>
                <option value="Conta Nubank">🟣 Nubank Empresa</option>
            `;
            caixaSelect.disabled = false;
        }
    }

    atualizarModalTotais();
    document.getElementById('modalFinalizar').style.display = 'flex';
}
function fecharModal() { document.getElementById('modalFinalizar').style.display = 'none'; }

function aplicarPrazo(btn) {
    const prazo = btn.dataset.prazo;
    const container = document.getElementById('prazoContainer');
    const prazoInfo = document.getElementById('prazoInfo');
    const prazoCustom = document.getElementById('prazoCustom');
    const today = new Date();
    container.style.display = 'block';
    prazoCustom.style.display = 'none';
    if (prazo === '0') {
        prazoInfo.innerHTML = `<b style="color:#15803d">Pagamento imediato</b> — Venc: <b>${today.toLocaleDateString('pt-BR')}</b> | Status: <b>Pago</b>`;
        container.style.background = '#f0fdf4'; container.style.borderColor = '#bbf7d0';
    } else if (prazo === '30') {
        const v = new Date(today); v.setDate(v.getDate() + 30);
        prazoInfo.innerHTML = `Crédito — Venc.: <b>30 dias</b> (${v.toLocaleDateString('pt-BR')}) | Status: <b>Pendente</b>`;
        container.style.background = '#fff7ed'; container.style.borderColor = '#fed7aa';
    } else {
        prazoInfo.innerHTML = `<b>${formaPagamentoSelecionada}</b> — Informe o vencimento:`;
        container.style.background = '#fef3c7'; container.style.borderColor = '#fde68a';
        prazoCustom.style.display = 'block';
        const sug = new Date(today); sug.setDate(sug.getDate() + 30);
        document.getElementById('vencimentoCustom').value = sug.toISOString().substring(0, 10);
    }
}

async function calcularVencimentoStatus() {
    const btn = document.querySelector('.pgto-btn.active');
    if (!btn) return { vencimento: new Date().toLocaleDateString('pt-BR'), status: 'Pendente' };
    const prazo = btn.dataset.prazo;
    const today = new Date();
    if (prazo === '0') {
        return { vencimento: today.toLocaleDateString('pt-BR'), status: 'Pago' };
    } else if (prazo === '30') {
        const v = new Date(today); v.setDate(v.getDate() + 30);
        return { vencimento: v.toLocaleDateString('pt-BR'), status: 'Pendente' };
    } else {
        const raw = document.getElementById('vencimentoCustom').value;
        if (!raw) { await CustomModal.alert('Informe a data de vencimento.', 'OK'); return null; }
        const [y, m, d] = raw.split('-');
        return { vencimento: `${d}/${m}/${y}`, status: 'Pendente' };
    }
}

function atualizarModalTotais() {
    // Subtotal BRUTO (sem nenhum desconto)
    const subtotalBruto = carrinho.reduce((s, i) => s + (i.quantidade * i.preco), 0);
    // Subtotal já com descontos por item
    const subtotalComItens = carrinho.reduce((s, i) => s + i.subtotal, 0);
    // Desconto dos itens individualmente
    const descontoItens = subtotalBruto - subtotalComItens;
    // Desconto geral adicional no modal
    const descontoGeral = parseFloat(document.getElementById('descontoGeralModal').value) || 0;
    // Desconto total
    const descontoTotal = descontoItens + descontoGeral;
    const total = Math.max(0, subtotalBruto - descontoTotal);

    document.getElementById('modalSubtotal').textContent = formatCurrencyBRL(subtotalBruto);
    document.getElementById('modalDesconto').textContent = `- ${formatCurrencyBRL(descontoTotal).replace('R$ ', 'R$ ')}`;
    document.getElementById('modalTotal').textContent = formatCurrencyBRL(total);
}


// ================================
// MONTAR PAYLOAD VENDA (compartilhado)
// ================================
function montarPayloadVenda() {
    const subtotalBruto = carrinho.reduce((s, i) => s + (i.quantidade * i.preco), 0);
    const subtotalItens = carrinho.reduce((s, i) => s + i.subtotal, 0);
    const descontoItens = subtotalBruto - subtotalItens;
    const descontoGeralEl = document.getElementById('descontoGeralModal');
    const descontoGeral = descontoGeralEl ? (parseFloat(descontoGeralEl.value) || 0) : 0;
    const descontoTotal = descontoItens + descontoGeral;
    const total = Math.max(0, subtotalBruto - descontoTotal);
    const qtdTotal = carrinho.reduce((s, i) => s + i.quantidade, 0);
    const itensStr = carrinho.map(i => {
        const d = i.desconto > 0 ? ` (-${i.desconto.toFixed(1)}%)` : '';
        return `${i.nome} (${i.quantidade}${d})`;
    }).join(', ');
    const caixaVal = document.getElementById('caixaVenda') ? document.getElementById('caixaVenda').value : 'Dinheiro';

    return {
        data: new Date().toLocaleDateString('pt-BR'),
        cliente: document.getElementById('cliente').value || 'Consumidor Interno',
        itens: itensStr,
        itensList: carrinho,
        quantidadeVendida: qtdTotal,
        subtotal: subtotalBruto,
        descontoPercentual: subtotalBruto > 0 ? ((descontoTotal / subtotalBruto) * 100).toFixed(2) : 0,
        descontoReal: descontoTotal,
        totalComDesconto: total,
        formaPagamento: formaPagamentoSelecionada,
        caixa: caixaVal,
        usuario: document.getElementById('usuario').value || 'Administrador'
    };
}


// ================================
// CONFIRMAR VENDA (Finalizar — Concluída)
// ================================
async function confirmarVenda() {
    const btn = document.getElementById('btnConfirmarVenda');
    await execWithSpinner(btn, async () => {
        if (!formaPagamentoSelecionada) { await CustomModal.alert('Selecione a forma de pagamento.', 'OK'); return; }
        if (carrinho.length === 0) { fecharModal(); return; }

        const prazoResult = await calcularVencimentoStatus();
        if (!prazoResult) return;

        const payload = montarPayloadVenda();
        payload.vencimento = prazoResult.vencimento;
        payload.statusFinanceiro = prazoResult.status;

        const action = vendaEditandoId ? 'finalizarPendente' : 'lancarVenda';
        if (vendaEditandoId) {
            payload.id = vendaEditandoId;
            payload.idVenda = vendaEditandoId;
        }

        try {
            const res = await fetch(window.MASTER_WEBHOOK_URL, {
                method: 'POST',
                body: JSON.stringify({ acao: action, action: action, data: payload, spreadsheetId: window.SPREADSHEET_ID })
            });
            const data = await res.json();
            fecharModal();
            exibirStatus(data);
            if (data.status === 'sucesso') {
                const cupomData = {
                    id: data.id || '?',
                    data: payload.data,
                    cliente: payload.cliente,
                    operador: document.getElementById('usuario').value,
                    itens: [...carrinho],
                    formaPagamento: formaPagamentoSelecionada,
                    vencimento: prazoResult.vencimento,
                    statusPgto: prazoResult.status,
                    subtotal: 0,
                    descontoGeral: 0,
                    total: payload.totalComDesconto
                };
                carrinho = [];
                vendaEditandoId = null;
                renderizarCarrinho();
                fecharSlideoverVenda(); // Fecha a gaveta
                CacheAPI.clear('cache_produtos'); // Invalida cache porque baixou estoque
                await carregarProdutos();
                await carregarHistoricoVendas();
                abrirCupom(cupomData);
            }
        } catch (e) {
            exibirStatus({ status: 'error', mensagem: 'Erro de comunicação: ' + e });
        }
    });
}

// ================================
// HISTÓRICO DE VENDAS — COM STATUS E AÇÕES ERP
// ================================
async function carregarHistoricoVendas(filtros = null, msgCarregando = 'Carregando...') {
    const tbody = document.getElementById('listaHistorico');
    tbody.innerHTML = `<tr><td colspan="10" class="td-empty">${msgCarregando}</td></tr>`;
    try {
        const payload = { action: 'obterVendas', spreadsheetId: window.SPREADSHEET_ID };
        if (filtros) Object.assign(payload, filtros);

        const res = await fetch(window.MASTER_WEBHOOK_URL, { method: 'POST', body: JSON.stringify(payload) });
        const text = await res.text();
        const data = JSON.parse(text);

        if (data.status === 'sucesso' && data.dados) {
            const arr = parseCompactData(data.dados);
            const vendas = [...arr].reverse().slice(0, 40);
            tbody.innerHTML = '';
            vendas.forEach(v => {
                const id = v['ID da Venda'] || '';
                const dataV = v['Data'] || '-';
                const cliente = v['Cliente'] || '-';
                const operador = v['Usuario'] || v['Usuário'] || '-';
                const itens = String(v['Itens'] || '');
                const pgto = v['Forma de Pagamento'] || '-';
                const total = isNaN(parseFloat(v['Total com Desconto'])) ? 0 : parseFloat(v['Total com Desconto']);
                const status = v['Status'] || '';
                const itensJSON = v['ItensJSON'] || '[]';

                const isMulti = itens.includes(',');
                const itensDisplay = isMulti ? itens.substring(0, 35) + '...' : itens;
                const expandBtn = isMulti
                    ? `<button class="expand-btn" onclick="toggleItens(this)">▶ ver itens</button>
                       <div class="items-detail">${itens.split(', ').join('<br>')}</div>`
                    : '';

                // Ícone WhatsApp oficial (SVG)
                const _waSvg = `<svg width="13" height="13" fill="#16a34a" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967c-.273-.099-.471-.148-.67.15c-.197.297-.767.966-.94 1.164c-.173.199-.347.223-.644.075c-.297-.15-1.255-.463-2.39-1.475c-.883-.788-1.48-1.761-1.653-2.059c-.173-.297-.018-.458.13-.606c.134-.133.298-.347.446-.52c.149-.174.198-.298.298-.497c.099-.198.05-.371-.025-.52c-.075-.149-.669-1.612-.916-2.207c-.242-.579-.487-.5-.669-.51c-.173-.008-.371-.01-.57-.01c-.198 0-.52.074-.792.372c-.272.297-1.04 1.016-1.04 2.479c0 1.462 1.065 2.875 1.213 3.074c.149.198 2.096 3.2 5.077 4.487c.709.306 1.262.489 1.694.625c.712.227 1.36.195 1.871.118c.571-.085 1.758-.719 2.006-1.413c.248-.694.248-1.289.173-1.413c-.074-.124-.272-.198-.57-.347zm-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214l-3.741.982l.998-3.648l-.235-.374A9.86 9.86 0 012.166 11.892C2.167 6.442 6.602 2.008 12.054 2.008c2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884zm8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>`;
                const whatsappBtn = `<button title="Enviar comprovante via WhatsApp" style="background:none;border:1px solid #22c55e;border-radius:4px;padding:3px 6px;cursor:pointer;display:inline-flex;align-items:center;" onclick="enviarWhatsApp(${id},'${encodeURIComponent(cliente)}','${encodeURIComponent(itensJSON)}',${total},'${dataV}')">${_waSvg}</button>`;

                let statusBadge = '';
                let acoes = '';
                if (status === 'Pendente') {
                    statusBadge = `<span style="background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;">&#128336; Pendente</span>`;
                    acoes = `
                        <button class="edit-btn" style="font-size:11px;" onclick="editarRascunho(${id}, '${encodeURIComponent(itensJSON)}')">&#9999;&#65039; Editar</button>
                        <button class="edit-btn" style="background:#16a34a;font-size:11px;" onclick="abrirModalFinalizarPendente(${id}, '${encodeURIComponent(itensJSON)}')">&#9989; Finalizar</button>
                    `;
                } else if (status === 'Concluída' || status === 'Concluda' || status === '') {
                    statusBadge = `<span style="background:#d1fae5;color:#065f46;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;">&#9989; Concluída</span>`;
                    const _bscCon = typeof Auth !== 'undefined' && Auth.isPlanBasico();
                    const _printCon = _bscCon ? '' : `<button title="Reimprimir cupom" data-print-btn style="background:none;border:1px solid #cbd5e1;border-radius:4px;padding:2px 6px;cursor:pointer;font-size:13px;" onclick="reimprimirCupom(${id},'${encodeURIComponent(itensJSON)}','${encodeURIComponent(cliente)}','${encodeURIComponent(operador)}','${encodeURIComponent(pgto)}',${total},'${dataV}')">&#128424;&#65039;</button>`;
                    acoes = `${_printCon}${whatsappBtn}<button class="delete-btn" style="background:#f59e0b;color:#fff;font-size:11px;" data-admin-btn onclick="confirmarEstorno(${id})">&#8617;&#65039; Estornar</button>`;
                } else if (status === 'Estornada') {
                    statusBadge = `<span style="background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;">&#8617;&#65039; Estornada</span>`;
                    const _basico2 = typeof Auth !== 'undefined' && Auth.isPlanBasico();
                    const _printBtn2 = _basico2 ? '' : `<button title="Reimprimir cupom" data-print-btn style="background:none;border:1px solid #cbd5e1;border-radius:4px;padding:2px 6px;cursor:pointer;font-size:13px;" onclick="reimprimirCupom(${id},'${encodeURIComponent(itensJSON)}','${encodeURIComponent(cliente)}','${encodeURIComponent(operador)}','${encodeURIComponent(pgto)}',${total},'${dataV}')">&#128424;&#65039;</button>`;
                    const btnReaproveitarX = `<button title="Copiar itens para nova venda (Reaproveitar)" class="edit-btn" style="background:none; border:1px solid #3b82f6; padding:3px 6px; font-size:13px; margin-right:2px; border-radius:4px; cursor:pointer;" onclick="reaproveitarVenda('${encodeURIComponent(itensJSON)}')">&#128260;</button>`;
                    acoes = `${btnReaproveitarX}${_printBtn2}`;
                }


                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${id}</td>
                    <td>${dataV}</td>
                    <td>${cliente}</td>
                    <td>${operador}</td>
                    <td><span>${itensDisplay}</span>${expandBtn}</td>
                    <td>${pgto}</td>
                    <td>${statusBadge}</td>
                    <td><strong>${formatCurrencyBRL(total)}</strong></td>
                    <td><div class="action-buttons">${acoes}</div></td>
                `;
                tbody.appendChild(tr);
            });
            if (typeof Auth !== 'undefined') Auth.applyUI();
        } else {
            tbody.innerHTML = `<tr><td colspan="10" class="td-empty">${data.mensagem || 'Nenhuma venda encontrada.'}</td></tr>`;
        }
    } catch (e) {
        console.error('Erro histórico:', e);
        tbody.innerHTML = `<tr><td colspan="10" class="td-empty" style="color:#ef4444;">Erro ao carregar histórico: ${e.message}</td></tr>`;
    }
}

async function filtrarHistoricoVendas() {
    const btn = document.getElementById('btnFiltrarHistorico');
    if (!btn) return;

    await execWithSpinner(btn, async () => {
        const dataInicio = document.getElementById('filtroInicio').value;
        const dataFim = document.getElementById('filtroFim').value;

        let hint = 'Carregando...';
        if (dataInicio) {
            const dInicio = new Date(dataInicio + 'T00:00:00');
            // Se buscar algo com mais de 60 dias, o script lerá a base de Historico
            if ((Date.now() - dInicio) > 60 * 24 * 60 * 60 * 1000) {
                hint = 'Consultando Arquivos (> 60 dias)...';
            }
        }

        btn.textContent = hint;
        await carregarHistoricoVendas({ dataInicio, dataFim }, hint);
        btn.textContent = 'Buscar';
    });
}

function toggleItens(btn) {
    const detail = btn.nextElementSibling;
    const isOpen = detail.classList.toggle('open');
    btn.textContent = isOpen ? '▼ ocultar' : '▶ ver itens';
}

// ================================
// EDITAR RASCUNHO (carrega de volta ao carrinho)
// ================================
function editarRascunho(id, itensJSONEncoded) {
    try {
        const itensJSON = decodeURIComponent(itensJSONEncoded);
        const itens = JSON.parse(itensJSON);
        if (!itens || itens.length === 0) {
            exibirStatus({ status: 'error', mensagem: 'ItensJSON vazio — não é possível editar este rascunho.' });
            return;
        }
        carrinho = itens.map(i => ({
            nome: i.nome, preco: parseFloat(i.preco) || 0,
            quantidade: parseFloat(i.quantidade) || 0,
            desconto: parseFloat(i.desconto) || 0,
            subtotal: parseFloat(i.subtotal) || 0
        }));
        vendaEditandoId = id;
        renderizarCarrinho();
        abrirSlideoverVenda();
        exibirStatus({ status: 'success', mensagem: `✏️ Rascunho #${id} carregado. Edite e finalize.` });
    } catch (e) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar rascunho: ' + e.message });
    }
}

// ================================
// REAPROVEITAR VENDA (Cria NOVO Carrinho)
// ================================
function reaproveitarVenda(itensJSONEncoded) {
    try {
        const itensJSON = decodeURIComponent(itensJSONEncoded);
        const itens = JSON.parse(itensJSON);
        if (!itens || itens.length === 0) {
            exibirStatus({ status: 'error', mensagem: 'Esta venda não possui itens para serem reaproveitados.' });
            return;
        }
        carrinho = itens.map(i => ({
            nome: i.nome, preco: parseFloat(i.preco) || 0,
            quantidade: parseFloat(i.quantidade) || 0,
            desconto: parseFloat(i.desconto) || 0,
            subtotal: parseFloat(i.subtotal) || 0
        }));

        vendaEditandoId = null; // Zera o ID: É uma nova venda
        renderizarCarrinho();
        abrirSlideoverVenda();
        exibirStatus({ status: 'success', mensagem: `🔄 Itens copiados! Edite as quantidades e crie o NOVO pedido.` });
    } catch (e) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao reaproveitar venda: ' + e.message });
    }
}

// ================================
// FINALIZAR VENDA PENDENTE (via histórico)
// ================================
function abrirModalFinalizarPendente(id, itensJSONEncoded) {
    try {
        const itensJSON = decodeURIComponent(itensJSONEncoded);
        const itens = JSON.parse(itensJSON);
        if (!itens || itens.length === 0) {
            exibirStatus({ status: 'error', mensagem: 'ItensJSON vazio — não é possível finalizar este rascunho.' });
            return;
        }
        carrinho = itens.map(i => ({
            nome: i.nome, preco: parseFloat(i.preco) || 0,
            quantidade: parseFloat(i.quantidade) || 0,
            desconto: parseFloat(i.desconto) || 0,
            subtotal: parseFloat(i.subtotal) || 0
        }));
        vendaEditandoId = id;
        renderizarCarrinho();
        abrirSlideoverVenda();
        abrirModal();
    } catch (e) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar itens para finalização: ' + e.message });
    }
}

// ================================
// ESTORNAR VENDA CONCLUÍDA
// ================================
async function confirmarEstorno(id) {
    if (!(await CustomModal.confirm(`⚠️ Estornar a Venda #${id}?\n\nEsta ação irá:\n• Devolver os itens ao estoque\n• Cancelar o lançamento financeiro\n• Marcar a venda como Estornada\n\nEsta operação não pode ser desfeita.`, 'Estornar', 'Cancelar'))) return;
    try {
        const res = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'estornarVenda', idVenda: id, spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await res.json();
        exibirStatus(data);
        if (data.status === 'sucesso') {
            await carregarProdutos();
            await carregarHistoricoVendas();
        }
    } catch (e) { exibirStatus({ status: 'error', mensagem: 'Erro ao estornar: ' + e.message }); }
}

async function excluirVenda(id) {
    if (!(await CustomModal.confirm(`Excluir o rascunho pendente #${id}?\nNão há estoque nem financeiro associados a este rascunho.`, 'Excluir', 'Cancelar'))) return;
    try {
        const res = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'excluirRascunho', idRascunho: id, spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await res.json();
        exibirStatus(data);
        if (data.status === 'sucesso') await carregarHistoricoVendas();
    } catch (e) { exibirStatus({ status: 'error', mensagem: 'Erro ao excluir rascunho: ' + e.message }); }
}

// Reimprimir cupom de uma venda do histórico
function reimprimirCupom(id, itensJSONEnc, clienteEnc, operadorEnc, pgtoEnc, total, data) {
    const cliente = decodeURIComponent(clienteEnc);
    const operador = decodeURIComponent(operadorEnc);
    const pgto = decodeURIComponent(pgtoEnc);
    let itens = [];
    try { itens = JSON.parse(decodeURIComponent(itensJSONEnc)); } catch (e) { }

    const cupom = {
        id, data, cliente, operador,
        itens: itens.map(i => ({
            nome: i.nome || '?',
            preco: parseFloat(i.preco) || 0,
            quantidade: parseFloat(i.quantidade) || 0,
            desconto: parseFloat(i.desconto) || 0,
            subtotal: parseFloat(i.subtotal) || 0
        })),
        formaPagamento: pgto,
        vencimento: '-',
        statusPgto: '-',
        subtotal: 0,
        descontoGeral: 0,
        total
    };
    abrirCupom(cupom);
}



function formatarData(valor) {
    if (!valor) return '-';
    const str = String(valor);
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) return str.substring(0, 10);
    const d = new Date(str);
    return isNaN(d) ? str : d.toLocaleDateString('pt-BR');
}

// ================================
// CUPOM DE VENDA
// ================================
let ultimoCupomData = null;

function abrirCupom(cupom) {
    ultimoCupomData = cupom;
    const linha = '--------------------------------';
    const linhaDupla = '================================';

    let subtotalBruto = 0;

    const itensHtml = cupom.itens.map(i => {
        const itemBruto = i.quantidade * i.preco;
        subtotalBruto += itemBruto;

        const sub = itemBruto.toFixed(2).replace('.', ',');
        const qtd = `${i.quantidade}`;
        const preco = `${i.preco.toFixed(2).replace('.', ',')}`;
        return `
        <div style="display:flex; justify-content:space-between; font-size:11px; margin-bottom: 3px; color:#000; font-weight:500;">
            <div style="width:12%;">${qtd}</div>
            <div style="width:48%; padding-right:4px; word-break:break-word;">${i.nome}</div>
            <div style="width:20%; text-align:right;">${preco}</div>
            <div style="width:20%; text-align:right;">${sub}</div>
        </div>`;
    }).join('');

    let descontoTotal = subtotalBruto - cupom.total;
    if (descontoTotal < 0) descontoTotal = 0;

    const empresaNome = (typeof Auth !== 'undefined' && typeof Auth.getEmpresa === 'function') ? Auth.getEmpresa().toUpperCase() : 'GESTÃO & CONTROLE';
    const html = `
        <div style="text-align:center;font-family:monospace;margin-top:20px;padding:0 10px;">
        <div style="text-align:center;font-weight:bold;font-size:14px;margin-bottom:4px;color:#000;">${empresaNome}</div>
        <div style="text-align:center;font-size:12px;margin-bottom:8px;color:#000;">Cupom não fiscal</div>
        </div>
        <div style="margin:6px 0;color:#000;">${linhaDupla}</div>
        <div style="color:#000;font-weight:500;"><b>Venda #:</b> ${cupom.id}</div>
        <div style="color:#000;font-weight:500;"><b>Data:</b> ${cupom.data}</div>
        <div style="color:#000;font-weight:500;"><b>Cliente:</b> ${cupom.cliente}</div>
        <div style="color:#000;font-weight:500;"><b>Operador:</b> ${cupom.operador}</div>
        <div style="margin:6px 0;color:#000;">${linha}</div>
        <div style="display:flex; justify-content:space-between; font-size:11px; font-weight:bold; margin-bottom: 4px; padding-bottom: 2px; border-bottom: 1px dashed #000; color:#000;">
            <div style="width:12%;">Qtd</div>
            <div style="width:48%;">Item</div>
            <div style="width:20%; text-align:right;">Unit.</div>
            <div style="width:20%; text-align:right;">Total</div>
        </div>
        ${itensHtml}
        <div style="margin:6px 0;color:#000;">${linha}</div>
        <div style="display:flex;justify-content:space-between;color:#000;font-weight:500;"><span>Subtotal:</span><span>${formatCurrencyBRL(subtotalBruto)}</span></div>
        ${descontoTotal > 0.005 ? `<div style="display:flex;justify-content:space-between;color:#000;font-weight:bold;"><span>Desconto Total:</span><span>- ${formatCurrencyBRL(descontoTotal)}</span></div>` : ''}
        <div style="margin:6px 0;color:#000;">${linhaDupla}</div>
        <div style="display:flex;justify-content:space-between;font-size:14px;font-weight:bold;color:#000;"><span>TOTAL:</span><span>${formatCurrencyBRL(cupom.total)}</span></div>
        <div style="margin:6px 0;color:#000;">${linha}</div>
        <div style="color:#000;font-weight:500;"><b>Pagamento:</b> ${cupom.formaPagamento}</div>
        <div style="color:#000;font-weight:500;"><b>Vencimento:</b> ${cupom.vencimento} (${cupom.statusPgto})</div>
        <div style="margin:6px 0;color:#000;">${linhaDupla}</div>
        <div style="text-align:center;font-size:11px;color:#000;font-weight:600;">Obrigado pela preferência!</div>
        <div style="text-align:center;font-size:10px;color:#000;font-weight:500;">${new Date().toLocaleString('pt-BR')}</div>
    `;
    document.getElementById('cupomConteudo').innerHTML = html;

    // Injeção de Botões Baseados no Plano
    const isBasico = typeof Auth !== 'undefined' && Auth.isPlanBasico();
    const btnContainer = document.getElementById('botoesCupomContainer');
    const msgData = `'${cupom.id}', '${encodeURIComponent(cupom.cliente)}', '${encodeURIComponent(JSON.stringify(cupom.itens))}', ${cupom.total}, '${cupom.data}'`;

    if (isBasico) {
        btnContainer.innerHTML = `
            <button onclick="fecharCupom()"
                style="flex:1; padding:0.6rem; border:2px solid #e2e8f0; border-radius:0.5rem; background:white; cursor:pointer; font-size:0.85rem;">Fechar</button>
            <button onclick="enviarWhatsApp(${msgData})"
                style="flex:2; padding:0.6rem; background:#25D366; color:white; border:none; border-radius:0.5rem; font-weight:600; cursor:pointer; font-size:0.85rem; display:flex; align-items:center; justify-content:center; gap:6px;">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"></path></svg>
                Enviar no WhatsApp
            </button>
        `;
    } else {
        btnContainer.innerHTML = `
            <button onclick="fecharCupom()"
                style="flex:1; padding:0.6rem; border:2px solid #e2e8f0; border-radius:0.5rem; background:white; cursor:pointer; font-size:0.85rem;">Fechar</button>
            <button onclick="imprimirCupomAgora()"
                style="flex:1; padding:0.6rem; background:#16a34a; color:white; border:none; border-radius:0.5rem; font-weight:600; cursor:pointer; font-size:0.85rem;">🖨️
                Imprimir</button>
            <button onclick="enviarWhatsApp(${msgData})"
                style="flex:1; padding:0.6rem; background:#25D366; color:white; border:none; border-radius:0.5rem; font-weight:600; cursor:pointer; font-size:0.85rem; display:flex; align-items:center; justify-content:center; gap:6px;">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"></path></svg>
                Zap</button>
        `;
    }

    document.getElementById('modalCupom').style.display = 'flex';
}

function fecharCupom() {
    document.getElementById('modalCupom').style.display = 'none';
}

function imprimirCupomAgora() {
    const conteudo = document.getElementById('cupomConteudo').innerHTML;
    const win = window.open('', '_blank', 'width=360,height=600');
    win.document.write(`
        <html><head><title>Cupom de Venda</title>
        <style>
            body { font-family: 'Courier New', monospace; font-size: 12px; margin: 10px; width: 280px; }
            @media print { body { margin: 0; } }
        </style></head>
        <body>${conteudo}</body></html>
    `);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); win.close(); }, 300);
}

// ================================
// WHATSAPP COMPROVANTE
// ================================
function enviarWhatsApp(id, clienteEnc, itensJSONEnc, total, dataV) {
    const cliente = decodeURIComponent(clienteEnc);
    const fmtBRL = v => 'R$ ' + parseFloat(v).toFixed(2).replace('.', ',');
    const isBasico = typeof Auth !== 'undefined' && Auth.isPlanBasico();

    let msg = '';
    let itens = [];
    try { itens = JSON.parse(decodeURIComponent(itensJSONEnc)); } catch (e) { }

    if (isBasico) {
        // Mensagem otimizada para plano Básico (com itens e quantidades resumidas)
        const itensListaSimples = itens.length > 0
            ? itens.map(i => `${i.quantidade}x ${i.nome}`).join(', ')
            : 'itens não disponíveis';

        msg = `*Pedido #${id}*\nItens: ${itensListaSimples}\n*Total: ${fmtBRL(total)}*\nObrigado pela preferência!`;
    } else {
        // Mensagem rica para Pro/Premium
        const linha = '------------------------------';
        const itensLista = itens.length > 0
            ? itens.map(i => `  • ${i.nome} x${i.quantidade} = ${fmtBRL(i.subtotal)}`).join('\n')
            : '  (itens não disponíveis)';

        msg = [
            `*Gestão&Controle - Comprovante de Venda*`,
            linha,
            `🛒 *Pedido:* #${id}`,
            `📅 *Data:* ${dataV}`,
            `👤 *Cliente:* ${cliente}`,
            `✅ *Status:* Concluída`,
            linha,
            `🛎️ *Itens:*`,
            itensLista,
            linha,
            `💰 *Total:* ${fmtBRL(total)}`,
            linha,
            `🙏 Obrigado pela preferência!`
        ].join('\n');
    }

    const encoded = encodeURIComponent(msg);
    window.open(`https://wa.me/?text=${encoded}`, '_blank');
}
