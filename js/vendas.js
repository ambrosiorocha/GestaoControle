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
    if (SCRIPT_URL === '') {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.SCRIPT_URL no config.js.' });
        return;
    }
    carregarProdutos();
    carregarClientes();
    carregarOperadores();
    carregarHistoricoVendas();

    document.getElementById('produto').addEventListener('change', atualizarPrecoUnitario);
    document.getElementById('quantidade').addEventListener('input', atualizarSubtotalItem);
    document.getElementById('descontoItemPct').addEventListener('input', () => sincronizarDesconto('pct'));
    document.getElementById('descontoItemReais').addEventListener('input', () => sincronizarDesconto('reais'));

    document.querySelectorAll('.pgto-btn').forEach(btn => {
        btn.addEventListener('click', function () {
            document.querySelectorAll('.pgto-btn').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            formaPagamentoSelecionada = this.dataset.pgto;
            aplicarPrazo(this);
            atualizarModalTotais();
        });
    });
    document.getElementById('descontoGeralModal').addEventListener('input', atualizarModalTotais);
});

// ================================
// STATUS MESSAGE
// ================================
function exibirStatus(resposta) {
    var el = document.getElementById('statusMessage');
    el.textContent = resposta.mensagem;
    el.className = '';
    if (resposta.status) el.classList.add(resposta.status);
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
        const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'obterProdutos' }) });
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
        const precoRaw = String(p['Preço'] || p.Preco || 0).replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
        const opt = document.createElement('option');
        opt.value = p.Nome;
        opt.textContent = `${p.Nome}  (Estoque: ${estoque})`;
        opt.dataset.preco = parseFloat(precoRaw) || 0;
        opt.dataset.estoque = estoque;
        sel.appendChild(opt);
    });
}

async function carregarOperadores() {
    const sel = document.getElementById('usuario');
    try {
        const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'obterOperadores' }) });
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
        const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'obterClientes' }) });
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
    document.getElementById('precoUnitario').value = preco > 0 ? preco.toFixed(2) : '';
    document.getElementById('descontoItemPct').value = '0';
    document.getElementById('descontoItemReais').value = '0';
    atualizarSubtotalItem();
}

function sincronizarDesconto(origem) {
    const preco = getPrecoUnitario();
    const qtd = parseFloat(document.getElementById('quantidade').value) || 0;
    const bruto = preco * qtd;
    if (origem === 'pct') {
        const pct = parseFloat(document.getElementById('descontoItemPct').value) || 0;
        document.getElementById('descontoItemReais').value = bruto > 0 ? (bruto * pct / 100).toFixed(2) : '0';
    } else {
        const reais = parseFloat(document.getElementById('descontoItemReais').value) || 0;
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
        el.textContent = `R$ ${subtotal.toFixed(2).replace('.', ',')}`;
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
            <td>${item.nome}</td>
            <td style="text-align:center;">${item.quantidade}</td>
            <td>${fmtBRL(item.preco)}</td>
            <td>${item.desconto > 0 ? item.desconto.toFixed(1) + '%' : '-'}</td>
            <td><strong>${fmtBRL(item.subtotal)}</strong></td>
            <td><button class="remove-item" onclick="removerItem(${idx})" title="Remover">✕</button></td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('totalCarrinho').textContent = fmtBRL(totalGeral);
    document.getElementById('qtdItensLabel').textContent = `${carrinho.length} item(ns)`;
    setBtns(true);
}


function removerItem(idx) { carrinho.splice(idx, 1); renderizarCarrinho(); }

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
        try {
            const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'salvarRascunho', data: payload }) });
            const data = await res.json();
            exibirStatus(data);
            if (data.status === 'sucesso') {
                carrinho = [];
                vendaEditandoId = null;
                renderizarCarrinho();
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

function calcularVencimentoStatus() {
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
        if (!raw) { alert('Informe a data de vencimento.'); return null; }
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

    document.getElementById('modalSubtotal').textContent = `R$ ${subtotalBruto.toFixed(2).replace('.', ',')}`;
    document.getElementById('modalDesconto').textContent = `- R$ ${descontoTotal.toFixed(2).replace('.', ',')}`;
    document.getElementById('modalTotal').textContent = `R$ ${total.toFixed(2).replace('.', ',')}`;
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
        usuario: document.getElementById('usuario').value || 'Administrador'
    };
}


// ================================
// CONFIRMAR VENDA (Finalizar — Concluída)
// ================================
async function confirmarVenda() {
    const btn = document.getElementById('btnConfirmarVenda');
    await execWithSpinner(btn, async () => {
        if (!formaPagamentoSelecionada) { alert('Selecione a forma de pagamento.'); return; }
        if (carrinho.length === 0) { fecharModal(); return; }

        const prazoResult = calcularVencimentoStatus();
        if (!prazoResult) return;

        const payload = montarPayloadVenda();
        payload.vencimento = prazoResult.vencimento;
        payload.statusFinanceiro = prazoResult.status;

        const action = vendaEditandoId ? 'finalizarPendente' : 'lancarVenda';
        if (vendaEditandoId) payload.id = vendaEditandoId;

        try {
            const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action, data: payload }) });
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
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:1rem;color:#94a3b8;">${msgCarregando}</td></tr>`;
    try {
        const payload = { action: 'obterVendas' };
        if (filtros) payload.data = filtros;

        const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify(payload) });
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

                // Badge de status
                let statusBadge = '';
                let acoes = '';
                if (status === 'Pendente') {
                    statusBadge = `<span style="background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;">🕐 Pendente</span>`;
                    acoes = `
                        <button class="edit-btn" style="font-size:11px;" onclick="editarRascunho(${id}, '${encodeURIComponent(itensJSON)}')">✏️ Editar</button>
                        <button class="edit-btn" style="background:#16a34a;font-size:11px;" onclick="abrirModalFinalizarPendente(${id}, '${encodeURIComponent(itensJSON)}')">✅ Finalizar</button>
                        <button class="delete-btn" style="font-size:11px;" data-admin-btn onclick="excluirVenda(${id})">🗑</button>
                    `;
                } else if (status === 'Concluda' || status === '') {
                    statusBadge = `<span style="background:#d1fae5;color:#065f46;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;">✅ Concluída</span>`;
                    acoes = `
                        <button title="Reimprimir cupom" style="background:none;border:1px solid #cbd5e1;border-radius:4px;padding:2px 6px;cursor:pointer;font-size:13px;"
                            onclick="reimprimirCupom(${id},'${encodeURIComponent(itensJSON)}','${encodeURIComponent(cliente)}','${encodeURIComponent(operador)}','${encodeURIComponent(pgto)}',${total},'${dataV}')">🖨️</button>
                        <button class="delete-btn" style="background:#f59e0b;color:#fff;font-size:11px;" data-admin-btn onclick="confirmarEstorno(${id})">↩️ Estornar</button>
                    `;
                } else if (status === 'Estornada') {
                    statusBadge = `<span style="background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;">↩️ Estornada</span>`;
                    acoes = `<button title="Reimprimir cupom" style="background:none;border:1px solid #cbd5e1;border-radius:4px;padding:2px 6px;cursor:pointer;font-size:13px;"
                        onclick="reimprimirCupom(${id},'${encodeURIComponent(itensJSON)}','${encodeURIComponent(cliente)}','${encodeURIComponent(operador)}','${encodeURIComponent(pgto)}',${total},'${dataV}')">🖨️</button>`;
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
                    <td><strong>R$ ${total.toFixed(2).replace('.', ',')}</strong></td>
                    <td><div class="action-buttons">${acoes}</div></td>
                `;
                tbody.appendChild(tr);
            });
            if (typeof Auth !== 'undefined') Auth.applyUI();
        } else {
            tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:1rem;">${data.mensagem || 'Nenhuma venda encontrada.'}</td></tr>`;
        }
    } catch (e) {
        console.error('Erro histórico:', e);
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:#ef4444;padding:1rem;">Erro ao carregar histórico: ${e.message}</td></tr>`;
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
        window.scrollTo({ top: 0, behavior: 'smooth' });
        exibirStatus({ status: 'success', mensagem: `✏️ Rascunho #${id} carregado. Edite e finalize.` });
    } catch (e) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar rascunho: ' + e.message });
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
        window.scrollTo({ top: 0, behavior: 'smooth' });
        abrirModal();
    } catch (e) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar itens para finalização: ' + e.message });
    }
}

// ================================
// ESTORNAR VENDA CONCLUÍDA
// ================================
async function confirmarEstorno(id) {
    if (!confirm(`⚠️ Estornar a Venda #${id}?\n\nEsta ação irá:\n• Devolver os itens ao estoque\n• Cancelar o lançamento financeiro\n• Marcar a venda como Estornada\n\nEsta operação não pode ser desfeita.`)) return;
    try {
        const res = await fetch(window.SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'estornarVenda', data: { id } }) });
        const data = await res.json();
        exibirStatus(data);
        if (data.status === 'sucesso') {
            await carregarProdutos();
            await carregarHistoricoVendas();
        }
    } catch (e) { exibirStatus({ status: 'error', mensagem: 'Erro ao estornar: ' + e.message }); }
}

async function excluirVenda(id) {
    if (!confirm(`Excluir o rascunho pendente #${id}?\nNão há estoque nem financeiro associados a este rascunho.`)) return;
    exibirStatus({ status: 'error', mensagem: 'Função de exclusão de rascunho em desenvolvimento.' });
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

    const html = `
        <div style="text-align:center;font-family:monospace;margin-top:20px;padding:0 10px;">
        <div style="text-align:center;font-weight:bold;font-size:14px;margin-bottom:4px;color:#000;">GESTÃO & CONTROLE</div>
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
        <div style="display:flex;justify-content:space-between;color:#000;font-weight:500;"><span>Subtotal:</span><span>R$ ${subtotalBruto.toFixed(2).replace('.', ',')}</span></div>
        ${descontoTotal > 0.005 ? `<div style="display:flex;justify-content:space-between;color:#000;font-weight:bold;"><span>Desconto Total:</span><span>- R$ ${descontoTotal.toFixed(2).replace('.', ',')}</span></div>` : ''}
        <div style="margin:6px 0;color:#000;">${linhaDupla}</div>
        <div style="display:flex;justify-content:space-between;font-size:14px;font-weight:bold;color:#000;"><span>TOTAL:</span><span>R$ ${cupom.total.toFixed(2).replace('.', ',')}</span></div>
        <div style="margin:6px 0;color:#000;">${linha}</div>
        <div style="color:#000;font-weight:500;"><b>Pagamento:</b> ${cupom.formaPagamento}</div>
        <div style="color:#000;font-weight:500;"><b>Vencimento:</b> ${cupom.vencimento} (${cupom.statusPgto})</div>
        <div style="margin:6px 0;color:#000;">${linhaDupla}</div>
        <div style="text-align:center;font-size:11px;color:#000;font-weight:600;">Obrigado pela preferência!</div>
        <div style="text-align:center;font-size:10px;color:#000;font-weight:500;">${new Date().toLocaleString('pt-BR')}</div>
    `;
    document.getElementById('cupomConteudo').innerHTML = html;
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
