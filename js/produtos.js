let produtos = []; // Array global para armazenar os produtos

document.addEventListener('DOMContentLoaded', function () {
    if (window.MASTER_WEBHOOK_URL === '' || window.MASTER_WEBHOOK_URL.includes('COLE_AQUI')) {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.MASTER_WEBHOOK_URL no config.js.' });
        return;
    }
    document.getElementById('produtoForm').addEventListener('submit', function (e) {
        e.preventDefault();
        execWithSpinner(document.querySelector('#produtoForm button[type="submit"]'), salvarProduto);
    });

    document.getElementById('pesquisa').addEventListener('input', filtrarProdutos);

    var btnSync = document.getElementById('btnSincronizar');
    if (btnSync) {
        btnSync.addEventListener('click', function (e) {
            execWithSpinner(e.target, async () => { await carregarProdutos(true); });
        });
    }

    // === Gate de Plano: Ocultar precificação para plano Básico ===
    if (typeof Auth !== 'undefined' && Auth.isPlanBasico()) {
        aplicarGatePlanBasico();
    }

    // Evento para fechar modal com tecla ESC
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') fecharSlideoverProduto();
    });

    carregarProdutos();
});

// ================================
// CONTROLE DO SLIDE-OVER (GAVETA)
// ================================
function abrirSlideoverProduto() {
    const overlay = document.getElementById('slideoverProduto');
    const container = document.getElementById('slideoverContainerProduto');
    
    // Se abrir sem ID, é um "Novo"
    const id = document.getElementById('idProduto').value;
    if (!id) {
        document.getElementById('produtoForm').reset();
        document.getElementById('slideoverLabelProduto').textContent = 'Novo Produto';
    } else {
        document.getElementById('slideoverLabelProduto').textContent = 'Editar Produto';
    }

    overlay.classList.add('open');
    container.classList.add('open');
    document.body.classList.add('slideover-open');
}

function fecharSlideoverProduto() {
    const overlay = document.getElementById('slideoverProduto');
    const container = document.getElementById('slideoverContainerProduto');
    
    overlay.classList.remove('open');
    container.classList.remove('open');
    document.body.classList.remove('slideover-open');
    
    // Limpar formulário e ID ao fechar
    setTimeout(() => {
        document.getElementById('produtoForm').reset();
        document.getElementById('idProduto').value = '';
    }, 300);
}

function validarFechamentoSlideoverProduto(e) {
    if (e.target.id === 'slideoverProduto') {
        fecharSlideoverProduto();
    }
}


function aplicarGatePlanBasico() {
    // Oculta bloco de custo/margem e mostra cadeado
    const bloco = document.getElementById('precificacaoBlock');
    const lock = document.getElementById('precificacaoLock');
    if (bloco) bloco.style.display = 'none';
    if (lock) lock.style.display = 'block';

    // Remove 'required' dos campos ocultos para não bloquear o form
    ['precoCusto', 'margemPct', 'margemRS'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.removeAttribute('required');
    });

    // Sincroniza campo precoBasico ↔ preco
    const precoBasico = document.getElementById('precoBasico');
    const precoMain = document.getElementById('preco');
    if (precoBasico && precoMain) {
        precoBasico.addEventListener('input', () => { precoMain.value = precoBasico.value; });
    }

    // Oculta colunas de custo e margem no <thead>
    document.querySelectorAll('.col-custo, .col-margem').forEach(th => th.style.display = 'none');
}

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

async function salvarProduto() {

    const produto = {
        idProduto: document.getElementById('idProduto').value || null,
        nome: document.getElementById('nome').value,
        unidadeVenda: document.getElementById('unidadeVenda').value,
        precoCusto: parseCurrencyBRL(document.getElementById('precoCusto').value),
        margemPct: parseFloat(document.getElementById('margemPct').value) || 0,
        margemRS: parseCurrencyBRL(document.getElementById('margemRS').value),
        preco: parseCurrencyBRL(document.getElementById('preco').value),
        quantidade: parseFloat(document.getElementById('quantidade').value) || 0,
        descricao: document.getElementById('descricao').value
    };

    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarProduto', spreadsheetId: window.SPREADSHEET_ID, data: produto })
        });

        const data = await response.json();
        exibirStatus(data);

        if (data.status === 'sucesso') {
            document.getElementById('produtoForm').reset();
            document.getElementById('idProduto').value = '';
            fecharSlideoverProduto();
            await carregarProdutos(true);
        }

    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro de comunicação: ' + error });
    }
}

async function carregarProdutos(forceSync = false) {
    const listaProdutos = document.getElementById('listaProdutos');

    if (!forceSync) {
        const cached = CacheAPI.get('cache_produtos');
        if (cached) {
            produtos = cached;
            renderizarTabela(produtos);
            return;
        }
    }

    listaProdutos.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Carregando produtos...</td></tr>';

    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterProdutos', spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await response.json();

        if (data.status === 'sucesso' && data.dados) {
            produtos = parseCompactData(data.dados);
            CacheAPI.set('cache_produtos', produtos);
            if (produtos.length > 0) {
                renderizarTabela(produtos);
            } else {
                listaProdutos.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Nenhum produto cadastrado.</td></tr>';
            }
        } else {
            listaProdutos.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Nenhum produto cadastrado.</td></tr>';
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar lista de produtos: ' + error.message });
        listaProdutos.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Erro ao carregar produtos.</td></tr>';
    }
}


function renderizarTabela(produtosParaRenderizar) {
    const listaProdutos = document.getElementById('listaProdutos');
    listaProdutos.innerHTML = '';

    const isBasico = typeof Auth !== 'undefined' && Auth.isPlanBasico();

    if (produtosParaRenderizar.length === 0) {
        listaProdutos.innerHTML = `<tr><td colspan="${isBasico ? 7 : 9}" class="td-empty">Nenhum produto encontrado.</td></tr>`;
        return;
    }

    produtosParaRenderizar.forEach((produto, idx) => {
        const precoVenda = parseCurrencyBRL(produto['Preço_de_venda'] || produto['Preço'] || 0);
        const precoCusto = parseCurrencyBRL(produto['Preço_de_custo'] || 0);
        const margemPct = parseCurrencyBRL(String(produto['Margem_de_lucro(%)'] || 0).replace('%', ''));
        const id = produto['ID do Produto'];

        const custoCel = isBasico ? '' : `<td style="text-align:right; color:#64748b;">${precoCusto > 0 ? formatCurrencyBRL(precoCusto) : '-'}</td>`;
        const margemCel = isBasico ? '' : `<td style="text-align:center; ${margemPct > 0 ? 'color:#15803d; font-weight:600;' : 'color:#94a3b8;'}">${margemPct > 0 ? margemPct.toFixed(1) + '%' : '-'}</td>`;

        const row = document.createElement('tr');
        row.innerHTML = `
            <td style="color:#94a3b8; font-size:0.78rem;">${id}</td>
            <td style="font-weight:500;">${produto.Nome || ''}</td>
            <td style="text-align:center;">${produto['Unidade de Venda'] || ''}</td>
            ${custoCel}
            ${margemCel}
            <td style="text-align:right; font-weight:600; color:#1d4ed8;">${formatCurrencyBRL(precoVenda)}</td>
            <td style="text-align:center;">${produto.Quantidade ?? ''}</td>
            <td style="color:#94a3b8; font-size:0.78rem;">${produto.Descrição || ''}</td>
            <td>
                <div class="action-buttons">
                    <button class="edit-btn" onclick="editarProduto(${id})">Editar</button>
                    <button class="delete-btn" data-admin-btn onclick="excluirProduto(${id})">Excluir</button>
                </div>
            </td>
        `;
        listaProdutos.appendChild(row);
    });
    if (typeof Auth !== 'undefined') Auth.applyUI();
}

function filtrarProdutos() {
    const termoPesquisa = document.getElementById('pesquisa').value.toLowerCase();
    const produtosFiltrados = produtos.filter(produto => {
        return produto.Nome.toLowerCase().includes(termoPesquisa) ||
            produto['ID do Produto'].toString().includes(termoPesquisa);
    });
    renderizarTabela(produtosFiltrados);
}

async function editarProduto(id) {
    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterProdutoPorId', spreadsheetId: window.SPREADSHEET_ID, data: { id: id } })
        });
        const data = await response.json();

        if (data.status === 'sucesso' && data.dados) {
            const produto = data.dados;
            document.getElementById('idProduto').value = produto['ID do Produto'];
            document.getElementById('nome').value = produto.Nome || '';
            document.getElementById('unidadeVenda').value = produto['Unidade de Venda'] || 'Un';

            const custo = parseCurrencyBRL(produto['Preço_de_custo']);
            const margemPct = parseCurrencyBRL(produto['Margem_de_lucro(%)']);
            const margemRS = parseCurrencyBRL(produto['Margem_de_lucro(R$)']);
            const preco = parseCurrencyBRL(produto['Preço_de_venda'] || produto['Preço']);

            const formatMoneyInput = (val) => val > 0 ? val.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';

            document.getElementById('precoCusto').value = formatMoneyInput(custo);
            document.getElementById('margemPct').value = margemPct || '';
            document.getElementById('margemRS').value = formatMoneyInput(margemRS);
            document.getElementById('preco').value = formatMoneyInput(preco);
            // Bug fix: sincroniza campo visível no plano Básico
            const precoBasicoEl = document.getElementById('precoBasico');
            if (precoBasicoEl) precoBasicoEl.value = formatMoneyInput(preco);
            document.getElementById('quantidade').value = parseFloat(produto.Quantidade) || 0;
            document.getElementById('descricao').value = produto.Descrição || '';
            exibirStatus({ status: 'success', mensagem: 'Campos preenchidos. Agora você pode editar.' });
            
            abrirSlideoverProduto();
        } else {
            exibirStatus({ status: 'error', mensagem: 'Produto não encontrado para edição.' });
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao obter dados do produto: ' + error.message });
    }
}

async function excluirProduto(id) {
    if (await CustomModal.confirm(`Tem certeza que deseja excluir o produto com ID ${id}?`, 'Excluir', 'Cancelar')) {
        try {
            const response = await fetch(window.MASTER_WEBHOOK_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'excluirProduto', spreadsheetId: window.SPREADSHEET_ID, data: { id: id } })
            });
            const data = await response.json();
            exibirStatus(data);
            if (data.status === 'sucesso') {
                CacheAPI.clear('cache_produtos');
                await carregarProdutos(true); // Recarrega a lista
            }
        } catch (error) {
            exibirStatus({ status: 'error', mensagem: 'Erro ao excluir o produto: ' + error.message });
        }
    }
}
