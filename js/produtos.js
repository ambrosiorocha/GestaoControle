let produtos = []; // Array global para armazenar os produtos

document.addEventListener('DOMContentLoaded', function () {
    if (SCRIPT_URL === '') {
        exibirStatus({ status: 'error', mensagem: 'Por favor, cole a URL do Apps Script no código.' });
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

    carregarProdutos();
});

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
        precoCusto: parseFloat(document.getElementById('precoCusto').value) || 0,
        margemPct: parseFloat(document.getElementById('margemPct').value) || 0,
        margemRS: parseFloat(document.getElementById('margemRS').value) || 0,
        preco: parseFloat(document.getElementById('preco').value) || 0,
        quantidade: parseFloat(document.getElementById('quantidade').value) || 0,
        descricao: document.getElementById('descricao').value
    };

    try {
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarProduto', data: produto })
        });

        const data = await response.json();
        exibirStatus(data);

        if (data.status === 'sucesso') {
            document.getElementById('produtoForm').reset();
            document.getElementById('idProduto').value = '';
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
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterProdutos' })
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

    if (produtosParaRenderizar.length === 0) {
        listaProdutos.innerHTML = '<tr><td colspan="9" class="table-cell p-4 text-center">Nenhum produto encontrado.</td></tr>';
        return;
    }

    produtosParaRenderizar.forEach((produto, idx) => {
        const precoVenda = parseFloat(String(produto['Preço_de_venda'] || produto['Preço'] || 0).replace('R$', '').replace(/\./g, '').replace(',', '.').trim()) || 0;
        const precoCusto = parseFloat(String(produto['Preço_de_custo'] || 0).replace('R$', '').replace(/\./g, '').replace(',', '.').trim()) || 0;
        const margemPct = parseFloat(String(produto['Margem_de_lucro(%)'] || 0).replace('%', '').replace(',', '.').trim()) || 0;
        const id = produto['ID do Produto'];

        const row = document.createElement('tr');
        row.className = idx % 2 === 0 ? '' : 'bg-gray-50';
        row.style.cssText = 'transition: background 0.15s;';
        row.onmouseover = () => row.style.background = '#eff6ff';
        row.onmouseout = () => row.style.background = idx % 2 === 0 ? '' : '#f9fafb';

        row.innerHTML = `
            <td class="table-cell px-3 py-2 text-xs text-gray-500">${id}</td>
            <td class="table-cell px-3 py-2 font-medium">${produto.Nome || ''}</td>
            <td class="table-cell px-3 py-2 text-center">${produto['Unidade de Venda'] || ''}</td>
            <td class="table-cell px-3 py-2 text-right text-gray-600">${precoCusto > 0 ? 'R$ ' + precoCusto.toFixed(2).replace('.', ',') : '-'}</td>
            <td class="table-cell px-3 py-2 text-center ${margemPct > 0 ? 'text-green-700 font-semibold' : 'text-gray-400'}">${margemPct > 0 ? margemPct.toFixed(1) + '%' : '-'}</td>
            <td class="table-cell px-3 py-2 text-right font-semibold text-blue-700">R$ ${precoVenda.toFixed(2).replace('.', ',')}</td>
            <td class="table-cell px-3 py-2 text-center">${produto.Quantidade ?? ''}</td>
            <td class="table-cell px-3 py-2 text-gray-500 text-xs">${produto.Descrição || ''}</td>
            <td class="table-cell px-3 py-2">
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
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterProdutoPorId', data: { id: id } })
        });
        const data = await response.json();

        if (data.status === 'sucesso' && data.dados) {
            const produto = data.dados;
            document.getElementById('idProduto').value = produto['ID do Produto'];
            document.getElementById('nome').value = produto.Nome || '';
            document.getElementById('unidadeVenda').value = produto['Unidade de Venda'] || 'Un';

            const parseNum = (v) => { const s = String(v || 0).replace('R$', '').replace(/\./g, '').replace(',', '.').trim(); const n = parseFloat(s); return isNaN(n) ? 0 : n; };

            const custo = parseNum(produto['Preço_de_custo']);
            const margemPct = parseNum(produto['Margem_de_lucro(%)']);
            const margemRS = parseNum(produto['Margem_de_lucro(R$)']);
            const preco = parseNum(produto['Preço_de_venda'] || produto['Preço']);

            document.getElementById('precoCusto').value = custo || '';
            document.getElementById('margemPct').value = margemPct || '';
            document.getElementById('margemRS').value = margemRS || '';
            document.getElementById('preco').value = preco || '';
            document.getElementById('quantidade').value = parseNum(produto.Quantidade);
            document.getElementById('descricao').value = produto.Descrição || '';
            exibirStatus({ status: 'success', mensagem: 'Campos preenchidos. Agora você pode editar.' });
        } else {
            exibirStatus({ status: 'error', mensagem: 'Produto não encontrado para edição.' });
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao obter dados do produto: ' + error.message });
    }
}

async function excluirProduto(id) {
    if (confirm(`Tem certeza que deseja excluir o produto com ID ${id}?`)) {
        try {
            const response = await fetch(window.SCRIPT_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'excluirProduto', data: { id: id } })
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
