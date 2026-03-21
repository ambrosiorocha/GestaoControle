let fornecedores = [];

document.addEventListener('DOMContentLoaded', function () {
    if (window.MASTER_WEBHOOK_URL === '' || window.MASTER_WEBHOOK_URL.includes('COLE_AQUI')) {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.MASTER_WEBHOOK_URL no config.js.' });
        return;
    }
    document.getElementById('fornecedorForm').addEventListener('submit', function (e) {
        e.preventDefault();
        execWithSpinner(document.querySelector('#fornecedorForm button[type="submit"]'), salvarFornecedor);
    });
    document.getElementById('pesquisa').addEventListener('input', filtrarFornecedores);

    carregarFornecedores();
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

async function salvarFornecedor() {

    // Chaves alinhadas com os cabeçalhos da planilha Fornecedores
    const fornecedor = {
        id: document.getElementById('idFornecedor').value || null,
        nome: document.getElementById('razaoSocial').value,       // Razão Social → nome
        telefone: document.getElementById('telefone').value,
        email: document.getElementById('email').value,
        produto: document.getElementById('categoria').value,       // Categoria → produto
        observacoes: document.getElementById('cnpj').value         // CNPJ → observacoes
    };

    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarFornecedor', spreadsheetId: window.SPREADSHEET_ID, data: fornecedor })
        });

        const data = await response.json();
        exibirStatus(data);

        if (data.status === 'sucesso') {
            document.getElementById('fornecedorForm').reset();
            document.getElementById('idFornecedor').value = '';
            await carregarFornecedores();
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro de comunicação: ' + error });
    }
}

async function carregarFornecedores() {
    const listaFornecedores = document.getElementById('listaFornecedores');
    listaFornecedores.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Carregando fornecedores...</td></tr>';

    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'listarFornecedores', spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await response.json();

        if (data.status === 'sucesso' && data.dados) {
            fornecedores = parseCompactData(data.dados);
            renderizarTabela(fornecedores);
        } else {
            listaFornecedores.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Nenhum fornecedor cadastrado.</td></tr>';
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar lista de fornecedores: ' + error.message });
        listaFornecedores.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Erro ao carregar fornecedores.</td></tr>';
    }
}

function renderizarTabela(dadosParaRenderizar) {
    const listaFornecedores = document.getElementById('listaFornecedores');
    listaFornecedores.innerHTML = '';

    if (dadosParaRenderizar.length === 0) {
        listaFornecedores.innerHTML = '<tr><td colspan="7" class="td-empty">Nenhum fornecedor encontrado.</td></tr>';
        return;
    }

    dadosParaRenderizar.forEach(f => {
        const fornecedorId = f.id || f.ID;
        const row = document.createElement('tr');
        row.innerHTML = `
            <td style="color:#94a3b8; font-size:0.78rem;">${fornecedorId}</td>
            <td style="font-weight:500;">${f.nome || ''}</td>
            <td>${f.observacoes || ''}</td>
            <td>${f.telefone || ''}</td>
            <td>${f.email || ''}</td>
            <td>${f.produto || ''}</td>
            <td>
                <div class="action-buttons">
                    <button class="edit-btn" onclick="editarFornecedor(${fornecedorId})">Editar</button>
                    <button class="delete-btn" data-admin-btn onclick="excluirFornecedor(${fornecedorId})">Excluir</button>
                </div>
            </td>
        `;
        listaFornecedores.appendChild(row);
    });
    if (typeof Auth !== 'undefined') Auth.applyUI();
}

function filtrarFornecedores() {
    const termoPesquisa = document.getElementById('pesquisa').value.toLowerCase();
    const filtrados = fornecedores.filter(f => {
        return (f.nome && f.nome.toLowerCase().includes(termoPesquisa)) ||
            (f.observacoes && f.observacoes.toLowerCase().includes(termoPesquisa));
    });
    renderizarTabela(filtrados);
}

function editarFornecedor(id) {
    const f = fornecedores.find(item => (item.id || item.ID) == id);
    if (f) {
        document.getElementById('idFornecedor').value = f.id || f.ID;
        document.getElementById('razaoSocial').value = f.nome || '';
        document.getElementById('cnpj').value = f.observacoes || '';
        document.getElementById('telefone').value = f.telefone || '';
        document.getElementById('email').value = f.email || '';
        document.getElementById('categoria').value = f.produto || '';
        exibirStatus({ status: 'success', mensagem: 'Campos preenchidos para edição.' });
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
}

async function excluirFornecedor(id) {
    if (await CustomModal.confirm(`Tem certeza que deseja excluir o fornecedor com ID ${id}?`, 'Excluir', 'Cancelar')) {
        try {
            const response = await fetch(window.MASTER_WEBHOOK_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'excluirFornecedor', spreadsheetId: window.SPREADSHEET_ID, data: { id: id } })
            });
            const data = await response.json();
            exibirStatus(data);
            if (data.status === 'sucesso') {
                await carregarFornecedores();
            }
        } catch (error) {
            exibirStatus({ status: 'error', mensagem: 'Erro ao excluir o fornecedor: ' + error.message });
        }
    }
}
