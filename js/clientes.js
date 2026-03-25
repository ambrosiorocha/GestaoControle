let clientes = [];

document.addEventListener('DOMContentLoaded', function () {
    if (window.MASTER_WEBHOOK_URL === '' || window.MASTER_WEBHOOK_URL.includes('COLE_AQUI')) {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.MASTER_WEBHOOK_URL no config.js.' });
        return;
    }
    document.getElementById('clienteForm').addEventListener('submit', function (e) {
        e.preventDefault();
        execWithSpinner(document.querySelector('#clienteForm button[type="submit"]'), salvarCliente);
    });
    document.getElementById('pesquisa').addEventListener('input', filtrarClientes);

    var btnSync = document.getElementById('btnSincronizar');
    if (btnSync) {
        btnSync.addEventListener('click', function (e) {
            execWithSpinner(e.target, async () => { await carregarClientes(true); });
        });
    }

    // Evento para fechar modal com tecla ESC
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') fecharSlideoverCliente();
    });

    carregarClientes();
});

// ================================
// CONTROLE DO SLIDE-OVER (GAVETA)
// ================================
function abrirSlideoverCliente() {
    const overlay = document.getElementById('slideoverCliente');
    const container = document.getElementById('slideoverContainerCliente');
    
    // Se abrir sem ID, é um "Novo"
    const id = document.getElementById('idCliente').value;
    if (!id) {
        document.getElementById('clienteForm').reset();
        document.getElementById('slideoverLabelCliente').textContent = 'Novo Cliente';
    } else {
        document.getElementById('slideoverLabelCliente').textContent = 'Editar Cliente';
    }

    overlay.classList.add('open');
    container.classList.add('open');
    document.body.classList.add('slideover-open');
}

function fecharSlideoverCliente() {
    const overlay = document.getElementById('slideoverCliente');
    const container = document.getElementById('slideoverContainerCliente');
    
    overlay.classList.remove('open');
    container.classList.remove('open');
    document.body.classList.remove('slideover-open');
    
    // Limpar formulário e ID ao fechar
    setTimeout(() => {
        document.getElementById('clienteForm').reset();
        document.getElementById('idCliente').value = '';
    }, 300);
}

function validarFechamentoSlideoverCliente(e) {
    if (e.target.id === 'slideoverCliente') {
        fecharSlideoverCliente();
    }
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

async function salvarCliente() {

    // Chaves alinhadas com os cabeçalhos da planilha (minúsculas, sem acentos)
    const cliente = {
        id: document.getElementById('idCliente').value || null,
        nome: document.getElementById('nome').value,
        telefone: document.getElementById('telefone').value,
        email: document.getElementById('email').value,
        endereco: document.getElementById('endereco').value,
        observacoes: document.getElementById('cpfCnpj').value // CPF/CNPJ salvo em observacoes
    };

    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarCliente', spreadsheetId: window.SPREADSHEET_ID, data: cliente })
        });

        const data = await response.json();
        exibirStatus(data);

        if (data.status === 'sucesso') {
            document.getElementById('clienteForm').reset();
            document.getElementById('idCliente').value = '';
            fecharSlideoverCliente();
            await carregarClientes(true); // Força sincronização para atualizar cache
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro de comunicação: ' + error });
    }
}

async function carregarClientes(forceSync = false) {
    const listaClientes = document.getElementById('listaClientes');

    if (!forceSync) {
        const cached = CacheAPI.get('cache_clientes');
        if (cached) {
            clientes = cached;
            renderizarTabela(clientes);
            return;
        }
    }

    listaClientes.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Carregando clientes...</td></tr>';

    try {
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterClientes', spreadsheetId: window.SPREADSHEET_ID })
        });
        const data = await response.json();

        if (data.status === 'sucesso' && data.dados) {
            clientes = parseCompactData(data.dados);
            CacheAPI.set('cache_clientes', clientes);
            if (clientes.length > 0) {
                renderizarTabela(clientes);
            } else {
                listaClientes.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Nenhum cliente cadastrado.</td></tr>';
            }
        } else {
            listaClientes.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Nenhum cliente cadastrado.</td></tr>';
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar lista de clientes: ' + error.message });
        listaClientes.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Erro ao carregar clientes.</td></tr>';
    }
}

function renderizarTabela(dadosParaRenderizar) {
    const listaClientes = document.getElementById('listaClientes');
    listaClientes.innerHTML = '';

    if (dadosParaRenderizar.length === 0) {
        listaClientes.innerHTML = '<tr><td colspan="7" class="table-cell p-4 text-center">Nenhum cliente encontrado.</td></tr>';
        return;
    }

    const trClasses = "table-row";
    const tdClasses = "table-cell align-middle";

    dadosParaRenderizar.forEach(cliente => {
        const clienteId = cliente.id || cliente.ID;
        const row = document.createElement('tr');
        row.className = trClasses;
        row.innerHTML = `
            <td class="${tdClasses}">${clienteId}</td>
            <td class="${tdClasses}">${cliente.nome || ''}</td>
            <td class="${tdClasses}">${cliente.observacoes || ''}</td>
            <td class="${tdClasses}">${cliente.telefone || ''}</td>
            <td class="${tdClasses}">${cliente.email || ''}</td>
            <td class="${tdClasses}">${cliente.endereco || ''}</td>
            <td class="${tdClasses}">
                <div class="action-buttons">
                    <button class="edit-btn" onclick="editarCliente(${clienteId})">Editar</button>
                    <button class="delete-btn" data-admin-btn onclick="excluirCliente(${clienteId})">Excluir</button>
                </div>
            </td>
        `;
        listaClientes.appendChild(row);
    });
    if (typeof Auth !== 'undefined') Auth.applyUI();
}

function filtrarClientes() {
    const termoPesquisa = document.getElementById('pesquisa').value.toLowerCase();
    const filtrados = clientes.filter(c => {
        return (c.nome && c.nome.toLowerCase().includes(termoPesquisa)) ||
            (c.observacoes && c.observacoes.toLowerCase().includes(termoPesquisa));
    });
    renderizarTabela(filtrados);
}

function editarCliente(id) {
    const cliente = clientes.find(c => (c.id || c.ID) == id);
    if (cliente) {
        document.getElementById('idCliente').value = cliente.id || cliente.ID;
        document.getElementById('nome').value = cliente.nome || '';
        document.getElementById('cpfCnpj').value = cliente.observacoes || '';
        document.getElementById('telefone').value = cliente.telefone || '';
        document.getElementById('email').value = cliente.email || '';
        document.getElementById('endereco').value = cliente.endereco || '';
        exibirStatus({ status: 'success', mensagem: 'Campos preenchidos para edição.' });
        
        abrirSlideoverCliente();
    }
}

async function excluirCliente(id) {
    if (await CustomModal.confirm(`Tem certeza que deseja excluir o cliente com ID ${id}?`, 'Excluir', 'Cancelar')) {
        try {
            const response = await fetch(window.MASTER_WEBHOOK_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'excluirCliente', spreadsheetId: window.SPREADSHEET_ID, data: { id: id } })
            });
            const data = await response.json();
            exibirStatus(data);
            if (data.status === 'sucesso') {
                CacheAPI.clear('cache_clientes'); // Limpa o cache após exclusão
                await carregarClientes(true);
            }
        } catch (error) {
            exibirStatus({ status: 'error', mensagem: 'Erro ao excluir o cliente: ' + error.message });
        }
    }
}
