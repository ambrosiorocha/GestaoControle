let registrosFinanceiros = [];

document.addEventListener('DOMContentLoaded', function () {
    if (SCRIPT_URL === '') {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.SCRIPT_URL no config.js.' });
        return;
    }
    const caixaSelect = document.getElementById('caixa');
    // Injeção de dependência dos Caixas baseada no Plano
    if (Auth.isPlanBasico()) {
        caixaSelect.innerHTML = '<option value="Dinheiro">💵 Dinheiro Mão</option>';
        caixaSelect.value = 'Dinheiro';
        caixaSelect.disabled = true; // Trava para o plano Básico
    } else {
        caixaSelect.innerHTML = `
            <option value="Dinheiro">💵 Dinheiro (Balcão)</option>
            <option value="Conta Banco do Brasil">🏦 Conta Banco do Brasil</option>
            <option value="Conta Itaú">🏦 Conta Itaú</option>
            <option value="Conta Caixa">🏦 Conta Bradesco</option>
            <option value="Conta Nubank">🟣 Nubank Empresa</option>
        `;
    }

    // Interceptação do Formulário para o Modal de Resumo
    document.getElementById('btnPréSalvar').addEventListener('click', function () {
        if (!document.getElementById('financeiroForm').checkValidity()) {
            document.getElementById('financeiroForm').reportValidity();
            return;
        }
        prepararResumoModal();
    });

    document.getElementById('btnConfirmarSalvar').addEventListener('click', function (e) {
        execWithSpinner(this, salvarFinanceiro);
    });

    document.getElementById('filtroTipo').addEventListener('change', aplicarFiltros);
    if (Auth.isPlanBasico()) {
        const kpiGrid = document.getElementById('kpiGridFinanceiro');
        if (kpiGrid) kpiGrid.style.display = 'none';

        // Removemos o bloqueio do tipo (agora Básico pode Lançar A Receber Pago)
        // Apenas o Status "Pendente" (Contas a prazo) fica oculto.
        const selectStatus = document.getElementById('status');
        if (selectStatus) {
            for (let i = 0; i < selectStatus.options.length; i++) {
                if (selectStatus.options[i].value === 'Pendente') selectStatus.options[i].style.display = 'none';
            }
            selectStatus.value = 'Pago';
        }

        // Injeta aviso visual informando o usuário Básico sobre a limitação do fiado/contas a receber futuras
        const formFin = document.getElementById('financeiroForm');
        if (formFin) {
            const aviso = document.createElement('div');
            aviso.innerHTML = `
                <div style="background:#fff7ed; border-left:4px solid #f97316; padding:0.8rem; margin-top:1rem; border-radius:0 0.5rem 0.5rem 0; font-size:0.85rem; color:#9a3412; display:flex; align-items:center; gap:0.5rem;">
                    <span style="font-size:1.2rem;">⭐</span>
                    <div>
                        <strong>Controle de Inadimplência VIP</strong><br>
                        Sua versão atual registra entradas e saídas <b>à vista (Quitado)</b>. Para gerenciar clientes pendentes (Fiado), programar despesas futuras e receber alertas de vencimento diários, faça o upgrade para o plano <strong>PRO</strong> ou <strong>Premium</strong>.
                    </div>
                </div>
            `;
            formFin.appendChild(aviso);
        }
        if (selectStatus) {
            for (let i = 0; i < selectStatus.options.length; i++) {
                if (selectStatus.options[i].value === 'Pendente') selectStatus.options[i].style.display = 'none';
            }
            selectStatus.value = 'Pago';
        }

        // Esconder os Filtros de "Todos os Tipos" etc para forçar visualização limpa
        const filtroArea = document.getElementById('filtroTipo');
        if (filtroArea) filtroArea.style.display = 'none';
    }
    carregarFinanceiro();
});

function exibirStatus(resposta) {
    var statusMessage = document.getElementById('statusMessage');
    statusMessage.textContent = resposta.mensagem;
    statusMessage.className = '';
    if (resposta.status) statusMessage.classList.add(resposta.status);
    statusMessage.style.display = 'block';
    setTimeout(() => statusMessage.style.display = 'none', 5000);
}

// ==================== FORMATAR DATA ====================
function formatarData(valor) {
    if (!valor) return '-';
    if (valor instanceof Date && !isNaN(valor)) return valor.toLocaleDateString('pt-BR');
    const str = String(valor).trim();
    // dd/mm/yyyy
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) return str.substring(0, 10);
    // ISO: 2026-02-26T...
    const d = new Date(str);
    if (!isNaN(d)) return d.toLocaleDateString('pt-BR');
    return str;
    return str;
}

// ==================== ABRIR MODAL RESUMO ====================
function prepararResumoModal() {
    const desc = document.getElementById('descricao').value;
    const valorRaw = parseCurrencyBRL(document.getElementById('valor').value);
    const tipo = document.getElementById('tipo').value;
    const dataVenc = document.getElementById('vencimento').value;
    const categ = document.getElementById('categoria').options[document.getElementById('categoria').selectedIndex].text;
    const caixa = document.getElementById('caixa').options[document.getElementById('caixa').selectedIndex].text;
    const status = document.getElementById('status').value;

    document.getElementById('resDescricao').textContent = desc;
    document.getElementById('resVencimento').textContent = formatarData(dataVenc);
    document.getElementById('resCategoria').textContent = categ;
    document.getElementById('resCaixa').textContent = caixa;

    const divValor = document.getElementById('resValor');
    const divTipoStatus = document.getElementById('resTipoStatus');

    divValor.textContent = formatCurrencyBRL(valorRaw);

    if (tipo === 'Receber') {
        divValor.style.color = '#16a34a'; // Verde
        divTipoStatus.innerHTML = `<span style="color:#16a34a font-weight:bold;">Entrada</span> • ${status}`;
    } else {
        divValor.style.color = '#dc2626'; // Vermelho
        divTipoStatus.innerHTML = `<span style="color:#dc2626; font-weight:bold;">Despesa</span> • ${status}`;
    }

    document.getElementById('modalResumoFinanceiro').style.display = 'flex';
}

// ==================== SALVAR ====================
async function salvarFinanceiro() {
    const registro = {
        id: document.getElementById('idFinanceiro').value || null,
        descricao: document.getElementById('descricao').value,
        valor: parseCurrencyBRL(document.getElementById('valor').value),
        tipo: document.getElementById('tipo').value,
        vencimento: document.getElementById('vencimento').value,
        status: document.getElementById('status').value,
        categoria: document.getElementById('categoria').value,
        caixa: document.getElementById('caixa').value
    };
    try {
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarFinanceiro', data: registro })
        });
        const data = await response.json();
        exibirStatus(data);
        if (data.status === 'sucesso') {
            document.getElementById('financeiroForm').reset();
            document.getElementById('idFinanceiro').value = '';

            // Re-trava o Caixa se for Básico
            if (Auth.isPlanBasico()) {
                document.getElementById('caixa').value = 'Dinheiro';
            }

            document.getElementById('modalResumoFinanceiro').style.display = 'none';
            await carregarFinanceiro();
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro de comunicação: ' + error });
    }
}

// ==================== CARREGAR ====================
async function carregarFinanceiro() {
    const listaFinanceiro = document.getElementById('listaFinanceiro');
    listaFinanceiro.innerHTML = '<tr><td colspan="8" class="table-cell p-4 text-center">Carregando...</td></tr>';
    try {
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterFinanceiro' })
        });
        const data = await response.json();
        if (data.status === 'sucesso' && data.dados) {
            registrosFinanceiros = parseCompactData(data.dados);
            aplicarFiltros();
        } else {
            registrosFinanceiros = [];
            listaFinanceiro.innerHTML = '<tr><td colspan="8" class="table-cell p-4 text-center">Nenhum registro encontrado.</td></tr>';
            atualizarResumo([]);
        }
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar dados: ' + error.message });
    }
}

// ==================== FILTROS ====================
function aplicarFiltros() {
    const filtroTipo = document.getElementById('filtroTipo').value;
    const dataInicio = document.getElementById('filtroDataInicio').value;
    const dataFim = document.getElementById('filtroDataFim').value;

    let filtrados = [...registrosFinanceiros];

    // Ocultar sempre registros financeiros de vendas Estornadas da UI limpa
    filtrados = filtrados.filter(r => {
        const isEstornado = r.status === 'Estornado' || r.status === 'Estornada';
        const isVenda = String(r.descricao).toLowerCase().includes('venda #') || (r.categoria && String(r.categoria).toLowerCase() === 'venda');
        return !(isEstornado && isVenda);
    });
    // Removemos os filtros rígidos do Básico. Agora eles podem ver as próprias Receitas lançadas à vista.

    if (filtroTipo && !(typeof Auth !== 'undefined' && Auth.isPlanBasico())) {
        filtrados = filtrados.filter(r => r.tipo === filtroTipo);
    }
    if (dataInicio) {
        const inicio = new Date(dataInicio + 'T00:00:00');
        filtrados = filtrados.filter(r => {
            const d = parseDataVencimento(r.vencimento);
            return d && d >= inicio;
        });
    }
    if (dataFim) {
        const fim = new Date(dataFim + 'T23:59:59');
        filtrados = filtrados.filter(r => {
            const d = parseDataVencimento(r.vencimento);
            return d && d <= fim;
        });
    }

    renderizarTabela(filtrados);
    atualizarResumo(filtrados);
}

function limparFiltros() {
    document.getElementById('filtroDataInicio').value = '';
    document.getElementById('filtroDataFim').value = '';
    document.getElementById('filtroTipo').value = '';
    aplicarFiltros();
}

function parseDataVencimento(valor) {
    if (!valor) return null;
    if (valor instanceof Date) return valor;
    const str = String(valor).trim();
    // dd/mm/yyyy
    const matchBr = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (matchBr) return new Date(`${matchBr[3]}-${matchBr[2].padStart(2, '0')}-${matchBr[1].padStart(2, '0')}`);
    const d = new Date(str);
    return isNaN(d) ? null : d;
}

// ==================== RENDERIZAR ====================
function renderizarTabela(dados) {
    const listaFinanceiro = document.getElementById('listaFinanceiro');
    listaFinanceiro.innerHTML = '';

    if (dados.length === 0) {
        listaFinanceiro.innerHTML = '<tr><td colspan="8" class="table-cell p-4 text-center">Nenhum registro encontrado.</td></tr>';
        return;
    }

    const trClasses = "table-row";
    const tdClasses = "table-cell align-middle";

    dados.forEach(r => {
        const registroId = r.id || r.ID;
        const row = document.createElement('tr');
        row.className = trClasses;
        const isEstornado = r.status === 'Estornado' || r.status === 'Estornada';
        const isVenda = String(r.descricao).toLowerCase().includes('venda #') || (r.categoria && String(r.categoria).toLowerCase() === 'venda');
        const isPago = r.status === 'Pago' || r.status === 'Quitado';

        let statusClass = 'font-bold ';
        if (isEstornado) statusClass += 'text-gray-500 line-through';
        else if (isPago) statusClass += 'text-green-600';
        else statusClass += 'text-red-600';

        const tipoClass = r.tipo === 'Receber' ? 'text-blue-600' : 'text-orange-600';

        let btnBaixar = '';
        if (!isEstornado) {
            btnBaixar = isPago
                ? `<span style="color:#16a34a; font-size:0.78rem;">✅ Quitado</span>`
                : `<button class="edit-btn" style="background:#16a34a; color:white; font-size:0.75rem;" onclick="execWithSpinner(this, () => baixarLancamento(${registroId}, '${r.tipo}'))">${r.tipo === 'Receber' ? '✅ Receber' : '✅ Pagar'}</button>`;
        } else {
            btnBaixar = `<span style="color:#64748b; font-size:0.78rem;">❌ Cancelado</span>`;
        }

        let acoesHtml = '';
        if (isVenda) {
            acoesHtml = `<span style="font-size:0.75rem; color:#64748b; font-style:italic;">Gerencie em Vendas</span>`;
            if (!isPago && !isEstornado) acoesHtml = btnBaixar + " " + acoesHtml;
        } else {
            if (!isEstornado) {
                // Se está ativo, o usuário só pode Estornar (Cancelar) para então Editar ou Excluir.
                acoesHtml = `
                    ${btnBaixar}
                    <button class="delete-btn" style="background:#f59e0b; color:white;" onclick="estornarLancamentoManual(${registroId})">Estornar</button>
                `;
            } else {
                // Se já está estornado, pode re-editar (voltando a ficar ativo e salvo) ou Excluir de vez da base
                acoesHtml = `
                    ${btnBaixar}
                    <button class="edit-btn" style="background:#3b82f6; color:white;" onclick="editarFinanceiro(${registroId})">Editar</button>
                    <button class="delete-btn" onclick="excluirFinanceiro(${registroId})">Excluir</button>
                `;
            }
        }

        row.innerHTML = `
            <td class="${tdClasses}">${registroId}</td>
            <td class="${tdClasses}">${r.descricao || ''}</td>
            <td class="${tdClasses} font-semibold">${formatCurrencyBRL(r.valor)}</td>
            <td class="${tdClasses} ${tipoClass}">${r.tipo || ''}</td>
            <td class="${tdClasses}">${formatarData(r.vencimento)}</td>
            <td class="${tdClasses} ${statusClass}">${r.status || ''}</td>
            <td class="${tdClasses}">${r.categoria || ''}</td>
            <td class="${tdClasses}">
                <div class="action-buttons" style="display:flex; gap:0.35rem; flex-wrap:wrap; align-items:center;">
                    ${acoesHtml}
                </div>
            </td>
        `;
        listaFinanceiro.appendChild(row);
    });
}

// ==================== RESUMO KPIs ====================
function atualizarResumo(dados) {
    const totalPagar = dados.filter(r => r.tipo === 'Pagar' && r.status !== 'Pago').reduce((acc, r) => acc + parseFloat(r.valor || 0), 0);
    const totalReceber = dados.filter(r => r.tipo === 'Receber' && r.status !== 'Pago').reduce((acc, r) => acc + parseFloat(r.valor || 0), 0);
    const saldo = totalReceber - totalPagar;
    document.getElementById('totalPagar').textContent = formatCurrencyBRL(totalPagar);
    document.getElementById('totalReceber').textContent = formatCurrencyBRL(totalReceber);
    document.getElementById('saldoPendencias').textContent = formatCurrencyBRL(saldo);
}

// ==================== BAIXAR LANÇAMENTO ====================
async function baixarLancamento(id, tipo) {
    const acao = tipo === 'Receber' ? 'receber' : 'pagar';
    const r = registrosFinanceiros.find(item => (item.id || item.ID) == id);
    const valor = r ? parseFloat(r.valor || 0) : 0;
    const valorStr = formatCurrencyBRL(valor);
    const msg = tipo === 'Receber' ? `Confirmar recebimento de ${valorStr}?` : `Confirmar pagamento de ${valorStr}?`;

    if (!(await CustomModal.confirm(msg, 'Confirmar', 'Cancelar'))) return;
    try {
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'baixarLancamento', data: { id: id } })
        });
        const data = await response.json();
        exibirStatus(data);
        if (data.status === 'sucesso') await carregarFinanceiro();
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro: ' + error.message });
    }
}

// ==================== ESTORNAR LANÇAMENTO MANUAL ====================
async function estornarLancamentoManual(id) {
    if (!(await CustomModal.confirm(`Cancelar Lançamento #${id}?\n\nEste lançamento ficará "Estornado" (inativo) e poderá ser editado ou excluído definitivamente depois.`, 'Cancelar Lançamento', 'Voltar'))) return;

    const r = registrosFinanceiros.find(item => (item.id || item.ID) == id);
    if (!r) return;

    const registroCancelado = {
        id: r.id || r.ID,
        descricao: r.descricao,
        valor: parseFloat(r.valor),
        tipo: r.tipo,
        vencimento: r.vencimento,
        status: 'Estornado',
        categoria: r.categoria
    };

    try {
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarFinanceiro', data: registroCancelado })
        });
        const data = await response.json();
        exibirStatus(data);
        if (data.status === 'sucesso') await carregarFinanceiro();
    } catch (error) {
        exibirStatus({ status: 'error', mensagem: 'Erro ao cancelar lançamento: ' + error.message });
    }
}

// ==================== EDITAR / EXCLUIR ====================
function editarFinanceiro(id) {
    const r = registrosFinanceiros.find(item => (item.id || item.ID) == id);
    if (r) {
        document.getElementById('idFinanceiro').value = r.id || r.ID;
        document.getElementById('descricao').value = r.descricao || '';
        document.getElementById('valor').value = r.valor || '';
        document.getElementById('categoria').value = r.categoria || '';
        document.getElementById('vencimento').value = r.vencimento || '';

        const isBsc = typeof Auth !== 'undefined' && Auth.isPlanBasico();
        document.getElementById('tipo').value = (r.tipo || '');
        document.getElementById('status').value = isBsc ? 'Pago' : (r.status || '');

        // Visual warning for 'Pendente' status on basic plans
        const warningElement = document.getElementById('pendenteWarning');
        if (isBsc && r.status === 'Pendente') {
            if (warningElement) {
                warningElement.style.display = 'block';
            } else {
                const form = document.getElementById('financeiroForm');
                const newWarning = document.createElement('p');
                newWarning.id = 'pendenteWarning';
                newWarning.className = 'text-red-500 text-sm mt-2';
                newWarning.textContent = 'O plano Básico não permite agendar contas Pendentes. Por favor, selecione "Pago" ou faça upgrade para o plano PRO/Premium.';
                form.parentNode.insertBefore(newWarning, form.nextSibling);
            }
        } else {
            if (warningElement) {
                warningElement.style.display = 'none';
            }
        }

        exibirStatus({ status: 'success', mensagem: 'Registro carregado para edição.' });
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
}

async function excluirFinanceiro(id) {
    if (await CustomModal.confirm(`Excluir registro ID ${id}?`, 'Excluir', 'Cancelar')) {
        try {
            const response = await fetch(window.SCRIPT_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'excluirFinanceiro', data: { id: id } })
            });
            const data = await response.json();
            exibirStatus(data);
            if (data.status === 'sucesso') await carregarFinanceiro();
        } catch (error) {
            exibirStatus({ status: 'error', mensagem: 'Erro ao excluir: ' + error.message });
        }
    }
}
