let registrosFinanceiros = [];

document.addEventListener('DOMContentLoaded', function () {
    if (window.MASTER_WEBHOOK_URL === '' || window.MASTER_WEBHOOK_URL.includes('COLE_AQUI')) {
        exibirStatus({ status: 'error', mensagem: 'Configure a window.MASTER_WEBHOOK_URL no config.js.' });
        return;
    }
    const caixaSelect = document.getElementById('caixa');
    // Injeção de dependência dos Caixas baseada no Plano
    // Injeção de dependência dos Caixas baseada no Plano
    if (Auth.isPlanBasico()) {
        caixaSelect.innerHTML = '<option value="Dinheiro">💵 Dinheiro</option>';
        caixaSelect.value = 'Dinheiro';
        caixaSelect.disabled = true; // Trava para o plano Básico
    } else {
        caixaSelect.innerHTML = `
            <option value="Dinheiro">💵 Dinheiro</option>
            <option value="Conta Banco do Brasil">🏦 Conta Banco do Brasil</option>
            <option value="Conta Itaú">🏦 Conta Itaú</option>
            <option value="Conta Caixa">🏦 Conta Bradesco</option>
            <option value="Conta Nubank">🟣 Nubank Empresa</option>
        `;
    }

    // Default Vencimento to today
    document.getElementById('vencimento').value = new Date().toISOString().split('T')[0];

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
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarFinanceiro', spreadsheetId: window.SPREADSHEET_ID, data: registro })
        });
        const data = await response.json();
        exibirStatus(data);
        if (data.status === 'sucesso') {
            document.getElementById('financeiroForm').reset();
            document.getElementById('idFinanceiro').value = '';
            document.getElementById('vencimento').value = new Date().toISOString().split('T')[0];

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
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterFinanceiro', spreadsheetId: window.SPREADSHEET_ID })
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
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'baixarLancamento', spreadsheetId: window.SPREADSHEET_ID, data: { id: id } })
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
        const response = await fetch(window.MASTER_WEBHOOK_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'salvarFinanceiro', spreadsheetId: window.SPREADSHEET_ID, data: registroCancelado })
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
            const response = await fetch(window.MASTER_WEBHOOK_URL, {
                method: 'POST',
                body: JSON.stringify({ action: 'excluirFinanceiro', spreadsheetId: window.SPREADSHEET_ID, data: { id: id } })
            });
            const data = await response.json();
            exibirStatus(data);
            if (data.status === 'sucesso') await carregarFinanceiro();
        } catch (error) {
            exibirStatus({ status: 'error', mensagem: 'Erro ao excluir: ' + error.message });
        }
    }
}

// ==================== ABAS DO FINANCEIRO ====================
function alternarAbaFinanceiro(aba) {
    const tabLancamentos = document.getElementById('tabLancamentos');
    const tabExtratos = document.getElementById('tabExtratos');
    const btnLanc = document.getElementById('btnTabLancamentos');
    const btnExt = document.getElementById('btnTabExtratos');

    if (aba === 'lancamentos') {
        tabLancamentos.style.display = 'block';
        tabExtratos.style.display = 'none';

        btnLanc.style.background = 'white';
        btnLanc.style.color = '#334155';
        btnLanc.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';

        btnExt.style.background = 'transparent';
        btnExt.style.color = '#64748b';
        btnExt.style.boxShadow = 'none';

    } else {
        tabLancamentos.style.display = 'none';
        tabExtratos.style.display = 'block';

        btnExt.style.background = 'white';
        btnExt.style.color = '#334155';
        btnExt.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';

        btnLanc.style.background = 'transparent';
        btnLanc.style.color = '#64748b';
        btnLanc.style.boxShadow = 'none';

        renderizarExtratosCaixa();
    }
}

// ==================== EXTRATO DE CAIXAS ====================
window.renderizarExtratosCaixa = function () {
    const isBsc = typeof Auth !== 'undefined' && Auth.isPlanBasico();

    // Identificar caixas existentes nos registros e definir listagem padrão
    const todosCaixasRegistrados = [...new Set(registrosFinanceiros.map(r => r.caixa || 'Dinheiro'))];
    let caixasAtivos = isBsc ? ['Dinheiro'] : ['Dinheiro', 'Conta Banco do Brasil', 'Conta Itaú', 'Conta Caixa', 'Conta Nubank'];

    // Adiciona caixas históricos caso existam nos dados mas não na listagem padrão
    todosCaixasRegistrados.forEach(c => {
        if (!caixasAtivos.includes(c)) caixasAtivos.push(c);
    });

    const kpiGrid = document.getElementById('kpiGridCaixas');
    const selectFiltro = document.getElementById('filtroExtratoCaixa');
    const dtInicioStr = document.getElementById('filtroExtratoInicio')?.value;
    const dtFimStr = document.getElementById('filtroExtratoFim')?.value;

    kpiGrid.innerHTML = '';

    // Se ainda não inicializamos as opções do select
    if (selectFiltro.options.length === 0) {
        selectFiltro.innerHTML = '<option value="TODOS">🧾 Mostrar Todos os Caixas</option>';
        caixasAtivos.forEach(c => {
            selectFiltro.innerHTML += `<option value="${c}">${c}</option>`;
        });
    }

    // Filtrar base para calcular saldos parciais nos Cards
    let baseParaCaixas = [...registrosFinanceiros];
    if (dtInicioStr) {
        const inicio = new Date(dtInicioStr + 'T00:00:00');
        baseParaCaixas = baseParaCaixas.filter(r => {
            const d = typeof parseDataVencimento === 'function' ? parseDataVencimento(r.vencimento) : new Date(r.vencimento);
            return d && d >= inicio;
        });
    }
    if (dtFimStr) {
        const fim = new Date(dtFimStr + 'T23:59:59');
        baseParaCaixas = baseParaCaixas.filter(r => {
            const d = typeof parseDataVencimento === 'function' ? parseDataVencimento(r.vencimento) : new Date(r.vencimento);
            return d && d <= fim;
        });
    }

    // Calcula saldo para cada caixa e gera os cards
    caixasAtivos.forEach(nomeCaixa => {
        const registrosDesteCaixa = baseParaCaixas.filter(r => (r.caixa || 'Dinheiro') === nomeCaixa);

        let entradas = 0;
        let saidas = 0;

        registrosDesteCaixa.forEach(r => {
            const val = parseFloat(r.valor) || 0;
            const statusPago = (r.status === 'Pago' || r.status === 'Quitado');
            const naoEstornado = !(r.status === 'Estornado' || r.status === 'Estornada');

            // Só conta no Extrato do Caixa o dinheiro que DE FATO ENTROU/SAIU
            // Desconsidera "Pendente" para o cálculo do saldo final no caixa
            if (statusPago && naoEstornado) {
                if (r.tipo === 'Receber') entradas += val;
                else saidas += val;
            }
        });

        const saldoFinal = entradas - saidas;

        // No grid, não mostrar cards de bancos se for básico exceto "Dinheiro"
        if (isBsc && nomeCaixa !== 'Dinheiro') return;

        const cardHtml = `
            <div class="kpi-card" style="background:#f8fafc; border-bottom: 4px solid ${saldoFinal >= 0 ? '#16a34a' : '#ef4444'}">
                <h4 style="color:#475569; font-weight:600; font-size:0.95rem;">🪙 ${nomeCaixa}</h4>
                <p style="color:${saldoFinal >= 0 ? '#16a34a' : '#ef4444'}; font-size:1.5rem; font-weight:700;">${formatCurrencyBRL(saldoFinal)}</p>
                <div style="display:flex; justify-content:space-between; margin-top:0.5rem; font-size:0.8rem;">
                    <span style="color:#16a34a">Entrou: ${formatCurrencyBRL(entradas)}</span>
                    <span style="color:#ef4444">Saiu: ${formatCurrencyBRL(saidas)}</span>
                </div>
            </div>
        `;
        kpiGrid.innerHTML += cardHtml;
    });

    renderizarTabelaExtrato();
}

window.renderizarTabelaExtrato = function () {
    const isBsc = typeof Auth !== 'undefined' && Auth.isPlanBasico();
    const tableWrapper = document.getElementById('historicoTableWrapper');
    const upgradeWall = document.getElementById('historicoUpgradeWall');

    if (isBsc) {
        if (tableWrapper) tableWrapper.style.display = 'none';
        if (upgradeWall) upgradeWall.style.display = 'block';
        return; // Basic users stop here, they only see the KPIs
    } else {
        if (tableWrapper) tableWrapper.style.display = 'block';
        if (upgradeWall) upgradeWall.style.display = 'none';
    }

    const selectFiltro = document.getElementById('filtroExtratoCaixa');
    const dtInicioStr = document.getElementById('filtroExtratoInicio')?.value;
    const dtFimStr = document.getElementById('filtroExtratoFim')?.value;
    const caixaFiltro = selectFiltro.value;

    const tbody = document.getElementById('listaExtratoCaixa');
    tbody.innerHTML = '';

    let filtrados = registrosFinanceiros.filter(r => {
        const isEstornado = r.status === 'Estornado' || r.status === 'Estornada';
        const isVenda = String(r.descricao).toLowerCase().includes('venda #') || (r.categoria && String(r.categoria).toLowerCase() === 'venda');
        return !(isEstornado && isVenda);
    });

    if (dtInicioStr) {
        const inicio = new Date(dtInicioStr + 'T00:00:00');
        filtrados = filtrados.filter(r => {
            const d = typeof parseDataVencimento === 'function' ? parseDataVencimento(r.vencimento) : new Date(r.vencimento);
            return d && d >= inicio;
        });
    }
    if (dtFimStr) {
        const fim = new Date(dtFimStr + 'T23:59:59');
        filtrados = filtrados.filter(r => {
            const d = typeof parseDataVencimento === 'function' ? parseDataVencimento(r.vencimento) : new Date(r.vencimento);
            return d && d <= fim;
        });
    }

    if (caixaFiltro !== 'TODOS') {
        filtrados = filtrados.filter(r => (r.caixa || 'Dinheiro') === caixaFiltro);
    }

    if (filtrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:1.5rem; color:#64748b;">Nenhuma movimentação neste período ou caixa.</td></tr>';
        return;
    }

    filtrados.forEach(r => {
        const tr = document.createElement('tr');

        const val = parseFloat(r.valor) || 0;
        const isEntrada = r.tipo === 'Receber';
        const isCancelado = r.status === 'Estornado' || r.status === 'Estornada';
        const isPendente = r.status === 'Pendente';

        let colorType = isEntrada ? '#16a34a' : '#ef4444';
        let spanValue = formatCurrencyBRL(val);

        if (isCancelado) {
            colorType = '#94a3b8';
            spanValue = `<span style="text-decoration:line-through;">${spanValue}</span>`;
        } else if (isPendente) {
            colorType = '#d97706';
        }

        const dataVencStr = formatarData(r.vencimento);
        const categStr = r.categoria || '-';
        const descStr = r.descricao || 'Sem descrição';

        tr.innerHTML = `
            <td>${dataVencStr}</td>
            <td>
                <strong>${descStr}</strong>
                ${isPendente ? '<br><span style="font-size:0.75rem; background:#fffbeb; color:#d97706; padding:1px 4px; border-radius:4px; font-weight:600;">&#128336; Futuro/Pendente</span>' : ''}
                ${isCancelado ? '<br><span style="font-size:0.75rem; background:#f1f5f9; color:#64748b; padding:1px 4px; border-radius:4px; font-weight:600;">❌ Cancelado</span>' : ''}
                ${caixaFiltro === 'TODOS' ? `<br><span style="font-size:0.75rem; color:#64748b;">🏦 ${r.caixa || 'Dinheiro'}</span>` : ''}
            </td>
            <td><span style="background:#f1f5f9; padding:2px 6px; border-radius:4px; font-size:0.8rem;">${categStr}</span></td>
            <td><span style="color:${colorType}; font-weight:600;">${isEntrada ? 'Entrada' : 'Saída'}</span></td>
            <td style="text-align:right; font-weight:700; color:${colorType};">${isEntrada ? '+' : '-'} ${spanValue}</td>
        `;
        tbody.appendChild(tr);
    });
}
