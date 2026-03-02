document.addEventListener('DOMContentLoaded', function () {
    if (SCRIPT_URL === '') {
        exibirStatus({ status: 'error', mensagem: 'Por favor, cole a URL do Apps Script no código.' });
        document.getElementById('loading').classList.add('hidden');
        return;
    }
    carregarRelatorios();
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

async function carregarRelatorios() {
    try {
        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterVendas' })
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

function renderizarKPIs(vendas) {
    const totalVendas = vendas.reduce((acc, venda) => acc + (parseFloat(venda['Total com Desconto']) || 0), 0);
    const totalTransacoes = vendas.length;
    const ticketMedio = totalTransacoes > 0 ? totalVendas / totalTransacoes : 0;

    document.getElementById('totalVendas').textContent = `R$ ${totalVendas.toFixed(2).replace('.', ',')}`;
    document.getElementById('totalTransacoes').textContent = totalTransacoes;
    document.getElementById('ticketMedio').textContent = `R$ ${ticketMedio.toFixed(2).replace('.', ',')}`;
}

/**
 * Extrai "MM/AAAA" de qualquer valor de data (string dd/mm/yyyy, timestamp, Date object).
 * Retorna null se não conseguir interpretar.
 */
function extrairMesAno(valorData) {
    if (!valorData) return null;

    // Se já for um objeto Date (ex: vindo do Google Sheets como Date)
    if (valorData instanceof Date && !isNaN(valorData)) {
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        const ano = valorData.getFullYear();
        return `${mes}/${ano}`;
    }

    const str = String(valorData).trim();

    // Formato dd/mm/yyyy ou dd/mm/yyyy hh:mm:ss
    const matchBr = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (matchBr) {
        const mes = matchBr[2].padStart(2, '0');
        const ano = matchBr[3];
        return `${mes}/${ano}`;
    }

    // Tenta parsear como Date genérico (ISO, etc.)
    const d = new Date(str);
    if (!isNaN(d)) {
        const mes = String(d.getMonth() + 1).padStart(2, '0');
        const ano = d.getFullYear();
        return `${mes}/${ano}`;
    }

    return null;
}

function renderizarGraficoVendasPorMes(vendas) {
    const vendasPorMes = vendas.reduce((acc, venda) => {
        const mesAno = extrairMesAno(venda.Data || venda['Data']);
        if (!mesAno) return acc;
        acc[mesAno] = (acc[mesAno] || 0) + (parseFloat(venda['Total com Desconto']) || 0);
        return acc;
    }, {});

    // Ordena cronologicamente (MM/AAAA → AAAA/MM para sort correto)
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
            labels: labels,
            datasets: [{
                label: 'Vendas Totais (R$)',
                data: dados,
                backgroundColor: 'rgba(52, 152, 219, 0.6)',
                borderColor: 'rgba(52, 152, 219, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

function renderizarGraficoVendasPorProduto(vendas) {
    const vendasPorProduto = vendas.reduce((acc, venda) => {
        const produto = venda.Itens || venda.itens || 'Desconhecido';
        acc[produto] = (acc[produto] || 0) + (parseFloat(venda['Total com Desconto']) || 0);
        return acc;
    }, {});

    const labels = Object.keys(vendasPorProduto);
    const dados = labels.map(produto => vendasPorProduto[produto]);

    const ctx = document.getElementById('vendasPorProdutoChart').getContext('2d');
    new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                label: 'Vendas por Produto (R$)',
                data: dados,
                backgroundColor: [
                    '#3498db', '#9b59b6', '#e74c3c', '#2ecc71', '#f1c40f', '#1abc9c', '#95a5a6'
                ],
                hoverOffset: 4
            }]
        },
        options: {
            responsive: true,
        }
    });
}
