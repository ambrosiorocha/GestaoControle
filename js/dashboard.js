document.addEventListener('DOMContentLoaded', function () {
    if (SCRIPT_URL === '') {
        exibirStatus({ status: 'error', mensagem: 'Por favor, cole a URL do Apps Script no código.' });
        document.getElementById('loading').classList.add('hidden');
        return;
    }
    carregarDadosDashboard();
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

async function carregarDadosDashboard() {
    try {
        const hoje = new Date();
        const umAnoAtras = new Date(hoje.setFullYear(hoje.getFullYear() - 1));
        const dataInicioStr = umAnoAtras.toISOString().split('T')[0];

        const response = await fetch(window.SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'obterVendas', data: { dataInicio: dataInicioStr } })
        });
        const data = await response.json();

        document.getElementById('loading').classList.add('hidden');
        document.getElementById('dashboard-content').classList.remove('hidden');

        if (data.status === 'sucesso' && data.dados) {
            const vendas = parseCompactData(data.dados);
            if (vendas.length === 0) {
                exibirStatus({ status: 'info', mensagem: 'Nenhuma venda encontrada no período.' });
                // Zerar KPIs
                document.getElementById('totalVendas').textContent = 'R$ 0,00';
                document.getElementById('totalTransacoes').textContent = '0';
                document.getElementById('ticketMedio').textContent = 'R$ 0,00';
                return;
            }

            const totalVendas = vendas.reduce((acc, venda) => {
                let val = venda['Total com Desconto'];
                if (typeof val === 'string') val = String(val).replace(/[R\$\s\.]/g, '').replace(',', '.');
                return acc + (parseFloat(val) || 0);
            }, 0);
            const totalTransacoes = vendas.length;
            const ticketMedio = totalTransacoes > 0 ? totalVendas / totalTransacoes : 0;

            document.getElementById('totalVendas').textContent = `R$ ${totalVendas.toFixed(2).replace('.', ',')}`;
            document.getElementById('totalTransacoes').textContent = totalTransacoes;
            document.getElementById('ticketMedio').textContent = `R$ ${ticketMedio.toFixed(2).replace('.', ',')}`;

            renderizarGraficoMensal(vendas);
        } else {
            exibirStatus({ status: 'error', mensagem: 'Nenhuma venda encontrada.' });
        }
    } catch (error) {
        document.getElementById('loading').classList.add('hidden');
        exibirStatus({ status: 'error', mensagem: 'Erro ao carregar dados do dashboard: ' + error.message });
    }
}

function renderizarGraficoMensal(vendas) {
    const dadosGrupados = {};

    vendas.forEach(v => {
        if (!v.Data) return;
        const [dia, mes, ano] = v.Data.split(' ')[0].split('/');
        if (!mes || !ano) return;

        const mesAno = `${mes}/${ano}`;
        if (!dadosGrupados[mesAno]) dadosGrupados[mesAno] = 0;

        let val = v['Total com Desconto'];
        if (typeof val === 'string') val = String(val).replace(/[R\$\s\.]/g, '').replace(',', '.');
        dadosGrupados[mesAno] += (parseFloat(val) || 0);
    });

    const meses = Object.keys(dadosGrupados).sort((a, b) => {
        const [mA, aA] = a.split('/');
        const [mB, aB] = b.split('/');
        return new Date(aA, mA - 1) - new Date(aB, mB - 1);
    });

    const totais = meses.map(m => dadosGrupados[m]);

    const graficoContainer = document.getElementById('graficoContainer');
    // Cria o canvas
    graficoContainer.innerHTML += '<canvas id="vendasChart" style="max-height: 300px; width:100%;"></canvas>';

    const ctx = document.getElementById('vendasChart').getContext('2d');
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: meses,
            datasets: [{
                label: 'Receita (R$)',
                data: totais,
                backgroundColor: 'rgba(59, 130, 246, 0.8)',
                borderColor: 'rgba(37, 99, 235, 1)',
                borderWidth: 1,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function (value) {
                            return 'R$ ' + value.toLocaleString('pt-BR');
                        }
                    }
                }
            }
        }
    });
}
