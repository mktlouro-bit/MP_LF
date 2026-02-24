// O URL de exportação CSV do seu Google Sheet.
const GOOGLE_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1cM20IYXbhuPhH3Z_3S-6jYXilL6nCGwL/gviz/tq?tqx=out:csv&gid=654429644';

// Intervalo de atualização: 5 minutos em milissegundos
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;

// Variáveis para guardar as instâncias dos gráficos, para que possam ser atualizadas sem piscar
let statusChartInstance = null;
let fornecedorChartInstance = null;

async function fetchData() {
    try {
        const response = await fetch(GOOGLE_SHEET_CSV_URL);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const csvText = await response.text();

        const { data, errors } = Papa.parse(csvText, {
            header: true,
            skipEmptyLines: true,
            trimHeaders: true
        });

        if (errors.length > 0) {
            console.error('Erros ao analisar CSV:', errors);
            document.getElementById('last-updated').textContent = "Erro ao carregar dados! (Consulte a consola)";
            return null;
        }

        console.log('Dados brutos do CSV:', data);
        return processData(data);

    } catch (error) {
        console.error('Erro ao buscar ou processar dados:', error);
        document.getElementById('last-updated').textContent = "Erro de conexão ou dados!";
        return null;
    }
}

function processData(rawData) {
    let totalCasos = 0;
    let abertos = 0;
    let resolvidosOK = 0;
    let resolvidosNOK = 0;
    const fornecedorCounts = {};
    const processedCases = [];

    rawData.forEach(row => {
        // Ignorar linhas que não têm um número de comunicação válido ou são incompletas
        // Usando os nomes EXATOS das colunas do CSV
        if (!row['Nº'] || !row['Data comunicação'] || !row['Estado']) {
            console.warn('Linha ignorada devido a dados essenciais incompletos:', row); 
            return;
        }

        totalCasos++;

        let finalStatus = 'Desconhecido';

        const estadoDaReclamacao = row['Estado'].trim().toLowerCase();
        // Usando o nome da coluna "Ok/NOK" do CSV
        const okNoStatus = row['Ok/NOK'] ? row['Ok/NOK'].trim().toLowerCase() : '';

        if (estadoDaReclamacao === 'aberto') {
            abertos++;
            finalStatus = 'Aberto';
        } else {
            // Se o estado não é "Aberto", verificamos a coluna 'Ok/NOK'
            if (okNoStatus === 'ok') {
                resolvidosOK++;
                finalStatus = 'OK';
            } else if (okNoStatus === 'nok') {
                resolvidosNOK++;
                finalStatus = 'NOK';
            } else {
                resolvidosNOK++; // Assume NOK se não for Aberto nem OK explícito
                finalStatus = 'NOK';
                console.warn(`Caso ${row['Nº']} com estado '${estadoDaReclamacao}' mas 'Ok/NOK' é '${okNoStatus || 'vazio'}'. Classificado como NOK.`);
            }
        }
        
        // Contar fornecedores para o gráfico
        // Usando o nome da coluna "Reclamações a Fornecedores 2026 Fornecedor"
        const fornecedor = row['Reclamações a Fornecedores 2026 Fornecedor'] ? row['Reclamações a Fornecedores 2026 Fornecedor'].trim() : 'Desconhecido';
        if (fornecedor && fornecedor !== 'Desconhecido' && fornecedor !== '-') {
            fornecedorCounts[fornecedor] = (fornecedorCounts[fornecedor] || 0) + 1;
        }

        processedCases.push({
            'Nº': row['Nº'],
            'Data comunicação': row['Data comunicação'],
            'Fornecedor': fornecedor,
            'Motivo': row['Motivo'] || 'N/A',
            'Estado': finalStatus
        });
    });

    processedCases.sort((a, b) => {
        const parseDate = (dateStr) => {
            if (!dateStr) return new Date(0);
            const parts = dateStr.split('/');
            // Assumindo que o formato é MM/DD/YYYY ou D/M/YYYY (pelo seu CSV "1/7/2026", "1/26/2026")
            // `new Date(year, monthIndex, day)`
            if (parts.length === 3) {
                 // Convertendo para o formato MM/DD/YYYY para new Date()
                const month = parseInt(parts[0]);
                const day = parseInt(parts[1]);
                const year = parseInt(parts[2]);
                return new Date(`${month}/${day}/${year}`);
            }
            return new Date(dateStr); 
        };
        try {
            const dateA = parseDate(a['Data comunicação']);
            const dateB = parseDate(b['Data comunicação']);
            return dateB.getTime() - dateA.getTime();
        } catch (e) {
            console.error("Erro ao analisar data para ordenação:", e, "Data A:", a['Data comunicação'], "Data B:", b['Data comunicação']);
            return 0;
        }
    });

    return {
        totalCasos,
        abertos,
        resolvidosOK,
        resolvidosNOK,
        fornecedorCounts,
        recentCases: processedCases.filter(c => c.Estado === 'Aberto').slice(0, 10)
    };
}

function updateDashboard(data) {
    if (!data) {
        console.warn('Nenhum dado para atualizar o dashboard.');
        return;
    }

    // Atualizar KPIs
    document.getElementById('kpi-total-casos').textContent = data.totalCasos;
    document.getElementById('kpi-abertos').textContent = data.abertos;
    document.getElementById('kpi-ok').textContent = data.resolvidosOK;
    document.getElementById('kpi-nok').textContent = data.resolvidosNOK;

    // Atualizar tempo da última atualização
    document.getElementById('last-updated').textContent = new Date().toLocaleTimeString('pt-PT');

    // Gráfico de Status (Rosca) - Atualização Suave
    const statusCtx = document.getElementById('statusChart').getContext('2d');
    const statusLabels = [`Abertos (${data.abertos})`, `Resolvidos OK (${data.resolvidosOK})`, `Resolvidos NOK (${data.resolvidosNOK})`];
    const statusChartData = [data.abertos, data.resolvidosOK, data.resolvidosNOK];

    if (statusChartInstance) {
        statusChartInstance.data.labels = statusLabels;
        statusChartInstance.data.datasets[0].data = statusChartData;
        statusChartInstance.update(); // Atualiza sem piscar
    } else {
        statusChartInstance = new Chart(statusCtx, {
            type: 'doughnut',
            data: {
                labels: statusLabels,
                datasets: [{
                    data: statusChartData,
                    backgroundColor: ['var(--kpi-aberto)', 'var(--kpi-ok)', 'var(--kpi-nok)'],
                    hoverOffset: 4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            font: { size: 14 }
                        }
                    },
                    title: { display: false }
                }
            }
        });
    }

    // Gráfico de Fornecedores (Top 5 - Barras Horizontais) - Atualização Suave
    const sortedFornecedores = Object.entries(data.fornecedorCounts)
                                    .sort(([,a], [,b]) => b - a)
                                    .slice(0, 5);
    const fornecedorLabels = sortedFornecedores.map(([forn, count]) => `${forn} (${count})`);
    const fornecedorData = sortedFornecedores.map(([,count]) => count);

    const fornecedorCtx = document.getElementById('fornecedorChart').getContext('2d');
    if (fornecedorChartInstance) {
        fornecedorChartInstance.data.labels = fornecedorLabels;
        fornecedorChartInstance.data.datasets[0].data = fornecedorData;
        fornecedorChartInstance.update(); // Atualiza sem piscar
    } else {
        fornecedorChartInstance = new Chart(fornecedorCtx, {
            type: 'bar',
            data: {
                labels: fornecedorLabels,
                datasets: [{
                    label: 'Número de Reclamações',
                    data: fornecedorData,
                    backgroundColor: 'rgba(54, 162, 235, 0.6)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    title: { display: false }
                },
                indexAxis: 'y',
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0
                        }
                    }
                }
            }
        });
    }

    // Tabela de Casos Abertos Recentes
    const tableBody = document.getElementById('recentCasesTable').getElementsByTagName('tbody')[0];
    tableBody.innerHTML = '';
    data.recentCases.forEach(caso => {
        const row = tableBody.insertRow();
        row.insertCell().textContent = caso['Data comunicação'];
        row.insertCell().textContent = caso['Fornecedor'];
        row.insertCell().textContent = caso['Motivo'].length > 50 ? caso['Motivo'].substring(0, 50) + '...' : caso['Motivo'];
        const estadoCell = row.insertCell();
        estadoCell.textContent = caso['Estado'];
        estadoCell.className = `table-status-${caso['Estado'].toLowerCase().replace(/\s/g, '-')}`; 
    });
}

async function initDashboard() {
    console.log('Inicializando dashboard...');
    const data = await fetchData();
    if (data) {
        updateDashboard(data);
    }
    setInterval(async () => {
        console.log('A atualizar dados...');
        const updatedData = await fetchData();
        if (updatedData) {
            updateDashboard(updatedData);
        }
    }, REFRESH_INTERVAL_MS);
}

document.addEventListener('DOMContentLoaded', initDashboard);
