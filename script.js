// O URL de exportação CSV do seu Google Sheet.
// Ele é construído a partir do ID do seu sheet e do GID da folha (654429644 para "2026").
const GOOGLE_SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1cM20IYXbhuPhH3Z_3S-6jYXilL6nCGwL/gviz/tq?tqx=out:csv&gid=654429644';

// Intervalo de atualização: 5 minutos em milissegundos
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;

let statusChartInstance = null;
let fornecedorChartInstance = null;

async function fetchData() {
    try {
        const response = await fetch(GOOGLE_SHEET_CSV_URL);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const csvText = await response.text();

        // Usar Papaparse para ler o CSV
        const { data, errors } = Papa.parse(csvText, {
            header: true,         // A primeira linha é o cabeçalho
            skipEmptyLines: true, // Ignorar linhas vazias
            trimHeaders: true     // Remover espaços em branco dos nomes dos cabeçalhos
            // REMOVIDO: delimiter: ';' -- Deixa o PapaParse auto-detetar (ele usará vírgula por padrão)
        });

        if (errors.length > 0) {
            console.error('Erros ao analisar CSV:', errors);
            // Mostrar aviso no dashboard se houver erros de parsing
            document.getElementById('last-updated').textContent = "Erro ao carregar dados! (Consulte a consola)";
            return null; // Retorna nulo se houver erros críticos
        }

        console.log('Dados brutos do CSV:', data); // Para depuração
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
        // console.log("Linha processada:", row); // Temporário para depuração de nomes de colunas

        // Ignorar linhas que não têm um número de comunicação válido ou são incompletas
        // ATUALIZADO: Usando os nomes EXATOS das colunas do CSV
        if (!row['Nº'] || !row['Data comunicação'] || !row['Estado']) {
            console.warn('Linha ignorada devido a dados essenciais incompletos:', row); 
            return;
        }

        totalCasos++;

        let finalStatus = 'Desconhecido'; // Estado padrão se não for classificado

        const estadoDaReclamacao = row['Estado'].trim().toLowerCase();
        // ATUALIZADO: Usando o nome da coluna "Ok/NOK" do CSV
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
                // Caso o 'Estado' não seja "Aberto" mas 'Ok/NOK' não seja "OK" nem "NOK"
                resolvidosNOK++;
                finalStatus = 'NOK';
                console.warn(`Caso ${row['Nº']} com estado '${estadoDaReclamacao}' mas 'Ok/NOK' é '${okNoStatus || 'vazio'}'. Classificado como NOK.`);
            }
        }
        
        // Contar fornecedores para o gráfico
        // ATUALIZADO: Usando o nome da coluna "Reclamações a Fornecedores 2026 Fornecedor"
        const fornecedor = row['Reclamações a Fornecedores 2026 Fornecedor'] ? row['Reclamações a Fornecedores 2026 Fornecedor'].trim() : 'Desconhecido';
        if (fornecedor && fornecedor !== 'Desconhecido' && fornecedor !== '-') { // Só conta se o fornecedor for válido e não for '-'
            fornecedorCounts[fornecedor] = (fornecedorCounts[fornecedor] || 0) + 1;
        }

        processedCases.push({
            'Nº': row['Nº'],
            'Data comunicação': row['Data comunicação'],
            'Fornecedor': fornecedor, // Aqui pode usar o nome da coluna do Google Sheet se for mais fácil, ou o nome limpo 'fornecedor'
            'Motivo': row['Motivo'] || 'N/A', // Se o motivo estiver vazio, usa 'N/A'
            'Estado': finalStatus // Usar o status processado
        });
    });

    // Ordenar casos recentes por data de comunicação (do mais recente para o mais antigo)
    processedCases.sort((a, b) => {
        const parseDate = (dateStr) => {
            if (!dateStr) return new Date(0); // Para lidar com datas vazias
            const parts = dateStr.split('/'); // Espera formato DD/MM/AAAA (seu CSV usa 1/7/2026 -> 1/MM/AAAA)
            // Note: Seu CSV está a usar MM/DD/AAAA ou D/M/AAAA para "1/7/2026" ou "1/26/2026". 
            // O JavaScript `new Date()` para "MM/DD/AAAA" funciona melhor. Se o seu Google Sheet 
            // estiver definido para "DD/MM/AAAA", isto pode causar problemas. Vamos ajustar para tentar
            // parsear de forma mais robusta.
            
            // Tentativa de parse DD/MM/AAAA e MM/DD/AAAA
            const dateParts = dateStr.split('/');
            if (dateParts.length === 3) {
                // Assumindo formato DD/MM/AAAA como é comum em PT
                const day = parseInt(dateParts[0]);
                const month = parseInt(dateParts[1]);
                const year = parseInt(dateParts[2]);
                // Se o mês for maior que 12, é provável que seja DD/MM/AAAA e não MM/DD/AAAA
                if (month > 12 && day <= 12) { // Tenta adivinhar se é MM/DD/AAAA
                    return new Date(year, day - 1, month);
                }
                return new Date(year, month - 1, day);
            }
            return new Date(dateStr); // Tenta parsear como string genérica
        };
        try {
            const dateA = parseDate(a['Data comunicação']);
            const dateB = parseDate(b['Data comunicação']);
            return dateB.getTime() - dateA.getTime(); // Ordem decrescente (mais recente primeiro)
        } catch (e) {
            console.error("Erro ao analisar data para ordenação:", e, "Data A:", a['Data comunicação'], "Data B:", b['Data comunicação']);
            return 0; // Se as datas forem inválidas, não ordena
        }
    });

    return {
        totalCasos,
        abertos,
        resolvidosOK,
        resolvidosNOK,
        fornecedorCounts,
        // Filtra para pegar apenas casos "Aberto" e pega os 10 mais recentes
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

    // Gráfico de Status (Rosca)
    const statusCtx = document.getElementById('statusChart').getContext('2d');
    if (statusChartInstance) {
        statusChartInstance.destroy(); // Destrói a instância anterior para atualizar o gráfico
    }
    statusChartInstance = new Chart(statusCtx, {
        type: 'doughnut',
        data: {
            labels: [`Abertos (${data.abertos})`, `Resolvidos OK (${data.resolvidosOK})`, `Resolvidos NOK (${data.resolvidosNOK})`],
            datasets: [{
                data: [data.abertos, data.resolvidosOK, data.resolvidosNOK],
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
                title: {
                    display: false // Título já no <h2>
                }
            }
        }
    });

    // Gráfico de Fornecedores (Top 5 - Barras Horizontais)
    // Converte o objeto de contagens para um array de pares [fornecedor, contagem]
    const sortedFornecedores = Object.entries(data.fornecedorCounts)
                                    .sort(([,a], [,b]) => b - a) // Ordena por contagem descendente
                                    .slice(0, 5); // Pega os Top 5

    const fornecedorLabels = sortedFornecedores.map(([forn, count]) => `${forn} (${count})`);
    const fornecedorData = sortedFornecedores.map(([,count]) => count);

    const fornecedorCtx = document.getElementById('fornecedorChart').getContext('2d');
    if (fornecedorChartInstance) {
        fornecedorChartInstance.destroy();
    }
    fornecedorChartInstance = new Chart(fornecedorCtx, {
        type: 'bar',
        data: {
            labels: fornecedorLabels,
            datasets: [{
                label: 'Número de Reclamações',
                data: fornecedorData,
                backgroundColor: 'rgba(54, 162, 235, 0.6)', // Cor azul
                borderColor: 'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }, // Não mostra legenda para barras únicas
                title: { display: false } // Título já no <h2>
            },
            indexAxis: 'y', // Para ter barras horizontais
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0 // Assegura que os rótulos do eixo X sejam inteiros
                    }
                }
            }
        }
    });

    // Tabela de Casos Abertos Recentes
    const tableBody = document.getElementById('recentCasesTable').getElementsByTagName('tbody')[0];
    tableBody.innerHTML = ''; // Limpa a tabela antes de preencher
    data.recentCases.forEach(caso => {
        const row = tableBody.insertRow();
        row.insertCell().textContent = caso['Data comunicação'];
        row.insertCell().textContent = caso['Fornecedor'];
        // Trunca o motivo se for muito longo para não quebrar o layout
        row.insertCell().textContent = caso['Motivo'].length > 50 ? caso['Motivo'].substring(0, 50) + '...' : caso['Motivo'];
        const estadoCell = row.insertCell();
        estadoCell.textContent = caso['Estado'];
        // Aplica a classe CSS para colorir o texto do estado na tabela
        estadoCell.className = `table-status-${caso['Estado'].toLowerCase().replace(/\s/g, '-')}`; 
    });
}

// Função para iniciar o dashboard e configurar o refresh automático
async function initDashboard() {
    console.log('Inicializando dashboard...');
    const data = await fetchData(); // Busca os dados iniciais
    if (data) {
        updateDashboard(data); // Atualiza a UI com os dados
    }
    // Configura o agendador para buscar e atualizar dados a cada REFRESH_INTERVAL_MS
    setInterval(async () => {
        console.log('A atualizar dados...');
        const updatedData = await fetchData();
        if (updatedData) {
            updateDashboard(updatedData);
        }
    }, REFRESH_INTERVAL_MS);
}

// Inicia o dashboard quando o DOM estiver completamente carregado
document.addEventListener('DOMContentLoaded', initDashboard);
