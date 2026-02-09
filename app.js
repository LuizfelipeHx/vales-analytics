// ===== Global State =====
let rawData = [];
let statusChart = null;

// ===== Constants =====
const EXCEL_URL = 'https://raw.githubusercontent.com/LuizfelipeHx/vales-analytics/main/dados.xlsx';
const HEADER_ROW = 7;
const DATA_START = 8;

// ===== DOM Elements =====
const loadingOverlay = document.getElementById('loadingOverlay');
const emptyState = document.getElementById('emptyState');
const dashboard = document.getElementById('dashboard');

// ===== Initialize =====
document.addEventListener('DOMContentLoaded', () => {
    setupEventListeners();
    loadDataFromGitHub();
});

function setupEventListeners() {
    // Refresh button
    document.getElementById('refreshBtn').addEventListener('click', loadDataFromGitHub);

    // Bottom nav
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => handleNavAction(btn.dataset.action));
    });

    // Tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });
}

// ===== Data Loading from GitHub =====
async function loadDataFromGitHub() {
    showLoading(true);

    try {
        console.log('Carregando dados de:', EXCEL_URL);

        const response = await fetch(EXCEL_URL + '?t=' + Date.now()); // Cache busting
        if (!response.ok) {
            throw new Error('Arquivo não encontrado. Verifique se dados.xlsx foi enviado para o GitHub.');
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Find the correct sheet
        const sheetName = findValidSheet(workbook);
        if (!sheetName) {
            throw new Error('Nenhuma aba válida encontrada');
        }

        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        rawData = parseSheetData(rows);
        console.log('Dados carregados:', rawData.length, 'registros');

        if (rawData.length === 0) {
            throw new Error('Nenhum dado encontrado na planilha');
        }

        showDashboard();
        updateAllData();
        updateLastUpdate();

    } catch (error) {
        console.error('Erro:', error);
        showError(error.message);
    } finally {
        showLoading(false);
    }
}

function findValidSheet(workbook) {
    const keywords = ['vales', 'dados', 'planilha', 'sheet'];

    for (const name of workbook.SheetNames) {
        const lower = name.toLowerCase();
        if (keywords.some(k => lower.includes(k))) {
            return name;
        }
    }

    // Return first sheet if no match
    return workbook.SheetNames[0];
}

function parseSheetData(rows) {
    const data = [];

    // Find column positions from header row
    let colMap = { data: 6, nome: 10, sala: 11, status: 14, valor: 19 };

    const headerRowData = rows[HEADER_ROW];
    if (headerRowData) {
        for (let j = 0; j < headerRowData.length; j++) {
            const cell = String(headerRowData[j] || '').toLowerCase().trim();

            if ((cell.includes('data') && (cell.includes('lan') || cell.includes('lcto'))) || cell === 'data') {
                colMap.data = j;
            }
            if (cell.includes('nome') || cell.includes('funcionário') || cell.includes('funcionario')) {
                colMap.nome = j;
            }
            if (cell.includes('sala') || cell === 'setor' || cell === 'unidade') {
                colMap.sala = j;
            }
            if (cell.includes('status') || cell.includes('situação') || cell.includes('situacao')) {
                colMap.status = j;
            }
            if (cell.includes('valor') || cell === 'vl' || cell === 'total') {
                colMap.valor = j;
            }
        }
    }

    // Parse data rows
    for (let i = DATA_START; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const nome = String(row[colMap.nome] || '').trim();
        const sala = String(row[colMap.sala] || '').trim();
        const status = String(row[colMap.status] || '').trim();
        const valor = parseFloat(row[colMap.valor]) || 0;

        if (!nome && !status) continue;

        data.push({
            nome: nome,
            sala: sala,
            status: normalizeStatus(status),
            valor: valor
        });
    }

    return data;
}

function normalizeStatus(status) {
    if (!status) return 'Não informado';
    return status.charAt(0).toUpperCase() + status.slice(1).toLowerCase();
}

// ===== UI Updates =====
function showLoading(show) {
    loadingOverlay.classList.toggle('active', show);
}

function showDashboard() {
    emptyState.style.display = 'none';
    dashboard.style.display = 'block';
}

function showError(message) {
    emptyState.style.display = 'block';
    dashboard.style.display = 'none';
    emptyState.querySelector('h2').textContent = 'Erro ao carregar';
    emptyState.querySelector('p').textContent = message;
}

function updateLastUpdate() {
    const now = new Date();
    const formatted = now.toLocaleString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
    document.getElementById('periodInfo').innerHTML = `<span>📅 ${rawData.length} registros • Atualizado: ${formatted}</span>`;
}

function handleNavAction(action) {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.action === action);
    });

    switch (action) {
        case 'refresh':
            loadDataFromGitHub();
            break;
    }
}

function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.tab === tabName);
    });

    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabName}`);
    });
}

// ===== Data Calculations =====
function updateAllData() {
    const totals = calculateTotals(rawData);
    updateKPIs(totals);
    updateOfensores();
    updateSalas();
    updateChart(totals);
}

function calculateTotals(data) {
    const result = {
        total: data.length,
        reprovado: { count: 0, value: 0 },
        abonado: { count: 0, value: 0 },
        analise: { count: 0, value: 0 },
        outros: { count: 0, value: 0 }
    };

    data.forEach(item => {
        const status = item.status.toLowerCase();

        if (status.includes('reprovad') || status.includes('cobrança')) {
            result.reprovado.count++;
            result.reprovado.value += item.valor;
        } else if (status.includes('abonad')) {
            result.abonado.count++;
            result.abonado.value += item.valor;
        } else if (status.includes('análise') || status.includes('analise')) {
            result.analise.count++;
            result.analise.value += item.valor;
        } else {
            result.outros.count++;
            result.outros.value += item.valor;
        }
    });

    return result;
}

function updateKPIs(totals) {
    document.getElementById('totalVales').textContent = totals.total;
    document.getElementById('totalReprovado').textContent = totals.reprovado.count;
    document.getElementById('valorReprovado').textContent = formatCurrency(totals.reprovado.value);
    document.getElementById('totalAbonado').textContent = totals.abonado.count;
    document.getElementById('valorAbonado').textContent = formatCurrency(totals.abonado.value);
    document.getElementById('totalAnalise').textContent = totals.analise.count;
    document.getElementById('valorAnalise').textContent = formatCurrency(totals.analise.value);

    // Secondary KPIs
    const valorMedio = totals.reprovado.count > 0 ? totals.reprovado.value / totals.reprovado.count : 0;
    document.getElementById('valorMedio').textContent = formatCurrency(valorMedio);

    const taxaReprovacao = totals.total > 0 ? (totals.reprovado.count / totals.total * 100) : 0;
    document.getElementById('taxaReprovacao').textContent = taxaReprovacao.toFixed(1) + '%';

    const valorTotal = rawData.reduce((sum, item) => sum + item.valor, 0);
    document.getElementById('valorTotal').textContent = formatCurrency(valorTotal);
}

function updateOfensores() {
    const reprovados = rawData.filter(item => {
        const status = item.status.toLowerCase();
        return status.includes('reprovad') || status.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(item => {
        if (!grouped[item.nome]) {
            grouped[item.nome] = { nome: item.nome, sala: item.sala, count: 0, valor: 0 };
        }
        grouped[item.nome].count++;
        grouped[item.nome].valor += item.valor;
    });

    const top10 = Object.values(grouped)
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

    renderOfensorList('ofensorList', top10);
}

function updateSalas() {
    const reprovados = rawData.filter(item => {
        const status = item.status.toLowerCase();
        return status.includes('reprovad') || status.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(item => {
        if (!grouped[item.sala]) {
            grouped[item.sala] = { sala: item.sala, count: 0, valor: 0 };
        }
        grouped[item.sala].count++;
        grouped[item.sala].valor += item.valor;
    });

    const top10 = Object.values(grouped)
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

    renderSalaList('salaList', top10);
}

function renderOfensorList(containerId, data) {
    const container = document.getElementById(containerId);

    if (data.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary)">Nenhum ofensor encontrado</p>';
        return;
    }

    container.innerHTML = data.map((item, i) => `
        <div class="ofensor-item">
            <div class="ofensor-rank ${i < 3 ? 'rank-' + (i + 1) : ''}">${i + 1}</div>
            <div class="ofensor-info">
                <div class="ofensor-name">${item.nome || 'N/A'}</div>
                <div class="ofensor-sala">${item.sala || 'Sala N/A'}</div>
            </div>
            <div class="ofensor-stats">
                <div class="ofensor-count">${item.count} vales</div>
                <div class="ofensor-value">${formatCurrency(item.valor)}</div>
            </div>
        </div>
    `).join('');
}

function renderSalaList(containerId, data) {
    const container = document.getElementById(containerId);

    if (data.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary)">Nenhuma sala encontrada</p>';
        return;
    }

    container.innerHTML = data.map((item, i) => `
        <div class="ofensor-item">
            <div class="ofensor-rank ${i < 3 ? 'rank-' + (i + 1) : ''}">${i + 1}</div>
            <div class="ofensor-info">
                <div class="ofensor-name">${item.sala || 'N/A'}</div>
            </div>
            <div class="ofensor-stats">
                <div class="ofensor-count">${item.count} vales</div>
                <div class="ofensor-value">${formatCurrency(item.valor)}</div>
            </div>
        </div>
    `).join('');
}

function updateChart(totals) {
    const ctx = document.getElementById('statusChart').getContext('2d');

    if (statusChart) {
        statusChart.destroy();
    }

    statusChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Reprovados', 'Abonados', 'Em Análise', 'Outros'],
            datasets: [{
                data: [
                    totals.reprovado.count,
                    totals.abonado.count,
                    totals.analise.count,
                    totals.outros.count
                ],
                backgroundColor: ['#ff7675', '#74b9ff', '#9b59b6', '#fdcb6e'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: '#a0a0b0',
                        padding: 16,
                        usePointStyle: true
                    }
                }
            }
        }
    });
}

// ===== Utilities =====
function formatCurrency(value) {
    return new Intl.NumberFormat('pt-BR', {
        style: 'currency',
        currency: 'BRL',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(value);
}

// ===== Service Worker Registration =====
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('sw.js')
            .then(reg => console.log('Service Worker registrado'))
            .catch(err => console.log('Erro no SW:', err));
    });
}
