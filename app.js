// ===== Global State =====
let rawData = [];
let filteredData = [];
let statusChart = null;
let evolutionChart = null;
let salaChart = null;
let ofensorSortBy = 'quantidade';
let salaSortBy = 'quantidade';

// ===== Constants =====
const EXCEL_URL = 'https://raw.githubusercontent.com/LuizfelipeHx/vales-analytics/main/dados.xlsx';
const SHEET_NAME = 'Acomp Físico'; // Nome exato da aba
const HEADER_ROW = 7;  // Linha 8 no Excel (0-indexed)
const DATA_START = 8;  // Linha 9 no Excel (0-indexed)
// Colunas confirmadas: G=Data, K=Nome, L=Sala, O=Status, T=Valor
const COL = { data: 6, nome: 10, sala: 11, status: 14, valor: 19 };

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

    // Filters toggle
    document.getElementById('filtersToggle').addEventListener('click', toggleFilters);

    // Filter changes
    document.getElementById('filterPeriodo').addEventListener('change', applyFilters);
    document.getElementById('filterSala').addEventListener('change', applyFilters);
    document.getElementById('filterStatus').addEventListener('change', applyFilters);
    document.getElementById('clearFilters').addEventListener('click', clearFilters);

    // Bottom nav
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => handleNavAction(btn.dataset.action));
    });

    // Tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });

    // Ranking toggles
    document.querySelectorAll('#tab-ofensores .ranking-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#tab-ofensores .ranking-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            ofensorSortBy = btn.dataset.sort;
            updateOfensores();
        });
    });

    document.querySelectorAll('#tab-salas .ranking-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('#tab-salas .ranking-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            salaSortBy = btn.dataset.sort;
            updateSalas();
        });
    });
}

// ===== Filters =====
function toggleFilters() {
    document.getElementById('filtersToggle').classList.toggle('active');
    document.getElementById('filtersContent').classList.toggle('active');
}

function isValidSala(sala) {
    if (!sala || sala === '-' || sala === '') return false;
    const lower = sala.toLowerCase().trim();
    // Only exclude exact summary rows
    if (lower === 'total' || lower === 'total de vale' || lower === 'total de vales' || lower === 'soma' || lower === 'subtotal') return false;
    return true;
}

function isValidNome(nome) {
    if (!nome || nome === '-' || nome === '') return false;
    const lower = nome.toLowerCase().trim();
    // Only exclude exact summary rows
    if (lower === 'total' || lower === 'total de vale' || lower === 'total de vales' || lower === 'soma' || lower === 'subtotal') return false;
    return true;
}

function populateFilters() {
    const periodos = new Set();
    const salas = new Set();
    const statuses = new Set();

    rawData.forEach(item => {
        if (item.periodo) periodos.add(item.periodo);
        if (item.sala && isValidSala(item.sala)) salas.add(item.sala);
        if (item.status && item.status !== 'Não informado') statuses.add(item.status);
    });

    // Populate period filter
    const periodoSelect = document.getElementById('filterPeriodo');
    periodoSelect.innerHTML = '<option value="">Todos</option>';
    [...periodos].sort().reverse().forEach(p => {
        periodoSelect.innerHTML += `<option value="${p}">${p}</option>`;
    });

    // Populate sala filter - sorted alphabetically
    const salaSelect = document.getElementById('filterSala');
    salaSelect.innerHTML = '<option value="">Todas</option>';
    [...salas].sort().forEach(s => {
        salaSelect.innerHTML += `<option value="${s}">${s}</option>`;
    });

    // Populate status filter
    const statusSelect = document.getElementById('filterStatus');
    statusSelect.innerHTML = '<option value="">Todos</option>';
    [...statuses].sort().forEach(s => {
        statusSelect.innerHTML += `<option value="${s}">${s}</option>`;
    });
}

function applyFilters() {
    const periodo = document.getElementById('filterPeriodo').value;
    const sala = document.getElementById('filterSala').value;
    const status = document.getElementById('filterStatus').value;

    filteredData = rawData.filter(item => {
        if (periodo && item.periodo !== periodo) return false;
        if (sala && item.sala !== sala) return false;
        if (status && item.status !== status) return false;
        return true;
    });

    updateAllData();
}

function clearFilters() {
    document.getElementById('filterPeriodo').value = '';
    document.getElementById('filterSala').value = '';
    document.getElementById('filterStatus').value = '';
    filteredData = [...rawData];
    updateAllData();
}

// ===== Data Loading from GitHub =====
async function loadDataFromGitHub() {
    showLoading(true);

    try {
        console.log('Carregando dados de:', EXCEL_URL);

        const response = await fetch(EXCEL_URL + '?t=' + Date.now());
        if (!response.ok) {
            throw new Error('Arquivo não encontrado. Verifique se dados.xlsx foi enviado.');
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        const sheetName = findValidSheet(workbook);
        if (!sheetName) {
            throw new Error('Nenhuma aba válida encontrada');
        }

        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        rawData = parseSheetData(rows);
        filteredData = [...rawData];
        console.log('Dados carregados:', rawData.length, 'registros');

        if (rawData.length === 0) {
            throw new Error('Nenhum dado encontrado na planilha');
        }

        populateFilters();
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
    console.log('Abas disponíveis:', workbook.SheetNames);

    // Primeiro tenta o nome exato
    if (workbook.SheetNames.includes(SHEET_NAME)) {
        console.log('Aba encontrada (exata):', SHEET_NAME);
        return SHEET_NAME;
    }

    // Busca parcial sem acento
    const normalize = (s) => s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const target = normalize(SHEET_NAME);

    for (const name of workbook.SheetNames) {
        if (normalize(name).includes('acomp')) {
            console.log('Aba encontrada (parcial):', name);
            return name;
        }
    }

    // Tenta keywords gerais
    const keywords = ['fisico', 'vales', 'dados', 'planilha'];
    for (const name of workbook.SheetNames) {
        const lower = normalize(name);
        if (keywords.some(k => lower.includes(k))) {
            console.log('Aba encontrada (keyword):', name);
            return name;
        }
    }

    console.warn('Nenhuma aba reconhecida, usando primeira:', workbook.SheetNames[0]);
    return workbook.SheetNames[0];
}

function parseSheetData(rows) {
    const data = [];

    console.log('Total de linhas brutas:', rows.length);
    console.log('Usando colunas: Data=G(' + COL.data + '), Nome=K(' + COL.nome + '), Sala=L(' + COL.sala + '), Status=O(' + COL.status + '), Valor=T(' + COL.valor + ')');

    // Log header row for debugging
    const headerRowData = rows[HEADER_ROW];
    if (headerRowData) {
        console.log('Cabeçalho (linha 8):', headerRowData.slice(0, 22));
    }

    // Log first data row for debugging
    if (rows[DATA_START]) {
        const firstRow = rows[DATA_START];
        console.log('Primeira linha de dados (linha 9):');
        console.log('  G(data):', firstRow[COL.data]);
        console.log('  K(nome):', firstRow[COL.nome]);
        console.log('  L(sala):', firstRow[COL.sala]);
        console.log('  O(status):', firstRow[COL.status]);
        console.log('  T(valor):', firstRow[COL.valor]);
    }

    for (let i = DATA_START; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const nome = String(row[COL.nome] || '').trim();
        const sala = String(row[COL.sala] || '').trim();
        const status = String(row[COL.status] || '').trim();
        const valor = parseFloat(row[COL.valor]) || 0;

        // Skip empty or summary rows
        if (!nome && !status) continue;
        if (!isValidNome(nome)) continue;

        // Parse date for period
        let periodo = '';
        const dataCell = row[COL.data];
        if (dataCell) {
            const dateObj = parseExcelDate(dataCell);
            if (dateObj) {
                periodo = dateObj.toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' });
            }
        }

        data.push({
            nome: nome,
            sala: isValidSala(sala) ? sala : 'N/A',
            status: normalizeStatus(status),
            valor: valor,
            periodo: periodo
        });
    }

    return data;
}

function parseExcelDate(value) {
    if (typeof value === 'number') {
        // Excel serial date
        const date = new Date((value - 25569) * 86400 * 1000);
        return date;
    } else if (typeof value === 'string') {
        const parts = value.split('/');
        if (parts.length === 3) {
            return new Date(parts[2], parts[1] - 1, parts[0]);
        }
    }
    return null;
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
    document.getElementById('periodInfo').innerHTML = `<span>📅 ${filteredData.length} registros • ${formatted}</span>`;
}

function handleNavAction(action) {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.action === action);
    });
    if (action === 'refresh') {
        loadDataFromGitHub();
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
    const totals = calculateTotals(filteredData);
    updateKPIs(totals);
    updateOfensores();
    updateSalas();
    updateCharts(totals);
    updateLastUpdate();
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

    const valorMedio = totals.reprovado.count > 0 ? totals.reprovado.value / totals.reprovado.count : 0;
    document.getElementById('valorMedio').textContent = formatCurrency(valorMedio);

    const taxaReprovacao = totals.total > 0 ? (totals.reprovado.count / totals.total * 100) : 0;
    document.getElementById('taxaReprovacao').textContent = taxaReprovacao.toFixed(1) + '%';

    const valorTotal = filteredData.reduce((sum, item) => sum + item.valor, 0);
    document.getElementById('valorTotal').textContent = formatCurrency(valorTotal);
}

function updateOfensores() {
    const reprovados = filteredData.filter(item => {
        const status = item.status.toLowerCase();
        return status.includes('reprovad') || status.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(item => {
        if (!isValidNome(item.nome)) return;
        if (!grouped[item.nome]) {
            grouped[item.nome] = { nome: item.nome, sala: item.sala, count: 0, valor: 0 };
        }
        grouped[item.nome].count++;
        grouped[item.nome].valor += item.valor;
    });

    const sortKey = ofensorSortBy === 'valor' ? 'valor' : 'count';
    const top10 = Object.values(grouped)
        .sort((a, b) => b[sortKey] - a[sortKey])
        .slice(0, 10);

    renderOfensorList('ofensorList', top10);
}

function updateSalas() {
    const reprovados = filteredData.filter(item => {
        const status = item.status.toLowerCase();
        return status.includes('reprovad') || status.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(item => {
        if (!isValidSala(item.sala)) return;
        if (!grouped[item.sala]) {
            grouped[item.sala] = { sala: item.sala, count: 0, valor: 0 };
        }
        grouped[item.sala].count++;
        grouped[item.sala].valor += item.valor;
    });

    const sortKey = salaSortBy === 'valor' ? 'valor' : 'count';
    const top10 = Object.values(grouped)
        .sort((a, b) => b[sortKey] - a[sortKey])
        .slice(0, 10);

    renderSalaList('salaList', top10);
}

function renderOfensorList(containerId, data) {
    const container = document.getElementById(containerId);
    if (data.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhum ofensor encontrado</p>';
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
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhuma sala encontrada</p>';
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

// ===== Charts =====
function updateCharts(totals) {
    updateStatusChart(totals);
    updateEvolutionChart();
    updateSalaChart();
}

function updateStatusChart(totals) {
    const ctx = document.getElementById('statusChart').getContext('2d');
    if (statusChart) statusChart.destroy();

    statusChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Reprovados', 'Abonados', 'Em Análise', 'Outros'],
            datasets: [{
                data: [totals.reprovado.count, totals.abonado.count, totals.analise.count, totals.outros.count],
                backgroundColor: ['#ff7675', '#00cec9', '#9b59b6', '#fdcb6e'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#a0a0b0', padding: 12, usePointStyle: true, font: { size: 11 } } }
            }
        }
    });
}

function updateEvolutionChart() {
    const ctx = document.getElementById('evolutionChart').getContext('2d');
    if (evolutionChart) evolutionChart.destroy();

    // Group by period
    const periodos = {};
    filteredData.forEach(item => {
        if (!item.periodo) return;
        if (!periodos[item.periodo]) {
            periodos[item.periodo] = { reprovado: 0, abonado: 0, analise: 0 };
        }
        const status = item.status.toLowerCase();
        if (status.includes('reprovad') || status.includes('cobrança')) {
            periodos[item.periodo].reprovado++;
        } else if (status.includes('abonad')) {
            periodos[item.periodo].abonado++;
        } else if (status.includes('análise') || status.includes('analise')) {
            periodos[item.periodo].analise++;
        }
    });

    const labels = Object.keys(periodos).sort();
    const reprovadoData = labels.map(p => periodos[p].reprovado);
    const abonadoData = labels.map(p => periodos[p].abonado);
    const analiseData = labels.map(p => periodos[p].analise);

    evolutionChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                { label: 'Reprovados', data: reprovadoData, backgroundColor: '#ff7675' },
                { label: 'Abonados', data: abonadoData, backgroundColor: '#00cec9' },
                { label: 'Análise', data: analiseData, backgroundColor: '#9b59b6' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true, ticks: { color: '#a0a0b0', font: { size: 10 } }, grid: { display: false } },
                y: { stacked: true, ticks: { color: '#a0a0b0' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: {
                legend: { position: 'bottom', labels: { color: '#a0a0b0', padding: 8, usePointStyle: true, font: { size: 10 } } }
            }
        }
    });
}

function updateSalaChart() {
    const ctx = document.getElementById('salaChart').getContext('2d');
    if (salaChart) salaChart.destroy();

    // Get top 5 salas by reprovados
    const reprovados = filteredData.filter(item => {
        const status = item.status.toLowerCase();
        return status.includes('reprovad') || status.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(item => {
        if (!isValidSala(item.sala)) return;
        if (!grouped[item.sala]) grouped[item.sala] = 0;
        grouped[item.sala]++;
    });

    const sorted = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 5);
    const labels = sorted.map(s => s[0] || 'N/A');
    const data = sorted.map(s => s[1]);

    salaChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Reprovados',
                data: data,
                backgroundColor: ['#ff7675', '#fd9644', '#fdcb6e', '#a29bfe', '#74b9ff']
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#a0a0b0' }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#a0a0b0', font: { size: 11 } }, grid: { display: false } }
            },
            plugins: {
                legend: { display: false }
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
