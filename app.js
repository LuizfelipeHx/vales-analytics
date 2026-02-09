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
const HEADER_ROW = 7;  // Linha 8 no Excel (0-indexed)
const DATA_START = 8;  // Linha 9 no Excel (0-indexed)

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
    document.getElementById('refreshBtn').addEventListener('click', loadDataFromGitHub);
    document.getElementById('filtersToggle').addEventListener('click', toggleFilters);
    document.getElementById('filterPeriodo').addEventListener('change', applyFilters);
    document.getElementById('filterSala').addEventListener('change', applyFilters);
    document.getElementById('filterStatus').addEventListener('change', applyFilters);
    document.getElementById('clearFilters').addEventListener('click', clearFilters);

    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', () => handleNavAction(btn.dataset.action));
    });

    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });

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

function isSummaryRow(text) {
    if (!text) return false;
    const lower = text.toLowerCase().trim();
    return lower === 'total' || lower === 'total de vale' || lower === 'total de vales' ||
        lower === 'soma' || lower === 'subtotal' || lower === 'grand total';
}

function populateFilters() {
    const periodos = new Set();
    const salas = new Set();
    const statuses = new Set();

    rawData.forEach(item => {
        if (item.periodo) periodos.add(item.periodo);
        if (item.sala && item.sala !== '-' && item.sala !== '') salas.add(item.sala);
        if (item.status && item.status !== 'Não informado') statuses.add(item.status);
    });

    const periodoSelect = document.getElementById('filterPeriodo');
    periodoSelect.innerHTML = '<option value="">Todos</option>';
    [...periodos].sort().reverse().forEach(p => {
        periodoSelect.innerHTML += `<option value="${p}">${p}</option>`;
    });

    const salaSelect = document.getElementById('filterSala');
    salaSelect.innerHTML = '<option value="">Todas</option>';
    [...salas].sort().forEach(s => {
        salaSelect.innerHTML += `<option value="${s}">${s}</option>`;
    });

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
        const response = await fetch(EXCEL_URL + '?t=' + Date.now());
        if (!response.ok) {
            throw new Error('Arquivo não encontrado no GitHub.');
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        console.log('Abas disponíveis:', workbook.SheetNames);

        // Find the correct sheet
        const sheetName = findValidSheet(workbook);
        console.log('Aba selecionada:', sheetName);

        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        console.log('Total linhas (com vazias):', rows.length);

        // Find columns dynamically
        const colMap = detectColumns(rows);
        console.log('Colunas detectadas:', colMap);

        rawData = parseSheetData(rows, colMap);
        filteredData = [...rawData];
        console.log('Registros válidos:', rawData.length);

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
    const normalize = (s) => s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    // Try exact names
    const names = ['Acomp Físico', 'Acomp Fisico'];
    for (const n of names) {
        if (workbook.SheetNames.includes(n)) return n;
    }

    // Try partial match
    for (const name of workbook.SheetNames) {
        const lower = normalize(name);
        if (lower.includes('acomp') || lower.includes('fisico') || lower.includes('vales') || lower.includes('dados')) {
            return name;
        }
    }

    return workbook.SheetNames[0];
}

function detectColumns(rows) {
    // Default positions: G=6, K=10, L=11, O=14, T=19
    const colMap = { data: 6, nome: 10, sala: 11, status: 14, valor: 19 };

    // Try to find header row (search rows 0-10)
    for (let r = 0; r <= Math.min(10, rows.length - 1); r++) {
        const row = rows[r];
        if (!row) continue;

        let foundHeaders = 0;
        for (let j = 0; j < row.length; j++) {
            const cell = String(row[j] || '').toLowerCase().trim();

            if (cell === 'data' || (cell.includes('data') && (cell.includes('lan') || cell.includes('lcto')))) {
                colMap.data = j; foundHeaders++;
            }
            if (cell.includes('nome') || cell.includes('funcionário') || cell.includes('funcionario')) {
                colMap.nome = j; foundHeaders++;
            }
            if (cell.includes('sala') || cell === 'setor' || cell === 'unidade') {
                colMap.sala = j; foundHeaders++;
            }
            if (cell.includes('status') || cell.includes('situação') || cell.includes('situacao')) {
                colMap.status = j; foundHeaders++;
            }
            if (cell.includes('valor') || cell === 'vl') {
                colMap.valor = j; foundHeaders++;
            }
        }

        if (foundHeaders >= 3) {
            colMap.headerRow = r;
            colMap.dataStart = r + 1;
            console.log('Cabeçalho encontrado na linha', r + 1, ':', row.slice(0, 22));
            return colMap;
        }
    }

    // If no header found, use defaults
    console.log('Cabeçalho padrão (linha 8) usado');
    if (rows[HEADER_ROW]) {
        console.log('Conteúdo linha 8:', rows[HEADER_ROW].slice(0, 22));
    }
    colMap.headerRow = HEADER_ROW;
    colMap.dataStart = DATA_START;
    return colMap;
}

function parseSheetData(rows, colMap) {
    const data = [];
    const startRow = colMap.dataStart || DATA_START;

    // Log first data rows for debugging
    for (let i = startRow; i < Math.min(startRow + 3, rows.length); i++) {
        const row = rows[i];
        if (row) {
            console.log(`Linha ${i + 1}: nome="${row[colMap.nome]}", sala="${row[colMap.sala]}", status="${row[colMap.status]}", valor="${row[colMap.valor]}"`);
        }
    }

    for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const nome = String(row[colMap.nome] || '').trim();
        const sala = String(row[colMap.sala] || '').trim();
        const status = String(row[colMap.status] || '').trim();
        const valor = parseFloat(row[colMap.valor]) || 0;

        // Skip completely empty rows
        if (!nome && !status) continue;
        // Skip summary rows
        if (isSummaryRow(nome)) continue;

        // Parse date for period
        let periodo = '';
        const dataCell = row[colMap.data];
        if (dataCell) {
            const dateObj = parseExcelDate(dataCell);
            if (dateObj) {
                periodo = dateObj.toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' });
            }
        }

        data.push({
            nome: nome || 'N/A',
            sala: isSummaryRow(sala) ? 'N/A' : (sala || 'N/A'),
            status: normalizeStatus(status),
            valor: valor,
            periodo: periodo
        });
    }

    return data;
}

function parseExcelDate(value) {
    if (typeof value === 'number') {
        return new Date((value - 25569) * 86400 * 1000);
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

// ===== UI =====
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
    const formatted = now.toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
    document.getElementById('periodInfo').innerHTML = `<span>📅 ${filteredData.length} registros • ${formatted}</span>`;
}

function handleNavAction(action) {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.action === action);
    });
    if (action === 'refresh') loadDataFromGitHub();
}

function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.tab === tabName);
    });
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabName}`);
    });
}

// ===== Calculations =====
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
        const s = item.status.toLowerCase();
        if (s.includes('reprovad') || s.includes('cobrança')) {
            result.reprovado.count++; result.reprovado.value += item.valor;
        } else if (s.includes('abonad')) {
            result.abonado.count++; result.abonado.value += item.valor;
        } else if (s.includes('análise') || s.includes('analise')) {
            result.analise.count++; result.analise.value += item.valor;
        } else {
            result.outros.count++; result.outros.value += item.valor;
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

    const taxa = totals.total > 0 ? (totals.reprovado.count / totals.total * 100) : 0;
    document.getElementById('taxaReprovacao').textContent = taxa.toFixed(1) + '%';

    const valorTotal = filteredData.reduce((sum, i) => sum + i.valor, 0);
    document.getElementById('valorTotal').textContent = formatCurrency(valorTotal);
}

function updateOfensores() {
    const reprovados = filteredData.filter(i => {
        const s = i.status.toLowerCase();
        return s.includes('reprovad') || s.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(i => {
        if (isSummaryRow(i.nome)) return;
        if (!grouped[i.nome]) grouped[i.nome] = { nome: i.nome, sala: i.sala, count: 0, valor: 0 };
        grouped[i.nome].count++;
        grouped[i.nome].valor += i.valor;
    });

    const sortKey = ofensorSortBy === 'valor' ? 'valor' : 'count';
    const top10 = Object.values(grouped).sort((a, b) => b[sortKey] - a[sortKey]).slice(0, 10);
    renderList('ofensorList', top10, 'nome');
}

function updateSalas() {
    const reprovados = filteredData.filter(i => {
        const s = i.status.toLowerCase();
        return s.includes('reprovad') || s.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(i => {
        if (isSummaryRow(i.sala) || i.sala === 'N/A') return;
        if (!grouped[i.sala]) grouped[i.sala] = { sala: i.sala, count: 0, valor: 0 };
        grouped[i.sala].count++;
        grouped[i.sala].valor += i.valor;
    });

    const sortKey = salaSortBy === 'valor' ? 'valor' : 'count';
    const top10 = Object.values(grouped).sort((a, b) => b[sortKey] - a[sortKey]).slice(0, 10);
    renderList('salaList', top10, 'sala');
}

function renderList(containerId, data, nameField) {
    const container = document.getElementById(containerId);
    if (data.length === 0) {
        container.innerHTML = '<p style="text-align:center;color:var(--text-secondary);padding:20px">Nenhum dado encontrado</p>';
        return;
    }

    container.innerHTML = data.map((item, i) => `
        <div class="ofensor-item">
            <div class="ofensor-rank ${i < 3 ? 'rank-' + (i + 1) : ''}">${i + 1}</div>
            <div class="ofensor-info">
                <div class="ofensor-name">${item[nameField] || 'N/A'}</div>
                ${nameField === 'nome' ? `<div class="ofensor-sala">${item.sala || ''}</div>` : ''}
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

    const periodos = {};
    filteredData.forEach(item => {
        if (!item.periodo) return;
        if (!periodos[item.periodo]) periodos[item.periodo] = { reprovado: 0, abonado: 0, analise: 0 };
        const s = item.status.toLowerCase();
        if (s.includes('reprovad') || s.includes('cobrança')) periodos[item.periodo].reprovado++;
        else if (s.includes('abonad')) periodos[item.periodo].abonado++;
        else if (s.includes('análise') || s.includes('analise')) periodos[item.periodo].analise++;
    });

    const labels = Object.keys(periodos).sort();

    evolutionChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                { label: 'Reprovados', data: labels.map(p => periodos[p].reprovado), backgroundColor: '#ff7675' },
                { label: 'Abonados', data: labels.map(p => periodos[p].abonado), backgroundColor: '#00cec9' },
                { label: 'Análise', data: labels.map(p => periodos[p].analise), backgroundColor: '#9b59b6' }
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

    const reprovados = filteredData.filter(i => {
        const s = i.status.toLowerCase();
        return s.includes('reprovad') || s.includes('cobrança');
    });

    const grouped = {};
    reprovados.forEach(i => {
        if (isSummaryRow(i.sala) || i.sala === 'N/A') return;
        if (!grouped[i.sala]) grouped[i.sala] = 0;
        grouped[i.sala]++;
    });

    const sorted = Object.entries(grouped).sort((a, b) => b[1] - a[1]).slice(0, 5);

    salaChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(s => s[0]),
            datasets: [{
                label: 'Reprovados',
                data: sorted.map(s => s[1]),
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
            plugins: { legend: { display: false } }
        }
    });
}

// ===== Utilities =====
function formatCurrency(value) {
    return new Intl.NumberFormat('pt-BR', {
        style: 'currency', currency: 'BRL',
        minimumFractionDigits: 0, maximumFractionDigits: 0
    }).format(value);
}

// ===== Service Worker =====
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('sw.js')
            .then(reg => console.log('SW registrado'))
            .catch(err => console.log('Erro SW:', err));
    });
}
