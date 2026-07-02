// --- Глобальные переменные ---
let rawExcelRows = [];
let originalHeaders = [];
let processedDataset = [];
let currentFilter = 'all';
let searchQuery = '';
let sortDirection = 'none';
let outputFileName = 'compiled_campaign.xlsx';
let header1Index = -1; // Исправлено: добавлена переменная
let headerRowGlobalIndex = -1;

// Хранилище связей индексов
const t1HeaderName = 'Заголовок 1';
const t2HeaderName = 'Заголовок 2';
const textHeaderName = 'Текст';

const fileInput = document.getElementById('excelFile');
const uploadText = document.getElementById('uploadText');
const tableHeaderRow = document.getElementById('tableHeaderRow');
const tableBody = document.getElementById('resultsTableBody');
const filterBtns = document.querySelectorAll('.filter-btn');
const downloadFileBtn = document.getElementById('downloadFileBtn');
const searchInput = document.getElementById('tableSearch');
const liveCounter = document.getElementById('liveCharCounter');

// --- Инициализация ---
lucide.createIcons();
fileInput.addEventListener('change', handleFileSelect);
downloadFileBtn.addEventListener('click', downloadUpdatedXLSX);
searchInput.addEventListener('input', (e) => { searchQuery = e.target.value.toLowerCase().trim(); renderFullTable(); });

filterBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
        filterBtns.forEach(b => b.classList.remove('bg-indigo-600', 'text-white', 'shadow-xs'));
        filterBtns.forEach(b => b.classList.add('bg-slate-100', 'text-slate-600', 'hover:bg-slate-200'));
        e.target.classList.remove('bg-slate-100', 'text-slate-600', 'hover:bg-slate-200');
        e.target.classList.add('bg-indigo-600', 'text-white', 'shadow-xs');
        currentFilter = e.target.getAttribute('data-filter');
        renderFullTable();
    });
});

// --- Функции парсинга ---
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    uploadText.innerText = file.name;
    const reader = new FileReader();

    reader.onload = function(e) {
        if (file.name.match(/\.(xlsx|xls)$/)) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            rawExcelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        } else {
            const text = e.target.result;
            const lines = text.split(/\r?\n/);
            const headerRowIndex = lines.findIndex(l => l.includes('Заголовок 1'));
            
            if (headerRowIndex === -1) {
                alert("Ошибка: Не могу найти 'Заголовок 1'.");
                return;
            }

            const headerLine = lines[headerRowIndex];
            const separator = ['\t', ';', ',', '|'].find(sep => headerLine.split(sep).length > 5) || '\t';
            
            rawExcelRows = lines
                .slice(headerRowIndex)
                .filter(line => line.trim() !== '')
                .map(line => line.split(separator).map(cell => cell.trim()));
        }

        if (rawExcelRows && rawExcelRows.length > 0) {
            analyzeStructureAndProcess();
            downloadFileBtn.removeAttribute('disabled');
            downloadFileBtn.className = "bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-3 px-5 rounded-xl transition-all flex items-center gap-2 shadow-md text-sm active:scale-98 cursor-pointer";
        }
    };

    if (file.name.match(/\.(xlsx|xls)$/)) {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file, 'windows-1251');
    }
}

function analyzeStructureAndProcess() {
    let headerRowIndex = -1;

    for (let i = 0; i < Math.min(rawExcelRows.length, 30); i++) {
        if (!rawExcelRows[i]) continue;
        const rowStr = rawExcelRows[i].map(cell => String(cell || '').trim());
        
        const foundT1Index = rowStr.findIndex(cell => cell === "Заголовок 1");
        const hasIdCol = rowStr.some(c => c.includes('ID объявления') || c.includes('ID группы'));

        if (foundT1Index !== -1 && hasIdCol) {
            headerRowIndex = i;
            headerRowGlobalIndex = i;
            originalHeaders = rowStr;
            header1Index = foundT1Index;
            break;
        }
    }

    if (headerRowIndex === -1) {
        alert("Не удалось найти строку заголовков.");
        return;
    }

    let startDataRow = headerRowIndex + 1;
    processedDataset = [];

    for (let i = startDataRow; i < rawExcelRows.length; i++) {
        const row = rawExcelRows[i];
        if (!row || row.length === 0) continue;

        const rowMap = {};
        originalHeaders.forEach((header, colIdx) => {
            rowMap[header] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });

        const title1 = rowMap[t1HeaderName] || '';
        if (!title1 || title1 === '-' || title1.startsWith('---')) continue; 

        const title2 = rowMap[t2HeaderName] || '';
        const analyzedRow = computeRowMetrics(i, title1, title2, rowMap);
        processedDataset.push(analyzedRow);
    }

    updateDashboardStats();
    buildTableHeader();
    renderFullTable();
}

// --- Вспомогательные функции (остаются без изменений) ---
function computeRowMetrics(rowIndex, t1, t2, rowMap) {
    const cleanTitle2 = (t2 === '-' || t2 === '0') ? '' : t2;
    let combinedTitle = t1;
    let isMerged = false;
    let totalLength = t1.length;
    let overflow = 0;

    if (cleanTitle2) {
        const potential = t1 + ". " + cleanTitle2;
        if (potential.length <= 56) {
            combinedTitle = potential;
            isMerged = true;
            totalLength = potential.length;
        } else {
            overflow = potential.length - 56;
        }
    } else {
        isMerged = true;
    }

    const utpAnalysis = analyzeUTPLoss(cleanTitle2);
    let statusType = 'success';
    let statusWeight = 1;

    if (!isMerged && cleanTitle2) {
        if (utpAnalysis.lost) { statusType = 'lost-utp'; statusWeight = 3; }
        else { statusType = 'lost-safe'; statusWeight = 2; }
    }

    return { rowIndex, t1, t2, combined: combinedTitle, isMerged, length: totalLength, overflow, statusType, statusWeight, utpReasons: utpAnalysis.reasons, rowMap };
}

function analyzeUTPLoss(t2) {
    if (!t2) return { lost: false, reasons: [] };
    const reasons = [];
    if (/[0-9]+\s*(?:руб|₽|\$|€|тыс|коп)/i.test(t2) || /(?:от|до|цена)\s*[0-9]+/i.test(t2)) reasons.push("Цена");
    if (/%\s*|скидк|акци|дисконт|процент/i.test(t2)) reasons.push("Скидка");
    if (/до\s+[0-9а-я]+|успей|дня|дне|только|срок/i.test(t2)) reasons.push("Сроки");
    if (/наличи|склад|осталось|в наличии/i.test(t2)) reasons.push("Наличие");
    if (/бесплат| 0 руб/i.test(t2)) reasons.push("Бесплатно");
    return { lost: reasons.length > 0, reasons: reasons };
}

function updateDashboardStats() {
    const total = processedDataset.length;
    const success = processedDataset.filter(d => d.statusType === 'success').length;
    const cut = processedDataset.filter(d => d.statusType === 'lost-safe').length;
    const loss = processedDataset.filter(d => d.statusType === 'lost-utp').length;
    const pct = total > 0 ? Math.round((success / total) * 100) : 0;

    document.getElementById('statTotal').innerText = total;
    document.getElementById('statSuccess').innerText = success;
    document.getElementById('statSuccessPct').innerText = `${pct}% от всех объявлений`;
    document.getElementById('statCut').innerText = cut;
    document.getElementById('statLoss').innerText = loss;
}

function buildTableHeader() {
    let html = `
        <th id="sortStatusBtn" class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-0 z-30 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.03)] min-w-[150px] cursor-pointer hover:bg-slate-200 transition-colors select-none">
            <div class="flex items-center gap-1.5">Статус переноса <span id="sortIndicator" class="text-indigo-600 font-mono text-xs">↕</span></div>
        </th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[150px] z-30 border-r border-slate-200 min-w-[200px]">Итоговый заголовок</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[350px] z-30 border-r border-slate-200 text-center min-w-[110px]">Превышение</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[460px] z-30 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.05)] min-w-[150px]">Проблемы</th>
    `;
    originalHeaders.forEach(header => html += `<th class="py-4 px-4 border-b border-slate-200 text-slate-600 font-semibold">${header || ''}</th>`);
    tableHeaderRow.innerHTML = html;
    document.getElementById('sortStatusBtn').addEventListener('click', toggleSort);
}

function toggleSort() {
    sortDirection = sortDirection === 'none' ? 'asc' : (sortDirection === 'asc' ? 'desc' : 'none');
    renderFullTable();
}

function formatTemplateText(text) {
    if (!text.includes('#')) return text;
    return text.replace(/#([^#]+)#/g, '<span class="yandex-template">#$1#</span>');
}

function renderFullTable() {
    tableBody.innerHTML = '';
    let displayData = [...processedDataset];
    
    if (currentFilter !== 'all') displayData = displayData.filter(item => item.statusType === currentFilter);
    if (searchQuery) displayData = displayData.filter(item => item.t1.toLowerCase().includes(searchQuery) || item.t2.toLowerCase().includes(searchQuery) || item.combined.toLowerCase().includes(searchQuery));
    if (sortDirection === 'asc') displayData.sort((a, b) => a.statusWeight - b.statusWeight);
    else if (sortDirection === 'desc') displayData.sort((a, b) => b.statusWeight - a.statusWeight);

    document.getElementById('tableCounter').innerText = `Строк: ${displayData.length}`;

    const textColIdx = originalHeaders.indexOf(textHeaderName);
    const lenIndices = (textColIdx !== -1) ? [textColIdx + 1, textColIdx + 2, textColIdx + 3] : [];

    displayData.forEach(item => {
        const tr = document.createElement('tr');
        tr.setAttribute('data-row-index', item.rowIndex);
        // ... (логика отрисовки строк остается вашей, из предыдущего сообщения)
        // ВАЖНО: убедитесь, что вы здесь используете ту же структуру rowHtml, что и была
        tableBody.appendChild(tr);
    });
    initInlineEditingEvents();
}
