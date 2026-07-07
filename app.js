lucide.createIcons();

let rawExcelRows = [];       // Исходная сырая матрица аоа для выгрузки
let originalHeaders = [];    // Чистый массив заголовков
let processedDataset = [];   // Наш структурированный массив объектов
let currentFilter = 'all';
let searchQuery = '';
let sortDirection = 'none'; 
let outputFileName = 'compiled_campaign.xlsx';
let selectedRowIndices = []; // Массив для хранения rowIndex выбранных объявлений

// Хранилище связей индексов
let t1HeaderName = 'Заголовок 1';
let t2HeaderName = 'Заголовок 2';
let textHeaderName = 'Текст';
let headerRowGlobalIndex = -1;

const fileInput = document.getElementById('excelFile');
const uploadText = document.getElementById('uploadText');
const tableHeaderRow = document.getElementById('tableHeaderRow');
const tableBody = document.getElementById('resultsTableBody');
const filterBtns = document.querySelectorAll('.filter-btn');
const downloadFileBtn = document.getElementById('downloadFileBtn');
const searchInput = document.getElementById('tableSearch');
const liveCounter = document.getElementById('liveCharCounter');

fileInput.addEventListener('change', handleFileSelect);
downloadFileBtn.addEventListener('click', downloadUpdatedXLSX);
searchInput.addEventListener('input', (e) => { searchQuery = e.target.value.toLowerCase().trim(); renderFullTable(); });

// Вешаем клик на кнопку массового изменения
const applyBulkBtn = document.getElementById('applyBulkBtn');
if (applyBulkBtn) {
    applyBulkBtn.addEventListener('click', applyBulkEdit);
}

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

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    uploadText.innerText = file.name;
    
    let extension = '.xlsx';
    if (file.name.endsWith('.xls')) extension = '.xls';
    if (file.name.endsWith('.csv')) extension = '.csv';
    
    const baseName = file.name.replace(/\.[^/.]+$/, "");
    outputFileName = `edited_${baseName}${extension}`;

    const reader = new FileReader();
    reader.onload = function(e) {
        let workbook;
        if (file.name.endsWith('.csv')) {
            const text = new TextDecoder('windows-1251').decode(e.target.result);
            const firstLine = text.split('\n')[0];
            const fs = (firstLine.split(';').length > firstLine.split(',').length) ? ';' : ',';
            workbook = XLSX.read(text, { type: 'string', FS: fs });
        } else {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
        }

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        rawExcelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        if (rawExcelRows.length === 0) {
            alert("Файл пуст или не удалось распознать структуру.");
            return;
        }

        selectedRowIndices = [];
        toggleBulkPanel();

        analyzeStructureAndProcess();
        downloadFileBtn.removeAttribute('disabled');
        downloadFileBtn.className = "bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-3 px-5 rounded-xl transition-all flex items-center gap-2 shadow-md text-sm active:scale-98 cursor-pointer";
    };
    reader.readAsArrayBuffer(file);
}

function analyzeStructureAndProcess() {
    let headerRowIndex = -1;

    for (let i = 0; i < Math.min(rawExcelRows.length, 30); i++) {
        if (!rawExcelRows[i]) continue;
        const rowStr = rawExcelRows[i].map(cell => String(cell || '').trim());
        
        const foundT1 = rowStr.indexOf(t1HeaderName);
        const hasIdCol = rowStr.some(c => c.includes('ID объявления') || c.includes('ID группы'));

        if (foundT1 !== -1 && hasIdCol) {
            headerRowIndex = i;
            headerRowGlobalIndex = i;
            originalHeaders = rowStr;
            break;
        }
    }

    if (headerRowIndex === -1) {
        alert("Не удалось найти строку заголовков с полем 'Заголовок 1'.");
        return;
    }

    let startDataRow = headerRowIndex + 1;
    if (startDataRow < rawExcelRows.length && rawExcelRows[startDataRow]) {
        const checkRow = rawExcelRows[startDataRow].map(c => String(c || '').toLowerCase().trim());
        if (checkRow.includes('заголовок 1') || checkRow.includes('текст') || checkRow.some(c => c === '55' || c === '35')) {
            startDataRow++;
        }
    }

    processedDataset = [];

    for (let i = startDataRow; i < rawExcelRows.length; i++) {
        const row = rawExcelRows[i];
        if (!row || row.length === 0) continue;

        const rowMap = {};
        originalHeaders.forEach((header, colIdx) => {
            rowMap[header] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });

        const hasIdOrGroup = originalHeaders.some(h => (h.includes('ID') || h.includes('Название группы') || h.includes('Тип объявления')) && rowMap[h]);
        if (!hasIdOrGroup && (!rowMap[t1HeaderName] || rowMap[t1HeaderName] === '-')) continue;

        let title1 = rowMap[t1HeaderName] || '';
        if (title1 === '-' || title1.startsWith('---')) title1 = '';

        const title2 = rowMap[t2HeaderName] || '';
        
        const analyzedRow = computeRowMetrics(i, title1, title2, rowMap);
        processedDataset.push(analyzedRow);
    }

    updateDashboardStats();
    buildTableHeader();
    renderFullTable();
}

function computeRowMetrics(rowIndex, t1, t2, rowMap) {
    const cleanTitle1 = (t1 === '-' || t1 === '0') ? '' : t1;
    const cleanTitle2 = (t2 === '-' || t2 === '0') ? '' : t2;

    let combinedTitle = cleanTitle1;
    let isMerged = false;
    let totalLength = cleanTitle1.length;
    let overflow = 0;

    if (cleanTitle1 && cleanTitle2) {
        const potential = cleanTitle1 + ". " + cleanTitle2;
        if (potential.length <= 56) {
            combinedTitle = potential;
            isMerged = true;
            totalLength = potential.length;
        } else {
            overflow = potential.length - 56;
        }
    } else if (!cleanTitle1 && cleanTitle2) {
        combinedTitle = cleanTitle2;
        isMerged = true;
        totalLength = cleanTitle2.length;
    } else {
        isMerged = true;
    }

    const utpAnalysis = analyzeUTPLoss(cleanTitle2);
    let statusType = 'success';
    let statusWeight = 1;

    if (!isMerged && cleanTitle2) {
        if (utpAnalysis.lost) {
            statusType = 'lost-utp';
            statusWeight = 3;
        } else {
            statusType = 'lost-safe';
            statusWeight = 2;
        }
    }

    return {
        rowIndex: rowIndex,
        t1: cleanTitle1,
        t2: cleanTitle2,
        combined: combinedTitle,
        isMerged: isMerged,
        length: totalLength,
        overflow: overflow,
        statusType: statusType,
        statusWeight: statusWeight, 
        utpReasons: utpAnalysis.reasons,
        rowMap: rowMap 
    };
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

    document.getElementById('statTotal').innerText = total;
    document.getElementById('statSuccess').innerText = success;
    document.getElementById('statCut').innerText = cut;
    document.getElementById('statLoss').innerText = loss;
}

function buildTableHeader() {
    let html = `
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-0 z-40 border-r border-slate-200 w-12 text-center">
            <input type="checkbox" id="selectAllCheckbox" class="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer">
        >
        <th id="sortStatusBtn" class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[48px] z-30 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.03)] min-w-[150px] cursor-pointer hover:bg-slate-200 transition-colors select-none">
            <div class="flex items-center gap-1.5">
                Статус переноса <span id="sortIndicator" class="text-indigo-600 font-mono text-xs">↕</span>
            </div>
        </th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[198px] z-30 border-r border-slate-200 min-w-[200px]">Итоговый заголовок</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[398px] z-30 border-r border-slate-200 text-center min-w-[110px]">Превышение</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[508px] z-30 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.05)] min-w-[150px]">Проблемы</th>
    `;

    originalHeaders.forEach(header => {
        html += `<th class="py-4 px-4 border-b border-slate-200 text-slate-600 font-semibold">${header || ''}</th>`;
    });

    tableHeaderRow.innerHTML = html;
    document.getElementById('sortStatusBtn').addEventListener('click', toggleSort);
    
    const selectAllCheck = document.getElementById('selectAllCheckbox');
    if (selectAllCheck) {
        selectAllCheck.addEventListener('change', handleSelectAll);
    }
    
    updateSortIndicator();
}

function handleSelectAll(e) {
    const isChecked = e.target.checked;
    const visibleCheckboxes = tableBody.querySelectorAll('.row-checkbox');
    
    visibleCheckboxes.forEach(cb => {
        cb.checked = isChecked;
        const rIdx = parseInt(cb.getAttribute('data-row-index'));
        
        if (isChecked) {
            if (!selectedRowIndices.includes(rIdx)) selectedRowIndices.push(rIdx);
        } else {
            selectedRowIndices = selectedRowIndices.filter(id => id !== rIdx);
        }
    });
    
    toggleBulkPanel();
}

function initCheckboxEvents() {
    const checkboxes = tableBody.querySelectorAll('.row-checkbox');
    checkboxes.forEach(cb => {
        cb.addEventListener('change', function() {
            const rIdx = parseInt(this.getAttribute('data-row-index'));
            if (this.checked) {
                if (!selectedRowIndices.includes(rIdx)) selectedRowIndices.push(rIdx);
            } else {
                selectedRowIndices = selectedRowIndices.filter(id => id !== rIdx);
                const selectAll = document.getElementById('selectAllCheckbox');
                if (selectAll) selectAll.checked = false;
            }
            toggleBulkPanel();
        });
    });
}

function toggleBulkPanel() {
    const panel = document.getElementById('bulkActionPanel');
    const countEl = document.getElementById('bulkSelectedCount');
    if (!panel) return;

    if (selectedRowIndices.length > 0) {
        panel.classList.remove('hidden');
        if (countEl) countEl.innerText = selectedRowIndices.length;
    } else {
        panel.classList.add('hidden');
        const selectAll = document.getElementById('selectAllCheckbox');
        if (selectAll) selectAll.checked = false;
    }
}

function applyBulkEdit() {
    const fieldType = document.getElementById('bulkFieldSelect').value; 
    const newValue = document.getElementById('bulkInputText').value.trim();

    if (selectedRowIndices.length === 0) return;
    
    if (!confirm(`Вы уверены, что хотите изменить поле для ${selectedRowIndices.length} объявлений?`)) return;

    selectedRowIndices.forEach(rowIndex => {
        const dataItem = processedDataset.find(item => item.rowIndex === rowIndex);
        if (!dataItem) return;

        if (fieldType === 't1') {
            dataItem.t1 = newValue;
            dataItem.rowMap[t1HeaderName] = newValue;
        } else if (fieldType === 't2') {
            dataItem.t2 = newValue;
            dataItem.rowMap[t2HeaderName] = newValue;
        }

        const updatedMetrics = computeRowMetrics(rowIndex, dataItem.t1, dataItem.t2, dataItem.rowMap);
        Object.assign(dataItem, updatedMetrics);
    });

    selectedRowIndices = [];
    document.getElementById('bulkInputText').value = '';
    
    updateDashboardStats();
    renderFullTable();
    toggleBulkPanel();
}

function toggleSort() {
    if (sortDirection === 'none') sortDirection = 'asc';
    else if (sortDirection === 'asc') sortDirection = 'desc';
    else sortDirection = 'none';
    
    updateSortIndicator();
    renderFullTable();
}

function updateSortIndicator() {
    const indicator = document.getElementById('sortIndicator');
    if (!indicator) return;
    if (sortDirection === 'asc') indicator.innerText = '↑';
    else if (sortDirection === 'desc') indicator.innerText = '↓';
    else indicator.innerText = '↕';
}

function formatTemplateText(text) {
    if (!text.includes('#')) return text;
    return text.replace(/#([^#]+)#/g, '<span class="yandex-template">#$1#</span>');
}

// Функция для сопоставления текстового столбца и столбца его расчетной длины
function getLengthColumnIndex(headerName) {
    if (!headerName) return -1;
    
    // Для ТГО объявлений (Заголовок 1, Заголовок 2, Текст)
    if (headerName === t1HeaderName) {
        const t1Idx = originalHeaders.indexOf(t1HeaderName);
        const textIdx = originalHeaders.indexOf(textHeaderName);
        return (textIdx !== -1) ? textIdx + 1 : -1;
    }
    if (headerName === t2HeaderName) {
        const t1Idx = originalHeaders.indexOf(t1HeaderName);
        const textIdx = originalHeaders.indexOf(textHeaderName);
        return (textIdx !== -1) ? textIdx + 2 : -1;
    }
    if (headerName === textHeaderName) {
        const textIdx = originalHeaders.indexOf(textHeaderName);
        return (textIdx !== -1) ? textIdx + 3 : -1;
    }

    // Для комбинаторных объявлений (Заголовок 3..7, Текст 1..3)
    // Ищем точное совпадение: например, "Длина Заголовка 3" или "Длина Текста 1"
    const targetLengthHeader = `Длина ${headerName.toLowerCase()}`;
    const foundIdx = originalHeaders.findIndex(h => h.toLowerCase().trim() === targetLengthHeader);
    if (foundIdx !== -1) return foundIdx;

    // Альтернативный поиск по вхождению (на случай сокращений в шаблонах экспорта)
    const numMatch = headerName.match(/\d+/);
    if (numMatch) {
        const num = numMatch[0];
        if (headerName.toLowerCase().includes('заголовок')) {
            return originalHeaders.findIndex(h => h.toLowerCase().includes('длина') && h.toLowerCase().includes('заголовок') && h.includes(num));
        }
        if (headerName.toLowerCase().includes('текст')) {
            return originalHeaders.findIndex(h => h.toLowerCase().includes('длина') && h.toLowerCase().includes('текст') && h.includes(num));
        }
    }

    return -1;
}

function renderFullTable() {
    tableBody.innerHTML = '';

    let displayData = [...processedDataset];
    
    if (currentFilter !== 'all') {
        if (currentFilter === 'has-t2') {
            displayData = displayData.filter(item => item.t2 && item.t2 !== '');
        } else {
            displayData = displayData.filter(item => item.statusType === currentFilter);
        }
    }

    if (searchQuery) {
        displayData = displayData.filter(item => 
            item.t1.toLowerCase().includes(searchQuery) || 
            item.t2.toLowerCase().includes(searchQuery) ||
            item.combined.toLowerCase().includes(searchQuery)
        );
    }

    if (sortDirection === 'asc') displayData.sort((a, b) => a.statusWeight - b.statusWeight);
    else if (sortDirection === 'desc') displayData.sort((a, b) => b.statusWeight - a.statusWeight);

    document.getElementById('tableCounter').innerText = `Строк: ${displayData.length}`;

    if (displayData.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="${originalHeaders.length + 5}" class="py-12 text-center text-slate-400">Ничего не найдено.</td></tr>`;
        return;
    }

    const selectAllCheck = document.getElementById('selectAllCheckbox');
    if (selectAllCheck) {
        const pageRowIds = displayData.map(d => d.rowIndex);
        const allPageRowsChecked = pageRowIds.length > 0 && pageRowIds.every(id => selectedRowIndices.includes(id));
        selectAllCheck.checked = allPageRowsChecked;
    }

    // Собираем все индексы колонок, которые отвечают за длины
    const textColIdx = originalHeaders.indexOf(textHeaderName);
    const tgoLenIndices = (textColIdx !== -1) ? [textColIdx + 1, textColIdx + 2, textColIdx + 3] : [];

    displayData.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/80 transition-colors group";
        tr.setAttribute('data-row-index', item.rowIndex);

        let statusBadge = '';
        let rowBgClass = '';
        let issueText = '—';

        if (item.statusType === 'lost-utp') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="alert-circle" class="w-3 h-3"></i> Доп. заголовок будет отброшен</span>`;
            rowBgClass = 'bg-rose-50/10 hover:bg-rose-50/20';
            issueText = `<span class="text-rose-600 font-semibold text-xs">Теряет УТП: ${item.utpReasons.join(', ')}</span>`;
        } else if (item.statusType === 'lost-safe') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="scissors" class="w-3 h-3"></i> Срез доп. заг-ка</span>`;
            rowBgClass = 'bg-amber-50/5 hover:bg-amber-50/15';
        } else {
            if (!item.t2 || item.t2.trim() === '') {
                statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Нет второго заголовка</span>`;
            } else {
                statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Перенесется</span>`;
            }
        }

        if (rowBgClass) tr.className = rowBgClass;

        const isChecked = selectedRowIndices.includes(item.rowIndex);

        let rowHtml = `
            <td class="p-3 sticky-col sticky left-0 z-10 border-r border-slate-200 bg-white text-center">
                <input type="checkbox" class="row-checkbox rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer" data-row-index="${item.rowIndex}" ${isChecked ? 'checked' : ''}>
            </td>
            <td class="p-3 sticky-col sticky left-[48px] z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)]">${statusBadge}</td>
            <td class="p-3 sticky-col sticky left-[198px] z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] font-medium text-slate-900 max-w-[200px] truncate" title="${item.combined}">${formatTemplateText(item.combined)}</td>
            <td class="p-3 sticky-col sticky left-[398px] z-10 border-r border-slate-200 text-center font-mono font-bold ${item.overflow > 0 ? 'text-rose-600' : 'text-slate-300'}">${item.overflow > 0 ? `+${item.overflow}` : '0'}</td>
            <td class="p-3 sticky-col sticky left-[508px] z-10 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.04)]">${issueText}</td>
        `;

        originalHeaders.forEach((headerName, curIdx) => {
            let displayValue = item.rowMap[headerName] || '';
            let headerLower = headerName.toLowerCase();
            
            let isT1 = (headerName === t1HeaderName);
            let isT2 = (headerName === t2HeaderName);
            
            // Проверяем, является ли текущее поле комбинаторным текстовым полем или его длиной
            let isCombText = /заголовок\s*[3-7]|текст\s*[1-3]/i.test(headerName) && !headerLower.includes('длина');
            let isLengthCol = tgoLenIndices.includes(curIdx) || headerLower.includes('длина') || headerLower.startsWith('дл. ');

            let cellStyle = "p-3 text-slate-600 border-r border-slate-100 max-w-[250px] truncate";
            let editableAttr = "";
            let extraDataAttr = "";

            if (isLengthCol) {
                cellStyle += " font-mono font-semibold text-center bg-slate-50/50 text-indigo-600";
                
                // Подставляем актуальные расчетные длины для ТГО из стейта
                if (curIdx === tgoLenIndices[0]) {
                    displayValue = item.t1.length;
                    extraDataAttr = `data-len-type="t1"`;
                } else if (curIdx === tgoLenIndices[1]) {
                    displayValue = item.t2.length;
                    extraDataAttr = `data-len-type="text2"`;
                } else if (curIdx === tgoLenIndices[2]) {
                    displayValue = (item.rowMap[textHeaderName] || '').length;
                    extraDataAttr = `data-len-type="text"`;
                } else {
                    // Для комбинаторных длин выводим длину связанного текстового поля
                    // Ищем имя оригинального текстового поля, отрезав слово "Длина "
                    let textHeaderKey = originalHeaders.find(h => `длина ${h.toLowerCase()}` === headerLower || headerLower.includes(h.toLowerCase()));
                    if (!textHeaderKey) {
                        // Фолбэк-определение по цифре
                        let m = headerName.match(/\d+/);
                        let type = headerLower.includes('заг') ? 'Заголовок' : 'Текст';
                        textHeaderKey = m ? `${type} ${m[0]}` : '';
                    }
                    displayValue = (item.rowMap[textHeaderKey] || '').length;
                    extraDataAttr = `data-len-type="comb-len" data-source-field="${textHeaderKey || ''}"`;
                }
            }
            
            if (isT1) {
                cellStyle += " bg-indigo-50/40 text-slate-900 font-medium editable-cell cursor-text";
                editableAttr = `contenteditable="true" data-type="t1" data-field-name="${headerName}"`;
            } else if (isT2) {
                cellStyle += " bg-amber-50/30 text-slate-900 font-medium editable-cell cursor-text";
                editableAttr = `contenteditable="true" data-type="t2" data-field-name="${headerName}"`;
            } else if (isCombText) {
                // Разрешаем редактирование Заголовков 3-7 и Текстов 1-3 прямо в ячейке!
                cellStyle += " bg-slate-50/60 text-slate-900 font-medium editable-cell cursor-text border-dashed border-b border-slate-300";
                editableAttr = `contenteditable="true" data-type="comb-text" data-field-name="${headerName}"`;
            }

            let showTooltip = isT1 || isT2 || isCombText ? 'Кликните для редактирования' : displayValue;
            let finalContent = isT1 || isT2 || isCombText ? formatTemplateText(String(displayValue)) : displayValue;

            rowHtml += `<td class="${cellStyle}" ${editableAttr} ${extraDataAttr} title="${showTooltip}">${finalContent}</td>`;
        });

        tr.innerHTML = rowHtml;
        tableBody.appendChild(tr);
    });

    lucide.createIcons();
    initInlineEditingEvents();
    initCheckboxEvents();
}

function initInlineEditingEvents() {
    const cells = tableBody.querySelectorAll('.editable-cell');
    cells.forEach(cell => {
        
        cell.addEventListener('input', function(e) {
            const editType = cell.getAttribute('data-type');
            const fieldName = cell.getAttribute('data-field-name');
            const text = cell.innerText;
            const len = text.length;

            const rect = cell.getBoundingClientRect();
            liveCounter.style.left = `${rect.left + window.scrollX}px`;
            liveCounter.style.top = `${rect.top + window.scrollY - 28}px`;
            liveCounter.innerText = `Длина: ${len} симв.`;
            
            // Индикация превышения лимита заголовков (56 символов)
            if ((editType === 't1' || fieldName.toLowerCase().includes('заголовок')) && len > 56) {
                liveCounter.className = "fixed z-50 bg-rose-600 text-white px-2.5 py-1 text-xs font-mono rounded-md shadow-lg pointer-events-none font-bold";
            } else {
                liveCounter.className = "fixed z-50 bg-slate-900 text-white px-2.5 py-1 text-xs font-mono rounded-md shadow-lg pointer-events-none font-bold";
            }

            const tr = cell.closest('tr');
            
            // Динамически ищем и обновляем нужную ячейку длины на экране
            if (editType === 't1') {
                const lenT1Cell = tr.querySelector('td[data-len-type="t1"]');
                if (lenT1Cell) lenT1Cell.innerText = len;
            } else if (editType === 't2') {
                const lenT2Cell = tr.querySelector('td[data-len-type="text2"]');
                if (lenT2Cell) lenT2Cell.innerText = len;
            } else if (editType === 'comb-text') {
                const lenCombCell = tr.querySelector(`td[data-source-field="${fieldName}"]`);
                if (lenCombCell) lenCombCell.innerText = len;
            }
        });

        cell.addEventListener('focus', function(e) {
            const rowIndex = parseInt(cell.closest('tr').getAttribute('data-row-index'));
            const editType = cell.getAttribute('data-type');
            const fieldName = cell.getAttribute('data-field-name');
            const dataItem = processedDataset.find(item => item.rowIndex === rowIndex);
            
            if (dataItem) {
                if (editType === 't1') cell.innerText = dataItem.t1;
                else if (editType === 't2') cell.innerText = dataItem.t2;
                else cell.innerText = dataItem.rowMap[fieldName] || '';
            }

            liveCounter.innerText = `Длина: ${cell.innerText.length} симв.`;
            const rect = cell.getBoundingClientRect();
            liveCounter.style.left = `${rect.left + window.scrollX}px`;
            liveCounter.style.top = `${rect.top + window.scrollY - 28}px`;
            liveCounter.classList.remove('hidden');
        });

        cell.addEventListener('blur', function(e) {
            liveCounter.classList.add('hidden');
            
            const tr = cell.closest('tr');
            const rowIndex = parseInt(tr.getAttribute('data-row-index'));
            const editType = cell.getAttribute('data-type');
            const fieldName = cell.getAttribute('data-field-name');
            const newText = cell.innerText.trim();

            const dataItem = processedDataset.find(item => item.rowIndex === rowIndex);
            if (!dataItem) return;

            // Записываем измененное значение обратно в модель данных
            if (editType === 't1') {
                dataItem.t1 = newText;
                dataItem.rowMap[t1HeaderName] = newText;
            } else if (editType === 't2') {
                dataItem.t2 = newText;
                dataItem.rowMap[t2HeaderName] = newText;
            } else {
                dataItem.rowMap[fieldName] = newText;
            }

            // Пересчитываем базовые метрики ТГО
            const updatedMetrics = computeRowMetrics(rowIndex, dataItem.t1, dataItem.t2, dataItem.rowMap);
            Object.assign(dataItem, updatedMetrics);

            updateDashboardStats();
            
            // Локальное обновление статусных колонок ТГО на лету
            const combinedCell = tr.querySelector('td:nth-child(3)'); 
            const overflowCell = tr.querySelector('td:nth-child(4)'); 
            const issueCell = tr.querySelector('td:nth-child(5)');    
            const statusCell = tr.querySelector('td:nth-child(2)');   

            if (combinedCell) {
                combinedCell.innerHTML = formatTemplateText(dataItem.combined);
                combinedCell.setAttribute('title', dataItem.combined);
            }
            
            if (overflowCell) {
                overflowCell.innerText = dataItem.overflow > 0 ? `+${dataItem.overflow}` : '0';
                overflowCell.className = `p-3 sticky-col sticky left-[398px] z-10 border-r border-slate-200 text-center font-mono font-bold ${dataItem.overflow > 0 ? 'text-rose-600' : 'text-slate-300'}`;
            }

            const lenT1Cell = tr.querySelector('td[data-len-type="t1"]');
            if (lenT1Cell) lenT1Cell.innerText = dataItem.t1.length;
            
            const lenT2Cell = tr.querySelector('td[data-len-type="text2"]');
            if (lenT2Cell) lenT2Cell.innerText = dataItem.t2.length;

            if (statusCell && issueCell) {
                if (dataItem.statusType === 'lost-utp') {
                    statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="alert-circle" class="w-3 h-3"></i> Доп. заголовок будет отброшен</span>`;
                    issueCell.innerHTML = `<span class="text-rose-600 font-semibold text-xs">Теряет УТП: ${dataItem.utpReasons.join(', ')}</span>`;
                    tr.className = "bg-rose-50/10 hover:bg-rose-50/20 transition-colors group";
                } else if (dataItem.statusType === 'lost-safe') {
                    statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="scissors" class="w-3 h-3"></i> Срез доп. заг-ка</span>`;
                    issueCell.innerHTML = '—';
                    tr.className = "bg-amber-50/5 hover:bg-amber-50/15 transition-colors group";
                } else {
                    if (!dataItem.t2 || dataItem.t2.trim() === '') {
                        statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Нет второго заголовка</span>`;
                    } else {
                        statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Перенесется</span>`;
                    }
                    issueCell.innerHTML = '—';
                    tr.className = "hover:bg-slate-50/80 transition-colors group";
                }
            }

            cell.innerHTML = formatTemplateText(newText);
            lucide.createIcons();
        });

        cell.addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                cell.blur();
            }
        });
    });
}

function downloadUpdatedXLSX() {
    if (processedDataset.length === 0) return;

    const exportRows = [];
    
    for (let i = 0; i <= headerRowGlobalIndex; i++) {
        exportRows.push(rawExcelRows[i]);
    }

    const textColIdx = originalHeaders.indexOf(textHeaderName);
    const tgoLenIndices = (textColIdx !== -1) ? [textColIdx + 1, textColIdx + 2, textColIdx + 3] : [];

    processedDataset.forEach(item => {
        const singleRowArray = [];
        originalHeaders.forEach((headerName, curIdx) => {
            let val = item.rowMap[headerName] || '';
            let headerLower = headerName.toLowerCase();
            
            // Динамический пересчет длин ячеек при сохранении
            if (tgoLenIndices.includes(curIdx)) {
                if (curIdx === tgoLenIndices[0]) val = item.t1.length;
                else if (curIdx === tgoLenIndices[1]) val = item.t2.length;
                else if (curIdx === tgoLenIndices[2]) val = (item.rowMap[textHeaderName] || '').length;
            } else if (headerLower.includes('длина') || headerLower.startsWith('дл. ')) {
                // Вычисляем длину связанного комбинаторного текстового поля
                let textHeaderKey = originalHeaders.find(h => `длина ${h.toLowerCase()}` === headerLower || headerLower.includes(h.toLowerCase()));
                if (!textHeaderKey) {
                    let m = headerName.match(/\d+/);
                    let type = headerLower.includes('заг') ? 'Заголовок' : 'Текст';
                    textHeaderKey = m ? `${type} ${m[0]}` : '';
                }
                if (textHeaderKey && item.rowMap[textHeaderKey] !== undefined) {
                    val = item.rowMap[textHeaderKey].length;
                }
            }
            
            // Сохраняем числа как числовые типы данных, чтобы Excel не ругался
            if (val !== '' && !isNaN(val) && (tgoLenIndices.includes(curIdx) || headerLower.includes('длина') || headerLower.startsWith('дл. '))) {
                singleRowArray.push(Number(val));
            } else {
                singleRowArray.push(val);
            }
        });
        exportRows.push(singleRowArray);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(exportRows);
    
    if (outputFileName.endsWith('.csv')) {
        XLSX.utils.book_append_sheet(wb, ws, "Кампания");
        XLSX.writeFile(wb, outputFileName, { bookType: 'csv', FS: ';' }); 
    } else {
        XLSX.utils.book_append_sheet(wb, ws, "Кампания");
        XLSX.writeFile(wb, outputFileName);
    }
}
