lucide.createIcons();

let rawExcelRows = [];    // Исходная сырая матрица аоа для выгрузки
let originalHeaders = [];    // Чистый массив заголовков
let processedDataset = [];   // Наш структурированный массив объектов
let currentFilter = 'all';
let searchQuery = '';
let sortDirection = 'none'; 
let outputFileName = 'compiled_campaign.xlsx';

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
const titleTooltip = document.getElementById('titleTooltip'); // Элемент для быстрого предпросмотра

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
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', codepage: 65001 });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        rawExcelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        
        if (rawExcelRows.length === 0) {
            alert("Файл пуст.");
            return;
        }

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

function computeRowMetrics(rowIndex, t1, t2, rowMap) {
    const cleanTitle2 = (t2 === '-' || t2 === '0') ? '' : t2;

    let combinedTitle = t1;
    let isMerged = false;
    let totalLength = t1.length;
    let overflow = 0;

    // Проверяем, пустой ли Заголовок 2
    if (!cleanTitle2) {
        return {
            rowIndex: rowIndex,
            t1: t1,
            t2: '',
            combined: t1,
            isMerged: true,
            length: t1.length,
            overflow: 0,
            statusType: 'no-t2', // Новый статус для пустых З2
            statusWeight: 4, 
            utpReasons: [],
            rowMap: rowMap 
        };
    }

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
        t1: t1,
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
            <div class="flex items-center gap-1.5">
                Статус переноса <span id="sortIndicator" class="text-indigo-600 font-mono text-xs">↕</span>
            </div>
        </th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[150px] z-30 border-r border-slate-200 min-w-[200px]">Итоговый заголовок</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[350px] z-30 border-r border-slate-200 text-center min-w-[110px]">Превышение</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[460px] z-30 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.05)] min-w-[150px]">Проблемы</th>
    `;

    originalHeaders.forEach(header => {
        html += `<th class="py-4 px-4 border-b border-slate-200 text-slate-600 font-semibold">${header || ''}</th>`;
    });

    tableHeaderRow.innerHTML = html;
    document.getElementById('sortStatusBtn').addEventListener('click', toggleSort);
    updateSortIndicator();
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

function renderFullTable() {
    tableBody.innerHTML = '';

    let displayData = [...processedDataset];
    
    // Новая логика фильтрации кнопок
    if (currentFilter === 'has-t2') {
        // Фильтр "Есть доп. заголовок": исключаем строки, где З2 пустой
        displayData = displayData.filter(item => item.statusType !== 'no-t2');
    } else if (currentFilter !== 'all') {
        displayData = displayData.filter(item => item.statusType === currentFilter);
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
        tableBody.innerHTML = `<tr><td colspan="${originalHeaders.length + 4}" class="py-12 text-center text-slate-400">Ничего не найдено.</td></tr>`;
        return;
    }

    const textColIdx = originalHeaders.indexOf(textHeaderName);
    const lenIndices = (textColIdx !== -1) ? [textColIdx + 1, textColIdx + 2, textColIdx + 3] : [];

    displayData.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/80 transition-colors group";
        tr.setAttribute('data-row-index', item.rowIndex);

        let statusBadge = '';
        let rowBgClass = '';
        let issueText = '—';

        if (item.statusType === 'no-t2') {
            // Вывод кастомного статуса для объявлений без З2
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-slate-100 text-slate-600 border border-slate-200"><i data-lucide="minus-circle" class="w-3 h-3"></i> Нет доп. заг-ка</span>`;
            rowBgClass = 'bg-slate-50/40 text-slate-400';
        } else if (item.statusType === 'lost-utp') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="alert-circle" class="w-3 h-3"></i> Доп. заголовок будет отброшен</span>`;
            rowBgClass = 'bg-rose-50/10 hover:bg-rose-50/20';
            issueText = `<span class="text-rose-600 font-semibold text-xs">Теряет УТП: ${item.utpReasons.join(', ')}</span>`;
        } else if (item.statusType === 'lost-safe') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="scissors" class="w-3 h-3"></i> Срез доп. заг-ка</span>`;
            rowBgClass = 'bg-amber-50/5 hover:bg-amber-50/15';
        } else {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Перенесется</span>`;
        }

        if (rowBgClass && item.statusType !== 'no-t2') tr.className = rowBgClass;
        if (item.statusType === 'no-t2') tr.className = rowBgClass + " transition-colors group";

        let rowHtml = `
            <td class="p-3 sticky-col sticky left-0 z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)]">${statusBadge}</td>
            <td class="p-3 sticky-col sticky left-[150px] z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] font-medium text-slate-900 max-w-[200px] truncate tooltip-target cursor-help" data-combined="${item.combined}">${formatTemplateText(item.combined)}</td>
            <td class="p-3 sticky-col sticky left-[350px] z-10 border-r border-slate-200 text-center font-mono font-bold ${item.overflow > 0 ? 'text-rose-600' : 'text-slate-300'}">${item.overflow > 0 ? `+${item.overflow}` : '0'}</td>
            <td class="p-3 sticky-col sticky left-[460px] z-10 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.04)]">${issueText}</td>
        `;

        originalHeaders.forEach((headerName, curIdx) => {
            let displayValue = item.rowMap[headerName] || '';
            
            let isT1 = (headerName === t1HeaderName);
            let isT2 = (headerName === t2HeaderName);
            
            let cellStyle = "p-3 text-slate-600 border-r border-slate-100 max-w-[250px] truncate";
            let editableAttr = "";
            let extraDataAttr = "";

            if (lenIndices.includes(curIdx)) {
                cellStyle += " font-mono font-semibold text-center bg-slate-50/50 text-indigo-600";
                if (curIdx === lenIndices[0]) {
                    displayValue = item.t1.length;
                    extraDataAttr = `data-len-type="t1"`;
                } else if (curIdx === lenIndices[1]) {
                    displayValue = item.t2.length;
                    extraDataAttr = `data-len-type="text2"`;
                } else if (curIdx === lenIndices[2]) {
                    displayValue = (item.rowMap[textHeaderName] || '').length;
                    extraDataAttr = `data-len-type="text"`;
                }
            }
            
            if (isT1) {
                cellStyle += " bg-indigo-50/40 text-slate-900 font-medium editable-cell cursor-text";
                editableAttr = `contenteditable="true" data-type="t1"`;
            }
            if (isT2) {
                cellStyle += " bg-amber-50/30 text-slate-900 font-medium editable-cell cursor-text";
                editableAttr = `contenteditable="true" data-type="t2"`;
            }

            rowHtml += `<td class="${cellStyle}" ${editableAttr} ${extraDataAttr} title="${isT1 || isT2 ? 'Кликните для редактирования' : displayValue}">${isT1 || isT2 ? formatTemplateText(String(displayValue)) : displayValue}</td>`;
        });

        tr.innerHTML = rowHtml;
        tableBody.appendChild(tr);
    });

    lucide.createIcons();
    initInlineEditingEvents();
    initTooltipEvents();
}

function initTooltipEvents() {
    const targets = tableBody.querySelectorAll('.tooltip-target');
    
    targets.forEach(target => {
        target.addEventListener('mouseenter', function() {
            const text = target.getAttribute('data-combined');
            if (!text) return;
            
            if (titleTooltip) {
                titleTooltip.innerHTML = formatTemplateText(text);
                titleTooltip.classList.remove('hidden');
                
                const rect = target.getBoundingClientRect();
                const tooltipRect = titleTooltip.getBoundingClientRect();
                
                let left = rect.left + window.scrollX + (rect.width / 2) - (tooltipRect.width / 2);
                let top = rect.top + window.scrollY - tooltipRect.height - 8;
                
                if (left < 10) left = 10;
                
                titleTooltip.style.left = `${left}px`;
                titleTooltip.style.top = `${top}px`;
            }
        });
        
        target.addEventListener('mouseleave', function() {
            if (titleTooltip) titleTooltip.classList.add('hidden');
        });
    });
}

function initInlineEditingEvents() {
    const cells = tableBody.querySelectorAll('.editable-cell');
    cells.forEach(cell => {
        
        cell.addEventListener('input', function(e) {
            const editType = cell.getAttribute('data-type');
            const text = cell.innerText;
            const len = text.length;

            const rect = cell.getBoundingClientRect();
            liveCounter.style.left = `${rect.left + window.scrollX}px`;
            liveCounter.style.top = `${rect.top + window.scrollY - 28}px`;
            liveCounter.innerText = `Длина: ${len} симв.`;
            
            if (editType === 't1' && len > 56) {
                liveCounter.className = "fixed z-50 bg-rose-600 text-white px-2.5 py-1 text-xs font-mono rounded-md shadow-lg pointer-events-none font-bold";
            } else {
                liveCounter.className = "fixed z-50 bg-slate-900 text-white px-2.5 py-1 text-xs font-mono rounded-md shadow-lg pointer-events-none font-bold";
            }

            const tr = cell.closest('tr');
            if (editType === 't1') {
                const lenT1Cell = tr.querySelector('td[data-len-type="t1"]');
                if (lenT1Cell) lenT1Cell.innerText = len;
            } else if (editType === 't2') {
                const lenT2Cell = tr.querySelector('td[data-len-type="text2"]');
                if (lenT2Cell) lenT2Cell.innerText = len;
            }
        });

        cell.addEventListener('focus', function(e) {
            const rowIndex = parseInt(cell.closest('tr').getAttribute('data-row-index'));
            const editType = cell.getAttribute('data-type');
            const dataItem = processedDataset.find(item => item.rowIndex === rowIndex);
            if (dataItem) {
                cell.innerText = editType === 't1' ? dataItem.t1 : dataItem.t2;
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
            const newText = cell.innerText.trim();

            const dataItem = processedDataset.find(item => item.rowIndex === rowIndex);
            if (!dataItem) return;

            if (editType === 't1') {
                dataItem.t1 = newText;
                dataItem.rowMap[t1HeaderName] = newText;
            } else if (editType === 't2') {
                dataItem.t2 = newText;
                dataItem.rowMap[t2HeaderName] = newText;
            }

            const updatedMetrics = computeRowMetrics(rowIndex, dataItem.t1, dataItem.t2, dataItem.rowMap);
            Object.assign(dataItem, updatedMetrics);

            updateDashboardStats();
            
            const combinedCell = tr.querySelector('td:nth-child(2)');
            const overflowCell = tr.querySelector('td:nth-child(3)');
            const issueCell = tr.querySelector('td:nth-child(4)');
            const statusCell = tr.querySelector('td:nth-child(1)');

            combinedCell.innerHTML = formatTemplateText(dataItem.combined);
            combinedCell.setAttribute('data-combined', dataItem.combined);
            
            overflowCell.innerText = dataItem.overflow > 0 ? `+${dataItem.overflow}` : '0';
            overflowCell.className = `p-3 sticky-col sticky left-[350px] z-10 border-r border-slate-200 text-center font-mono font-bold ${dataItem.overflow > 0 ? 'text-rose-600' : 'text-slate-300'}`;

            const lenT1Cell = tr.querySelector('td[data-len-type="t1"]');
            if (lenT1Cell) lenT1Cell.innerText = dataItem.t1.length;
            
            const lenT2Cell = tr.querySelector('td[data-len-type="text2"]');
            if (lenT2Cell) lenT2Cell.innerText = dataItem.t2.length;

            if (dataItem.statusType === 'no-t2') {
                statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-slate-100 text-slate-600 border border-slate-200"><i data-lucide="minus-circle" class="w-3 h-3"></i> Нет доп. заг-ка</span>`;
                issueCell.innerHTML = '—';
                tr.className = "bg-slate-50/40 text-slate-400 transition-colors group";
            } else if (dataItem.statusType === 'lost-utp') {
                statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="alert-circle" class="w-3 h-3"></i> Доп. заголовок будет отброшен</span>`;
                issueCell.innerHTML = `<span class="text-rose-600 font-semibold text-xs">Теряет УТП: ${dataItem.utpReasons.join(', ')}</span>`;
                tr.className = "bg-rose-50/10 hover:bg-rose-50/20 transition-colors group";
            } else if (dataItem.statusType === 'lost-safe') {
                statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="scissors" class="w-3 h-3"></i> Срез доп. заг-ка</span>`;
                issueCell.innerHTML = '—';
                tr.className = "bg-amber-50/5 hover:bg-amber-50/15 transition-colors group";
            } else {
                statusCell.innerHTML = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Перенесется</span>`;
                issueCell.innerHTML = '—';
                tr.className = "hover:bg-slate-50/80 transition-colors group";
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
    const lenIndices = (textColIdx !== -1) ? [textColIdx + 1, textColIdx + 2, textColIdx + 3] : [];

    processedDataset.forEach(item => {
        const singleRowArray = [];
        originalHeaders.forEach((headerName, curIdx) => {
            let val = item.rowMap[headerName] || '';
            
            if (lenIndices.includes(curIdx)) {
                if (curIdx === lenIndices[0]) val = item.t1.length;
                else if (curIdx === lenIndices[1]) val = item.t2.length;
                else if (curIdx === lenIndices[2]) val = (item.rowMap[textHeaderName] || '').length;
            }
            
            if (val !== '' && !isNaN(val) && (lenIndices.includes(curIdx) || headerName.toLowerCase().includes('длина'))) {
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
