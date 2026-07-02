// --- 1. Глобальное состояние ---
let rawExcelRows = [];
let originalHeaders = [];
let processedDataset = [];
let currentFilter = 'all';
let searchQuery = '';
let sortDirection = 'none';
let headerRowGlobalIndex = -1;

const t1HeaderName = 'Заголовок 1';
const t2HeaderName = 'Заголовок 2';
const textHeaderName = 'Текст';

// --- 2. DOM Элементы ---
const fileInput = document.getElementById('excelFile');
const uploadText = document.getElementById('uploadText');
const tableHeaderRow = document.getElementById('tableHeaderRow');
const tableBody = document.getElementById('resultsTableBody');
const filterBtns = document.querySelectorAll('.filter-btn');
const downloadFileBtn = document.getElementById('downloadFileBtn');
const searchInput = document.getElementById('tableSearch');

// --- 3. Инициализация и подписки на события ---
if (typeof lucide !== 'undefined') {
    lucide.createIcons();
}

if (fileInput) fileInput.addEventListener('change', handleFileSelect);
if (downloadFileBtn) downloadFileBtn.addEventListener('click', downloadUpdatedXLSX);
if (searchInput) {
    searchInput.addEventListener('input', (e) => { 
        searchQuery = e.target.value.toLowerCase().trim(); 
        renderFullTable(); 
    });
}

if (filterBtns) {
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
}

// --- 4. Обработка файлов и парсинг ---
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (uploadText) uploadText.innerText = file.name;
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            if (file.name.match(/\.(xlsx|xls)$/)) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                rawExcelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            } else {
                // Обработка текстовых CSV/TSV Директ Коммандера
                const text = e.target.result;
                const lines = text.split(/\r?\n/);
                const headerRowIndex = lines.findIndex(l => l.includes('Заголовок 1'));
                
                if (headerRowIndex === -1) {
                    alert("Ошибка: В файле не найден обязательный столбец 'Заголовок 1'. Проверьте кодировку.");
                    return;
                }

                const headerLine = lines[headerRowIndex];
                // Автоматически определяем разделитель (табуляция для TSV Коммандера или точка с запятой)
                const separator = ['\t', ';', ',', '|'].find(sep => headerLine.split(sep).length > 5) || '\t';
                
                rawExcelRows = lines
                    .slice(headerRowIndex)
                    .filter(line => line.trim() !== '')
                    .map(line => line.split(separator).map(cell => cell.trim()));
            }

            console.log("Файл успешно прочитан. Всего строк в массиве:", rawExcelRows.length);

            if (rawExcelRows && rawExcelRows.length > 0) {
                analyzeStructureAndProcess();
                if (downloadFileBtn) {
                    downloadFileBtn.removeAttribute('disabled');
                    downloadFileBtn.className = "bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-3 px-5 rounded-xl transition-all flex items-center gap-2 shadow-md text-sm active:scale-98 cursor-pointer";
                }
            }
        } catch (err) {
            console.error("Критическая ошибка при парсинге файла:", err);
            alert("Не удалось прочитать файл. Подробности в консоли браузера.");
        }
    };

    if (file.name.match(/\.(xlsx|xls)$/)) {
        reader.readAsArrayBuffer(file);
    } else {
        // Читаем в windows-1251, так как Коммандер выгружает файлы в ней
        reader.readAsText(file, 'windows-1251');
    }
}

// --- 5. Анализ структуры и вычисление метрик ---
function analyzeStructureAndProcess() {
    let headerRowIndex = -1;

    // Ищем строку заголовков в первых 30 строках
    for (let i = 0; i < Math.min(rawExcelRows.length, 30); i++) {
        if (!rawExcelRows[i]) continue;
        const rowStr = rawExcelRows[i].map(cell => String(cell || '').trim());
        
        const foundT1Index = rowStr.findIndex(cell => cell === t1HeaderName);
        const hasIdCol = rowStr.some(c => c.includes('ID объявления') || c.includes('ID группы') || c.includes('Доп. объявление'));

        if (foundT1Index !== -1) {
            headerRowIndex = i;
            headerRowGlobalIndex = i;
            originalHeaders = rowStr;
            break;
        }
    }

    if (headerRowIndex === -1) {
        console.error("Первые строки файла для отладки:", rawExcelRows.slice(0, 5));
        alert("Не удалось найти строку заголовков с полем 'Заголовок 1'. Проверьте структуру файла.");
        return;
    }

    console.log("Строка заголовков найдена на позиции:", headerRowIndex, originalHeaders);

    let startDataRow = headerRowIndex + 1;
    processedDataset = [];

    for (let i = startDataRow; i < rawExcelRows.length; i++) {
        const row = rawExcelRows[i];
        if (!row || row.length === 0) continue;

        // Создаем карту строка -> значение
        const rowMap = {};
        originalHeaders.forEach((header, colIdx) => {
            rowMap[header] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });

        const title1 = rowMap[t1HeaderName] || '';
        // Пропускаем пустые строки или технические маркеры Директовских шаблонов
        if (!title1 || title1 === '-' || title1.startsWith('---')) continue; 

        const title2 = rowMap[t2HeaderName] || '';
        const analyzedRow = computeRowMetrics(i, title1, title2, rowMap);
        processedDataset.push(analyzedRow);
    }

    console.log("Обработка завершена. Валидных строк для таблицы:", processedDataset.length);

    updateDashboardStats();
    buildTableHeader();
    renderFullTable();
}

function computeRowMetrics(rowIndex, t1, t2, rowMap) {
    const cleanTitle2 = (t2 === '-' || t2 === '0' || t2 === 'none') ? '' : t2;
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
        if (utpAnalysis.lost) { 
            statusType = 'lost-utp'; 
            statusWeight = 3; 
        } else { 
            statusType = 'lost-safe'; 
            statusWeight = 2; 
        }
    }

    return { 
        rowIndex, 
        t1, 
        t2: cleanTitle2, 
        combined: combinedTitle, 
        isMerged, 
        length: totalLength, 
        overflow, 
        statusType, 
        statusWeight, 
        utpReasons: utpAnalysis.reasons, 
        rowMap 
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

// --- 6. Рендеринг интерфейса ---
function updateDashboardStats() {
    const total = processedDataset.length;
    const success = processedDataset.filter(d => d.statusType === 'success').length;
    const cut = processedDataset.filter(d => d.statusType === 'lost-safe').length;
    const loss = processedDataset.filter(d => d.statusType === 'lost-utp').length;
    const pct = total > 0 ? Math.round((success / total) * 100) : 0;

    const sTotal = document.getElementById('statTotal');
    const sSuccess = document.getElementById('statSuccess');
    const sSuccessPct = document.getElementById('statSuccessPct');
    const sCut = document.getElementById('statCut');
    const sLoss = document.getElementById('statLoss');

    if (sTotal) sTotal.innerText = total;
    if (sSuccess) sSuccess.innerText = success;
    if (sSuccessPct) sSuccessPct.innerText = `${pct}% от всех объявлений`;
    if (sCut) sCut.innerText = cut;
    if (sLoss) sLoss.innerText = loss;
}

function buildTableHeader() {
    if (!tableHeaderRow) return;
    
    let html = `
        <th id="sortStatusBtn" class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-0 z-30 border-r border-slate-200 shadow-sm min-w-[150px] cursor-pointer hover:bg-slate-200 transition-colors select-none">
            <div class="flex items-center gap-1.5">Статус переноса <span id="sortIndicator" class="text-indigo-600 font-mono text-xs">↕</span></div>
        </th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 sticky left-[150px] z-30 border-r border-slate-200 min-w-[220px]">Итоговый заголовок (подстановка)</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 text-center min-w-[110px]">Длина (ост.)</th>
        <th class="py-4 px-4 bg-slate-100 text-slate-700 min-w-[150px]">Потерянные УТП</th>
    `;
    
    originalHeaders.forEach(header => {
        html += `<th class="py-4 px-4 border-b border-slate-200 text-slate-600 font-medium bg-slate-50 text-left min-w-[120px]">${header || ''}</th>`;
    });
    
    tableHeaderRow.innerHTML = html;
    
    const sortBtn = document.getElementById('sortStatusBtn');
    if (sortBtn) sortBtn.addEventListener('click', toggleSort);
}

function toggleSort() {
    sortDirection = sortDirection === 'none' ? 'asc' : (sortDirection === 'asc' ? 'desc' : 'none');
    const indicator = document.getElementById('sortIndicator');
    if (indicator) indicator.innerText = sortDirection === 'none' ? '↕' : (sortDirection === 'asc' ? '↑' : '↓');
    renderFullTable();
}

function formatTemplateText(text) {
    if (!text || !text.includes('#')) return text;
    return text.replace(/#([^#]+)#/g, '<span class="px-1 bg-amber-100 text-amber-800 rounded font-mono text-xs">#$1#</span>');
}

function renderFullTable() {
    if (!tableBody) return;
    tableBody.innerHTML = '';
    
    let displayData = [...processedDataset];
    
    // Фильтрация по табам
    if (currentFilter !== 'all') {
        displayData = displayData.filter(item => item.statusType === currentFilter);
    }
    
    // Поиск
    if (searchQuery) {
        displayData = displayData.filter(item => 
            item.t1.toLowerCase().includes(searchQuery) || 
            item.t2.toLowerCase().includes(searchQuery) || 
            item.combined.toLowerCase().includes(searchQuery)
        );
    }
    
    // Сортировка
    if (sortDirection === 'asc') displayData.sort((a, b) => a.statusWeight - b.statusWeight);
    else if (sortDirection === 'desc') displayData.sort((a, b) => b.statusWeight - a.statusWeight);

    const counterEl = document.getElementById('tableCounter');
    if (counterEl) counterEl.innerText = `Отображено строк: ${displayData.length}`;

    if (displayData.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="${originalHeaders.length + 4}" class="p-8 text-center text-slate-400">Нет данных для отображения</td></tr>`;
        return;
    }

    displayData.forEach(item => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/80 transition-colors border-b border-slate-100 align-top text-sm";
        tr.setAttribute('data-row-index', item.rowIndex);

        // Настройка бейджей статуса
        let statusBadge = '';
        if (item.statusType === 'success') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check-circle" class="w-3.5 h-3.5"></i> Склеен полностью</span>`;
        } else if (item.statusType === 'lost-safe') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="alert-triangle" class="w-3.5 h-3.5"></i> Не влез (без УТП)</span>`;
        } else {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-medium bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="x-circle" class="w-3.5 h-3.5"></i> Потеряно УТП!</span>`;
        }

        // Логика цвета счетчика длины
        const remaining = 56 - item.length;
        const lengthBadge = remaining >= 0 
            ? `<span class="text-emerald-600 font-semibold">${item.length} <span class="text-slate-400 font-normal">(${remaining})</span></span>`
            : `<span class="text-rose-600 font-bold">${item.length} <span class="bg-rose-100 px-1 rounded">строка > 56!</span></span>`;

        // Проблемы / Причины потери УТП
        const problems = item.utpReasons.length > 0 
            ? item.utpReasons.map(r => `<span class="inline-block bg-rose-100 text-rose-800 text-xs px-1.5 py-0.5 rounded mr-1 mb-1 font-medium">${r}</span>`).join('')
            : '<span class="text-slate-400 text-xs">—</span>';

        // Формируем статичные первые 4 колонки
        let rowHtml = `
            <td class="p-3 bg-white sticky left-0 z-20 border-r border-slate-200 shadow-sm">${statusBadge}</td>
            <td class="p-3 bg-white sticky left-[150px] z-20 border-r border-slate-200 font-medium text-slate-800 max-w-xs truncate">${formatTemplateText(item.combined)}</td>
            <td class="p-3 text-center border-r border-slate-100">${lengthBadge}</td>
            <td class="p-3 border-r border-slate-200">${problems}</td>
        `;

        // Динамически рендерим все оригинальные столбцы из файла
        originalHeaders.forEach(header => {
            const val = item.rowMap[header] || '';
            const isT1 = header === t1HeaderName;
            const isT2 = header === t2HeaderName;
            
            // Если это Заголовок 1 или 2, делаем их инлайн-редактируемыми
            if (isT1 || isT2) {
                rowHtml += `
                    <td class="p-3 border-r border-slate-100 min-w-[200px]">
                        <div class="editable-cell px-2 py-1 rounded border border-transparent hover:border-slate-300 hover:bg-white focus:bg-white focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 outline-none transition-all" 
                             contenteditable="true" 
                             data-field="${header}" 
                             data-origin="${val}">${formatTemplateText(val)}</div>
                    </td>`;
            } else {
                rowHtml += `<td class="p-3 border-r border-slate-100 text-slate-500 max-w-sm truncate" title="${val}">${formatTemplateText(val)}</td>`;
            }
        });

        tr.innerHTML = rowHtml;
        tableBody.appendChild(tr);
    });

    initInlineEditingEvents();
}

// --- 7. Инлайн-редактирование и обновление ---
function initInlineEditingEvents() {
    document.querySelectorAll('.editable-cell').forEach(cell => {
        // Убираем HTML-теги форматирования при фокусе, чтобы редактировать чистый текст
        cell.addEventListener('focus', (e) => {
            const tr = e.target.closest('tr');
            const origIndex = parseInt(tr.getAttribute('data-row-index'));
            const field = e.target.getAttribute('data-field');
            e.target.innerText = processedDataset.find(d => d.rowIndex === origIndex).rowMap[field];
        });

        cell.addEventListener('blur', (e) => {
            const tr = e.target.closest('tr');
            if (!tr) return;
            
            const origIndex = parseInt(tr.getAttribute('data-row-index'));
            const field = e.target.getAttribute('data-field');
            const newValue = e.target.innerText.trim();

            // Находим строку в массиве данных
            const dataItem = processedDataset.find(d => d.rowIndex === origIndex);
            if (!dataItem) return;

            // Если значение изменилось, пересчитываем строку
            if (dataItem.rowMap[field] !== newValue) {
                dataItem.rowMap[field] = newValue;
                
                // Обновляем исходный массив rawExcelRows для последующей выгрузки
                const headerColIdx = originalHeaders.indexOf(field);
                if (headerColIdx !== -1) {
                    rawExcelRows[origIndex][headerColIdx] = newValue;
                }

                // Пересчитываем метрики конкретно этой строки
                const updatedMetrics = computeRowMetrics(
                    origIndex, 
                    dataItem.rowMap[t1HeaderName] || '', 
                    dataItem.rowMap[t2HeaderName] || '', 
                    dataItem.rowMap
                );

                // Вносим изменения обратно в массив обработанных данных
                Object.assign(dataItem, updatedMetrics);

                // Точечно перерисовываем UI без сброса фокуса всей таблицы
                updateDashboardStats();
                
                // Для стабильности вызываем рендер (оптимизированный сборщик подхватит изменения)
                renderFullTable();
            } else {
                // Если ничего не менялось — возвращаем красивую подсветку шаблонов обратно
                e.target.innerHTML = formatTemplateText(newValue);
            }
        });

        // Завершение редактирования по Enter
        cell.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                e.target.blur();
            }
        });
    });

    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
}

// --- 8. Выгрузка измененного файла ---
function downloadUpdatedXLSX() {
    if (rawExcelRows.length === 0) {
        alert("Нет данных для выгрузки.");
        return;
    }

    try {
        const ws = XLSX.utils.aoa_to_sheet(rawExcelRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Редактированная Кампания");
        
        // Генерация имени файла
        const date = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `compiled_campaign_${date}.xlsx`);
    } catch (err) {
        console.error("Ошибка при генерации файла:", err);
        alert("Не удалось сгенерировать XLSX файл.");
    }
}
