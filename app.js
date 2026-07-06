// Глобальные переменные (убедитесь, что они объявлены в вашем коде)
let rawExcelRows = []; // Исходные данные из xlsx
let processedDataset = []; // Наш обработанный массив для таблицы
let originalHeaders = []; // Массив заголовков исходного файла
let headerRowGlobalIndex = -1;
let currentFilter = 'all';
let searchQuery = '';
let sortDirection = 'none';

// Константы имен колонок (базовые)
const t1HeaderName = 'Заголовок 1';
const t2HeaderName = 'Заголовок 2';
const textHeaderName = 'Текст';
const idHeaderName = 'ID объявления';

/**
 * 1. ОСНОВНОЙ АНАЛИЗ СТРУКТУРЫ И ГРУППИРОВКА СТРОК
 */
function analyzeStructureAndProcess() {
    let headerRowIndex = -1;

    // Ищем строку заголовков (без привязки к регистру букв)
    for (let i = 0; i < Math.min(rawExcelRows.length, 30); i++) {
        if (!rawExcelRows[i]) continue;
        const rowStr = rawExcelRows[i].map(cell => String(cell || '').trim());
        
        const foundT1 = rowStr.indexOf(t1HeaderName); 
        const hasIdCol = rowStr.some(c => {
            const low = c.toLowerCase();
            return low.includes('id объявления') || low.includes('id группы') || low === 'id';
        });

        if (foundT1 !== -1 && hasIdCol) {
            headerRowIndex = i;
            headerRowGlobalIndex = i;
            originalHeaders = rowStr;
            break;
        }
    }

    if (headerRowIndex === -1) {
        alert("Не удалось найти строку заголовков с полем 'Заголовок 1' и колонкой ID.");
        return;
    }

    // Находим точные имена ключевых колонок в текущем файле
    const actualIdHeader = originalHeaders.find(c => {
        const low = c.toLowerCase();
        return low.includes('id объявления') || low.includes('id группы') || low === 'id';
    }) || idHeaderName;

    const typeHeaderName = originalHeaders.find(c => c.toLowerCase().includes('тип объявления')) || 'Тип объявления';

    // Определяем начало данных (пропускаем техническую строку длин, если она есть)
    let startDataRow = headerRowIndex + 1;
    if (startDataRow < rawExcelRows.length && rawExcelRows[startDataRow]) {
        const checkRow = rawExcelRows[startDataRow].map(c => String(c || '').toLowerCase().trim());
        if (checkRow.includes('заголовок 1') || checkRow.includes('текст') || checkRow.some(c => c === '55' || c === '35')) {
            startDataRow++;
        }
    }

    const groupedData = {};
    let virtualIndex = 0;

    for (let i = startDataRow; i < rawExcelRows.length; i++) {
        const row = rawExcelRows[i];
        if (!row || row.length === 0) continue;

        const rowMap = {};
        originalHeaders.forEach((header, colIdx) => {
            rowMap[header] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });

        const adType = (rowMap[typeHeaderName] || '').toLowerCase().trim();
        const title1 = rowMap[t1HeaderName] || '';
        const title2 = rowMap[t2HeaderName] || '';
        const idValue = rowMap[actualIdHeader] || `no-id-${virtualIndex}`;

        // Текстово-графическим считаем тип ТГО или пустую ячейку
        const isTextGraphic = adType.includes('текстово-графическое') || adType === '';

        if (!isTextGraphic) {
            // Спец-типы (Комбинаторные и др.): выводим «как есть» без склейки
            let previewTitle = title1;
            if (!previewTitle) {
                previewTitle = rowMap['Заголовок 1 (Комбинаторика)'] || rowMap['Заголовок 3'] || 'Комбинаторное объявление';
            }

            const uniqueTypeKey = `special_${i}_${idValue}`;
            groupedData[uniqueTypeKey] = {
                realRowIndices: [i],
                title1: title1,
                title2: title2,
                combinedTitle: previewTitle,
                rowMap: rowMap,
                isSpecialType: true,
                displayType: rowMap[typeHeaderName] || 'Другое'
            };
        } else {
            // Текстово-графические: фильтруем пустые/технические строки
            if (!title1 || title1 === '-' || title1.startsWith('---')) continue; 
            
            // Ключ склейки: ID + Заголовки (склеиваем только фразовые дубли)
            const uniqueGroupKey = `${idValue}_[T1:${title1}]_[T2:${title2}]`;

            if (!groupedData[uniqueGroupKey]) {
                groupedData[uniqueGroupKey] = {
                    realRowIndices: [i], 
                    title1: title1,
                    title2: title2,
                    rowMap: rowMap,
                    isSpecialType: false
                };
            } else {
                groupedData[uniqueGroupKey].realRowIndices.push(i);
            }
        }
        virtualIndex++;
    }

    processedDataset = [];

    Object.values(groupedData).forEach((group, index) => {
        let analyzedRow;
        if (group.isSpecialType) {
            analyzedRow = {
                rowIndex: index,
                t1: group.title1,
                t2: group.title2,
                combined: group.combinedTitle,
                isMerged: false,
                length: 0,
                overflow: 0,
                statusType: 'special-type',
                statusWeight: 0,
                utpReasons: [],
                rowMap: group.rowMap,
                isSpecialType: true,
                displayType: group.displayType
            };
        } else {
            analyzedRow = computeRowMetrics(index, group.title1, group.title2, group.rowMap);
            analyzedRow.isSpecialType = false;
        }
        analyzedRow.realRowIndices = group.realRowIndices; 
        processedDataset.push(analyzedRow);
    });

    updateDashboardStats();
    buildTableHeader();
    renderFullTable();
    saveStateToDB();
}

/**
 * 2. МЕТРИКИ СТРОКИ И ПРОВЕРКА ПРАВИЛ ДИРЕКТА
 */
function computeRowMetrics(index, t1, t2, rowMap) {
    // Соединяем строго через точку и пробел
    const combined = t2 ? `${t1}. ${t2}` : t1;
    
    // Чистая длина для лимитов Яндекса (без учета решеток # шаблона)
    const cleanLength = combined.replace(/#/g, '').length;
    const overflow = Math.max(0, cleanLength - 56); 

    let statusType = 'ok';
    let statusWeight = 1;
    let utpReasons = [];

    // Поиск УТП триггеров в доп. заголовке (проценты, валюты, скидки)
    const hasUtpTrigger = /[\d%]|руб|коп|usd|eur|¥|штук|от|до|—|-|–|бесплатно|акция|скид|гарант|достав/i.test(t2);

    if (!t2) {
        statusType = 'no-t2';
        statusWeight = 2;
    } else if (cleanLength > 56) {
        if (hasUtpTrigger) {
            statusType = 'lost-utp';
            statusWeight = 4;
            // Короткое описание почему отброшен
            if (/[\d%]/.test(t2)) utpReasons.push("цифры/скидки");
            if (/руб|usd|eur|¥/i.test(t2)) utpReasons.push("валюта");
            if (/от|до/i.test(t2)) utpReasons.push("порог цены");
            if (utpReasons.length === 0) utpReasons.push("важный текст");
        } else {
            statusType = 'lost-safe';
            statusWeight = 3;
        }
    }

    return {
        rowIndex: index,
        t1: t1,
        t2: t2,
        combined: combined,
        length: cleanLength,
        overflow: overflow,
        statusType: statusType,
        statusWeight: statusWeight,
        utpReasons: utpReasons,
        rowMap: rowMap
    };
}

/**
 * 3. ОТРИСОВКА ТАБЛИЦЫ В ИНТЕРФЕЙСЕ
 */
function renderFullTable() {
    const tableBody = document.getElementById('tableBody');
    if (!tableBody) return;
    tableBody.innerHTML = '';

    let displayData = [...processedDataset];
    
    // Фильтрация по вкладкам
    if (currentFilter === 'has-t2') {
        displayData = displayData.filter(item => item.statusType !== 'no-t2' && !item.isSpecialType);
    } else if (currentFilter !== 'all') {
        displayData = displayData.filter(item => item.statusType === currentFilter);
    }

    // Фильтрация по поиску
    if (searchQuery) {
        const q = searchQuery.toLowerCase();
        displayData = displayData.filter(item => 
            item.t1.toLowerCase().includes(q) || 
            item.t2.toLowerCase().includes(q) ||
            item.combined.toLowerCase().includes(q)
        );
    }

    // Сортировка
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
        let statusBadge = '';
        let rowBgClass = 'hover:bg-slate-50/80';
        let issueText = '—';

        if (item.isSpecialType) {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-purple-50 text-purple-700 border border-purple-200"><i data-lucide="layers" class="w-3 h-3"></i> ${item.displayType}</span>`;
            rowBgClass = 'bg-purple-50/5 hover:bg-purple-50/10';
            issueText = '<span class="text-slate-400 text-xs">Не валидируется</span>';
        } else if (item.statusType === 'no-t2') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="minus-circle" class="w-3 h-3"></i> Нет доп. заголовка</span>`;
        } else if (item.statusType === 'lost-utp') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="alert-circle" class="w-3 h-3"></i> Доп. заголовок отброшен</span>`;
            rowBgClass = 'bg-rose-50/10 hover:bg-rose-50/20';
            issueText = `<span class="text-rose-600 font-semibold text-xs">Теряет УТП: ${item.utpReasons.join(', ')}</span>`;
        } else if (item.statusType === 'lost-safe') {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="scissors" class="w-3 h-3"></i> Срез доп. заг-ка</span>`;
            rowBgClass = 'bg-amber-50/5 hover:bg-amber-50/15';
        } else {
            statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Перенесется</span>`;
        }

        tr.className = `${rowBgClass} transition-colors group`;
        tr.setAttribute('data-row-index', item.rowIndex);

        // Фиксированные системные колонки утилиты
        let rowHtml = `
            <td class="p-3 sticky-col sticky left-0 z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] status-cell bg-white group-hover:bg-inherit">${statusBadge}</td>
            <td class="p-3 sticky-col sticky left-[150px] z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] font-medium text-slate-900 max-w-[200px] truncate combined-cell bg-white group-hover:bg-inherit" data-combined="${item.combined}">${item.isSpecialType ? item.combined : formatTemplateText(item.combined)}</td>
            <td class="p-3 sticky-col sticky left-[350px] z-10 border-r border-slate-200 text-center font-mono font-bold overflow-cell bg-white group-hover:bg-inherit ${item.overflow > 0 ? 'text-rose-600' : 'text-slate-300'}">${item.isSpecialType ? '—' : (item.overflow > 0 ? `+${item.overflow}` : '0')}</td>
            <td class="p-3 sticky-col sticky left-[460px] z-10 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.04)] issue-cell bg-white group-hover:bg-inherit">${issueText}</td>
        `;

        // Отрендерить оригинальные колонки из файла (включая Комбинаторику)
        originalHeaders.forEach((headerName, curIdx) => {
            let displayValue = item.rowMap[headerName] || '';
            let isT1 = (headerName === t1HeaderName);
            let isT2 = (headerName === t2HeaderName);
            
            let cellStyle = "p-3 text-slate-600 border-r border-slate-100 max-w-[250px] min-w-[150px] truncate";
            let editableAttr = "";
            let extraDataAttr = "";

            // Технические ячейки счетчиков длин (если они сгенерированы парсером)
            if (lenIndices.includes(curIdx) && !item.isSpecialType) {
                cellStyle = "p-3 font-mono font-semibold text-center bg-slate-50/50 text-indigo-600 border-r border-slate-100 min-w-[70px]";
                if (curIdx === lenIndices[0]) { displayValue = item.t1.length; extraDataAttr = `data-len-type="t1"`; }
                else if (curIdx === lenIndices[1]) { displayValue = item.t2.length; extraDataAttr = `data-len-type="text2"`; }
                else if (curIdx === lenIndices[2]) { displayValue = (item.rowMap[textHeaderName] || '').length; extraDataAttr = `data-len-type="text"`; }
            }
            
            // Разрешаем инлайн-редактирование Т1 и Т2 только для ТГО
            if (!item.isSpecialType) {
                if (isT1) {
                    cellStyle = "p-3 bg-indigo-50/40 text-slate-900 font-medium editable-cell cursor-text border-r border-slate-100 min-w-[250px] max-w-[350px] whitespace-pre-wrap break-all";
                    editableAttr = `contenteditable="true" data-type="t1"`;
                }
                if (isT2) {
                    cellStyle = "p-3 bg-amber-50/30 text-slate-900 font-medium editable-cell cursor-text border-r border-slate-100 min-w-[250px] max-w-[350px] whitespace-pre-wrap break-all";
                    editableAttr = `contenteditable="true" data-type="t2"`;
                }
            }

            const valueToRender = (isT1 || isT2) && !item.isSpecialType ? formatTemplateText(String(displayValue)) : displayValue;
            rowHtml += `<td class="${cellStyle}" ${editableAttr} ${extraDataAttr} title="${displayValue}">${valueToRender}</td>`;
        });

        tr.innerHTML = rowHtml;
        tableBody.appendChild(tr);
    });

    if (typeof lucide !== 'undefined') lucide.createIcons();
    initInlineEditingEvents(); 
}

/**
 * 4. ПЕРЕСЧЕТ ПРЕВЫШЕНИЯ И ДЛИНЫ В РЕАЛЬНОМ ВРЕМЕНИ (ЖИВОЙ ВВОД)
 */
function initInlineEditingEvents() {
    const table = document.getElementById('mainTable');
    if (!table) return;
    
    table.querySelectorAll('.editable-cell').forEach(cell => {
        if (cell.dataset.editingInitialized) return;
        cell.dataset.editingInitialized = "true";

        cell.addEventListener('input', function() {
            const tr = this.closest('tr');
            const rowIndex = parseInt(tr.getAttribute('data-row-index'));
            const item = processedDataset.find(d => d.rowIndex === rowIndex);
            
            if (!item || item.isSpecialType) return;

            const t1Cell = tr.querySelector('[data-type="t1"]');
            const t2Cell = tr.querySelector('[data-type="t2"]');
            
            const newT1 = t1Cell ? t1Cell.innerText.trim() : '';
            const newT2 = t2Cell ? t2Cell.innerText.trim() : '';

            // Обновляем данные в оперативной памяти утилиты
            item.t1 = newT1;
            item.t2 = newT2;
            item.rowMap[t1HeaderName] = newT1;
            item.rowMap[t2HeaderName] = newT2;
            
            // Соединяем СТРОГО через точку и пробел
            item.combined = newT2 ? `${newT1}. ${newT2}` : newT1;

            // Вычисляем чистую длину БЕЗ знаков # (шаблонов Яндекса)
            const cleanLength = item.combined.replace(/#/g, '').length;
            item.overflow = Math.max(0, cleanLength - 56);

            // Мгновенно обновляем счетчики длин в строке интерфейса
            const t1LenCell = tr.querySelector('[data-len-type="t1"]');
            if (t1LenCell) t1LenCell.innerText = newT1.length;

            const t2LenCell = tr.querySelector('[data-len-type="text2"]');
            if (t2LenCell) t2LenCell.innerText = newT2.length;

            // Обновляем ячейку текстового превью
            const combinedCell = tr.querySelector('.combined-cell');
            if (combinedCell) {
                combinedCell.setAttribute('data-combined', item.combined);
                combinedCell.innerHTML = formatTemplateText(item.combined); 
            }

            // Динамически перекрашиваем и меняем цифру в ячейке Превышения
            const overflowCell = tr.querySelector('.overflow-cell');
            if (overflowCell) {
                if (item.overflow > 0) {
                    overflowCell.innerText = `+${item.overflow}`;
                    overflowCell.className = "p-3 sticky-col sticky left-[350px] z-10 border-r border-slate-200 text-center font-mono font-bold overflow-cell bg-white group-hover:bg-inherit text-rose-600";
                } else {
                    overflowCell.innerText = '0';
                    overflowCell.className = "p-3 sticky-col sticky left-[350px] z-10 border-r border-slate-200 text-center font-mono font-bold overflow-cell bg-white group-hover:bg-inherit text-slate-300";
                }
            }

            // Пересчитываем статусы логики отсечения УТП Директа
            const updatedMetrics = computeRowMetrics(rowIndex, newT1, newT2, item.rowMap);
            item.statusType = updatedMetrics.statusType;
            item.utpReasons = updatedMetrics.utpReasons;
            
            // Точечно меняем бейджи и текст проблемы, не сбивая курсор ввода (focus)
            const statusCell = tr.querySelector('.status-cell');
            const issueCell = tr.querySelector('.issue-cell');
            
            if (statusCell) {
                let statusBadge = '';
                if (item.statusType === 'no-t2') {
                    statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="minus-circle" class="w-3 h-3"></i> Нет доп. заголовка</span>`;
                    tr.className = "hover:bg-slate-50/80 transition-colors group";
                } else if (item.statusType === 'lost-utp') {
                    statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-bold bg-rose-50 text-rose-700 border border-rose-200"><i data-lucide="alert-circle" class="w-3 h-3"></i> Доп. заголовок отброшен</span>`;
                    tr.className = "bg-rose-50/10 hover:bg-rose-50/20 transition-colors group";
                } else if (item.statusType === 'lost-safe') {
                    statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-amber-50 text-amber-700 border border-amber-200"><i data-lucide="scissors" class="w-3 h-3"></i> Срез доп. заг-ка</span>`;
                    tr.className = "bg-amber-50/5 hover:bg-amber-50/15 transition-colors group";
                } else {
                    statusBadge = `<span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200"><i data-lucide="check" class="w-3 h-3"></i> Перенесется</span>`;
                    tr.className = "hover:bg-slate-50/80 transition-colors group";
                }
                statusCell.innerHTML = statusBadge;
            }

            if (issueCell) {
                issueCell.innerHTML = item.statusType === 'lost-utp' 
                    ? `<span class="text-rose-600 font-semibold text-xs">Теряет УТП: ${item.utpReasons.join(', ')}</span>`
                    : '—';
            }

            if (typeof lucide !== 'undefined') lucide.createIcons({ attrs: { class: ['w-3', 'h-3'] } });
            
            saveStateToDB(); // Синхронизируем с локальным хранилищем (IndexedDB)
        });
    });
}

/**
 * Вспомогательная функция красивой подсветки шаблонов Яндекс.Директа (#текст#)
 */
function formatTemplateText(text) {
    if (!text) return '';
    return text.replace(/(#[^#\s]+#)/g, '<span class="px-1 py-0.5 rounded bg-indigo-100 text-indigo-800 font-mono text-xs border border-indigo-200">$1</span>');
}

// Заглушки для функций аналитики и БД, чтобы код не вызывал ошибок, если они определены в других файлах скриптов
if (typeof updateDashboardStats !== 'function') { function updateDashboardStats() {} }
if (typeof buildTableHeader !== 'function') { function buildTableHeader() {} }
if (typeof saveStateToDB !== 'function') { function saveStateToDB() {} }
