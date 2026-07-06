/**
 * ФИНАЛЬНЫЙ СБОРНЫЙ СКРИПТ УТИЛИТЫ ВАЛИДАЦИИ И СКЛЕЙКИ ОБЪЯВЛЕНИЙ
 * 
 * Что учтено:
 * 1. Соединение Заголовка 1 и 2 строго через ТОЧКУ И ПРОБЕЛ.
 * 2. Вырезание решеток (#) при расчете лимитов Яндекса (56 символов).
 * 3. Поддержка не-ТГО объявлений (Комбинаторных и др.) — выводятся без склейки и валидации.
 * 4. Мгновенный пересчет счетчиков длин и превышения (+X) в реальном времени при вводе букв.
 * 5. Полноценное сохранение состояния в IndexedDB, экспорт и фильтрация.
 */

// ==========================================
// 1. ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ И НАСТРОЙКИ
// ==========================================
let rawExcelRows = [];        // Исходный двумерный массив из Excel файла
let processedDataset = [];    // Наш обработанный массив агрегированных объявлений
let originalHeaders = [];     // Заголовки столбцов из исходного файла
let headerRowGlobalIndex = -1;// Индекс строки с заголовками

let currentFilter = 'all';     // Текущий фильтр вкладок: all, no-t2, lost-utp, lost-safe, ok
let searchQuery = '';         // Строка поиска
let sortDirection = 'none';    // Направление сортировки: none, asc, desc

// Константы имен колонок для логики утилиты
const t1HeaderName = 'Заголовок 1';
const t2HeaderName = 'Заголовок 2';
const textHeaderName = 'Текст';
const idHeaderName = 'ID объявления';

// Ссылка на БД IndexedDB
let db = null;
const DB_NAME = 'AdValidatorDB';
const STORE_NAME = 'app_state';

// ==========================================
// 2. ИНИЦИАЛИЗАЦИЯ И ИНТЕГРАЦИЯ С INDEXEDDB
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    initIndexedDB(() => {
        loadStateFromDB();
    });
    setupAppEventListeners();
});

function initIndexedDB(callback) {
    const request = indexedDB.open(DB_NAME, 1);
    
    request.onupgradeneeded = function(e) {
        const dbRef = e.target.result;
        if (!dbRef.objectStoreNames.contains(STORE_NAME)) {
            dbRef.createObjectStore(STORE_NAME, { keyPath: 'id' });
        }
    };
    
    request.onsuccess = function(e) {
        db = e.target.result;
        if (callback) callback();
    };
    
    request.onerror = function(e) {
        console.error('Ошибка инициализации IndexedDB:', e.target.error);
        if (callback) callback();
    };
}

function saveStateToDB() {
    if (!db) return;
    const transaction = db.transaction([STORE_NAME], 'readwrite');
    const store = transaction.objectStore(STORE_NAME);
    
    const state = {
        id: 'current_state',
        rawExcelRows: rawExcelRows,
        processedDataset: processedDataset,
        originalHeaders: originalHeaders,
        headerRowGlobalIndex: headerRowGlobalIndex,
        currentFilter: currentFilter,
        searchQuery: searchQuery,
        sortDirection: sortDirection
    };
    
    store.put(state);
}

function loadStateFromDB() {
    if (!db) return;
    const transaction = db.transaction([STORE_NAME], 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const request = store.get('current_state');
    
    request.onsuccess = function(e) {
        const state = e.target.result;
        if (state) {
            rawExcelRows = state.rawExcelRows || [];
            processedDataset = state.processedDataset || [];
            originalHeaders = state.originalHeaders || [];
            headerRowGlobalIndex = state.headerRowGlobalIndex !== undefined ? state.headerRowGlobalIndex : -1;
            currentFilter = state.currentFilter || 'all';
            searchQuery = state.searchQuery || '';
            sortDirection = state.sortDirection || 'none';
            
            // Восстанавливаем интерфейс, если данные были
            if (rawExcelRows.length > 0 && processedDataset.length > 0) {
                // Синхронизируем поисковую строку, если она была сохранена
                const searchInput = document.getElementById('tableSearch');
                if (searchInput) searchInput.value = searchQuery;
                
                updateDashboardStats();
                buildTableHeader();
                renderFullTable();
                updateFilterTabStyles();
            }
        }
    };
}

function clearAppState() {
    rawExcelRows = [];
    processedDataset = [];
    originalHeaders = [];
    headerRowGlobalIndex = -1;
    currentFilter = 'all';
    searchQuery = '';
    sortDirection = 'none';
    
    const searchInput = document.getElementById('tableSearch');
    if (searchInput) searchInput.value = '';
    
    if (db) {
        const transaction = db.transaction([STORE_NAME], 'readwrite');
        transaction.objectStore(STORE_NAME).clear();
    }
    
    // Очищаем DOM элементы
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    if (tableHead) tableHead.innerHTML = '';
    if (tableBody) tableBody.innerHTML = '<tr><td class="p-8 text-center text-slate-400">Файл не загружен. Перетащите его сюда.</td></tr>';
    
    updateDashboardStats();
    updateFilterTabStyles();
}

// ==========================================
// 3. НАСТРОЙКА НАТИВНЫХ ИВЕНТОВ СТРАНИЦЫ
// ==========================================
function setupAppEventListeners() {
    const fileInput = document.getElementById('excelFileInput');
    const dropZone = document.getElementById('dropZone');
    const clearBtn = document.getElementById('clearStateBtn');
    const searchInput = document.getElementById('tableSearch');
    const exportBtn = document.getElementById('exportExcelBtn');
    
    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) handleIncomingFile(file);
        });
    }
    
    if (dropZone) {
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('border-indigo-500', 'bg-indigo-50/20'); });
        dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('border-indigo-500', 'bg-indigo-50/20'); });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('border-indigo-500', 'bg-indigo-50/20');
            const file = e.dataTransfer.files[0];
            if (file) handleIncomingFile(file);
        });
    }
    
    if (clearBtn) clearBtn.addEventListener('click', clearAppState);
    
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            searchQuery = e.target.value.trim().toLowerCase();
            renderFullTable();
            saveStateToDB();
        });
    }
    
    if (exportBtn) exportBtn.addEventListener('click', exportToExcelFile);
    
    // Навешиваем клики на табы фильтров
    document.querySelectorAll('[data-filter-tab]').forEach(tab => {
        tab.addEventListener('click', function() {
            currentFilter = this.getAttribute('data-filter-tab');
            updateFilterTabStyles();
            renderFullTable();
            saveStateToDB();
        });
    });
}

function updateFilterTabStyles() {
    document.querySelectorAll('[data-filter-tab]').forEach(tab => {
        const filterType = tab.getAttribute('data-filter-tab');
        if (filterType === currentFilter) {
            tab.className = "px-4 py-2 text-sm font-medium border-b-2 border-indigo-600 text-indigo-600 focus:outline-none";
        } else {
            tab.className = "px-4 py-2 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300 focus:outline-none";
        }
    });
}

// Пакетный разбор файла библиотекой XLSX
function handleIncomingFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Превращаем в массив строк-массивов
            rawExcelRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (!rawExcelRows || rawExcelRows.length === 0) {
                alert("Кажется, файл пустой.");
                return;
            }
            
            analyzeStructureAndProcess();
        } catch (err) {
            console.error(err);
            alert("Ошибка при чтении Excel файла. Убедитесь, что это валидный .xlsx");
        }
    };
    reader.readAsArrayBuffer(file);
}

// ==========================================
// 4. ЯДРО УТИЛИТЫ: ОБРАБОТКА ДАННЫХ И МЕТРИКИ
// ==========================================
function analyzeStructureAndProcess() {
    let headerRowIndex = -1;

    // Ищем строку заголовков без привязки к регистру букв
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

    // Пропускаем техническую строку длин, если она есть сразу под заголовками
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
            
            // Ключ склейки: ID объявления + Заголовки (сворачиваем только фразовые копии)
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

function computeRowMetrics(index, t1, t2, rowMap) {
    // Соединяем СТРОГО через точку и пробел
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

// ==========================================
// 5. ОТРЕНДЕРИТЬ ИНТЕРФЕЙС ТАБЛИЦЫ И СТАТИСТИКУ
// ==========================================
function updateDashboardStats() {
    const sAll = document.getElementById('statAllCount');
    const sNoT2 = document.getElementById('statNoT2Count');
    const sLostUtp = document.getElementById('statLostUtpCount');
    const sLostSafe = document.getElementById('statLostSafeCount');
    const sOk = document.getElementById('statOkCount');
    
    if (!sAll) return; // Если элементов статистики нет в DOM
    
    const countAll = processedDataset.length;
    const countNoT2 = processedDataset.filter(d => d.statusType === 'no-t2').length;
    const countLostUtp = processedDataset.filter(d => d.statusType === 'lost-utp').length;
    const countLostSafe = processedDataset.filter(d => d.statusType === 'lost-safe').length;
    const countOk = processedDataset.filter(d => d.statusType === 'ok').length;
    
    sAll.innerText = countAll;
    sNoT2.innerText = countNoT2;
    sLostUtp.innerText = countLostUtp;
    sLostSafe.innerText = countLostSafe;
    sOk.innerText = countOk;
}

function buildTableHeader() {
    const tableHead = document.getElementById('tableHead');
    if (!tableHead) return;
    
    // Генерируем заголовки. Сначала 4 фиксированные колонки утилиты, затем оригинальные
    let headHtml = `
        <tr>
            <th class="p-3 sticky-col sticky left-0 z-20 text-left bg-slate-100 font-semibold text-slate-700 border-b border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] w-[150px]">Статус Директа</th>
            <th class="p-3 sticky-col sticky left-[150px] z-20 text-left bg-slate-100 font-semibold text-slate-700 border-b border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] w-[200px]">Превью Склейки</th>
            <th class="p-3 sticky-col sticky left-[350px] z-20 text-center bg-slate-100 font-semibold text-slate-700 border-b border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] w-[110px] cursor-pointer select-none" id="sortOverflowBtn">Превышение <span id="sortDirectionArrow">↕</span></th>
            <th class="p-3 sticky-col sticky left-[460px] z-20 text-left bg-slate-100 font-semibold text-slate-700 border-b border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.04)] w-[250px]">Проблематика</th>
    `;
    
    originalHeaders.forEach(hName => {
        headHtml += `<th class="p-3 text-left bg-slate-50 font-medium text-slate-500 border-b border-r border-slate-200 min-w-[150px] whitespace-nowrap">${hName}</th>`;
    });
    
    headHtml += `</tr>`;
    tableHead.innerHTML = headHtml;
    
    // Вешаем обработчик сортировки на колонку Превышения
    const sortBtn = document.getElementById('sortOverflowBtn');
    if (sortBtn) {
        sortBtn.addEventListener('click', () => {
            if (sortDirection === 'none') sortDirection = 'desc';
            else if (sortDirection === 'desc') sortDirection = 'asc';
            else sortDirection = 'none';
            
            const arrow = document.getElementById('sortDirectionArrow');
            if (arrow) {
                arrow.innerText = sortDirection === 'none' ? '↕' : (sortDirection === 'desc' ? '↓' : '↑');
            }
            renderFullTable();
        });
    }
}

function renderFullTable() {
    const tableBodyElement = document.getElementById('tableBody');
    if (!tableBodyElement) return;
    tableBodyElement.innerHTML = '';

    let displayData = [...processedDataset];
    
    // Фильтрация вкладок
    const currentFilterVal = typeof currentFilter !== 'undefined' ? currentFilter : 'all';
    if (currentFilterVal === 'has-t2') {
        displayData = displayData.filter(item => item.statusType !== 'no-t2' && !item.isSpecialType);
    } else if (currentFilterVal !== 'all') {
        displayData = displayData.filter(item => item.statusType === currentFilterVal);
    }

    // Поиск по ключевым словам
    const searchQueryVal = typeof searchQuery !== 'undefined' ? searchQuery : '';
    if (searchQueryVal) {
        const q = searchQueryVal.toLowerCase();
        displayData = displayData.filter(item => 
            item.t1.toLowerCase().includes(q) || 
            item.t2.toLowerCase().includes(q) ||
            item.combined.toLowerCase().includes(q)
        );
    }

    // Сортировка
    const sortDirectionVal = typeof sortDirection !== 'undefined' ? sortDirection : 'none';
    if (sortDirectionVal === 'asc') displayData.sort((a, b) => a.statusWeight - b.statusWeight);
    else if (sortDirectionVal === 'desc') displayData.sort((a, b) => b.statusWeight - a.statusWeight);

    const counterEl = document.getElementById('tableCounter');
    if (counterEl) counterEl.innerText = `Строк: ${displayData.length}`;

    if (displayData.length === 0) {
        tableBodyElement.innerHTML = `<tr><td colspan="${originalHeaders.length + 4}" class="py-12 text-center text-slate-400">Ничего не найдено.</td></tr>`;
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

        let rowHtml = `
            <td class="p-3 sticky-col sticky left-0 z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] status-cell bg-white group-hover:bg-inherit">${statusBadge}</td>
            <td class="p-3 sticky-col sticky left-[150px] z-10 border-r border-slate-200 shadow-[2px_0_5px_rgba(0,0,0,0.02)] font-medium text-slate-900 max-w-[200px] truncate combined-cell bg-white group-hover:bg-inherit" data-combined="${item.combined}">${item.isSpecialType ? item.combined : formatTemplateText(item.combined)}</td>
            <td class="p-3 sticky-col sticky left-[350px] z-10 border-r border-slate-200 text-center font-mono font-bold overflow-cell bg-white group-hover:bg-inherit ${item.overflow > 0 ? 'text-rose-600' : 'text-slate-300'}">${item.isSpecialType ? '—' : (item.overflow > 0 ? `+${item.overflow}` : '0')}</td>
            <td class="p-3 sticky-col sticky left-[460px] z-10 border-r border-slate-300 shadow-[3px_0_5px_rgba(0,0,0,0.04)] issue-cell bg-white group-hover:bg-inherit">${issueText}</td>
        `;

        originalHeaders.forEach((headerName, curIdx) => {
            let displayValue = item.rowMap[headerName] || '';
            let isT1 = (headerName === t1HeaderName);
            let isT2 = (headerName === t2HeaderName);
            
            let cellStyle = "p-3 text-slate-600 border-r border-slate-100 max-w-[250px] min-w-[150px] truncate";
            let editableAttr = "";
            let extraDataAttr = "";

            if (lenIndices.includes(curIdx) && !item.isSpecialType) {
                cellStyle = "p-3 font-mono font-semibold text-center bg-slate-50/50 text-indigo-600 border-r border-slate-100 min-w-[70px]";
                if (curIdx === lenIndices[0]) { displayValue = item.t1.length; extraDataAttr = `data-len-type="t1"`; }
                else if (curIdx === lenIndices[1]) { displayValue = item.t2.length; extraDataAttr = `data-len-type="text2"`; }
                else if (curIdx === lenIndices[2]) { displayValue = (item.rowMap[textHeaderName] || '').length; extraDataAttr = `data-len-type="text"`; }
            }
            
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
        tableBodyElement.appendChild(tr);
    });

    if (typeof lucide !== 'undefined') lucide.createIcons();
    initInlineEditingEvents(); 
}

// ==========================================
// 6. ЖИВОЙ ПЕРЕСЧЕТ В РЕАЛЬНОМ ВРЕМЕНИ (INPUT)
// ==========================================
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

            // Обновляем данные в памяти приложения
            item.t1 = newT1;
            item.t2 = newT2;
            item.rowMap[t1HeaderName] = newT1;
            item.rowMap[t2HeaderName] = newT2;
            
            // Соединяем строго через точку и пробел
            item.combined = newT2 ? `${newT1}. ${newT2}` : newT1;

            // Вычисляем длину без решеток
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
            
            // Точечно меняем бейджи и текст проблемы, не сбивая курсор ввода
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
            
            // Динамически обновляем дашборд и сохраняем изменения в локальную БД
            updateDashboardStats();
            saveStateToDB(); 
        });
    });
}

function formatTemplateText(text) {
    if (!text) return '';
    return text.replace(/(#[^#\s]+#)/g, '<span class="px-1 py-0.5 rounded bg-indigo-100 text-indigo-800 font-mono text-xs border border-indigo-200">$1</span>');
}

// ==========================================
// 7. РАЗГРУЗКА И СБОРКА ИТОГОВОГО XLSX ФАЙЛА
// ==========================================
function exportToExcelFile() {
    if (processedDataset.length === 0) {
        alert("Нет данных для экспорта.");
        return;
    }
    
    // Создаем массив строк для нового Excel. Первая строка — шапка файла
    const exportRows = [];
    exportRows.push(originalHeaders);
    
    // Бежим по нашему агрегированному датасету
    processedDataset.forEach(item => {
        // Если объявление текстово-графическое и склеило в себе несколько строк из исходника
        if (!item.isSpecialType && item.realRowIndices && item.realRowIndices.length > 0) {
            item.realRowIndices.forEach(origIdx => {
                const originalRow = [...rawExcelRows[origIdx]];
                
                // Переносим отредактированные Т1 и Т2 обратно в массив строки перед выгрузкой
                const t1Idx = originalHeaders.indexOf(t1HeaderName);
                const t2Idx = originalHeaders.indexOf(t2HeaderName);
                
                if (t1Idx !== -1) originalRow[t1Idx] = item.t1;
                if (t2Idx !== -1) originalRow[t2Idx] = item.t2;
                
                exportRows.push(originalRow);
            });
        } else {
            // Для комбинаторных и иных типов объявлений выгружаем строку как есть
            const originalRow = [];
            originalHeaders.forEach(header => {
                originalRow.push(item.rowMap[header] || '');
            });
            exportRows.push(originalRow);
        }
    });
    
    // Формируем XLSX книгу
    const newWorksheet = XLSX.utils.aoa_to_sheet(exportRows);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Валидированные объявления");
    
    // Скачиваем файл на компьютер
    XLSX.writeFile(newWorkbook, `объявления_склейка_${new Date().toISOString().slice(0,10)}.xlsx`);
}
