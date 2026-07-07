function analyzeStructureAndProcess() {
    let headerRowIndex = -1;

    // 1. Ищем строку заголовков ТГО
    for (let i = 0; i < Math.min(rawExcelRows.length, 30); i++) {
        if (!rawExcelRows[i]) continue;
        const rowStr = rawExcelRows[i].map(cell => String(cell || '').trim());
        
        const foundT1 = rowStr.indexOf(t1HeaderName); // "Заголовок 1"
        const hasIdCol = rowStr.some(c => c.includes('ID объявления') || c.includes('ID группы') || c.includes('Доп. объявление'));

        if (foundT1 !== -1 && hasIdCol) {
            headerRowIndex = i;
            headerRowGlobalIndex = i;
            originalHeaders = rowStr;
            break;
        }
    }

    if (headerRowIndex === -1) {
        // Если жесткий поиск не сработал, берем строку, где просто есть "Заголовок 1"
        for (let i = 0; i < Math.min(rawExcelRows.length, 30); i++) {
            if (!rawExcelRows[i]) continue;
            const rowStr = rawExcelRows[i].map(cell => String(cell || '').trim());
            if (rowStr.indexOf(t1HeaderName) !== -1) {
                headerRowIndex = i;
                headerRowGlobalIndex = i;
                originalHeaders = rowStr;
                break;
            }
        }
    }

    if (headerRowIndex === -1) {
        alert("Не удалось найти строку заголовков с полем 'Заголовок 1'. Проверьте формат файла.");
        return;
    }

    // Определяем точные индексы ПЕРВЫХ (базовых) полей ТГО (с заглавной буквы)
    baseT1Idx = originalHeaders.indexOf(t1HeaderName);
    baseT2Idx = originalHeaders.indexOf(t2HeaderName);
    baseTextIdx = originalHeaders.indexOf(textHeaderName);

    // Корректный пропуск служебных строк ограничений Директа (например, "56", "35", "81")
    let startDataRow = headerRowIndex + 1;
    while (startDataRow < rawExcelRows.length && rawExcelRows[startDataRow]) {
        const checkRow = rawExcelRows[startDataRow].map(c => String(c || '').toLowerCase().trim());
        // Если в строке содержатся числа ограничений или повторы названий заголовков — это служебная строка
        if (checkRow.includes('заголовок 1') || checkRow.includes('текст') || checkRow.some(c => c === '56' || c === '35' || c === '81')) {
            startDataRow++;
        } else {
            break;
        }
    }

    processedDataset = [];

    for (let i = startDataRow; i < rawExcelRows.length; i++) {
        const row = rawExcelRows[i];
        if (!row || row.length === 0) continue;

        // Сохраняем значения по изолированным ключам "НазваниеКолонки_Индекс"
        const rowMap = {};
        originalHeaders.forEach((header, colIdx) => {
            let uniqueKey = `${header || 'Пусто'}_${colIdx}`;
            rowMap[uniqueKey] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });

        let title1 = row[baseT1Idx] !== undefined ? String(row[baseT1Idx]).trim() : '';
        if (title1 === '-' || title1.startsWith('---')) title1 = '';

        let title2 = row[baseT2Idx] !== undefined ? String(row[baseT2Idx]).trim() : '';
        
        const analyzedRow = computeRowMetrics(i, title1, title2, rowMap);
        processedDataset.push(analyzedRow);
    }

    updateDashboardStats();
    buildTableHeader();
    renderFullTable();
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

    // Умный поиск индексов колонок длин базового ТГО (ищем в нижнем регистре, как в файле Коммандера)
    // Находим "заголовок 1", который идет ПОСЛЕ базового "Заголовок 1"
    let lenT1Idx = originalHeaders.findIndex((h, idx) => h.trim() === 'заголовок 1' && idx > baseT1Idx);
    let lenT2Idx = originalHeaders.findIndex((h, idx) => h.trim() === 'заголовок 2' && idx > baseT2Idx);
    let lenTextIdx = originalHeaders.findIndex((h, idx) => h.trim() === 'текст' && idx > baseTextIdx);

    // Массив базовых индексов длин для быстрой проверки
    const tgoLenIndices = [lenT1Idx, lenT2Idx, lenTextIdx].filter(idx => idx !== -1);

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
            let uniqueKey = `${headerName || 'Пусто'}_${curIdx}`;
            let displayValue = item.rowMap[uniqueKey] || '';
            let headerLower = headerName.toLowerCase().trim();
            
            let isT1 = (curIdx === baseT1Idx);
            let isT2 = (curIdx === baseT2Idx);
            
            // Проверка, является ли поле текстовым блоком из комбинаторной части кампании
            let isCombText = (curIdx > baseTextIdx) && 
                             (headerLower.includes('заголовок') || headerLower.includes('текст')) && 
                             !headerLower.includes('длина') && !headerLower.startsWith('дл.') &&
                             curIdx !== lenT1Idx && curIdx !== lenT2Idx && curIdx !== lenTextIdx;

            let isLengthCol = tgoLenIndices.includes(curIdx) || headerLower.includes('длина') || headerLower.startsWith('дл.');

            let cellStyle = "p-3 text-slate-600 border-r border-slate-100 max-w-[250px] truncate";
            let editableAttr = "";
            let extraDataAttr = "";

            if (isLengthCol) {
                cellStyle += " font-mono font-semibold text-center bg-slate-50/50 text-indigo-600";
                
                if (curIdx === lenT1Idx) {
                    displayValue = item.t1.length;
                    extraDataAttr = `data-len-type="t1"`;
                } else if (curIdx === lenT2Idx) {
                    displayValue = item.t2.length;
                    extraDataAttr = `data-len-type="text2"`;
                } else if (curIdx === lenTextIdx) {
                    let baseTextKey = `${textHeaderName}_${baseTextIdx}`;
                    displayValue = (item.rowMap[baseTextKey] || '').length;
                    extraDataAttr = `data-len-type="text"`;
                } else {
                    // Расчет длины для колонок комбинаторики
                    let targetTextIdx = originalHeaders.findIndex((h, hIdx) => {
                        if (hIdx <= baseTextIdx) return false;
                        let hLower = h.toLowerCase().trim();
                        return headerLower.includes(hLower) && hIdx !== curIdx;
                    });

                    if (targetTextIdx !== -1) {
                        let targetKey = `${originalHeaders[targetTextIdx]}_${targetTextIdx}`;
                        displayValue = (item.rowMap[targetKey] || '').length;
                        extraDataAttr = `data-len-type="comb-len" data-source-field="${originalHeaders[targetTextIdx]}_${targetTextIdx}"`;
                    } else {
                        displayValue = '0';
                    }
                }
            }
            
            if (isT1) {
                cellStyle += " bg-indigo-50/40 text-slate-900 font-medium editable-cell cursor-text";
                editableAttr = `contenteditable="true" data-type="t1" data-field-key="${uniqueKey}"`;
            } else if (isT2) {
                cellStyle += " bg-amber-50/30 text-slate-900 font-medium editable-cell cursor-text";
                editableAttr = `contenteditable="true" data-type="t2" data-field-key="${uniqueKey}"`;
            } else if (isCombText) {
                cellStyle += " bg-slate-50/60 text-slate-900 font-medium editable-cell cursor-text border-dashed border-b border-slate-300";
                editableAttr = `contenteditable="true" data-type="comb-text" data-field-key="${uniqueKey}"`;
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

function downloadUpdatedXLSX() {
    if (processedDataset.length === 0) return;

    const exportRows = [];
    
    // 1. Копируем ВСЕ исходные строки шапки (включая служебные строки ограничений)
    for (let i = 0; i <= headerRowGlobalIndex; i++) {
        exportRows.push([...rawExcelRows[i]]);
    }

    // Определяем точные индексы длин для экспорта
    let lenT1Idx = originalHeaders.findIndex((h, idx) => h.trim() === 'заголовок 1' && idx > baseT1Idx);
    let lenT2Idx = originalHeaders.findIndex((h, idx) => h.trim() === 'заголовок 2' && idx > baseT2Idx);
    let lenTextIdx = originalHeaders.findIndex((h, idx) => h.trim() === 'текст' && idx > baseTextIdx);
    const tgoLenIndices = [lenT1Idx, lenT2Idx, lenTextIdx].filter(idx => idx !== -1);

    // 2. Добавляем обновленные строки данных
    processedDataset.forEach(item => {
        const singleRowArray = [];
        originalHeaders.forEach((headerName, curIdx) => {
            let uniqueKey = `${headerName || 'Пусто'}_${curIdx}`;
            let val = item.rowMap[uniqueKey] !== undefined ? item.rowMap[uniqueKey] : '';
            let headerLower = headerName ? headerName.toLowerCase().trim() : '';
            
            // Вычисляем базовые длины для выгрузки в файл
            if (tgoLenIndices.includes(curIdx)) {
                if (curIdx === lenT1Idx) val = item.t1.length;
                else if (curIdx === lenT2Idx) val = item.t2.length;
                else if (curIdx === lenTextIdx) {
                    let baseTextKey = `${textHeaderName}_${baseTextIdx}`;
                    val = (item.rowMap[baseTextKey] || '').length;
                }
            } 
            // Вычисляем комбинаторные длины для выгрузки в файл
            else if (headerLower.includes('длина') || headerLower.startsWith('дл.')) {
                let targetTextIdx = originalHeaders.findIndex((h, hIdx) => {
                    if (hIdx <= baseTextIdx) return false;
                    let hLower = h.toLowerCase().trim();
                    return headerLower.includes(hLower) && hIdx !== curIdx;
                });

                if (targetTextIdx !== -1) {
                    let targetKey = `${originalHeaders[targetTextIdx]}_${targetTextIdx}`;
                    val = (item.rowMap[targetKey] || '').length;
                } else {
                    val = 0;
                }
            }

            // Форматируем строго как число, чтобы Директ не ругался на строковый формат в ячейках длин
            let isLengthColumn = tgoLenIndices.includes(curIdx) || headerLower.includes('длина') || headerLower.startsWith('дл.');
            if (val !== '' && !isNaN(val) && isLengthColumn) {
                singleRowArray.push(Number(val));
            } else {
                singleRowArray.push(val);
            }
        });
        exportRows.push(singleRowArray);
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(exportRows);
    
    if (currentWorkbook && currentWorkbook.Sheets[currentWorkbook.SheetNames[0]]['!merges']) {
        ws['!merges'] = currentWorkbook.Sheets[currentWorkbook.SheetNames[0]]['!merges'];
    }

    if (outputFileName.endsWith('.csv')) {
        XLSX.utils.book_append_sheet(wb, ws, "Кампания");
        XLSX.writeFile(wb, outputFileName, { bookType: 'csv', FS: ';' });
    } else {
        XLSX.utils.book_append_sheet(wb, ws, "Кампания");
        XLSX.writeFile(wb, outputFileName);
    }
}
