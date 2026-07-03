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
    const seenAdsKeys = new Set();

    // Находим индексы всех колонок, связанных с заголовками и текстами (включая комбинаторные)
    const textContentIndices = [];
    originalHeaders.forEach((header, idx) => {
        const hLower = header.toLowerCase();
        if (hLower.includes('заголовок') || hLower.includes('текст объявления') || hLower === 'текст') {
            textContentIndices.push(idx);
        }
    });

    for (let i = startDataRow; i < rawExcelRows.length; i++) {
        const row = rawExcelRows[i];
        if (!row || row.length === 0) continue;

        const rowMap = {};
        originalHeaders.forEach((header, colIdx) => {
            rowMap[header] = row[colIdx] !== undefined ? String(row[colIdx]).trim() : '';
        });

        const title1 = rowMap[t1HeaderName] || '';
        const title2 = rowMap[t2HeaderName] || '';
        const text = rowMap[textHeaderName] || '';
        const adType = rowMap[typeHeaderName] || '—';
        const adId = rowMap[idAdHeaderName] || '';

        // Собираем полный текстовый слепок строки для точной идентификации комбинации
        const textContentSnapshot = textContentIndices.map(idx => String(row[idx] || '').trim()).join('|');

        // Ключ теперь уникален не просто по ID, а по связке ID + весь текстовый состав.
        // Если ID нет, то по типу объявления + всему контенту.
        const uniqueKey = adId 
            ? `${adType}_id_${adId}_content_${textContentSnapshot}` 
            : `${adType}_content_${textContentSnapshot}`;

        if (seenAdsKeys.has(uniqueKey)) {
            continue;
        }

        seenAdsKeys.add(uniqueKey);
        
        const analyzedRow = computeRowMetrics(i, title1, title2, adType, rowMap, uniqueKey);
        processedDataset.push(analyzedRow);
    }

    updateDashboardStats();
    buildTableHeader();
    renderFullTable();
}
