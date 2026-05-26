/**
 * script.js - Dashboard Ejecutivo GAFI Ferrelectrico
 * - Hoja "resumen": minicards premium, primera fila numérica tamaño +8 medidas.
 * - Hoja "DN": eliminación de filas vacías y de las últimas 2 filas con datos.
 * - Resto de hojas sin cambios (familias, cedis cartera vencida, mes, etc.)
 * - Gráficos, temas, exportación, etc. completamente funcionales.
 */

let currentSheetsData = [];
let selectedSheetName = null;
let currentChart = null;
let familiasMode = {};

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelInput');
    const exportAllBtn = document.getElementById('exportAllBtn');
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const closeChartBtn = document.getElementById('closeChartBtn');
    const chartContainer = document.getElementById('chartContainer');

    // Tema día/noche
    function setTheme(theme) {
        document.body.classList.remove('dark-theme', 'light-theme');
        document.body.classList.add(theme);
        localStorage.setItem('gafi-theme', theme);
        const icon = themeToggleBtn.querySelector('i');
        const span = themeToggleBtn.querySelector('span');
        if (theme === 'dark-theme') {
            icon.className = 'fas fa-moon';
            span.textContent = 'Modo Claro';
        } else {
            icon.className = 'fas fa-sun';
            span.textContent = 'Modo Oscuro';
        }
        if (currentChart && chartContainer.style.display !== 'none') {
            const activeSheetName = document.getElementById('chartTitle').textContent.replace('Gráfico: ', '');
            const sheet = currentSheetsData.find(s => s.sheetName === activeSheetName);
            if (sheet) {
                showChartForSheet(sheet.sheetName, sheet.headers, sheet.rowsData, sheet.worksheet);
            }
        }
    }
    function toggleTheme() {
        const isDark = document.body.classList.contains('dark-theme');
        if (isDark) setTheme('light-theme');
        else setTheme('dark-theme');
    }
    const savedTheme = localStorage.getItem('gafi-theme');
    if (savedTheme && (savedTheme === 'dark-theme' || savedTheme === 'light-theme')) {
        setTheme(savedTheme);
    } else {
        setTheme('dark-theme');
    }
    themeToggleBtn.addEventListener('click', toggleTheme);

    closeChartBtn.addEventListener('click', () => {
        chartContainer.style.display = 'none';
        if (currentChart) {
            currentChart.destroy();
            currentChart = null;
        }
    });

    fileInput.addEventListener('change', handleFileSelect);
    exportAllBtn.addEventListener('click', () => exportAllSheetsToExcel());

    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetsData = [];

            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "", header: 1 });
                if (jsonRows.length === 0) {
                    sheetsData.push({ sheetName, rowsData: [], headers: [], worksheet, rawRows: [] });
                    return;
                }
                const headers = jsonRows[0] || [];
                const rowsData = jsonRows.slice(1).map(row => headers.map((_, i) => row[i] ?? ""));
                sheetsData.push({ sheetName, rowsData, headers, worksheet, rawRows: jsonRows });
            });

            if (sheetsData.length === 0) {
                showEmptyState("El archivo no contiene hojas válidas.");
                return;
            }

            currentSheetsData = sheetsData;
            familiasMode = {};
            if (currentSheetsData.length > 0) {
                selectedSheetName = currentSheetsData[0].sheetName;
            }
            renderRadioButtons();
            renderSelectedTable();
        };
        reader.onerror = () => alert("Error al leer el archivo.");
        reader.readAsArrayBuffer(file);
    }

    function renderRadioButtons() {
        const container = document.getElementById('sheetsRadioGroup');
        if (!container) return;
        if (!currentSheetsData.length) {
            container.innerHTML = '<div class="placeholder-text">Cargue un archivo Excel para visualizar hojas.</div>';
            return;
        }
        container.innerHTML = '';
        currentSheetsData.forEach(sheet => {
            const sheetName = sheet.sheetName;
            const isChecked = (selectedSheetName === sheetName);
            const div = document.createElement('div');
            div.className = 'radio-item';
            div.innerHTML = `
                <input type="radio" name="sheetSelector" id="radio_${escapeId(sheetName)}" value="${escapeHtml(sheetName)}" ${isChecked ? 'checked' : ''}>
                <label for="radio_${escapeId(sheetName)}">${escapeHtml(sheetName)}</label>
            `;
            const radio = div.querySelector('input');
            radio.addEventListener('change', (e) => {
                if (e.target.checked) {
                    selectedSheetName = sheetName;
                    renderSelectedTable();
                    chartContainer.style.display = 'none';
                    if (currentChart) {
                        currentChart.destroy();
                        currentChart = null;
                    }
                }
            });
            container.appendChild(div);
        });
    }

    function renderSelectedTable() {
        const viewport = document.getElementById('sheetsViewport');
        viewport.innerHTML = '';
        if (!currentSheetsData.length) {
            showEmptyState("No hay hojas para mostrar.");
            return;
        }
        const selectedSheet = currentSheetsData.find(s => s.sheetName === selectedSheetName);
        if (selectedSheet) {
            const card = createSheetCard(selectedSheet);
            viewport.appendChild(card);
        } else {
            showEmptyState("Seleccione una hoja válida.");
        }
    }

    function isNumericValue(value) {
        if (value === null || value === undefined || value === "") return false;
        const num = parseFloat(String(value).replace(/[^0-9.-]/g, ''));
        return !isNaN(num);
    }

    function createSheetCard(sheet) {
        const { sheetName, rowsData, headers, worksheet, rawRows } = sheet;
        const card = document.createElement('div');
        card.className = 'sheet-card';
        card.setAttribute('data-sheetname', sheetName);
        card.addEventListener('click', (e) => {
            if (e.target.closest('.btn-export-sheet')) return;
            showChartForSheet(sheetName, headers, rowsData, worksheet);
        });

        const headerDiv = document.createElement('div');
        headerDiv.className = 'sheet-header';
        headerDiv.innerHTML = `
            <h3 class="sheet-title">${escapeHtml(sheetName)}</h3>
            <button class="btn-export-sheet" data-sheet="${escapeHtml(sheetName)}"><i class="fas fa-file-export"></i> Exportar hoja</button>
        `;
        
        const contentWrapper = document.createElement('div');
        contentWrapper.className = 'sheet-content';
        
        // ========== HOJA "RESUMEN" (minicards) ==========
        if (normalizeString(sheetName) === 'resumen' && rawRows && rawRows.length >= 2) {
            const mainHeaders = rawRows[0] || [];
            const subHeaders = rawRows[1] || [];
            const groups = [];
            let i = 0;
            while (i < mainHeaders.length) {
                const main = mainHeaders[i];
                if (main && main.toString().trim() !== "") {
                    if (i + 1 < subHeaders.length) {
                        const sub1 = subHeaders[i] || "";
                        const sub2 = subHeaders[i+1] || "";
                        const dataRows = rawRows.slice(2);
                        const values1 = dataRows.map(row => row[i] ?? "");
                        const values2 = dataRows.map(row => row[i+1] ?? "");
                        groups.push({
                            title: main.toString(),
                            subTitle1: sub1.toString(),
                            subTitle2: sub2.toString(),
                            data1: values1,
                            data2: values2,
                            colIndex: i
                        });
                    }
                    i += 2;
                } else {
                    i++;
                }
            }
            const maxGroups = 8;
            const groupsToShow = groups.slice(0, maxGroups);
            const gridContainer = document.createElement('div');
            gridContainer.className = 'resumen-grid';
            
            const count = groupsToShow.length;
            let cols = 4;
            if (count <= 4) cols = 2;
            else if (count <= 6) cols = 3;
            else cols = 4;
            gridContainer.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
            gridContainer.style.gridTemplateRows = `repeat(2, 1fr)`;
            
            // Umbral para DN desde celda F4 (rawRows[3][5])
            let dnThreshold = null;
            if (rawRows.length > 3 && rawRows[3] && rawRows[3][5] !== undefined && rawRows[3][5] !== "") {
                const thresholdRaw = rawRows[3][5];
                dnThreshold = parseFloat(String(thresholdRaw).replace(/[^0-9.-]/g, ''));
                if (isNaN(dnThreshold)) dnThreshold = null;
            }
            
            groupsToShow.forEach(group => {
                const normalizedTitle = normalizeString(group.title);
                const isDN = normalizedTitle === 'dn';
                const singleRowGroups = ['mes', 'cedis mes', 'trimestre', 'cartera vencida', 'familias'];
                const showOnlyFirstRow = singleRowGroups.includes(normalizedTitle);
                const isSpecialGroup = (normalizedTitle === 'dn' || normalizedTitle === 'familias');
                
                const miniCard = document.createElement('div');
                miniCard.className = 'resumen-mini-card';
                
                const table = document.createElement('table');
                table.className = 'resumen-mini-table';
                const thead = document.createElement('thead');
                const headerRow = document.createElement('tr');
                const th1 = document.createElement('th');
                th1.textContent = group.subTitle1;
                const th2 = document.createElement('th');
                th2.textContent = group.subTitle2;
                headerRow.appendChild(th1);
                headerRow.appendChild(th2);
                thead.appendChild(headerRow);
                table.appendChild(thead);
                
                const tbody = document.createElement('tbody');
                
                // Funciones auxiliares
                const parsePercentageNum = (value) => {
                    if (value === null || value === undefined || value === "") return NaN;
                    let num = 0;
                    if (typeof value === 'number') {
                        if (value <= 1 && value >= -1) num = value * 100;
                        else num = value;
                    } else {
                        const str = String(value).trim().replace('%', '');
                        let parsed = parseFloat(str);
                        if (isNaN(parsed)) return NaN;
                        if (parsed <= 1 && parsed >= -1 && !String(value).includes('%')) parsed = parsed * 100;
                        num = parsed;
                    }
                    return num;
                };
                
                const getColorClassForNormalGroup = (value, groupTitle) => {
                    const titleNorm = normalizeString(groupTitle);
                    const num = parsePercentageNum(value);
                    if (isNaN(num)) return '';
                    if (titleNorm === 'mes' || titleNorm === 'cedis mes' || titleNorm === 'trimestre') {
                        if (num > 99.9) return 'cell-green';
                        if (num < 90) return 'cell-red';
                        return 'cell-yellow';
                    } else if (titleNorm === 'cartera vencida') {
                        if (num > 3.5) return 'cell-yellow';
                        return 'cell-green';
                    }
                    return '';
                };
                
                const formatNormalCell = (value, isRightColumn, groupTitle) => {
                    if (value === null || value === undefined || value === "") return { display: "—", color: '' };
                    let display = "—";
                    let color = '';
                    if (!isRightColumn) {
                        const num = parseFloat(String(value).replace(/[^0-9.-]/g, ''));
                        if (!isNaN(num)) {
                            display = formatCurrencyForResumen(value, true);
                        } else {
                            display = value.toString();
                        }
                    } else {
                        const num = parsePercentageNum(value);
                        if (!isNaN(num)) {
                            const rounded = (Math.round(num * 10) / 10).toFixed(1);
                            display = rounded + "%";
                            color = getColorClassForNormalGroup(value, groupTitle);
                        } else {
                            display = value.toString();
                        }
                    }
                    return { display, color };
                };
                
                const formatFamiliesCell = (value) => {
                    if (value === null || value === undefined || value === "") return { display: "—", color: '' };
                    const num = parsePercentageNum(value);
                    if (!isNaN(num)) {
                        const rounded = Math.round(num);
                        const color = (num > 0) ? 'cell-green' : 'cell-red';
                        return { display: rounded + "%", color };
                    }
                    return { display: value.toString(), color: '' };
                };
                
                const formatDNCell = (value, rowIndex) => {
                    if (value === null || value === undefined || value === "") return { display: "—", color: '' };
                    const num = parsePercentageNum(value);
                    let display = "—";
                    let color = '';
                    if (!isNaN(num)) {
                        const rounded = Math.round(num);
                        display = rounded + "%";
                        if (rowIndex === 0 && dnThreshold !== null && !isNaN(dnThreshold)) {
                            color = (num > dnThreshold) ? 'cell-green' : 'cell-red';
                        }
                    } else {
                        display = value.toString();
                    }
                    return { display, color };
                };
                
                // Fila 2 (índice 0) - primera fila de datos
                if (group.data1.length > 0 && group.data2.length > 0) {
                    const row2 = document.createElement('tr');
                    const leftValue = group.data1[0];
                    const rightValue = group.data2[0];
                    const isLeftNumeric = isNumericValue(leftValue);
                    const isRightNumeric = isNumericValue(rightValue);
                    if (isLeftNumeric || isRightNumeric) {
                        row2.className = 'resumen-row-first';
                    }
                    let left, right;
                    if (isDN) {
                        left = formatDNCell(leftValue, 0);
                        right = formatDNCell(rightValue, 0);
                    } else if (normalizedTitle === 'familias') {
                        left = formatFamiliesCell(leftValue);
                        right = formatFamiliesCell(rightValue);
                    } else {
                        left = formatNormalCell(leftValue, false, group.title);
                        right = formatNormalCell(rightValue, true, group.title);
                    }
                    const td1 = document.createElement('td');
                    td1.textContent = left.display;
                    if (left.color) td1.className = left.color;
                    const td2 = document.createElement('td');
                    td2.textContent = right.display;
                    if (right.color) td2.className = right.color;
                    row2.appendChild(td1);
                    row2.appendChild(td2);
                    tbody.appendChild(row2);
                }
                
                // Fila 3 (índice 1) - solo si no es de fila única
                if (!showOnlyFirstRow && group.data1.length > 1 && group.data2.length > 1) {
                    const row3 = document.createElement('tr');
                    let left, right;
                    if (isDN) {
                        left = formatDNCell(group.data1[1], 1);
                        right = formatDNCell(group.data2[1], 1);
                    } else if (normalizedTitle === 'familias') {
                        left = formatFamiliesCell(group.data1[1]);
                        right = formatFamiliesCell(group.data2[1]);
                    } else {
                        left = formatNormalCell(group.data1[1], false, group.title);
                        right = formatNormalCell(group.data2[1], true, group.title);
                    }
                    const td1 = document.createElement('td');
                    td1.textContent = left.display;
                    if (left.color) td1.className = left.color;
                    const td2 = document.createElement('td');
                    td2.textContent = right.display;
                    if (right.color) td2.className = right.color;
                    if (isDN) {
                        td1.style.fontSize = '0.65rem';
                        td2.style.fontSize = '0.65rem';
                    }
                    row3.appendChild(td1);
                    row3.appendChild(td2);
                    tbody.appendChild(row3);
                }
                // Se eliminó la cuarta fila para DN
                
                table.appendChild(tbody);
                const miniHeader = document.createElement('div');
                miniHeader.className = 'resumen-mini-header';
                miniHeader.textContent = group.title;
                miniCard.appendChild(miniHeader);
                miniCard.appendChild(table);
                gridContainer.appendChild(miniCard);
            });
            contentWrapper.appendChild(gridContainer);
        } else {
            // Tablas normales
            const tableWrapper = document.createElement('div');
            tableWrapper.className = 'table-wrapper';
            const table = buildDataTable(sheetName, headers, rowsData, worksheet);
            tableWrapper.appendChild(table);
            contentWrapper.appendChild(tableWrapper);
        }
        
        card.appendChild(headerDiv);
        card.appendChild(contentWrapper);

        const exportBtn = headerDiv.querySelector('.btn-export-sheet');
        exportBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            exportSingleSheetToExcel(sheetName);
        });
        return card;
    }
    
    // Funciones auxiliares para resumen
    function formatCurrencyForResumen(value, integerMode = true) {
        if (value === null || value === undefined || value === "") return "—";
        let num = parseFloat(String(value).replace(/[^0-9.-]/g, ''));
        if (isNaN(num)) return value.toString();
        const options = integerMode 
            ? { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 }
            : { style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2 };
        return new Intl.NumberFormat('es-MX', options).format(num);
    }
    
    function formatPercentageForResumen(value, decimals = 1) {
        if (value === null || value === undefined || value === "") return "—";
        let num = 0;
        if (typeof value === 'number') {
            if (value <= 1 && value >= -1) num = value * 100;
            else num = value;
        } else {
            const str = String(value).trim().replace('%', '');
            let parsed = parseFloat(str);
            if (isNaN(parsed)) return value.toString();
            if (parsed <= 1 && parsed >= -1 && !String(value).includes('%')) parsed = parsed * 100;
            num = parsed;
        }
        const rounded = (decimals === 0) ? Math.round(num) : (Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals));
        return rounded.toFixed(decimals) + "%";
    }

    // ========== CONSTRUCCIÓN DE TABLA ESTÁNDAR (CON FILTRADO PARA DN) ==========
    function buildDataTable(sheetName, headers, rowsData, worksheet) {
        const table = document.createElement('table');
        table.className = 'data-table';

        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headers.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        let thresholdDN = null;
        if (normalizeString(sheetName) === 'dn' && worksheet) {
            const cellC70 = worksheet['C70'];
            if (cellC70 && cellC70.v !== undefined && cellC70.v !== null) {
                thresholdDN = parseFloat(cellC70.v);
                if (isNaN(thresholdDN)) thresholdDN = null;
            }
        }

        // Filtrar filas para hoja DN (vacías y últimas 2)
        let dataRows = rowsData;
        const sheetLower = normalizeString(sheetName);
        if (sheetLower === 'dn') {
            // 1. Eliminar filas completamente vacías
            const nonEmptyRows = rowsData.filter(row => 
                row.some(cell => cell !== undefined && cell !== null && cell !== "")
            );
            // 2. Eliminar las últimas 2 filas (si existen)
            const filteredRows = nonEmptyRows.slice(0, -2);
            dataRows = filteredRows;
        }

        const tbody = document.createElement('tbody');
        dataRows.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            headers.forEach((header, idx) => {
                const rawValue = row[idx];
                let displayValue = rawValue;
                let cellClass = '';
                const lowerHeader = normalizeString(header);
                const sheetLowerLocal = normalizeString(sheetName);

                // Reglas especiales para distintas hojas
                if (sheetLowerLocal === 'familias') {
                    if (idx >= 1 && idx <= 4) {
                        displayValue = formatCurrency(rawValue, true);
                        cellClass = 'currency-cell';
                    }
                    else if (idx === 5 || idx === 6) {
                        const p = parseToPercentage(rawValue, 1);
                        displayValue = p.formatted;
                        cellClass = 'percentage-cell';
                    }
                    if (lowerHeader.includes('periodo act. vs periodo ant.') || lowerHeader.includes('trimestre act. vs trimestre ant.')) {
                        const numVal = parsePercentageValue(rawValue);
                        if (!isNaN(numVal)) {
                            if (numVal > 0) cellClass += ' cell-green';
                            else if (numVal < 0) cellClass += ' cell-red';
                        }
                    }
                }
                else if (sheetLowerLocal === 'cedis cartera vencida') {
                    if (idx >= 1 && idx <= 3) {
                        displayValue = formatCurrency(rawValue, true);
                        cellClass = 'currency-cell';
                    }
                }
                else if (sheetLowerLocal === 'venta diaria' && (rowIndex + 1) % 2 === 0) {
                    displayValue = formatCurrency(rawValue, false);
                    cellClass = 'currency-cell';
                }
                else if (sheetLowerLocal === 'cartera vencida' && lowerHeader.includes('suma de > 15 dias')) {
                    displayValue = formatCurrency(rawValue, true);
                    cellClass = 'currency-cell';
                }
                else if (sheetLowerLocal === 'cartera vencida' && lowerHeader.includes('suma de % 15 dias')) {
                    const p = parseToPercentage(rawValue, 2);
                    displayValue = p.formatted;
                    cellClass = 'percentage-cell';
                    if (p.numeric > 3.50) cellClass += ' cell-yellow';
                }
                else if (sheetLowerLocal === 'semaforizacion gerente') {
                    const percentCols = ['4to trim 2024', '1er trim 2025', '2do trim 2025', '3er trim 2025', 'promedio'];
                    const isPercentCol = percentCols.some(c => lowerHeader.includes(c));
                    if (isPercentCol) {
                        const p = parseToPercentage(rawValue, 0);
                        displayValue = p.formatted;
                        cellClass = 'percentage-cell';
                        if (p.numeric < 90) cellClass += ' cell-red';
                        else if (p.numeric > 99) cellClass += ' cell-green';
                        else cellClass += ' cell-yellow';
                    }
                }
                else {
                    const isSemaforizacion = (sheetLowerLocal === 'semaforizacion');
                    if (isSemaforizacion && idx >= 2 && idx <= 6) {
                        const percentVal = parseToPercentage(rawValue, 2);
                        displayValue = percentVal.formatted;
                        cellClass = 'percentage-cell';
                        const numeric = percentVal.numeric;
                        if (numeric >= 100) cellClass += ' cell-green';
                        else if (numeric >= 90) cellClass += ' cell-yellow';
                        else cellClass += ' cell-red';
                    } else {
                        const format = getFormatForCell(sheetName, lowerHeader, rawValue);
                        if (format.formatted) displayValue = format.formatted;
                        if (format.isPercentage) cellClass = 'percentage-cell';
                        if (format.isCurrency) cellClass = 'currency-cell';

                        if (sheetLowerLocal === 'dn' && thresholdDN !== null) {
                            if (lowerHeader.includes('% cub cuota venta') || lowerHeader.includes('% cub clientes')) {
                                const percentValue = parseToPercentage(rawValue, 2);
                                const numericPercent = percentValue.numeric;
                                if (numericPercent < thresholdDN) cellClass += ' cell-red';
                                else if (numericPercent > thresholdDN) cellClass += ' cell-green';
                                else cellClass += ' cell-yellow';
                            }
                        }

                        if (lowerHeader.includes('cubrimiento') && lowerHeader.includes('cuota') && sheetLowerLocal !== 'dn' && !isSemaforizacion && sheetLowerLocal !== 'semaforizacion gerente' && sheetLowerLocal !== 'familias' && sheetLowerLocal !== 'cedis cartera vencida' && sheetLowerLocal !== 'resumen') {
                            const percentValue = parseToPercentage(rawValue, 1);
                            const numericPercent = percentValue.numeric;
                            if (numericPercent < 90.0) cellClass += ' cell-red';
                            else if (numericPercent > 99.9) cellClass += ' cell-green';
                            else cellClass += ' cell-yellow';
                        }
                    }
                }

                const td = document.createElement('td');
                td.textContent = (displayValue !== undefined && displayValue !== "") ? displayValue : "—";
                if (cellClass) td.className = cellClass.trim();
                td.style.textAlign = 'center';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        return table;
    }

    // ========== FUNCIONES AUXILIARES ==========
    function normalizeString(str) {
        return str.toLowerCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .trim()
            .replace(/\s+/g, ' ');
    }

    function getFormatForCell(sheetName, colNameNormalized, rawValue) {
        const sheetLower = normalizeString(sheetName);
        const colLower = colNameNormalized;

        if (sheetLower === 'cedis mes') {
            if (colLower.includes('venta pesos') || colLower.includes('estimado al cierre') || colLower === 'cuota') {
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            }
            if (colLower.includes('cubrimiento') && colLower.includes('cuota')) {
                return { formatted: parseToPercentage(rawValue, 1).formatted, isPercentage: true };
            }
            if (colLower.includes('%var') || colLower.includes('% var')) {
                return { formatted: parseToPercentage(rawValue, 0).formatted, isPercentage: true };
            }
        }
        if (sheetLower === 'mes') {
            if (colLower.includes('venta pesos') || colLower.includes('estimado al cierre') || colLower === 'cuota') {
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            }
            if (colLower.includes('cubrimiento') && colLower.includes('cuota')) {
                return { formatted: parseToPercentage(rawValue, 1).formatted, isPercentage: true };
            }
            if (colLower.includes('%var') || colLower.includes('% var')) {
                return { formatted: parseToPercentage(rawValue, 0).formatted, isPercentage: true };
            }
        }
        if (sheetLower === 'trimestre') {
            if (colLower.includes('venta pesos') || colLower.includes('estimado al cierre') || colLower === 'cuota') {
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            }
        }
        if (sheetLower === 'cartera vencida') {
            if (colLower === 'b' || colLower === 'c' || colLower === 'd' || colLower === 'abr' || colLower === 'feb' || colLower === 'mzo') {
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            }
        }
        if (sheetLower === 'dn') {
            if (colLower.includes('no. ctes que faltan p/cuota')) {
                const num = parseFloat(String(rawValue).replace(/[^0-9.-]/g, ''));
                const formatted = isNaN(num) ? "—" : Math.round(num).toLocaleString('es-MX');
                return { formatted: formatted, isCurrency: false, isPercentage: false };
            }
        }
        if (colLower.includes('cubrimiento') || colLower.includes('%var') || colLower.includes('% cub')) {
            const p = parseToPercentage(rawValue, 1);
            return { formatted: p.formatted, isPercentage: true };
        }
        const currencyColumnsGeneral = ['venta en pesos', 'estimado al cierre', 'cuota', 'abr', 'feb', 'mzo', 'venta faltante'];
        if (currencyColumnsGeneral.some(c => colLower.includes(c))) {
            if (sheetLower === 'dn' && colLower.includes('venta faltante')) return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            return { formatted: formatCurrency(rawValue, false), isCurrency: true };
        }
        return { formatted: null };
    }

    function formatCurrency(value, integerMode = false) {
        if (value === null || value === undefined || value === "") return "—";
        let num = parseFloat(String(value).replace(/[^0-9.-]/g, ''));
        if (isNaN(num)) return String(value);
        const options = integerMode 
            ? { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 }
            : { style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2 };
        return new Intl.NumberFormat('es-MX', options).format(num);
    }

    function parseToPercentage(value, decimals = 1) {
        if (value === null || value === undefined || value === "") return { formatted: "—", numeric: 0 };
        let num = 0;
        if (typeof value === 'number') {
            if (value <= 1 && value >= -1) num = value * 100;
            else num = value;
        } else {
            const str = String(value).trim().replace('%', '');
            let parsed = parseFloat(str);
            if (isNaN(parsed)) return { formatted: String(value), numeric: 0 };
            if (parsed <= 1 && parsed >= -1 && !String(value).includes('%')) parsed = parsed * 100;
            num = parsed;
        }
        const rounded = (decimals === 0) ? Math.round(num) : (Math.round(num * Math.pow(10, decimals)) / Math.pow(10, decimals));
        const formatted = (decimals === 0) ? rounded.toFixed(0) + "%" : rounded.toFixed(decimals) + "%";
        return { formatted: formatted, numeric: rounded };
    }

    function parsePercentageValue(value) {
        if (value === null || value === undefined || value === "") return NaN;
        let num = 0;
        if (typeof value === 'number') {
            if (value <= 1 && value >= -1) num = value * 100;
            else num = value;
        } else {
            const str = String(value).trim().replace('%', '');
            let parsed = parseFloat(str);
            if (isNaN(parsed)) return NaN;
            if (parsed <= 1 && parsed >= -1 && !String(value).includes('%')) parsed = parsed * 100;
            num = parsed;
        }
        return num;
    }

    // ========== GRÁFICOS ==========
    function showFamiliesPieChart(sheetName, headers, rowsData, mode) {
        const isPeriodo = (mode === 0);
        const ventaColName = isPeriodo ? 'venta periodo act.' : 'venta trimestre act.';
        const varColName = isPeriodo ? 'periodo act. vs periodo ant.' : 'trimestre act. vs trimestre ant.';
        
        let ventaIdx = -1, varIdx = -1;
        for (let i = 0; i < headers.length; i++) {
            const h = normalizeString(headers[i]);
            if (h.includes(ventaColName)) ventaIdx = i;
            if (h.includes(varColName)) varIdx = i;
        }
        if (ventaIdx === -1 || varIdx === -1) {
            alert(`No se encontraron las columnas requeridas (${isPeriodo ? 'Venta periodo Act.' : 'Venta trimestre Act.'} o la de variación).`);
            return null;
        }
        
        let items = [];
        for (let i = 0; i < rowsData.length; i++) {
            let rawVenta = rowsData[i][ventaIdx];
            let ventaNum = parseFloat(String(rawVenta).replace(/[^0-9.-]/g, ''));
            if (isNaN(ventaNum)) continue;
            let rawVar = rowsData[i][varIdx];
            let varNum = parsePercentageValue(rawVar);
            if (isNaN(varNum)) continue;
            let label = (headers[0] && ventaIdx !== 0) ? String(rowsData[i][0] || `Fila ${i+1}`) : `Fila ${i+1}`;
            items.push({ label, venta: ventaNum, variacion: varNum });
        }
        
        items.sort((a, b) => b.venta - a.venta);
        const top40 = items.slice(0, 40);
        let negativos = top40.filter(item => item.variacion < 0);
        negativos.sort((a, b) => a.variacion - b.variacion);
        const topNegativos = negativos.slice(0, 10);
        
        if (topNegativos.length === 0) {
            alert(`No se encontraron valores negativos en la variación para mostrar.`);
            return null;
        }
        
        const labels = topNegativos.map(item => item.label);
        const values = topNegativos.map(item => Math.abs(item.variacion));
        const datasetLabel = isPeriodo ? "Periodo Act. vs Periodo Ant. (%)" : "Trimestre Act. vs Trimestre Ant. (%)";
        return { labels, values, datasetLabel };
    }

    function showCedisCarteraVencidaChart(sheetName, headers, rowsData) {
        let mayIdx = -1;
        for (let i = 0; i < headers.length; i++) {
            if (normalizeString(headers[i]) === 'may') {
                mayIdx = i;
                break;
            }
        }
        if (mayIdx === -1) {
            alert("No se encontró la columna 'May' en la hoja 'cedis cartera vencida'.");
            return null;
        }
        
        const maxRows = Math.min(rowsData.length, 15);
        const labels = [];
        const values = [];
        for (let i = 0; i < maxRows; i++) {
            let label = `Fila ${i+1}`;
            if (headers[0] && mayIdx !== 0) {
                let firstVal = rowsData[i][0];
                if (firstVal !== undefined && firstVal !== "") label = String(firstVal).substring(0, 30);
            }
            let rawVal = rowsData[i][mayIdx];
            let num = parseFloat(String(rawVal).replace(/[^0-9.-]/g, ''));
            if (isNaN(num)) num = 0;
            labels.push(label);
            values.push(num);
        }
        return { labels, values, datasetLabel: headers[mayIdx] };
    }

    function showChartForSheet(sheetName, headers, rowsData, worksheet) {
        const sheetLower = normalizeString(sheetName);
        
        if (sheetLower === 'dn' || sheetLower === 'resumen') {
            const chartContainer = document.getElementById('chartContainer');
            chartContainer.style.display = 'none';
            if (currentChart) {
                currentChart.destroy();
                currentChart = null;
            }
            return;
        }
        
        if (sheetLower === 'familias') {
            if (familiasMode[sheetName] === undefined) familiasMode[sheetName] = 0;
            const currentMode = familiasMode[sheetName];
            const chartData = showFamiliesPieChart(sheetName, headers, rowsData, currentMode);
            if (!chartData) return;
            const { labels, values, datasetLabel } = chartData;
            
            const chartContainer = document.getElementById('chartContainer');
            const chartTitle = document.getElementById('chartTitle');
            const modeDisplay = (currentMode === 0) ? "Periodo vs Periodo" : "Trimestre vs Trimestre";
            chartTitle.textContent = `Gráfico: ${sheetName} (${modeDisplay})`;
            
            familiasMode[sheetName] = (currentMode + 1) % 2;
            
            const isDark = document.body.classList.contains('dark-theme');
            const textColor = isDark ? '#ffffff' : '#1e293b';
            const backgroundColors = labels.map((_, i) => `hsl(${(i * 360 / labels.length) % 360}, 70%, 50%)`);
            const legendFontSize = 10;
            const tooltipFontSize = 9;
            
            const ctx = document.getElementById('barChart').getContext('2d');
            if (currentChart) currentChart.destroy();
            currentChart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: labels,
                    datasets: [{
                        label: datasetLabel,
                        data: values,
                        backgroundColor: backgroundColors,
                        borderColor: isDark ? '#333' : '#fff',
                        borderWidth: 1,
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {
                        legend: { position: 'right', labels: { color: textColor, font: { size: legendFontSize, weight: 'bold' } } },
                        tooltip: { 
                            callbacks: { label: (ctx) => `${ctx.label}: ${ctx.raw.toFixed(1)}%` },
                            titleColor: textColor,
                            bodyColor: textColor,
                            backgroundColor: isDark ? '#333' : '#fff',
                            titleFont: { size: tooltipFontSize },
                            bodyFont: { size: tooltipFontSize }
                        }
                    }
                }
            });
            chartContainer.style.display = 'block';
            return;
        }
        
        if (sheetLower === 'cedis cartera vencida') {
            const chartData = showCedisCarteraVencidaChart(sheetName, headers, rowsData);
            if (!chartData) return;
            const { labels, values, datasetLabel } = chartData;
            
            const chartContainer = document.getElementById('chartContainer');
            const chartTitle = document.getElementById('chartTitle');
            chartTitle.textContent = `Gráfico: ${sheetName} - ${datasetLabel} (primeros 15)`;
            
            const isDark = document.body.classList.contains('dark-theme');
            const barColor = isDark ? '#d32f2f' : '#2563eb';
            const textColor = isDark ? '#ffffff' : '#1e293b';
            const gridColor = isDark ? '#555' : '#ccc';
            const legendFontSize = 10;
            const tooltipFontSize = 9;
            const yTickFont = 9;
            const xTickFont = 8;
            
            const ctx = document.getElementById('barChart').getContext('2d');
            if (currentChart) currentChart.destroy();
            currentChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: datasetLabel,
                        data: values,
                        backgroundColor: barColor,
                        borderColor: barColor,
                        borderWidth: 1,
                        borderRadius: 6,
                        barPercentage: 0.5,
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {
                        legend: { labels: { color: textColor, font: { size: legendFontSize, weight: 'bold' } } },
                        tooltip: { 
                            callbacks: { label: (ctx) => `${ctx.dataset.label}: ${ctx.raw.toLocaleString('es-MX')}` },
                            titleColor: textColor,
                            bodyColor: textColor,
                            backgroundColor: isDark ? '#333' : '#fff',
                            borderColor: '#d32f2f',
                            titleFont: { size: tooltipFontSize },
                            bodyFont: { size: tooltipFontSize }
                        }
                    },
                    scales: {
                        y: { grid: { color: gridColor }, ticks: { color: textColor, font: { size: yTickFont } } },
                        x: { ticks: { color: textColor, maxRotation: 45, minRotation: 45, font: { size: xTickFont } } }
                    }
                }
            });
            chartContainer.style.display = 'block';
            return;
        }
        
        // Resto de hojas (Mes, Cartera vencida, Semaforizacion gerente, Venta diaria, Cedis Mes, Trimestre, etc.)
        const chartContainer = document.getElementById('chartContainer');
        const chartTitle = document.getElementById('chartTitle');
        chartTitle.textContent = `Gráfico: ${sheetName}`;

        let labels = [];
        let values = [];
        let datasetLabel = '';
        let barColors = [];

        if (sheetLower === 'mes') {
            let valueColIndex = -1;
            for (let i = 0; i < headers.length; i++) {
                const lower = normalizeString(headers[i]);
                if (lower.includes('cubrimiento') && lower.includes('cuota')) {
                    valueColIndex = i;
                    datasetLabel = headers[i];
                    break;
                }
            }
            if (valueColIndex === -1) {
                alert("No se encontró la columna 'cubrimiento de cuota' en la hoja Mes.");
                chartContainer.style.display = 'none';
                return;
            }
            const dataPoints = [];
            for (let i = 0; i < rowsData.length; i++) {
                let label = `Fila ${i+1}`;
                if (headers[0] && valueColIndex !== 0) {
                    let firstVal = rowsData[i][0];
                    if (firstVal !== undefined && firstVal !== "") label = String(firstVal).substring(0, 30);
                }
                let rawVal = rowsData[i][valueColIndex];
                let percentVal = parseToPercentage(rawVal, 1);
                let numeric = percentVal.numeric;
                if (isNaN(numeric)) numeric = 0;
                dataPoints.push({ label, value: numeric });
            }
            dataPoints.sort((a, b) => b.value - a.value);
            labels = dataPoints.map(dp => dp.label);
            values = dataPoints.map(dp => dp.value);
            const isDark = document.body.classList.contains('dark-theme');
            barColors = values.map(v => {
                if (v >= 100) return isDark ? '#1e4d3a' : '#22c55e';
                if (v >= 90) return isDark ? '#6b5a1a' : '#eab308';
                return isDark ? '#5e2a2a' : '#ef4444';
            });
            datasetLabel = "Cubrimiento de cuota (%)";
        }
        else if (sheetLower === 'cartera vencida') {
            let valueColIndex = -1;
            for (let i = 0; i < headers.length; i++) {
                if (normalizeString(headers[i]).includes('suma de % 15 dias')) {
                    valueColIndex = i;
                    datasetLabel = headers[i];
                    break;
                }
            }
            if (valueColIndex === -1) {
                alert("No se encontró la columna 'Suma de % 15 dias'.");
                chartContainer.style.display = 'none';
                return;
            }
            const dataPoints = [];
            for (let i = 0; i < rowsData.length; i++) {
                let label = `Fila ${i+1}`;
                if (headers[0] && valueColIndex !== 0) {
                    let firstVal = rowsData[i][0];
                    if (firstVal !== undefined && firstVal !== "") label = String(firstVal).substring(0, 30);
                }
                let rawVal = rowsData[i][valueColIndex];
                let percentVal = parseToPercentage(rawVal, 2);
                let numeric = percentVal.numeric;
                if (isNaN(numeric)) numeric = 0;
                dataPoints.push({ label, value: numeric });
            }
            dataPoints.sort((a, b) => b.value - a.value);
            labels = dataPoints.map(dp => dp.label);
            values = dataPoints.map(dp => dp.value);
            const isDark = document.body.classList.contains('dark-theme');
            barColors = dataPoints.map(dp => {
                if (dp.value > 3.50) return isDark ? '#f59e0b' : '#eab308';
                return isDark ? '#d32f2f' : '#2563eb';
            });
            datasetLabel = "Suma de % 15 dias (%)";
        }
        else if (sheetLower === 'semaforizacion gerente') {
            let gerenteIdx = -1, promedioIdx = -1;
            for (let i = 0; i < headers.length; i++) {
                const h = normalizeString(headers[i]);
                if (h === 'gerente') gerenteIdx = i;
                if (h === 'promedio') promedioIdx = i;
            }
            if (gerenteIdx === -1 || promedioIdx === -1) {
                alert("No se encontraron las columnas 'gerente' o 'promedio'.");
                chartContainer.style.display = 'none';
                return;
            }
            const dataPoints = [];
            for (let i = 0; i < rowsData.length; i++) {
                let gerente = String(rowsData[i][gerenteIdx] || `Fila ${i+1}`);
                let promedioRaw = rowsData[i][promedioIdx];
                let percentVal = parseToPercentage(promedioRaw, 0);
                let numeric = percentVal.numeric;
                if (isNaN(numeric)) numeric = 0;
                dataPoints.push({ label: gerente, value: numeric });
            }
            dataPoints.sort((a, b) => b.value - a.value);
            labels = dataPoints.map(dp => dp.label);
            values = dataPoints.map(dp => dp.value);
            const isDark = document.body.classList.contains('dark-theme');
            barColors = dataPoints.map(dp => {
                if (dp.value < 90) return isDark ? '#ef4444' : '#dc2626';
                if (dp.value > 99) return isDark ? '#22c55e' : '#16a34a';
                return isDark ? '#eab308' : '#ca8a04';
            });
            datasetLabel = "Promedio (%)";
        }
        else if (sheetLower === 'venta diaria' && worksheet) {
            const columns = ['A', 'B', 'C', 'D', 'E', 'F'];
            for (let row = 18; row >= 2; row -= 2) {
                for (let col of columns) {
                    const xCellAddr = `${col}${row}`;
                    const yCellAddr = `${col}${row + 1}`;
                    const xCell = worksheet[xCellAddr];
                    const yCell = worksheet[yCellAddr];
                    if (xCell && xCell.v !== undefined && xCell.v !== null && 
                        yCell && yCell.v !== undefined && yCell.v !== null) {
                        let label = String(xCell.v).trim();
                        if (label === "") label = `${col}${row}`;
                        let yValue = parseFloat(yCell.v);
                        if (!isNaN(yValue)) {
                            labels.push(label);
                            values.push(yValue);
                        }
                    }
                }
            }
            const isDark = document.body.classList.contains('dark-theme');
            barColors = labels.map(() => isDark ? '#3b82f6' : '#2563eb');
            datasetLabel = "Valor (moneda)";
        }
        else if (sheetLower === 'cedis mes' || sheetLower === 'trimestre') {
            let valueColIndex = -1;
            for (let i = 0; i < headers.length; i++) {
                const lower = normalizeString(headers[i]);
                if (lower.includes('cubrimiento') && lower.includes('cuota')) {
                    valueColIndex = i;
                    datasetLabel = headers[i];
                    break;
                }
            }
            if (valueColIndex === -1) {
                alert("No se encontró la columna 'cubrimiento de cuota'.");
                chartContainer.style.display = 'none';
                return;
            }
            const dataPoints = [];
            for (let i = 0; i < rowsData.length; i++) {
                let label = `Fila ${i+1}`;
                if (headers[0] && valueColIndex !== 0) {
                    let firstVal = rowsData[i][0];
                    if (firstVal !== undefined && firstVal !== "") label = String(firstVal).substring(0, 30);
                }
                let rawVal = rowsData[i][valueColIndex];
                let percentVal = parseToPercentage(rawVal, 1);
                let numeric = percentVal.numeric;
                if (isNaN(numeric)) numeric = 0;
                dataPoints.push({ label, value: numeric });
            }
            dataPoints.sort((a, b) => b.value - a.value);
            labels = dataPoints.map(dp => dp.label);
            values = dataPoints.map(dp => dp.value);
            const isDark = document.body.classList.contains('dark-theme');
            barColors = values.map(v => {
                if (v >= 100) return isDark ? '#1e4d3a' : '#22c55e';
                if (v >= 90) return isDark ? '#6b5a1a' : '#eab308';
                return isDark ? '#5e2a2a' : '#ef4444';
            });
            datasetLabel = "Cubrimiento de cuota (%)";
        }
        else {
            let valueColIndex = -1;
            for (let i = 0; i < headers.length; i++) {
                const lower = normalizeString(headers[i]);
                if (lower.includes('venta') || lower.includes('cuota') || lower.includes('estimado') || 
                    lower.includes('cubrimiento') || lower.includes('%var') || lower === 'abr' || lower === 'feb' || lower === 'mzo') {
                    valueColIndex = i;
                    datasetLabel = headers[i];
                    break;
                }
            }
            if (valueColIndex === -1 && headers.length > 1) {
                valueColIndex = 1;
                datasetLabel = headers[1];
            }
            if (valueColIndex === -1) {
                alert("No se encontró una columna numérica adecuada.");
                chartContainer.style.display = 'none';
                return;
            }
            const maxRows = Math.min(rowsData.length, 15);
            for (let i = 0; i < maxRows; i++) {
                let label = `Fila ${i+1}`;
                if (headers[0] && valueColIndex !== 0) {
                    let firstVal = rowsData[i][0];
                    if (firstVal !== undefined && firstVal !== "") label = String(firstVal).substring(0, 20);
                }
                let rawVal = rowsData[i][valueColIndex];
                let num = parseFloat(String(rawVal).replace(/[^0-9.-]/g, ''));
                if (isNaN(num)) num = 0;
                labels.push(label);
                values.push(num);
            }
            const isDark = document.body.classList.contains('dark-theme');
            barColors = labels.map(() => isDark ? '#d32f2f' : '#2563eb');
        }

        if (labels.length === 0) {
            alert("No se encontraron datos válidos para generar el gráfico.");
            chartContainer.style.display = 'none';
            return;
        }

        const isDark = document.body.classList.contains('dark-theme');
        const textColor = isDark ? '#ffffff' : '#1e293b';
        const gridColor = isDark ? '#555' : '#ccc';

        let legendFontSize = 13;
        let tooltipTitleFont = 12;
        let tooltipBodyFont = 12;
        let yTickFont = 12;
        let xTickFont = 11;

        if (sheetLower === 'mes' || sheetLower === 'cedis cartera vencida') {
            legendFontSize = 10;
            tooltipTitleFont = 9;
            tooltipBodyFont = 9;
            yTickFont = 9;
            xTickFont = 8;
        }

        const ctx = document.getElementById('barChart').getContext('2d');
        if (currentChart) currentChart.destroy();
        currentChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: datasetLabel,
                    data: values,
                    backgroundColor: barColors,
                    borderColor: barColors,
                    borderWidth: 1,
                    borderRadius: 6,
                    barPercentage: 0.5,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { labels: { color: textColor, font: { size: legendFontSize, weight: 'bold' } } },
                    tooltip: { 
                        callbacks: { 
                            label: (ctx) => {
                                let suffix = '';
                                if (sheetLower === 'cedis mes' || sheetLower === 'trimestre' || sheetLower === 'cartera vencida' || sheetLower === 'semaforizacion gerente' || sheetLower === 'mes')
                                    suffix = '%';
                                return `${ctx.dataset.label}: ${ctx.raw.toLocaleString('es-MX')}${suffix}`;
                            }
                        },
                        titleColor: textColor,
                        bodyColor: textColor,
                        backgroundColor: isDark ? '#333' : '#fff',
                        borderColor: '#d32f2f',
                        titleFont: { size: tooltipTitleFont },
                        bodyFont: { size: tooltipBodyFont }
                    }
                },
                scales: {
                    y: { grid: { color: gridColor }, ticks: { color: textColor, font: { size: yTickFont } } },
                    x: { ticks: { color: textColor, maxRotation: 45, minRotation: 45, font: { size: xTickFont } } }
                }
            }
        });
        chartContainer.style.display = 'block';
    }

    // ========== EXPORTACIÓN Y UTILIDADES ==========
    function exportSingleSheetToExcel(sheetName) {
        const sheet = currentSheetsData.find(s => s.sheetName === sheetName);
        if (!sheet) return;
        const { headers, rowsData } = sheet;
        const sheetData = [headers, ...rowsData];
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        XLSX.writeFile(wb, `${sheetName.replace(/[\\/:*?"<>|]/g, '_')}.xlsx`);
    }

    function exportAllSheetsToExcel() {
        if (!currentSheetsData.length) {
            alert("No hay datos para exportar. Primero cargue un archivo.");
            return;
        }
        const wb = XLSX.utils.book_new();
        currentSheetsData.forEach(sheet => {
            const { sheetName, headers, rowsData } = sheet;
            const sheetData = [headers, ...rowsData];
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        });
        XLSX.writeFile(wb, `GAFI_export_${new Date().toISOString().slice(0,19)}.xlsx`);
    }

    function showEmptyState(message) {
        const viewport = document.getElementById('sheetsViewport');
        viewport.innerHTML = `<div class="empty-state"><i class="fas fa-file-excel"></i><h3>${escapeHtml(message)}</h3><p>Seleccione un archivo Excel válido.</p></div>`;
    }

    function escapeHtml(str) {
        if (!str) return '';
        return str.replace(/[&<>]/g, function(m) {
            if (m === '&') return '&amp;';
            if (m === '<') return '&lt;';
            if (m === '>') return '&gt;';
            return m;
        });
    }
    function escapeId(str) {
        return str.replace(/[^a-z0-9]/gi, '_');
    }
});