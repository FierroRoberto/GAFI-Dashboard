/**
 * script.js - Dashboard Ejecutivo GAFI Ferrelectrico
 * - Hoja "familias": gráfico circular (pie) con colores, celdas de variación coloreadas (verde/rojo)
 * - Hoja "cedis cartera vencida": columnas B/C/D moneda sin decimales, gráfico de columna "May" (primeros 15)
 * - Resto de hojas sin cambios
 */

let currentSheetsData = [];
let sheetVisibility = {};
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
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                if (jsonRows.length === 0) {
                    sheetsData.push({ sheetName, rowsData: [], headers: [], worksheet });
                    return;
                }
                const headers = Object.keys(jsonRows[0]);
                const rowsData = jsonRows.map(row => headers.map(h => row[h] ?? ""));
                sheetsData.push({ sheetName, rowsData, headers, worksheet });
            });

            if (sheetsData.length === 0) {
                showEmptyState("El archivo no contiene hojas válidas.");
                return;
            }

            currentSheetsData = sheetsData;
            sheetVisibility = {};
            familiasMode = {};
            currentSheetsData.forEach(sheet => { sheetVisibility[sheet.sheetName] = true; });

            renderCheckboxesPanel();
            renderAllTables();
        };
        reader.onerror = () => alert("Error al leer el archivo.");
        reader.readAsArrayBuffer(file);
    }

    function renderCheckboxesPanel() {
        const container = document.getElementById('sheetsCheckboxes');
        if (!container) return;
        if (!currentSheetsData.length) {
            container.innerHTML = '<div class="placeholder-text">Cargue un archivo Excel para visualizar hojas.</div>';
            return;
        }
        container.innerHTML = '';
        currentSheetsData.forEach(sheet => {
            const sheetName = sheet.sheetName;
            const isChecked = sheetVisibility[sheetName] !== false;
            const div = document.createElement('div');
            div.className = 'checkbox-item';
            div.innerHTML = `
                <input type="checkbox" id="chk_${escapeId(sheetName)}" ${isChecked ? 'checked' : ''}>
                <label for="chk_${escapeId(sheetName)}">${escapeHtml(sheetName)}</label>
            `;
            const checkbox = div.querySelector('input');
            checkbox.addEventListener('change', (e) => {
                sheetVisibility[sheetName] = e.target.checked;
                updateTablesVisibility();
            });
            container.appendChild(div);
        });
    }

    function renderAllTables() {
        const viewport = document.getElementById('sheetsViewport');
        viewport.innerHTML = '';
        if (!currentSheetsData.length) {
            showEmptyState("No hay hojas para mostrar.");
            return;
        }
        currentSheetsData.forEach(sheet => {
            const card = createSheetCard(sheet);
            viewport.appendChild(card);
        });
        updateTablesVisibility();
    }

    function createSheetCard(sheet) {
        const { sheetName, rowsData, headers, worksheet } = sheet;
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
        const tableWrapper = document.createElement('div');
        tableWrapper.className = 'table-wrapper';
        const table = buildDataTable(sheetName, headers, rowsData, worksheet);
        tableWrapper.appendChild(table);
        card.appendChild(headerDiv);
        card.appendChild(tableWrapper);

        const exportBtn = headerDiv.querySelector('.btn-export-sheet');
        exportBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            exportSingleSheetToExcel(sheetName);
        });
        return card;
    }

    // ========== CONSTRUCCIÓN DE TABLA ==========
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

        const tbody = document.createElement('tbody');
        rowsData.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            headers.forEach((header, idx) => {
                const rawValue = row[idx];
                let displayValue = rawValue;
                let cellClass = '';
                const lowerHeader = normalizeString(header);
                const sheetLower = normalizeString(sheetName);

                // === REGLAS PARA "FAMILIAS" ===
                if (sheetLower === 'familias') {
                    // Columnas B, C, D, E (índices 1-4) moneda sin decimales
                    if (idx >= 1 && idx <= 4) {
                        displayValue = formatCurrency(rawValue, true);
                        cellClass = 'currency-cell';
                    }
                    // Columnas F, G (índices 5,6) porcentaje 1 decimal
                    else if (idx === 5 || idx === 6) {
                        const p = parseToPercentage(rawValue, 1);
                        displayValue = p.formatted;
                        cellClass = 'percentage-cell';
                    }
                    // Columnas de variación: "Periodo Act. vs Periodo Ant." y "Trimestre Act. vs Trimestre Ant."
                    if (lowerHeader.includes('periodo act. vs periodo ant.') || lowerHeader.includes('trimestre act. vs trimestre ant.')) {
                        const numVal = parsePercentageValue(rawValue);
                        if (!isNaN(numVal)) {
                            if (numVal > 0) cellClass += ' cell-green';
                            else if (numVal < 0) cellClass += ' cell-red';
                        }
                    }
                }

                // === REGLAS PARA "CEDIS CARTERA VENCIDA" ===
                else if (sheetLower === 'cedis cartera vencida') {
                    // Columnas B, C, D (índices 1,2,3) moneda sin decimales
                    if (idx >= 1 && idx <= 3) {
                        displayValue = formatCurrency(rawValue, true);
                        cellClass = 'currency-cell';
                    }
                }

                // Resto de reglas especiales (venta diaria, cartera vencida, semaforizacion gerente, etc.)
                else if (sheetLower === 'venta diaria' && (rowIndex + 1) % 2 === 0) {
                    displayValue = formatCurrency(rawValue, false);
                    cellClass = 'currency-cell';
                }
                else if (sheetLower === 'cartera vencida' && lowerHeader.includes('suma de > 15 dias')) {
                    displayValue = formatCurrency(rawValue, true);
                    cellClass = 'currency-cell';
                }
                else if (sheetLower === 'cartera vencida' && lowerHeader.includes('suma de % 15 dias')) {
                    const p = parseToPercentage(rawValue, 2);
                    displayValue = p.formatted;
                    cellClass = 'percentage-cell';
                    if (p.numeric > 3.50) cellClass += ' cell-yellow';
                }
                else if (sheetLower === 'semaforizacion gerente') {
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
                    const isSemaforizacion = (sheetLower === 'semaforizacion');
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

                        if (sheetLower === 'dn' && thresholdDN !== null) {
                            if (lowerHeader.includes('% cub cuota venta') || lowerHeader.includes('% cub clientes')) {
                                const percentValue = parseToPercentage(rawValue, 2);
                                const numericPercent = percentValue.numeric;
                                if (numericPercent < thresholdDN) cellClass += ' cell-red';
                                else if (numericPercent > thresholdDN) cellClass += ' cell-green';
                                else cellClass += ' cell-yellow';
                            }
                        }

                        if (lowerHeader.includes('cubrimiento') && lowerHeader.includes('cuota') && sheetLower !== 'dn' && !isSemaforizacion && sheetLower !== 'semaforizacion gerente' && sheetLower !== 'familias' && sheetLower !== 'cedis cartera vencida') {
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
        if (isNaN(num)) return "—";
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
            if (isNaN(parsed)) return { formatted: "—", numeric: 0 };
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

    // ========== GRÁFICO PARA "FAMILIAS" (circular) ==========
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
        const values = topNegativos.map(item => Math.abs(item.variacion)); // valores absolutos para el círculo
        const datasetLabel = isPeriodo ? "Periodo Act. vs Periodo Ant. (%)" : "Trimestre Act. vs Trimestre Ant. (%)";
        return { labels, values, datasetLabel };
    }

    // ========== GRÁFICO PARA "CEDIS CARTERA VENCIDA" (columna May, primeros 15) ==========
    function showCedisCarteraVencidaChart(sheetName, headers, rowsData) {
        // Buscar columna "May"
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
        
        // Tomar primeros 15 registros
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

    // ========== GRÁFICO PRINCIPAL ==========
    function showChartForSheet(sheetName, headers, rowsData, worksheet) {
        const sheetLower = normalizeString(sheetName);
        
        if (sheetLower === 'dn') {
            const chartContainer = document.getElementById('chartContainer');
            chartContainer.style.display = 'none';
            if (currentChart) {
                currentChart.destroy();
                currentChart = null;
            }
            return;
        }
        
        // --- Hoja "familias": gráfico circular ---
        if (sheetLower === 'familias') {
            if (familiasMode[sheetName] === undefined) familiasMode[sheetName] = 0;
            const currentMode = familiasMode[sheetName];
            const chartData = showFamiliesPieChart(sheetName, headers, rowsData, currentMode);
            if (!chartData) return;
            const { labels, values, datasetLabel } = chartData;
            
            const chartContainer = document.getElementById('chartContainer');
            const chartTitle = document.getElementById('chartTitle');
            const modeText = currentMode === 0 ? "Periodo" : "Trimestre";
            chartTitle.textContent = `Gráfico: ${sheetName} (${modeText} - Top 10 negativos)`;
            
            familiasMode[sheetName] = (currentMode + 1) % 2;
            
            const isDark = document.body.classList.contains('dark-theme');
            const textColor = isDark ? '#ffffff' : '#1e293b';
            
            // Generar colores variados para el gráfico circular
            const backgroundColors = labels.map((_, i) => `hsl(${(i * 360 / labels.length) % 360}, 70%, 50%)`);
            
            // Tamaño de fuente reducido para esta hoja
            const legendFontSize = 10;
            const tooltipFontSize = 9;
            
            const ctx = document.getElementById('barChart').getContext('2d');
            if (currentChart) currentChart.destroy();
            currentChart = new Chart(ctx, {
                type: 'pie',  // gráfico circular
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
        
        // --- Hoja "cedis cartera vencida": gráfico de barras con columna May (primeros 15) ---
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
            
            // Fuente reducida para esta hoja
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
        
        // ========== RESTO DE HOJAS (sin cambios) ==========
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
                const header = headers[i];
                const lower = normalizeString(header);
                if (lower.includes('cubrimiento') && lower.includes('cuota')) {
                    valueColIndex = i;
                    datasetLabel = header;
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
                const header = headers[i];
                const lower = normalizeString(header);
                if (lower.includes('cubrimiento') && lower.includes('cuota')) {
                    valueColIndex = i;
                    datasetLabel = header;
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
                const header = headers[i];
                const lower = normalizeString(header);
                if (lower.includes('venta') || lower.includes('cuota') || lower.includes('estimado') || 
                    lower.includes('cubrimiento') || lower.includes('%var') || lower === 'abr' || lower === 'feb' || lower === 'mzo') {
                    valueColIndex = i;
                    datasetLabel = header;
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

    // ========== FUNCIONES DE ACTUALIZACIÓN Y EXPORTACIÓN ==========
    function updateTablesVisibility() {
        const cards = document.querySelectorAll('.sheet-card');
        cards.forEach(card => {
            const sheetName = card.getAttribute('data-sheetname');
            card.style.display = sheetVisibility[sheetName] ? '' : 'none';
        });
        const anyVisible = Array.from(cards).some(c => c.style.display !== 'none');
        if (!anyVisible && currentSheetsData.length > 0) {
            if (!document.getElementById('tempEmptyMsg')) {
                const viewport = document.getElementById('sheetsViewport');
                const msg = document.createElement('div');
                msg.id = 'tempEmptyMsg';
                msg.className = 'empty-state';
                msg.innerHTML = '<i class="fas fa-eye-slash"></i><p>Todas las hojas están ocultas. Active al menos un checkbox.</p>';
                viewport.appendChild(msg);
            }
        } else {
            const tempMsg = document.getElementById('tempEmptyMsg');
            if (tempMsg) tempMsg.remove();
        }
        if (currentSheetsData.length === 0) showEmptyState("Cargue un archivo Excel.");
    }

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