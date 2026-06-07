/**
 * script.js — Dashboard Ejecutivo GAFI Ferrelectrico v2.0
 * Gráficas animadas: barras elásticas, dispersión con draw-SVG, dona con contador.
 */

'use strict';

let currentSheetsData  = [];
let selectedSheetName  = null;
let currentChart       = null;
let familiasMode       = {};
let ventaDiariaMode    = {};

/* ══════════════════════════════════════════════════════
   UTILIDADES
══════════════════════════════════════════════════════ */
const normalizeString = str =>
    str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().replace(/\s+/g, ' ');

const escapeHtml = str => {
    if (!str) return '';
    return str.replace(/[&<>]/g, m => ({ '&':'&amp;','<':'&lt;','>':'&gt;' }[m]));
};

const escapeId = str => str.replace(/[^a-z0-9]/gi, '_');

const isNumericValue = value => {
    if (value === null || value === undefined || value === "") return false;
    return !isNaN(parseFloat(String(value).replace(/[^0-9.-]/g, '')));
};

function showToast(msg, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.style.borderLeftColor = type === 'error' ? '#ef4444' : type === 'success' ? '#22c55e' : 'var(--accent-color)';
    toast.innerHTML = `<i class="fas fa-${type === 'error' ? 'exclamation-circle' : type === 'success' ? 'check-circle' : 'info-circle'}" style="margin-right:6px;"></i>${escapeHtml(msg)}`;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3200);
}

/* ══════════════════════════════════════════════════════
   CHART.JS — PLUGINS GLOBALES DE ANIMACIÓN
══════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════
   CHART.JS — PLUGINS GLOBALES COMPATIBLES CON v4
   En Chart.js 4, los plugins inline en el array plugins:[]
   del config NO se ejecutan. Se deben registrar con Chart.register().
══════════════════════════════════════════════════════ */

/* Estado compartido para el contador central de la dona */
const _donutState = { progress: 0, total: 0, text: '#e8edf4' };

/* Plugin global: dibuja el contador animado en el centro de la dona */
const centerCounterPlugin = {
    id: 'centerCounterPlugin',
    afterDraw(chart) {
        if (chart.config.type !== 'doughnut') return;
        const { ctx, chartArea } = chart;
        if (!chartArea) return;

        const progress = _donutState.progress;
        const current  = Math.round(_donutState.total * progress);
        const cx = (chartArea.left + chartArea.right)  / 2;
        const cy = (chartArea.top  + chartArea.bottom) / 2;

        ctx.save();
        ctx.textAlign    = 'center';
        ctx.textBaseline = 'middle';
        ctx.font         = `500 9px 'Space Mono', monospace`;
        ctx.fillStyle    = _donutState.text;
        ctx.globalAlpha  = 0.55;
        ctx.fillText('TOTAL', cx, cy - 13);
        ctx.globalAlpha  = 1;
        ctx.font         = `bold 18px 'DM Sans', sans-serif`;
        ctx.fillStyle    = _donutState.text;
        ctx.fillText(current.toLocaleString('es-MX'), cx, cy + 4);
        ctx.restore();
    }
};

/* Plugin global: dibuja la línea del scatter progresivamente */
const _scatterState = { progress: 1, accentColor: '#d32f2f', bgColor: '#0f1520' };

const drawSVGGlobalPlugin = {
    id: 'drawSVGGlobalPlugin',
    afterDatasetsDraw(chart) {
        if (chart.config.type !== 'scatter') return;
        const progress = _scatterState.progress;
        if (progress >= 1) return;

        const meta = chart.getDatasetMeta(0);
        if (!meta?.data?.length) return;

        const { ctx, chartArea } = chart;
        const points = meta.data;
        const total  = points.length;
        if (total < 2) return;

        const upTo = Math.floor(progress * (total - 1));
        const frac = (progress * (total - 1)) - upTo;

        ctx.save();
        ctx.beginPath();
        ctx.rect(chartArea.left, chartArea.top, chartArea.width, chartArea.height);
        ctx.clip();

        ctx.beginPath();
        ctx.strokeStyle = _scatterState.accentColor;
        ctx.lineWidth   = 2.5;
        ctx.lineJoin    = 'round';
        ctx.lineCap     = 'round';

        ctx.moveTo(points[0].x, points[0].y);
        for (let i = 1; i <= upTo && i < total; i++) {
            ctx.lineTo(points[i].x, points[i].y);
        }
        if (upTo < total - 1 && points[upTo + 1]) {
            ctx.lineTo(
                points[upTo].x + (points[upTo + 1].x - points[upTo].x) * frac,
                points[upTo].y + (points[upTo + 1].y - points[upTo].y) * frac
            );
        }
        ctx.stroke();

        /* Tapar los puntos que aún no deben verse */
        for (let i = upTo + 1; i < total; i++) {
            ctx.beginPath();
            ctx.arc(points[i].x, points[i].y, 8, 0, Math.PI * 2);
            ctx.fillStyle = _scatterState.bgColor;
            ctx.fill();
        }
        ctx.restore();
    }
};

Chart.register(centerCounterPlugin, drawSVGGlobalPlugin);

/* NO se toca Chart.defaults.animation — cada gráfica define la suya propia */

/* ══════════════════════════════════════════════════════
   INIT
══════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
    const fileInput      = document.getElementById('excelInput');
    const exportAllBtn   = document.getElementById('exportAllBtn');
    const themeToggleBtn = document.getElementById('themeToggleBtn');
    const closeChartBtn  = document.getElementById('closeChartBtn');
    const chartContainer = document.getElementById('chartContainer');

    /* ── TEMA ── */
    function setTheme(theme) {
        document.body.classList.remove('dark-theme', 'light-theme');
        document.body.classList.add(theme);
        localStorage.setItem('gafi-theme', theme);
        const icon = themeToggleBtn.querySelector('i');
        const span = themeToggleBtn.querySelector('span.hide-mobile') ||
                     themeToggleBtn.querySelector('span');
        if (theme === 'dark-theme') {
            icon.className = 'fas fa-moon';
            if (span) span.textContent = 'Modo Claro';
        } else {
            icon.className = 'fas fa-sun';
            if (span) span.textContent = 'Modo Oscuro';
        }
        if (currentChart && chartContainer.style.display !== 'none') {
            const titleEl = document.getElementById('chartTitle');
            if (!titleEl) return;
            let sn = titleEl.textContent.replace('Gráfico: ', '');
            if (sn.includes(' (')) sn = sn.split(' (')[0];
            const sheet = currentSheetsData.find(s => s.sheetName === sn);
            if (sheet) showChartForSheet(sheet.sheetName, sheet.headers, sheet.rowsData, sheet.worksheet);
        }
    }

    const savedTheme = localStorage.getItem('gafi-theme');
    setTheme((savedTheme === 'dark-theme' || savedTheme === 'light-theme') ? savedTheme : 'dark-theme');
    themeToggleBtn.addEventListener('click', () => {
        setTheme(document.body.classList.contains('dark-theme') ? 'light-theme' : 'dark-theme');
    });

    /* ── CERRAR CHART ── */
    closeChartBtn.addEventListener('click', () => {
        chartContainer.style.display = 'none';
        if (currentChart) { currentChart.destroy(); currentChart = null; }
    });

    /* ── CLICK EN CHART PARA ROTAR ── */
    chartContainer.addEventListener('click', e => {
        if (e.target.closest('.btn-close-chart') || chartContainer.style.display !== 'block') return;
        const titleEl = document.getElementById('chartTitle');
        if (!titleEl) return;
        let sn = titleEl.textContent.replace('Gráfico: ', '');
        if (sn.includes(' (')) sn = sn.split(' (')[0];
        const sheet = currentSheetsData.find(s => s.sheetName === sn);
        if (!sheet) return;
        const sl = normalizeString(sn);
        if (sl === 'familias') {
            showChartForSheet(sheet.sheetName, sheet.headers, sheet.rowsData, sheet.worksheet);
        } else if (sl === 'venta diaria') {
            if (ventaDiariaMode[sn] === undefined) ventaDiariaMode[sn] = 0;
            ventaDiariaMode[sn] = (ventaDiariaMode[sn] + 1) % 2;
            showChartForSheet(sheet.sheetName, sheet.headers, sheet.rowsData, sheet.worksheet);
        }
    });

    /* ── ARCHIVO EXCEL ── */
    fileInput.addEventListener('change', handleFileSelect);
    exportAllBtn.addEventListener('click', exportAllSheetsToExcel);

    /* ══════════════════════════════════════════════════════
       NORMALIZACIÓN DE PORCENTAJES — LÓGICA UNIVERSAL
       ──────────────────────────────────────────────────────
       Excel almacena porcentajes de 3 maneras distintas:

         A) Formato PORCENTAJE (cell.t === 'n', cell.z contiene '%'):
            Excel guarda internamente 0.11 y lo muestra como "11%".
            → Multiplicamos × 100 → resultado: 11.

         B) Formato NÚMERO DECIMAL (cell.t === 'n', sin '%' en z):
            Valor puede ser 0.11 (→ 11%) o 11 (→ ya es porcentaje).
            → Usamos heurística: si |v| ≤ 1.5 → es decimal → × 100.
            → Si |v| > 1.5 → ya está en escala de porcentaje → directo.

         C) Formato TEXTO (cell.t === 's'):
            Puede ser "11%", "-11%", "0.11", "11".
            → Extraemos número, detectamos '%' en string → si tiene '%'
              y |v| ≤ 1.5 → × 100; si tiene '%' y |v| > 1.5 → directo;
              sin '%' → aplicamos misma heurística que B.

         La bandera `_isExcelPercent` se guarda EN EL PROPIO VALOR
         como un objeto enriquecido { v, isPercent } para que las
         funciones de display posteriores sepan que ese campo es %.
    ══════════════════════════════════════════════════════ */

    /* ══════════════════════════════════════════════════════
       NORMALIZACIÓN DE PORCENTAJES — BASADA EN DATOS REALES
       ──────────────────────────────────────────────────────
       Tras inspeccionar los archivos Excel reales, se identificaron
       EXACTAMENTE estos patrones de numFmt en SheetJS:

       CASO A — Porcentaje decimal (valor 0.xx → mostrar XX%):
         numFmt = '0%'        → valor 0.992 → mostrar 99%
         numFmt = '0.0%'      → valor 107.3 → mostrar 107.3% ← OJO ver caso B
         numFmt = '0.00%'     → valor 0.0639 → mostrar 6.39%
         REGLA: numFmt termina en '%' SIN comillas antes → decimal → × 100

       CASO B — Número ya en escala (valor XX → mostrar XX%):
         numFmt = '0" %"'     → valor 121.3 → mostrar 121.3%
         numFmt = '0.0" %"'   → valor 112.7 → mostrar 112.7%
         REGLA: numFmt tiene '" %"' (% entre comillas) → ya en escala → directo

       CASO C — Resumen con numFmt='0.0%' pero valor ya en escala (107.3):
         La heurística: si numFmt termina en '%' Y |valor| > 1.5 → ya en escala.
         Si numFmt termina en '%' Y |valor| ≤ 1.5 → decimal → × 100.

       MANEJO DE NEGATIVOS: funciona automáticamente en todos los casos.
    ══════════════════════════════════════════════════════ */

    /**
     * Analiza la celda de SheetJS y devuelve el tipo de porcentaje.
     * Funciona con y sin cellNF/cellStyles (compatible con iOS Safari).
     *
     * Fuentes de informacion inspeccionadas en orden de confiabilidad:
     *   1. cell.z  — numFmt string (disponible sin cellNF en muchos casos)
     *   2. cell.w  — valor formateado como string ("11%", "107.3 %")
     *               SheetJS SIEMPRE lo genera, incluso sin opciones extra.
     *   3. cell.t === 's' con cell.v que contiene '%' — texto literal
     *
     * Retorna: 'quoted' | 'decimal' | 'text' | 'none'
     */
    function _detectPctFormat(cell) {
        if (!cell) return 'none';

        // ── Fuente 1: cell.z (numFmt string) ──
        var fmt = cell.z || cell.numFmt || '';
        if (typeof fmt === 'string' && fmt !== '') {
            if (fmt.indexOf('" %"') !== -1 || fmt.indexOf("' %'") !== -1) return 'quoted';
            var stripped = fmt.replace(/"[^"]*"/g, '').replace(/'[^']*'/g, '');
            if (stripped.indexOf('%') !== -1) return 'decimal';
        }

        // ── Fuente 2: cell.w (valor formateado por SheetJS, siempre presente) ──
        // Ejemplos: "11%", "107.3 %", "-5%", "99 %"
        var w = cell.w;
        if (typeof w === 'string' && w.indexOf('%') !== -1) {
            // Si el formato es "107.3 %" (espacio antes de %) → quoted (ya en escala)
            // Si el formato es "11%"  (sin espacio)            → decimal (era 0.11)
            // Distinguimos por el espacio antes del %
            if (w.indexOf(' %') !== -1) return 'quoted';
            return 'decimal';
        }

        // ── Fuente 3: celda de texto con simbolo % ──
        if (cell.t === 's' && typeof cell.v === 'string' && cell.v.indexOf('%') !== -1) {
            return 'text';
        }

        return 'none';
    }

    /**
     * Normaliza una celda de SheetJS a { value, isPercent }.
     * value siempre está en escala 0–100 si isPercent=true.
     */
    function normalizeCell(rawValue, cell) {
        if (rawValue === null || rawValue === undefined || rawValue === '') {
            return { value: '', isPercent: false };
        }

        const pctType = _detectPctFormat(cell);

        if (pctType === 'none') {
            return { value: rawValue, isPercent: false };
        }

        if (pctType === 'quoted') {
            // Valor ya en escala: 121.3 → 121.3%
            const n = typeof rawValue === 'number' ? rawValue : parseFloat(String(rawValue).replace('%',''));
            return { value: isNaN(n) ? rawValue : n, isPercent: true };
        }

        if (pctType === 'decimal') {
            // numFmt tiene '%' real (ej: '0%', '0.00%') → Excel SIEMPRE guarda decimal (0–1 o mayor si >100%)
            // Multiplicamos × 100 siempre — no heurística, porque el formato lo confirma.
            // EXCEPCIÓN: algunos campos de Resumen tienen numFmt='0.0%' pero valor ya en escala (107.3).
            // Distinguimos: si el valor viene de sheet_to_json con raw:true Y |v|>2 → ya en escala.
            // Umbral 2 (no 1.5) porque porcentajes > 200% son válidos en cubrimiento.
            const n = typeof rawValue === 'number' ? rawValue : parseFloat(String(rawValue).replace('%',''));
            if (isNaN(n)) return { value: rawValue, isPercent: false };
            // Si |n| > 3 → probablemente ya en escala (ej: 107.3% cubrimiento en Resumen)
            // Si |n| <= 3 → decimal de Excel (ej: 0.99=99%, 1.07=107%, 2.33=233%)
            const finalVal = Math.abs(n) > 3 ? n : n * 100;
            return { value: finalVal, isPercent: true };
        }

        if (pctType === 'text') {
            // Texto como "11%" o "-11%"
            const cleaned = String(rawValue).replace('%','').trim();
            const n = parseFloat(cleaned);
            if (isNaN(n)) return { value: rawValue, isPercent: false };
            // Si el número ya está en escala 0–100 (|n|>1.5) usarlo directo
            const finalVal = Math.abs(n) > 1.5 ? n : n * 100;
            return { value: finalVal, isPercent: true };
        }

        return { value: rawValue, isPercent: false };
    }

    /**
     * Construye el mapa de dirección celda → metadatos del worksheet.
     * Clave: "col,row" (0-indexed). Valor: objeto celda SheetJS.
     * Así se puede consultar la celda al procesar cada valor.
     */
    function _buildCellMap(worksheet) {
        const map = {};
        if (!worksheet) return map;
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        for (let R = range.s.r; R <= range.e.r; R++) {
            for (let C = range.s.c; C <= range.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                if (worksheet[addr]) map[`${C},${R}`] = worksheet[addr];
            }
        }
        return map;
    }

    function handleFileSelect(event) {
        var files = event.target.files;
        if (!files || files.length === 0) return;
        var file = files[0];
        if (!file) return;

        var ext = file.name.split('.').pop().toLowerCase();
        if (['xlsx','xls','xlsm','xlsb'].indexOf(ext) === -1) {
            showToast('Formato no soportado. Use .xlsx, .xls o .xlsm', 'error');
            return;
        }
        if (file.size > 52428800) {
            showToast('Archivo muy grande (max 50 MB).', 'error');
            return;
        }

        showToast('Leyendo archivo...', 'info');

        function parseAndProcess(arrayBuffer) {
            var uint8 = new Uint8Array(arrayBuffer);
            var wb;
            // Intentar parse con opciones minimas (sin cellStyles/cellNF que rompen iOS)
            try {
                wb = XLSX.read(uint8, { type: 'array', raw: true, cellDates: false });
            } catch (e1) {
                try {
                    wb = XLSX.read(uint8, { type: 'array' });
                } catch (e2) {
                    var msg = e2 && e2.message ? e2.message : String(e2);
                    showToast('No se pudo procesar el archivo: ' + msg.substring(0, 50), 'error');
                    try { event.target.value = ''; } catch(e) {}
                    return;
                }
            }
            processWorkbook(wb);
        }

        function processWorkbook(workbook) {
            var sheetsData = [];

            for (var si = 0; si < workbook.SheetNames.length; si++) {
                var sheetName = workbook.SheetNames[si];
                var worksheet = workbook.Sheets[sheetName];
                if (!worksheet || !worksheet['!ref']) {
                    sheetsData.push({ sheetName: sheetName, rowsData: [], headers: [], worksheet: worksheet, rawRows: [], cellMap: {} });
                    continue;
                }
                var jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', header: 1, raw: true });
                if (jsonRows.length === 0) {
                    sheetsData.push({ sheetName: sheetName, rowsData: [], headers: [], worksheet: worksheet, rawRows: [], cellMap: {} });
                    continue;
                }
                var cellMap = _buildCellMap(worksheet);
                var headers = (jsonRows[0] || []).map(function(h) {
                    return (h === null || h === undefined) ? '' : String(h);
                });
                var rowsData = [];
                for (var rIdx = 0; rIdx < jsonRows.length - 1; rIdx++) {
                    var row = jsonRows[rIdx + 1];
                    var rowArr = [];
                    for (var cIdx = 0; cIdx < headers.length; cIdx++) {
                        var rawVal = (row && row[cIdx] !== undefined && row[cIdx] !== null) ? row[cIdx] : '';
                        var cell   = cellMap[cIdx + ',' + (rIdx + 1)];
                        var norm   = normalizeCell(rawVal, cell);
                        if (!norm.isPercent) {
                            rowArr.push(rawVal);
                        } else {
                            var fmt = (cell && (cell.z || cell.numFmt)) || '';
                            var dec = fmt.indexOf('.00') !== -1 ? 2 : fmt.indexOf('.0') !== -1 ? 1 : 0;
                            rowArr.push({ _pct: true, _val: norm.value, _raw: rawVal, _dec: dec });
                        }
                    }
                    rowsData.push(rowArr);
                }
                sheetsData.push({ sheetName: sheetName, rowsData: rowsData, headers: headers, worksheet: worksheet, rawRows: jsonRows, cellMap: cellMap });
            }

            if (!sheetsData.length) {
                showEmptyState('El archivo no contiene hojas validas.');
                return;
            }
            currentSheetsData = sheetsData;
            familiasMode      = {};
            ventaDiariaMode   = {};
            selectedSheetName = sheetsData[0].sheetName;
            renderRadioButtons();
            renderSelectedTable();
            showToast('Archivo cargado: ' + sheetsData.length + ' hoja(s)', 'success');
        }

        /* ── FileReader iniciado SINCRONO en el mismo tick del evento ──
           CRITICO para iOS: el acceso al File object se revoca rapidamente
           cuando viene de iCloud Drive u otros proveedores externos.
           FileReader.readAsArrayBuffer iniciado sincrono conserva el acceso. */
        var reader = new FileReader();

        reader.onload = function(ev) {
            try {
                parseAndProcess(ev.target.result);
            } catch (err) {
                var msg = err && err.message ? err.message : String(err);
                showToast('Error al procesar: ' + msg.substring(0, 60), 'error');
                try { event.target.value = ''; } catch(e) {}
            }
        };

        reader.onerror = function() {
            var errMsg = reader.error ? reader.error.message : 'desconocido';
            console.error('[GAFI] FileReader.onerror:', errMsg);
            showToast('No se pudo leer el archivo.', 'error');
            try { event.target.value = ''; } catch(e) {}
        };

        reader.onabort = function() {
            showToast('Lectura cancelada.', 'error');
        };

        // Iniciar lectura inmediatamente en el mismo tick del evento (critico para iOS)
        reader.readAsArrayBuffer(file);
    }

    /* ── RADIO BUTTONS ── */
    function renderRadioButtons() {
        const container = document.getElementById('sheetsRadioGroup');
        if (!container) return;
        if (!currentSheetsData.length) {
            container.innerHTML = '<div class="placeholder-text"><i class="fas fa-file-upload" style="font-size:1.4rem;opacity:0.4;"></i><span>Cargue un archivo Excel</span></div>';
            return;
        }
        container.innerHTML = '';
        currentSheetsData.forEach((sheet, idx) => {
            const sn        = sheet.sheetName;
            const isChecked = selectedSheetName === sn;
            const div       = document.createElement('div');
            div.className   = 'radio-item' + (isChecked ? ' is-active' : '');
            div.style.animationDelay = `${idx * 40}ms`;
            div.innerHTML = `
                <input type="radio" name="sheetSelector" id="radio_${escapeId(sn)}" value="${escapeHtml(sn)}" ${isChecked ? 'checked' : ''}>
                <label for="radio_${escapeId(sn)}">${escapeHtml(sn)}</label>
            `;
            const radio = div.querySelector('input');
            radio.addEventListener('change', ev => {
                if (!ev.target.checked) return;
                // Quitar activo anterior
                container.querySelectorAll('.radio-item').forEach(el => el.classList.remove('is-active'));
                div.classList.add('is-active');
                selectedSheetName = sn;
                renderSelectedTable();
                chartContainer.style.display = 'none';
                if (currentChart) { currentChart.destroy(); currentChart = null; }
            });
            container.appendChild(div);
        });
    }

    /* ── TABLA SELECCIONADA ── */
    function renderSelectedTable() {
        const viewport = document.getElementById('sheetsViewport');
        viewport.innerHTML = '';
        if (!currentSheetsData.length) { showEmptyState("No hay hojas para mostrar."); return; }
        const sel = currentSheetsData.find(s => s.sheetName === selectedSheetName);
        if (sel) viewport.appendChild(createSheetCard(sel));
        else showEmptyState("Seleccione una hoja válida.");
    }

    /* ══════════════════════════════════════════════════════
       SHEET CARD
    ══════════════════════════════════════════════════════ */
    function createSheetCard(sheet) {
        const { sheetName, rowsData, headers, worksheet, rawRows } = sheet;
        const card = document.createElement('div');
        card.className = 'sheet-card';
        card.setAttribute('data-sheetname', sheetName);
        card.addEventListener('click', e => {
            if (e.target.closest('.btn-export-sheet')) return;
            showChartForSheet(sheetName, headers, rowsData, worksheet);
        });

        const headerDiv = document.createElement('div');
        headerDiv.className = 'sheet-header';
        headerDiv.innerHTML = `
            <h3 class="sheet-title">${escapeHtml(sheetName)}</h3>
            <button class="btn-export-sheet" data-sheet="${escapeHtml(sheetName)}">
                <i class="fas fa-file-export"></i> Exportar hoja
            </button>
        `;

        const contentWrapper = document.createElement('div');
        contentWrapper.className = 'sheet-content';

        /* ── HOJA RESUMEN ── */
        if (normalizeString(sheetName) === 'resumen' && rawRows && rawRows.length >= 2) {
            const mainHeaders = rawRows[0] || [];
            const subHeaders  = rawRows[1] || [];
            const groups = [];
            let i = 0;
            while (i < mainHeaders.length) {
                const main = mainHeaders[i];
                if (main && main.toString().trim() !== "") {
                    if (i + 1 < subHeaders.length) {
                        const dataRows = rawRows.slice(2);
                        groups.push({
                            title:     main.toString(),
                            subTitle1: (subHeaders[i]   || "").toString(),
                            subTitle2: (subHeaders[i+1] || "").toString(),
                            data1: dataRows.map(r => r[i]   ?? ""),
                            data2: dataRows.map(r => r[i+1] ?? ""),
                            colIndex: i,
                        });
                    }
                    i += 2;
                } else { i++; }
            }

            const groupsToShow = groups.slice(0, 8);
            const gridContainer = document.createElement('div');
            gridContainer.className = 'resumen-grid';
            gridContainer.style.gridTemplateColumns = 'repeat(2, 1fr)';

            let dnThreshold = null;
            if (rawRows.length > 3 && rawRows[3] && rawRows[3][5] !== undefined && rawRows[3][5] !== "") {
                const t = parseFloat(String(rawRows[3][5]).replace(/[^0-9.-]/g, ''));
                if (!isNaN(t)) dnThreshold = t;
            }

            const parsePercentageNum = value => {
                if (value === null || value === undefined || value === "") return NaN;
                let num;
                if (typeof value === 'number') num = Math.abs(value) >= 1 ? value : value * 100;
                else {
                    const parsed = parseFloat(String(value).trim().replace('%', ''));
                    if (isNaN(parsed)) return NaN;
                    num = Math.abs(parsed) >= 1 ? parsed : parsed * 100;
                }
                return num;
            };

            const getColorClassForNormalGroup = (value, title) => {
                const t = normalizeString(title);
                const num = parsePercentageNum(value);
                if (isNaN(num)) return '';
                if (t === 'mes' || t === 'cedis mes' || t === 'trimestre') {
                    if (num > 99.9) return 'cell-green';
                    if (num < 90)   return 'cell-red';
                    return 'cell-yellow';
                }
                if (t === 'cartera vencida') return num > 3.5 ? 'cell-yellow' : 'cell-green';
                return '';
            };

            const formatNormalCell = (value, isRight, title) => {
                if (value === null || value === undefined || value === "") return { display: "—", color: '' };
                if (!isRight) {
                    const num = parseFloat(String(value).replace(/[^0-9.-]/g, ''));
                    if (!isNaN(num)) return { display: formatCurrencyForResumen(value, true), color: '' };
                    return { display: value.toString(), color: '' };
                }
                const num = parsePercentageNum(value);
                if (!isNaN(num)) {
                    return { display: (Math.round(num * 10) / 10).toFixed(1) + "%", color: getColorClassForNormalGroup(value, title) };
                }
                return { display: value.toString(), color: '' };
            };

            const formatFamiliesCell = value => {
                if (value === null || value === undefined || value === "") return { display: "—", color: '' };
                const num = parsePercentageNum(value);
                if (!isNaN(num)) {
                    return { display: Math.round(num) + "%", color: num > 0 ? 'cell-green' : 'cell-red' };
                }
                return { display: value.toString(), color: '' };
            };

            const formatDNCell = (value, rowIdx) => {
                if (value === null || value === undefined || value === "") return { display: "—", color: '' };
                const num = parsePercentageNum(value);
                if (!isNaN(num)) {
                    const rounded = Math.round(num);
                    let color = '';
                    if (rowIdx === 0 && dnThreshold !== null) color = num > dnThreshold ? 'cell-green' : 'cell-red';
                    return { display: rounded + "%", color };
                }
                return { display: value.toString(), color: '' };
            };

            groupsToShow.forEach((group, gIdx) => {
                const nt = normalizeString(group.title);
                const isDN = nt === 'dn';
                const showOnlyFirst = ['mes', 'cedis mes', 'trimestre', 'cartera vencida', 'familias'].includes(nt);

                const miniCard = document.createElement('div');
                miniCard.className = 'resumen-mini-card';
                miniCard.style.animationDelay = `${gIdx * 55}ms`;

                const table = document.createElement('table');
                table.className = 'resumen-mini-table';
                const thead = document.createElement('thead');
                thead.innerHTML = `<tr><th>${escapeHtml(group.subTitle1)}</th><th>${escapeHtml(group.subTitle2)}</th></tr>`;
                table.appendChild(thead);

                const tbody = document.createElement('tbody');

                const makeCells = (left, right, extraSmall) => {
                    const tr = document.createElement('tr');
                    if (extraSmall) tr.className = 'resumen-row-small';
                    [left, right].forEach(cell => {
                        const td = document.createElement('td');
                        td.textContent = cell.display;
                        if (cell.color) td.className = cell.color;
                        tr.appendChild(td);
                    });
                    tbody.appendChild(tr);
                };

                if (group.data1.length > 0) {
                    const isLeftNum  = isNumericValue(group.data1[0]);
                    const isRightNum = isNumericValue(group.data2[0]);
                    const row1 = document.createElement('tr');
                    if (isLeftNum || isRightNum) row1.className = 'resumen-row-first';
                    let left, right;
                    if (isDN)              { left = formatDNCell(group.data1[0], 0); right = formatDNCell(group.data2[0], 0); }
                    else if (nt === 'familias') { left = formatFamiliesCell(group.data1[0]); right = formatFamiliesCell(group.data2[0]); }
                    else                   { left = formatNormalCell(group.data1[0], false, group.title); right = formatNormalCell(group.data2[0], true, group.title); }
                    [left, right].forEach(cell => {
                        const td = document.createElement('td');
                        td.textContent = cell.display;
                        if (cell.color) td.className = cell.color;
                        row1.appendChild(td);
                    });
                    tbody.appendChild(row1);
                }

                if (!showOnlyFirst && group.data1.length > 1) {
                    let left, right;
                    if (isDN)              { left = formatDNCell(group.data1[1], 1); right = formatDNCell(group.data2[1], 1); }
                    else if (nt === 'familias') { left = formatFamiliesCell(group.data1[1]); right = formatFamiliesCell(group.data2[1]); }
                    else                   { left = formatNormalCell(group.data1[1], false, group.title); right = formatNormalCell(group.data2[1], true, group.title); }
                    makeCells(left, right, isDN);
                }

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
            /* ── TABLA NORMAL ── */
            const tableWrapper = document.createElement('div');
            tableWrapper.className = 'table-wrapper';
            tableWrapper.appendChild(buildDataTable(sheetName, headers, rowsData, worksheet));
            contentWrapper.appendChild(tableWrapper);
        }

        card.appendChild(headerDiv);
        card.appendChild(contentWrapper);

        headerDiv.querySelector('.btn-export-sheet').addEventListener('click', e => {
            e.stopPropagation();
            exportSingleSheetToExcel(sheetName);
        });
        return card;
    }

    /* ══════════════════════════════════════════════════════
       BUILD DATA TABLE
    ══════════════════════════════════════════════════════ */
    function buildDataTable(sheetName, headers, rowsData, worksheet) {
        const table = document.createElement('table');
        table.className = 'data-table';

        const thead     = document.createElement('thead');
        const headerRow = document.createElement('tr');
        headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; headerRow.appendChild(th); });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        let thresholdDN = null;
        if (normalizeString(sheetName) === 'dn' && worksheet) {
            const c = worksheet['C70'];
            if (c && c.v !== undefined) { thresholdDN = parseFloat(c.v); if (isNaN(thresholdDN)) thresholdDN = null; }
        }

        const sheetLower = normalizeString(sheetName);
        let dataRows = rowsData;
        if (sheetLower === 'dn') {
            const nonEmpty = rowsData.filter(r => r.some(c => c !== undefined && c !== null && c !== ""));
            dataRows = nonEmpty.slice(0, -2);
        }

        const tbody = document.createElement('tbody');
        dataRows.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            headers.forEach((header, idx) => {
                const rawValue      = row[idx];
                // Desempacar objeto enriquecido para determinar si es porcentaje
                const isEnrichedPct = rawValue && typeof rawValue === 'object' && rawValue._pct === true;
                const cellRaw       = isEnrichedPct ? rawValue._raw : rawValue;   // valor original del Excel
                let   displayValue  = isEnrichedPct ? rawValue._val : rawValue;   // valor numérico normalizado
                let   cellClass     = '';
                const lh            = normalizeString(header);
                const sl            = sheetLower;

                // Si la celda fue detectada como porcentaje por SheetJS y no hay
                // regla específica por hoja, la mostramos como % de forma genérica.
                if (isEnrichedPct && sl !== 'familias' && sl !== 'semaforizacion gerente'
                    && sl !== 'semaforizacion' && sl !== 'dn' && sl !== 'cartera vencida'
                    && sl !== 'cedis cartera vencida' && sl !== 'resumen') {
                    const numPct = Number(rawValue._val);
                    const dec    = rawValue._dec ?? 1;
                    if (!isNaN(numPct)) {
                        displayValue = numPct.toFixed(dec) + '%';
                        cellClass = 'percentage-cell';
                        if (lh.includes('cubr') || lh.includes('cump') || lh.includes('%')) {
                            if (numPct >= 100)     cellClass += ' cell-green';
                            else if (numPct >= 90) cellClass += ' cell-yellow';
                            else                   cellClass += ' cell-red';
                        }
                    }
                }

                // ── DN: porcentajes en columnas C y D (enriquecidos o crudos)
                if (sl === 'dn' && (lh.includes('% cub cuota venta') || lh.includes('% cub clientes'))) {
                    const numPct = isEnrichedPct ? Math.round(Number(rawValue._val)) : Math.round(_resolveNumericPct(cellRaw));
                    if (!isNaN(numPct)) {
                        displayValue = numPct + '%';
                        cellClass    = 'percentage-cell';
                        if (thresholdDN !== null) {
                            if (numPct < thresholdDN)      cellClass += ' cell-red';
                            else if (numPct > thresholdDN) cellClass += ' cell-green';
                            else                           cellClass += ' cell-yellow';
                        }
                    }
                } else if (sl === 'familias') {
                    if (idx >= 1 && idx <= 4) { displayValue = formatCurrency(cellRaw, true);              cellClass = 'currency-cell'; }
                    else if (idx === 5 || idx === 6) { displayValue = parseToPercentage(rawValue, 1).formatted; cellClass = 'percentage-cell'; }
                    if (lh.includes('periodo act. vs periodo ant.') || lh.includes('trimestre act. vs trimestre ant.')) {
                        const n = parsePercentageValue(rawValue);
                        if (!isNaN(n)) cellClass += n > 0 ? ' cell-green' : n < 0 ? ' cell-red' : '';
                    }
                } else if (sl === 'cedis cartera vencida') {
                    if (idx >= 1 && idx <= 3) { displayValue = formatCurrency(cellRaw, true); cellClass = 'currency-cell'; }
                } else if (sl === 'venta diaria' && (rowIndex + 1) % 2 === 0) {
                    displayValue = formatCurrency(cellRaw, false); cellClass = 'currency-cell';
                } else if (sl === 'cartera vencida' && lh.includes('suma de > 15 dias')) {
                    displayValue = formatCurrency(cellRaw, true); cellClass = 'currency-cell';
                } else if (sl === 'cartera vencida' && lh.includes('suma de % 15 dias')) {
                    const p = parseToPercentage(rawValue, 2);
                    displayValue = p.formatted; cellClass = 'percentage-cell';
                    if (p.numeric > 3.50) cellClass += ' cell-yellow';
                } else if (sl === 'semaforizacion gerente') {
                    const pc = ['4to trim 2024','1er trim 2025','2do trim 2025','3er trim 2025','promedio'];
                    if (pc.some(c => lh.includes(c))) {
                        // normalizeCell detecta numFmt='0%' → decimal → ×100
                        // _dec=0 porque numFmt='0%' sin decimales → mostrar entero
                        const rawNum  = isEnrichedPct ? Number(rawValue._val) : _resolveNumericPct(cellRaw);
                        const dec     = isEnrichedPct ? (rawValue._dec ?? 0) : 0;
                        const num     = dec === 0 ? Math.round(rawNum) : +(rawNum.toFixed(dec));
                        if (!isNaN(num)) {
                            displayValue = num + '%'; cellClass = 'percentage-cell';
                            if (num < 90)      cellClass += ' cell-red';
                            else if (num > 99) cellClass += ' cell-green';
                            else               cellClass += ' cell-yellow';
                        }
                    }
                } else if (sl === 'semaforizacion') {
                    if (idx >= 2 && idx <= 6) {
                        const p = parseToPercentage(rawValue, 2);
                        displayValue = p.formatted; cellClass = 'percentage-cell';
                        if (p.numeric >= 100)     cellClass += ' cell-green';
                        else if (p.numeric >= 90) cellClass += ' cell-yellow';
                        else                      cellClass += ' cell-red';
                    }
                } else if (!isEnrichedPct) {
                    // Solo aplicar getFormatForCell si no fue ya manejado como porcentaje
                    const fmt = getFormatForCell(sheetName, lh, cellRaw);
                    if (fmt.formatted) displayValue = fmt.formatted;
                    if (fmt.isPercentage) cellClass = 'percentage-cell';
                    if (fmt.isCurrency)   cellClass = 'currency-cell';

                    if (sl === 'dn') {
                        if (lh.includes('venta faltante') || lh.includes('no. ctes que faltan p/cuota')) {
                            const n = parseFloat(String(cellRaw).replace(/[^0-9.-]/g, ''));
                            if (!isNaN(n)) cellClass += n > 0 ? ' cell-green' : ' cell-red';
                        }
                    }
                    if (lh.includes('cubrimiento') && lh.includes('cuota')
                        && !['dn','semaforizacion','semaforizacion gerente','familias','cedis cartera vencida','resumen'].includes(sl)) {
                        const pv = parseToPercentage(rawValue, 1);
                        if (pv.numeric < 90)        cellClass += ' cell-red';
                        else if (pv.numeric > 99.9) cellClass += ' cell-green';
                        else                        cellClass += ' cell-yellow';
                    }
                }

                const td = document.createElement('td');
                // displayValue puede ser objeto si no fue procesado — asegurar string
                const finalDisplay = (displayValue !== undefined && displayValue !== '' && displayValue !== null)
                    ? (typeof displayValue === 'object' ? JSON.stringify(displayValue) : String(displayValue))
                    : '—';
                td.textContent = finalDisplay;
                if (cellClass) td.className = cellClass.trim();
                td.style.textAlign = 'center';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        return table;
    }

    /* ══════════════════════════════════════════════════════
       FORMAT HELPERS
    ══════════════════════════════════════════════════════ */
    function formatCurrencyForResumen(value, integerMode = true) {
        const v = (value && typeof value === 'object' && '_pct' in value) ? value._raw : value;
        if (v === null || v === undefined || v === "") return "—";
        const num = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
        if (isNaN(num)) return String(v);
        const opts = integerMode
            ? { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 }
            : { style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2 };
        return new Intl.NumberFormat('es-MX', opts).format(num);
    }

    function formatCurrency(value, integerMode = false) {
        const v = (value && typeof value === 'object' && '_pct' in value) ? value._raw : value;
        if (v === null || v === undefined || v === "") return "—";
        const num = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
        if (isNaN(num)) return String(v);
        const opts = integerMode
            ? { style: 'currency', currency: 'MXN', minimumFractionDigits: 0, maximumFractionDigits: 0 }
            : { style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2 };
        return new Intl.NumberFormat('es-MX', opts).format(num);
    }

    /**
     * Extrae el valor numérico de un posible objeto enriquecido { _pct, _val }
     * o de un valor primitivo. Devuelve número en escala 0–100 (o NaN).
     * Usado por parseToPercentage y parsePercentageValue.
     */
    function _resolveNumericPct(value) {
        if (value === null || value === undefined || value === '') return NaN;
        // Objeto enriquecido por normalizeCell — ya en escala 0–100
        if (value && typeof value === 'object' && value._pct) return Number(value._val);
        // Número primitivo — aplicar heurística: >1.5 → ya en escala; ≤1.5 → decimal
        if (typeof value === 'number') {
            return Math.abs(value) > 1.5 ? value : value * 100;
        }
        // String con o sin '%'
        if (typeof value === 'string') {
            const hasSymbol = value.includes('%');
            const n = parseFloat(value.replace('%', '').trim());
            if (isNaN(n)) return NaN;
            if (hasSymbol) return Math.abs(n) > 1.5 ? n : n * 100;
            return Math.abs(n) > 1.5 ? n : n * 100;
        }
        return NaN;
    }

    /** Devuelve el valor crudo de una celda (primitivo), quitando el wrapper si existe */
    function _rawOf(value) {
        if (value && typeof value === 'object' && '_pct' in value) return value._raw;
        return value;
    }

    function parseToPercentage(value, decimals = 1) {
        if (value === null || value === undefined || value === '') return { formatted: '—', numeric: 0 };
        const num = _resolveNumericPct(value);
        if (isNaN(num)) return { formatted: String(_rawOf(value)), numeric: 0 };
        // Si el objeto enriquecido trae pista de decimales, usarla
        const dec = (value && typeof value === 'object' && value._dec !== undefined) ? value._dec : decimals;
        const factor  = 10 ** dec;
        const rounded = Math.round(num * factor) / factor;
        return { formatted: rounded.toFixed(dec) + '%', numeric: rounded };
    }

    function parsePercentageValue(value) {
        return _resolveNumericPct(value);
    }

    function getFormatForCell(sheetName, colLower, rawValue) {
        const sl = normalizeString(sheetName);

        const currencyColumns = ['venta en pesos','estimado al cierre','cuota','abr','feb','mzo','venta faltante'];
        const isPercent = c => c.includes('cubrimiento') || c.includes('%var') || c.includes('% cub');

        if (sl === 'cedis mes' || sl === 'mes') {
            if (colLower.includes('venta pesos') || colLower.includes('estimado al cierre') || colLower === 'cuota')
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            if (colLower.includes('cubrimiento') && colLower.includes('cuota'))
                return { formatted: parseToPercentage(rawValue, 1).formatted, isPercentage: true };
            if (colLower.includes('%var') || colLower.includes('% var'))
                return { formatted: parseToPercentage(rawValue, 0).formatted, isPercentage: true };
        }
        if (sl === 'trimestre') {
            if (colLower.includes('venta pesos') || colLower.includes('estimado al cierre') || colLower === 'cuota')
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
        }
        if (sl === 'cartera vencida') {
            if (['b','c','d','abr','feb','mzo'].includes(colLower))
                return { formatted: formatCurrency(rawValue, true), isCurrency: true };
        }
        if (sl === 'dn' && colLower.includes('no. ctes que faltan p/cuota')) {
            const n = parseFloat(String(rawValue).replace(/[^0-9.-]/g, ''));
            return { formatted: isNaN(n) ? "—" : Math.round(n).toLocaleString('es-MX') };
        }
        if (isPercent(colLower)) return { formatted: parseToPercentage(rawValue, 1).formatted, isPercentage: true };
        if (currencyColumns.some(c => colLower.includes(c))) {
            if (sl === 'dn' && colLower.includes('venta faltante')) return { formatted: formatCurrency(rawValue, true), isCurrency: true };
            return { formatted: formatCurrency(rawValue, false), isCurrency: true };
        }
        return { formatted: null };
    }

    /* ══════════════════════════════════════════════════════
       GRÁFICAS — ANIMACIONES PREMIUM
    ══════════════════════════════════════════════════════ */

    /** Configuración común de colores según tema */
    function getThemeColors() {
        const dark = document.body.classList.contains('dark-theme');
        return {
            isDark:    dark,
            text:      dark ? '#e8edf4' : '#0f1c2e',
            grid:      dark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.08)',
            tooltip:   dark ? '#0d1520' : '#ffffff',
            accent:    dark ? '#d32f2f' : '#2563eb',
            green:     dark ? '#22c55e' : '#16a34a',
            yellow:    dark ? '#eab308' : '#ca8a04',
            red:       dark ? '#ef4444' : '#dc2626',
            greenBg:   dark ? 'rgba(34,197,94,0.18)' : 'rgba(22,163,74,0.15)',
            yellowBg:  dark ? 'rgba(234,179,8,0.18)' : 'rgba(202,138,4,0.12)',
            redBg:     dark ? 'rgba(239,68,68,0.18)' : 'rgba(220,38,38,0.12)',
        };
    }

    /** Barras con animación stagger por barra (Chart.js 4 compatible) */
    function buildBarConfig(labels, values, colors, datasetLabel, smallFonts = false) {
        const c  = getThemeColors();
        const fs = smallFonts
            ? { legend: 9, tooltip: 8, y: 8, x: 7 }
            : { legend: 11, tooltip: 10, y: 10, x: 9 };

        return {
            type: 'bar',
            data: {
                labels,
                datasets: [{
                    label: datasetLabel,
                    data: values,
                    backgroundColor: colors,
                    borderColor: colors,
                    borderWidth: 0,
                    borderRadius: 7,
                    borderSkipped: false,
                    barPercentage: 0.55,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                animation: {
                    duration: 700,
                    easing: 'easeOutQuart',
                },
                plugins: {
                    legend: {
                        labels: {
                            color: c.text,
                            font: { size: fs.legend, weight: '600', family: 'DM Sans' },
                            boxWidth: 12,
                            padding: 10,
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${ctx.raw.toLocaleString('es-MX')}`,
                        },
                        titleColor: c.text,
                        bodyColor:  c.text,
                        backgroundColor: c.tooltip,
                        borderColor: c.accent,
                        borderWidth: 1,
                        cornerRadius: 8,
                        padding: 8,
                        titleFont: { size: fs.tooltip, family: 'Space Mono' },
                        bodyFont:  { size: fs.tooltip, family: 'DM Sans' },
                    },
                    /* Deshabilitar explícitamente el centerCounterPlugin en barras */
                    centerCounterPlugin: false,
                },
                scales: {
                    y: {
                        grid:  { color: c.grid },
                        ticks: { color: c.text, font: { size: fs.y, family: 'Space Mono' } }
                    },
                    x: {
                        grid:  { display: false },
                        ticks: { color: c.text, maxRotation: 45, minRotation: 45, font: { size: fs.x, family: 'DM Sans' } }
                    }
                }
            }
        };
    }

    /** Gráfica de dispersión con efecto "Draw SVG" (Chart.js 4 compatible) */
    function buildScatterConfig(dataPoints, label) {
        const c = getThemeColors();

        /* Inicializar estado global del scatter */
        _scatterState.progress    = 0;
        _scatterState.accentColor = c.accent;
        _scatterState.bgColor     = getComputedStyle(document.documentElement)
            .getPropertyValue('--bg-card').trim() || '#0f1520';

        return {
            type: 'scatter',
            data: {
                datasets: [{
                    label,
                    data: dataPoints,
                    backgroundColor: c.accent,
                    borderColor: c.accent,
                    borderWidth: 2.5,
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    pointHoverBackgroundColor: c.accent,
                    pointHoverBorderColor: '#ffffff',
                    pointHoverBorderWidth: 2,
                    showLine: true,
                    fill: false,
                    tension: 0,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                animation: {
                    duration: 0,   /* La animación la hace requestAnimationFrame */
                },
                plugins: {
                    legend: {
                        labels: {
                            color: c.text,
                            font: { size: 11, weight: '600', family: 'DM Sans' },
                        }
                    },
                    tooltip: {
                        callbacks: { label: ctx => ` Semana ${ctx.parsed.x}: ${ctx.parsed.y.toLocaleString('es-MX')}` },
                        titleColor: c.text,
                        bodyColor:  c.text,
                        backgroundColor: c.tooltip,
                        borderColor: c.accent,
                        borderWidth: 1,
                        cornerRadius: 8,
                        padding: 8,
                    }
                },
                scales: {
                    x: {
                        type: 'linear',
                        title: { display: true, text: 'Semana', color: c.text, font: { size: 10 } },
                        ticks: { color: c.text, stepSize: 1, font: { size: 9, family: 'Space Mono' } },
                        grid:  { color: c.grid }
                    },
                    y: {
                        title: { display: true, text: 'Monto', color: c.text, font: { size: 10 } },
                        ticks: { color: c.text, callback: v => v.toLocaleString('es-MX'), font: { size: 9, family: 'Space Mono' } },
                        grid:  { color: c.grid }
                    }
                }
            }
            /* drawSVGGlobalPlugin está registrado globalmente — sin plugins:[] inline */
        };
    }

    /** Gráfica de dona con animación + contador central (Chart.js 4 compatible) */
    function buildDoughnutConfig(labels, values, backgroundColors, datasetLabel) {
        const c = getThemeColors();
        const total = values.reduce((a, b) => a + b, 0);

        /* Inicializar estado global del contador */
        _donutState.progress = 0;
        _donutState.total    = total;
        _donutState.text     = c.text;

        return {
            type: 'doughnut',
            data: {
                labels,
                datasets: [{
                    label: datasetLabel,
                    data: values,
                    backgroundColor: backgroundColors,
                    borderColor: c.tooltip,
                    borderWidth: 2,
                    hoverBorderWidth: 3,
                    hoverBorderColor: c.accent,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                cutout: '62%',
                rotation: -90,
                circumference: 360,
                animation: {
                    animateRotate: true,
                    animateScale: false,
                    duration: 900,
                    easing: 'easeOutQuart',
                    onProgress(anim) {
                        _donutState.progress = anim.currentStep / anim.numSteps;
                    },
                    onComplete() {
                        _donutState.progress = 1;
                    },
                },
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            color: c.text,
                            font:  { size: 9, weight: '600', family: 'DM Sans' },
                            boxWidth: 10,
                            padding: 6,
                        }
                    },
                    tooltip: {
                        callbacks: { label: ctx => ` ${ctx.label}: ${ctx.raw.toFixed(1)}%` },
                        titleColor: c.text,
                        bodyColor:  c.text,
                        backgroundColor: c.tooltip,
                        borderColor: c.accent,
                        borderWidth: 1,
                        cornerRadius: 8,
                        padding: 8,
                        titleFont: { size: 9,  family: 'Space Mono' },
                        bodyFont:  { size: 10, family: 'DM Sans' },
                    },
                    /* Deshabilitar el plugin del scatter en la dona */
                    drawSVGGlobalPlugin: false,
                }
            }
            /* Sin plugins:[] inline — centerCounterPlugin está registrado globalmente */
        };
    }

    /* ══════════════════════════════════════════════════════
       CHART ORCHESTRATOR
    ══════════════════════════════════════════════════════ */
    function renderChart(config) {
        const container  = document.getElementById('chartContainer');
        const ctx        = document.getElementById('barChart').getContext('2d');
        if (currentChart) { currentChart.destroy(); currentChart = null; }
        currentChart = new Chart(ctx, config);
        container.style.display = 'block';
    }

    function showChartForSheet(sheetName, headers, rowsData, worksheet) {
        const sl = normalizeString(sheetName);
        const container = document.getElementById('chartContainer');

        /* Hojas sin gráfica */
        if (sl === 'dn' || sl === 'resumen') {
            container.style.display = 'none';
            if (currentChart) { currentChart.destroy(); currentChart = null; }
            return;
        }

        const titleEl = document.getElementById('chartTitle');
        const c       = getThemeColors();

        /* ── FAMILIAS → Dona animada ── */
        if (sl === 'familias') {
            if (familiasMode[sheetName] === undefined) familiasMode[sheetName] = 0;
            const mode    = familiasMode[sheetName];
            const chartData = showFamiliesPieChart(sheetName, headers, rowsData, mode);
            if (!chartData) return;
            const { labels, values, datasetLabel } = chartData;
            familiasMode[sheetName] = (mode + 1) % 2;
            const modeLabel = mode === 0 ? "Caída Periodo vs Periodo" : "Caída Trimestre vs Trimestre";
            titleEl.textContent = `Gráfico: ${sheetName} (${modeLabel})`;
            const bgColors = labels.map((_, i) => `hsl(${Math.round(i * 360 / labels.length)}, 70%, 55%)`);
            renderChart(buildDoughnutConfig(labels, values, bgColors, datasetLabel));
            return;
        }

        /* ── CEDIS CARTERA VENCIDA ── */
        if (sl === 'cedis cartera vencida') {
            const chartData = showCedisCarteraVencidaChart(sheetName, headers, rowsData);
            if (!chartData) return;
            const { labels, values, datasetLabel } = chartData;
            titleEl.textContent = `Gráfico: ${sheetName}`;
            const barColors = labels.map(() => c.accent);
            renderChart(buildBarConfig(labels, values, barColors, datasetLabel, true));
            return;
        }

        /* ── VENTA DIARIA ── */
        if (sl === 'venta diaria' && worksheet) {
            if (ventaDiariaMode[sheetName] === undefined) ventaDiariaMode[sheetName] = 0;
            const mode = ventaDiariaMode[sheetName];

            if (mode === 0) {
                /* Barras */
                const cols = ['A','B','C','D','E','F'];
                const labels = [], values = [];
                for (let row = 2; row <= 18; row += 2) {
                    for (const col of cols) {
                        const xC = worksheet[`${col}${row}`];
                        const yC = worksheet[`${col}${row+1}`];
                        if (xC?.v != null && yC?.v != null) {
                            const label = String(xC.v).trim() || `${col}${row}`;
                            const yVal  = parseFloat(yC.v);
                            if (!isNaN(yVal)) { labels.push(label); values.push(yVal); }
                        }
                    }
                }
                titleEl.textContent = `Gráfico: ${sheetName} (Venta Diaria)`;
                renderChart(buildBarConfig(labels, values, labels.map(() => c.accent), "Valor (moneda)"));
            } else {
                /* Dispersión con Draw SVG
                   La col G alterna texto (header de semana) y número (total).
                   Solo tomamos las celdas numéricas y les asignamos x=1,2,3... */
                const dataPoints = [];
                let semana = 1;
                for (let row = 1; row <= 200; row++) {
                    const cell = worksheet[`G${row}`];
                    if (!cell || cell.v == null || cell.v === '') break; // fin de datos
                    const yVal = parseFloat(cell.v);
                    if (!isNaN(yVal)) {               // celda numérica → punto válido
                        dataPoints.push({ x: semana, y: yVal });
                        semana++;
                    }
                    // celda de texto (header de semana) → la saltamos sin incrementar semana
                }
                if (!dataPoints.length) { showToast("Sin datos en columna G.", 'error'); container.style.display = 'none'; return; }
                titleEl.textContent = `Gráfico: ${sheetName} (Venta Por Semana)`;
                renderChart(buildScatterConfig(dataPoints, 'Venta por Semana'));

                /* Animar la línea con requestAnimationFrame usando estado global */
                let start = null;
                const duration = 1000;
                const chart = currentChart;
                function animate(ts) {
                    if (!start) start = ts;
                    const progress = Math.min((ts - start) / duration, 1);
                    _scatterState.progress = progress;
                    chart.draw();
                    if (progress < 1) requestAnimationFrame(animate);
                    else _scatterState.progress = 1;
                }
                requestAnimationFrame(animate);
            }
            return;
        }

        /* ── HOJAS ESTÁNDAR ── */
        titleEl.textContent = `Gráfico: ${sheetName}`;
        let labels = [], values = [], barColors = [], datasetLabel = '';

        let effectiveRows = rowsData;
        if (['mes','cedis mes','trimestre','cartera vencida'].includes(sl) && rowsData.length > 0)
            effectiveRows = rowsData.slice(0, -1);

        if (sl === 'mes') {
            const colIdx = headers.findIndex(h => normalizeString(h).includes('cubrimiento') && normalizeString(h).includes('cuota'));
            if (colIdx === -1) { showToast("No se encontró columna de cubrimiento.", 'error'); return; }
            const pts = effectiveRows.map(r => {
                const lbl = headers[0] && colIdx !== 0 && r[0] !== "" ? String(r[0]).substring(0, 30) : 'Fila';
                const n   = parseToPercentage(r[colIdx], 1).numeric;
                return { label: lbl, value: isNaN(n) ? 0 : n };
            }).sort((a, b) => b.value - a.value);
            labels = pts.map(p => p.label); values = pts.map(p => p.value);
            barColors = values.map(v => v >= 100 ? c.greenBg : v >= 90 ? c.yellowBg : c.redBg);
            datasetLabel = "Cubrimiento de cuota (%)";

        } else if (sl === 'cartera vencida') {
            const colIdx = headers.findIndex(h => normalizeString(h).includes('suma de % 15 dias'));
            if (colIdx === -1) { showToast("No se encontró columna 'Suma de % 15 dias'.", 'error'); return; }
            const pts = effectiveRows.map(r => {
                const lbl = headers[0] && colIdx !== 0 && r[0] !== "" ? String(r[0]).substring(0, 30) : 'Fila';
                const n   = parseToPercentage(r[colIdx], 2).numeric;
                return { label: lbl, value: isNaN(n) ? 0 : n };
            }).sort((a, b) => b.value - a.value);
            labels = pts.map(p => p.label); values = pts.map(p => p.value);
            barColors = values.map(v => v > 3.50 ? c.yellowBg : c.accent);
            datasetLabel = "Suma de % 15 dias (%)";

        } else if (sl === 'semaforizacion gerente') {
            const gIdx = headers.findIndex(h => normalizeString(h) === 'gerente');
            const pIdx = headers.findIndex(h => normalizeString(h) === 'promedio');
            if (gIdx === -1 || pIdx === -1) { showToast("No se encontraron columnas gerente/promedio.", 'error'); return; }
            // _resolveNumericPct maneja objetos enriquecidos y heurística correctamente
            const pts = effectiveRows.map(r => {
                const lbl = String(r[gIdx] || 'Gerente');
                const n   = Math.round(_resolveNumericPct(r[pIdx]));
                return { label: lbl, value: isNaN(n) ? 0 : n };
            }).sort((a, b) => b.value - a.value);
            labels = pts.map(p => p.label); values = pts.map(p => p.value);
            barColors = values.map(v => v < 90 ? c.redBg : v > 99 ? c.greenBg : c.yellowBg);
            datasetLabel = "Promedio (%)";

        } else if (sl === 'cedis mes' || sl === 'trimestre') {
            const colIdx = headers.findIndex(h => normalizeString(h).includes('cubrimiento') && normalizeString(h).includes('cuota'));
            if (colIdx === -1) { showToast("No se encontró columna de cubrimiento.", 'error'); return; }
            const pts = effectiveRows.map(r => {
                const lbl = headers[0] && colIdx !== 0 && r[0] !== "" ? String(r[0]).substring(0, 30) : 'Fila';
                const n   = parseToPercentage(r[colIdx], 1).numeric;
                return { label: lbl, value: isNaN(n) ? 0 : n };
            }).sort((a, b) => b.value - a.value);
            labels = pts.map(p => p.label); values = pts.map(p => p.value);
            barColors = values.map(v => v >= 100 ? c.greenBg : v >= 90 ? c.yellowBg : c.redBg);
            datasetLabel = headers.find(h => normalizeString(h).includes('cubrimiento')) || "Cubrimiento";

        } else {
            /* Hoja genérica */
            const TARGET_KEYWORDS = ['venta','cuota','estimado','cubrimiento','%var','abr','feb','mzo'];
            let colIdx = headers.findIndex(h => TARGET_KEYWORDS.some(kw => normalizeString(h).includes(kw)));
            if (colIdx === -1 && headers.length > 1) colIdx = 1;
            if (colIdx === -1) { showToast("Sin columna numérica válida.", 'error'); return; }
            const maxR = Math.min(rowsData.length, 15);
            for (let i = 0; i < maxR; i++) {
                const lbl = headers[0] && colIdx !== 0 && rowsData[i][0] !== "" ? String(rowsData[i][0]).substring(0, 20) : `Fila ${i+1}`;
                const n   = parseFloat(String(rowsData[i][colIdx]).replace(/[^0-9.-]/g, ''));
                labels.push(lbl); values.push(isNaN(n) ? 0 : n);
            }
            barColors = labels.map(() => c.accent);
            datasetLabel = headers[colIdx] || "Valor";
        }

        if (!labels.length) { showToast("Sin datos válidos para la gráfica.", 'error'); return; }

        /* Usar colores sólidos con gradiente en barras */
        const solidColors = barColors.map(bg => {
            if (!bg || bg === c.accent) return c.accent;
            if (bg === c.greenBg)  return c.green;
            if (bg === c.yellowBg) return c.yellow;
            if (bg === c.redBg)    return c.red;
            return bg;
        });

        const small = sl === 'mes' || sl === 'cedis cartera vencida';
        renderChart(buildBarConfig(labels, values, solidColors, datasetLabel, small));
    }

    /* ══════════════════════════════════════════════════════
       CHART HELPERS
    ══════════════════════════════════════════════════════ */
    function showFamiliesPieChart(sheetName, headers, rowsData, mode) {
        const isPeriodo   = mode === 0;
        const ventaCol    = isPeriodo ? 'venta periodo act.' : 'venta trimestre act.';
        const varCol      = isPeriodo ? 'periodo act. vs periodo ant.' : 'trimestre act. vs trimestre ant.';
        const ventaIdx    = headers.findIndex(h => normalizeString(h).includes(ventaCol));
        const varIdx      = headers.findIndex(h => normalizeString(h).includes(varCol));
        if (ventaIdx === -1 || varIdx === -1) { showToast("Columnas de familias no encontradas.", 'error'); return null; }

        const items = [];
        rowsData.forEach((r, i) => {
            const ventaNum = parseFloat(String(r[ventaIdx]).replace(/[^0-9.-]/g, ''));
            if (isNaN(ventaNum)) return;
            const varNum = parsePercentageValue(r[varIdx]);
            if (isNaN(varNum)) return;
            const label = headers[0] && ventaIdx !== 0 ? String(r[0] || `Fila ${i+1}`) : `Fila ${i+1}`;
            items.push({ label, venta: ventaNum, variacion: varNum });
        });

        items.sort((a, b) => b.venta - a.venta);
        const negativos = items.slice(0, 40).filter(it => it.variacion < 0).sort((a, b) => a.variacion - b.variacion).slice(0, 10);
        if (!negativos.length) { showToast("Sin valores negativos para mostrar.", 'error'); return null; }

        return {
            labels:       negativos.map(it => it.label),
            values:       negativos.map(it => Math.abs(it.variacion)),
            datasetLabel: isPeriodo ? "Periodo Act. vs Periodo Ant. (%)" : "Trimestre Act. vs Trimestre Ant. (%)",
        };
    }

    function showCedisCarteraVencidaChart(sheetName, headers, rowsData) {
        const mayIdx = headers.findIndex(h => normalizeString(h) === 'may');
        if (mayIdx === -1) { showToast("No se encontró columna 'May'.", 'error'); return null; }
        const maxRows = Math.min(rowsData.length, 15);
        const labels = [], values = [];
        for (let i = 0; i < maxRows; i++) {
            const label = headers[0] && mayIdx !== 0 && rowsData[i][0] !== "" ? String(rowsData[i][0]).substring(0, 30) : `Fila ${i+1}`;
            const num   = parseFloat(String(rowsData[i][mayIdx]).replace(/[^0-9.-]/g, ''));
            labels.push(label); values.push(isNaN(num) ? 0 : num);
        }
        return { labels, values, datasetLabel: headers[mayIdx] };
    }

    /* ══════════════════════════════════════════════════════
       EXPORTACIÓN
    ══════════════════════════════════════════════════════ */
    function exportSingleSheetToExcel(sheetName) {
        const sheet = currentSheetsData.find(s => s.sheetName === sheetName);
        if (!sheet) return;
        const ws = XLSX.utils.aoa_to_sheet([sheet.headers, ...sheet.rowsData]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        XLSX.writeFile(wb, `${sheetName.replace(/[\\/:*?"<>|]/g, '_')}.xlsx`);
        showToast(`Hoja "${sheetName}" exportada`, 'success');
    }

    function exportAllSheetsToExcel() {
        if (!currentSheetsData.length) { showToast("Sin datos para exportar. Cargue un archivo primero.", 'error'); return; }
        const wb = XLSX.utils.book_new();
        currentSheetsData.forEach(({ sheetName, headers, rowsData }) => {
            const ws = XLSX.utils.aoa_to_sheet([headers, ...rowsData]);
            XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));
        });
        XLSX.writeFile(wb, `GAFI_export_${new Date().toISOString().slice(0, 19)}.xlsx`);
        showToast("Exportación completa", 'success');
    }

    function showEmptyState(message) {
        const viewport = document.getElementById('sheetsViewport');
        viewport.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon-wrap"><i class="fas fa-file-excel"></i></div>
                <h3>${escapeHtml(message)}</h3>
                <p>Seleccione un archivo Excel válido (.xlsx / .xls).</p>
            </div>`;
    }
});
