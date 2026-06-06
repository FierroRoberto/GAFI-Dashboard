# GAFI Ferrelectrico — Executive Dashboard

Dashboard ejecutivo de KPIs para GAFI Ferrelectrico. Lee archivos Excel directamente en el navegador (sin backend), normaliza porcentajes automáticamente y visualiza datos con gráficas animadas.

## ✨ Características

| Funcionalidad | Detalle |
|---|---|
| 📊 **Lectura de Excel** | SheetJS — procesamiento 100% local, datos nunca salen del dispositivo |
| 🔢 **Normalización de %** | Detecta automáticamente formato texto, decimal y porcentaje de Excel |
| 📈 **Gráficas animadas** | Barras (Chart.js 4), dispersión con Draw-SVG, dona con contador central |
| 🎨 **Tema dual** | Oscuro / Claro con persistencia en `localStorage` |
| 📱 **PWA** | Instalable como app, funciona offline para el shell |
| 🔒 **Privacidad total** | Todo el procesamiento es local — ningún dato se envía a servidores |

## 🚀 Uso rápido

1. Clona o descarga el repositorio
2. Abre `index.html` en un navegador moderno (o sirve con un servidor local)
3. Haz clic en **Seleccionar Excel** y carga tu archivo `.xlsx`
4. Navega entre hojas en la barra lateral

> **Nota:** Para que el Service Worker funcione correctamente, sirve el proyecto desde un servidor local o GitHub Pages (no desde `file://`).

## 🖥️ Servidor local rápido

```bash
# Python 3
python -m http.server 8080

# Node.js (npx)
npx serve .

# VS Code
# Instala la extensión "Live Server" y haz clic en "Go Live"
```

Luego abre: `http://localhost:8080`

## 🌐 Deploy en GitHub Pages

1. Sube el repositorio a GitHub
2. Ve a **Settings → Pages**
3. En "Source" selecciona **Deploy from a branch** → rama `main` → carpeta `/ (root)`
4. GitHub Pages publicará la URL en minutos
5. La PWA estará disponible para instalación desde esa URL

## 📁 Estructura del proyecto

```
gafi-dashboard/
├── index.html          # Shell principal + PWA meta tags
├── style.css           # Estilos premium con variables CSS + animaciones
├── script.js           # Lógica principal + normalización de % + gráficas
├── sw.js               # Service Worker (Cache-First para shell, Network para Excel)
├── manifest.json       # Manifiesto PWA
├── logo-gafi.png       # Logo (proveer externamente)
├── icons/              # Íconos PWA (proveer externamente)
│   ├── icon-72.png
│   ├── icon-96.png
│   ├── icon-128.png
│   ├── icon-144.png
│   ├── icon-152.png
│   ├── icon-192.png
│   ├── icon-384.png
│   └── icon-512.png
├── .gitignore
└── README.md
```

## 🧠 Lógica de normalización de porcentajes

Excel almacena los porcentajes de tres formas distintas según cómo fue ingresado el dato:

| Formato en Excel | Valor interno SheetJS | Resultado normalizado |
|---|---|---|
| `11%` (formato %) | `0.11` + `cell.z` contiene `%` | `11.0%` ✅ |
| `-11%` (formato %) | `-0.11` + `cell.z` contiene `%` | `-11.0%` ✅ |
| `"11%"` (texto) | `"11%"` (string) | `11.0%` ✅ |
| `0.11` (número sin formato %) | `0.11` sin `cell.z` con % | Heurística: `≤1.5 → ×100 = 11%` ✅ |
| `11` (número ya en escala %) | `11` sin `cell.z` con % | Heurística: `>1.5 → directo = 11%` ✅ |

**Por qué esta lógica es la mejor opción:**
- Usa los **metadatos de celda de SheetJS** (`cell.z`, `cell.t`, `cell.numFmt`) como fuente de verdad primaria — es la forma más confiable de saber si Excel consideraba ese campo como porcentaje.
- Solo recurre a **heurística** (`|v| ≤ 1.5`) cuando no hay metadatos (p. ej. valores pegados como texto plano).
- Maneja **negativos** en todos los casos sin lógica especial adicional.
- Es **no destructiva**: el valor original siempre se conserva en `_raw` para operaciones de exportación y moneda.

## 🛠️ Tecnologías

- [SheetJS (xlsx)](https://sheetjs.com/) — lectura de Excel en frontend
- [Chart.js 4](https://www.chartjs.org/) — gráficas animadas
- [DM Sans + Space Mono](https://fonts.google.com/) — tipografía
- [Font Awesome 6](https://fontawesome.com/) — íconos
- Service Worker API — PWA offline

## 📄 Licencia

Uso interno GAFI Ferrelectrico. Por GVR RFA.
