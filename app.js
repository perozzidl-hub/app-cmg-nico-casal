// ==========================================
// CONFIGURACIÓN GENERAL Y ESTADO GLOBAL
// ==========================================
const AppState = {
    rawData: { ventas: [], recetas: [], costos: [] },
    filteredData: [],
    chartInstances: { ventas: null, costos: null, composicion: null },
    isDarkMode: false,
};

// Utilidades de formato para directivos (Formato corto: $1.2M, $450K)
const Utils = {
    formatMoneyShort(val) {
        if (!val) return '$0';
        if (Math.abs(val) >= 1e9) return '$' + (val / 1e9).toFixed(2) + 'B';
        if (Math.abs(val) >= 1e6) return '$' + (val / 1e6).toFixed(1) + 'M';
        if (Math.abs(val) >= 1e3) return '$' + (val / 1e3).toFixed(1) + 'K';
        return '$' + val.toFixed(2);
    },
    parseMonth(fechaStr) {
        if (!fechaStr) return 'Sin Mes';
        const parts = String(fechaStr).trim().split(' ')[0].split('-');
        if (parts.length < 2) return 'Sin Mes';
        return parts[0] + '-' + parts[1];
    },
    monthName(ym) {
        const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
        const parts = ym.split('-');
        return meses[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
    },
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(this, args), wait);
        };
    }
};

// ==========================================
// MANEJO DE ARCHIVOS (Lectura unificada)
// ==========================================
// Optimización Senior: Lee un solo Excel, buscando automáticamente las hojas
// por nombre (Ventas, Recetas, Costos) para no saturar al usuario.
const FileService = {
    async handleUpload(file) {
        const reader = new FileReader();
        return new Promise((resolve, reject) => {
            reader.onload = (e) => {
                try {
                    const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                    const data = {};
                    
                    // Lógica de mapeo dinámico de hojas
                    const sheetNames = wb.SheetNames;
                    const findSheet = (names) => sheetNames.find(name => names.some(n => name.toLowerCase().includes(n)));
                    
                    const ventasSheet = findSheet(['venta', 'data', 'datos']);
                    const recetasSheet = findSheet(['receta', 'bom']);
                    const costosSheet = findSheet(['costo', 'insumo']);

                    if (ventasSheet) data.ventas = XLSX.utils.sheet_to_json(wb.Sheets[ventasSheet], { header: 1, defval: '' });
                    else data.ventas = XLSX.utils.sheet_to_json(wb.Sheets[sheetNames[0]], { header: 1, defval: '' });
                    
                    if (recetasSheet) data.recetas = XLSX.utils.sheet_to_json(wb.Sheets[recetasSheet], { header: 1, defval: '' });
                    if (costosSheet) data.costos = XLSX.utils.sheet_to_json(wb.Sheets[costosSheet], { header: 1, defval: '' });

                    resolve(data);
                } catch (err) { reject(err); }
            };
            reader.readAsArrayBuffer(file);
        });
    }
};

// ==========================================
// LÓGICA DE NEGOCIO (Procesamiento Eficiente)
// ==========================================
const DataProcessor = {
    normalizeHeaders(row) {
        return row.map(h => h ? h.toString().trim().replace(/[^a-zA-Z0-9_ ]/g, '').trim().replace(/ /g, '_') : `col_${row.indexOf(h)}`);
    },
    
    processData(dataArray) {
        if (!dataArray || dataArray.length < 2) return [];
        const headers = this.normalizeHeaders(dataArray[0]);
        const rows = dataArray.slice(1);
        return rows.map(row => {
            const obj = {};
            headers.forEach((key, idx) => obj[key] = row[idx] !== undefined ? row[idx] : '');
            // Convertir numéricos clave
            ['Físicos', 'Facturación_Neta', 'TOTAL_INSUMOS'].forEach(c => obj[c] = parseFloat(obj[c]) || 0);
            obj._mes = Utils.parseMonth(obj.FECHA);
            return obj;
        }).filter(o => o.ART && o.ART !== '');
    },

    // OPTIMIZACIÓN CLAVE: Uso de Map para pasar de O(NxM) a O(N)
    recalculateCosts() {
        if (!AppState.rawData.recetas.length || !AppState.rawData.costos.length) return AppState.rawData.ventas;

        const costosMap = new Map();
        AppState.rawData.costos.forEach(c => {
            const key = `${c.idInsumo}|${c._mes}`;
            if (!costosMap.has(key) || c.fecha > costosMap.get(key).fecha) costosMap.set(key, c);
        });

        const recetasMap = new Map();
        AppState.rawData.recetas.forEach(r => {
            if (!recetasMap.has(r.idArticulo)) recetasMap.set(r.idArticulo, []);
            recetasMap.get(r.idArticulo).push(r);
        });

        return AppState.rawData.ventas.map(venta => {
            const recetasArt = recetasMap.get(venta.ART) || [];
            let costoTotalRecalculado = 0;
            
            recetasArt.forEach(rec => {
                const costo = costosMap.get(`${rec.idInsumo}|${venta._mes}`);
                if (costo) costoTotalRecalculado += (rec.cantidad * costo.precioUnitario);
            });

            venta.TOTAL_INSUMOS = costoTotalRecalculado; // Actualizar costo
            return venta;
        });
    }
};

// ==========================================
// RENDERIZADO UI (Experiencia de Usuario)
// ==========================================
const UIRenderer = {
    updateKPIs(data) {
        const totalUnidades = data.reduce((acc, o) => acc + (o.Físicos || 0), 0);
        const sumFacturacion = data.reduce((acc, o) => acc + (o.Facturación_Neta || 0), 0);
        const sumCosto = data.reduce((acc, o) => acc + ((o.TOTAL_INSUMOS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0)), 0);
        const margen = sumFacturacion > 0 ? ((sumFacturacion - sumCosto) / sumFacturacion) * 100 : 0;
        
        // Formato elegante para directivos
        document.getElementById('totalUnidades').textContent = totalUnidades.toLocaleString('es-AR');
        document.getElementById('facturacionNeta').textContent = Utils.formatMoneyShort(sumFacturacion);
        document.getElementById('costoTotal').textContent = Utils.formatMoneyShort(sumCosto);
        document.getElementById('margenBruto').textContent = margen.toFixed(1) + '%';
        document.getElementById('costoUnitario').textContent = Utils.formatMoneyShort(sumCosto / (totalUnidades || 1));
    },

    updateCharts(data) {
        // Destruir gráficos anteriores para evitar memory leaks
        Object.values(AppState.chartInstances).forEach(chart => chart && chart.destroy());

        // 1. Tendencia de Ventas (Área/Línea)
        const ventasPorMes = {};
        data.forEach(o => ventasPorMes[o._mes] = (ventasPorMes[o._mes] || 0) + (o.Físicos || 0));
        const meses = Object.keys(ventasPorMes).sort();
        const ctx1 = document.getElementById('chartVentas').getContext('2d');
        AppState.chartInstances.ventas = new Chart(ctx1, {
            type: 'line',
            data: { labels: meses.map(Utils.monthName), datasets: [{ label: 'Unidades', data: meses.map(m => ventasPorMes[m]), borderColor: '#ED1C24', backgroundColor: 'rgba(237,28,36,0.1)', fill: true, tension: 0.4 }]},
            options: { responsive: true, plugins: { tooltip: { callbacks: { label: (ctx) => ctx.parsed.y.toLocaleString() + ' u.' } } } }
        });

        // 2. Composición de Costos (Dona)
        const insumos = ['Total_Concentrado_($)','Total_Fructosa_($)','Total_Gas_Carbónico_($)','Total_Tapas_($)'];
        const labels = ['Concentrado','Fructosa','Gas','Tapas'];
        const dataComp = insumos.map((col, i) => data.reduce((acc, o) => acc + (o[col] || 0), 0));
        const ctx3 = document.getElementById('chartComposicion').getContext('2d');
        AppState.chartInstances.composicion = new Chart(ctx3, {
            type: 'doughnut',
            data: { labels, datasets: [{ data: dataComp, backgroundColor: ['#ED1C24', '#FF6B6B', '#FFC300', '#3498DB'] }] },
            options: { cutout: '60%', plugins: { legend: { position: 'bottom' } } }
        });
    },

    updateTable(data) {
        // Optimización: En lugar de una tabla gigante, mostramos los Top 10 productos por Facturación.
        const sorted = [...data].sort((a, b) => (b.Facturación_Neta || 0) - (a.Facturación_Neta || 0)).slice(0, 10);
        const cols = ['FECHA','Nombre','Físicos','Facturación_Neta','TOTAL_INSUMOS','LOCACION___SAP'];
        
        let thead = '<tr>' + cols.map(c => `<th>${c.replace(/_/g,' ')}</th>`).join('') + '</tr>';
        document.getElementById('tableHead').innerHTML = thead;
        
        let tbody = '';
        sorted.forEach(o => {
            tbody += `<tr>` + cols.map(c => `<td>${typeof o[c] === 'number' ? o[c].toLocaleString('es-AR', {minimumFractionDigits: 2}) : o[c] || '-'}</td>`).join('') + `</tr>`;
        });
        document.getElementById('tableBody').innerHTML = tbody;
        document.getElementById('registrosCount').textContent = `${data.length} registros totales (Top 10 por facturación)`;
    }
};

// ==========================================
// CONTROLADOR PRINCIPAL (Inicialización)
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    const processBtn = document.getElementById('processDataBtn');
    const fileInput = document.querySelector('.upload-box input[type="file"]');
    const filters = { selectMes: document.getElementById('selectMes'), selectUbicacion: document.getElementById('selectUbicacion') };

    // Carga unificada: El usuario sube un solo Excel
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const status = document.querySelector('.upload-box .file-status-msg');
        status.textContent = '⏳ Procesando archivo...';
        
        try {
            const raw = await FileService.handleUpload(file);
            AppState.rawData = {
                ventas: DataProcessor.processData(raw.ventas),
                recetas: DataProcessor.processData(raw.recetas),
                costos: DataProcessor.processData(raw.costos)
            };
            
            // Procesar automáticamente sin botón adicional
            const datos = DataProcessor.recalculateCosts();
            AppState.filteredData = datos;
            
            status.textContent = `✅ ${AppState.rawData.ventas.length} registros cargados`;
            document.getElementById('filtrosSection').style.display = 'flex';
            document.getElementById('resumenSection').style.display = 'grid';
            
            UIRenderer.updateKPIs(datos);
            UIRenderer.updateCharts(datos);
            UIRenderer.updateTable(datos);
            
        } catch (err) {
            alert('Error al leer el Excel: ' + err.message);
            status.textContent = '❌ Error al procesar';
        }
    });

    // Filtros con Debounce (Para no recargar el DOM en cada tecla)
    const applyFilters = Utils.debounce(() => {
        const mes = filters.selectMes.value;
        const ubic = filters.selectUbicacion.value;
        
        const datos = AppState.filteredData.filter(o => {
            if (mes !== 'todos' && o._mes !== mes) return false;
            if (ubic !== 'todas' && o.LOCACION___SAP !== ubic) return false;
            return true;
        });
        
        UIRenderer.updateKPIs(datos);
        UIRenderer.updateCharts(datos);
        UIRenderer.updateTable(datos);
    }, 300); // 300ms de espera

    filters.selectMes.addEventListener('change', applyFilters);
    filters.selectUbicacion.addEventListener('change', applyFilters);

    // Modo Oscuro
    document.getElementById('darkModeToggle').addEventListener('click', () => {
        document.body.classList.toggle('dark-mode');
        // Re-renderizar gráficos para ajustar colores
        if (AppState.filteredData.length) UIRenderer.updateCharts(AppState.filteredData);
    });
});
