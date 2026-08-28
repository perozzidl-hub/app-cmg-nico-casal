// Contenedores globales de bases de datos
let DB = {
    ventas: [],
    recetas: [],
    precios: [],
    teoricos: []
};

let insumosChart = null;

// Escuchar eventos de la interfaz
document.getElementById('file-input').addEventListener('change', handleFileUpload);
document.getElementById('select-mes').addEventListener('change', applyFilters);
document.getElementById('select-articulo').addEventListener('change', applyFilters);

// Convierte cualquier formato de Fecha/Excel a "mmm-aa" (ej: sep-25)
function formatToMesAnio(value) {
    if (!value) return '';
    const meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
    
    if (value instanceof Date && !isNaN(value)) {
        return `${meses[value.getUTCMonth()]}-${String(value.getUTCFullYear()).slice(-2)}`;
    }
    
    if (typeof value === 'number' && value > 30000) {
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        return `${meses[date.getUTCMonth()]}-${String(date.getUTCFullYear()).slice(-2)}`;
    }
    
    const str = String(value).trim();
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) {
        const date = new Date(parsed);
        return `${meses[date.getUTCMonth()]}-${String(date.getUTCFullYear()).slice(-2)}`;
    }
    
    return str;
}

function handleFileUpload(event) {
    const files = event.target.files;
    if (!files.length) return;

    let loadedCount = 0;
    DB = { ventas: [], recetas: [], precios: [], teoricos: [] };

    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheet = workbook.SheetNames[0];
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);

            classifyAndStoreData(file.name, jsonData);
            loadedCount++;

            if (loadedCount === files.length) {
                document.getElementById('data-status').innerText = '✅ Bases cargadas correctamente';
                populateFilterOptions();
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function classifyAndStoreData(fileName, data) {
    if (!data.length) return;

    const nameLower = fileName.toLowerCase();
    const sampleKeys = Object.keys(data[0]).map(k => k.toLowerCase().trim());

    if (nameLower.includes('precio') || sampleKeys.some(k => k.includes('precio') || k.includes('compra'))) {
        DB.precios = data;
    } else if (nameLower.includes('receta') || sampleKeys.some(k => k.includes('teorica') || k.includes('teórica') || k.includes('cant'))) {
        DB.recetas = data;
    } else if (nameLower.includes('venta') || sampleKeys.some(k => k === 'fecha' || k.includes('venta') || k.includes('volumen'))) {
        // Mapear la columna FECHA explícitamente
        DB.ventas = data.map(row => {
            let newRow = { ...row };
            Object.keys(row).forEach(key => {
                const kClean = key.toLowerCase().trim();
                if (kClean === 'fecha' || kClean.includes('mes') || kClean.includes('periodo')) {
                    newRow['MesVentaClean'] = formatToMesAnio(row[key]);
                }
            });
            return newRow;
        });
    } else {
        DB.teoricos = data;
    }
}

function populateFilterOptions() {
    const selectMes = document.getElementById('select-mes');
    const selectArticulo = document.getElementById('select-articulo');

    // 1. Obtener los meses únicos leyendo la columna FECHA procesada de Ventas
    const mesesVentas = [...new Set(DB.ventas.map(i => i.MesVentaClean))].filter(Boolean);
    
    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    mesesVentas.forEach(m => {
        selectMes.innerHTML += `<option value="${m}">${m}</option>`;
    });
    selectMes.disabled = false;

    // 2. Obtener Artículos mostrando "CÓDIGO - DESCRIPCIÓN"
    let articulosMap = new Map();

    // Cruzamos la Receta y Ventas para extraer código y nombre del producto
    const listaCombinada = [...DB.recetas, ...DB.ventas];
    listaCombinada.forEach(row => {
        let codigo = '';
        let descripcion = '';

        Object.keys(row).forEach(k => {
            const kClean = k.toLowerCase().trim();
            const val = String(row[k] || '').trim();

            // Detectar Código de Artículo
            if ((kClean.includes('cod') || kClean.includes('articulo') || kClean === 'art') && !kClean.includes('insumo') && !kClean.includes('desc')) {
                if (!codigo && val) codigo = val;
            }
            // Detectar Descripción de Artículo
            if (kClean.includes('descrip') || kClean.includes('nombre') || kClean.includes('producto')) {
                if (!kClean.includes('insumo') && !descripcion && val) descripcion = val;
            }
        });

        if (codigo && !articulosMap.has(codigo)) {
            const textoMostrar = descripcion ? `${codigo} - ${descripcion}` : codigo;
            articulosMap.set(codigo, textoMostrar);
        }
    });

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulosMap.forEach((texto, codigo) => {
        selectArticulo.innerHTML += `<option value="${codigo}">${texto}</option>`;
    });
    selectArticulo.disabled = false;
}

function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!artSel) return;

    // 1. Buscar Receta del Artículo (Sin filtrar por mes)
    const recetaArticulo = DB.recetas.filter(r => {
        return Object.keys(r).some(k => {
            const kClean = k.toLowerCase().trim();
            return (kClean.includes('cod') || kClean.includes('articulo')) && !kClean.includes('insumo') && String(r[k]).trim() === artSel;
        });
    });

    let costoUnitarioTotal = 0;
    let desgloseInsumos = [];

    // 2. Cruzar con Precios de Insumos
    recetaArticulo.forEach(item => {
        let codInsumo = '';
        let cantTeorica = 0;

        Object.keys(item).forEach(k => {
            const kClean = k.toLowerCase().trim();
            if (kClean.includes('insumo') && (kClean.includes('cod') || kClean.includes('código'))) {
                codInsumo = String(item[k]).trim();
            }
            if (kClean.includes('cant') || kClean.includes('teorica') || kClean.includes('teórica')) {
                cantTeorica = parseFloat(item[k] || 0);
            }
        });

        // Buscar precio del insumo en la tabla Precios
        const precioItem = DB.precios.find(p => {
            return Object.keys(p).some(k => {
                const kClean = k.toLowerCase().trim();
                return kClean.includes('insumo') && String(p[k]).trim() === codInsumo;
            });
        });

        let precioCompra = 0;
        let descInsumo = 'Insumo ' + codInsumo;

        if (precioItem) {
            Object.keys(precioItem).forEach(k => {
                const kClean = k.toLowerCase().trim();
                if (kClean.includes('precio') || kClean.includes('compra') || kClean.includes('costo')) {
                    precioCompra = parseFloat(precioItem[k] || 0);
                }
                if (kClean.includes('descrip') || kClean.includes('nombre')) {
                    if (isNaN(precioItem[k])) descInsumo = precioItem[k];
                }
            });
        }

        const subtotal = cantTeorica * precioCompra;
        costoUnitarioTotal += subtotal;

        desgloseInsumos.push({
            codInsumo,
            descripcion: descInsumo,
            cantTeorica,
            precioCompra,
            subtotal
        });
    });

    // 3. Buscar Unidades Vendidas filtradas por el MES de la columna FECHA y por ARTÍCULO
    let cantVendida = 0;
    if (mesSel) {
        const ventaItem = DB.ventas.find(v => {
            const matchMes = v.MesVentaClean === mesSel;
            const matchArt = Object.keys(v).some(k => {
                const kClean = k.toLowerCase().trim();
                return (kClean.includes('cod') || kClean.includes('articulo')) && !kClean.includes('insumo') && String(v[k]).trim() === artSel;
            });
            return matchMes && matchArt;
        });

        if (ventaItem) {
            Object.keys(ventaItem).forEach(k => {
                const kClean = k.toLowerCase().trim();
                if (kClean.includes('cant') || kClean.includes('volumen') || kClean.includes('unidades') || kClean.includes('venta')) {
                    if (!isNaN(ventaItem[k])) cantVendida = parseFloat(ventaItem[k]);
                }
            });
        }
    }

    const costoVentaTotal = costoUnitarioTotal * cantVendida;

    // Update KPIs
    document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
    document.getElementById('kpi-costo-unitario').innerText = '$ ' + costoUnitarioTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('kpi-costo-total').innerText = '$ ' + costoVentaTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});

    renderTable(desgloseInsumos, costoUnitarioTotal);
    renderChart(desgloseInsumos);
}

function renderTable(insumos, costoTotal) {
    const tbody = document.querySelector('#table-receta tbody');
    tbody.innerHTML = '';

    if (insumos.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-msg">No se encontraron insumos para este artículo.</td></tr>';
        return;
    }

    insumos.forEach(i => {
        const part = costoTotal > 0 ? ((i.subtotal / costoTotal) * 100).toFixed(1) : '0.0';
        tbody.innerHTML += `
            <tr>
                <td><b>${i.codInsumo}</b></td>
                <td>${i.descripcion}</td>
                <td>${i.cantTeorica.toLocaleString('es-AR')}</td>
                <td>$ ${i.precioCompra.toLocaleString('es-AR', {minimumFractionDigits: 2})}</td>
                <td><b>$ ${i.subtotal.toLocaleString('es-AR', {minimumFractionDigits: 2})}</b></td>
                <td>${part}%</td>
            </tr>
        `;
    });
}

function renderChart(insumos) {
    const ctx = document.getElementById('chart-insumos').getContext('2d');
    if (insumosChart) insumosChart.destroy();

    const insumosConCosto = insumos.filter(i => i.subtotal > 0);

    insumosChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: insumosConCosto.map(i => i.descripcion),
            datasets: [{
                data: insumosConCosto.map(i => i.subtotal),
                backgroundColor: ['#F40009', '#111111', '#555555', '#999999', '#DDDDDD', '#AA0000', '#FF4D4D', '#333333']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' }
            }
        }
    });
}
