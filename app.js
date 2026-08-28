let DB = {
    ventas: [],
    recetas: [],
    precios: [],
    teoricos: []
};

let insumosChart = null;

document.getElementById('file-input').addEventListener('change', handleFileUpload);
document.getElementById('select-mes').addEventListener('change', applyFilters);
document.getElementById('select-articulo').addEventListener('change', applyFilters);

// Función auxiliar para limpiar nombres de columnas
function cleanKey(key) {
    return String(key || '')
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Quita tildes
        .toLowerCase()
        .trim();
}

// Formateador de Fechas/Meses
function formatToMesAnio(value) {
    if (value === null || value === undefined || value === '') return '';
    const meses = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
    
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
    
    return str.toUpperCase();
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
            const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });

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
    const sampleKeys = Object.keys(data[0]).map(cleanKey);

    if (nameLower.includes('precio') || sampleKeys.some(k => k.includes('precio') || k.includes('compra'))) {
        DB.precios = data;
    } else if (nameLower.includes('receta') || sampleKeys.some(k => k.includes('teorica') || k.includes('cant'))) {
        DB.recetas = data;
    } else if (nameLower.includes('venta') || sampleKeys.some(k => k.includes('fecha') || k.includes('venta') || k.includes('volumen'))) {
        DB.ventas = data.map(row => {
            let newRow = { ...row };
            Object.keys(row).forEach(key => {
                const kClean = cleanKey(key);
                if (kClean.includes('fecha') || kClean.includes('mes') || kClean.includes('periodo')) {
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

    // 1. OBTENER MESES DESDE VENTAS
    let mesesVentas = [];
    DB.ventas.forEach(row => {
        if (row.MesVentaClean && !mesesVentas.includes(row.MesVentaClean)) {
            mesesVentas.push(row.MesVentaClean);
        }
        // Búsqueda profunda en la fila si la clave no hizo match
        Object.keys(row).forEach(k => {
            const kClean = cleanKey(k);
            if (kClean.includes('fecha') || kClean.includes('mes')) {
                const formatted = formatToMesAnio(row[k]);
                if (formatted && formatted.length <= 7 && formatted.includes('-') && !mesesVentas.includes(formatted)) {
                    mesesVentas.push(formatted);
                }
            }
        });
    });

    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    mesesVentas.forEach(m => {
        selectMes.innerHTML += `<option value="${m}">${m}</option>`;
    });
    selectMes.disabled = false;

    // 2. OBTENER ARTÍCULOS EXCLUSIVAMENTE DESDE LA BASE DE VENTAS
    let articulosMap = new Map();

    // Priorizamos la lectura de Ventas para obtener CÓDIGO + NOMBRE DEL PRODUCTO
    const fuenteArticulos = DB.ventas.length > 0 ? DB.ventas : DB.recetas;

    fuenteArticulos.forEach(row => {
        let codigo = '';
        let descripcion = '';

        Object.keys(row).forEach(k => {
            const kClean = cleanKey(k);
            const val = String(row[k] || '').trim();

            if (!val) return;

            // Detectar columna de código (ej: Codigo, Cod Venta, Articulo)
            if ((kClean.includes('cod') || kClean.includes('art')) && !kClean.includes('insumo') && !kClean.includes('desc') && !kClean.includes('nom')) {
                if (!codigo) codigo = val;
            }

            // Detectar columna de descripción (ej: Descripcion, Nombre, Producto)
            if ((kClean.includes('desc') || kClean.includes('nom') || kClean.includes('prod') || kClean.includes('denominacion')) && !kClean.includes('insumo')) {
                if (!descripcion) descripcion = val;
            }
        });

        if (codigo && !articulosMap.has(codigo)) {
            const textoMostrar = (descripcion && descripcion !== codigo) ? `${codigo} - ${descripcion}` : codigo;
            articulosMap.set(codigo, textoMostrar);
        }
    });

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulosMap.forEach((textoMostrar, codigo) => {
        selectArticulo.innerHTML += `<option value="${codigo}">${textoMostrar}</option>`;
    });
    selectArticulo.disabled = false;
}

function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!artSel) return;

    // Receta del Artículo
    const recetaArticulo = DB.recetas.filter(r => {
        return Object.keys(r).some(k => {
            const kClean = cleanKey(k);
            return (kClean.includes('cod') || kClean.includes('art')) && !kClean.includes('insumo') && String(r[k]).trim() === artSel;
        });
    });

    let costoUnitarioTotal = 0;
    let desgloseInsumos = [];

    recetaArticulo.forEach(item => {
        let codInsumo = '';
        let cantTeorica = 0;

        Object.keys(item).forEach(k => {
            const kClean = cleanKey(k);
            if (kClean.includes('insumo') && (kClean.includes('cod') || kClean.includes('codigo'))) {
                codInsumo = String(item[k]).trim();
            }
            if (kClean.includes('cant') || kClean.includes('teorica')) {
                cantTeorica = parseFloat(item[k] || 0);
            }
        });

        // Buscar Precio Insumo
        const precioItem = DB.precios.find(p => {
            return Object.keys(p).some(k => {
                const kClean = cleanKey(k);
                return kClean.includes('insumo') && String(p[k]).trim() === codInsumo;
            });
        });

        let precioCompra = 0;
        let descInsumo = '';

        Object.keys(item).forEach(k => {
            const kClean = cleanKey(k);
            if ((kClean.includes('desc') || kClean.includes('nom')) && kClean.includes('insumo')) {
                descInsumo = item[k];
            }
        });

        if (precioItem) {
            Object.keys(precioItem).forEach(k => {
                const kClean = cleanKey(k);
                if (kClean.includes('precio') || kClean.includes('compra') || kClean.includes('costo')) {
                    precioCompra = parseFloat(precioItem[k] || 0);
                }
                if (!descInsumo && (kClean.includes('desc') || kClean.includes('nom'))) {
                    descInsumo = precioItem[k];
                }
            });
        }

        if (!descInsumo) descInsumo = 'Insumo ' + codInsumo;

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

    // Unidades Vendidas
    let cantVendida = 0;
    if (mesSel) {
        const ventaItem = DB.ventas.find(v => {
            const matchMes = v.MesVentaClean === mesSel || Object.values(v).some(val => formatToMesAnio(val) === mesSel);
            const matchArt = Object.keys(v).some(k => {
                const kClean = cleanKey(k);
                return (kClean.includes('cod') || kClean.includes('art')) && !kClean.includes('insumo') && String(v[k]).trim() === artSel;
            });
            return matchMes && matchArt;
        });

        if (ventaItem) {
            Object.keys(ventaItem).forEach(k => {
                const kClean = cleanKey(k);
                if (kClean.includes('cant') || kClean.includes('volumen') || kClean.includes('unidades') || kClean.includes('venta')) {
                    if (!isNaN(ventaItem[k]) && ventaItem[k] !== '') cantVendida = parseFloat(ventaItem[k]);
                }
            });
        }
    }

    const costoVentaTotal = costoUnitarioTotal * cantVendida;

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

    const coloresCorporativos = [
        '#F40009', '#111111', '#555555', '#888888', '#BBBBBB', 
        '#990000', '#CC0000', '#FF4D4D', '#333333', '#D4AF37', 
        '#777777', '#E60000'
    ];

    insumosChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: insumosConCosto.map(i => i.descripcion),
            datasets: [{
                data: insumosConCosto.map(i => i.subtotal),
                backgroundColor: coloresCorporativos.slice(0, insumosConCosto.length),
                borderWidth: 2,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: { top: 10, bottom: 10 }
            },
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        boxWidth: 12,
                        padding: 10,
                        font: { size: 10, family: 'Helvetica Neue' }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw || 0;
                            return ' ' + context.label + ': $' + value.toLocaleString('es-AR', {minimumFractionDigits: 2});
                        }
                    }
                }
            }
        }
    });
}
