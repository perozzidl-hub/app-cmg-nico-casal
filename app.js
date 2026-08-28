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

// Convierte cualquier formato de fecha/mes a MMM-AA
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

    // Clasificación por nombre del archivo
    if (nameLower.includes('precio')) {
        DB.precios = data;
    } else if (nameLower.includes('receta')) {
        DB.recetas = data;
    } else if (nameLower.includes('venta')) {
        DB.ventas = data;
    } else {
        DB.teoricos = data;
    }
}

function populateFilterOptions() {
    const selectMes = document.getElementById('select-mes');
    const selectArticulo = document.getElementById('select-articulo');

    // 1. OBTENER MESES DIRECTAMENTE DE LA COLUMNA "FECHA" DE VENTAS
    let mesesSet = new Set();
    DB.ventas.forEach(row => {
        // Busca la clave de la columna sin importar espacios o mayúsculas
        const keyFecha = Object.keys(row).find(k => k.trim().toUpperCase() === 'FECHA');
        if (keyFecha && row[keyFecha]) {
            const mesFormatted = formatToMesAnio(row[keyFecha]);
            if (mesFormatted) mesesSet.add(mesFormatted);
        }
    });

    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    mesesSet.forEach(m => {
        selectMes.innerHTML += `<option value="${m}">${m}</option>`;
    });
    selectMes.disabled = false;

    // 2. OBTENER ARTÍCULOS DIRECTAMENTE DE LA COLUMNA "NOMBRE" DE VENTAS
    let articulosMap = new Map(); // Guarda CÓDIGO -> NOMBRE

    DB.ventas.forEach(row => {
        const keyNombre = Object.keys(row).find(k => k.trim().toUpperCase() === 'NOMBRE');
        const keyCod = Object.keys(row).find(k => {
            const kUpper = k.trim().toUpperCase();
            return kUpper.includes('COD') || kUpper.includes('ARTICULO') || kUpper === 'ART';
        });

        if (keyNombre && row[keyNombre]) {
            const nombreArt = String(row[keyNombre]).trim();
            const codArt = keyCod ? String(row[keyCod]).trim() : '';

            if (codArt && !articulosMap.has(codArt)) {
                articulosMap.set(codArt, `${codArt} - ${nombreArt}`);
            } else if (!codArt && !articulosMap.has(nombreArt)) {
                articulosMap.set(nombreArt, nombreArt);
            }
        }
    });

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulosMap.forEach((textoDisplay, valor) => {
        selectArticulo.innerHTML += `<option value="${valor}">${textoDisplay}</option>`;
    });
    selectArticulo.disabled = false;
}

function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!artSel) return;

    // Buscar Receta asociando por código o por nombre
    const recetaArticulo = DB.recetas.filter(r => {
        return Object.keys(r).some(k => {
            const val = String(r[k]).trim();
            return val === artSel;
        });
    });

    let costoUnitarioTotal = 0;
    let desgloseInsumos = [];

    recetaArticulo.forEach(item => {
        let codInsumo = '';
        let cantTeorica = 0;
        let descInsumo = '';

        Object.keys(item).forEach(k => {
            const kUpper = k.trim().toUpperCase();
            if (kUpper.includes('INSUMO') && (kUpper.includes('COD') || kUpper.includes('CÓDIGO'))) {
                codInsumo = String(item[k]).trim();
            }
            if (kUpper.includes('CANT') || kUpper.includes('TEORICA') || kUpper.includes('TEÓRICA')) {
                cantTeorica = parseFloat(item[k] || 0);
            }
            if (kUpper.includes('DESC') || kUpper.includes('NOMBRE')) {
                descInsumo = String(item[k]).trim();
            }
        });

        // Buscar Precio
        const precioItem = DB.precios.find(p => {
            return Object.keys(p).some(k => String(p[k]).trim() === codInsumo);
        });

        let precioCompra = 0;

        if (precioItem) {
            Object.keys(precioItem).forEach(k => {
                const kUpper = k.trim().toUpperCase();
                if (kUpper.includes('PRECIO') || kUpper.includes('COMPRA') || kUpper.includes('COSTO')) {
                    precioCompra = parseFloat(precioItem[k] || 0);
                }
                if (!descInsumo && (kUpper.includes('DESC') || kUpper.includes('NOMBRE'))) {
                    descInsumo = String(precioItem[k]).trim();
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

    // Unidades vendidas desde la base de Ventas
    let cantVendida = 0;
    if (mesSel) {
        const ventaItem = DB.ventas.find(v => {
            const keyFecha = Object.keys(v).find(k => k.trim().toUpperCase() === 'FECHA');
            const matchMes = keyFecha && formatToMesAnio(v[keyFecha]) === mesSel;

            const matchArt = Object.values(v).some(val => String(val).trim() === artSel);

            return matchMes && matchArt;
        });

        if (ventaItem) {
            Object.keys(ventaItem).forEach(k => {
                const kUpper = k.trim().toUpperCase();
                if (kUpper.includes('CANT') || kUpper.includes('VOLUMEN') || kUpper.includes('UNIDADES') || kUpper.includes('VENTA')) {
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
