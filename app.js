// Contenedores globales de bases de datos
let DB = {
    ventas: [],
    recetas: [],
    precios: [],
    teoricos: []
};

let insumosChart = null;

// Escuchar la carga de archivos Excel / CSV
document.getElementById('file-input').addEventListener('change', handleFileUpload);
document.getElementById('select-mes').addEventListener('change', applyFilters);
document.getElementById('select-articulo').addEventListener('change', applyFilters);

// Convertir fechas de Excel (ej: 45901) a formato mes-año (ej: sept-25)
function parseExcelDate(value) {
    if (!value) return '';
    if (typeof value === 'number' && value > 30000) {
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        const meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
        const mes = meses[date.getUTCMonth()];
        const anio = String(date.getUTCFullYear()).slice(-2);
        return `${mes}-${anio}`;
    }
    return String(value).trim();
}

function handleFileUpload(event) {
    const files = event.target.files;
    if (!files.length) return;

    let loadedCount = 0;

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

// Clasificación automática de la tabla según las columnas
function classifyAndStoreData(fileName, data) {
    if (!data.length) return;

    // Normalizar nombres de columnas para evitar fallos por espacios
    const normalizedData = data.map(row => {
        let newRow = {};
        Object.keys(row).forEach(key => {
            const cleanKey = key.trim();
            if (cleanKey.toLowerCase().includes('mes')) {
                newRow['MesClean'] = parseExcelDate(row[key]);
            } else {
                newRow[cleanKey] = row[key];
            }
        });
        return newRow;
    });

    const sampleKeys = Object.keys(normalizedData[0]).map(k => k.toLowerCase());

    if (sampleKeys.some(k => k.includes('precio') || k.includes('compra'))) {
        DB.precios = normalizedData;
    } else if (sampleKeys.some(k => k.includes('teorica') || k.includes('teórica') || k.includes('cant'))) {
        DB.recetas = normalizedData;
    } else if (sampleKeys.some(k => k.includes('venta') || k.includes('factura') || k.includes('volumen'))) {
        DB.ventas = normalizedData;
    } else {
        DB.teoricos = normalizedData;
    }
}

// Carga de opciones en los selectores de filtro
function populateFilterOptions() {
    const selectMes = document.getElementById('select-mes');
    const selectArticulo = document.getElementById('select-articulo');

    // Extraer meses únicos normalizados
    const meses = [...new Set([
        ...DB.recetas.map(i => i.MesClean),
        ...DB.precios.map(i => i.MesClean),
        ...DB.ventas.map(i => i.MesClean)
    ])].filter(Boolean);

    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    meses.forEach(m => selectMes.innerHTML += `<option value="${m}">${m}</option>`);
    selectMes.disabled = false;

    // Extraer artículos únicos por código o nombre
    const articulos = [...new Set([
        ...DB.recetas.map(i => String(i['Cod. Venta'] || i.Cod_Venta || i.Articulo || '').trim()),
        ...DB.ventas.map(i => String(i['Cod. Venta'] || i.Cod_Venta || i.Articulo || '').trim())
    ])].filter(Boolean);

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulos.forEach(a => selectArticulo.innerHTML += `<option value="${a}">${a}</option>`);
    selectArticulo.disabled = false;
}

// Ejecución del cálculo relacional
function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!mesSel || !artSel) return;

    // 1. Filtrar Receta
    const recetaArticulo = DB.recetas.filter(r => 
        r.MesClean === mesSel && 
        String(r['Cod. Venta'] || r.Cod_Venta || r.Articulo || '').trim() === artSel
    );

    let costoUnitarioTotal = 0;
    let desgloseInsumos = [];

    // 2. Cruzar con Precios de Insumos
    recetaArticulo.forEach(item => {
        const codInsumo = String(item['Código Insumo'] || item.Codigo_Insumo || item.Cod_Insumo || '').trim();
        const cantTeorica = parseFloat(item['Cant. Teorica'] || item['Cant. Teórica'] || item.Cant_Teorica || 0);

        const precioItem = DB.precios.find(p => 
            p.MesClean === mesSel && 
            String(p['Código Insumo'] || p.Codigo_Insumo || p.Cod_Insumo || '').trim() === codInsumo
        );

        const precioCompra = precioItem ? parseFloat(precioItem['Precio Compra'] || precioItem.Precio_Compra || 0) : 0;
        const subtotal = cantTeorica * precioCompra;
        costoUnitarioTotal += subtotal;

        desgloseInsumos.push({
            codInsumo,
            descripcion: precioItem ? (precioItem.Descripción || precioItem.Descripcion || 'Insumo ' + codInsumo) : 'Insumo ' + codInsumo,
            cantTeorica,
            precioCompra,
            subtotal
        });
    });

    // 3. Obtener Ventas
    const ventaItem = DB.ventas.find(v => 
        v.MesClean === mesSel && 
        String(v['Cod. Venta'] || v.Cod_Venta || v.Articulo || '').trim() === artSel
    );

    const cantVendida = ventaItem ? parseFloat(ventaItem.Cantidad || ventaItem.cantidad_vendida || ventaItem.Volumen || 0) : 0;
    const costoVentaTotal = costoUnitarioTotal * cantVendida;

    // 4. Actualizar KPIs
    document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
    document.getElementById('kpi-costo-unitario').innerText = '$ ' + costoUnitarioTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('kpi-costo-total').innerText = '$ ' + costoVentaTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});

    // 5. Renderizar Vistas
    renderTable(desgloseInsumos, costoUnitarioTotal);
    renderChart(desgloseInsumos);
}

function renderTable(insumos, costoTotal) {
    const tbody = document.querySelector('#table-receta tbody');
    tbody.innerHTML = '';

    if (insumos.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-msg">No se encontraron insumos para este artículo y mes.</td></tr>';
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
                backgroundColor: ['#F40009', '#111111', '#555555', '#999999', '#DDDDDD', '#AA0000', '#FF4D4D']
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
