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

// Función de formateo estricto a MES-AÑO (ej: sep-25)
function formatToMesAnio(value) {
    if (!value) return '';
    const meses = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
    
    let date;
    if (value instanceof Date) {
        date = value;
    } else if (typeof value === 'number' && value > 30000) {
        date = new Date(Math.round((value - 25569) * 86400 * 1000));
    } else {
        const parsed = Date.parse(value);
        if (!isNaN(parsed)) {
            date = new Date(parsed);
        } else {
            return String(value).trim();
        }
    }
    
    const mes = meses[date.getUTCMonth()];
    const anio = String(date.getUTCFullYear()).slice(-2);
    return `${mes}-${anio}`;
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

function classifyAndStoreData(fileName, data) {
    if (!data.length) return;

    const sampleKeys = Object.keys(data[0]).map(k => k.toLowerCase());

    if (sampleKeys.some(k => k.includes('precio') || k.includes('compra'))) {
        DB.precios = data;
    } else if (sampleKeys.some(k => k.includes('teorica') || k.includes('teórica') || k.includes('cant'))) {
        DB.recetas = data;
    } else if (sampleKeys.some(k => k.includes('venta') || k.includes('factura') || k.includes('volumen'))) {
        // Formatear mes de ventas al momento de guardar
        DB.ventas = data.map(row => {
            let newRow = { ...row };
            Object.keys(row).forEach(key => {
                if (key.toLowerCase().includes('mes') || key.toLowerCase().includes('fecha')) {
                    newRow['MesVentaFormateado'] = formatToMesAnio(row[key]);
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

    // 1. Filtro de Mes proveniente EXCLUSIVAMENTE del archivo de Ventas
    const mesesVentas = [...new Set(DB.ventas.map(i => i.MesVentaFormateado))].filter(Boolean);
    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    mesesVentas.forEach(m => selectMes.innerHTML += `<option value="${m}">${m}</option>`);
    selectMes.disabled = false;

    // 2. Filtro de Artículo con NOMBRE / DESCRIPCIÓN legible
    let articulosMap = new Map();

    // Extraer nombres desde Ventas o Receta
    const fuenteArticulos = DB.ventas.length ? DB.ventas : DB.recetas;
    fuenteArticulos.forEach(row => {
        const codigo = String(row['Cod. Venta'] || row.Cod_Venta || row.Codigo || row.Articulo || '').trim();
        const descripcion = row['Descripción'] || row.Descripcion || row.Nombre || row.Articulo_Nombre || row.Articulo || codigo;
        
        if (codigo && !articulosMap.has(codigo)) {
            articulosMap.set(codigo, descripcion !== codigo ? `${codigo} - ${descripcion}` : codigo);
        }
    });

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulosMap.forEach((nombreMostrar, codigo) => {
        selectArticulo.innerHTML += `<option value="${codigo}">${nombreMostrar}</option>`;
    });
    selectArticulo.disabled = false;
}

function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!artSel) return;

    // 1. Receta del Artículo (Sin filtrar por mes)
    const recetaArticulo = DB.recetas.filter(r => 
        String(r['Cod. Venta'] || r.Cod_Venta || r.Articulo || '').trim() === artSel
    );

    let costoUnitarioTotal = 0;
    let desgloseInsumos = [];

    // 2. Obtener costo de los insumos (Sin filtrar por mes)
    recetaArticulo.forEach(item => {
        const codInsumo = String(item['Código Insumo'] || item.Codigo_Insumo || item.Cod_Insumo || '').trim();
        const cantTeorica = parseFloat(item['Cant. Teorica'] || item['Cant. Teórica'] || item.Cant_Teorica || 0);

        // Buscar el precio de compra del insumo
        const precioItem = DB.precios.find(p => 
            String(p['Código Insumo'] || p.Codigo_Insumo || p.Cod_Insumo || '').trim() === codInsumo
        );

        const precioCompra = precioItem ? parseFloat(precioItem['Precio Compra'] || precioItem.Precio_Compra || 0) : 0;
        const subtotal = cantTeorica * precioCompra;
        costoUnitarioTotal += subtotal;

        desgloseInsumos.push({
            codInsumo,
            descripcion: precioItem ? (precioItem.Descripción || precioItem.Descripcion) : (item.Descripción || 'Insumo ' + codInsumo),
            cantTeorica,
            precioCompra,
            subtotal
        });
    });

    // 3. Obtener Ventas asociadas al Artículo y al Mes Seleccionado en Ventas
    let cantVendida = 0;
    if (mesSel) {
        const ventaItem = DB.ventas.find(v => 
            v.MesVentaFormateado === mesSel && 
            String(v['Cod. Venta'] || v.Cod_Venta || v.Articulo || '').trim() === artSel
        );
        cantVendida = ventaItem ? parseFloat(ventaItem.Cantidad || ventaItem.cantidad_vendida || ventaItem.Volumen || 0) : 0;
    }

    const costoVentaTotal = costoUnitarioTotal * cantVendida;

    // 4. Actualizar KPIs en pantalla
    document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
    document.getElementById('kpi-costo-unitario').innerText = '$ ' + costoUnitarioTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('kpi-costo-total').innerText = '$ ' + costoVentaTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});

    // 5. Renderizar vista detallada y gráfico
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
