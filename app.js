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

function handleFileUpload(event) {
    const files = event.target.files;
    if (!files.length) return;

    let loadedCount = 0;

    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
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
    const keys = Object.keys(data[0]).map(k => k.toLowerCase());

    if (keys.some(k => k.includes('insumo') && k.includes('precio'))) {
        DB.precios = data;
    } else if (keys.some(k => k.includes('cant') && k.includes('teorica'))) {
        DB.recetas = data;
    } else if (keys.some(k => k.includes('venta') || k.includes('factura'))) {
        DB.ventas = data;
    } else {
        DB.teoricos = data;
    }
}

// Carga de opciones en los selectores de filtro
function populateFilterOptions() {
    const selectMes = document.getElementById('select-mes');
    const selectArticulo = document.getElementById('select-articulo');

    // Extraer meses únicos de recetas o precios
    const meses = [...new Set(DB.recetas.map(item => item.Mes || item.mes))].filter(Boolean);
    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    meses.forEach(m => selectMes.innerHTML += `<option value="${m}">${m}</option>`);
    selectMes.disabled = false;

    // Extraer artículos únicos
    const articulos = [...new Set(DB.recetas.map(item => item.Articulo || item['Cod. Venta'] || item.Cod_Venta))].filter(Boolean);
    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulos.forEach(a => selectArticulo.innerHTML += `<option value="${a}">${a}</option>`);
    selectArticulo.disabled = false;
}

// Ejecución del cálculo relacional al cambiar filtros
function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!mesSel || !artSel) return;

    // 1. Filtrar la Receta del Artículo y Mes seleccionado
    const recetaArticulo = DB.recetas.filter(r => 
        (r.Mes === mesSel || r.mes === mesSel) && 
        (r.Articulo === artSel || r['Cod. Venta'] == artSel || r.Cod_Venta == artSel)
    );

    // 2. Cruzar cada Insumo con la tabla de Precios por Código Insumo y Mes
    let costoUnitarioTotal = 0;
    let desgloseInsumos = [];

    recetaArticulo.forEach(item => {
        const codInsumo = item['Código Insumo'] || item.Codigo_Insumo || item.cod_insumo;
        const cantTeorica = parseFloat(item['Cant. Teorica'] || item.Cant_Teorica || item.cantidad || 0);

        // Buscar Precio Compra
        const precioItem = DB.precios.find(p => 
            (p.Mes === mesSel || p.mes === mesSel) && 
            (p['Código Insumo'] == codInsumo || p.Codigo_Insumo == codInsumo)
        );

        const precioCompra = precioItem ? parseFloat(precioItem['Precio Compra'] || precioItem.Precio_Compra || 0) : 0;
        const subtotal = cantTeorica * precioCompra;
        costoUnitarioTotal += subtotal;

        desgloseInsumos.push({
            codInsumo,
            descripcion: precioItem ? (precioItem.Descripción || precioItem.Descripcion) : 'Insumo ' + codInsumo,
            cantTeorica,
            precioCompra,
            subtotal
        });
    });

    // 3. Obtener Ventas del Artículo
    const ventaItem = DB.ventas.find(v => 
        (v.Mes === mesSel || v.mes === mesSel) && 
        (v['Cod. Venta'] == artSel || v.Articulo === artSel)
    );
    const cantVendida = ventaItem ? parseFloat(ventaItem.Cantidad || ventaItem.cantidad_vendida || 1000) : 1000;
    const costoVentaTotal = costoUnitarioTotal * cantVendida;

    // 4. Actualizar KPIs en la pantalla
    document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString();
    document.getElementById('kpi-costo-unitario').innerText = '$ ' + costoUnitarioTotal.toLocaleString('es-AR', {minimumFractionDigits: 2});
    document.getElementById('kpi-costo-total').innerText = '$ ' + costoVentaTotal.toLocaleString('es-AR', {minimumFractionDigits: 2});

    // 5. Renderizar Tabla y Gráfico
    renderTable(desgloseInsumos, costoUnitarioTotal);
    renderChart(desgloseInsumos);
}

function renderTable(insumos, costoTotal) {
    const tbody = document.querySelector('#table-receta tbody');
    tbody.innerHTML = '';

    insumos.forEach(i => {
        const part = costoTotal > 0 ? ((i.subtotal / costoTotal) * 100).toFixed(1) : 0;
        tbody.innerHTML += `
            <tr>
                <td><b>${i.codInsumo}</b></td>
                <td>${i.descripcion}</td>
                <td>${i.cantTeorica}</td>
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

    insumosChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: insumos.map(i => i.descripcion),
            datasets: [{
                data: insumos.map(i => i.subtotal),
                backgroundColor: ['#F40009', '#111111', '#555555', '#999999', '#DDDDDD', '#AA0000']
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
