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

// Función auxiliar para encontrar la primera clave que coincida con un patrón
function findKey(obj, patterns) {
    const keys = Object.keys(obj);
    for (let pattern of patterns) {
        const found = keys.find(k => k.trim().toUpperCase().includes(pattern.toUpperCase()));
        if (found) return found;
    }
    return null;
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

    // 1. Obtener meses únicos desde la columna "FECHA" de ventas
    let mesesSet = new Set();
    DB.ventas.forEach(row => {
        const keyFecha = findKey(row, ['FECHA']);
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

    // 2. Obtener nombres de artículos desde la columna "NOMBRE" de ventas
    let articulosSet = new Set(); // Usamos un Set para evitar duplicados
    DB.ventas.forEach(row => {
        const keyNombre = findKey(row, ['NOMBRE', 'ARTICULO', 'PRODUCTO']);
        if (keyNombre && row[keyNombre]) {
            const nombreArt = String(row[keyNombre]).trim();
            if (nombreArt) articulosSet.add(nombreArt);
        }
    });

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    // Opcional: si además quieres mostrar el código, podrías obtenerlo, pero el valor será el nombre
    articulosSet.forEach(nombre => {
        selectArticulo.innerHTML += `<option value="${nombre}">${nombre}</option>`;
    });
    selectArticulo.disabled = false;
}

function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value; // Este es el NOMBRE del artículo

    if (!artSel) {
        // Limpiar KPIs y tabla si no hay artículo seleccionado
        document.getElementById('kpi-ventas').innerText = '0';
        document.getElementById('kpi-costo-unitario').innerText = '$ 0.00';
        document.getElementById('kpi-costo-total').innerText = '$ 0.00';
        document.querySelector('#table-receta tbody').innerHTML = '<tr><td colspan="6" class="empty-msg">Seleccione un artículo para ver su receta.</td></tr>';
        if (insumosChart) insumosChart.destroy();
        return;
    }

    // ---- 1. Calcular cantidad vendida en el mes para el artículo ----
    let cantVendida = 0;
    if (mesSel) {
        DB.ventas.forEach(row => {
            const keyFecha = findKey(row, ['FECHA']);
            const keyNombre = findKey(row, ['NOMBRE', 'ARTICULO', 'PRODUCTO']);
            const keyCant = findKey(row, ['CANTIDAD', 'UNIDADES', 'VOLUMEN', 'VENTA']);

            if (keyFecha && keyNombre && keyCant) {
                const mesFila = formatToMesAnio(row[keyFecha]);
                const nombreFila = String(row[keyNombre]).trim();
                if (mesFila === mesSel && nombreFila === artSel) {
                    const cant = parseFloat(row[keyCant]) || 0;
                    cantVendida += cant;
                }
            }
        });
    }

    // ---- 2. Buscar la receta del artículo ----
    // Buscamos en DB.recetas todas las filas cuyo nombre de artículo coincida con artSel
    // La columna que contiene el nombre del artículo en la tabla de recetas puede variar.
    // Usamos findKey para localizarla.
    let recetasFiltradas = [];
    if (DB.recetas.length > 0) {
        // Primero, determinar cuál es la columna que contiene el nombre del artículo en recetas
        // Tomamos la primera fila como referencia
        const sampleRow = DB.recetas[0];
        const keyArtReceta = findKey(sampleRow, ['ARTICULO', 'PRODUCTO', 'NOMBRE']);
        if (keyArtReceta) {
            recetasFiltradas = DB.recetas.filter(row => {
                const nombre = String(row[keyArtReceta]).trim();
                return nombre === artSel;
            });
        } else {
            // Si no encontramos columna de artículo, fallback: buscar en cualquier columna (no recomendado)
            recetasFiltradas = DB.recetas.filter(row => {
                return Object.values(row).some(val => String(val).trim() === artSel);
            });
        }
    }

    if (recetasFiltradas.length === 0) {
        document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
        document.getElementById('kpi-costo-unitario').innerText = '$ 0.00';
        document.getElementById('kpi-costo-total').innerText = '$ 0.00';
        document.querySelector('#table-receta tbody').innerHTML = '<tr><td colspan="6" class="empty-msg">No se encontró receta para este artículo.</td></tr>';
        if (insumosChart) insumosChart.destroy();
        return;
    }

    // ---- 3. Procesar cada fila de la receta (insumos) ----
    let desgloseInsumos = [];
    let costoUnitarioTotal = 0;

    recetasFiltradas.forEach(item => {
        // Extraer código de insumo, cantidad teórica, descripción
        const keyCodInsumo = findKey(item, ['CODIGO INSUMO', 'COD INSUMO', 'CODIGO', 'INSUMO']);
        const keyCantTeorica = findKey(item, ['CANTIDAD', 'TEORICA', 'CANT']);
        const keyDescInsumo = findKey(item, ['DESCRIPCION', 'NOMBRE INSUMO', 'NOMBRE']);

        let codInsumo = keyCodInsumo ? String(item[keyCodInsumo]).trim() : '';
        let cantTeorica = keyCantTeorica ? parseFloat(item[keyCantTeorica]) || 0 : 0;
        let descInsumo = keyDescInsumo ? String(item[keyDescInsumo]).trim() : '';

        // Si no se encontró descripción, usar el código
        if (!descInsumo) descInsumo = 'Insumo ' + codInsumo;

        // ---- 4. Buscar precio de compra para ese insumo ----
        let precioCompra = 0;
        if (codInsumo) {
            // Buscar en DB.precios una fila cuyo código coincida
            for (let pRow of DB.precios) {
                const keyCodPrecio = findKey(pRow, ['CODIGO', 'INSUMO', 'COD']);
                if (keyCodPrecio && String(pRow[keyCodPrecio]).trim() === codInsumo) {
                    // Encontrar la columna de precio
                    const keyPrecio = findKey(pRow, ['PRECIO', 'COSTO', 'COMPRA']);
                    if (keyPrecio) {
                        precioCompra = parseFloat(pRow[keyPrecio]) || 0;
                    }
                    break; // asumimos que solo hay una fila por código
                }
            }
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

    // ---- 5. Calcular costo total de venta ----
    const costoVentaTotal = costoUnitarioTotal * cantVendida;

    // ---- 6. Actualizar KPIs ----
    document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
    document.getElementById('kpi-costo-unitario').innerText = '$ ' + costoUnitarioTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('kpi-costo-total').innerText = '$ ' + costoVentaTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});

    // ---- 7. Renderizar tabla y gráfico ----
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
    if (insumosConCosto.length === 0) {
        insumosChart = null;
        return;
    }

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
