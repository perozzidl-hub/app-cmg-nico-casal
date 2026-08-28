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

// ========== UTILIDADES ==========
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

// Busca la primera clave que contenga alguno de los patrones (case insensitive)
function findKey(obj, patterns) {
    const keys = Object.keys(obj);
    for (let p of patterns) {
        const found = keys.find(k => k.trim().toUpperCase().includes(p.toUpperCase()));
        if (found) return found;
    }
    return null;
}

// ========== CARGA Y CLASIFICACIÓN INTELIGENTE ==========
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

            // Clasificar por contenido (más robusto que solo el nombre)
            classifyByContent(jsonData, file.name);
            loadedCount++;

            if (loadedCount === files.length) {
                // Mostrar qué tablas se cargaron
                let status = '✅ Bases cargadas: ';
                if (DB.ventas.length) status += `Ventas (${DB.ventas.length} filas) `;
                if (DB.recetas.length) status += `Recetas (${DB.recetas.length} filas) `;
                if (DB.precios.length) status += `Precios (${DB.precios.length} filas) `;
                if (DB.teoricos.length) status += `Teóricos (${DB.teoricos.length} filas) `;
                document.getElementById('data-status').innerText = status || '⚠️ No se detectaron datos válidos';

                // Depuración: mostrar columnas detectadas en consola
                console.log('Columnas detectadas:');
                if (DB.ventas.length) console.log('Ventas:', Object.keys(DB.ventas[0]));
                if (DB.recetas.length) console.log('Recetas:', Object.keys(DB.recetas[0]));
                if (DB.precios.length) console.log('Precios:', Object.keys(DB.precios[0]));
                if (DB.teoricos.length) console.log('Teóricos:', Object.keys(DB.teoricos[0]));

                populateFilterOptions();
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function classifyByContent(data, fileName) {
    if (!data.length) return;
    const firstRow = data[0];
    const keys = Object.keys(firstRow).map(k => k.trim().toUpperCase());

    // Detectar si es VENTAS: debe tener fecha y cantidad/nombre
    const hasDate = keys.some(k => /FECHA|DATE|DIA|MES/.test(k));
    const hasQty = keys.some(k => /CANT|UNIDAD|VOLUMEN|VENTA/.test(k));
    const hasName = keys.some(k => /NOMBRE|PRODUCTO|ARTICULO|DESCRIPCION/.test(k));

    // Detectar si es RECETA: debe tener insumo y cantidad teórica
    const hasInsumo = keys.some(k => /INSUMO|CODIGO/.test(k));
    const hasTeorica = keys.some(k => /TEORICA|TEÓRICA|CANT/.test(k));

    // Detectar si es PRECIOS: debe tener código y precio
    const hasCod = keys.some(k => /CODIGO|COD/.test(k));
    const hasPrecio = keys.some(k => /PRECIO|COSTO|COMPRA/.test(k));

    // Clasificar por nombre de archivo (fallback)
    const nameLower = fileName.toLowerCase();
    if (nameLower.includes('venta') || (hasDate && hasQty && hasName)) {
        DB.ventas = data;
    } else if (nameLower.includes('receta') || (hasInsumo && hasTeorica)) {
        DB.recetas = data;
    } else if (nameLower.includes('precio') || (hasCod && hasPrecio)) {
        DB.precios = data;
    } else if (nameLower.includes('teorico') || (hasInsumo && hasName)) {
        DB.teoricos = data;
    } else {
        // Si no se pudo clasificar, lo guardamos como teórico por defecto
        DB.teoricos = data;
    }
}

// ========== POBLAR FILTROS ==========
function populateFilterOptions() {
    const selectMes = document.getElementById('select-mes');
    const selectArticulo = document.getElementById('select-articulo');

    // 1. Meses desde Ventas
    let mesesSet = new Set();
    DB.ventas.forEach(row => {
        const keyFecha = findKey(row, ['FECHA', 'DATE', 'DIA', 'MES']);
        if (keyFecha && row[keyFecha]) {
            const mes = formatToMesAnio(row[keyFecha]);
            if (mes) mesesSet.add(mes);
        }
    });

    selectMes.innerHTML = '<option value="">-- Seleccionar Mes --</option>';
    mesesSet.forEach(m => {
        selectMes.innerHTML += `<option value="${m}">${m}</option>`;
    });
    selectMes.disabled = (mesesSet.size === 0);

    // 2. Artículos desde Ventas (columna NOMBRE)
    let articulosSet = new Set();
    DB.ventas.forEach(row => {
        const keyNombre = findKey(row, ['NOMBRE', 'PRODUCTO', 'ARTICULO', 'DESCRIPCION']);
        if (keyNombre && row[keyNombre]) {
            const nombre = String(row[keyNombre]).trim();
            if (nombre) articulosSet.add(nombre);
        }
    });

    selectArticulo.innerHTML = '<option value="">-- Seleccionar Artículo --</option>';
    articulosSet.forEach(nombre => {
        selectArticulo.innerHTML += `<option value="${nombre}">${nombre}</option>`;
    });
    selectArticulo.disabled = (articulosSet.size === 0);

    if (!articulosSet.size) {
        document.getElementById('data-status').innerText += ' ⚠️ No se encontraron artículos en la tabla de ventas.';
    }
}

// ========== APLICAR FILTROS ==========
function applyFilters() {
    const mesSel = document.getElementById('select-mes').value;
    const artSel = document.getElementById('select-articulo').value;

    if (!artSel) {
        limpiarPantalla('Seleccione un artículo para ver su receta.');
        return;
    }

    // --- 1. Calcular cantidad vendida en el mes ---
    let cantVendida = 0;
    if (mesSel) {
        DB.ventas.forEach(row => {
            const keyFecha = findKey(row, ['FECHA', 'DATE', 'DIA', 'MES']);
            const keyNombre = findKey(row, ['NOMBRE', 'PRODUCTO', 'ARTICULO', 'DESCRIPCION']);
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

    // --- 2. Buscar receta del artículo ---
    let recetasFiltradas = [];
    if (DB.recetas.length > 0) {
        // Intentar encontrar la columna que contiene el nombre del artículo
        const sampleRow = DB.recetas[0];
        const keyArtReceta = findKey(sampleRow, ['ARTICULO', 'PRODUCTO', 'NOMBRE']);
        if (keyArtReceta) {
            recetasFiltradas = DB.recetas.filter(row => {
                return String(row[keyArtReceta]).trim() === artSel;
            });
        } else {
            // Fallback: buscar en cualquier columna
            recetasFiltradas = DB.recetas.filter(row => {
                return Object.values(row).some(val => String(val).trim() === artSel);
            });
        }
    }

    if (recetasFiltradas.length === 0) {
        document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
        document.getElementById('kpi-costo-unitario').innerText = '$ 0.00';
        document.getElementById('kpi-costo-total').innerText = '$ 0.00';
        document.getElementById('kpi-desvio').innerText = '0.0%';
        document.querySelector('#table-receta tbody').innerHTML = 
            '<tr><td colspan="6" class="empty-msg">No se encontró receta para este artículo.</td></tr>';
        if (insumosChart) insumosChart.destroy();
        return;
    }

    // --- 3. Procesar insumos de la receta ---
    let desgloseInsumos = [];
    let costoUnitarioTotal = 0;

    recetasFiltradas.forEach(item => {
        const keyCod = findKey(item, ['CODIGO INSUMO', 'COD INSUMO', 'CODIGO', 'INSUMO']);
        const keyCant = findKey(item, ['CANTIDAD', 'TEORICA', 'CANT']);
        const keyDesc = findKey(item, ['DESCRIPCION', 'NOMBRE INSUMO', 'NOMBRE']);

        let codInsumo = keyCod ? String(item[keyCod]).trim() : '';
        let cantTeorica = keyCant ? parseFloat(item[keyCant]) || 0 : 0;
        let descInsumo = keyDesc ? String(item[keyDesc]).trim() : '';

        if (!descInsumo) descInsumo = 'Insumo ' + codInsumo;

        // Buscar precio
        let precioCompra = 0;
        if (codInsumo) {
            for (let pRow of DB.precios) {
                const keyCodP = findKey(pRow, ['CODIGO', 'INSUMO', 'COD']);
                if (keyCodP && String(pRow[keyCodP]).trim() === codInsumo) {
                    const keyPrecio = findKey(pRow, ['PRECIO', 'COSTO', 'COMPRA']);
                    if (keyPrecio) {
                        precioCompra = parseFloat(pRow[keyPrecio]) || 0;
                    }
                    break;
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

    const costoVentaTotal = costoUnitarioTotal * cantVendida;

    // --- 4. Calcular desvío vs teórico (si existe tabla teóricos) ---
    let desvio = 0;
    if (DB.teoricos.length > 0 && costoUnitarioTotal > 0) {
        // Buscar el costo teórico para este artículo
        const teoricoRow = DB.teoricos.find(row => {
            const keyArt = findKey(row, ['ARTICULO', 'PRODUCTO', 'NOMBRE']);
            return keyArt && String(row[keyArt]).trim() === artSel;
        });
        if (teoricoRow) {
            const keyCostoTeorico = findKey(teoricoRow, ['COSTO', 'PRECIO', 'TEORICO']);
            if (keyCostoTeorico) {
                const costoTeorico = parseFloat(teoricoRow[keyCostoTeorico]) || 0;
                desvio = ((costoUnitarioTotal - costoTeorico) / costoTeorico) * 100;
            }
        }
    }

    // --- 5. Actualizar KPIs ---
    document.getElementById('kpi-ventas').innerText = cantVendida.toLocaleString('es-AR');
    document.getElementById('kpi-costo-unitario').innerText = '$ ' + costoUnitarioTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('kpi-costo-total').innerText = '$ ' + costoVentaTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
    document.getElementById('kpi-desvio').innerText = desvio.toFixed(1) + '%';

    // --- 6. Renderizar tabla y gráfico ---
    renderTable(desgloseInsumos, costoUnitarioTotal);
    renderChart(desgloseInsumos);
}

// ========== RENDER ==========
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

    const colores = ['#F40009', '#111111', '#555555', '#888888', '#BBBBBB', '#990000', '#CC0000', '#FF4D4D', '#333333', '#D4AF37', '#777777', '#E60000'];
    insumosChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: insumosConCosto.map(i => i.descripcion),
            datasets: [{
                data: insumosConCosto.map(i => i.subtotal),
                backgroundColor: colores.slice(0, insumosConCosto.length),
                borderWidth: 2,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { boxWidth: 12, padding: 10, font: { size: 10 } } },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return ' ' + context.label + ': $' + context.raw.toLocaleString('es-AR', {minimumFractionDigits: 2});
                        }
                    }
                }
            }
        }
    });
}

function limpiarPantalla(mensaje) {
    document.getElementById('kpi-ventas').innerText = '0';
    document.getElementById('kpi-costo-unitario').innerText = '$ 0.00';
    document.getElementById('kpi-costo-total').innerText = '$ 0.00';
    document.getElementById('kpi-desvio').innerText = '0.0%';
    document.querySelector('#table-receta tbody').innerHTML = `<tr><td colspan="6" class="empty-msg">${mensaje}</td></tr>`;
    if (insumosChart) insumosChart.destroy();
}
