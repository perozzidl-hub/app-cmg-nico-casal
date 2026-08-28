// app.js
document.addEventListener('DOMContentLoaded', function () {
  // ========== REFERENCIAS DOM ==========
  const uploadBoxes = document.querySelectorAll('.upload-box');
  const loadDemoBtn = document.getElementById('loadDemoBtn');
  const processBtn = document.getElementById('processDataBtn');
  const fileStatus = document.getElementById('fileStatus');

  const articulosContainer = document.getElementById('articulosContainer');
  const selectMes = document.getElementById('selectMes');
  const selectUbicacion = document.getElementById('selectUbicacion');
  const selectTipo = document.getElementById('selectTipo');
  const selectAllArticulos = document.getElementById('selectAllArticulos');
  const deselectAllArticulos = document.getElementById('deselectAllArticulos');

  const totalUnidadesEl = document.getElementById('totalUnidades');
  const facturacionNetaEl = document.getElementById('facturacionNeta');
  const costoTotalEl = document.getElementById('costoTotal');
  const margenBrutoEl = document.getElementById('margenBruto');
  const costoUnitarioEl = document.getElementById('costoUnitario');

  const tableHead = document.getElementById('tableHead');
  const tableBody = document.getElementById('tableBody');
  const registrosCount = document.getElementById('registrosCount');
  const exportCsvBtn = document.getElementById('exportCsvBtn');
  const exportPdfBtn = document.getElementById('exportPdfBtn');
  const darkModeToggle = document.getElementById('darkModeToggle');

  const chartVentasCanvas = document.getElementById('chartVentas');
  const chartCostoCanvas = document.getElementById('chartCostoUnitario');
  const chartComposicionCanvas = document.getElementById('chartComposicion');

  // ========== ESTADO GLOBAL ==========
  let state = {
    ventas: [],          // Array de objetos (ventas procesadas)
    recetas: [],         // Array de objetos { idArticulo, idInsumo, cantidad }
    costos: [],          // Array de objetos { idInsumo, fecha, precio }
    datosFiltrados: [],
    archivosCargados: { ventas: false, recetas: false, costos: false },
    modoOscuro: false,
  };

  let chartVentas = null;
  let chartCosto = null;
  let chartComposicion = null;

  // ========== FUNCIONES AUXILIARES ==========
  function formatearMoneda(valor) {
    return '$' + Number(valor).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  function formatearPorcentaje(valor) {
    return Number(valor).toFixed(1) + '%';
  }
  function obtenerMes(fechaStr) {
    if (!fechaStr) return '';
    const partes = fechaStr.trim().split(' ');
    const fechaParte = partes[0];
    if (!fechaParte) return '';
    const sub = fechaParte.split('-');
    return sub.length >= 2 ? sub[0] + '-' + sub[1] : '';
  }
  function nombreMes(ym) {
    const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                   'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const parts = ym.split('-');
    if (parts.length !== 2) return ym;
    const mes = parseInt(parts[1], 10) - 1;
    return meses[mes] + ' ' + parts[0];
  }

  // ========== PROCESAR DATOS DE EXCEL ==========
  function procesarVentas(dataArray) {
    if (!dataArray || dataArray.length < 2) return [];
    const headers = dataArray[0];
    const rows = dataArray.slice(1);
    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((h, idx) => {
        let key = h ? h.toString().trim() : `col_${idx}`;
        key = key.replace(/[^a-zA-Z0-9_ ]/g, '').trim().replace(/ /g, '_');
        obj[key] = row[idx] !== undefined ? row[idx] : '';
      });
      return obj;
    });
    // Filtrar filas sin artículo
    const validos = objetos.filter(o => o.ART && o.ART !== '');
    // Convertir columnas numéricas
    const numCols = ['Físicos', 'Unit', 'Facturación_Lista', 'Facturación_Neta',
      'TOTAL_INSUMOS', 'Precio_Prod._Reventa', 'Incidencia_de_Reventa',
      'Ajuste_de_concentrado_por_compra', 'TOTAL_PRODUCTO_TERCEROS',
      'AJUSTE_DE_STOCK_(Cajas_Físicas)', 'COSTO_AJUSTES_DE_STOCK',
      'Total_Concentrado_($)', 'Total_Fructosa_($)', 'Total_Gas_Carbónico_($)',
      'Total_Tapas_($)', 'Total_Etiquetas_($)', 'Total_Preformas_($)',
      'Total_Termocontraible_($)', 'Total_Stretch_($)', 'Total_Separadores_($)',
      'Total_Caja_BIB_($)', 'Total_Bolsa_BIB_($)', 'Total_Nitrógeno_($)'
    ];
    validos.forEach(o => {
      numCols.forEach(col => {
        if (o[col] !== undefined && o[col] !== '') {
          const val = parseFloat(o[col]);
          o[col] = isNaN(val) ? 0 : val;
        } else o[col] = 0;
      });
      if (o.Físicos && typeof o.Físicos === 'string') {
        o.Físicos = parseFloat(o.Físicos.replace(/,/g, '')) || 0;
      }
      o._mes = obtenerMes(o.FECHA);
      o._key = (o.ART || '') + (o.LOCACION___SAP || '');
    });
    return validos;
  }

  function procesarRecetas(dataArray) {
    if (!dataArray || dataArray.length < 2) return [];
    const headers = dataArray[0];
    const rows = dataArray.slice(1);
    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((h, idx) => {
        let key = h ? h.toString().trim() : `col_${idx}`;
        key = key.replace(/[^a-zA-Z0-9_ ]/g, '').trim().replace(/ /g, '_');
        obj[key] = row[idx] !== undefined ? row[idx] : '';
      });
      return obj;
    });
    // Esperamos columnas: idArticulo, idInsumo, cantidad
    const validos = objetos.filter(o => o.idArticulo && o.idInsumo);
    validos.forEach(o => {
      o.cantidad = parseFloat(o.cantidad) || 0;
    });
    return validos;
  }

  function procesarCostos(dataArray) {
    if (!dataArray || dataArray.length < 2) return [];
    const headers = dataArray[0];
    const rows = dataArray.slice(1);
    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((h, idx) => {
        let key = h ? h.toString().trim() : `col_${idx}`;
        key = key.replace(/[^a-zA-Z0-9_ ]/g, '').trim().replace(/ /g, '_');
        obj[key] = row[idx] !== undefined ? row[idx] : '';
      });
      return obj;
    });
    const validos = objetos.filter(o => o.idInsumo && o.fecha);
    validos.forEach(o => {
      o.precioUnitario = parseFloat(o.precioUnitario) || 0;
      o._mes = obtenerMes(o.fecha);
    });
    return validos;
  }

  // ========== RECALCULAR COSTOS USANDO RECETAS Y COSTOS ==========
  function recalcularCostos() {
    if (!state.recetas.length || !state.costos.length) {
      // Si no hay recetas o costos, usamos los valores ya presentes en ventas
      return state.ventas;
    }

    // Crear mapa de costos por insumo y mes (tomar el más reciente)
    const costosMap = {};
    state.costos.forEach(c => {
      const key = c.idInsumo + '|' + (c._mes || '');
      if (!costosMap[key] || c.fecha > costosMap[key].fecha) {
        costosMap[key] = c;
      }
    });

    // Recalcular cada venta
    const nuevasVentas = state.ventas.map(venta => {
      const idArt = venta.ART;
      // Buscar recetas para este artículo
      const recetasArt = state.recetas.filter(r => r.idArticulo == idArt);
      let costoTotalRecalculado = 0;
      let desglose = {};

      recetasArt.forEach(rec => {
        const key = rec.idInsumo + '|' + (venta._mes || '');
        const costo = costosMap[key];
        if (costo) {
          const importe = rec.cantidad * costo.precioUnitario;
          costoTotalRecalculado += importe;
          desglose[rec.idInsumo] = (desglose[rec.idInsumo] || 0) + importe;
        }
      });

      // Actualizar campos de insumos en la venta
      // Asignar los valores calculados a las columnas de insumos (para gráficos)
      // Esto es complejo porque no sabemos qué columna corresponde a cada insumo.
      // Lo dejamos simple: si se usan recetas, el TOTAL_INSUMOS se recalcula,
      // pero los insumos individuales se dejan como 0 (o se podría mapear).
      // Para el gráfico de composición, usaremos el desglose.
      venta._costoRecalculado = costoTotalRecalculado;
      venta._desgloseInsumos = desglose;
      // Actualizar TOTAL_INSUMOS
      venta.TOTAL_INSUMOS = costoTotalRecalculado;
      return venta;
    });
    return nuevasVentas;
  }

  // ========== CARGAR ARCHIVOS ==========
  function handleFileUpload(e, type) {
    const file = e.target.files[0];
    if (!file) return;
    const box = e.target.closest('.upload-box');
    const statusSpan = box.querySelector('.file-status-msg');

    const reader = new FileReader();
    reader.onload = function (loadEvent) {
      try {
        const data = new Uint8Array(loadEvent.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const dataArray = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

        if (type === 'ventas') {
          state.ventas = procesarVentas(dataArray);
          state.archivosCargados.ventas = true;
          statusSpan.textContent = `✅ ${state.ventas.length} registros`;
          statusSpan.className = 'file-status-msg cargado';
        } else if (type === 'recetas') {
          state.recetas = procesarRecetas(dataArray);
          state.archivosCargados.recetas = true;
          statusSpan.textContent = `✅ ${state.recetas.length} recetas`;
          statusSpan.className = 'file-status-msg cargado';
        } else if (type === 'costos') {
          state.costos = procesarCostos(dataArray);
          state.archivosCargados.costos = true;
          statusSpan.textContent = `✅ ${state.costos.length} costos`;
          statusSpan.className = 'file-status-msg cargado';
        }
        checkFilesReady();
      } catch (err) {
        alert('Error al leer ' + type + ': ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // Asignar eventos a cada caja de carga
  uploadBoxes.forEach(box => {
    const input = box.querySelector('input[type="file"]');
    const type = box.dataset.type;
    input.addEventListener('change', (e) => handleFileUpload(e, type));
  });

  function checkFilesReady() {
    if (state.archivosCargados.ventas) {
      processBtn.disabled = false;
      fileStatus.textContent = '✅ Archivo de ventas cargado. Haz clic en "Procesar datos".';
    } else {
      processBtn.disabled = true;
      fileStatus.textContent = '⚠️ Carga al menos el archivo de Ventas.';
    }
  }

  // ========== PROCESAR DATOS (botón) ==========
  processBtn.addEventListener('click', function () {
    if (!state.archivosCargados.ventas) {
      alert('Debes cargar un archivo de Ventas.');
      return;
    }

    // Recalcular costos si hay recetas y costos
    let datos = state.ventas;
    if (state.recetas.length && state.costos.length) {
      datos = recalcularCostos();
    }

    // Guardar y mostrar
    state.datosProcesados = datos;
    mostrarSecciones();
    poblarFiltros(datos);
    aplicarFiltros();
    fileStatus.textContent = '✅ Datos procesados correctamente.';
  });

  // ========== DATOS DE EJEMPLO ==========
  loadDemoBtn.addEventListener('click', function () {
    // Cargar ventas de ejemplo (simulamos)
    const demoVentas = [
      { FECHA: '2026-07-01', Tipo_de_Prod_: 'P', MARCA: 'COCA-COLA', LOCACION___SAP: 'Ranelagh', ART: '105', Nombre: 'COLA 250CC', Físicos: 1594.5, Facturación_Neta: 11924249.09, TOTAL_INSUMOS: 139800, COSTO_AJUSTES_DE_STOCK: 0, TOTAL_PRODUCTO_TERCEROS: 0, 'Total_Concentrado_($)': 100000, 'Total_Fructosa_($)': 20000, 'Total_Gas_Carbónico_($)': 5000, 'Total_Tapas_($)': 3000, 'Total_Etiquetas_($)': 2000, 'Total_Preformas_($)': 8000, 'Total_Termocontraible_($)': 1000, 'Total_Stretch_($)': 500, 'Total_Separadores_($)': 300, 'Total_Caja_BIB_($)': 0, 'Total_Bolsa_BIB_($)': 0, 'Total_Nitrógeno_($)': 0 },
      { FECHA: '2026-08-01', Tipo_de_Prod_: 'P', MARCA: 'COCA-COLA', LOCACION___SAP: 'Ranelagh', ART: '105', Nombre: 'COLA 250CC', Físicos: 1800, Facturación_Neta: 13000000, TOTAL_INSUMOS: 153780, COSTO_AJUSTES_DE_STOCK: 0, TOTAL_PRODUCTO_TERCEROS: 0, 'Total_Concentrado_($)': 110000, 'Total_Fructosa_($)': 22000, 'Total_Gas_Carbónico_($)': 5500, 'Total_Tapas_($)': 3300, 'Total_Etiquetas_($)': 2200, 'Total_Preformas_($)': 8800, 'Total_Termocontraible_($)': 1100, 'Total_Stretch_($)': 550, 'Total_Separadores_($)': 330, 'Total_Caja_BIB_($)': 0, 'Total_Bolsa_BIB_($)': 0, 'Total_Nitrógeno_($)': 0 },
      { FECHA: '2026-07-01', Tipo_de_Prod_: 'P', MARCA: 'FANTA', LOCACION___SAP: 'Ranelagh', ART: '310', Nombre: 'FANTA 350CC', Físicos: 324, Facturación_Neta: 7088360.90, TOTAL_INSUMOS: 105900, COSTO_AJUSTES_DE_STOCK: 0, TOTAL_PRODUCTO_TERCEROS: 0, 'Total_Concentrado_($)': 80000, 'Total_Fructosa_($)': 15000, 'Total_Gas_Carbónico_($)': 3000, 'Total_Tapas_($)': 4000, 'Total_Etiquetas_($)': 2500, 'Total_Preformas_($)': 0, 'Total_Termocontraible_($)': 800, 'Total_Stretch_($)': 400, 'Total_Separadores_($)': 200, 'Total_Caja_BIB_($)': 0, 'Total_Bolsa_BIB_($)': 0, 'Total_Nitrógeno_($)': 0 },
      { FECHA: '2026-09-01', Tipo_de_Prod_: 'r', MARCA: 'COCA-COLA', LOCACION___SAP: 'Dolores', ART: '180', Nombre: 'COLA 10L BIB', Físicos: 82, Facturación_Neta: 10952205.36, TOTAL_INSUMOS: 0, COSTO_AJUSTES_DE_STOCK: 0, TOTAL_PRODUCTO_TERCEROS: 9500000, 'Total_Concentrado_($)': 0, 'Total_Fructosa_($)': 0, 'Total_Gas_Carbónico_($)': 0, 'Total_Tapas_($)': 0, 'Total_Etiquetas_($)': 0, 'Total_Preformas_($)': 0, 'Total_Termocontraible_($)': 0, 'Total_Stretch_($)': 0, 'Total_Separadores_($)': 0, 'Total_Caja_BIB_($)': 0, 'Total_Bolsa_BIB_($)': 0, 'Total_Nitrógeno_($)': 0 }
    ];
    // Añadir _mes
    demoVentas.forEach(o => { o._mes = obtenerMes(o.FECHA); });
    state.ventas = demoVentas;
    state.archivosCargados.ventas = true;
    // Marcamos recetas y costos como no disponibles (usamos los costos embebidos)
    state.recetas = [];
    state.costos = [];
    processBtn.disabled = false;
    fileStatus.textContent = '✅ Datos de ejemplo cargados. Haz clic en "Procesar datos".';
    // Actualizar UI de carga
    document.querySelector('.upload-box[data-type="ventas"] .file-status-msg').textContent = '✅ Demo cargada';
    document.querySelector('.upload-box[data-type="ventas"] .file-status-msg').className = 'file-status-msg cargado';
  });

  // ========== MOSTRAR SECCIONES ==========
  function mostrarSecciones() {
    document.getElementById('filtrosSection').style.display = 'flex';
    document.getElementById('resumenSection').style.display = 'grid';
    document.getElementById('graficosSection').style.display = 'grid';
    document.getElementById('tablaSection').style.display = 'block';
    document.getElementById('uploadArea').style.display = 'none';
  }

  // ========== POBLAR FILTROS ==========
  function poblarFiltros(datos) {
    // Artículos (checkboxes)
    const articulosSet = new Set();
    datos.forEach(o => {
      if (o.ART && o.Nombre) {
        articulosSet.add(o.ART + '|' + o.Nombre);
      }
    });
    const articulosArr = Array.from(articulosSet).map(str => {
      const [cod, nom] = str.split('|');
      return { cod, nom };
    });
    articulosArr.sort((a, b) => a.nom.localeCompare(b.nom));

    articulosContainer.innerHTML = '';
    articulosArr.forEach(a => {
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = a.cod;
      cb.checked = true;
      label.appendChild(cb);
      label.appendChild(document.createTextNode(a.cod + ' - ' + a.nom));
      articulosContainer.appendChild(label);
    });

    // Meses
    const mesesSet = new Set();
    datos.forEach(o => { if (o._mes) mesesSet.add(o._mes); });
    const mesesArr = Array.from(mesesSet).sort();
    selectMes.innerHTML = '<option value="todos">Todos los meses</option>';
    mesesArr.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m;
      opt.textContent = nombreMes(m);
      selectMes.appendChild(opt);
    });

    // Ubicaciones
    const ubicSet = new Set();
    datos.forEach(o => { if (o.LOCACION___SAP) ubicSet.add(o.LOCACION___SAP); });
    const ubicArr = Array.from(ubicSet).sort();
    selectUbicacion.innerHTML = '<option value="todas">Todas</option>';
    ubicArr.forEach(u => {
      const opt = document.createElement('option');
      opt.value = u;
      opt.textContent = u;
      selectUbicacion.appendChild(opt);
    });

    // Eventos para checkboxes
    document.querySelectorAll('#articulosContainer input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', aplicarFiltros);
    });
    selectAllArticulos.addEventListener('click', () => {
      document.querySelectorAll('#articulosContainer input[type="checkbox"]').forEach(cb => cb.checked = true);
      aplicarFiltros();
    });
    deselectAllArticulos.addEventListener('click', () => {
      document.querySelectorAll('#articulosContainer input[type="checkbox"]').forEach(cb => cb.checked = false);
      aplicarFiltros();
    });
  }

  // ========== FILTRAR DATOS ==========
  function aplicarFiltros() {
    const checkboxes = document.querySelectorAll('#articulosContainer input[type="checkbox"]:checked');
    const articulosSeleccionados = Array.from(checkboxes).map(cb => cb.value);
    const mes = selectMes.value;
    const ubic = selectUbicacion.value;
    const tipo = selectTipo.value;

    const datos = state.datosProcesados || [];
    state.datosFiltrados = datos.filter(o => {
      if (articulosSeleccionados.length > 0 && !articulosSeleccionados.includes(o.ART)) return false;
      if (mes !== 'todos' && o._mes !== mes) return false;
      if (ubic !== 'todas' && o.LOCACION___SAP !== ubic) return false;
      if (tipo !== 'todos' && o.Tipo_de_Prod_ !== tipo) return false;
      return true;
    });

    actualizarDashboard();
  }

  // Escuchar cambios en filtros (selects)
  selectMes.addEventListener('change', aplicarFiltros);
  selectUbicacion.addEventListener('change', aplicarFiltros);
  selectTipo.addEventListener('change', aplicarFiltros);

  // ========== ACTUALIZAR DASHBOARD ==========
  function actualizarDashboard() {
    const datos = state.datosFiltrados;
    if (datos.length === 0) {
      totalUnidadesEl.textContent = '0';
      facturacionNetaEl.textContent = '$0.00';
      costoTotalEl.textContent = '$0.00';
      margenBrutoEl.textContent = '0%';
      costoUnitarioEl.textContent = '$0.00';
      tableBody.innerHTML = '';
      tableHead.innerHTML = '';
      registrosCount.textContent = '0 registros';
      destruirGraficos();
      return;
    }

    // Cálculos agregados
    let totalUnidades = 0, sumFacturacion = 0, sumCostoTotal = 0;
    datos.forEach(o => {
      const units = o.Físicos || 0;
      totalUnidades += units;
      sumFacturacion += (o.Facturación_Neta || 0);
      let costo = (o.TOTAL_INSUMOS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      if (o.Tipo_de_Prod_ === 'r') {
        costo = (o.TOTAL_PRODUCTO_TERCEROS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      }
      sumCostoTotal += costo;
    });

    const costoUnitario = totalUnidades > 0 ? sumCostoTotal / totalUnidades : 0;
    const margen = sumFacturacion > 0 ? ((sumFacturacion - sumCostoTotal) / sumFacturacion) * 100 : 0;

    totalUnidadesEl.textContent = totalUnidades.toFixed(0);
    facturacionNetaEl.textContent = formatearMoneda(sumFacturacion);
    costoTotalEl.textContent = formatearMoneda(sumCostoTotal);
    margenBrutoEl.textContent = formatearPorcentaje(margen);
    costoUnitarioEl.textContent = formatearMoneda(costoUnitario);

    actualizarGraficos(datos);
    actualizarTabla(datos);
  }

  // ========== GRÁFICOS ==========
  function destruirGraficos() {
    if (chartVentas) { chartVentas.destroy(); chartVentas = null; }
    if (chartCosto) { chartCosto.destroy(); chartCosto = null; }
    if (chartComposicion) { chartComposicion.destroy(); chartComposicion = null; }
  }

  function actualizarGraficos(datos) {
    destruirGraficos();

    // 1. Ventas mensuales
    const ventasPorMes = {};
    datos.forEach(o => {
      const mes = o._mes || 'sin mes';
      ventasPorMes[mes] = (ventasPorMes[mes] || 0) + (o.Físicos || 0);
    });
    const mesesOrd = Object.keys(ventasPorMes).sort();
    const ctx1 = chartVentasCanvas.getContext('2d');
    chartVentas = new Chart(ctx1, {
      type: 'bar',
      data: {
        labels: mesesOrd.map(m => nombreMes(m) || m),
        datasets: [{ label: 'Unidades', data: mesesOrd.map(m => ventasPorMes[m]), backgroundColor: '#ED1C24', borderRadius: 4 }]
      },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
    });

    // 2. Costo unitario mensual
    const costoPorMes = {}, unitsPorMes = {};
    datos.forEach(o => {
      const mes = o._mes || 'sin mes';
      let costo = (o.TOTAL_INSUMOS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      if (o.Tipo_de_Prod_ === 'r') costo = (o.TOTAL_PRODUCTO_TERCEROS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      costoPorMes[mes] = (costoPorMes[mes] || 0) + costo;
      unitsPorMes[mes] = (unitsPorMes[mes] || 0) + (o.Físicos || 0);
    });
    const mesesCosto = Object.keys(costoPorMes).sort();
    const costoUnitarioData = mesesCosto.map(m => unitsPorMes[m] > 0 ? costoPorMes[m] / unitsPorMes[m] : 0);
    const ctx2 = chartCostoCanvas.getContext('2d');
    chartCosto = new Chart(ctx2, {
      type: 'line',
      data: {
        labels: mesesCosto.map(m => nombreMes(m) || m),
        datasets: [{ label: 'Costo unitario ($)', data: costoUnitarioData, borderColor: '#ED1C24', backgroundColor: 'rgba(237,28,36,0.1)', fill: true, tension: 0.2, pointBackgroundColor: '#ED1C24', pointRadius: 5 }]
      },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
    });

    // 3. Composición del costo (solo productos propios)
    const insumos = ['Total_Concentrado_($)','Total_Fructosa_($)','Total_Gas_Carbónico_($)','Total_Tapas_($)','Total_Etiquetas_($)','Total_Preformas_($)','Total_Termocontraible_($)','Total_Stretch_($)','Total_Separadores_($)','Total_Caja_BIB_($)','Total_Bolsa_BIB_($)','Total_Nitrógeno_($)'];
    const nombres = ['Concentrado','Fructosa','Gas Carbónico','Tapas','Etiquetas','Preformas','Termocontraible','Stretch','Separadores','Caja BIB','Bolsa BIB','Nitrógeno'];
    const sumas = nombres.map(() => 0);
    datos.forEach(o => {
      if (o.Tipo_de_Prod_ !== 'P') return;
      insumos.forEach((col, idx) => {
        sumas[idx] += (o[col] || 0);
      });
    });
    const dataComp = [];
    const labelsComp = [];
    const colores = ['#ED1C24','#FF6B6B','#FFA07A','#FFD93D','#6BCB77','#4D96FF','#9B59B6','#F39C12','#1ABC9C','#E67E22','#2ECC71','#3498DB'];
    sumas.forEach((val, idx) => {
      if (val > 0) {
        dataComp.push(val);
        labelsComp.push(nombres[idx]);
      }
    });
    const ctx3 = chartComposicionCanvas.getContext('2d');
    chartComposicion = new Chart(ctx3, {
      type: 'pie',
      data: {
        labels: labelsComp,
        datasets: [{ data: dataComp, backgroundColor: colores.slice(0, dataComp.length), borderColor: '#fff', borderWidth: 1 }]
      },
      options: { responsive: true, plugins: { legend: { position: 'right', labels: { boxWidth: 12, font: { size: 12 } } } } }
    });
  }

  // ========== TABLA ==========
  function actualizarTabla(datos) {
    if (datos.length === 0) { tableBody.innerHTML = ''; tableHead.innerHTML = ''; registrosCount.textContent = '0 registros'; return; }
    const cols = ['FECHA','Nombre','Físicos','Facturación_Neta','TOTAL_INSUMOS','COSTO_AJUSTES_DE_STOCK','TOTAL_PRODUCTO_TERCEROS','LOCACION___SAP','Tipo_de_Prod_'];
    const headers = { 'FECHA':'Fecha','Nombre':'Artículo','Físicos':'Unidades','Facturación_Neta':'Fact. Neta','TOTAL_INSUMOS':'Costo Insumos','COSTO_AJUSTES_DE_STOCK':'Ajuste Stock','TOTAL_PRODUCTO_TERCEROS':'Costo Terceros','LOCACION___SAP':'Ubicación','Tipo_de_Prod_':'Tipo' };
    let thead = '<tr>' + cols.map(c => `<th>${headers[c]||c}</th>`).join('') + '</tr>';
    tableHead.innerHTML = thead;
    const maxRows = 100;
    const subset = datos.slice(0, maxRows);
    let tbody = '';
    subset.forEach(o => {
      tbody += '<tr>' + cols.map(c => {
        let val = o[c] !== undefined ? o[c] : '';
        if (typeof val === 'number') val = val.toFixed(2);
        return `<td>${val}</td>`;
      }).join('') + '</tr>';
    });
    tableBody.innerHTML = tbody;
    registrosCount.textContent = `${datos.length} registros (mostrando ${Math.min(datos.length, maxRows)})`;
  }

  // ========== EXPORTAR CSV ==========
  exportCsvBtn.addEventListener('click', function () {
    if (!state.datosFiltrados || state.datosFiltrados.length === 0) { alert('No hay datos para exportar.'); return; }
    const cols = ['FECHA','Nombre','Físicos','Facturación_Neta','TOTAL_INSUMOS','COSTO_AJUSTES_DE_STOCK','TOTAL_PRODUCTO_TERCEROS','LOCACION___SAP','Tipo_de_Prod_'];
    const headers = { 'FECHA':'Fecha','Nombre':'Artículo','Físicos':'Unidades','Facturación_Neta':'Fact. Neta','TOTAL_INSUMOS':'Costo Insumos','COSTO_AJUSTES_DE_STOCK':'Ajuste Stock','TOTAL_PRODUCTO_TERCEROS':'Costo Terceros','LOCACION___SAP':'Ubicación','Tipo_de_Prod_':'Tipo' };
    let csv = cols.map(c => headers[c]||c).join(',') + '\n';
    state.datosFiltrados.forEach(o => {
      csv += cols.map(c => { let v = o[c] !== undefined ? o[c] : ''; if (typeof v === 'string' && v.includes(',')) v = '"'+v+'"'; return v; }).join(',') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'ventas_filtradas.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  });

  // ========== EXPORTAR PDF ==========
  exportPdfBtn.addEventListener('click', function () {
    const element = document.getElementById('dashboardMain');
    html2canvas(element, { scale: 2, useCORS: true, backgroundColor: getComputedStyle(document.body).backgroundColor }).then(canvas => {
      const imgData = canvas.toDataURL('image/png');
      const { jsPDF } = window.jspdf; // Nota: necesitamos cargar jsPDF
      // Como no lo cargamos por CDN, usamos una ventana nueva para imprimir/guardar como PDF
      // Alternativa: abrir una nueva ventana con la imagen y usar print -> save as PDF
      const win = window.open('', '_blank');
      win.document.write('<html><head><title>Dashboard CMG</title><style>body{margin:0;display:flex;justify-content:center;align-items:center;background:#fff;}img{max-width:100%;}</style></head><body>');
      win.document.write(`<img src="${imgData}" />`);
      win.document.write('</body></html>');
      win.document.close();
      win.focus();
      win.print(); // Esto permite guardar como PDF
    }).catch(err => alert('Error al generar PDF: ' + err.message));
  });

  // ========== MODO OSCURO ==========
  darkModeToggle.addEventListener('click', function () {
    document.body.classList.toggle('dark-mode');
    const icon = this.querySelector('i');
    if (document.body.classList.contains('dark-mode')) {
      icon.className = 'fas fa-sun';
    } else {
      icon.className = 'fas fa-moon';
    }
    // Refrescar gráficos para que se vean bien
    if (state.datosFiltrados) actualizarGraficos(state.datosFiltrados);
  });

  // ========== INICIO ==========
  // Si no hay datos, mostrar mensaje
  fileStatus.textContent = 'Carga un archivo Excel o usa datos de ejemplo.';
});
