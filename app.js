// app.js - VERSIÓN FINAL: Detecta automáticamente dónde empiezan los encabezados
document.addEventListener('DOMContentLoaded', function () {
  // ========== REFERENCIAS DOM ==========
  const uploadBox = document.querySelector('.upload-box input[type="file"]');
  const loadDemoBtn = document.getElementById('loadDemoBtn');
  const fileStatus = document.getElementById('fileStatus');

  const selectMes = document.getElementById('selectMes');
  const selectUbicacion = document.getElementById('selectUbicacion');
  const selectTipo = document.getElementById('selectTipo');

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
  const chartComposicionCanvas = document.getElementById('chartComposicion');

  // ========== ESTADO GLOBAL ==========
  let state = {
    ventas: [],
    datosFiltrados: [],
    modoOscuro: false,
  };

  let chartVentas = null;
  let chartComposicion = null;

  // ========== UTILIDADES DE FORMATO ==========
  function normalizarHeader(header) {
    if (!header) return '';
    // Elimina $, (, ), ., etc. y reemplaza CUALQUIER espacio por un solo guion bajo
    return header.toString().trim().replace(/[^a-zA-Z0-9_ ]/g, '').replace(/\s+/g, '_');
  }

  function obtenerMes(fechaStr) {
    if (!fechaStr) return 'Sin Mes';
    const meses = { ene: '01', feb: '02', mar: '03', abr: '04', may: '05', jun: '06', jul: '07', ago: '08', sep: '09', oct: '10', nov: '11', dic: '12' };
    const partes = String(fechaStr).trim().toLowerCase().split('-');
    if (partes.length !== 2) return 'Sin Mes';
    const mesAbr = partes[0].substring(0, 3);
    const anio = partes[1];
    if (meses[mesAbr] && anio) return '20' + anio + '-' + meses[mesAbr];
    return String(fechaStr).trim();
  }

  function nombreMes(ym) {
    const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const parts = ym.split('-');
    if (parts.length !== 2) return ym;
    return meses[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
  }

  function formatearMonedaCorta(valor) {
    if (!valor) return '$0';
    if (Math.abs(valor) >= 1e9) return '$' + (valor / 1e9).toFixed(2) + 'B';
    if (Math.abs(valor) >= 1e6) return '$' + (valor / 1e6).toFixed(1) + 'M';
    if (Math.abs(valor) >= 1e3) return '$' + (valor / 1e3).toFixed(1) + 'K';
    return '$' + valor.toFixed(2);
  }

  // ========== LÓGICA PARA ENCONTRAR LA FILA DE ENCABEZADOS ==========
  // Recorre las primeras 20 filas hasta encontrar la que tiene los títulos reales
  function findHeaderRow(dataArray) {
    const keywords = ['ART', 'FECHA', 'Nombre', 'Facturación', 'Físicos', 'TOTAL INSUMOS', 'LOCACION'];
    
    for (let i = 0; i < Math.min(dataArray.length, 20); i++) {
        const row = dataArray[i] || [];
        const rowText = row.map(c => String(c).toUpperCase()).join(' ');
        let matchCount = 0;

        keywords.forEach(kw => {
            if (rowText.includes(kw)) matchCount++;
        });

        // Si encuentra al menos 2 palabras clave, asume que es la fila de encabezados
        if (matchCount >= 2) return i;
    }
    return 0; // Si no encuentra nada, vuelve a la fila 0 (comportamiento original)
  }

  // ========== PROCESAMIENTO DE DATOS (CON DETECCIÓN DE ENCABEZADOS) ==========
  function procesarVentas(dataArray) {
    // Detectar el índice real de la fila de encabezados
    const headerIndex = findHeaderRow(dataArray);
    if (dataArray.length < headerIndex + 1) return [];

    const headers = dataArray[headerIndex].map(normalizarHeader);
    const rows = dataArray.slice(headerIndex + 1); // Solo toma los datos debajo de los encabezados

    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((key, idx) => obj[key] = row[idx] !== undefined ? row[idx] : '');
      return obj;
    }).filter(o => o.ART && o.ART !== '');

    const numCols = ['Físicos', 'Facturación_Neta', 'TOTAL_INSUMOS', 'COSTO_AJUSTES_DE_STOCK', 'TOTAL_PRODUCTO_TERCEROS', 
                     'Total_Concentrado_', 'Total_Fructosa_', 'Total_Gas_Carbónico_', 'Total_Tapas_', 
                     'Total_Etiquetas_', 'Total_Preformas_', 'Total_Termocontraible_', 'Total_Stretch_', 
                     'Total_Separadores_', 'Total_Caja_BIB_', 'Total_Bolsa_BIB_', 'Total_Nitrógeno_'];
    
    objetos.forEach(o => {
      numCols.forEach(col => o[col] = parseFloat(o[col]) || 0);
      o._mes = obtenerMes(o.FECHA);
    });
    return objetos;
  }

  // ========== CARGA DE ARCHIVOS (UNIFICADA) ==========
  uploadBox.addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;
    const statusSpan = document.querySelector('.upload-box .file-status-msg');

    const reader = new FileReader();
    reader.onload = function (loadEvent) {
      try {
        const data = new Uint8Array(loadEvent.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const dataArray = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });

        // Procesar con la nueva lógica de detección
        state.ventas = procesarVentas(dataArray);

        if (state.ventas.length === 0) {
            statusSpan.textContent = '⚠️ No se encontraron registros debajo de los encabezados';
            return;
        }

        statusSpan.textContent = `✅ ${state.ventas.length} registros cargados`;
        statusSpan.className = 'file-status-msg cargado';
        
        // Ocultar zona de carga y mostrar dashboard automáticamente
        document.getElementById('uploadArea').style.display = 'none';
        document.getElementById('filtrosSection').style.display = 'flex';
        document.getElementById('resumenSection').style.display = 'grid';
        document.getElementById('graficosSection').style.display = 'grid';
        document.getElementById('tablaSection').style.display = 'block';

        poblarFiltros(state.ventas);
        aplicarFiltros();
        fileStatus.textContent = '✅ Datos procesados correctamente.';
        
      } catch (err) {
        alert('Error al leer ' + file.name + ': ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // ========== DATOS DE EJEMPLO ==========
  loadDemoBtn.addEventListener('click', function () {
    state.ventas = [
      { FECHA: 'jul-26', Tipo_de_Prod_: 'P', LOCACION_SAP: 'Ranelagh', ART: '105', Nombre: 'COLA 250CC', Físicos: 1594.5, Facturación_Neta: 11924249.09, TOTAL_INSUMOS: 4981763, Total_Concentrado_: 2408698 },
      { FECHA: 'jul-26', Tipo_de_Prod_: 'R', LOCACION_SAP: 'Ranelagh', ART: '118', Nombre: 'COLA 354CC LATA', Físicos: 12672, Facturación_Neta: 91504598, TOTAL_INSUMOS: 0, Total_Concentrado_: 0 }
    ];
    state.ventas.forEach(o => o._mes = obtenerMes(o.FECHA));
    document.getElementById('uploadArea').style.display = 'none';
    document.getElementById('filtrosSection').style.display = 'flex';
    document.getElementById('resumenSection').style.display = 'grid';
    document.getElementById('graficosSection').style.display = 'grid';
    document.getElementById('tablaSection').style.display = 'block';
    poblarFiltros(state.ventas);
    aplicarFiltros();
  });

  // ========== FILTROS Y DASHBOARD ==========
  function poblarFiltros(datos) {
    const mesesSet = new Set(datos.map(o => o._mes));
    selectMes.innerHTML = '<option value="todos">Todos los meses</option>';
    Array.from(mesesSet).sort().forEach(m => selectMes.innerHTML += `<option value="${m}">${nombreMes(m)}</option>`);

    const ubicSet = new Set(datos.map(o => o.LOCACION_SAP).filter(Boolean));
    selectUbicacion.innerHTML = '<option value="todas">Todas</option>';
    Array.from(ubicSet).sort().forEach(u => selectUbicacion.innerHTML += `<option value="${u}">${u}</option>`);
  }

  function aplicarFiltros() {
    const mes = selectMes.value;
    const ubic = selectUbicacion.value;
    const tipo = selectTipo.value;

    state.datosFiltrados = state.ventas.filter(o => {
      if (mes !== 'todos' && o._mes !== mes) return false;
      if (ubic !== 'todas' && o.LOCACION_SAP !== ubic) return false;
      if (tipo !== 'todos' && o.Tipo_de_Prod_ !== tipo) return false;
      return true;
    });

    actualizarDashboard();
  }

  selectMes.addEventListener('change', aplicarFiltros);
  selectUbicacion.addEventListener('change', aplicarFiltros);
  selectTipo.addEventListener('change', aplicarFiltros);

  // ========== ACTUALIZAR DASHBOARD ==========
  function actualizarDashboard() {
    const datos = state.datosFiltrados;
    if (datos.length === 0) {
      totalUnidadesEl.textContent = '0';
      facturacionNetaEl.textContent = '$0';
      costoTotalEl.textContent = '$0';
      margenBrutoEl.textContent = '0%';
      costoUnitarioEl.textContent = '$0';
      tableHead.innerHTML = ''; tableBody.innerHTML = ''; registrosCount.textContent = '0 registros';
      destruirGraficos(); return;
    }

    let totalUnidades = 0, sumFacturacion = 0, sumCostoTotal = 0;
    datos.forEach(o => {
      totalUnidades += o.Físicos || 0;
      sumFacturacion += o.Facturación_Neta || 0;
      sumCostoTotal += (o.TOTAL_INSUMOS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
    });

    const costoUnitario = totalUnidades > 0 ? sumCostoTotal / totalUnidades : 0;
    const margen = sumFacturacion > 0 ? ((sumFacturacion - sumCostoTotal) / sumFacturacion) * 100 : 0;

    totalUnidadesEl.textContent = totalUnidades.toLocaleString('es-AR');
    facturacionNetaEl.textContent = formatearMonedaCorta(sumFacturacion);
    costoTotalEl.textContent = formatearMonedaCorta(sumCostoTotal);
    margenBrutoEl.textContent = margen.toFixed(1) + '%';
    costoUnitarioEl.textContent = formatearMonedaCorta(costoUnitario);

    actualizarGraficos(datos);
    actualizarTabla(datos);
  }

  // ========== GRÁFICOS ==========
  function destruirGraficos() {
    if (chartVentas) { chartVentas.destroy(); chartVentas = null; }
    if (chartComposicion) { chartComposicion.destroy(); chartComposicion = null; }
  }

  function actualizarGraficos(datos) {
    destruirGraficos();

    // 1. Tendencia de Ventas
    const ventasPorMes = {};
    datos.forEach(o => ventasPorMes[o._mes] = (ventasPorMes[o._mes] || 0) + (o.Físicos || 0));
    const meses = Object.keys(ventasPorMes).sort();
    const ctx1 = chartVentasCanvas.getContext('2d');
    chartVentas = new Chart(ctx1, {
      type: 'line',
      data: { labels: meses.map(nombreMes), datasets: [{ label: 'Unidades', data: meses.map(m => ventasPorMes[m]), borderColor: '#ED1C24', backgroundColor: 'rgba(237,28,36,0.1)', fill: true, tension: 0.4 }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
    });

    // 2. Composición de Costos
    const insumos = ['Total_Concentrado_', 'Total_Fructosa_', 'Total_Gas_Carbónico_', 'Total_Tapas_', 'Total_Etiquetas_', 'Total_Preformas_', 'Total_Termocontraible_'];
    const labels = ['Concentrado','Fructosa','Gas','Tapas','Etiquetas','Preformas','Termocontraible'];
    const dataComp = insumos.map(col => datos.reduce((acc, o) => acc + (o[col] || 0), 0));
    const ctx3 = chartComposicionCanvas.getContext('2d');
    chartComposicion = new Chart(ctx3, {
      type: 'doughnut',
      data: { labels, datasets: [{ data: dataComp, backgroundColor: ['#ED1C24', '#FF6B6B', '#FFC300', '#3498DB', '#2ECC71', '#9B59B6', '#E67E22'] }] },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }

  // ========== TABLA (TOP 10) ==========
  function actualizarTabla(datos) {
    const sorted = [...datos].sort((a, b) => (b.Facturación_Neta || 0) - (a.Facturación_Neta || 0)).slice(0, 10);
    const cols = ['FECHA','Nombre','Físicos','Facturación_Neta','TOTAL_INSUMOS','LOCACION_SAP'];
    const headers = { 'FECHA':'Fecha','Nombre':'Artículo','Físicos':'Unidades','Facturación_Neta':'Fact. Neta','TOTAL_INSUMOS':'Costo Insumos','LOCACION_SAP':'Ubicación' };

    tableHead.innerHTML = '<tr>' + cols.map(c => `<th>${headers[c]||c}</th>`).join('') + '</tr>';
    let tbody = '';
    sorted.forEach(o => {
      tbody += '<tr>' + cols.map(c => {
        let val = o[c] !== undefined ? o[c] : '';
        if (typeof val === 'number') val = val.toLocaleString('es-AR', { minimumFractionDigits: 2 });
        return `<td>${val}</td>`;
      }).join('') + '</tr>';
    });
    tableBody.innerHTML = tbody;
    registrosCount.textContent = `${datos.length} registros totales (Top 10 por Facturación)`;
  }

  // ========== EXPORTAR CSV ==========
  exportCsvBtn.addEventListener('click', function () {
    if (!state.datosFiltrados || state.datosFiltrados.length === 0) { alert('No hay datos para exportar.'); return; }
    const cols = ['FECHA','Nombre','Físicos','Facturación_Neta','TOTAL_INSUMOS','LOCACION_SAP'];
    let csv = cols.join(',') + '\n';
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
      const win = window.open('', '_blank');
      win.document.write('<html><head><title>Dashboard CMG</title><style>body{margin:0;}img{max-width:100%;}</style></head><body>');
      win.document.write(`<img src="${imgData}" />`);
      win.document.write('</body></html>');
      win.document.close();
      win.focus();
      win.print();
    }).catch(err => alert('Error al generar PDF: ' + err.message));
  });

  // ========== MODO OSCURO ==========
  darkModeToggle.addEventListener('click', function () {
    document.body.classList.toggle('dark-mode');
    const icon = this.querySelector('i');
    icon.className = document.body.classList.contains('dark-mode') ? 'fas fa-sun' : 'fas fa-moon';
    if (state.datosFiltrados) actualizarGraficos(state.datosFiltrados);
  });

  // ========== INICIO ==========
  fileStatus.textContent = 'Carga un archivo Excel (los encabezados pueden estar en cualquier fila).';
});
