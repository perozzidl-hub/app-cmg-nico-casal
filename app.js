document.addEventListener('DOMContentLoaded', function () {
  // ========== REFERENCIAS DOM (Todas las existentes) ==========
  const uploadArea = document.getElementById('uploadArea');
  const filtrosSection = document.getElementById('filtrosSection');
  const resumenSection = document.getElementById('resumenSection');
  const graficosSection = document.getElementById('graficosSection');
  const tablaSection = document.getElementById('tablaSection');
  
  const fileInput = document.querySelector('.upload-box input[type="file"]');
  const loadRepoBtn = document.getElementById('loadRepoBtn');
  const fileStatus = document.getElementById('fileStatus');

  const selectArticulo = document.getElementById('selectArticulo');
  const selectMes = document.getElementById('selectMes');
  const selectUbicacion = document.getElementById('selectUbicacion');
  const selectTipo = document.getElementById('selectTipo');

  const totalUnidadesEl = document.getElementById('totalUnidades');
  const facturacionNetaEl = document.getElementById('facturacionNeta');
  const costoTotalEl = document.getElementById('costoTotal');
  const margenBrutoEl = document.getElementById('margenBruto');
  const costoUnitarioEl = document.getElementById('costoUnitario');
  const promedioUnidadesEl = document.getElementById('promedioUnidades');
  const margenUnidadEl = document.getElementById('margenUnidad');
  const totalRegistrosEl = document.getElementById('totalRegistros');

  const tableHead = document.getElementById('tableHead');
  const tableBody = document.getElementById('tableBody');
  const registrosCount = document.getElementById('registrosCount');
  const exportCsvBtn = document.getElementById('exportCsvBtn');
  const exportPdfBtn = document.getElementById('exportPdfBtn');
  const darkModeToggle = document.getElementById('darkModeToggle');

  const chartVentasCanvas = document.getElementById('chartVentas');
  const chartComposicionCanvas = document.getElementById('chartComposicion');

  // ========== ESTADO GLOBAL ==========
  let state = { ventas: [], datosFiltrados: [] };
  let chartVentas = null;
  let chartComposicion = null;

  // ========== UTILIDADES ==========
  function formatearMonedaCorta(valor) {
    if (!valor) return '$0';
    if (Math.abs(valor) >= 1e9) return '$' + (valor / 1e9).toFixed(2) + 'B';
    if (Math.abs(valor) >= 1e6) return '$' + (valor / 1e6).toFixed(1) + 'M';
    if (Math.abs(valor) >= 1e3) return '$' + (valor / 1e3).toFixed(1) + 'K';
    return '$' + valor.toFixed(2);
  }

  function obtenerMes(valor) {
    if (!valor) return 'Sin Mes';
    if (typeof valor === 'number') {
      const date = new Date(Math.round((valor - 25569) * 86400 * 1000));
      return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
    }
    const str = String(valor).trim();
    const meses = { ene: '01', feb: '02', mar: '03', abr: '04', may: '05', jun: '06', jul: '07', ago: '08', sep: '09', oct: '10', nov: '11', dic: '12' };
    if (str.includes('-')) {
      const parts = str.toLowerCase().split('-');
      if (meses[parts[0].substring(0, 3)] && parts[1]) return '20' + parts[1] + '-' + meses[parts[0].substring(0, 3)];
    }
    if (str.includes('/')) {
      const p = str.split('/');
      if (p.length === 3 && p[2].length === 4) return p[2] + '-' + p[1];
    }
    return 'Sin Mes';
  }

  function nombreMes(ym) {
    const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const parts = ym.split('-');
    if (parts.length !== 2) return ym;
    return meses[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
  }

  // ========== LECTURA DEL EXCEL (Ignora filas de arriba) ==========
  function normalizarHeader(header) {
    if (!header) return '';
    return header.toString().trim().replace(/[^a-zA-Z0-9_ ]/g, '').replace(/\s+/g, '_');
  }

  function findHeaderRow(dataArray) {
    const keywords = ['ART', 'FECHA', 'Nombre', 'Facturación', 'Físicos', 'TOTAL INSUMOS', 'LOCACION'];
    for (let i = 0; i < Math.min(dataArray.length, 20); i++) {
      const rowText = (dataArray[i] || []).map(c => String(c).toUpperCase()).join(' ');
      let matchCount = 0;
      keywords.forEach(kw => { if (rowText.includes(kw)) matchCount++; });
      if (matchCount >= 2) return i;
    }
    return 0;
  }

  function procesarVentas(dataArray) {
    const headerIndex = findHeaderRow(dataArray);
    if (dataArray.length < headerIndex + 1) return [];
    const headers = dataArray[headerIndex].map(normalizarHeader);
    const rows = dataArray.slice(headerIndex + 1);
    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((key, idx) => obj[key] = row[idx] !== undefined ? row[idx] : '');
      return obj;
    }).filter(o => o.ART && o.ART !== '');

    const numCols = ['Físicos', 'Facturación_Neta', 'TOTAL_INSUMOS', 'COSTO_AJUSTES_DE_STOCK', 'TOTAL_PRODUCTO_TERCEROS', 'Total_Concentrado_', 'Total_Fructosa_', 'Total_Gas_Carbónico_', 'Total_Tapas_'];
    objetos.forEach(o => {
      numCols.forEach(col => o[col] = parseFloat(o[col]) || 0);
      o._mes = obtenerMes(o.FECHA);
    });
    return objetos;
  }

  // ========== MOSTRAR DASHBOARD ==========
  function mostrarDashboard() {
    uploadArea.style.display = 'none';
    filtrosSection.style.display = 'flex';
    resumenSection.style.display = 'grid';
    graficosSection.style.display = 'grid';
    tablaSection.style.display = 'block';
  }

  // ========== CARGAR ARCHIVO MANUAL ==========
  function procesarArchivo(dataArray) {
    state.ventas = procesarVentas(dataArray);
    if (state.ventas.length === 0) {
      alert('No se encontraron datos debajo de los encabezados. Revisa el archivo.');
      return;
    }
    mostrarDashboard();
    poblarFiltros(state.ventas);
    aplicarFiltros();
    fileStatus.textContent = '✅ Datos cargados correctamente.';
  }

  fileInput.addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (loadEvent) {
      try {
        const data = new Uint8Array(loadEvent.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const dataArray = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
        procesarArchivo(dataArray);
      } catch (err) {
        alert('Error al leer el archivo: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // ========== CARGAR DESDE REPOSITORIO (SOLUCIONADO) ==========
  async function cargarDesdeRepo() {
    fileStatus.textContent = '⏳ Intentando cargar ventas.xlsx...';
    // Si el usuario está en local (file://), avisarle que use un servidor o GitHub Pages
    if (window.location.protocol === 'file:') {
        fileStatus.textContent = '❌ Abre este archivo con "Live Server" o súbelo a GitHub Pages para que el botón funcione. Mientras tanto, puedes subirlo manualmente.';
        return;
    }

    try {
      const response = await fetch('ventas.xlsx');
      if (!response.ok) throw new Error('Archivo no encontrado');
      
      const arrayBuffer = await response.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      const dataArray = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });

      state.ventas = procesarVentas(dataArray);
      
      if (state.ventas.length === 0) {
        fileStatus.textContent = '⚠️ El Excel no tiene datos debajo de los encabezados.';
        return;
      }

      mostrarDashboard();
      poblarFiltros(state.ventas);
      aplicarFiltros();
      fileStatus.textContent = '✅ Datos cargados automáticamente desde ventas.xlsx';

    } catch (error) {
      fileStatus.textContent = '❌ No se pudo cargar desde el repo. Sube el archivo manualmente.';
      console.error('Error de CORS o archivo no encontrado:', error);
    }
  }

  // Vincular botón del repositorio
  loadRepoBtn.addEventListener('click', cargarDesdeRepo);

  // Intentar carga automática al abrir
  cargarDesdeRepo();

  // ========== FILTROS Y DASHBOARD ==========
  function poblarFiltros(datos) {
    const artSet = new Set(datos.map(o => o.ART + ' - ' + o.Nombre).filter(Boolean));
    selectArticulo.innerHTML = '<option value="todos">Todos los Artículos</option>';
    Array.from(artSet).sort().forEach(a => selectArticulo.innerHTML += `<option value="${a.split(' - ')[0]}">${a}</option>`);

    const mesesSet = new Set(datos.map(o => o._mes).filter(m => m !== 'Sin Mes'));
    selectMes.innerHTML = '<option value="todos">Todos los meses</option>';
    Array.from(mesesSet).sort().forEach(m => selectMes.innerHTML += `<option value="${m}">${nombreMes(m)}</option>`);

    const ubicSet = new Set(datos.map(o => o.LOCACION_SAP).filter(Boolean));
    selectUbicacion.innerHTML = '<option value="todas">Todas</option>';
    Array.from(ubicSet).sort().forEach(u => selectUbicacion.innerHTML += `<option value="${u}">${u}</option>`);
    
    selectArticulo.addEventListener('change', aplicarFiltros);
    selectMes.addEventListener('change', aplicarFiltros);
    selectUbicacion.addEventListener('change', aplicarFiltros);
    selectTipo.addEventListener('change', aplicarFiltros);
  }

  function aplicarFiltros() {
    const art = selectArticulo.value;
    const mes = selectMes.value;
    const ubic = selectUbicacion.value;
    const tipo = selectTipo.value;

    state.datosFiltrados = state.ventas.filter(o => {
      if (art !== 'todos' && o.ART !== art) return false;
      if (mes !== 'todos' && o._mes !== mes) return false;
      if (ubic !== 'todas' && o.LOCACION_SAP !== ubic) return false;
      if (tipo !== 'todos' && o.Tipo_de_Prod_ !== tipo) return false;
      return true;
    });

    actualizarDashboard();
  }

  // ========== ACTUALIZAR DASHBOARD ==========
  function actualizarDashboard() {
    const datos = state.datosFiltrados;
    if (datos.length === 0) {
      totalUnidadesEl.textContent = '0';
      facturacionNetaEl.textContent = '$0';
      costoTotalEl.textContent = '$0';
      margenBrutoEl.textContent = '0%';
      costoUnitarioEl.textContent = '$0';
      promedioUnidadesEl.textContent = '0';
      margenUnidadEl.textContent = '$0';
      totalRegistrosEl.textContent = '0';
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
    const promedioUnidades = datos.length > 0 ? totalUnidades / datos.length : 0;
    const margenUnidad = totalUnidades > 0 ? (sumFacturacion - sumCostoTotal) / totalUnidades : 0;

    totalUnidadesEl.textContent = totalUnidades.toLocaleString('es-AR');
    facturacionNetaEl.textContent = formatearMonedaCorta(sumFacturacion);
    costoTotalEl.textContent = formatearMonedaCorta(sumCostoTotal);
    margenBrutoEl.textContent = margen.toFixed(1) + '%';
    costoUnitarioEl.textContent = formatearMonedaCorta(costoUnitario);
    promedioUnidadesEl.textContent = promedioUnidades.toFixed(0);
    margenUnidadEl.textContent = formatearMonedaCorta(margenUnidad);
    totalRegistrosEl.textContent = datos.length;

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

    const ventasPorMes = {};
    datos.forEach(o => ventasPorMes[o._mes] = (ventasPorMes[o._mes] || 0) + (o.Físicos || 0));
    const meses = Object.keys(ventasPorMes).sort();
    const ctx1 = chartVentasCanvas.getContext('2d');
    chartVentas = new Chart(ctx1, {
      type: 'line',
      data: { labels: meses.map(nombreMes), datasets: [{ label: 'Unidades', data: meses.map(m => ventasPorMes[m]), borderColor: '#E61A27', backgroundColor: 'rgba(230,26,39,0.1)', fill: true, tension: 0.4 }] },
      options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
    });

    const insumos = ['Total_Concentrado_', 'Total_Fructosa_', 'Total_Gas_Carbónico_', 'Total_Tapas_'];
    const labels = ['Concentrado','Fructosa','Gas','Tapas'];
    const dataComp = insumos.map(col => datos.reduce((acc, o) => acc + (o[col] || 0), 0));
    const ctx3 = chartComposicionCanvas.getContext('2d');
    chartComposicion = new Chart(ctx3, {
      type: 'doughnut',
      data: { labels, datasets: [{ data: dataComp, backgroundColor: ['#E61A27', '#FF6B6B', '#FFC300', '#3498DB'] }] },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }

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

  // ========== EXPORTAR ==========
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

  exportPdfBtn.addEventListener('click', function () {
    const element = document.getElementById('dashboardMain');
    html2canvas(element, { scale: 2, useCORS: true }).then(canvas => {
      const imgData = canvas.toDataURL('image/png');
      const win = window.open('', '_blank');
      win.document.write('<html><head><title>Dashboard CMG</title><style>body{margin:0;}img{max-width:100%;}</style></head><body>');
      win.document.write(`<img src="${imgData}" />`);
      win.document.write('</body></html>');
      win.document.close();
      win.focus();
      win.print();
    });
  });

  // ========== MODO OSCURO (CORREGIDO) ==========
  darkModeToggle.addEventListener('click', function () {
    document.body.classList.toggle('dark-mode');
    const icon = this.querySelector('i');
    if (document.body.classList.contains('dark-mode')) {
      icon.className = 'fas fa-sun';
    } else {
      icon.className = 'fas fa-moon';
    }
    if (state.datosFiltrados) actualizarGraficos(state.datosFiltrados);
  });
});
