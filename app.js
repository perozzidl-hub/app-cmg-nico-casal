// app.js
document.addEventListener('DOMContentLoaded', function () {
  // ========== REFERENCIAS DOM ==========
  const fileInput = document.getElementById('fileInput');
  const loadDemoBtn = document.getElementById('loadDemoBtn');
  const fileStatus = document.getElementById('fileStatus');
  const uploadArea = document.getElementById('uploadArea');

  const selectArticulo = document.getElementById('selectArticulo');
  const searchArticulo = document.getElementById('searchArticulo');
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

  const chartVentasCanvas = document.getElementById('chartVentas');
  const chartCostoCanvas = document.getElementById('chartCostoUnitario');
  const chartComposicionCanvas = document.getElementById('chartComposicion');

  // ========== VARIABLES GLOBALES ==========
  let datosOriginales = [];      // Array de objetos (todas las filas)
  let datosFiltrados = [];       // Array filtrado según selecciones
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

  // Extraer mes (AAAA-MM) de una fecha en formato "2026-07-01 00:00:00"
  function obtenerMes(fechaStr) {
    if (!fechaStr) return '';
    const partes = fechaStr.trim().split(' ');
    const fechaParte = partes[0]; // "2026-07-01"
    if (!fechaParte) return '';
    const sub = fechaParte.split('-');
    if (sub.length < 2) return '';
    return sub[0] + '-' + sub[1]; // "2026-07"
  }

  // Obtener nombre de mes a partir de "YYYY-MM"
  function nombreMes(ym) {
    const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
                   'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const parts = ym.split('-');
    if (parts.length !== 2) return ym;
    const mes = parseInt(parts[1], 10) - 1;
    return meses[mes] + ' ' + parts[0];
  }

  // ========== PROCESAR DATOS DESDE EXCEL ==========
  function procesarDatosExcel(dataArray) {
    // Convertir a objetos con claves limpias
    const headers = dataArray[0];
    const rows = dataArray.slice(1);
    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((h, idx) => {
        let key = h ? h.toString().trim() : `col_${idx}`;
        // Limpiar nombres de columnas (eliminar caracteres extraños)
        key = key.replace(/[^a-zA-Z0-9_ ]/g, '').trim().replace(/ /g, '_');
        obj[key] = row[idx] !== undefined ? row[idx] : '';
      });
      return obj;
    });

    // Filtrar filas vacías (si el artículo está vacío, descartar)
    const validos = objetos.filter(o => o.ART && o.ART !== '');

    // Convertir columnas numéricas
    const columnasNumericas = [
      'Físicos', 'Unit', 'Facturación_Lista', 'Facturación_Neta',
      'TOTAL_INSUMOS', 'Precio_Prod._Reventa', 'Incidencia_de_Reventa',
      'Ajuste_de_concentrado_por_compra', 'TOTAL_PRODUCTO_TERCEROS',
      'AJUSTE_DE_STOCK_(Cajas_Físicas)', 'COSTO_AJUSTES_DE_STOCK',
      // Insumos individuales
      'Total_Concentrado_($)', 'Total_Fructosa_($)', 'Total_Gas_Carbónico_($)',
      'Total_Tapas_($)', 'Total_Etiquetas_($)', 'Total_Preformas_($)',
      'Total_Termocontraible_($)', 'Total_Stretch_($)', 'Total_Separadores_($)',
      'Total_Caja_BIB_($)', 'Total_Bolsa_BIB_($)', 'Total_Nitrógeno_($)'
    ];

    validos.forEach(o => {
      columnasNumericas.forEach(col => {
        if (o[col] !== undefined && o[col] !== '') {
          const val = parseFloat(o[col]);
          o[col] = isNaN(val) ? 0 : val;
        } else {
          o[col] = 0;
        }
      });
      // Asegurar que Físicos es numérico
      if (o.Físicos && typeof o.Físicos === 'string') {
        o.Físicos = parseFloat(o.Físicos.replace(/,/g, '')) || 0;
      }
      // Extraer mes
      if (o.FECHA) {
        o._mes = obtenerMes(o.FECHA);
      } else {
        o._mes = '';
      }
      // Crear una clave combinada para identificación
      o._key = (o.ART || '') + (o.LOCACION___SAP || '');
    });

    return validos;
  }

  // ========== CARGAR DATOS ==========
  function cargarDatosDesdeExcel(workbook) {
    // Asumimos que la primera hoja es la que interesa
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataArray = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    if (dataArray.length < 2) {
      alert('El archivo no contiene datos válidos.');
      return;
    }
    const datos = procesarDatosExcel(dataArray);
    if (datos.length === 0) {
      alert('No se encontraron registros con datos de artículo.');
      return;
    }
    datosOriginales = datos;
    fileStatus.innerHTML = '<i class="fas fa-check-circle"></i> Datos cargados correctamente. ' + datos.length + ' registros.';
    mostrarSecciones();
    poblarFiltros();
    aplicarFiltros();
  }

  // Cargar archivo subido
  fileInput.addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (loadEvent) {
      try {
        const data = new Uint8Array(loadEvent.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        cargarDatosDesdeExcel(workbook);
      } catch (err) {
        alert('Error al leer el archivo: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // Cargar datos de ejemplo (DEMO) – usamos los datos incrustados en el código
  loadDemoBtn.addEventListener('click', function () {
    // Los datos de ejemplo se definen al final del script
    if (typeof datosDemo !== 'undefined' && datosDemo.length > 0) {
      datosOriginales = datosDemo;
      fileStatus.innerHTML = '<i class="fas fa-check-circle"></i> Datos de ejemplo cargados. ' + datosDemo.length + ' registros.';
      mostrarSecciones();
      poblarFiltros();
      aplicarFiltros();
    } else {
      alert('No hay datos de ejemplo disponibles. Sube un archivo Excel.');
    }
  });

  // ========== MOSTRAR SECCIONES ==========
  function mostrarSecciones() {
    document.getElementById('filtrosSection').style.display = 'flex';
    document.getElementById('resumenSection').style.display = 'grid';
    document.getElementById('graficosSection').style.display = 'grid';
    document.getElementById('tablaSection').style.display = 'block';
    uploadArea.style.display = 'none';
  }

  // ========== POBLAR FILTROS ==========
  function poblarFiltros() {
    // Artículos
    const articulosSet = new Set();
    datosOriginales.forEach(o => {
      if (o.ART && o.Nombre) {
        articulosSet.add(o.ART + '|' + o.Nombre);
      }
    });
    const articulosArr = Array.from(articulosSet).map(str => {
      const [cod, nom] = str.split('|');
      return { cod, nom };
    });
    articulosArr.sort((a, b) => a.nom.localeCompare(b.nom));

    selectArticulo.innerHTML = '<option value="">Todos los artículos</option>';
    articulosArr.forEach(a => {
      const opt = document.createElement('option');
      opt.value = a.cod;
      opt.textContent = a.cod + ' - ' + a.nom;
      selectArticulo.appendChild(opt);
    });

    // Meses
    const mesesSet = new Set();
    datosOriginales.forEach(o => {
      if (o._mes) mesesSet.add(o._mes);
    });
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
    datosOriginales.forEach(o => {
      if (o.LOCACION___SAP) ubicSet.add(o.LOCACION___SAP);
    });
    const ubicArr = Array.from(ubicSet).sort();
    selectUbicacion.innerHTML = '<option value="todas">Todas</option>';
    ubicArr.forEach(u => {
      const opt = document.createElement('option');
      opt.value = u;
      opt.textContent = u;
      selectUbicacion.appendChild(opt);
    });

    // Tipo (ya tiene opciones fijas)
  }

  // ========== FILTRAR DATOS ==========
  function aplicarFiltros() {
    const artCod = selectArticulo.value;
    const mes = selectMes.value;
    const ubic = selectUbicacion.value;
    const tipo = selectTipo.value;

    // Búsqueda por nombre de artículo (si hay texto)
    const searchText = searchArticulo.value.toLowerCase().trim();

    datosFiltrados = datosOriginales.filter(o => {
      if (artCod && o.ART !== artCod) return false;
      if (mes !== 'todos' && o._mes !== mes) return false;
      if (ubic !== 'todas' && o.LOCACION___SAP !== ubic) return false;
      if (tipo !== 'todos' && o.Tipo_de_Prod_ !== tipo) return false;
      if (searchText) {
        const nombre = (o.Nombre || '').toLowerCase();
        if (!nombre.includes(searchText)) return false;
      }
      return true;
    });

    actualizarDashboard();
  }

  // Escuchar cambios en filtros
  selectArticulo.addEventListener('change', aplicarFiltros);
  selectMes.addEventListener('change', aplicarFiltros);
  selectUbicacion.addEventListener('change', aplicarFiltros);
  selectTipo.addEventListener('change', aplicarFiltros);
  searchArticulo.addEventListener('input', aplicarFiltros);

  // ========== ACTUALIZAR DASHBOARD ==========
  function actualizarDashboard() {
    if (datosFiltrados.length === 0) {
      // Mostrar ceros y vaciar gráficos
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
    let totalUnidades = 0;
    let sumFacturacion = 0;
    let sumCostoTotal = 0; // Usamos TOTAL_INSUMOS + ajustes

    datosFiltrados.forEach(o => {
      const unidades = o.Físicos || 0;
      totalUnidades += unidades;
      sumFacturacion += (o.Facturación_Neta || 0);
      // Costo total = TOTAL_INSUMOS + COSTO_AJUSTES_DE_STOCK + (TOTAL_PRODUCTO_TERCEROS si aplica)
      let costo = (o.TOTAL_INSUMOS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      // Si es reventa, usar TOTAL_PRODUCTO_TERCEROS en lugar de insumos
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

    // Gráficos
    actualizarGraficos();

    // Tabla
    actualizarTabla();
  }

  // ========== GRÁFICOS ==========
  function destruirGraficos() {
    if (chartVentas) { chartVentas.destroy(); chartVentas = null; }
    if (chartCosto) { chartCosto.destroy(); chartCosto = null; }
    if (chartComposicion) { chartComposicion.destroy(); chartComposicion = null; }
  }

  function actualizarGraficos() {
    destruirGraficos();

    // ----- 1. Ventas mensuales (barras) -----
    const ventasPorMes = {};
    datosFiltrados.forEach(o => {
      const mes = o._mes || 'sin mes';
      if (!ventasPorMes[mes]) ventasPorMes[mes] = 0;
      ventasPorMes[mes] += (o.Físicos || 0);
    });
    const mesesOrdenados = Object.keys(ventasPorMes).sort();
    const etiquetas = mesesOrdenados.map(m => nombreMes(m) || m);
    const valores = mesesOrdenados.map(m => ventasPorMes[m]);

    const ctx1 = chartVentasCanvas.getContext('2d');
    chartVentas = new Chart(ctx1, {
      type: 'bar',
      data: {
        labels: etiquetas,
        datasets: [{
          label: 'Unidades vendidas',
          data: valores,
          backgroundColor: '#ED1C24',
          borderRadius: 4,
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false }
        },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });

    // ----- 2. Evolución del costo unitario (línea) -----
    const costoPorMes = {};
    const unidadesPorMes = {};
    datosFiltrados.forEach(o => {
      const mes = o._mes || 'sin mes';
      if (!costoPorMes[mes]) costoPorMes[mes] = 0;
      if (!unidadesPorMes[mes]) unidadesPorMes[mes] = 0;
      let costo = (o.TOTAL_INSUMOS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      if (o.Tipo_de_Prod_ === 'r') {
        costo = (o.TOTAL_PRODUCTO_TERCEROS || 0) + (o.COSTO_AJUSTES_DE_STOCK || 0);
      }
      costoPorMes[mes] += costo;
      unidadesPorMes[mes] += (o.Físicos || 0);
    });
    const mesesCosto = Object.keys(costoPorMes).sort();
    const costoUnitarioPorMes = mesesCosto.map(m => {
      const units = unidadesPorMes[m] || 0;
      return units > 0 ? costoPorMes[m] / units : 0;
    });
    const etiquetasCosto = mesesCosto.map(m => nombreMes(m) || m);

    const ctx2 = chartCostoCanvas.getContext('2d');
    chartCosto = new Chart(ctx2, {
      type: 'line',
      data: {
        labels: etiquetasCosto,
        datasets: [{
          label: 'Costo unitario ($)',
          data: costoUnitarioPorMes,
          borderColor: '#ED1C24',
          backgroundColor: 'rgba(237,28,36,0.1)',
          fill: true,
          tension: 0.2,
          pointBackgroundColor: '#ED1C24',
          pointRadius: 5,
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false }
        },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });

    // ----- 3. Composición del costo (pastel) -----
    // Sumar cada insumo para los datos filtrados
    const insumos = {
      'Concentrado': 'Total_Concentrado_($)',
      'Fructosa': 'Total_Fructosa_($)',
      'Gas Carbónico': 'Total_Gas_Carbónico_($)',
      'Tapas': 'Total_Tapas_($)',
      'Etiquetas': 'Total_Etiquetas_($)',
      'Preformas': 'Total_Preformas_($)',
      'Termocontraible': 'Total_Termocontraible_($)',
      'Stretch': 'Total_Stretch_($)',
      'Separadores': 'Total_Separadores_($)',
      'Caja BIB': 'Total_Caja_BIB_($)',
      'Bolsa BIB': 'Total_Bolsa_BIB_($)',
      'Nitrógeno': 'Total_Nitrógeno_($)'
    };

    const sumInsumos = {};
    Object.keys(insumos).forEach(key => {
      sumInsumos[key] = 0;
    });

    datosFiltrados.forEach(o => {
      // Solo para productos propios (P)
      if (o.Tipo_de_Prod_ !== 'P') return;
      Object.keys(insumos).forEach(key => {
        const col = insumos[key];
        sumInsumos[key] += (o[col] || 0);
      });
    });

    // Filtrar insumos con valor > 0
    const labelsComposicion = [];
    const dataComposicion = [];
    const colores = ['#ED1C24','#FF6B6B','#FFA07A','#FFD93D','#6BCB77','#4D96FF','#9B59B6','#F39C12','#1ABC9C','#E67E22','#2ECC71','#3498DB'];
    Object.keys(sumInsumos).forEach((key, idx) => {
      if (sumInsumos[key] > 0) {
        labelsComposicion.push(key);
        dataComposicion.push(sumInsumos[key]);
      }
    });

    const ctx3 = chartComposicionCanvas.getContext('2d');
    chartComposicion = new Chart(ctx3, {
      type: 'pie',
      data: {
        labels: labelsComposicion,
        datasets: [{
          data: dataComposicion,
          backgroundColor: colores.slice(0, dataComposicion.length),
          borderColor: '#fff',
          borderWidth: 1,
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: {
            position: 'right',
            labels: {
              boxWidth: 12,
              font: { size: 12 }
            }
          }
        }
      }
    });
  }

  // ========== TABLA DETALLADA ==========
  function actualizarTabla() {
    if (datosFiltrados.length === 0) {
      tableBody.innerHTML = '';
      tableHead.innerHTML = '';
      registrosCount.textContent = '0 registros';
      return;
    }

    // Definir columnas a mostrar
    const columnasMostrar = [
      'FECHA', 'Nombre', 'Físicos', 'Facturación_Neta', 'TOTAL_INSUMOS',
      'COSTO_AJUSTES_DE_STOCK', 'TOTAL_PRODUCTO_TERCEROS', 'LOCACION___SAP',
      'Tipo_de_Prod_'
    ];
    // Encabezados más amigables
    const encabezados = {
      'FECHA': 'Fecha',
      'Nombre': 'Artículo',
      'Físicos': 'Unidades',
      'Facturación_Neta': 'Facturación Neta',
      'TOTAL_INSUMOS': 'Costo Insumos',
      'COSTO_AJUSTES_DE_STOCK': 'Ajuste Stock',
      'TOTAL_PRODUCTO_TERCEROS': 'Costo Terceros',
      'LOCACION___SAP': 'Ubicación',
      'Tipo_de_Prod_': 'Tipo'
    };

    // Construir cabecera
    let theadHtml = '<tr>';
    columnasMostrar.forEach(col => {
      const label = encabezados[col] || col;
      theadHtml += `<th>${label}</th>`;
    });
    theadHtml += '</tr>';
    tableHead.innerHTML = theadHtml;

    // Construir cuerpo (solo los primeros 100 registros para rendimiento)
    const maxRows = 100;
    const mostrar = datosFiltrados.slice(0, maxRows);
    let tbodyHtml = '';
    mostrar.forEach(o => {
      tbodyHtml += '<tr>';
      columnasMostrar.forEach(col => {
        let val = o[col] !== undefined ? o[col] : '';
        if (typeof val === 'number') {
          val = val.toFixed(2);
        }
        tbodyHtml += `<td>${val}</td>`;
      });
      tbodyHtml += '</tr>';
    });
    tableBody.innerHTML = tbodyHtml;

    const total = datosFiltrados.length;
    registrosCount.textContent = `${total} registros (mostrando ${Math.min(total, maxRows)})`;
  }

  // ========== EXPORTAR CSV ==========
  exportCsvBtn.addEventListener('click', function () {
    if (datosFiltrados.length === 0) {
      alert('No hay datos para exportar.');
      return;
    }
    // Usar las mismas columnas que la tabla
    const columnasMostrar = [
      'FECHA', 'Nombre', 'Físicos', 'Facturación_Neta', 'TOTAL_INSUMOS',
      'COSTO_AJUSTES_DE_STOCK', 'TOTAL_PRODUCTO_TERCEROS', 'LOCACION___SAP',
      'Tipo_de_Prod_'
    ];
    const encabezados = {
      'FECHA': 'Fecha',
      'Nombre': 'Artículo',
      'Físicos': 'Unidades',
      'Facturación_Neta': 'Facturación Neta',
      'TOTAL_INSUMOS': 'Costo Insumos',
      'COSTO_AJUSTES_DE_STOCK': 'Ajuste Stock',
      'TOTAL_PRODUCTO_TERCEROS': 'Costo Terceros',
      'LOCACION___SAP': 'Ubicación',
      'Tipo_de_Prod_': 'Tipo'
    };

    let csv = columnasMostrar.map(c => encabezados[c] || c).join(',') + '\n';
    datosFiltrados.forEach(o => {
      const row = columnasMostrar.map(col => {
        let val = o[col] !== undefined ? o[col] : '';
        if (typeof val === 'string' && val.includes(',')) {
          val = '"' + val + '"';
        }
        return val;
      });
      csv += row.join(',') + '\n';
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'ventas_filtradas.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  });

  // ========== DATOS DE EJEMPLO (EMBEBIDOS) ==========
  // Estos datos simulan la estructura del Excel original
  const datosDemo = [
    {
      "FECHA": "2026-07-01 00:00:00",
      "Tipo_de_Prod_": "P",
      "MARCA": "COCA-COLA",
      "LOCACION___SAP": "Org.Vtas Ranelagh",
      "ART": "105",
      "Nombre": "COCA-COLA 250CC PETX12",
      "Físicos": 1594.5,
      "Unit": 841.896,
      "Facturación_Lista": 12698256.52,
      "Facturación_Neta": 11924249.09,
      "Total_Concentrado_($)": 100000,
      "Total_Fructosa_($)": 20000,
      "Total_Gas_Carbónico_($)": 5000,
      "Total_Tapas_($)": 3000,
      "Total_Etiquetas_($)": 2000,
      "Total_Preformas_($)": 8000,
      "Total_Termocontraible_($)": 1000,
      "Total_Stretch_($)": 500,
      "Total_Separadores_($)": 300,
      "Total_Caja_BIB_($)": 0,
      "Total_Bolsa_BIB_($)": 0,
      "Total_Nitrógeno_($)": 0,
      "TOTAL_INSUMOS": 139800,
      "Precio_Prod._Reventa": 0,
      "Incidencia_de_Reventa": 0,
      "Ajuste_de_concentrado_por_compra": 0,
      "TOTAL_PRODUCTO_TERCEROS": 0,
      "AJUSTE_DE_STOCK_(Cajas_Físicas)": 0,
      "COSTO_AJUSTES_DE_STOCK": 0,
      "_mes": "2026-07"
    },
    {
      "FECHA": "2026-07-01 00:00:00",
      "Tipo_de_Prod_": "P",
      "MARCA": "COCA-COLA",
      "LOCACION___SAP": "Org.Vtas Ranelagh",
      "ART": "110",
      "Nombre": "COCA-COLA 350CC VIDRIOX24",
      "Físicos": 3092,
      "Unit": 4573.068,
      "Facturación_Lista": 77629980.33,
      "Facturación_Neta": 67711004.86,
      "Total_Concentrado_($)": 250000,
      "Total_Fructosa_($)": 45000,
      "Total_Gas_Carbónico_($)": 8000,
      "Total_Tapas_($)": 12000,
      "Total_Etiquetas_($)": 6000,
      "Total_Preformas_($)": 0,
      "Total_Termocontraible_($)": 2000,
      "Total_Stretch_($)": 1000,
      "Total_Separadores_($)": 500,
      "Total_Caja_BIB_($)": 0,
      "Total_Bolsa_BIB_($)": 0,
      "Total_Nitrógeno_($)": 0,
      "TOTAL_INSUMOS": 324500,
      "Precio_Prod._Reventa": 0,
      "Incidencia_de_Reventa": 0,
      "Ajuste_de_concentrado_por_compra": 0,
      "TOTAL_PRODUCTO_TERCEROS": 0,
      "AJUSTE_DE_STOCK_(Cajas_Físicas)": 0,
      "COSTO_AJUSTES_DE_STOCK": 0,
      "_mes": "2026-07"
    },
    {
      "FECHA": "2026-08-01 00:00:00",
      "Tipo_de_Prod_": "P",
      "MARCA": "COCA-COLA",
      "LOCACION___SAP": "Org.Vtas Ranelagh",
      "ART": "105",
      "Nombre": "COCA-COLA 250CC PETX12",
      "Físicos": 1800,
      "Unit": 850.0,
      "Facturación_Lista": 14000000,
      "Facturación_Neta": 13000000,
      "Total_Concentrado_($)": 110000,
      "Total_Fructosa_($)": 22000,
      "Total_Gas_Carbónico_($)": 5500,
      "Total_Tapas_($)": 3300,
      "Total_Etiquetas_($)": 2200,
      "Total_Preformas_($)": 8800,
      "Total_Termocontraible_($)": 1100,
      "Total_Stretch_($)": 550,
      "Total_Separadores_($)": 330,
      "Total_Caja_BIB_($)": 0,
      "Total_Bolsa_BIB_($)": 0,
      "Total_Nitrógeno_($)": 0,
      "TOTAL_INSUMOS": 153780,
      "Precio_Prod._Reventa": 0,
      "Incidencia_de_Reventa": 0,
      "Ajuste_de_concentrado_por_compra": 0,
      "TOTAL_PRODUCTO_TERCEROS": 0,
      "AJUSTE_DE_STOCK_(Cajas_Físicas)": 0,
      "COSTO_AJUSTES_DE_STOCK": 0,
      "_mes": "2026-08"
    },
    {
      "FECHA": "2026-08-01 00:00:00",
      "Tipo_de_Prod_": "P",
      "MARCA": "FANTA",
      "LOCACION___SAP": "Org.Vtas Ranelagh",
      "ART": "310",
      "Nombre": "FANTA NARANJA 350CC VIDRIOX24",
      "Físicos": 324,
      "Unit": 479.196,
      "Facturación_Lista": 8165469.11,
      "Facturación_Neta": 7088360.90,
      "Total_Concentrado_($)": 80000,
      "Total_Fructosa_($)": 15000,
      "Total_Gas_Carbónico_($)": 3000,
      "Total_Tapas_($)": 4000,
      "Total_Etiquetas_($)": 2500,
      "Total_Preformas_($)": 0,
      "Total_Termocontraible_($)": 800,
      "Total_Stretch_($)": 400,
      "Total_Separadores_($)": 200,
      "Total_Caja_BIB_($)": 0,
      "Total_Bolsa_BIB_($)": 0,
      "Total_Nitrógeno_($)": 0,
      "TOTAL_INSUMOS": 105900,
      "Precio_Prod._Reventa": 0,
      "Incidencia_de_Reventa": 0,
      "Ajuste_de_concentrado_por_compra": 0,
      "TOTAL_PRODUCTO_TERCEROS": 0,
      "AJUSTE_DE_STOCK_(Cajas_Físicas)": 0,
      "COSTO_AJUSTES_DE_STOCK": 0,
      "_mes": "2026-08"
    },
    {
      "FECHA": "2026-09-01 00:00:00",
      "Tipo_de_Prod_": "r",
      "MARCA": "COCA-COLA",
      "LOCACION___SAP": "Org.Vtas Dolores",
      "ART": "180",
      "Nombre": "COCA-COLA 10L BIB",
      "Físicos": 82,
      "Unit": 924.304,
      "Facturación_Lista": 10952205.36,
      "Facturación_Neta": 10952205.36,
      "Total_Concentrado_($)": 0,
      "Total_Fructosa_($)": 0,
      "Total_Gas_Carbónico_($)": 0,
      "Total_Tapas_($)": 0,
      "Total_Etiquetas_($)": 0,
      "Total_Preformas_($)": 0,
      "Total_Termocontraible_($)": 0,
      "Total_Stretch_($)": 0,
      "Total_Separadores_($)": 0,
      "Total_Caja_BIB_($)": 0,
      "Total_Bolsa_BIB_($)": 0,
      "Total_Nitrógeno_($)": 0,
      "TOTAL_INSUMOS": 0,
      "Precio_Prod._Reventa": 0,
      "Incidencia_de_Reventa": 0,
      "Ajuste_de_concentrado_por_compra": 0,
      "TOTAL_PRODUCTO_TERCEROS": 9500000,
      "AJUSTE_DE_STOCK_(Cajas_Físicas)": 0,
      "COSTO_AJUSTES_DE_STOCK": 0,
      "_mes": "2026-09"
    }
  ];

  // Si no se carga ningún archivo, los datos de ejemplo se cargan automáticamente al inicio
  // (opcional: descomentar la siguiente línea para cargar demo automáticamente)
  // setTimeout(() => { if (datosOriginales.length === 0) loadDemoBtn.click(); }, 500);
});
