// app.js - VERSIÓN 2.0 (Con Fecha robusta, Filtro Artículo y más KPIs)
document.addEventListener('DOMContentLoaded', function () {
  // ... [Referencias DOM iguales a las del HTML anterior más las nuevas] ...
  const selectArticulo = document.getElementById('selectArticulo');
  const promedioUnidadesEl = document.getElementById('promedioUnidades');
  const margenUnidadEl = document.getElementById('margenUnidad');
  const totalRegistrosEl = document.getElementById('totalRegistros');

  let state = { ventas: [], datosFiltrados: [] };
  let chartVentas = null;
  let chartComposicion = null;

  // ========== FUNCIÓN DE FECHA ULTRA ROBUSTA ==========
  function obtenerMes(valor) {
    if (!valor) return 'Sin Mes';
    // 1. Si es número serial de Excel (ej. 45000)
    if (typeof valor === 'number') {
      const date = new Date(Math.round((valor - 25569) * 86400 * 1000));
      return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
    }
    // 2. Si es texto "jul-26"
    const str = String(valor).trim();
    const meses = { ene: '01', feb: '02', mar: '03', abr: '04', may: '05', jun: '06', jul: '07', ago: '08', sep: '09', oct: '10', nov: '11', dic: '12' };
    
    if (str.includes('-')) {
      const parts = str.toLowerCase().split('-');
      if (meses[parts[0].substring(0,3)] && parts[1]) {
        return '20' + parts[1] + '-' + meses[parts[0].substring(0,3)];
      }
    }
    // 3. Si es texto "dd/mm/aaaa"
    if (str.includes('/')) {
      const p = str.split('/');
      if (p.length === 3 && p[2].length === 4) return p[2] + '-' + p[1];
    }
    return 'Sin Mes';
  }

  // ========== PROCESAR DATOS (Detecta encabezados en cualquier fila) ==========
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
    const rows = dataArray.slice(headerIndex + 1); // Debajo de los títulos

    const objetos = rows.map(row => {
      const obj = {};
      headers.forEach((key, idx) => obj[key] = row[idx] !== undefined ? row[idx] : '');
      return obj;
    }).filter(o => o.ART && o.ART !== '');

    // Números
    const numCols = ['Físicos', 'Facturación_Neta', 'TOTAL_INSUMOS', 'COSTO_AJUSTES_DE_STOCK', 'TOTAL_PRODUCTO_TERCEROS', 'Total_Concentrado_', 'Total_Fructosa_', 'Total_Gas_Carbónico_', 'Total_Tapas_'];
    objetos.forEach(o => {
      numCols.forEach(col => o[col] = parseFloat(o[col]) || 0);
      o._mes = obtenerMes(o.FECHA); // ¡Corregido!
    });
    return objetos;
  }

  // ... (Función de carga de archivos igual, pero actualizando los nuevos KPIs al final) ...
  
  function poblarFiltros(datos) {
    // Artículos
    const artSet = new Set(datos.map(o => o.ART + ' - ' + o.Nombre).filter(Boolean));
    selectArticulo.innerHTML = '<option value="todos">Todos los Artículos</option>';
    Array.from(artSet).sort().forEach(a => selectArticulo.innerHTML += `<option value="${a.split(' - ')[0]}">${a}</option>`);

    // Meses
    const mesesSet = new Set(datos.map(o => o._mes).filter(m => m !== 'Sin Mes'));
    selectMes.innerHTML = '<option value="todos">Todos los meses</option>';
    Array.from(mesesSet).sort().forEach(m => selectMes.innerHTML += `<option value="${m}">${nombreMes(m)}</option>`);

    // Ubicaciones
    const ubicSet = new Set(datos.map(o => o.LOCACION_SAP).filter(Boolean));
    selectUbicacion.innerHTML = '<option value="todas">Todas</option>';
    Array.from(ubicSet).sort().forEach(u => selectUbicacion.innerHTML += `<option value="${u}">${u}</option>`);
    
    // Eventos
    selectArticulo.addEventListener('change', aplicarFiltros);
    selectMes.addEventListener('change', aplicarFiltros);
    selectUbicacion.addEventListener('change', aplicarFiltros);
    selectTipo.addEventListener('change', aplicarFiltros);
  }

  // ========== NUEVOS KPIs Y GRÁFICOS ==========
  function actualizarDashboard() {
    const datos = state.datosFiltrados;
    if (datos.length === 0) { /* Resetear valores a 0 */ return; }

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

    // Actualizar DOM
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

  function actualizarGraficos(datos) {
    destruirGraficos();
    // 1. Evolución Mensual
    const ventasPorMes = {};
    datos.forEach(o => ventasPorMes[o._mes] = (ventasPorMes[o._mes] || 0) + (o.Físicos || 0));
    const meses = Object.keys(ventasPorMes).sort();
    // ... Crear gráfico de línea aquí ...

    // 2. Composición del Costo
    const insumos = ['Total_Concentrado_', 'Total_Fructosa_', 'Total_Gas_Carbónico_', 'Total_Tapas_'];
    const labels = ['Concentrado','Fructosa','Gas','Tapas'];
    const dataComp = insumos.map(col => datos.reduce((acc, o) => acc + (o[col] || 0), 0));
    // ... Crear gráfico de dona aquí (esto hará que si filtras un artículo, veas su composición exacta) ...
  }
  async function cargarDesdeRepo() {
  const response = await fetch('data.xlsx'); // Asegúrate de que el archivo se llame data.xlsx y esté en la raíz
  const arrayBuffer = await response.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const dataArray = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  // Procesar igual que en el cambio de archivo...
}
  // ... (Resto del código: Exportar CSV, PDF, Dark Mode)
});
