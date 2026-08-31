document.addEventListener('DOMContentLoaded', function () {
  // ... [Referencias DOM iguales a tu HTML actualizado] ...
  const loadRepoBtn = document.getElementById('loadRepoBtn');
  const fileStatus = document.getElementById('fileStatus');

  let state = { ventas: [], datosFiltrados: [] };
  let chartVentas = null;
  let chartComposicion = null;

  // ========== FUNCIONES DE LECTURA Y PROCESO ==========
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

  // ========== FUNCIÓN PARA CARGAR DESDE EL REPOSITORIO ==========
  async function cargarDesdeRepo() {
    fileStatus.textContent = '⏳ Intentando cargar ventas.xlsx desde el repositorio...';
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

      // OCULTAR LA SECCIÓN DE CARGA Y MOSTRAR EL DASHBOARD
      document.getElementById('uploadArea').style.display = 'none';
      document.getElementById('filtrosSection').style.display = 'flex';
      document.getElementById('resumenSection').style.display = 'grid';
      document.getElementById('graficosSection').style.display = 'grid';
      document.getElementById('tablaSection').style.display = 'block';

      poblarFiltros(state.ventas);
      aplicarFiltros();
      fileStatus.textContent = '✅ Datos cargados automáticamente desde ventas.xlsx';

    } catch (error) {
      fileStatus.textContent = '❌ No se pudo cargar desde el repo. Sube el archivo manualmente.';
      console.error('Error de CORS o archivo no encontrado:', error);
    }
  }

  // Botón manual para cargar
  loadRepoBtn.addEventListener('click', cargarDesdeRepo);

  // ========== SUBIR ARCHIVO MANUALMENTE (RESPALDO) ==========
  // ... [Tu código de input type="file" para procesar el archivo local] ...
  
  // ========== INTENTAR CARGA AUTOMÁTICA AL ABRIR ==========
  cargarDesdeRepo();

  // ... [Resto de las funciones: poblarFiltros, aplicarFiltros, actualizarDashboard, etc. igual que antes] ...
});
