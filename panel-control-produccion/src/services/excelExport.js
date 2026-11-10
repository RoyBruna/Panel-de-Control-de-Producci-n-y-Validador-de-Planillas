const XLSX = require('xlsx');

// Generación de reporte Excel con hojas Resumen y Detalle
function generateExcelReport(metricas, detalle, filtros = {}) {
  const wb = XLSX.utils.book_new();

  // 1. Hoja RESUMEN
  const resumenRows = [
    ['REPORTE EJECUTIVO DE PRODUCCIÓN'],
    ['Generado:', new Date().toISOString().replace('T', ' ').substring(0, 19)],
    ['Filtro Fecha Desde:', filtros.desde || 'Todas'],
    ['Filtro Fecha Hasta:', filtros.hasta || 'Todas'],
    ['Filtro Línea:', filtros.linea || 'Todas'],
    [],
    ['INDICADORES CLAVE DE DESEMPEÑO (KPI)'],
    ['Métrica', 'Valor'],
    ['Total Producido (u)', metricas.total_producido],
    ['Total Defectuoso (u)', metricas.total_defectuoso],
    ['Tasa de Defectos (%)', `${metricas.porcentaje_defectuoso}%`],
    ['Lotes / Registros Procesados', metricas.total_registros],
    [],
    ['DESGLOSE POR TURNO'],
    ['Turno', 'Producido (u)', 'Defectuoso (u)', 'Tasa Defecto (%)']
  ];

  for (const t of metricas.produccion_por_turno) {
    resumenRows.push([t.turno, t.total_producido, t.total_defectuoso, `${t.pct_defectuoso}%`]);
  }

  resumenRows.push([]);
  resumenRows.push(['TOP PRODUCTOS DE MAYOR VOLUMEN']);
  resumenRows.push(['Producto', 'Producido (u)', 'Defectuoso (u)', 'Tasa Defecto (%)']);

  for (const p of metricas.top_productos) {
    resumenRows.push([p.producto, p.total_producido, p.total_defectuoso, `${p.pct_defectuoso}%`]);
  }

  const wsResumen = XLSX.utils.aoa_to_sheet(resumenRows);
  wsResumen['!cols'] = [
    { wch: 32 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 }
  ];

  // 2. Hoja DETALLE
  const detalleHeaders = [
    'ID',
    'Fecha',
    'Turno',
    'Línea',
    'Producto',
    'Cant. Producida',
    'Cant. Defectuosa',
    '% Defectuoso',
    'Operario'
  ];

  const detalleRows = detalle.map(r => [
    r.id,
    r.fecha,
    r.turno,
    r.linea,
    r.producto,
    r.cantidad_producida,
    r.cantidad_defectuosa,
    r.pct_defectuoso !== undefined ? `${r.pct_defectuoso}%` : '0.00%',
    r.operario
  ]);

  const wsDetalle = XLSX.utils.aoa_to_sheet([detalleHeaders, ...detalleRows]);
  wsDetalle['!cols'] = [
    { wch: 8 },
    { wch: 14 },
    { wch: 12 },
    { wch: 18 },
    { wch: 24 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 22 }
  ];

  XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');
  XLSX.utils.book_append_sheet(wb, wsDetalle, 'Detalle');

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

module.exports = {
  generateExcelReport
};

