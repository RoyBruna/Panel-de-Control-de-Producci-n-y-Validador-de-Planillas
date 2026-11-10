const XLSX = require('xlsx');

function normalizeKey(key) {
  return String(key)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\s\-_]+/g, '_');
}

// Mapeo flexible de alias de columnas a nombres canónicos
const COLUMN_ALIASES = {
  fecha: ['fecha', 'date', 'fec', 'dia'],
  turno: ['turno', 'shift', 'horario'],
  linea: ['linea', 'line', 'linea_produccion', 'linea_prod'],
  producto: ['producto', 'product', 'item', 'articulo', 'codigo_producto', 'pieza'],
  cantidad_producida: ['cantidad_producida', 'cant_producida', 'producido', 'produccion', 'unidades_producidas', 'cant_prod', 'qty_produced'],
  cantidad_defectuosa: ['cantidad_defectuosa', 'cant_defectuosa', 'defectuoso', 'defectos', 'unidades_defectuosas', 'cant_def', 'qty_defective', 'scrap'],
  operario: ['operario', 'operador', 'operator', 'responsable', 'empleado', 'legajo_operario']
};

function resolveCanonicalRow(rawObj) {
  const row = {};
  const rawKeys = Object.keys(rawObj);

  for (const rawKey of rawKeys) {
    const normKey = normalizeKey(rawKey);
    for (const [canonical, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (aliases.includes(normKey)) {
        row[canonical] = rawObj[rawKey];
        break;
      }
    }
  }

  return row;
}

function parseExcelDate(val) {
  if (val === undefined || val === null || val === '') return null;

  if (typeof val === 'number') {
    const parsed = XLSX.SSF.parse_date_code(val);
    if (!parsed) return null;
    const y = parsed.y;
    const m = String(parsed.m).padStart(2, '0');
    const d = String(parsed.d).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  if (val instanceof Date && !isNaN(val)) {
    return val.toISOString().split('T')[0];
  }

  const str = String(val).trim();
  // Formato ISO YYYY-MM-DD o YYYY/MM/DD
  const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const y = isoMatch[1];
    const m = isoMatch[2].padStart(2, '0');
    const d = isoMatch[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // Formato DD-MM-YYYY o DD/MM/YYYY
  const dmyMatch = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmyMatch) {
    const d = dmyMatch[1].padStart(2, '0');
    const m = dmyMatch[2].padStart(2, '0');
    const y = dmyMatch[3];
    return `${y}-${m}-${d}`;
  }

  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    return d.toISOString().split('T')[0];
  }

  return null;
}

// Validación y sanitización rigurosa de cada fila de la planilla
function parseBuffer(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error('El archivo no contiene hojas de cálculo.');
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: null });

  if (rawRows.length === 0) {
    throw new Error('La planilla se encuentra vacía.');
  }

  const validRows = [];
  const invalidRows = [];

  rawRows.forEach((rawRow, index) => {
    const rowNum = index + 2; // Fila Excel considerando encabezado en fila 1
    const canonical = resolveCanonicalRow(rawRow);
    const errors = [];

    // 1. Fecha
    const fecha = parseExcelDate(canonical.fecha);
    if (!fecha) {
      errors.push(`Fecha inválida o ausente ('${canonical.fecha ?? ''}')`);
    }

    // 2. Turno
    const turno = canonical.turno ? String(canonical.turno).trim() : '';
    if (!turno) {
      errors.push('Turno obligatorio no especificado');
    }

    // 3. Línea
    const linea = canonical.linea ? String(canonical.linea).trim() : '';
    if (!linea) {
      errors.push('Línea de producción obligatoria no especificada');
    }

    // 4. Producto
    const producto = canonical.producto ? String(canonical.producto).trim() : '';
    if (!producto) {
      errors.push('Producto obligatorio no especificado');
    }

    // 5. Cantidad producida
    const cantProdRaw = canonical.cantidad_producida;
    const cantProd = Number(cantProdRaw);
    if (cantProdRaw === null || cantProdRaw === '' || isNaN(cantProd) || !Number.isInteger(cantProd) || cantProd < 0) {
      errors.push(`Cantidad producida inválida ('${cantProdRaw ?? ''}'). Debe ser un entero mayor o igual a 0`);
    }

    // 6. Cantidad defectuosa
    const cantDefRaw = canonical.cantidad_defectuosa;
    const cantDef = Number(cantDefRaw);
    if (cantDefRaw === null || cantDefRaw === '' || isNaN(cantDef) || !Number.isInteger(cantDef) || cantDef < 0) {
      errors.push(`Cantidad defectuosa inválida ('${cantDefRaw ?? ''}'). Debe ser un entero mayor o igual a 0`);
    } else if (!isNaN(cantProd) && cantDef > cantProd) {
      errors.push(`Inconsistencia: Cantidad defectuosa (${cantDef}) excede la producida (${cantProd})`);
    }

    // 7. Operario
    const operario = canonical.operario ? String(canonical.operario).trim() : '';
    if (!operario) {
      errors.push('Operario responsable obligatorio no especificado');
    }

    if (errors.length > 0) {
      invalidRows.push({
        fila: rowNum,
        datos: rawRow,
        errores: errors
      });
    } else {
      validRows.push({
        fecha,
        turno,
        linea,
        producto,
        cantidad_producida: cantProd,
        cantidad_defectuosa: cantDef,
        operario
      });
    }
  });

  return {
    totalFilas: rawRows.length,
    validas: validRows.length,
    rechazadas: invalidRows.length,
    validRows,
    invalidRows
  };
}

module.exports = {
  parseBuffer,
  parseExcelDate,
  resolveCanonicalRow
};

