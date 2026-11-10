const path = require('path');
const fs = require('fs');

let DatabaseSync;
try {
  ({ DatabaseSync } = require('node:sqlite'));
} catch (e) {
  try {
    DatabaseSync = require('better-sqlite3');
  } catch (err) {
    throw new Error('SQLite driver not found. Node.js >= 22.5.0 with node:sqlite or better-sqlite3 is required.');
  }
}

const dbDir = path.join(__dirname, '..', 'data');
if (!fs.existsSync(dbDir)) {
  fs.mkdirSync(dbDir, { recursive: true });
}

const dbPath = path.join(dbDir, 'produccion.db');
const db = new DatabaseSync(dbPath);

// Configuración de integridad y rendimiento para SQLite
db.exec('PRAGMA foreign_keys = ON;');
db.exec('PRAGMA journal_mode = WAL;');

function initSchema() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS lineas (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nombre TEXT UNIQUE NOT NULL COLLATE NOCASE
    );

    CREATE TABLE IF NOT EXISTS productos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nombre TEXT UNIQUE NOT NULL COLLATE NOCASE
    );

    CREATE TABLE IF NOT EXISTS produccion (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      fecha TEXT NOT NULL,
      turno TEXT NOT NULL,
      linea_id INTEGER NOT NULL REFERENCES lineas(id),
      producto_id INTEGER NOT NULL REFERENCES productos(id),
      cantidad_producida INTEGER NOT NULL,
      cantidad_defectuosa INTEGER NOT NULL,
      operario TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE INDEX IF NOT EXISTS idx_produccion_fecha ON produccion(fecha);
    CREATE INDEX IF NOT EXISTS idx_produccion_linea ON produccion(linea_id);
    CREATE INDEX IF NOT EXISTS idx_produccion_producto ON produccion(producto_id);
  `);
}

initSchema();

const stmtGetLinea = db.prepare('SELECT id FROM lineas WHERE nombre = ?');
const stmtInsertLinea = db.prepare('INSERT INTO lineas (nombre) VALUES (?)');

function getOrCreateLinea(nombre) {
  const clean = nombre.trim();
  const existing = stmtGetLinea.get(clean);
  if (existing) return existing.id;
  const res = stmtInsertLinea.run(clean);
  return Number(res.lastInsertRowid);
}

const stmtGetProducto = db.prepare('SELECT id FROM productos WHERE nombre = ?');
const stmtInsertProducto = db.prepare('INSERT INTO productos (nombre) VALUES (?)');

function getOrCreateProducto(nombre) {
  const clean = nombre.trim();
  const existing = stmtGetProducto.get(clean);
  if (existing) return existing.id;
  const res = stmtInsertProducto.run(clean);
  return Number(res.lastInsertRowid);
}

const stmtInsertProduccion = db.prepare(`
  INSERT INTO produccion (fecha, turno, linea_id, producto_id, cantidad_producida, cantidad_defectuosa, operario)
  VALUES (?, ?, ?, ?, ?, ?, ?)
`);

// Inserción transaccional atómica de lotes de producción
function insertBatch(records) {
  if (!records || records.length === 0) return 0;

  db.exec('BEGIN TRANSACTION;');
  try {
    for (const r of records) {
      const lineaId = getOrCreateLinea(r.linea);
      const productoId = getOrCreateProducto(r.producto);

      stmtInsertProduccion.run(
        r.fecha,
        r.turno,
        lineaId,
        productoId,
        r.cantidad_producida,
        r.cantidad_defectuosa,
        r.operario
      );
    }
    db.exec('COMMIT;');
    return records.length;
  } catch (err) {
    db.exec('ROLLBACK;');
    throw err;
  }
}

function buildWhereClause({ desde, hasta, linea }) {
  const conditions = [];
  const params = [];

  if (desde) {
    conditions.push('p.fecha >= ?');
    params.push(desde);
  }
  if (hasta) {
    conditions.push('p.fecha <= ?');
    params.push(hasta);
  }
  if (linea) {
    conditions.push('l.nombre = ?');
    params.push(linea);
  }

  const whereStr = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
  return { whereStr, params };
}

function getMetricas(filtros = {}) {
  const { whereStr, params } = buildWhereClause(filtros);

  // Totales generales
  const queryTotales = `
    SELECT 
      COALESCE(SUM(p.cantidad_producida), 0) AS total_producido,
      COALESCE(SUM(p.cantidad_defectuosa), 0) AS total_defectuoso,
      COUNT(p.id) AS total_registros
    FROM produccion p
    JOIN lineas l ON p.linea_id = l.id
    JOIN productos pr ON p.producto_id = pr.id
    ${whereStr}
  `;
  const totales = db.prepare(queryTotales).get(...params);
  const totalProducido = Number(totales.total_producido);
  const totalDefectuoso = Number(totales.total_defectuoso);
  const pctDefectuoso = totalProducido > 0 ? ((totalDefectuoso / totalProducido) * 100).toFixed(2) : '0.00';

  // Top 3 productos
  const queryTopProductos = `
    SELECT 
      pr.nombre AS producto,
      SUM(p.cantidad_producida) AS total_producido,
      SUM(p.cantidad_defectuosa) AS total_defectuoso
    FROM produccion p
    JOIN lineas l ON p.linea_id = l.id
    JOIN productos pr ON p.producto_id = pr.id
    ${whereStr}
    GROUP BY pr.id, pr.nombre
    ORDER BY total_producido DESC
    LIMIT 3
  `;
  const topProductos = db.prepare(queryTopProductos).all(...params).map(row => ({
    producto: row.producto,
    total_producido: Number(row.total_producido),
    total_defectuoso: Number(row.total_defectuoso),
    pct_defectuoso: row.total_producido > 0 ? ((row.total_defectuoso / row.total_producido) * 100).toFixed(2) : '0.00'
  }));

  // Distribución por turno
  const queryTurnos = `
    SELECT 
      p.turno,
      SUM(p.cantidad_producida) AS total_producido,
      SUM(p.cantidad_defectuosa) AS total_defectuoso
    FROM produccion p
    JOIN lineas l ON p.linea_id = l.id
    JOIN productos pr ON p.producto_id = pr.id
    ${whereStr}
    GROUP BY p.turno
    ORDER BY total_producido DESC
  `;
  const porTurno = db.prepare(queryTurnos).all(...params).map(row => ({
    turno: row.turno,
    total_producido: Number(row.total_producido),
    total_defectuoso: Number(row.total_defectuoso),
    pct_defectuoso: row.total_producido > 0 ? ((row.total_defectuoso / row.total_producido) * 100).toFixed(2) : '0.00'
  }));

  // Serie temporal diaria para gráfico de evolución
  const querySerie = `
    SELECT 
      p.fecha,
      SUM(p.cantidad_producida) AS total_producido,
      SUM(p.cantidad_defectuosa) AS total_defectuoso
    FROM produccion p
    JOIN lineas l ON p.linea_id = l.id
    JOIN productos pr ON p.producto_id = pr.id
    ${whereStr}
    GROUP BY p.fecha
    ORDER BY p.fecha ASC
  `;
  const serieTemporal = db.prepare(querySerie).all(...params).map(row => ({
    fecha: row.fecha,
    total_producido: Number(row.total_producido),
    total_defectuoso: Number(row.total_defectuoso)
  }));

  return {
    total_producido: totalProducido,
    total_defectuoso: totalDefectuoso,
    porcentaje_defectuoso: Number(pctDefectuoso),
    total_registros: Number(totales.total_registros),
    top_productos: topProductos,
    produccion_por_turno: porTurno,
    serie_temporal: serieTemporal
  };
}

function getDetalle(filtros = {}, limit = 200) {
  const { whereStr, params } = buildWhereClause(filtros);
  const query = `
    SELECT 
      p.id,
      p.fecha,
      p.turno,
      l.nombre AS linea,
      pr.nombre AS producto,
      p.cantidad_producida,
      p.cantidad_defectuosa,
      p.operario
    FROM produccion p
    JOIN lineas l ON p.linea_id = l.id
    JOIN productos pr ON p.producto_id = pr.id
    ${whereStr}
    ORDER BY p.fecha DESC, p.id DESC
    LIMIT ?
  `;
  return db.prepare(query).all(...params, limit);
}

function getAllDetalleForExport(filtros = {}) {
  const { whereStr, params } = buildWhereClause(filtros);
  const query = `
    SELECT 
      p.id,
      p.fecha,
      p.turno,
      l.nombre AS linea,
      pr.nombre AS producto,
      p.cantidad_producida,
      p.cantidad_defectuosa,
      ROUND((CAST(p.cantidad_defectuosa AS FLOAT) / p.cantidad_producida) * 100, 2) AS pct_defectuoso,
      p.operario
    FROM produccion p
    JOIN lineas l ON p.linea_id = l.id
    JOIN productos pr ON p.producto_id = pr.id
    ${whereStr}
    ORDER BY p.fecha DESC, p.id ASC
  `;
  return db.prepare(query).all(...params);
}

function getLineas() {
  const rows = db.prepare('SELECT DISTINCT nombre FROM lineas ORDER BY nombre ASC').all();
  return rows.map(r => r.nombre);
}

module.exports = {
  db,
  insertBatch,
  getMetricas,
  getDetalle,
  getAllDetalleForExport,
  getLineas
};

