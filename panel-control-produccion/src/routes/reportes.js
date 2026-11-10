const express = require('express');
const { getMetricas, getDetalle, getAllDetalleForExport, getLineas } = require('../db');
const { generateExcelReport } = require('../services/excelExport');

const router = express.Router();

router.get('/metricas', (req, res) => {
  try {
    const { desde, hasta, linea } = req.query;
    const metricas = getMetricas({ desde, hasta, linea });
    return res.json(metricas);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

router.get('/detalle', (req, res) => {
  try {
    const { desde, hasta, linea, limit } = req.query;
    const parsedLimit = limit ? parseInt(limit, 10) : 50;
    const detalle = getDetalle({ desde, hasta, linea }, parsedLimit);
    return res.json(detalle);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

router.get('/lineas', (req, res) => {
  try {
    const lineas = getLineas();
    return res.json(lineas);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

router.get('/exportar', (req, res) => {
  try {
    const { desde, hasta, linea } = req.query;
    const filtros = { desde, hasta, linea };
    const metricas = getMetricas(filtros);
    const detalle = getAllDetalleForExport(filtros);

    const buffer = generateExcelReport(metricas, detalle, filtros);
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const filename = `reporte_produccion_${timestamp}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    return res.send(buffer);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

module.exports = router;

