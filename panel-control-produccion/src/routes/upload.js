const express = require('express');
const multer = require('multer');
const { parseBuffer } = require('../services/excelParser');
const { insertBatch } = require('../db');

const router = express.Router();
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB
});

router.post('/', upload.single('archivo'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No se envió ningún archivo de planilla Excel.' });
    }

    const { totalFilas, validas, rechazadas, validRows, invalidRows } = parseBuffer(req.file.buffer);

    let insertados = 0;
    if (validRows.length > 0) {
      insertados = insertBatch(validRows);
    }

    return res.json({
      success: true,
      resumen: {
        totalFilas,
        procesadas: insertados,
        rechazadas
      },
      errores: invalidRows
    });
  } catch (err) {
    return res.status(400).json({ error: err.message || 'Error procesando la planilla de Excel.' });
  }
});

module.exports = router;

