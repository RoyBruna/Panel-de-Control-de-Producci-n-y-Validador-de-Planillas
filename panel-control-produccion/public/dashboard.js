// Lógica de cliente para visualización de KPIs, control de filtros y gráficos
let chartEvolucion = null;
let chartTurnos = null;

const currentFiltros = {
  desde: '',
  hasta: '',
  linea: ''
};

function formatNumber(num) {
  return new Intl.NumberFormat('es-AR').format(num || 0);
}

document.addEventListener('DOMContentLoaded', () => {
  setupUploadEvents();
  setupFilterEvents();
  loadLineas();
  loadDashboardData();
});

// Carga del selector de líneas de producción
async function loadLineas() {
  try {
    const res = await fetch('/api/reportes/lineas');
    const lineas = await res.json();
    const select = document.getElementById('filtro-linea');
    select.innerHTML = '<option value="">Todas las líneas</option>';
    lineas.forEach(l => {
      const opt = document.createElement('option');
      opt.value = l;
      opt.textContent = l;
      select.appendChild(opt);
    });
  } catch (err) {
    console.error('Error al cargar líneas:', err);
  }
}

// Configuración de zona de carga de archivos (drag & drop y botón)
function setupUploadEvents() {
  const form = document.getElementById('upload-form');
  const fileInput = document.getElementById('file-input');
  const dropzone = document.getElementById('dropzone');
  const label = document.getElementById('dropzone-label');
  const btnUpload = document.getElementById('btn-upload');

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
      label.textContent = `Archivo seleccionado: ${fileInput.files[0].name}`;
      btnUpload.disabled = false;
    } else {
      label.textContent = 'Seleccionar archivo Excel o arrastrar aquí';
      btnUpload.disabled = true;
    }
  });

  ['dragenter', 'dragover'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files.length > 0) {
      fileInput.files = e.dataTransfer.files;
      label.textContent = `Archivo seleccionado: ${fileInput.files[0].name}`;
      btnUpload.disabled = false;
    }
  });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!fileInput.files.length) return;

    const file = fileInput.files[0];
    const formData = new FormData();
    formData.append('archivo', file);

    btnUpload.disabled = true;
    btnUpload.textContent = 'Procesando...';

    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();

      if (!res.ok) {
        showValidationAlert('error', 'Error en el procesamiento', data.error || 'Ocurrió un error al procesar el archivo.');
      } else {
        renderValidationResult(data);
        await loadLineas();
        await loadDashboardData();
        form.reset();
        label.textContent = 'Seleccionar archivo Excel o arrastrar aquí';
      }
    } catch (err) {
      showValidationAlert('error', 'Fallo de comunicación', 'No se pudo conectar con el servidor.');
    } finally {
      btnUpload.disabled = false;
      btnUpload.textContent = 'Procesar Planilla';
    }
  });
}

function showValidationAlert(type, title, description) {
  const container = document.getElementById('alert-container');
  const section = document.getElementById('validation-report');

  container.className = `alert-box ${type}`;
  container.innerHTML = `
    <h3 class="alert-title">${title}</h3>
    <p class="alert-desc">${description}</p>
  `;
  section.classList.remove('hidden');
}

function renderValidationResult(data) {
  const { totalFilas, procesadas, rechazadas } = data.resumen;
  const container = document.getElementById('alert-container');
  const section = document.getElementById('validation-report');

  if (rechazadas === 0) {
    container.className = 'alert-box success';
    container.innerHTML = `
      <h3 class="alert-title">Carga completada con éxito</h3>
      <p class="alert-desc">Se importaron y persistieron ${procesadas} registros de ${totalFilas} filas sin inconsistencias.</p>
    `;
  } else {
    container.className = 'alert-box warning';
    let rowsHtml = '';
    data.errores.forEach(err => {
      const motivo = err.errores.join('; ');
      rowsHtml += `
        <tr>
          <td style="font-family: monospace; font-weight: 600;">Fila ${err.fila}</td>
          <td style="color: #b91c1c;">${motivo}</td>
          <td>${escapeHtml(JSON.stringify(err.datos))}</td>
        </tr>
      `;
    });

    container.innerHTML = `
      <h3 class="alert-title">Carga completada con observaciones</h3>
      <p class="alert-desc">Se procesaron <strong>${procesadas}</strong> registros válidos. <strong>${rechazadas}</strong> filas fueron rechazadas por violar reglas de validación:</p>
      <div class="error-table-wrapper">
        <table class="error-table">
          <thead>
            <tr>
              <th>Ubicación</th>
              <th>Motivo de Rechazo</th>
              <th>Datos Originales</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    `;
  }

  section.classList.remove('hidden');
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Configuración de filtros y exportación
function setupFilterEvents() {
  const filterForm = document.getElementById('filter-form');
  const btnReset = document.getElementById('btn-reset-filters');
  const btnExportar = document.getElementById('btn-exportar');

  filterForm.addEventListener('submit', (e) => {
    e.preventDefault();
    currentFiltros.desde = document.getElementById('filtro-desde').value;
    currentFiltros.hasta = document.getElementById('filtro-hasta').value;
    currentFiltros.linea = document.getElementById('filtro-linea').value;
    loadDashboardData();
  });

  btnReset.addEventListener('click', () => {
    filterForm.reset();
    currentFiltros.desde = '';
    currentFiltros.hasta = '';
    currentFiltros.linea = '';
    loadDashboardData();
  });

  btnExportar.addEventListener('click', () => {
    const params = new URLSearchParams();
    if (currentFiltros.desde) params.append('desde', currentFiltros.desde);
    if (currentFiltros.hasta) params.append('hasta', currentFiltros.hasta);
    if (currentFiltros.linea) params.append('linea', currentFiltros.linea);
    window.location.href = `/api/reportes/exportar?${params.toString()}`;
  });
}

// Carga centralizada de métricas, gráficos y grilla de detalle
async function loadDashboardData() {
  const params = new URLSearchParams();
  if (currentFiltros.desde) params.append('desde', currentFiltros.desde);
  if (currentFiltros.hasta) params.append('hasta', currentFiltros.hasta);
  if (currentFiltros.linea) params.append('linea', currentFiltros.linea);

  try {
    const [resMetricas, resDetalle] = await Promise.all([
      fetch(`/api/reportes/metricas?${params.toString()}`),
      fetch(`/api/reportes/detalle?${params.toString()}&limit=50`)
    ]);

    const metricas = await resMetricas.json();
    const detalle = await resDetalle.json();

    updateKPIs(metricas);
    renderCharts(metricas);
    renderTable(detalle);
  } catch (err) {
    console.error('Error al sincronizar dashboard:', err);
  }
}

function updateKPIs(m) {
  // KPI 1: Volumen Producido
  document.getElementById('kpi-total-producido').textContent = formatNumber(m.total_producido);
  document.getElementById('kpi-total-registros').textContent = `${formatNumber(m.total_registros)} lotes registrados`;

  // KPI 2: Tasa de Descarte
  const scrapVal = document.getElementById('kpi-pct-defectuoso');
  scrapVal.textContent = `${m.porcentaje_defectuoso.toFixed(2)}%`;
  document.getElementById('kpi-total-defectuoso').textContent = `${formatNumber(m.total_defectuoso)} unidades defectuosas`;

  // Indicador de color para descarte
  const badge = document.getElementById('kpi-scrap-badge');
  if (m.porcentaje_defectuoso <= 2.0) {
    badge.className = 'kpi-tag tag-neutral';
  } else if (m.porcentaje_defectuoso <= 4.0) {
    badge.className = 'kpi-tag';
    badge.style.backgroundColor = '#fef3c7';
    badge.style.color = '#92400e';
  } else {
    badge.className = 'kpi-tag tag-defect';
  }

  // KPI 3: Desglose por Turno
  const turnosList = document.getElementById('kpi-turnos-list');
  if (m.produccion_por_turno && m.produccion_por_turno.length > 0) {
    turnosList.innerHTML = m.produccion_por_turno.map(t => `
      <div class="breakdown-row">
        <span class="breakdown-name">${t.turno}</span>
        <span class="breakdown-val">${formatNumber(t.total_producido)} u (${t.pct_defectuoso}%)</span>
      </div>
    `).join('');
  } else {
    turnosList.innerHTML = '<span class="text-muted">Sin registros</span>';
  }

  // KPI 4: Top 3 Productos
  const prodList = document.getElementById('kpi-productos-list');
  if (m.top_productos && m.top_productos.length > 0) {
    prodList.innerHTML = m.top_productos.map((p, idx) => `
      <div class="breakdown-row">
        <span class="breakdown-name">${idx + 1}. ${p.producto}</span>
        <span class="breakdown-val">${formatNumber(p.total_producido)} u</span>
      </div>
    `).join('');
  } else {
    prodList.innerHTML = '<span class="text-muted">Sin registros</span>';
  }
}

function renderCharts(m) {
  // 1. Gráfico de Evolución Diaria
  const ctxEvolucion = document.getElementById('chart-evolucion').getContext('2d');
  const labelsEvolucion = m.serie_temporal.map(s => s.fecha);
  const dataProducido = m.serie_temporal.map(s => s.total_producido);
  const dataDefectos = m.serie_temporal.map(s => s.total_defectuoso);

  if (chartEvolucion) {
    chartEvolucion.destroy();
  }

  chartEvolucion = new Chart(ctxEvolucion, {
    type: 'bar',
    data: {
      labels: labelsEvolucion,
      datasets: [
        {
          label: 'Unidades Producidas',
          data: dataProducido,
          backgroundColor: '#334155',
          borderRadius: 2,
          yAxisID: 'y'
        },
        {
          label: 'Unidades Defectuosas',
          data: dataDefectos,
          backgroundColor: '#dc2626',
          borderRadius: 2,
          yAxisID: 'y'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false
      },
      plugins: {
        legend: {
          position: 'top',
          labels: {
            font: { size: 11, family: '-apple-system, sans-serif' },
            boxWidth: 12
          }
        },
        tooltip: {
          backgroundColor: '#0f172a',
          titleFont: { size: 12 },
          bodyFont: { size: 11 }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { size: 10 } }
        },
        y: {
          beginAtZero: true,
          grid: { color: '#e2e8f0' },
          ticks: { font: { size: 10 } }
        }
      }
    }
  });

  // 2. Gráfico de Turnos (Donut sobrio)
  const ctxTurnos = document.getElementById('chart-turnos').getContext('2d');
  const labelsTurnos = m.produccion_por_turno.map(t => t.turno);
  const dataTurnos = m.produccion_por_turno.map(t => t.total_producido);

  if (chartTurnos) {
    chartTurnos.destroy();
  }

  chartTurnos = new Chart(ctxTurnos, {
    type: 'doughnut',
    data: {
      labels: labelsTurnos,
      datasets: [{
        data: dataTurnos,
        backgroundColor: ['#1e293b', '#475569', '#94a3b8', '#cbd5e1'],
        borderWidth: 1,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            font: { size: 11 },
            boxWidth: 12
          }
        }
      },
      cutout: '65%'
    }
  });
}

function renderTable(detalle) {
  const tbody = document.getElementById('table-body');
  const hint = document.getElementById('table-hint');

  if (!detalle || detalle.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="text-center py-4 text-muted">No se encontraron registros con los filtros seleccionados.</td>
      </tr>
    `;
    hint.textContent = '0 registros encontrados';
    return;
  }

  hint.textContent = `Mostrando ${detalle.length} registros recientes`;

  tbody.innerHTML = detalle.map(r => {
    const rate = r.cantidad_producida > 0 
      ? ((r.cantidad_defectuosa / r.cantidad_producida) * 100).toFixed(2)
      : '0.00';
    
    let badgeClass = 'low';
    if (rate > 4.0) badgeClass = 'high';
    else if (rate > 2.0) badgeClass = 'medium';

    return `
      <tr>
        <td style="font-family: var(--font-mono); font-size: 0.78rem;">${r.fecha}</td>
        <td>${r.turno}</td>
        <td>${r.linea}</td>
        <td style="font-weight: 500;">${r.producto}</td>
        <td class="text-right" style="font-family: var(--font-mono);">${formatNumber(r.cantidad_producida)}</td>
        <td class="text-right" style="font-family: var(--font-mono); color: ${r.cantidad_defectuosa > 0 ? '#b91c1c' : 'inherit'};">${formatNumber(r.cantidad_defectuosa)}</td>
        <td class="text-right">
          <span class="badge-defect ${badgeClass}">${rate}%</span>
        </td>
        <td>${r.operario}</td>
      </tr>
    `;
  }).join('');
}

