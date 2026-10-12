let allData = [];
let charts = {};
let dateField = '';
const chartPalette = ['#2563EB', '#06B6D4', '#F59E0B', '#D946EF', '#10B981', '#F97316', '#7C3AED', '#EC4899'];
const maxHighlightColor = '#d62f2f';
const impactColor = '#F59E0B';
const infrastructureColor = '#2563EB';

function colorByValue(values) {
  const maxValue = Math.max(0, ...values);
  return values.map((value, index) =>
    maxValue > 0 && value === maxValue
      ? maxHighlightColor
      : chartPalette[index % chartPalette.length]
  );
}

const horizontalBarValues = {
  id: 'horizontalBarValues',
  afterDatasetsDraw(chart) {
    const dataset = chart.data.datasets[0];
    if (!dataset) return;

    const context = chart.ctx;
    const meta = chart.getDatasetMeta(0);
    context.save();
    context.font = '600 11px Inter, sans-serif';
    context.fillStyle = '#555';
    context.textAlign = 'left';
    context.textBaseline = 'middle';

    meta.data.forEach((bar, index) => {
      const value = Number(dataset.data[index] || 0);
      const position = bar.tooltipPosition();
      context.fillText(String(value), position.x + 6, position.y);
    });

    context.restore();
  }
};

const verticalBarValues = {
  id: 'verticalBarValues',
  afterDatasetsDraw(chart) {
    const context = chart.ctx;
    context.save();
    context.font = '600 11px Inter, sans-serif';
    context.fillStyle = '#555';
    context.textAlign = 'center';
    context.textBaseline = 'bottom';

    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if (meta.hidden || !chart.isDatasetVisible(datasetIndex)) return;

      meta.data.forEach((bar, index) => {
        const value = Number(dataset.data[index] || 0);
        const position = bar.tooltipPosition();
        context.fillText(String(value), position.x, position.y - 5);
      });
    });

    context.restore();
  }
};

const els = {
  fileInput: document.getElementById('fileInput'),
  clearDataBtn: document.getElementById('clearDataBtn'),
  ufFilter: document.getElementById('ufFilter'),
  siteTypeFilter: document.getElementById('siteTypeFilter'),
  impactFilter: document.getElementById('impactFilter'),
  tableUfFilter: document.getElementById('tableUfFilter'),
  dateFromFilter: document.getElementById('dateFromFilter'),
  dateToFilter: document.getElementById('dateToFilter'),
  dateFilterHint: document.getElementById('dateFilterHint'),
  totalImpactos: document.getElementById('totalImpactos'),
  totalInfra: document.getElementById('totalInfra'),
  totalCorrigidos: document.getElementById('totalCorrigidos'),
  deviceDateValue: document.getElementById('deviceDateValue'),
  imageInput: document.getElementById('imageInput'),
  clearImageBtn: document.getElementById('clearImageBtn'),
  uploadedImage: document.getElementById('uploadedImage'),
  noImagePlaceholder: document.getElementById('noImagePlaceholder'),
  infraTableHead: document.getElementById('infraTableHead'),
  infraTableBody: document.getElementById('infraTableBody')
};

function formatDeviceDate(date = new Date()) {
  return new Intl.DateTimeFormat('pt-BR', {
    weekday: 'long',
    day: '2-digit',
    month: 'long',
    year: 'numeric'
  }).format(date);
}

function syncDeviceDateCard() {
  if (!els.deviceDateValue) return;
  els.deviceDateValue.textContent = formatDeviceDate();
}

function buildDefaultFilters() {
  const sortedUfs = [...new Set(allData.map(item => item.uf))]
    .sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const ufOptions = ['all', ...sortedUfs];
  populateSelect(els.ufFilter, ufOptions, 'Todas as UFs');
  populateSelect(els.tableUfFilter, ufOptions, 'Todas as UFs');
  populateSelect(els.siteTypeFilter, ['all', ...new Set(allData.map(item => item.tipoSite).filter(Boolean))], 'Todos');
  populateSelect(els.impactFilter, ['all', ...new Set(allData.map(item => item.tipoImpacto).filter(Boolean))], 'Todos');
  configureDateFilters();
}

function configureDateFilters() {
  const columns = [...new Set(allData.flatMap(item => Object.keys(item.source || {})))]
    .filter(isDateHeader);
  const preferredColumns = [
    'dia',
    'data incidente', 'data do incidente', 'data abertura', 'data de abertura',
    'data impacto', 'data do impacto', 'data criacao', 'data de criacao',
    'data', 'date', 'created at'
  ];
  dateField = preferredColumns.find(column => columns.includes(column)) || columns[0] || '';

  const enabled = Boolean(dateField);
  els.dateFromFilter.disabled = !enabled;
  els.dateToFilter.disabled = !enabled;
  els.dateFromFilter.value = '';
  els.dateToFilter.value = '';
  els.dateFilterHint.textContent = enabled
    ? `Análise por período usando a coluna “${displayColumnName(dateField)}”.`
    : 'Esta planilha não contém uma coluna de data reconhecida; o filtro por data está indisponível.';
}

function isDateHeader(header) {
  return /^(dia|date|data|data incidente|data do incidente|data abertura|data de abertura|data impacto|data do impacto|abertura|fechamento|criacao|resolucao|created at|updated at)$/.test(header) ||
    /^data (do |de )?(incidente|impacto|abertura|fechamento|criacao|resolucao|acao|registro)$/.test(header);
}

function populateSelect(selectEl, values, allLabel) {
  selectEl.innerHTML = '';
  const baseOption = document.createElement('option');
  baseOption.value = values[0];
  baseOption.textContent = allLabel;
  selectEl.appendChild(baseOption);

  values.slice(1).forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });
}

function parseImportFile(file) {
  return new Promise((resolve, reject) => {
    const isCsv = file.name.toLowerCase().endsWith('.csv');
    const reader = new FileReader();

    reader.onload = event => {
      try {
        let rows = [];
        if (isCsv) {
          rows = csvToRows(event.target.result);
        } else {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          const sheetName = workbook.SheetNames[0];
          const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1,
            defval: '',
            blankrows: false
          });
          rows = matrixToObjects(matrix);
        }

        const normalized = rows
          .map(normalizeItem)
          .filter(item => item && item.uf && item.impactos !== undefined && item.infraestrutura !== undefined && item.corrigidos !== undefined);

        if (!normalized.length) {
          throw new Error('Não encontrei linhas válidas. Verifique se a planilha contém as colunas INC, IMPACTO, INC. INFRA., SITE, TIPO DE SITE, FALHA / INFORMAÇÃO e AÇÃO; a linha de cabeçalho pode ficar abaixo de um título.');
        }

        allData = normalized;
        buildDefaultFilters();
        renderDashboard();
        resolve();
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error('Não foi possível ler o arquivo.'));

    if (isCsv) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

function csvToRows(text) {
  const sample = text.split(/\r?\n/).slice(0, 30).join('\n');
  const delimiter = (sample.match(/;/g) || []).length > (sample.match(/,/g) || []).length ? ';' : ',';
  const rows = [];
  let row = [];
  let value = '';
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === '"' && text[index + 1] === '"' && quoted) {
      value += '"';
      index += 1;
    } else if (character === '"') {
      quoted = !quoted;
    } else if (character === delimiter && !quoted) {
      row.push(value.trim());
      value = '';
    } else if ((character === '\n' || character === '\r') && !quoted) {
      if (character === '\r' && text[index + 1] === '\n') index += 1;
      row.push(value.trim());
      if (row.some(cell => cell !== '')) rows.push(row);
      row = [];
      value = '';
    } else {
      value += character;
    }
  }

  row.push(value.trim());
  if (row.some(cell => cell !== '')) rows.push(row);
  return matrixToObjects(rows);
}

function matrixToObjects(matrix) {
  if (!matrix.length) return [];

  const recognizedHeaders = new Set([
    'inc', 'incidente', 'impacto', 'designado', 'grupo designado', 'relacionado',
    'inc infra', 'inc infraestrutura', 'site', 'tipo de site',
    'falha', 'falha informacao', 'acao', 'uf', 'estado', 'infraestrutura',
    'corrigidos', 'total impactos', 'total', 'dia', 'data', 'date', 'data incidente',
    'data do incidente', 'data abertura', 'data de abertura', 'data impacto',
    'data do impacto'
  ]);
  let headerIndex = -1;
  let highestScore = 0;

  matrix.slice(0, 30).forEach((row, index) => {
    const normalized = row.map(normalizeHeader);
    const score = normalized.filter(header => recognizedHeaders.has(header)).length;
    if (score > highestScore) {
      highestScore = score;
      headerIndex = index;
    }
  });

  if (headerIndex < 0 || highestScore < 2) return [];

  const headers = matrix[headerIndex].map(normalizeHeader);
  return matrix.slice(headerIndex + 1)
    .filter(values => values.some(value => String(value).trim() !== ''))
    .map(values => Object.fromEntries(
      headers.map((header, index) => [header, values[index] ?? ''])
    ));
}

function normalizeItem(item) {
  const fields = Object.fromEntries(
    Object.entries(item).map(([key, value]) => [normalizeHeader(key), value])
  );
  const valueFor = (...keys) => {
    const key = keys.find(candidate => fields[candidate] !== undefined && fields[candidate] !== '');
    return key ? fields[key] : '';
  };
  const itemDateField = Object.keys(fields).find(isDateHeader) || '';
  const rawDate = itemDateField ? fields[itemDateField] : '';
  const date = itemDateField ? parseDate(rawDate) : null;
  const dayOfMonth = itemDateField === 'dia' && !date ? parseDayOfMonth(rawDate) : null;

  const site = String(valueFor('site') || '').trim();
  const infraId = String(valueFor(
    'inc infra', 'inc infraestrutura', 'incidente infra', 'incidente infraestrutura',
    'infra', 'infraestrutura', 'incidentes infraestrutura'
  ) || valueFor('incidente') || valueFor('inc')).trim();
  const isDetailedIncident = fields.inc !== undefined || fields.impacto !== undefined ||
    fields.incidente !== undefined || fields['tipo de site'] !== undefined ||
    fields.falha !== undefined || fields['falha informacao'] !== undefined;

  if (isDetailedIncident) {
    const uf = normalizeUf(valueFor('uf', 'estado')) || extractUf(site);
    if (!uf) return null;
    const impactDescription = String(valueFor('impacto') || '').trim();
    const impactType = normalizeImpactType(
      String(impactDescription || valueFor('falha informacao', 'falha', 'tipo impacto') || 'Não informado').trim()
    );
    const action = String(valueFor('acao', 'status', 'situacao') || '').trim();

    return {
      uf,
      impactos: impactDescription ? 1 : 0,
      infraestrutura: infraId ? 1 : 0,
      corrigidos: isResolved(action) ? 1 : 0,
      incidentId: String(valueFor('incidente', 'inc') || '').trim() || null,
      infraId: infraId || null,
      tipoSite: String(valueFor('tipo de site', 'tipo site', 'site type') || 'N/C').trim(),
      tipoImpacto: impactType,
      detailed: true,
      source: fields,
      date,
      dayOfMonth
    };
  }

  const uf = normalizeUf(valueFor('uf', 'estado'));
  if (!uf) return null;

  return {
    uf,
    impactos: parseNumber(valueFor('impactos', 'impactos total', 'total impactos', 'total')),
    infraestrutura: parseNumber(valueFor('infraestrutura', 'incidentes infraestrutura', 'total infraestrutura')),
    corrigidos: parseNumber(valueFor('corrigidos', 'incidentes corrigidos', 'total corrigidos')),
    infraId: null,
    tipoSite: String(valueFor('tipo site', 'site type', 'tipo de site', 'site') || 'N/C').trim(),
    tipoImpacto: String(valueFor('tipo impacto', 'impact type', 'impacto', 'tipo') || 'Não informado').trim(),
    detailed: false,
    source: fields,
    date,
    dayOfMonth
  };
}

function normalizeHeader(value) {
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeImpactType(value) {
  if (normalizeHeader(value) === 'falha massiva soo') {
    return 'FALHA MASSIVA-SOOS';
  }
  return value;
}

function parseNumber(value) {
  const normalized = String(value || '0').replace(/[^0-9,.-]/g, '').replace(',', '.');
  return Number(normalized) || 0;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDate(value);
  }

  if (typeof value === 'number' && value >= 20000 && value <= 80000) {
    const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    return formatDate(date);
  }

  const text = String(value || '').trim();
  if (!text) return null;

  if (/^\d{5}(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial >= 20000 && serial <= 80000) {
      return formatDate(new Date(Date.UTC(1899, 11, 30) + serial * 86400000));
    }
  }

  const brazilianDate = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (brazilianDate) {
    const [, day, month, year] = brazilianDate;
    const fullYear = year.length === 2 ? `20${year}` : year;
    const parsed = new Date(Date.UTC(Number(fullYear), Number(month) - 1, Number(day)));
    if (parsed.getUTCFullYear() !== Number(fullYear) || parsed.getUTCMonth() !== Number(month) - 1 || parsed.getUTCDate() !== Number(day)) {
      return null;
    }
    return formatDate(parsed);
  }

  const isoDate = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoDate) {
    const parsed = new Date(Date.UTC(Number(isoDate[1]), Number(isoDate[2]) - 1, Number(isoDate[3])));
    if (parsed.getUTCFullYear() !== Number(isoDate[1]) || parsed.getUTCMonth() !== Number(isoDate[2]) - 1 || parsed.getUTCDate() !== Number(isoDate[3])) {
      return null;
    }
    return formatDate(parsed);
  }
  return null;
}

function parseDayOfMonth(value) {
  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= 31) {
    return value;
  }

  const match = String(value || '').trim().match(/^(\d{1,2})$/);
  if (!match) return null;
  const day = Number(match[1]);
  return day >= 1 && day <= 31 ? day : null;
}

function dayOccursInRange(day, from, to) {
  if (!from || !to || from > to) return false;

  const start = new Date(`${from}T00:00:00Z`);
  const end = new Date(`${to}T00:00:00Z`);
  const month = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), 1));

  while (month <= end) {
    const candidate = new Date(Date.UTC(month.getUTCFullYear(), month.getUTCMonth(), day));
    if (candidate.getUTCMonth() === month.getUTCMonth() && candidate >= start && candidate <= end) {
      return true;
    }
    month.setUTCMonth(month.getUTCMonth() + 1);
  }

  return false;
}

function formatDate(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}

function normalizeUf(value) {
  const uf = String(value || '').trim().toUpperCase();
  return /^[A-Z]{2}$/.test(uf) ? uf : '';
}

function extractUf(site) {
  const match = String(site || '').toUpperCase().match(/^\s*(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)(?=$|[^A-Z])/);
  return match ? match[1] : '';
}

function isResolved(action) {
  const normalized = normalizeHeader(action);
  return /\b(corrigido|corrigida|resolvido|resolvida|concluido|concluida|encerrado|encerrada|restabelecido|restabelecida|regularizado|regularizada)\b/i
    .test(normalized) && !/\bnao\s+(corrigido|corrigida|resolvido|resolvida|concluido|concluida|encerrado|encerrada)\b/i.test(normalized);
}

function getFilteredData() {
  const uf = els.ufFilter.value;
  const siteType = els.siteTypeFilter.value;
  const impact = els.impactFilter.value;
  const dateFrom = els.dateFromFilter.value;
  const dateTo = els.dateToFilter.value;

  return allData.filter(item => {
    if (uf !== 'all' && item.uf !== uf) return false;
    if (siteType !== 'all' && item.tipoSite !== siteType) return false;
    if (impact !== 'all' && item.tipoImpacto !== impact) return false;
    if (dateFrom || dateTo) {
      if (item.date) {
        if (dateFrom && item.date < dateFrom) return false;
        if (dateTo && item.date > dateTo) return false;
      } else if (item.dayOfMonth) {
        if (!dayOccursInRange(item.dayOfMonth, dateFrom, dateTo)) return false;
      } else {
        return false;
      }
    }
    return true;
  });
}

function renderSummaryBoxes(filtered) {
  const totalImpactos = filtered.reduce((sum, item) => sum + item.impactos, 0);
  const totalInfra = countInfrastructure(filtered);
  const totalCorrigidos = countCorrected(filtered);

  els.totalImpactos.textContent = totalImpactos;
  els.totalInfra.textContent = totalInfra;
  els.totalCorrigidos.textContent = totalCorrigidos;

  const ufCounts = groupBy(filtered, 'uf');
  const siteCounts = groupBy(filtered, 'tipoSite');
  const impactCounts = groupByMetric(filtered, 'tipoImpacto', items =>
    countUniqueIncidents(items, 'incidentId', item => item.impactos)
  );
  const infraByUf = groupByMetric(filtered, 'uf', countInfrastructure);
  const impactByUf = groupByMetric(filtered, 'uf', items =>
    items.reduce((sum, item) => sum + item.impactos, 0)
  );

  document.getElementById('summaryImpactByUf').textContent = buildSummary('UFs com maior concentração', Object.entries(ufCounts).sort((a, b) => b[1] - a[1]).slice(0, 3));
  document.getElementById('summaryImpactBySite').textContent = buildSummary('Tipos de site mais impactados', Object.entries(siteCounts).sort((a, b) => b[1] - a[1]).slice(0, 3));
  const impactInfraEntries = Object.keys(impactByUf)
    .sort((a, b) => impactByUf[b] - impactByUf[a] || a.localeCompare(b, 'pt-BR'))
    .slice(0, 3)
    .map(uf => {
      const impacts = impactByUf[uf];
      const incidents = infraByUf[uf] || 0;
      const average = incidents ? (impacts / incidents).toLocaleString('pt-BR', { maximumFractionDigits: 2 }) : '0';
      return [uf, `${impacts} impactos / ${incidents} incidentes únicos (${average} impactos por incidente)`];
    });
  document.getElementById('summaryImpactVsInfra').textContent = buildSummary('Impactos por incidente único de infraestrutura', impactInfraEntries);
  document.getElementById('summaryImpactTypes').textContent = buildSummary('Incidentes únicos por tipo de impacto', Object.entries(impactCounts).sort((a, b) => b[1] - a[1]).slice(0, 3));
  document.getElementById('summaryUniqueInfraByUf').textContent = buildSummary(
    'UFs com mais incidentes únicos de infraestrutura',
    Object.entries(infraByUf)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'))
      .slice(0, 3)
  );
  const dateRange = els.dateFromFilter.value || els.dateToFilter.value
    ? ` Período: ${els.dateFromFilter.value || 'sem limite inicial'} a ${els.dateToFilter.value || 'sem limite final'}.`
    : '';
  updateDateFilterHint(filtered);
  document.getElementById('summaryImage').textContent = 'Importe uma imagem da tabela de incidentes abertos acima de 120 horas para visualizá-la aqui.';
}

function updateDateFilterHint(filtered) {
  if (!dateField) {
    els.dateFilterHint.textContent = 'Esta planilha não contém uma coluna de data reconhecida; o filtro por data está indisponível.';
    return;
  }

  const dateFrom = els.dateFromFilter.value;
  const dateTo = els.dateToFilter.value;
  if (!dateFrom && !dateTo) {
    els.dateFilterHint.textContent = `Análise por período usando a coluna “${displayColumnName(dateField)}”.`;
    return;
  }

  if (dateFrom && dateTo && dateFrom > dateTo) {
    els.dateFilterHint.textContent = 'A data inicial deve ser anterior ou igual à data final.';
    return;
  }

  const hasDayOnlyDates = allData.some(item => item.dayOfMonth);
  if (hasDayOnlyDates && (!dateFrom || !dateTo)) {
    els.dateFilterHint.textContent = 'A coluna DIA contém apenas o número do dia. Selecione a data inicial e a data final para filtrar.';
    return;
  }

  const readableDates = allData.filter(item => item.date || item.dayOfMonth).length;
  els.dateFilterHint.textContent = filtered.length
    ? `${filtered.length} registro(s) encontrado(s) no período selecionado.`
    : `Nenhum registro encontrado nesse período. ${readableDates} de ${allData.length} linhas têm DIA reconhecido; confira o intervalo e os demais filtros.`;
}

function buildSummary(label, entries) {
  if (!entries.length) return `${label}: sem dados.`;
  const text = entries.map(([key, value]) => `${key}: ${value}`).join(' | ');
  return `${label}: ${text}.`;
}

function groupBy(data, key) {
  return data.reduce((acc, item) => {
    const value = item[key];
    acc[value] = (acc[value] || 0) + item.impactos;
    return acc;
  }, {});
}

function countInfrastructure(data) {
  const detailedIds = new Set(data.filter(item => item.detailed && item.infraId).map(item => item.infraId));
  const aggregateCount = data.filter(item => !item.detailed).reduce((sum, item) => sum + item.infraestrutura, 0);
  return detailedIds.size + aggregateCount;
}

function countCorrected(data) {
  const correctedIds = new Set(
    data.filter(item => item.detailed && item.corrigidos && item.infraId).map(item => item.infraId)
  );
  const aggregateCount = data.filter(item => !item.detailed).reduce((sum, item) => sum + item.corrigidos, 0);
  return correctedIds.size + aggregateCount;
}

function groupByMetric(data, key, metric) {
  const groups = new Map();
  data.forEach(item => {
    const group = groups.get(item[key]) || [];
    group.push(item);
    groups.set(item[key], group);
  });
  return Object.fromEntries([...groups].map(([name, items]) => [name, metric(items)]));
}

function countUniqueIncidents(data, identifierKey, fallbackMetric) {
  const identifiers = new Set();
  let fallbackCount = 0;

  data.forEach(item => {
    if (item.detailed && item[identifierKey]) {
      identifiers.add(item[identifierKey]);
    } else {
      fallbackCount += fallbackMetric(item);
    }
  });

  return identifiers.size + fallbackCount;
}

function renderCharts(filtered) {
  const ufLabels = [...new Set(filtered.map(item => item.uf))];
  const ufImpactCounts = ufLabels.map(uf => filtered.filter(item => item.uf === uf).reduce((sum, item) => sum + item.impactos, 0));
  const ufInfraCounts = ufLabels.map(uf => countInfrastructure(filtered.filter(item => item.uf === uf)));
  const sortedUfComparison = ufLabels
    .map((uf, index) => ({
      uf,
      impacts: ufImpactCounts[index],
      infrastructure: ufInfraCounts[index]
    }))
    .sort((a, b) =>
      Math.max(b.impacts, b.infrastructure) - Math.max(a.impacts, a.infrastructure) ||
      b.impacts - a.impacts ||
      a.uf.localeCompare(b.uf, 'pt-BR')
    );
  const sortedUfImpacts = ufLabels
    .map((uf, index) => ({ uf, count: ufImpactCounts[index] }))
    .sort((a, b) => b.count - a.count || a.uf.localeCompare(b.uf, 'pt-BR'));

  const sortedSiteImpacts = [...new Set(filtered.map(item => item.tipoSite))]
    .map(site => ({
      site,
      count: filtered.filter(item => item.tipoSite === site).reduce((sum, item) => sum + item.impactos, 0)
    }))
    .sort((a, b) => b.count - a.count || a.site.localeCompare(b.site, 'pt-BR'));

  const sortedImpactCounts = Object.entries(groupByMetric(filtered, 'tipoImpacto', items =>
    countUniqueIncidents(items, 'incidentId', item => item.impactos)
  ))
    .map(([type, count]) => ({ type, count }))
    .sort((a, b) => b.count - a.count || a.type.localeCompare(b.type, 'pt-BR'));

  const infraLabels = [...new Set(filtered.map(item => item.uf))];
  const infraCounts = infraLabels.map(uf => countInfrastructure(filtered.filter(item => item.uf === uf)));
  const sortedInfraCounts = infraLabels
    .map((uf, index) => ({ uf, count: infraCounts[index] }))
    .sort((a, b) => b.count - a.count || a.uf.localeCompare(b.uf, 'pt-BR'));

  buildPieChart(
    'impactByUfChart',
    sortedUfImpacts.map(item => item.uf),
    sortedUfImpacts.map(item => item.count)
  );
  buildHorizontalBarChart(
    'impactBySiteChart',
    sortedSiteImpacts.map(item => item.site),
    sortedSiteImpacts.map(item => item.count)
  );
  buildGroupedBarChart(
    'impactVsInfraChart',
    sortedUfComparison.map(item => item.uf),
    sortedUfComparison.map(item => item.impacts),
    sortedUfComparison.map(item => item.infrastructure),
    'Impactos e incidentes únicos de infraestrutura por UF'
  );
  buildColumnChart(
    'impactTypesChart',
    sortedImpactCounts.map(item => item.type),
    sortedImpactCounts.map(item => item.count)
  );
  buildColumnChart(
    'uniqueInfraByUfChart',
    sortedInfraCounts.map(item => item.uf),
    sortedInfraCounts.map(item => item.count)
  );
}

function buildColumnChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }

  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colorByValue(data),
        borderColor: '#666',
        borderWidth: 1,
        borderRadius: 0,
        borderSkipped: false,
        barPercentage: 0.78,
        categoryPercentage: 0.88
      }]
    },
    plugins: [verticalBarValues],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: { top: 16 }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: context => ` ${context.raw}`
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            color: '#222',
            font: { size: 11, weight: 'bold' }
          }
        },
        y: {
          beginAtZero: true,
          ticks: { precision: 0 },
          grid: {
            color: '#d7d7d7',
            borderDash: [2, 3]
          }
        }
      }
    }
  });
}

function buildHorizontalBarChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }

  const maxValue = Math.max(0, ...data);
  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colorByValue(data),
        borderColor: '#fff',
        borderWidth: 1,
        borderRadius: 0,
        borderSkipped: false,
        barPercentage: 0.78,
        categoryPercentage: 0.88
      }]
    },
    plugins: [horizontalBarValues],
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: { right: 30 }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: context => ` ${context.raw}`
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          max: maxValue ? maxValue * 1.12 : undefined,
          ticks: { precision: 0 },
          grid: {
            color: '#d7d7d7',
            borderDash: [2, 3]
          }
        },
        y: {
          grid: { display: false },
          ticks: {
            color: '#222',
            font: { size: 10 }
          }
        }
      }
    }
  });
}

function buildPieChart(canvasId, labels, data) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }

  charts[canvasId] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colorByValue(data),
        borderColor: '#fff',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '58%',
      radius: '60%',
      layout: {
        padding: { top: 12, right: 12, bottom: 12, left: 12 }
      },
      plugins: {
        legend: {
          display: true,
          position: 'right',
          labels: {
            usePointStyle: true,
            pointStyle: 'circle',
            boxWidth: 8,
            padding: 8,
            font: { size: 11 },
            generateLabels(chart) {
              const dataset = chart.data.datasets[0];
              return chart.data.labels.map((label, index) => ({
                text: `${label}   ${dataset.data[index]}`,
                fillStyle: dataset.backgroundColor[index],
                strokeStyle: dataset.borderColor,
                lineWidth: dataset.borderWidth,
                hidden: !chart.getDataVisibility(index),
                index
              }));
            }
          }
        },
        tooltip: {
          callbacks: {
            label(context) {
              const value = Number(context.raw || 0);
              const total = context.dataset.data.reduce((sum, item) => sum + Number(item || 0), 0);
              const percent = total ? (value / total * 100).toLocaleString('pt-BR', { maximumFractionDigits: 2 }) : '0';
              return `${context.label}: ${value} (${percent}%)`;
            }
          }
        }
      }
    }
  });
}

function buildBarChart(canvasId, labels, data, color, title) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }

  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: title,
        data,
        backgroundColor: color,
        borderRadius: 6,
        borderSkipped: false
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: false }
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, ticks: { precision: 0 } }
      }
    }
  });
}

function buildGroupedBarChart(canvasId, labels, primaryData, secondaryData, title) {
  const ctx = document.getElementById(canvasId);
  if (!ctx) return;

  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }

  const primaryColors = primaryData.map((value, index) =>
    value > 0 && value >= secondaryData[index] ? maxHighlightColor : impactColor
  );
  const secondaryColors = secondaryData.map((value, index) =>
    value > 0 && value >= primaryData[index] ? maxHighlightColor : infrastructureColor
  );
  charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: 'Impactos',
          data: primaryData,
          backgroundColor: primaryColors,
          borderRadius: 6
        },
        {
          label: 'Incidentes únicos de infraestrutura',
          data: secondaryData,
          backgroundColor: secondaryColors,
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' },
        title: { display: true, text: title, font: { size: 14 } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, ticks: { precision: 0 } }
      }
    },
    plugins: [verticalBarValues]
  });
}

function renderTable(filtered) {
  els.infraTableHead.replaceChildren();
  els.infraTableBody.replaceChildren();

  const tableUf = els.tableUfFilter.value;
  const tableData = tableUf === 'all'
    ? filtered
    : filtered.filter(item => item.uf === tableUf);

  if (!tableData.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 1;
    cell.textContent = filtered.length ? 'Nenhum registro para a UF selecionada.' : 'Nenhum dado importado.';
    row.appendChild(cell);
    els.infraTableBody.appendChild(row);
    updateTableSummary(0, filtered.length, tableUf, dateRangeText());
    return;
  }

  const columns = [...new Set(tableData.flatMap(item => Object.keys(item.source || {}).filter(Boolean)))];
  if (!columns.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.textContent = 'O arquivo não contém colunas para exibir.';
    row.appendChild(cell);
    els.infraTableBody.appendChild(row);
    return;
  }

  const headerRow = document.createElement('tr');
  columns.forEach(column => {
    const heading = document.createElement('th');
    heading.textContent = displayColumnName(column);
    headerRow.appendChild(heading);
  });
  els.infraTableHead.appendChild(headerRow);

  tableData.forEach(item => {
    const row = document.createElement('tr');
    columns.forEach(column => {
      const cell = document.createElement('td');
      const value = item.source?.[column];
      cell.textContent = value == null ? '' : String(value);
      row.appendChild(cell);
    });
    els.infraTableBody.appendChild(row);
  });
  updateTableSummary(tableData.length, filtered.length, tableUf, dateRangeText());
}

function dateRangeText() {
  if (!els.dateFromFilter.value && !els.dateToFilter.value) return '';
  return ` Período: ${els.dateFromFilter.value || 'sem limite inicial'} a ${els.dateToFilter.value || 'sem limite final'}.`;
}

function updateTableSummary(visibleRows, filteredRows, tableUf, dateRange) {
  const ufText = tableUf === 'all' ? 'todas as UFs' : `UF ${tableUf}`;
  document.getElementById('summaryInfraTable').textContent =
    `Linhas exibidas: ${visibleRows} de ${filteredRows} (${ufText}). A área da tabela mostra até 10 linhas; use a rolagem para ver as demais.${dateRange}`;
}

function displayColumnName(column) {
  const labels = {
    inc: 'INC',
    impacto: 'IMPACTO',
    designado: 'DESIGNADO',
    'grupo designado': 'GRUPO DESIGNADO',
    relacionado: 'RELACIONADO',
    'inc infra': 'INC. INFRA.',
    'inc infraestrutura': 'INC. INFRA.',
    site: 'SITE',
    'tipo de site': 'TIPO DE SITE',
    'falha informacao': 'FALHA / INFORMAÇÃO',
    acao: 'AÇÃO'
  };
  return labels[column] || column.toLocaleUpperCase('pt-BR');
}

function syncFilters() {
  const ufValue = els.ufFilter.value;
  const siteValue = els.siteTypeFilter.value;
  const impactValue = els.impactFilter.value;

  const filtered = getFilteredData();
  renderSummaryBoxes(filtered);
  renderCharts(filtered);
  renderTable(filtered);

  els.ufFilter.value = ufValue;
  els.siteTypeFilter.value = siteValue;
  els.impactFilter.value = impactValue;
}

function initEventHandlers() {
  els.fileInput.addEventListener('change', async event => {
    const [file] = event.target.files;
    if (!file) return;

    try {
      await parseImportFile(file);
      syncDeviceDateCard();
      alert('Arquivo importado com sucesso!');
    } catch (error) {
      alert(error.message || 'Erro ao importar o arquivo.');
    } finally {
      event.target.value = '';
    }
  });

  els.clearDataBtn.addEventListener('click', () => {
    allData = [];
    buildDefaultFilters();
    els.ufFilter.value = 'all';
    els.siteTypeFilter.value = 'all';
    els.impactFilter.value = 'all';
    renderDashboard();
  });

  els.ufFilter.addEventListener('change', syncFilters);
  els.siteTypeFilter.addEventListener('change', syncFilters);
  els.impactFilter.addEventListener('change', syncFilters);
  els.dateFromFilter.addEventListener('change', syncFilters);
  els.dateToFilter.addEventListener('change', syncFilters);
  els.tableUfFilter.addEventListener('change', () => renderTable(getFilteredData()));

  els.imageInput.addEventListener('change', event => {
    const [file] = event.target.files;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = e => {
      const img = els.uploadedImage;
      img.src = e.target.result;
      img.hidden = false;
      els.noImagePlaceholder.hidden = true;
    };
    reader.readAsDataURL(file);
    event.target.value = '';
  });

  els.clearImageBtn.addEventListener('click', () => {
    els.uploadedImage.src = '';
    els.uploadedImage.hidden = true;
    els.noImagePlaceholder.hidden = false;
  });
}

function renderDashboard() {
  const filtered = getFilteredData();
  renderSummaryBoxes(filtered);
  renderCharts(filtered);
  renderTable(filtered);
}

syncDeviceDateCard();
setInterval(syncDeviceDateCard, 60000);

buildDefaultFilters();
initEventHandlers();
renderDashboard();
