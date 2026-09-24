/* LedgerLens is deliberately client-side: uploaded files are parsed in this browser only. */
const state = { files: [], records: [], excerpts: [], chart: null, activeYear: null };
const $ = (selector) => document.querySelector(selector);
const fileInput = $('#fileInput');
const dropzone = $('#dropzone');

const moneyFormat = (value, unit = 'reported') => {
  if (value === null || value === undefined || Number.isNaN(value)) return '—';
  if (unit === 'reported') return compactNumber(value);
  const abs = Math.abs(value);
  const valueText = abs >= 100000 ? (value / 100000).toFixed(1) + 'L' : abs >= 1000 ? (value / 1000).toFixed(1) + 'K' : value.toLocaleString('en-IN', { maximumFractionDigits: 1 });
  return `${value < 0 ? '−' : ''}${unit} ${valueText}`;
};

const compactNumber = (value) => {
  if (value === null || value === undefined) return '—';
  return value >= 1000 ? `${(value / 1000).toFixed(value >= 10000 ? 0 : 1)}K` : String(Math.round(value));
};

function addFiles(files) {
  [...files].forEach((file) => {
    if (!state.files.some((entry) => entry.name === file.name && entry.size === file.size)) state.files.push(file);
  });
  renderFiles();
}

function renderFiles() {
  const list = $('#fileList');
  list.innerHTML = '';
  if (!state.files.length) {
    list.innerHTML = '<div class="empty-files">No sources added yet. Try the sample workspace to explore the analyzer.</div>';
  } else {
    state.files.forEach((file, index) => {
      const node = $('#fileTemplate').content.firstElementChild.cloneNode(true);
      const extension = file.name.split('.').pop().toUpperCase();
      node.querySelector('.file-type').textContent = extension === 'XLSX' ? 'XLS' : extension;
      node.querySelector('.file-info strong').textContent = file.name;
      node.querySelector('.file-info span').textContent = `${(file.size / 1024 / 1024).toFixed(file.size > 1024 * 1024 ? 1 : 2)} MB · ready`;
      node.querySelector('.remove-file').onclick = () => { state.files.splice(index, 1); renderFiles(); };
      list.appendChild(node);
    });
  }
  $('#analyzeButton').disabled = state.files.length === 0;
}

fileInput.addEventListener('change', (event) => addFiles(event.target.files));
['dragenter', 'dragover'].forEach((eventName) => dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.add('dragging'); }));
['dragleave', 'drop'].forEach((eventName) => dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.remove('dragging'); }));
dropzone.addEventListener('drop', (event) => addFiles(event.dataTransfer.files));

function normaliseNumber(input) {
  if (typeof input === 'number') return input;
  const text = String(input || '').replace(/[₹,$€£]/g, '').replace(/,/g, '').trim();
  if (!text || !/[0-9]/.test(text)) return null;
  const isNegative = /^\(|-$/.test(text) || text.includes('(');
  const value = Number.parseFloat(text.replace(/[()]/g, ''));
  return Number.isFinite(value) ? (isNegative ? -value : value) : null;
}

function findMetric(text, keys) {
  const lines = text.split(/\r?\n| {3,}/).map((line) => line.trim()).filter(Boolean);
  for (const line of lines) {
    const normalized = line.toLowerCase();
    if (keys.some((key) => normalized.includes(key))) {
      const values = line.match(/(?:\(?-?[₹$€£]?\s*[\d,]+(?:\.\d+)?\)?)/g) || [];
      const numeric = values.map(normaliseNumber).filter((value) => value !== null);
      if (numeric.length) return numeric[numeric.length - 1];
    }
  }
  return null;
}

function findYear(text, fallbackYear) {
  const matches = [...text.matchAll(/(?:FY\s*|fiscal year\s*)?(20\d{2})/gi)].map((match) => Number(match[1]));
  return matches.length ? Math.max(...matches) : fallbackYear;
}

function recordFromText(text, source, fallbackYear) {
  return {
    year: findYear(text, fallbackYear), source,
    revenue: findMetric(text, ['revenue from operations', 'total revenue', 'net sales', 'revenue']),
    ebitda: findMetric(text, ['ebitda', 'earnings before interest']),
    netIncome: findMetric(text, ['profit after tax', 'net profit', 'net income', 'profit for the year']),
    operatingCash: findMetric(text, ['cash flow from operating', 'net cash generated from operating']),
    text, unit: 'reported',
  };
}

async function readPdf(file) {
  if (!window.pdfjsLib) throw new Error('PDF reader could not be loaded. Check your connection and try again.');
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];
  for (let pageNumber = 1; pageNumber <= Math.min(pdf.numPages, 60); pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => item.str).join(' '));
  }
  return pages.join('\n');
}

async function readWorkbook(file) {
  if (!window.XLSX) throw new Error('Spreadsheet reader could not be loaded. Check your connection and try again.');
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  return workbook.SheetNames.map((name) => `SHEET: ${name}\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}`).join('\n');
}

async function readFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'pdf') return readPdf(file);
  if (['xlsx', 'xls'].includes(ext)) return readWorkbook(file);
  return file.text();
}

function extractExcerpts(text, source) {
  const themes = [
    { title: 'Growth outlook', type: 'opportunity', terms: ['growth', 'expansion', 'opportunity', 'demand', 'momentum'] },
    { title: 'Risk & uncertainty', type: 'risk', terms: ['risk', 'uncertainty', 'headwind', 'challenging', 'volatility'] },
    { title: 'Operating priorities', type: 'standard', terms: ['margin', 'efficiency', 'strategy', 'investment', 'operating'] },
  ];
  const sentences = text.replace(/\s+/g, ' ').match(/[^.!?]{45,260}[.!?]/g) || [];
  return themes.map((theme) => {
    const sentence = sentences.find((item) => theme.terms.some((term) => item.toLowerCase().includes(term)));
    return sentence ? { ...theme, quote: sentence.trim(), source } : null;
  }).filter(Boolean);
}

async function analyzeFiles() {
  const button = $('#analyzeButton');
  button.disabled = true; button.innerHTML = '<span>◌</span> Reading documents…';
  try {
    const parsed = await Promise.all(state.files.map(async (file, index) => {
      const text = await readFile(file);
      return { text, source: file.name, fallback: new Date().getFullYear() - index };
    }));
    state.records = parsed.map((result) => recordFromText(result.text, result.source, result.fallback)).sort((a, b) => a.year - b.year);
    state.excerpts = parsed.flatMap((result) => extractExcerpts(result.text, result.source)).slice(0, 3);
    state.activeYear = state.records.at(-1)?.year;
    renderDashboard();
  } catch (error) {
    alert(error.message || 'We could not read one of the documents.');
  } finally { button.disabled = false; button.innerHTML = '<span>✦</span> Analyze documents'; }
}

function percentageChange(current, previous) {
  if (current === null || previous === null || previous === 0) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function renderDashboard() {
  const records = state.records;
  const selected = records.find((record) => record.year === state.activeYear) || records.at(-1);
  const prev = records[records.indexOf(selected) - 1];
  const usable = records.filter((record) => Object.values(record).some((value) => typeof value === 'number'));
  $('#companyName').textContent = selected?.source?.replace(/\.[^.]+$/, '') || 'Workspace';
  $('#yearRange').textContent = records.length > 1 ? `${records[0].year}—${records.at(-1).year}` : selected?.year || '—';
  $('#yearTabs').innerHTML = records.map((record) => `<button class="${record.year === state.activeYear ? 'active' : ''}" data-year="${record.year}">FY ${record.year}</button>`).join('');
  $('#yearTabs').querySelectorAll('button').forEach((button) => button.onclick = () => { state.activeYear = Number(button.dataset.year); renderDashboard(); });
  const definitions = [
    ['Revenue', 'revenue'], ['EBITDA', 'ebitda'], ['Net income', 'netIncome'], ['Operating cash flow', 'operatingCash'],
  ];
  $('#metricsGrid').innerHTML = definitions.map(([label, key]) => {
    const change = percentageChange(selected?.[key], prev?.[key]);
    const direction = change === null ? '' : change >= 0 ? 'up' : 'down';
    const changeText = change === null ? 'No comparable prior value' : `<b>${change >= 0 ? '+' : ''}${change.toFixed(1)}%</b> vs FY ${prev.year}`;
    return `<article class="metric-card"><span class="metric-label">${label.toUpperCase()}</span><div class="metric-value">${moneyFormat(selected?.[key], selected?.unit)}</div><div class="metric-change ${direction}">${changeText}</div></article>`;
  }).join('');
  renderDrivers(selected, prev); renderChart(usable); renderCommentary();
}

function renderDrivers(current, previous) {
  const target = $('#driversList');
  const drivers = [['Revenue', 'revenue'], ['EBITDA', 'ebitda'], ['Net income', 'netIncome']].map(([name, key]) => ({ name, key, change: percentageChange(current?.[key], previous?.[key]) })).filter((driver) => driver.change !== null).sort((a, b) => Math.abs(b.change) - Math.abs(a.change));
  target.innerHTML = drivers.length ? drivers.map((driver, index) => `<div class="driver"><span class="driver-badge">0${index + 1}</span><div class="driver-text"><strong>${driver.name} ${driver.change >= 0 ? 'increased' : 'decreased'} ${Math.abs(driver.change).toFixed(1)}%</strong><p>Change observed between FY ${previous.year} and FY ${current.year}, based on extracted statement values.</p></div></div>`).join('') : '<p class="placeholder">Add reports from at least two periods to identify financial drivers.</p>';
}

function renderChart(records) {
  if (state.chart) state.chart.destroy();
  const ctx = $('#trendChart');
  if (!window.Chart || !records.length) return;
  state.chart = new Chart(ctx, { type: 'line', data: { labels: records.map((record) => `FY ${record.year}`), datasets: [
    { label: 'Revenue', data: records.map((record) => record.revenue), borderColor: '#194d3b', backgroundColor: '#194d3b', borderWidth: 2, tension: .35, spanGaps: true, pointRadius: 4, pointHoverRadius: 5 },
    { label: 'Net income', data: records.map((record) => record.netIncome), borderColor: '#ed805e', backgroundColor: '#ed805e', borderWidth: 2, tension: .35, spanGaps: true, pointRadius: 4, pointHoverRadius: 5 },
  ] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { backgroundColor: '#173b2d', padding: 10, titleFont: { family: 'Manrope' }, bodyFont: { family: 'DM Mono' } } }, scales: { x: { grid: { display: false }, ticks: { color: '#78857b', font: { family: 'DM Mono', size: 10 } }, border: { display: false } }, y: { grid: { color: '#ebeee9' }, ticks: { color: '#78857b', font: { family: 'DM Mono', size: 10 }, callback: (value) => compactNumber(value) }, border: { display: false } } } } });
  $('#chartUnit').textContent = 'reported units';
}

function renderCommentary() {
  const grid = $('#commentaryGrid');
  if (!state.excerpts.length) { grid.innerHTML = '<div class="empty-state commentary-empty"><span class="empty-icon">◌</span><h3>No commentary extracted</h3><p>These source documents did not contain readable narrative text. Try a text-based PDF or a management discussion section.</p></div>'; return; }
  grid.innerHTML = state.excerpts.map((excerpt) => `<article class="commentary-card ${excerpt.type}"><span class="commentary-tag">${excerpt.type === 'risk' ? 'WATCH ITEM' : excerpt.type === 'opportunity' ? 'OPPORTUNITY' : 'MANAGEMENT FOCUS'}</span><h3>${excerpt.title}</h3><blockquote>“${excerpt.quote.replace(/"/g, '”')}”</blockquote><span class="commentary-source">SOURCE · ${excerpt.source}</span></article>`).join('');
}

function loadSample() {
  state.records = [
    { year: 2023, source: 'Northstar Industries FY2023.pdf', revenue: 18420, ebitda: 2940, netIncome: 1380, operatingCash: 2110, unit: '₹ Cr' },
    { year: 2024, source: 'Northstar Industries FY2024.pdf', revenue: 21560, ebitda: 3620, netIncome: 1750, operatingCash: 2650, unit: '₹ Cr' },
    { year: 2025, source: 'Northstar Industries FY2025.pdf', revenue: 24780, ebitda: 4090, netIncome: 1980, operatingCash: 3110, unit: '₹ Cr' },
  ];
  state.excerpts = [
    { title: 'Growth outlook', type: 'opportunity', quote: 'We remain confident that demand for our core platforms will support sustainable growth, aided by expansion in high-value enterprise accounts.', source: 'Annual report FY2025' },
    { title: 'Risk & uncertainty', type: 'risk', quote: 'Input-cost volatility and changing trade conditions remain areas of active monitoring, though our diversified supply base provides meaningful resilience.', source: 'Annual report FY2025' },
    { title: 'Operating priorities', type: 'standard', quote: 'Margin improvement was driven by disciplined pricing, a richer product mix and ongoing productivity initiatives across the operating model.', source: 'Annual report FY2025' },
  ];
  state.activeYear = 2025; renderDashboard();
  document.querySelector('#financials').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

$('#analyzeButton').addEventListener('click', analyzeFiles);
$('#sampleButton').addEventListener('click', loadSample);
$('#exportButton').addEventListener('click', () => {
  if (!state.records.length) return alert('Analyze documents or load the sample workspace first.');
  const latest = state.records.at(-1); const previous = state.records.at(-2);
  const lines = [`LedgerLens analysis summary`, `Generated ${new Date().toLocaleDateString()}`, '', `Latest period: FY ${latest.year}`, ...['revenue', 'ebitda', 'netIncome', 'operatingCash'].map((key) => `${key}: ${latest[key] ?? 'not found'}`), '', `Source: ${latest.source}`];
  if (previous) lines.push(`Revenue YoY: ${percentageChange(latest.revenue, previous.revenue)?.toFixed(1) ?? 'n/a'}%`);
  const blob = new Blob([lines.join('\n')], { type: 'text/plain' }); const anchor = document.createElement('a'); anchor.href = URL.createObjectURL(blob); anchor.download = 'ledgerlens-summary.txt'; anchor.click(); URL.revokeObjectURL(anchor.href);
});
$('#resetButton').addEventListener('click', () => { state.files = []; state.records = []; state.excerpts = []; state.activeYear = null; if (state.chart) state.chart.destroy(); $('#yearTabs').innerHTML = ''; $('#metricsGrid').innerHTML = '<div class="empty-state"><span class="empty-icon">⌁</span><h3>Waiting for source data</h3><p>Add at least two annual reports or financial data files to surface year-over-year changes.</p></div>'; $('#driversList').innerHTML = '<p class="placeholder">Financial drivers will appear after analysis.</p>'; $('#commentaryGrid').innerHTML = '<div class="empty-state commentary-empty"><span class="empty-icon">◌</span><h3>No commentary extracted</h3><p>Upload an annual report PDF or text file to identify management discussion themes.</p></div>'; $('#companyName').textContent = 'Workspace'; $('#yearRange').textContent = '—'; renderFiles(); });
