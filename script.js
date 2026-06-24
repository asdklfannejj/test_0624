const state = {
  rawRows: [],
  headers: [],
  filters: {},
  search: '',
  sort: { index: -1, direction: 'asc' },
  fileName: '',
};

const els = {
  fileInput: document.getElementById('fileInput'),
  searchInput: document.getElementById('searchInput'),
  resetBtn: document.getElementById('resetBtn'),
  columnFilters: document.getElementById('columnFilters'),
  tableHead: document.getElementById('tableHead'),
  tableBody: document.getElementById('tableBody'),
  emptyMessage: document.getElementById('emptyMessage'),
  fileName: document.getElementById('fileName'),
  totalRows: document.getElementById('totalRows'),
  visibleRows: document.getElementById('visibleRows'),
};

function decodeExcelValue(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value;
  return value;
}

function readFirstSheet(file) {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false,
      raw: false,
    });
  });
}

function normalizeRows(rows) {
  if (!rows.length) return { headers: [], data: [] };
  const headers = rows[0].map((header, index) => String(header || `Column ${index + 1}`).trim());
  const data = rows.slice(1).map((row) => {
    return headers.map((_, index) => decodeExcelValue(row[index] ?? ''));
  });
  return { headers, data };
}

function isDateLike(value) {
  if (value instanceof Date) return true;
  if (typeof value !== 'string') return false;
  const s = value.trim();
  if (!s) return false;
  const d = new Date(s);
  return !Number.isNaN(d.getTime()) && /[-/.]/.test(s);
}

function isNumericLike(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return true;
  if (typeof value !== 'string') return false;
  const s = value.replace(/,/g, '').trim();
  return s !== '' && Number.isFinite(Number(s));
}

function toComparableValue(value) {
  if (value instanceof Date) return value.getTime();
  if (isNumericLike(value)) return Number(String(value).replace(/,/g, ''));
  if (isDateLike(value)) return new Date(value).getTime();
  return String(value).toLowerCase();
}

function formatCell(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value);
}

function uniqueValuesForColumn(data, colIndex) {
  const values = new Set();
  data.forEach((row) => {
    values.add(formatCell(row[colIndex]));
  });
  return Array.from(values).sort((a, b) => a.localeCompare(b, 'ko'));
}

function applyFiltersAndSort() {
  const searchTerm = state.search.trim().toLowerCase();
  const filtered = state.rawRows.filter((row) => {
    const matchesSearch = !searchTerm || row.some((cell) => formatCell(cell).toLowerCase().includes(searchTerm));
    const matchesColumns = Object.entries(state.filters).every(([idx, value]) => {
      if (!value) return true;
      return formatCell(row[Number(idx)]) === value;
    });
    return matchesSearch && matchesColumns;
  });

  if (state.sort.index >= 0) {
    const { index, direction } = state.sort;
    filtered.sort((a, b) => {
      const left = toComparableValue(a[index]);
      const right = toComparableValue(b[index]);
      if (left < right) return direction === 'asc' ? -1 : 1;
      if (left > right) return direction === 'asc' ? 1 : -1;
      return 0;
    });
  }

  return filtered;
}

function renderColumnFilters() {
  els.columnFilters.innerHTML = '';
  state.headers.forEach((header, index) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'filter-group';

    const label = document.createElement('label');
    label.textContent = header;

    const select = document.createElement('select');
    select.className = 'select';
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = `All ${header}`;
    select.appendChild(defaultOption);

    uniqueValuesForColumn(state.rawRows, index).forEach((value) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value === '' ? '(empty)' : value;
      select.appendChild(option);
    });

    select.addEventListener('change', () => {
      state.filters[index] = select.value;
      renderTable();
    });

    wrapper.append(label, select);
    els.columnFilters.appendChild(wrapper);
  });
}

function renderTableHead() {
  els.tableHead.innerHTML = '';
  const tr = document.createElement('tr');

  state.headers.forEach((header, index) => {
    const th = document.createElement('th');
    th.dataset.index = String(index);
    th.innerHTML = `${header}<span class="sort-indicator">${state.sort.index === index ? (state.sort.direction === 'asc' ? '▲' : '▼') : ''}</span>`;
    th.addEventListener('click', () => {
      if (state.sort.index === index) {
        state.sort.direction = state.sort.direction === 'asc' ? 'desc' : 'asc';
      } else {
        state.sort.index = index;
        state.sort.direction = 'asc';
      }
      renderTable();
    });
    tr.appendChild(th);
  });

  els.tableHead.appendChild(tr);
}

function renderTableBody(rows) {
  els.tableBody.innerHTML = '';

  if (!rows.length) {
    els.emptyMessage.classList.remove('hidden');
    els.tableBody.innerHTML = '<tr class="empty-row"><td colspan="100%">표시할 데이터가 없습니다.</td></tr>';
    return;
  }

  els.emptyMessage.classList.add('hidden');
  const fragment = document.createDocumentFragment();

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    row.forEach((cell) => {
      const td = document.createElement('td');
      td.textContent = formatCell(cell);
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });

  els.tableBody.appendChild(fragment);
}

function updateStats(visibleCount) {
  els.totalRows.textContent = String(state.rawRows.length);
  els.visibleRows.textContent = String(visibleCount);
}

function renderTable() {
  renderTableHead();
  const rows = applyFiltersAndSort();
  renderTableBody(rows);
  updateStats(rows.length);
}

function resetFilters() {
  state.search = '';
  state.filters = {};
  state.sort = { index: -1, direction: 'asc' };
  els.searchInput.value = '';
  [...els.columnFilters.querySelectorAll('select')].forEach((select) => {
    select.value = '';
  });
  renderTable();
}

async function handleFileUpload(file) {
  const rows = await readFirstSheet(file);
  const normalized = normalizeRows(rows);
  state.rawRows = normalized.data;
  state.headers = normalized.headers;
  state.fileName = file.name;
  state.filters = {};
  state.sort = { index: -1, direction: 'asc' };

  els.fileName.textContent = file.name;
  els.searchInput.disabled = false;
  els.resetBtn.disabled = false;
  renderColumnFilters();
  renderTable();
}

function initEvents() {
  els.fileInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    await handleFileUpload(file);
  });

  els.searchInput.addEventListener('input', () => {
    state.search = els.searchInput.value;
    renderTable();
  });

  els.resetBtn.addEventListener('click', resetFilters);
}

initEvents();
