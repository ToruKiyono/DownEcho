const state = {
  records: [],
  filtered: [],
  sortKey: 'downloadTime',
  sortDirection: 'desc',
  settings: null
};

const dom = {
  body: document.body,
  search: document.getElementById('searchInput'),
  sizeMin: document.getElementById('sizeMin'),
  sizeMax: document.getElementById('sizeMax'),
  dateStart: document.getElementById('dateStart'),
  dateEnd: document.getElementById('dateEnd'),
  regexFilter: document.getElementById('regexFilter'),
  duplicateFilter: document.getElementById('duplicateFilter'),
  importButton: document.getElementById('importButton'),
  importInput: document.getElementById('popupImportInput'),
  exportButton: document.getElementById('exportButton'),
  refreshButton: document.getElementById('refreshButton'),
  tableBody: document.getElementById('recordsBody'),
  table: document.getElementById('recordsTable')
};

function applyTheme(theme) {
  const valid = ['sky', 'forest', 'rose'];
  const resolved = valid.includes(theme) ? theme : 'sky';
  dom.body.classList.remove('theme-sky', 'theme-forest', 'theme-rose');
  dom.body.classList.add(`theme-${resolved}`);
}

async function loadSettings() {
  const response = await chrome.runtime.sendMessage({ type: 'GET_SETTINGS' });
  if (response?.ok) {
    state.settings = response.settings;
    applyTheme(state.settings.theme);
  }
}

async function loadRecords() {
  const response = await chrome.runtime.sendMessage({ type: 'GET_RECORDS' });
  if (response?.ok) {
    state.records = Array.isArray(response.records) ? response.records : [];
  } else {
    state.records = [];
  }
  applyFilters();
}

function normalize(value) {
  if (!value && value !== 0) return '';
  return String(value).toLowerCase();
}

function formatSize(bytes) {
  if (!bytes || Number(bytes) === 0) return '—';
  const value = Number(bytes);
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  if (value < 1024 * 1024 * 1024) return `${(value / 1024 / 1024).toFixed(2)} MB`;
  return `${(value / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

function formatTime(value) {
  if (!value) return '—';
  try {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
  } catch (error) {
    return value;
  }
}

function matchesDateRange(value, start, end) {
  if (!start && !end) return true;
  const target = new Date(value);
  if (Number.isNaN(target.getTime())) return false;
  if (start) {
    const startDate = new Date(start);
    if (target < startDate) return false;
  }
  if (end) {
    const endDate = new Date(end);
    endDate.setHours(23, 59, 59, 999);
    if (target > endDate) return false;
  }
  return true;
}

function matchesSizeRange(bytes, min, max) {
  const size = Number(bytes) || 0;
  const minValue = min !== '' ? Number(min) * 1024 : null;
  const maxValue = max !== '' ? Number(max) * 1024 : null;
  if (minValue !== null && size < minValue) return false;
  if (maxValue !== null && size > maxValue) return false;
  return true;
}

function applyFilters() {
  const keyword = normalize(dom.search.value);
  const min = dom.sizeMin.value;
  const max = dom.sizeMax.value;
  const start = dom.dateStart.value;
  const end = dom.dateEnd.value;
  const regexFilter = dom.regexFilter.value;
  const duplicateFilter = dom.duplicateFilter.value;

  const filtered = state.records.filter(record => {
    const name = normalize(record.fileName);
    const source = normalize(record.sourceUrl);
    if (keyword && !name.includes(keyword) && !source.includes(keyword)) {
      return false;
    }
    if (!matchesSizeRange(record.fileSize, min, max)) {
      return false;
    }
    if (!matchesDateRange(record.downloadTime, start, end)) {
      return false;
    }
    if (regexFilter === 'matched' && !record.matchedRegex) {
      return false;
    }
    if (regexFilter === 'unmatched' && record.matchedRegex) {
      return false;
    }
    if (duplicateFilter === 'duplicates' && !record.duplicate) {
      return false;
    }
    if (duplicateFilter === 'unique' && record.duplicate) {
      return false;
    }
    return true;
  });

  state.filtered = sortRecords(filtered);
  render();
}

function sortRecords(records) {
  const { sortKey, sortDirection } = state;
  const sorted = [...records];
  sorted.sort((a, b) => {
    const valueA = a[sortKey];
    const valueB = b[sortKey];
    if (sortKey === 'fileSize') {
      return sortDirection === 'asc' ? (valueA - valueB) : (valueB - valueA);
    }
    if (sortKey === 'downloadTime') {
      const timeA = new Date(valueA).getTime();
      const timeB = new Date(valueB).getTime();
      return sortDirection === 'asc' ? (timeA - timeB) : (timeB - timeA);
    }
    const strA = normalize(valueA);
    const strB = normalize(valueB);
    if (strA < strB) return sortDirection === 'asc' ? -1 : 1;
    if (strA > strB) return sortDirection === 'asc' ? 1 : -1;
    return 0;
  });
  return sorted;
}

function render() {
  const rows = state.filtered;
  dom.tableBody.innerHTML = '';
  if (!rows.length) {
    const emptyRow = document.createElement('tr');
    emptyRow.className = 'placeholder';
    emptyRow.innerHTML = '<td colspan="5">暂无符合条件的记录</td>';
    dom.tableBody.appendChild(emptyRow);
    return;
  }

  const highlightRegex = state.settings?.highlightRegexHits !== false;
  for (const record of rows) {
    const tr = document.createElement('tr');
    const regexFlag = highlightRegex && record.matchedRegex
      ? `<span class="record-flag record-flag--regex" title="命中过滤规则：${escapeHtml(record.matchedRegex)}">REG</span>`
      : '';
    const duplicateFlag = record.duplicate
      ? `<span class="record-flag record-flag--duplicate" title="${escapeHtml(record.duplicateReason || '可能重复')}">DUP</span>`
      : '';
    tr.innerHTML = `
      <td>
        <div class="file-name">
          ${regexFlag}
          ${duplicateFlag}
          <span class="file-name__text" title="${escapeHtml(record.fileName)}">${escapeHtml(record.fileName)}</span>
        </div>
      </td>
      <td>${formatSize(record.fileSize)}</td>
      <td>${formatTime(record.downloadTime)}</td>
      <td>${renderSource(record.sourceUrl)}</td>
      <td>${renderStatus(record.status)}</td>
    `;
    dom.tableBody.appendChild(tr);
  }
}

function renderSource(url) {
  if (!url) return '—';
  const safe = escapeHtml(url);
  return `<a href="${safe}" target="_blank" rel="noreferrer noopener">${safe}</a>`;
}

function renderStatus(status) {
  const value = status || 'in_progress';
  const cls = value === 'complete' ? 'status-pill status-pill--complete'
    : (value === 'interrupted' || value === 'canceled') ? 'status-pill status-pill--interrupted'
      : 'status-pill';
  const label = {
    complete: '已完成',
    interrupted: '已中断',
    canceled: '已取消',
    in_progress: '下载中',
    imported: '已导入'
  }[value] || value;
  return `<span class="${cls}">${label}</span>`;
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, ch => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[ch] || ch));
}

async function importFromExcel(file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
      throw new Error('Excel 文件不包含可用数据');
    }
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const response = await chrome.runtime.sendMessage({ type: 'IMPORT_RECORDS', records: rows });
    if (!response?.ok) {
      throw new Error(response?.error || '导入失败');
    }
    await loadRecords();
    const added = Number(response.added) || 0;
    const updated = Number(response.updated) || 0;
    if (added > 0 && updated > 0) {
      alert(`导入完成，新增 ${added} 条并更新 ${updated} 条记录`);
    } else if (added > 0) {
      alert(`导入完成，新增 ${added} 条记录`);
    } else if (updated > 0) {
      alert(`导入完成，更新 ${updated} 条记录`);
    } else {
      alert('导入完成，没有新的记录');
    }
  } catch (error) {
    console.error('导入失败', error);
    alert(`导入失败：${error.message}`);
  }
}

function handleSort(event) {
  const th = event.target.closest('th');
  if (!th) return;
  const key = th.dataset.sort;
  if (!key) return;
  if (state.sortKey === key) {
    state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
  } else {
    state.sortKey = key;
    state.sortDirection = key === 'fileName' ? 'asc' : 'desc';
  }
  state.filtered = sortRecords(state.filtered);
  render();
}

function collectExportRows() {
  return state.filtered.map(record => ({
    '文件名': record.fileName,
    '文件大小': record.fileSize,
    '下载时间': record.downloadTime,
    '来源网址': record.sourceUrl || '',
    '状态': record.status || ''
  }));
}

function exportToExcel() {
  if (!state.filtered.length) {
    alert('没有可导出的记录');
    return;
  }
  const rows = collectExportRows();
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'DownEcho');
  const now = new Date();
  const filename = `DownEcho_Record_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}.xlsx`;
  XLSX.writeFile(workbook, filename);
}

function addEventListeners() {
  dom.search.addEventListener('input', debounce(applyFilters, 200));
  dom.sizeMin.addEventListener('input', debounce(applyFilters, 200));
  dom.sizeMax.addEventListener('input', debounce(applyFilters, 200));
  dom.dateStart.addEventListener('change', applyFilters);
  dom.dateEnd.addEventListener('change', applyFilters);
  dom.regexFilter.addEventListener('change', applyFilters);
  dom.duplicateFilter.addEventListener('change', applyFilters);
  dom.table.querySelector('thead').addEventListener('click', handleSort);
  dom.importButton.addEventListener('click', () => {
    dom.importInput.value = '';
    dom.importInput.click();
  });
  dom.importInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (file) {
      await importFromExcel(file);
    }
    dom.importInput.value = '';
  });
  dom.exportButton.addEventListener('click', exportToExcel);
  dom.refreshButton.addEventListener('click', () => {
    loadRecords();
  });
}

function debounce(fn, delay) {
  let timer = null;
  return function debounced(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

(async function init() {
  addEventListeners();
  await loadSettings();
  await loadRecords();
})();
