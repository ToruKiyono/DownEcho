const controls = {
  duplicateDetection: document.getElementById('duplicateDetection'),
  regexPrompt: document.getElementById('regexPrompt'),
  notificationsEnabled: document.getElementById('notificationsEnabled'),
  highlightRegex: document.getElementById('highlightRegex'),
  autoCleanEnabled: document.getElementById('autoCleanEnabled'),
  autoCleanDays: document.getElementById('autoCleanDays'),
  autoCleanDaysLabel: document.getElementById('autoCleanDaysLabel'),
  themeSelect: document.getElementById('themeSelect'),
  regexInput: document.getElementById('regexInput'),
  regexList: document.getElementById('regexList'),
  addRegex: document.getElementById('addRegex'),
  importInput: document.getElementById('importInput'),
  triggerImport: document.getElementById('triggerImport'),
  exportRecords: document.getElementById('exportRecords'),
  refreshPreview: document.getElementById('refreshPreview'),
  clearRecords: document.getElementById('clearRecords'),
  preview: document.getElementById('preview')
};

let settings = null;
let regexFilters = [];
let recordsCache = [];

const SIZE_UNITS = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];

function humanFileSize(bytes) {
  let value = Number(bytes);
  if (!Number.isFinite(value) || value < 0) {
    const text = String(bytes || '').trim().toLowerCase();
    const match = text.match(/^(\d+(?:\.\d+)?)\s*(b|kb|mb|gb|tb|pb)?$/);
    if (match) {
      value = parseFloat(match[1]);
      const unit = match[2] || 'b';
      const index = SIZE_UNITS.findIndex(item => item.toLowerCase() === unit);
      if (index >= 0) {
        value *= Math.pow(1024, index);
      }
    } else {
      value = 0;
    }
  }
  if (!Number.isFinite(value) || value < 0) {
    value = 0;
  }
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < SIZE_UNITS.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(2)} ${SIZE_UNITS[unitIndex]}`;
}

function extractRowsFromWorkbook(workbook) {
  if (!workbook || typeof workbook !== 'object') {
    return [];
  }
  const candidates = Array.isArray(workbook.SheetNames) && workbook.SheetNames.length
    ? workbook.SheetNames
    : Object.keys(workbook.Sheets || {});
  for (const name of candidates) {
    const sheet = workbook.Sheets?.[name];
    if (!sheet) continue;

    const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    const populatedRows = jsonRows.filter(row => Object.values(row).some(value => String(value ?? '').trim() !== ''));
    if (populatedRows.length) {
      return populatedRows;
    }

    const matrix = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1, raw: false });
    if (matrix.length > 1) {
      const [headers, ...dataRows] = matrix;
      const normalizedHeaders = headers.map(header => String(header || '').trim());
      const remapped = dataRows
        .map(row => {
          const hasContent = row.some(cell => String(cell ?? '').trim() !== '');
          if (!hasContent) return null;
          const entry = {};
          normalizedHeaders.forEach((header, index) => {
            if (header) {
              entry[header] = row[index];
            }
          });
          return entry;
        })
        .filter(Boolean);
      if (remapped.length) {
        return remapped;
      }
    }
  }
  return [];
}

function applyTheme(theme) {
  const root = document.documentElement;
  root.dataset.theme = theme;
  switch (theme) {
    case 'forest':
      root.style.setProperty('background', 'linear-gradient(180deg, #e9fff0 0%, #f9fff6 100%)');
      document.body.style.color = '#145a32';
      break;
    case 'rose':
      root.style.setProperty('background', 'linear-gradient(180deg, #fff0f6 0%, #fff7fb 100%)');
      document.body.style.color = '#7a1f4b';
      break;
    default:
      root.style.setProperty('background', 'linear-gradient(180deg, #e8f4ff 0%, #fdfbff 100%)');
      document.body.style.color = '#0a3d62';
  }
}

function updateAutoCleanLabel(value) {
  controls.autoCleanDaysLabel.textContent = value || settings?.autoCleanDays || 30;
}

async function loadSettings() {
  const response = await chrome.runtime.sendMessage({ type: 'GET_SETTINGS' });
  if (!response?.ok) {
    console.error(response?.error || '无法加载设置');
    return;
  }
  settings = response.settings;
  regexFilters = [...(settings.regexFilters || [])];
  controls.duplicateDetection.checked = Boolean(settings.duplicateDetection);
  controls.regexPrompt.checked = Boolean(settings.regexPromptEnabled);
  controls.notificationsEnabled.checked = Boolean(settings.notificationsEnabled);
  controls.highlightRegex.checked = Boolean(settings.highlightRegexHits);
  controls.autoCleanEnabled.checked = Boolean(settings.autoCleanEnabled);
  controls.autoCleanDays.value = settings.autoCleanDays || 30;
  controls.themeSelect.value = settings.theme || 'sky';
  applyTheme(settings.theme);
  updateAutoCleanLabel(settings.autoCleanDays);
  renderRegexList();
}

async function saveSettings(partial) {
  const payload = { ...settings, ...partial };
  const response = await chrome.runtime.sendMessage({ type: 'SAVE_SETTINGS', settings: payload });
  if (response?.ok) {
    settings = response.settings;
    regexFilters = [...(settings.regexFilters || [])];
    updateAutoCleanLabel(settings.autoCleanDays);
  } else {
    throw new Error(response?.error || '保存设置失败');
  }
}

function renderRegexList() {
  controls.regexList.innerHTML = '';
  if (!regexFilters.length) {
    const empty = document.createElement('li');
    empty.textContent = '尚未设置过滤规则';
    empty.className = 'regex-item regex-item--empty';
    controls.regexList.appendChild(empty);
    return;
  }

  regexFilters.forEach((rule, index) => {
    const li = document.createElement('li');
    li.className = 'regex-item';
    li.innerHTML = `
      <span class="regex-item__rule" title="${escapeHtml(rule)}">${escapeHtml(rule)}</span>
      <div class="regex-item__actions">
        <button class="btn btn--secondary" data-action="edit" data-index="${index}">编辑</button>
        <button class="btn btn--danger" data-action="remove" data-index="${index}">删除</button>
      </div>
    `;
    controls.regexList.appendChild(li);
  });
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

function validateRegexInput(value) {
  const trimmed = value.trim();
  if (!trimmed) {
    throw new Error('请输入正则表达式');
  }
  try {
    // eslint-disable-next-line no-new
    new RegExp(trimmed);
  } catch (error) {
    throw new Error('无效的正则表达式');
  }
  return trimmed;
}

async function addRegexRule() {
  try {
    const value = validateRegexInput(controls.regexInput.value);
    if (regexFilters.includes(value)) {
      alert('该规则已存在');
      return;
    }
    regexFilters.push(value);
    await saveSettings({ regexFilters });
    controls.regexInput.value = '';
    renderRegexList();
    alert('规则已添加');
  } catch (error) {
    alert(error.message);
  }
}

async function editRegexRule(index) {
  const current = regexFilters[index];
  const next = prompt('修改正则规则', current);
  if (next === null) return;
  try {
    const value = validateRegexInput(next);
    regexFilters[index] = value;
    await saveSettings({ regexFilters });
    renderRegexList();
    alert('规则已更新');
  } catch (error) {
    alert(error.message);
  }
}

async function removeRegexRule(index) {
  if (!confirm('确定删除该规则吗？')) return;
  regexFilters.splice(index, 1);
  await saveSettings({ regexFilters });
  renderRegexList();
  alert('规则已删除');
}

function bindSettingListeners() {
  controls.duplicateDetection.addEventListener('change', () => saveSettings({ duplicateDetection: controls.duplicateDetection.checked }).catch(handleError));
  controls.regexPrompt.addEventListener('change', () => saveSettings({ regexPromptEnabled: controls.regexPrompt.checked }).catch(handleError));
  controls.notificationsEnabled.addEventListener('change', () => saveSettings({ notificationsEnabled: controls.notificationsEnabled.checked }).catch(handleError));
  controls.highlightRegex.addEventListener('change', () => saveSettings({ highlightRegexHits: controls.highlightRegex.checked }).catch(handleError));
  controls.autoCleanEnabled.addEventListener('change', () => saveSettings({ autoCleanEnabled: controls.autoCleanEnabled.checked }).catch(handleError));
  controls.autoCleanDays.addEventListener('change', () => {
    const value = Number(controls.autoCleanDays.value) || 1;
    controls.autoCleanDays.value = value;
    updateAutoCleanLabel(value);
    saveSettings({ autoCleanDays: value }).catch(handleError);
  });
  controls.themeSelect.addEventListener('change', () => {
    const value = controls.themeSelect.value;
    applyTheme(value);
    saveSettings({ theme: value }).catch(handleError);
  });
}

controls.regexList.addEventListener('click', event => {
  const button = event.target.closest('button');
  if (!button) return;
  const index = Number(button.dataset.index);
  if (Number.isNaN(index)) return;
  const action = button.dataset.action;
  if (action === 'edit') {
    editRegexRule(index).catch(handleError);
  } else if (action === 'remove') {
    removeRegexRule(index).catch(handleError);
  }
});

controls.addRegex.addEventListener('click', () => {
  addRegexRule().catch(handleError);
});

controls.triggerImport.addEventListener('click', () => {
  controls.importInput.click();
});

controls.regexInput.addEventListener('keydown', event => {
  if (event.key === 'Enter') {
    event.preventDefault();
    addRegexRule().catch(handleError);
  }
});

controls.importInput.addEventListener('change', async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const rows = extractRowsFromWorkbook(workbook);
    if (!rows.length) {
      throw new Error('Excel 文件不包含可用数据');
    }
    const response = await chrome.runtime.sendMessage({ type: 'IMPORT_RECORDS', records: rows });
    if (!response?.ok) throw new Error(response?.error || '导入失败');
    recordsCache = response.records;
    renderPreview();
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
    console.error(error);
    alert(`导入失败：${error.message}`);
  } finally {
    controls.importInput.value = '';
  }
});

controls.exportRecords.addEventListener('click', async () => {
  try {
    const response = await chrome.runtime.sendMessage({ type: 'GET_RECORDS' });
    if (!response?.ok) throw new Error(response?.error || '无法获取记录');
    const rows = response.records.map(record => ({
      '文件名': record.fileName,
      '文件大小': humanFileSize(record.fileSize),
      '下载时间': record.downloadTime,
      '来源网址': record.sourceUrl || '',
      '状态': record.status || ''
    }));
    if (!rows.length) {
      alert('没有可导出的记录');
      return;
    }
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'DownEcho');
    const now = new Date();
    const filename = `DownEcho_Record_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}.xlsx`;
    XLSX.writeFile(workbook, filename);
  } catch (error) {
    handleError(error);
  }
});

controls.refreshPreview.addEventListener('click', async () => {
  try {
    const response = await chrome.runtime.sendMessage({ type: 'GET_RECORDS' });
    if (!response?.ok) throw new Error(response?.error || '无法加载记录');
    recordsCache = response.records;
    renderPreview();
  } catch (error) {
    handleError(error);
  }
});

controls.clearRecords.addEventListener('click', async () => {
  if (!confirm('确定清空所有记录？此操作不可恢复。')) return;
  try {
    const response = await chrome.runtime.sendMessage({ type: 'CLEAR_RECORDS' });
    if (!response?.ok) throw new Error(response?.error || '清空失败');
    recordsCache = [];
    renderPreview();
    alert('记录已清空');
  } catch (error) {
    handleError(error);
  }
});

function renderPreview() {
  if (!recordsCache.length) {
    controls.preview.innerHTML = '<p>暂无记录，请导入或等待下载事件。</p>';
    return;
  }
  const head = ['文件名', '文件大小', '下载时间', '来源网址', '状态'];
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  head.forEach(title => {
    const th = document.createElement('th');
    th.textContent = title;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  recordsCache.slice(0, 100).forEach(record => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(record.fileName || '')}</td>
      <td>${escapeHtml(formatPreviewSize(record.fileSize))}</td>
      <td>${escapeHtml(record.downloadTime || '')}</td>
      <td>${escapeHtml(record.sourceUrl || '')}</td>
      <td>${escapeHtml(record.status || '')}</td>
    `;
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  controls.preview.innerHTML = '';
  controls.preview.appendChild(table);
}

function formatPreviewSize(value) {
  if (value == null || value === '') return '—';
  const formatted = humanFileSize(value);
  return formatted === '0.00 B' ? '—' : formatted;
}

function handleError(error) {
  console.error(error);
  alert(error?.message || '发生未知错误');
}

(async function init() {
  bindSettingListeners();
  await loadSettings();
  try {
    const response = await chrome.runtime.sendMessage({ type: 'GET_RECORDS' });
    if (response?.ok) {
      recordsCache = response.records;
      renderPreview();
    }
  } catch (error) {
    console.error('加载记录失败', error);
  }
})();
