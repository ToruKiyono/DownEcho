const DEFAULT_SETTINGS = {
  duplicateDetection: true,
  notificationsEnabled: true,
  regexFilters: [],
  regexPromptEnabled: true,
  autoCleanEnabled: true,
  autoCleanDays: 30,
  theme: 'sky',
  highlightRegexHits: true
};

const STORAGE_KEYS = {
  records: 'downloadRecords',
  settings: 'downEchoSettings'
};

const pendingDecisions = new Map();

const FALLBACK_ICON_DATA_URL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/Pm7kWQAAAABJRU5ErkJggg==';

function resolveNotificationIcon() {
  try {
    return chrome.runtime.getURL('icons/icon128.png');
  } catch (error) {
    console.warn('Failed to resolve notification icon URL', error);
    return FALLBACK_ICON_DATA_URL;
  }
}

async function safeCreateNotification(notificationId, options) {
  try {
    await chrome.notifications.create(notificationId, options);
  } catch (error) {
    if (error && typeof error.message === 'string' && error.message.includes('Unable to download all specified images')) {
      const fallbackOptions = { ...options, iconUrl: FALLBACK_ICON_DATA_URL };
      try {
        await chrome.notifications.create(notificationId, fallbackOptions);
        return;
      } catch (innerError) {
        console.error('Fallback notification creation failed', innerError);
      }
    }
    console.error('Notification creation failed', error);
    throw error;
  }
}

function sanitize(text) {
  if (!text) return '';
  return String(text).replace(/[\u0000-\u001f\u007f]/g, '').trim();
}

function decodeNameSegment(value) {
  if (!value) return '';
  let decoded = value;
  try {
    decoded = decodeURIComponent(value.replace(/\+/g, ' '));
  } catch (error) {
    try {
      decoded = decodeURI(value);
    } catch (innerError) {
      decoded = value;
    }
  }
  return sanitize(decoded) || '';
}

function extractFileName(input) {
  const sanitized = sanitize(input);
  if (!sanitized) return '';
  let candidate = sanitized;
  try {
    const url = new URL(sanitized);
    if (url.pathname && url.pathname !== '/') {
      candidate = url.pathname;
    } else if (url.hostname) {
      candidate = url.hostname;
    }
  } catch (error) {
    // Not a URL, keep the sanitized value
  }
  const segments = candidate.split(/[\\\/]/).filter(Boolean);
  const baseName = segments.length ? segments[segments.length - 1] : candidate;
  const decoded = decodeNameSegment(baseName);
  return decoded || sanitize(baseName) || sanitized;
}

function normalizedName(name) {
  return extractFileName(name).toLowerCase();
}

async function getSettings() {
  const stored = await chrome.storage.local.get(STORAGE_KEYS.settings);
  return { ...DEFAULT_SETTINGS, ...(stored[STORAGE_KEYS.settings] || {}) };
}

async function setSettings(settings) {
  await chrome.storage.local.set({
    [STORAGE_KEYS.settings]: settings
  });
}

function normalizeRecord(record) {
  if (!record || typeof record !== 'object') return record;
  const normalizedFileName = extractFileName(record.fileName || '');
  if (normalizedFileName && normalizedFileName !== record.fileName) {
    return { ...record, fileName: normalizedFileName };
  }
  return record;
}

async function getRecords() {
  const stored = await chrome.storage.local.get(STORAGE_KEYS.records);
  const rawRecords = Array.isArray(stored[STORAGE_KEYS.records]) ? stored[STORAGE_KEYS.records] : [];
  let changed = false;
  const normalizedRecords = rawRecords.map(record => {
    const normalized = normalizeRecord(record);
    if (normalized !== record) {
      changed = true;
    }
    return normalized;
  });
  if (changed) {
    await saveRecords(normalizedRecords);
    return normalizedRecords;
  }
  return rawRecords;
}

async function saveRecords(records) {
  const normalizedList = Array.isArray(records) ? records.map(normalizeRecord) : [];
  await chrome.storage.local.set({
    [STORAGE_KEYS.records]: normalizedList
  });
}

function formatDate(date = new Date()) {
  return date.toISOString();
}

function deriveFileName(downloadItem) {
  const candidates = [
    downloadItem.filename,
    downloadItem.suggestedFilename,
    downloadItem.targetPath,
    downloadItem.finalUrl,
    downloadItem.url
  ];
  for (const value of candidates) {
    const extracted = extractFileName(value);
    if (extracted) {
      return extracted;
    }
  }
  return '未知文件';
}

function computeSize(item) {
  if (typeof item.fileSize === 'number') return item.fileSize;
  if (typeof item.totalBytes === 'number' && item.totalBytes > 0) {
    return item.totalBytes;
  }
  if (typeof item.bytesReceived === 'number') {
    return item.bytesReceived;
  }
  return 0;
}

const SIZE_UNITS = ['b', 'kb', 'mb', 'gb', 'tb', 'pb'];

function parseFileSize(value) {
  if (value == null) return 0;
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  const text = sanitize(String(value).replace(/,/g, ''));
  if (!text) return 0;
  const normalized = text.toLowerCase();
  const match = normalized.match(/^(\d+(?:\.\d+)?)\s*(b|kb|mb|gb|tb|pb)?$/);
  if (match) {
    const numeric = parseFloat(match[1]);
    if (!Number.isFinite(numeric)) return 0;
    const unit = match[2] || 'b';
    const index = SIZE_UNITS.indexOf(unit);
    const multiplier = index >= 0 ? Math.pow(1024, index) : 1;
    return Math.round(numeric * multiplier);
  }
  const fallback = Number(normalized);
  return Number.isFinite(fallback) ? fallback : 0;
}

function evaluateRegex(fileName, filters) {
  if (!Array.isArray(filters)) return null;
  for (const rule of filters) {
    if (typeof rule !== 'string' || !rule.trim()) continue;
    try {
      const regex = new RegExp(rule);
      if (regex.test(fileName)) {
        return rule;
      }
    } catch (error) {
      console.warn('Invalid regex rule skipped', rule, error);
    }
  }
  return null;
}

async function ensureDefaults() {
  const settings = await getSettings();
  await setSettings(settings);
  const existing = await getRecords();
  if (!Array.isArray(existing)) {
    await saveRecords([]);
  }
}

async function cleanupOldRecords(settings) {
  if (!settings.autoCleanEnabled) return;
  const records = await getRecords();
  if (!records.length) return;
  const now = Date.now();
  const threshold = settings.autoCleanDays * 24 * 60 * 60 * 1000;
  const filtered = records.filter(record => {
    const time = new Date(record.downloadTime || 0).getTime();
    if (!Number.isFinite(time)) return true;
    return now - time <= threshold;
  });
  if (filtered.length !== records.length) {
    await saveRecords(filtered);
  }
}

function createNotificationId(downloadId, reason) {
  return `downecho_${downloadId}_${reason}_${Date.now()}`;
}

function createInfoNotificationId() {
  return `downecho_info_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

async function sendDecisionNotification({
  download,
  reason,
  matchedRegex,
  settings,
  message,
  recordId,
  force = false
}) {
  if (!force && !settings.notificationsEnabled) return null;
  const notificationId = createNotificationId(download.id, reason);
  const buttons = [
    { title: '继续下载' },
    { title: '取消下载' }
  ];
  const title = reason === 'regex'
    ? '命中下载过滤规则'
    : '检测到可能的重复下载';
  const context = matchedRegex ? `规则：${matchedRegex}` : '';
  const iconUrl = resolveNotificationIcon();
  const options = {
    type: 'basic',
    iconUrl,
    title,
    message,
    contextMessage: context,
    buttons
  };
  await safeCreateNotification(notificationId, options);
  pendingDecisions.set(notificationId, {
    downloadId: download.id,
    reason,
    matchedRegex,
    recordId: recordId ? String(recordId) : String(download.id)
  });
  return notificationId;
}

async function showSimpleNotification(settings, title, message) {
  if (!settings.notificationsEnabled) return null;
  const notificationId = createInfoNotificationId();
  const iconUrl = resolveNotificationIcon();
  const options = {
    type: 'basic',
    iconUrl,
    title,
    message
  };
  await safeCreateNotification(notificationId, options);
  return notificationId;
}

async function updateRecord(recordId, updates) {
  const records = await getRecords();
  const index = records.findIndex(record => record.id === recordId);
  if (index >= 0) {
    records[index] = { ...records[index], ...updates };
    await saveRecords(records);
  }
}

async function handleDownloadCreated(downloadItem) {
  const [settings, records] = await Promise.all([
    getSettings(),
    getRecords()
  ]);
  cleanupOldRecords(settings).catch(error => console.warn('Cleanup failed', error));
  const fileName = deriveFileName(downloadItem);
  const downloadTime = formatDate(new Date());
  const sourceUrl = sanitize(downloadItem.finalUrl || downloadItem.referrer || downloadItem.url || '');
  const fileSize = computeSize(downloadItem);

  const duplicateReasons = [];
  const normalized = normalizedName(fileName);
  if (settings.duplicateDetection) {
    const existingByName = records.find(record => normalizedName(record.fileName || '') === normalized);
    if (existingByName) {
      duplicateReasons.push('相同文件名');
    }

    const tolerance = 1024;
    const targetSize = fileSize;
    if (targetSize > 0) {
      const similar = records.find(record => {
        const size = Number(record.fileSize) || 0;
        return Math.abs(size - targetSize) <= tolerance;
      });
      if (similar) {
        duplicateReasons.push('相近文件大小');
      }
    }
  }

  const matchedRegex = evaluateRegex(fileName, settings.regexFilters);
  const regexReason = matchedRegex && settings.regexPromptEnabled
    ? `命中过滤规则：${matchedRegex}`
    : null;

  const duplicate = duplicateReasons.length > 0 && settings.duplicateDetection;
  const regexFlagged = Boolean(regexReason);
  const requiresDecision = duplicate || regexFlagged;

  let pausedForDecision = false;
  if (requiresDecision) {
    try {
      await chrome.downloads.pause(downloadItem.id);
      pausedForDecision = true;
    } catch (error) {
      console.warn('Pause download for decision failed', error);
    }
  }

  const reasonNotes = [];
  if (duplicateReasons.length) {
    reasonNotes.push(duplicateReasons.join('；'));
  }
  if (regexReason) {
    reasonNotes.push(regexReason);
  }

  const record = {
    id: String(downloadItem.id),
    fileName,
    fileSize,
    downloadTime,
    sourceUrl,
    status: pausedForDecision ? 'awaiting_user_confirmation' : (downloadItem.state || 'in_progress'),
    duplicate,
    matchedRegex: matchedRegex || undefined,
    duplicateReason: requiresDecision ? (reasonNotes.join('；') || undefined) : undefined
  };

  records.push(record);
  await saveRecords(records);

  if (requiresDecision) {
    const messages = [];
    if (duplicate) {
      const reasonText = duplicateReasons.length ? duplicateReasons.join('；') : '检测到重复下载';
      messages.push(`下载的文件可能重复：${reasonText}`);
    }
    if (regexFlagged) {
      messages.push(regexReason || `文件名命中过滤规则：${matchedRegex}`);
    }
    const decisionReason = duplicate && regexFlagged ? 'duplicate_regex'
      : duplicate ? 'duplicate'
        : 'regex';
    try {
      await sendDecisionNotification({
        download: downloadItem,
        reason: decisionReason,
        matchedRegex,
        settings,
        message: messages.join('\n'),
        recordId: record.id,
        force: true
      });
    } catch (error) {
      console.error('Failed to deliver decision notification', error);
      try {
        await chrome.downloads.resume(downloadItem.id);
      } catch (resumeError) {
        console.warn('Resume download after notification failure failed', resumeError);
      }
      await updateRecord(record.id, {
        status: downloadItem.state || 'in_progress'
      });
    }
  }
}

async function handleDownloadChanged(delta) {
  const downloadId = delta.id;
  const updates = {};
  if (delta.filename && delta.filename.current) {
    updates.fileName = extractFileName(delta.filename.current);
  }
  if (delta.state && delta.state.current) {
    updates.status = delta.state.current;
  }
  if (delta.bytesReceived && typeof delta.bytesReceived.current === 'number') {
    updates.fileSize = delta.bytesReceived.current;
  }
  if (delta.totalBytes && typeof delta.totalBytes.current === 'number') {
    updates.fileSize = delta.totalBytes.current;
  }
  if (Object.keys(updates).length > 0) {
    await updateRecord(String(downloadId), updates);
  }
}

chrome.runtime.onInstalled.addListener(async () => {
  await ensureDefaults();
});

chrome.runtime.onStartup.addListener(async () => {
  const settings = await getSettings();
  await cleanupOldRecords(settings);
});

chrome.downloads.onCreated.addListener(handleDownloadCreated);
chrome.downloads.onChanged.addListener(handleDownloadChanged);

chrome.notifications.onButtonClicked.addListener(async (notificationId, buttonIndex) => {
  if (!pendingDecisions.has(notificationId)) return;
  const entry = pendingDecisions.get(notificationId);
  pendingDecisions.delete(notificationId);
  const downloadId = entry.downloadId;
  const recordId = entry.recordId;
  if (buttonIndex === 0) {
    if (recordId) {
      await updateRecord(String(recordId), { status: 'in_progress' });
    }
    try {
      await chrome.downloads.resume(downloadId);
    } catch (error) {
      console.warn('Resume download failed', error);
    }
  } else if (buttonIndex === 1) {
    try {
      await chrome.downloads.cancel(downloadId);
      if (recordId) {
        await updateRecord(String(recordId), { status: 'canceled' });
      }
    } catch (error) {
      console.warn('Cancel download failed', error);
    }
  }
  try {
    await chrome.notifications.clear(notificationId);
  } catch (error) {
    console.warn('Failed to clear notification', error);
  }
});

chrome.notifications.onClosed.addListener(notificationId => {
  if (!pendingDecisions.has(notificationId)) return;
  const entry = pendingDecisions.get(notificationId);
  pendingDecisions.delete(notificationId);
  const downloadId = entry.downloadId;
  const recordId = entry.recordId;
  (async () => {
    try {
      await chrome.downloads.cancel(downloadId);
    } catch (error) {
      console.warn('Cancel download after notification closed failed', error);
    }
    if (recordId) {
      await updateRecord(String(recordId), { status: 'canceled' });
    }
  })();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  const respond = (data) => {
    try {
      sendResponse(data);
    } catch (error) {
      console.warn('Failed to send response', error);
    }
  };
  if (!message || typeof message.type !== 'string') {
    respond({ ok: false, error: 'Invalid message' });
    return false;
  }
  switch (message.type) {
    case 'GET_RECORDS':
      getRecords().then(records => respond({ ok: true, records })).catch(error => respond({ ok: false, error: error?.message }));
      return true;
    case 'GET_SETTINGS':
      getSettings().then(settings => respond({ ok: true, settings })).catch(error => respond({ ok: false, error: error?.message }));
      return true;
    case 'SAVE_SETTINGS':
      getSettings().then(current => {
        const merged = { ...current, ...(message.settings || {}) };
        return setSettings(merged).then(() => respond({ ok: true, settings: merged }));
      }).catch(error => respond({ ok: false, error: error?.message }));
      return true;
    case 'CLEAR_RECORDS':
      saveRecords([]).then(() => respond({ ok: true })).catch(error => respond({ ok: false, error: error?.message }));
      return true;
    case 'IMPORT_RECORDS': {
      const incoming = Array.isArray(message.records) ? message.records : [];
      Promise.all([getRecords(), getSettings()]).then(async ([records, settings]) => {
        const merged = [...records];
        const existingIndexByName = new Map();
        records.forEach((record, index) => {
          const key = normalizedName(record.fileName || '');
          if (key) {
            existingIndexByName.set(key, index);
          }
        });
        let added = 0;
        let updated = 0;
        for (const item of incoming) {
          if (!item || typeof item !== 'object') continue;
          const name = extractFileName(item.fileName || item['文件名'] || '');
          if (!name) continue;
          const normalizedImportName = normalizedName(name);
          const size = parseFileSize(item.fileSize ?? item['文件大小']);
          const time = sanitize(item.downloadTime || item['下载时间'] || formatDate());
          const sourceUrl = sanitize(item.sourceUrl || item['来源网址'] || '');
          const status = sanitize(item.status || item['状态'] || 'imported');
          const matchedRegex = evaluateRegex(name, settings.regexFilters);
          if (existingIndexByName.has(normalizedImportName)) {
            const index = existingIndexByName.get(normalizedImportName);
            const existing = merged[index];
            const updatedRecord = {
              ...existing,
              fileName: name,
              fileSize: size || existing.fileSize || 0,
              downloadTime: time || existing.downloadTime || formatDate(),
              sourceUrl: sourceUrl || existing.sourceUrl || '',
              status: status || existing.status || 'imported',
              matchedRegex: matchedRegex || existing.matchedRegex,
              duplicateReason: existing.duplicateReason
            };
            if (JSON.stringify(existing) !== JSON.stringify(updatedRecord)) {
              merged[index] = updatedRecord;
              updated += 1;
            }
            continue;
          }
          const newRecord = {
            id: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
            fileName: name,
            fileSize: size,
            downloadTime: time,
            sourceUrl,
            status: status || 'imported',
            duplicate: false,
            matchedRegex: matchedRegex || undefined,
            duplicateReason: undefined
          };
          merged.push(newRecord);
          existingIndexByName.set(normalizedImportName, merged.length - 1);
          added += 1;
        }
        await saveRecords(merged);
        let message = '没有新的记录需要导入';
        if (added > 0 && updated > 0) {
          message = `新增 ${added} 条，更新 ${updated} 条记录`;
        } else if (added > 0) {
          message = `成功导入 ${added} 条新记录`;
        } else if (updated > 0) {
          message = `更新 ${updated} 条已存在记录`;
        }
        await showSimpleNotification(settings, '导入完成', message);
        respond({ ok: true, records: merged, added, updated });
      }).catch(async error => {
        const settings = await getSettings();
        await showSimpleNotification(settings, '导入失败', error?.message || '导入失败');
        respond({ ok: false, error: error?.message });
      });
      return true;
    }
    case 'DELETE_RECORD':
      if (!message.id) {
        respond({ ok: false, error: 'Missing id' });
        return false;
      }
      getRecords().then(async records => {
        const filtered = records.filter(record => record.id !== message.id);
        await saveRecords(filtered);
        respond({ ok: true, records: filtered });
      }).catch(error => respond({ ok: false, error: error?.message }));
      return true;
    default:
      respond({ ok: false, error: 'Unknown message type' });
      return false;
  }
});
