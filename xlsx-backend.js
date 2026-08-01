/* ChadLair GitHub Pages XLSX backend.
 * Static hosting cannot write to files in a GitHub repository. This adapter:
 *  - loads data/chadlair-main.xlsx and data/chadlair-pump.xlsx by default;
 *  - persists changes in IndexedDB;
 *  - writes directly to user-selected XLSX files when File System Access is available;
 *  - exports updated XLSX workbooks in every modern browser.
 */
(function () {
  'use strict';

  const BUILD_ID = 'CHADLAIR-GITHUB-XLSX-20260731-01';
  const TIME_ZONE = 'America/Denver';
  const DEFAULT_MAIN_URL = 'data/chadlair-main.xlsx';
  const DEFAULT_PUMP_URL = 'data/chadlair-pump.xlsx';
  const MEDIA_MANIFEST_URL = 'media/manifest.json';
  const DB_NAME = 'ChadLairStaticXlsx';
  const DB_VERSION = 1;
  const MAX_MEDIA_BYTES = 5 * 1024 * 1024;

  const DIMENSIONS = ['Body', 'Mind', 'Heart', 'Spirit', 'Stewardship'];
  const FREQUENCIES = ['Daily', 'Weekly', 'Monthly', 'Quarterly', 'Annually'];
  const STATUS_CHOICES = ['Default', 'Due', 'Complete', 'Not Completed'];
  const POINTS = { 'To-Do': 1, Daily: 1, Weekly: 5, Monthly: 15, Quarterly: 35, Annually: 120 };

  const ACTIVITY_HEADERS = [
    'Activity ID', 'Dimension', 'Category', 'Subcategory', 'Activity',
    'Frequency', 'Interval', 'Active', 'Default Notes', 'Anchor Period',
    'Deleted', 'Revision', 'Created At', 'Updated At'
  ];
  const TASK_HEADERS = [
    'Task ID', 'Task Type', 'Source Activity ID', 'Occurrence Key',
    'Generation Source', 'Task Name', 'Frequency', 'Period Start',
    'Period End', 'Due Date', 'Dimension', 'Category', 'Subcategory',
    'Points', 'Status Mode', 'Manual Status', 'Effective Status',
    'Previous Status', 'Notes', 'Deleted', 'Revision', 'Created At',
    'Completed At', 'Deleted At', 'Updated At'
  ];
  const ARCHIVE_HEADERS = TASK_HEADERS.concat(['Archived At', 'Archive Reason']);
  const STATS_HEADERS = [
    'Rollup ID', 'Time Basis', 'Date', 'Dimension', 'Category',
    'Subcategory', 'Activity', 'Frequency', 'Task Type', 'Complete Count',
    'Not Completed Count', 'Due Count', 'Earned Points', 'Available Points',
    'Source Revision', 'Updated At'
  ];
  const AFFIRMATION_HEADERS = [
    'Affirmation ID', 'Affirmation', 'Active', 'Deleted', 'Created At', 'Updated At'
  ];
  const TAUNT_HEADERS = [
    'Taunt ID', 'Taunt', 'Active', 'Deleted', 'Created At', 'Updated At'
  ];
  const NAME_HEADERS = ['Name'];
  const SYSTEM_MEDIA_NAMES = new Set([
    'ko.png', 'r1f.gif', 'r2f.gif', 'r3f.gif', 'final.gif',
    '1.gif', '2.gif', '3.gif', '4.gif', '5.gif', '6.gif', '7.gif', '8.gif', '9.gif'
  ]);

  function assert(condition, message) {
    if (!condition) throw new Error(message);
  }

  function clean(value, maxLength, preserveNewlines) {
    let text = String(value == null ? '' : value);
    if (!preserveNewlines) text = text.replace(/\s+/g, ' ');
    text = text.trim();
    return maxLength ? text.slice(0, maxLength) : text;
  }

  function bool(value) {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    return ['true', 'yes', 'y', '1', 'active'].includes(String(value || '').trim().toLowerCase());
  }

  function uuid() {
    if (crypto && typeof crypto.randomUUID === 'function') return crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(16).slice(2)}-${Math.random().toString(16).slice(2)}`;
  }

  function dateKey(value) {
    if (!value) return '';
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, '0')}-${String(value.getUTCDate()).padStart(2, '0')}`;
    }
    if (typeof value === 'number' && window.XLSX && XLSX.SSF) {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
    const text = String(value).trim();
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.valueOf())) {
      return `${parsed.getUTCFullYear()}-${String(parsed.getUTCMonth() + 1).padStart(2, '0')}-${String(parsed.getUTCDate()).padStart(2, '0')}`;
    }
    return text;
  }

  function todayKey() {
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: TIME_ZONE, year: 'numeric', month: '2-digit', day: '2-digit'
    }).formatToParts(new Date());
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return `${values.year}-${values.month}-${values.day}`;
  }

  function utcDate(key) {
    const match = String(key || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    assert(match, `Invalid date: ${key}`);
    return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12));
  }

  function keyFromDate(date) {
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
  }

  function addDays(key, days) {
    const date = utcDate(key);
    date.setUTCDate(date.getUTCDate() + Number(days || 0));
    return keyFromDate(date);
  }

  function addMonths(key, months) {
    const date = utcDate(key);
    return keyFromDate(new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + Number(months || 0), 1, 12)));
  }

  function periodForDate(key, frequency) {
    key = dateKey(key);
    const date = utcDate(key);
    const year = date.getUTCFullYear();
    const month = date.getUTCMonth() + 1;
    if (frequency === 'Daily') return { start: key, end: key };
    if (frequency === 'Weekly') {
      const start = addDays(key, -date.getUTCDay());
      return { start, end: addDays(start, 6) };
    }
    if (frequency === 'Monthly') {
      const start = `${year}-${String(month).padStart(2, '0')}-01`;
      const end = keyFromDate(new Date(Date.UTC(year, month, 0, 12)));
      return { start, end };
    }
    if (frequency === 'Quarterly') {
      const startMonth = Math.floor((month - 1) / 3) * 3 + 1;
      const start = `${year}-${String(startMonth).padStart(2, '0')}-01`;
      const end = keyFromDate(new Date(Date.UTC(year, startMonth + 2, 0, 12)));
      return { start, end };
    }
    if (frequency === 'Annually') return { start: `${year}-01-01`, end: `${year}-12-31` };
    throw new Error(`Unsupported frequency: ${frequency}`);
  }

  function nextPeriodForDate(key, frequency) {
    const current = periodForDate(key, frequency);
    const next = frequency === 'Daily' ? addDays(current.start, 1)
      : frequency === 'Weekly' ? addDays(current.start, 7)
      : frequency === 'Monthly' ? addMonths(current.start, 1)
      : frequency === 'Quarterly' ? addMonths(current.start, 3)
      : addMonths(current.start, 12);
    return periodForDate(next, frequency);
  }

  function periodsBetween(startKey, endKey, frequency) {
    const start = utcDate(startKey);
    const end = utcDate(endKey);
    if (frequency === 'Daily' || frequency === 'Weekly') {
      const days = Math.round((end - start) / 86400000);
      return frequency === 'Weekly' ? days / 7 : days;
    }
    if (frequency === 'Monthly') return (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + (end.getUTCMonth() - start.getUTCMonth());
    if (frequency === 'Quarterly') {
      const a = start.getUTCFullYear() * 4 + Math.floor(start.getUTCMonth() / 3);
      const b = end.getUTCFullYear() * 4 + Math.floor(end.getUTCMonth() / 3);
      return b - a;
    }
    if (frequency === 'Annually') return end.getUTCFullYear() - start.getUTCFullYear();
    return 0;
  }

  function eligible(activity, period) {
    const interval = Math.max(1, Number(activity.Interval || 1));
    const anchor = dateKey(activity['Anchor Period']) || period.start;
    const difference = periodsBetween(anchor, period.start, activity.Frequency);
    return difference >= 0 && difference % interval === 0;
  }

  function occurrenceKey(activity, period) {
    return `${activity['Activity ID'] || ''}::${activity.Frequency || ''}::${period.start}`;
  }

  function taskFromActivity(activity, period, source, now) {
    return {
      'Task ID': uuid(), 'Task Type': 'Activity',
      'Source Activity ID': String(activity['Activity ID'] || ''),
      'Occurrence Key': occurrenceKey(activity, period),
      'Generation Source': source, 'Task Name': String(activity.Activity || ''),
      Frequency: String(activity.Frequency || ''), 'Period Start': period.start,
      'Period End': period.end, 'Due Date': period.end,
      Dimension: String(activity.Dimension || ''), Category: String(activity.Category || ''),
      Subcategory: String(activity.Subcategory || ''), Points: Number(POINTS[activity.Frequency] || 0),
      'Status Mode': 'Default', 'Manual Status': '', 'Effective Status': 'Due',
      'Previous Status': '', Notes: String(activity['Default Notes'] || ''),
      Deleted: false, Revision: 1, 'Created At': now, 'Completed At': '',
      'Deleted At': '', 'Updated At': now
    };
  }

  function autoStatus(task, now) {
    if (task['Task Type'] === 'To-Do') return 'Due';
    const end = dateKey(task['Period End'] || task['Due Date']);
    return !end || now <= end ? 'Due' : 'Not Completed';
  }

  function normalizeRecord(record, headers) {
    const output = {};
    headers.forEach((header) => {
      let value = record && Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
      if (header.includes('Date') || header.endsWith(' At') || header === 'Anchor Period' || header === 'Period Start' || header === 'Period End') {
        value = dateKey(value);
      }
      output[header] = value == null ? '' : value;
    });
    return output;
  }

  function publicActivity(record, row) {
    return {
      id: String(record['Activity ID'] || ''), row: row || 0, revision: Number(record.Revision || 1),
      dimension: String(record.Dimension || ''), category: String(record.Category || ''),
      subcategory: String(record.Subcategory || ''), name: String(record.Activity || ''),
      frequency: String(record.Frequency || ''), interval: Number(record.Interval || 1),
      active: bool(record.Active), notes: String(record['Default Notes'] || ''),
      anchorPeriod: dateKey(record['Anchor Period']), createdAt: dateKey(record['Created At']),
      updatedAt: dateKey(record['Updated At'])
    };
  }

  function publicTask(record, row) {
    const createdAt = dateKey(record['Created At']);
    const completedAt = dateKey(record['Completed At']);
    return {
      id: String(record['Task ID'] || ''), row: row || 0, revision: Number(record.Revision || 1),
      type: String(record['Task Type'] || ''), sourceActivityId: String(record['Source Activity ID'] || ''),
      occurrenceKey: String(record['Occurrence Key'] || ''), generationSource: String(record['Generation Source'] || ''),
      name: String(record['Task Name'] || ''), frequency: String(record.Frequency || ''),
      periodStart: dateKey(record['Period Start']), periodEnd: dateKey(record['Period End']),
      dueDate: dateKey(record['Due Date']), dimension: String(record.Dimension || ''),
      category: String(record.Category || ''), subcategory: String(record.Subcategory || ''),
      points: Number(record.Points || 0), statusMode: String(record['Status Mode'] || 'Default'),
      manualStatus: String(record['Manual Status'] || ''), effectiveStatus: String(record['Effective Status'] || 'Due'),
      previousStatus: String(record['Previous Status'] || ''), notes: String(record.Notes || ''),
      createdAt, createdDate: createdAt, completedAt, completedDate: completedAt,
      updatedAt: dateKey(record['Updated At'])
    };
  }

  function hierarchy(activities, tasks) {
    const seen = new Set();
    const rows = [];
    const add = (dimension, category, subcategory, activity) => {
      const parts = [dimension, category, subcategory, activity].map((value) => String(value || ''));
      if (!parts.some(Boolean)) return;
      const key = parts.join('\u001f').toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      rows.push({ dimension: parts[0], category: parts[1], subcategory: parts[2], activity: parts[3] });
    };
    activities.forEach((item) => add(item.dimension, item.category, item.subcategory, item.name));
    tasks.forEach((item) => add(item.dimension, item.category, item.subcategory, item.type === 'Activity' ? item.name : ''));
    return rows;
  }

  function sortPublicTasks(tasks) {
    const rank = { Due: 0, Complete: 1, 'Not Completed': 2 };
    tasks.sort((a, b) => (rank[a.effectiveStatus] ?? 9) - (rank[b.effectiveStatus] ?? 9)
      || (a.dueDate && b.dueDate ? a.dueDate.localeCompare(b.dueDate) : a.dueDate ? -1 : b.dueDate ? 1 : 0)
      || String(b.createdAt).localeCompare(String(a.createdAt)));
  }

  function roleForMedia(name, mimeType) {
    const lower = String(name || '').toLowerCase();
    if (lower === 'ko.png') return 'ko';
    if (/^[1-9]\.gif$/.test(lower)) return 'win';
    if (lower === 'r1f.gif') return 'round1';
    if (lower === 'r2f.gif') return 'round2';
    if (lower === 'r3f.gif') return 'round3';
    if (lower === 'final.gif') return 'finalRound';
    if (/^[a-z]{6}\.(gif|png|jpe?g|webp|bmp)$/i.test(lower)) return 'loss';
    return mimeType ? 'win' : '';
  }

  function bytesToDataUrl(bytes, mimeType) {
    let binary = '';
    const array = new Uint8Array(bytes);
    const chunk = 0x8000;
    for (let i = 0; i < array.length; i += chunk) binary += String.fromCharCode(...array.subarray(i, i + chunk));
    return `data:${mimeType};base64,${btoa(binary)}`;
  }

  function workbookBytes(workbook) {
    return XLSX.write(workbook, { bookType: 'xlsx', type: 'array', compression: true });
  }

  function readRows(workbook, sheetName, headers) {
    const sheet = workbook && workbook.Sheets[sheetName];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true }).map((row) => normalizeRecord(row, headers));
  }

  function putSheet(workbook, sheetName, rows, headers) {
    const normalized = rows.map((row) => normalizeRecord(row, headers));
    const sheet = XLSX.utils.json_to_sheet(normalized, { header: headers, skipHeader: false });
    sheet['!freeze'] = { xSplit: 0, ySplit: 1 };
    workbook.Sheets[sheetName] = sheet;
    if (!workbook.SheetNames.includes(sheetName)) workbook.SheetNames.push(sheetName);
  }

  function emptyMainWorkbook() {
    const wb = XLSX.utils.book_new();
    putSheet(wb, 'Activity', [], ACTIVITY_HEADERS);
    putSheet(wb, 'Tasks', [], TASK_HEADERS);
    putSheet(wb, 'TaskArchive', [], ARCHIVE_HEADERS);
    putSheet(wb, 'StatsRollup', [], STATS_HEADERS);
    return wb;
  }

  function emptyPumpWorkbook() {
    const wb = XLSX.utils.book_new();
    putSheet(wb, 'Affirmations', [], AFFIRMATION_HEADERS);
    putSheet(wb, 'Taunts', [], TAUNT_HEADERS);
    putSheet(wb, 'Names', [], NAME_HEADERS);
    return wb;
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains('files')) db.createObjectStore('files');
        if (!db.objectStoreNames.contains('handles')) db.createObjectStore('handles');
        if (!db.objectStoreNames.contains('media')) db.createObjectStore('media');
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function idbGet(store, key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const request = db.transaction(store, 'readonly').objectStore(store).get(key);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function idbPut(store, value, key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const request = db.transaction(store, 'readwrite').objectStore(store).put(value, key);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }

  async function idbDelete(store, key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const request = db.transaction(store, 'readwrite').objectStore(store).delete(key);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }

  function download(bytes, filename, mimeType) {
    const blob = new Blob([bytes], { type: mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  class Backend {
    constructor() {
      this.readyPromise = null;
      this.mainWb = null;
      this.pumpWb = null;
      this.main = { Activity: [], Tasks: [], TaskArchive: [], StatsRollup: [] };
      this.pump = { Affirmations: [], Taunts: [], Names: [] };
      this.handles = { main: null, pump: null };
      this.media = [];
      this.deletedRepoMedia = new Set(JSON.parse(localStorage.getItem('ChadLair-deleted-repo-media') || '[]'));
      this.connectedNames = { main: 'chadlair-main.xlsx', pump: 'chadlair-pump.xlsx' };
    }

    async ensureReady() {
      if (!this.readyPromise) this.readyPromise = this.initialize();
      return this.readyPromise;
    }

    async initialize() {
      assert(window.XLSX, 'SheetJS failed to load. Check the browser connection or vendor xlsx.full.min.js locally.');
      this.handles.main = await idbGet('handles', 'main').catch(() => null);
      this.handles.pump = await idbGet('handles', 'pump').catch(() => null);
      this.mainWb = await this.loadWorkbook('main', DEFAULT_MAIN_URL, emptyMainWorkbook);
      this.pumpWb = await this.loadWorkbook('pump', DEFAULT_PUMP_URL, emptyPumpWorkbook);
      this.loadRecords();
      await this.loadMediaManifest();
      const changed = this.reconcileCurrent();
      if (changed) await this.persistMain();
      this.updateConnectionUi();
      return this;
    }

    async loadWorkbook(kind, url, fallbackFactory) {
      const saved = await idbGet('files', kind).catch(() => null);
      if (saved) return XLSX.read(saved, { type: 'array', cellDates: true });
      try {
        const response = await fetch(url, { cache: 'no-store' });
        if (!response.ok) throw new Error(`${response.status}`);
        const bytes = await response.arrayBuffer();
        await idbPut('files', bytes, kind);
        return XLSX.read(bytes, { type: 'array', cellDates: true });
      } catch (error) {
        return fallbackFactory();
      }
    }

    loadRecords() {
      this.main.Activity = readRows(this.mainWb, 'Activity', ACTIVITY_HEADERS);
      this.main.Tasks = readRows(this.mainWb, 'Tasks', TASK_HEADERS);
      this.main.TaskArchive = readRows(this.mainWb, 'TaskArchive', ARCHIVE_HEADERS);
      this.main.StatsRollup = readRows(this.mainWb, 'StatsRollup', STATS_HEADERS);
      this.pump.Affirmations = readRows(this.pumpWb, 'Affirmations', AFFIRMATION_HEADERS);
      this.pump.Taunts = readRows(this.pumpWb, 'Taunts', TAUNT_HEADERS);
      this.pump.Names = readRows(this.pumpWb, 'Names', NAME_HEADERS);
    }

    async loadMediaManifest() {
      try {
        const response = await fetch(MEDIA_MANIFEST_URL, { cache: 'no-store' });
        if (!response.ok) throw new Error('No media manifest');
        const manifest = await response.json();
        this.media = (Array.isArray(manifest) ? manifest : manifest.media || []).map((item) => {
          const name = clean(item.name, 180) || clean(item.path, 180).split('/').pop();
          const mimeType = clean(item.mimeType, 100) || (name.endsWith('.gif') ? 'image/gif' : name.endsWith('.png') ? 'image/png' : 'image/jpeg');
          return { id: `repo:${name}`, name, mimeType, bytes: Number(item.bytes || 0), role: item.role || roleForMedia(name, mimeType), path: item.path || `media/${name}`, source: 'repo' };
        }).filter((item) => !this.deletedRepoMedia.has(item.id));
      } catch (error) {
        this.media = [];
      }
      const localMeta = JSON.parse(localStorage.getItem('ChadLair-local-media') || '[]');
      this.media.push(...localMeta.map((item) => Object.assign({}, item, { source: 'local' })));
    }

    reconcileCurrent() {
      const now = todayKey();
      let changed = false;
      const survivors = [];
      for (const task of this.main.Tasks) {
        const deleted = bool(task.Deleted);
        const isTodo = task['Task Type'] === 'To-Do';
        const end = dateKey(task['Period End'] || task['Due Date']);
        const shouldArchive = isTodo ? deleted : Boolean(end && end < now && !(task['Status Mode'] === 'Manual' && task['Effective Status'] === 'Due'));
        if (shouldArchive) {
          if (!deleted && (task['Status Mode'] || 'Default') === 'Default' && task['Effective Status'] !== 'Complete') {
            task['Previous Status'] = task['Effective Status'] || 'Due';
            task['Effective Status'] = 'Not Completed';
            task.Revision = Number(task.Revision || 0) + 1;
            task['Updated At'] = now;
          }
          if (!this.main.TaskArchive.some((row) => String(row['Task ID']) === String(task['Task ID']))) {
            this.main.TaskArchive.push(Object.assign({}, task, {
              'Archived At': now, 'Archive Reason': deleted ? 'Deleted' : 'Period Closed'
            }));
          }
          changed = true;
          continue;
        }
        if (!deleted && (task['Status Mode'] || 'Default') === 'Default') {
          const automatic = autoStatus(task, now);
          if (task['Effective Status'] !== automatic) {
            task['Previous Status'] = task['Effective Status'] || '';
            task['Effective Status'] = automatic;
            if (automatic !== 'Complete') task['Completed At'] = '';
            task.Revision = Number(task.Revision || 0) + 1;
            task['Updated At'] = now;
            changed = true;
          }
        }
        survivors.push(task);
      }
      this.main.Tasks = survivors;
      const keys = new Set(survivors.map((task) => String(task['Occurrence Key'] || '')).filter(Boolean));
      for (const activity of this.main.Activity) {
        if (bool(activity.Deleted) || !bool(activity.Active)) continue;
        const period = periodForDate(now, activity.Frequency);
        if (!eligible(activity, period)) continue;
        const key = occurrenceKey(activity, period);
        if (keys.has(key)) continue;
        this.main.Tasks.push(taskFromActivity(activity, period, 'Automatic', now));
        keys.add(key);
        changed = true;
      }
      return changed;
    }

    writeMainSheets() {
      putSheet(this.mainWb, 'Activity', this.main.Activity, ACTIVITY_HEADERS);
      putSheet(this.mainWb, 'Tasks', this.main.Tasks, TASK_HEADERS);
      putSheet(this.mainWb, 'TaskArchive', this.main.TaskArchive, ARCHIVE_HEADERS);
      putSheet(this.mainWb, 'StatsRollup', this.main.StatsRollup, STATS_HEADERS);
    }

    writePumpSheets() {
      putSheet(this.pumpWb, 'Affirmations', this.pump.Affirmations, AFFIRMATION_HEADERS);
      putSheet(this.pumpWb, 'Taunts', this.pump.Taunts, TAUNT_HEADERS);
      putSheet(this.pumpWb, 'Names', this.pump.Names, NAME_HEADERS);
    }

    async persistMain() {
      this.writeMainSheets();
      const bytes = workbookBytes(this.mainWb);
      await idbPut('files', bytes, 'main');
      await this.writeHandle('main', bytes);
      this.updateConnectionUi();
    }

    async persistPump() {
      this.writePumpSheets();
      const bytes = workbookBytes(this.pumpWb);
      await idbPut('files', bytes, 'pump');
      await this.writeHandle('pump', bytes);
      this.updateConnectionUi();
    }

    async writeHandle(kind, bytes) {
      const handle = this.handles[kind];
      if (!handle || typeof handle.createWritable !== 'function') return false;
      try {
        const permission = typeof handle.queryPermission === 'function' ? await handle.queryPermission({ mode: 'readwrite' }) : 'granted';
        if (permission !== 'granted') return false;
        const writable = await handle.createWritable();
        await writable.write(bytes);
        await writable.close();
        return true;
      } catch (error) {
        return false;
      }
    }

    buildFullState(message) {
      const activities = this.main.Activity.filter((row) => !bool(row.Deleted)).map((row, index) => publicActivity(row, index + 2));
      const tasks = this.main.Tasks.filter((row) => !bool(row.Deleted)).map((row, index) => publicTask(row, index + 2));
      activities.sort((a, b) => a.dimension.localeCompare(b.dimension) || a.category.localeCompare(b.category) || a.subcategory.localeCompare(b.subcategory) || a.name.localeCompare(b.name));
      sortPublicTasks(tasks);
      return {
        ok: true, mode: 'full', message: message || '', appName: 'ChadLair', buildId: BUILD_ID,
        spreadsheetId: this.connectedNames.main, timeZone: TIME_ZONE, today: todayKey(),
        sourceShape: { activityRows: this.main.Activity.length, taskRows: this.main.Tasks.length },
        dimensions: DIMENSIONS.slice(), frequencies: FREQUENCIES.slice(), statusChoices: STATUS_CHOICES.slice(),
        points: Object.assign({}, POINTS), lossEvents: [], activities, tasks, hierarchy: hierarchy(activities, tasks)
      };
    }

    patch(message, options) {
      const changes = options || {};
      return {
        ok: true, mode: 'patch', message: message || '', buildId: BUILD_ID,
        spreadsheetId: this.connectedNames.main, today: todayKey(),
        changes: {
          activities: (changes.activities || []).map((record) => publicActivity(record, this.main.Activity.indexOf(record) + 2)),
          tasks: (changes.tasks || []).map((record) => publicTask(record, this.main.Tasks.indexOf(record) + 2)),
          removeActivityIds: (changes.removeActivityIds || []).slice(),
          removeTaskIds: (changes.removeTaskIds || []).slice(),
          hierarchyChanged: Boolean(changes.hierarchyChanged), statsChanged: Boolean(changes.statsChanged),
          lossEvents: (changes.lossEvents || []).slice()
        }
      };
    }

    findActivity(id) { return this.main.Activity.find((row) => String(row['Activity ID']) === String(id)); }
    findTask(id) { return this.main.Tasks.find((row) => String(row['Task ID']) === String(id)); }

    async call(method, ...args) {
      await this.ensureReady();
      assert(typeof this[method] === 'function' && !method.startsWith('_'), `Unsupported static method: ${method}`);
      return this[method](...args);
    }

    async callForm(method, formElement) {
      await this.ensureReady();
      if (method !== 'savePumpMedia') throw new Error(`Unsupported form method: ${method}`);
      const data = new FormData(formElement);
      return this.savePumpMedia({
        mediaFile: data.get('mediaFile'), mediaName: data.get('mediaName'), mediaRole: data.get('mediaRole')
      });
    }

    async bootstrapApp() {
      const changed = this.reconcileCurrent();
      if (changed) await this.persistMain();
      return this.buildFullState('');
    }

    async saveActivity(payload) {
      const now = todayKey();
      const dimension = clean(payload.dimension, 50);
      const frequency = clean(payload.frequency, 30);
      const name = clean(payload.name, 200);
      const interval = Number(payload.interval);
      assert(DIMENSIONS.includes(dimension), 'Choose a valid Dimension.');
      assert(name, 'Activity is required.');
      assert(FREQUENCIES.includes(frequency), 'Choose a valid Frequency.');
      assert(Number.isInteger(interval) && interval >= 1 && interval <= 100, 'Choose an Interval from 1 through 100.');
      let record = this.findActivity(payload.id);
      const isNew = Boolean(payload.isNew || !record);
      let message;
      if (isNew) {
        const period = periodForDate(now, frequency);
        record = {
          'Activity ID': clean(payload.id, 100) || uuid(), Dimension: dimension,
          Category: clean(payload.category, 120), Subcategory: clean(payload.subcategory, 120),
          Activity: name, Frequency: frequency, Interval: interval, Active: bool(payload.active),
          'Default Notes': clean(payload.notes, 10000, true), 'Anchor Period': period.start,
          Deleted: false, Revision: 1, 'Created At': now, 'Updated At': now
        };
        this.main.Activity.push(record);
        message = record.Active ? 'Activity created.' : 'Inactive Activity created.';
      } else {
        const frequencyChanged = record.Frequency !== frequency;
        Object.assign(record, {
          Dimension: dimension, Category: clean(payload.category, 120),
          Subcategory: clean(payload.subcategory, 120), Activity: name, Frequency: frequency,
          Interval: interval, Active: bool(payload.active), 'Default Notes': clean(payload.notes, 10000, true),
          Revision: Number(record.Revision || 0) + 1, 'Updated At': now
        });
        if (frequencyChanged) record['Anchor Period'] = nextPeriodForDate(now, frequency).start;
        message = 'Activity updated. Existing Tasks were left unchanged.';
      }
      let generated = null;
      if (!bool(record.Deleted) && bool(record.Active)) {
        const period = periodForDate(now, record.Frequency);
        const key = occurrenceKey(record, period);
        if (eligible(record, period) && !this.main.Tasks.some((task) => String(task['Occurrence Key']) === key)) {
          generated = taskFromActivity(record, period, 'Automatic', now);
          this.main.Tasks.push(generated);
          message = isNew ? 'Activity created with its current-period Task.' : 'Activity updated and its current-period Task was created.';
        }
      }
      await this.persistMain();
      return this.patch(message, { activities: [record], tasks: generated ? [generated] : [], hierarchyChanged: true, statsChanged: Boolean(generated) });
    }

    async deleteActivity(id) {
      const record = this.findActivity(id);
      assert(record && !bool(record.Deleted), 'Activity not found.');
      record.Active = false; record.Deleted = true; record.Revision = Number(record.Revision || 0) + 1; record['Updated At'] = todayKey();
      await this.persistMain();
      return this.patch('Activity archived. Existing Task history was preserved.', { removeActivityIds: [String(id)], hierarchyChanged: true });
    }

    async renewCreateActivity(id) {
      const activity = this.findActivity(id);
      assert(activity && !bool(activity.Deleted), 'Activity not found.');
      assert(bool(activity.Active), 'Make this Activity active before using Renew/Create.');
      const now = todayKey();
      const period = periodForDate(now, activity.Frequency);
      const key = occurrenceKey(activity, period);
      let task = this.main.Tasks.find((row) => String(row['Occurrence Key']) === key);
      let message;
      if (!task) {
        task = taskFromActivity(activity, period, 'Renew/Create', now);
        this.main.Tasks.push(task);
        message = 'Current-period Task created as Due.';
      } else {
        const prior = task['Effective Status'] || 'Due';
        Object.assign(task, { Deleted: false, 'Deleted At': '', 'Previous Status': prior,
          'Status Mode': 'Default', 'Manual Status': '', 'Effective Status': 'Due',
          'Completed At': '', 'Generation Source': 'Renew/Create', Revision: Number(task.Revision || 0) + 1,
          'Updated At': now });
        message = prior === 'Complete' ? 'Completed Task reopened as Due.' : prior === 'Deleted' ? 'Deleted Task restored as Due.' : 'Current-period Task renewed as Due.';
      }
      await this.persistMain();
      return this.patch(message, { tasks: [task], statsChanged: true });
    }

    async saveTodo(payload) {
      const now = todayKey();
      const name = clean(payload.name, 200);
      const dimension = clean(payload.dimension, 50);
      assert(name, 'To-Do is required.');
      assert(!dimension || DIMENSIONS.includes(dimension), 'Choose a valid Dimension or leave it blank.');
      let task = this.findTask(payload.id);
      const isNew = Boolean(payload.isNew || !task);
      let message;
      if (isNew) {
        task = {
          'Task ID': clean(payload.id, 100) || uuid(), 'Task Type': 'To-Do', 'Source Activity ID': '',
          'Occurrence Key': '', 'Generation Source': 'To-Do', 'Task Name': name, Frequency: 'To-Do',
          'Period Start': '', 'Period End': '', 'Due Date': '', Dimension: dimension,
          Category: clean(payload.category, 120), Subcategory: clean(payload.subcategory, 120),
          Points: POINTS['To-Do'], 'Status Mode': 'Default', 'Manual Status': '',
          'Effective Status': 'Due', 'Previous Status': '', Notes: clean(payload.notes, 10000, true),
          Deleted: false, Revision: 1, 'Created At': now, 'Completed At': '', 'Deleted At': '', 'Updated At': now
        };
        this.main.Tasks.push(task); message = 'To-Do created.';
      } else {
        assert(task['Task Type'] === 'To-Do' && !bool(task.Deleted), 'To-Do not found.');
        Object.assign(task, { 'Task Name': name, Dimension: dimension, Category: clean(payload.category, 120),
          Subcategory: clean(payload.subcategory, 120), Notes: clean(payload.notes, 10000, true),
          Revision: Number(task.Revision || 0) + 1, 'Updated At': now });
        message = 'To-Do updated.';
      }
      await this.persistMain();
      return this.patch(message, { tasks: [task], hierarchyChanged: true, statsChanged: true });
    }

    async updateTaskStatus(id, choice) {
      assert(STATUS_CHOICES.includes(choice), 'Invalid status.');
      const task = this.findTask(id); assert(task && !bool(task.Deleted), 'Task not found.');
      const prior = task['Effective Status'] || autoStatus(task, todayKey());
      task['Previous Status'] = prior;
      if (choice === 'Default') {
        task['Status Mode'] = 'Default'; task['Manual Status'] = ''; task['Effective Status'] = autoStatus(task, todayKey());
      } else {
        task['Status Mode'] = 'Manual'; task['Manual Status'] = choice; task['Effective Status'] = choice;
      }
      task['Completed At'] = task['Effective Status'] === 'Complete' ? (dateKey(task['Completed At']) || todayKey()) : '';
      task.Revision = Number(task.Revision || 0) + 1; task['Updated At'] = todayKey();
      const lossEvents = ((prior === 'Complete' && task['Effective Status'] !== 'Complete') || (prior !== 'Not Completed' && task['Effective Status'] === 'Not Completed'))
        ? [{ id: String(task['Task ID']), name: String(task['Task Name']), revision: Number(task.Revision), previousStatus: prior, effectiveStatus: task['Effective Status'] }] : [];
      await this.persistMain();
      return this.patch('Task status saved.', { tasks: [task], statsChanged: true, lossEvents });
    }

    async updateTaskNotes(id, notes) {
      const task = this.findTask(id); assert(task && !bool(task.Deleted), 'Task not found.');
      task.Notes = clean(notes, 10000, true); task.Revision = Number(task.Revision || 0) + 1; task['Updated At'] = todayKey();
      await this.persistMain();
      return this.patch('Notes saved.', { tasks: [task] });
    }

    async deleteTask(id) {
      const task = this.findTask(id); assert(task && !bool(task.Deleted), 'Task not found.');
      task['Previous Status'] = task['Effective Status'] || 'Due'; task.Deleted = true;
      task['Effective Status'] = 'Deleted'; task['Deleted At'] = todayKey();
      task.Revision = Number(task.Revision || 0) + 1; task['Updated At'] = task['Deleted At'];
      await this.persistMain();
      return this.patch('Task deleted and hidden.', { removeTaskIds: [String(id)], hierarchyChanged: true, statsChanged: true });
    }

    buildStatsRows() {
      const map = new Map();
      const add = (task, basis, date, status) => {
        const activity = task['Task Type'] === 'Activity' ? String(task['Task Name'] || '') : '';
        const parts = [basis, date, task.Dimension || '', task.Category || '', task.Subcategory || '', activity, task.Frequency || '', task['Task Type'] || ''];
        const key = parts.join('\u001f');
        if (!map.has(key)) map.set(key, {
          'Rollup ID': key, 'Time Basis': basis, Date: date, Dimension: String(task.Dimension || ''),
          Category: String(task.Category || ''), Subcategory: String(task.Subcategory || ''), Activity: activity,
          Frequency: String(task.Frequency || ''), 'Task Type': String(task['Task Type'] || ''),
          'Complete Count': 0, 'Not Completed Count': 0, 'Due Count': 0,
          'Earned Points': 0, 'Available Points': 0, 'Source Revision': 0, 'Updated At': todayKey()
        });
        const row = map.get(key); const points = Number(task.Points || 0);
        if (status === 'Complete') { row['Complete Count'] += 1; row['Earned Points'] += points; }
        else if (status === 'Not Completed') row['Not Completed Count'] += 1;
        else if (status === 'Due') { row['Due Count'] += 1; row['Available Points'] += points; }
        row['Source Revision'] += Number(task.Revision || 1);
      };
      this.main.TaskArchive.filter((task) => !bool(task.Deleted)).forEach((task) => {
        const status = String(task['Effective Status'] || '');
        const due = dateKey(task['Due Date'] || task['Created At']);
        if (due) add(task, 'Due', due, status);
        const completed = dateKey(task['Completed At']);
        if (status === 'Complete' && completed) add(task, 'Completion', completed, status);
      });
      return Array.from(map.values()).sort((a, b) => String(a.Date).localeCompare(String(b.Date)));
    }

    async getStatsSnapshot() {
      this.main.StatsRollup = this.buildStatsRows();
      await this.persistMain();
      const rows = this.main.StatsRollup.map((row) => ({
        id: String(row['Rollup ID'] || ''), basis: String(row['Time Basis']).toLowerCase() === 'completion' ? 'completion' : 'due',
        date: dateKey(row.Date), dimension: String(row.Dimension || ''), category: String(row.Category || ''),
        subcategory: String(row.Subcategory || ''), activity: String(row.Activity || ''), frequency: String(row.Frequency || ''),
        taskType: String(row['Task Type'] || ''), complete: Number(row['Complete Count'] || 0),
        notCompleted: Number(row['Not Completed Count'] || 0), due: Number(row['Due Count'] || 0),
        earned: Number(row['Earned Points'] || 0), available: Number(row['Available Points'] || 0)
      }));
      const rowsHierarchy = hierarchy([], rows.map((row) => ({
        dimension: row.dimension, category: row.category, subcategory: row.subcategory,
        type: row.taskType, name: row.activity
      })));
      return { ok: true, rows, hierarchy: rowsHierarchy, updatedAt: todayKey() };
    }

    async getPumpSnapshot() {
      return {
        ok: true, configured: true, media: this.media.map(({ path, source, ...item }) => item),
        affirmations: this.pump.Affirmations.filter((row) => bool(row.Active) && !bool(row.Deleted)).map((row) => ({ id: String(row['Affirmation ID']), text: String(row.Affirmation || '') })),
        taunts: this.pump.Taunts.filter((row) => bool(row.Active) && !bool(row.Deleted)).map((row) => ({ id: String(row['Taunt ID']), text: String(row.Taunt || '') })),
        names: this.pump.Names.map((row) => clean(row.Name, 100)).filter(Boolean),
        maxUploadBytes: MAX_MEDIA_BYTES, message: ''
      };
    }

    async getPumpMediaAsset(id) {
      const item = this.media.find((row) => String(row.id) === String(id));
      assert(item, 'Media file is unavailable.');
      let bytes;
      if (item.source === 'local') {
        const blob = await idbGet('media', item.id);
        assert(blob, 'Media file is unavailable.');
        bytes = await blob.arrayBuffer();
      } else {
        const response = await fetch(item.path, { cache: 'no-store' });
        assert(response.ok, `Could not load ${item.name}.`);
        bytes = await response.arrayBuffer();
      }
      assert(bytes.byteLength <= MAX_MEDIA_BYTES, 'Media file is too large to preview.');
      return { ok: true, id: item.id, name: item.name, mimeType: item.mimeType, dataUrl: bytesToDataUrl(bytes, item.mimeType) };
    }

    async savePumpMedia(formObject) {
      const file = formObject && formObject.mediaFile;
      assert(file instanceof Blob && file.size, 'Choose an image or GIF first.');
      assert(file.size <= MAX_MEDIA_BYTES, 'Media must be 5 MB or smaller.');
      assert(/^image\/(gif|png|jpeg|webp|bmp)$/i.test(file.type), 'Use a GIF, PNG, JPG, WEBP, or BMP image.');
      const requested = clean(formObject.mediaName || file.name, 180);
      const system = SYSTEM_MEDIA_NAMES.has(requested.toLowerCase()) ? requested.toLowerCase() : '';
      const extension = file.type === 'image/jpeg' ? 'jpg' : file.type.split('/')[1];
      const role = clean(formObject.mediaRole, 20).toLowerCase() === 'loss' ? 'loss' : 'win';
      const random = role === 'loss' ? Math.random().toString(36).replace(/[^a-z]/g, '').slice(2, 8).padEnd(6, 'x').toUpperCase()
        : Math.random().toString(36).replace(/[^a-z0-9]/g, '').slice(2, 7).padEnd(5, 'x').toUpperCase();
      const name = system || `${random}.${extension}`;
      const id = `local:${uuid()}`;
      const media = { id, name, mimeType: file.type, bytes: file.size, role: roleForMedia(name, file.type) || role, source: 'local' };
      await idbPut('media', file, id);
      this.media.push(media);
      localStorage.setItem('ChadLair-local-media', JSON.stringify(this.media.filter((item) => item.source === 'local').map(({ source, path, ...item }) => item)));
      return { ok: true, media: (({ source, path, ...item }) => item)(media), message: `${name} uploaded to browser storage.` };
    }

    async savePumpMediaData(payload) {
      const mimeType = clean(payload && payload.mimeType, 100).toLowerCase();
      assert(/^image\/(gif|png|jpeg|webp|bmp)$/i.test(mimeType), 'Use a GIF, PNG, JPG, WEBP, or BMP image.');
      const binary = atob(String(payload.base64 || '').replace(/\s/g, ''));
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
      const blob = new Blob([bytes], { type: mimeType });
      Object.defineProperty(blob, 'name', { value: clean(payload.name, 180) || 'pasted-image' });
      return this.savePumpMedia({ mediaFile: blob, mediaName: payload.name, mediaRole: payload.mediaRole });
    }

    async deletePumpMedia(id) {
      const index = this.media.findIndex((item) => String(item.id) === String(id));
      assert(index >= 0, 'Media file is unavailable.');
      const item = this.media[index];
      this.media.splice(index, 1);
      if (item.source === 'local') {
        await idbDelete('media', item.id);
        localStorage.setItem('ChadLair-local-media', JSON.stringify(this.media.filter((row) => row.source === 'local').map(({ source, path, ...row }) => row)));
      } else {
        this.deletedRepoMedia.add(item.id);
        localStorage.setItem('ChadLair-deleted-repo-media', JSON.stringify(Array.from(this.deletedRepoMedia)));
      }
      return { ok: true, id, message: `${item.name} hidden from this browser.` };
    }

    async savePumpAffirmation(text) {
      text = clean(text, 500, true); assert(text, 'Write an affirmation first.');
      const row = { 'Affirmation ID': uuid(), Affirmation: text, Active: true, Deleted: false, 'Created At': todayKey(), 'Updated At': todayKey() };
      this.pump.Affirmations.push(row); await this.persistPump();
      return { ok: true, affirmation: { id: row['Affirmation ID'], text }, message: 'Affirmation added.' };
    }

    async savePumpTaunt(text) {
      text = clean(text, 500, true); assert(text, 'Write a taunt first.');
      const row = { 'Taunt ID': uuid(), Taunt: text, Active: true, Deleted: false, 'Created At': todayKey(), 'Updated At': todayKey() };
      this.pump.Taunts.push(row); await this.persistPump();
      return { ok: true, taunt: { id: row['Taunt ID'], text }, message: 'Taunt added.' };
    }

    async connectFiles() {
      await this.ensureReady();
      let selections = [];
      if (window.showOpenFilePicker) {
        selections = await window.showOpenFilePicker({
          multiple: true,
          types: [{ description: 'Excel workbooks', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
        });
        for (const handle of selections) {
          const file = await handle.getFile();
          await this.acceptWorkbookFile(file, handle);
        }
      } else {
        const files = await new Promise((resolve) => {
          const input = document.createElement('input'); input.type = 'file'; input.accept = '.xlsx'; input.multiple = true;
          input.onchange = () => resolve(Array.from(input.files || [])); input.click();
        });
        for (const file of files) await this.acceptWorkbookFile(file, null);
      }
      this.loadRecords();
      const changed = this.reconcileCurrent();
      if (changed) await this.persistMain();
      this.updateConnectionUi();
      return this.buildFullState('XLSX workbooks connected.');
    }

    async acceptWorkbookFile(file, handle) {
      const bytes = await file.arrayBuffer();
      const wb = XLSX.read(bytes, { type: 'array', cellDates: true });
      const isMain = ['Activity', 'Tasks'].every((name) => wb.SheetNames.includes(name));
      const isPump = ['Affirmations', 'Taunts', 'Names'].some((name) => wb.SheetNames.includes(name));
      assert(isMain || isPump, `${file.name} is not a ChadLair main or PUMP workbook.`);
      const kind = isMain ? 'main' : 'pump';
      if (kind === 'main') this.mainWb = wb; else this.pumpWb = wb;
      this.connectedNames[kind] = file.name;
      this.handles[kind] = handle;
      await idbPut('files', bytes, kind);
      if (handle) {
        if (typeof handle.requestPermission === 'function') {
          await handle.requestPermission({ mode: 'readwrite' }).catch(() => 'prompt');
        }
        await idbPut('handles', handle, kind).catch(() => null);
      }
    }

    async flush() {
      await this.ensureReady();
      await this.persistMain(); await this.persistPump();
      return { ok: true, message: 'XLSX changes saved locally.' };
    }

    async exportWorkbooks() {
      await this.ensureReady();
      this.writeMainSheets(); this.writePumpSheets();
      download(workbookBytes(this.mainWb), this.connectedNames.main || 'chadlair-main.xlsx');
      setTimeout(() => download(workbookBytes(this.pumpWb), this.connectedNames.pump || 'chadlair-pump.xlsx'), 250);
    }

    updateConnectionUi() {
      const button = document.getElementById('xlsx-connect-button');
      if (button) {
        button.title = `Main: ${this.connectedNames.main}\nPUMP: ${this.connectedNames.pump}`;
        button.textContent = this.handles.main || this.handles.pump ? 'XLSX Connected' : 'Connect XLSX';
      }
    }
  }

  const backend = new Backend();
  window.ChadLairXlsxBackend = {
    call: (method, ...args) => backend.call(method, ...args),
    callForm: (method, form) => backend.callForm(method, form),
    connectFiles: () => backend.connectFiles(),
    flush: () => backend.flush(),
    exportWorkbooks: () => backend.exportWorkbooks(),
    backend
  };

  window.addEventListener('DOMContentLoaded', () => {
    const connect = document.getElementById('xlsx-connect-button');
    const save = document.getElementById('xlsx-save-button');
    const exportButton = document.getElementById('xlsx-export-button');
    if (connect) connect.addEventListener('click', async () => {
      try {
        const state = await backend.connectFiles();
        if (typeof window.ChadLairApplyState === 'function') window.ChadLairApplyState(state, { silent: false, quiet: false });
      } catch (error) {
        if (error && error.name !== 'AbortError') {
          if (typeof window.ChadLairHandleError === 'function') window.ChadLairHandleError(error); else alert(error.message || error);
        }
      }
    });
    if (save) save.addEventListener('click', async () => {
      try {
        const result = await backend.flush();
        if (typeof window.ChadLairShowToast === 'function') window.ChadLairShowToast(result.message);
      } catch (error) {
        if (typeof window.ChadLairHandleError === 'function') window.ChadLairHandleError(error); else alert(error.message || error);
      }
    });
    if (exportButton) exportButton.addEventListener('click', () => backend.exportWorkbooks());
    backend.ensureReady().catch((error) => {
      if (typeof window.ChadLairHandleError === 'function') window.ChadLairHandleError(error); else console.error(error);
    });
  });
})();
