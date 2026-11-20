/**
 * 定期試験日程 検索・管理データベース (script.js)
 * - シート数1固定・行数可変（毎年再利用可能）のExcelファイルをブラウザ上で直接解析
 * - SheetJSマルチCDNフォールバック搭載
 * - PC向けテーブル表示 / スマホモード（カード表示・折りたたみ絞り込み）切替対応
 * - IndexedDB（フォールバック: localStorage）によるローカルデータベース永続化
 */

(() => {
  'use strict';

  // =========================================================================
  // 0. Dynamic SheetJS Loader (Fallback if primary CDN is blocked/slow)
  // =========================================================================
  const XLSX_CDNS = [
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
  ];

  function ensureXlsxLoaded() {
    if (typeof window.XLSX !== 'undefined') {
      return Promise.resolve(window.XLSX);
    }
    return new Promise((resolve, reject) => {
      let idx = 0;
      const tryNext = () => {
        if (typeof window.XLSX !== 'undefined') {
          resolve(window.XLSX);
          return;
        }
        if (idx >= XLSX_CDNS.length) {
          reject(new Error('Excel解析ライブラリ(SheetJS)の読み込みに失敗しました。ネットワーク接続をご確認ください。'));
          return;
        }
        const script = document.createElement('script');
        script.src = XLSX_CDNS[idx++];
        script.async = true;
        script.onload = () => {
          if (typeof window.XLSX !== 'undefined') {
            resolve(window.XLSX);
          } else {
            tryNext();
          }
        };
        script.onerror = () => tryNext();
        document.head.appendChild(script);
      };
      tryNext();
    });
  }

  // =========================================================================
  // 1. Storage Layer (IndexedDB with localStorage fallback)
  // =========================================================================
  const DB_NAME = 'UniversityExamScheduleDB_v1';
  const DB_VERSION = 1;
  const STORE_NAME = 'appState';
  const MOBILE_PREF_KEY = `${DB_NAME}_mobileMode`;

  const StorageService = {
    openDB() {
      return new Promise((resolve, reject) => {
        if (!('indexedDB' in window)) {
          reject(new Error('IndexedDB not supported'));
          return;
        }
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = (event) => {
          const db = event.target.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) {
            db.createObjectStore(STORE_NAME);
          }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    },

    async get(key) {
      try {
        const db = await this.openDB();
        return await new Promise((resolve, reject) => {
          const tx = db.transaction(STORE_NAME, 'readonly');
          const store = tx.objectStore(STORE_NAME);
          const req = store.get(key);
          req.onsuccess = () => resolve(req.result ?? null);
          req.onerror = () => reject(req.error);
        });
      } catch {
        const raw = localStorage.getItem(`${DB_NAME}_${key}`);
        return raw ? JSON.parse(raw) : null;
      }
    },

    async set(key, value) {
      try {
        const db = await this.openDB();
        await new Promise((resolve, reject) => {
          const tx = db.transaction(STORE_NAME, 'readwrite');
          const store = tx.objectStore(STORE_NAME);
          const req = store.put(value, key);
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch {
        localStorage.setItem(`${DB_NAME}_${key}`, JSON.stringify(value));
      }
    },

    async clearAll() {
      try {
        const db = await this.openDB();
        await new Promise((resolve, reject) => {
          const tx = db.transaction(STORE_NAME, 'readwrite');
          const store = tx.objectStore(STORE_NAME);
          const req = store.clear();
          req.onsuccess = () => resolve();
          req.onerror = () => reject(req.error);
        });
      } catch {
        localStorage.removeItem(`${DB_NAME}_records`);
        localStorage.removeItem(`${DB_NAME}_meta`);
        localStorage.removeItem(`${DB_NAME}_pinnedKeys`);
      }
    }
  };

  // =========================================================================
  // 2. Normalization & Parsing Utilities
  // =========================================================================
  const DAY_ORDER = { '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6, '日': 7 };

  function normalizeText(value) {
    if (value === null || value === undefined) return '';
    return String(value)
      .normalize('NFKC')
      .toLowerCase()
      .replace(/[\s　]+/g, '');
  }

  function cleanCell(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim();
  }

  function normalizePeriodDisplay(rawPeriod) {
    const str = cleanCell(rawPeriod);
    if (!str) return '';
    return str.normalize('NFKC').replace(/[･・]/g, '・');
  }

  function parsePeriodNumbers(periodStr) {
    const norm = normalizePeriodDisplay(periodStr);
    const matches = norm.match(/\d+/g);
    if (!matches) return [];
    return matches.map((n) => parseInt(n, 10));
  }

  /**
   * 試験日時文字列を構造化データへ変換
   * 例: "12月21日(月)3時限", "1月12日(火)1・2時限"
   * 4月〜12月を年度前半(month)、1月〜3月を年度後半(month + 12)として時系列ソートキーを算出
   */
  function parseExamDateTime(rawDateTime) {
    const text = cleanCell(rawDateTime);
    if (!text) {
      return {
        raw: '',
        dateLabel: '日時未定・記載なし',
        periodLabel: '',
        periods: [],
        sortKey: Number.MAX_SAFE_INTEGER,
        dateSortKey: Number.MAX_SAFE_INTEGER,
        month: null,
        day: null,
        weekday: ''
      };
    }

    const norm = text.normalize('NFKC').replace(/[･・]/g, '・');
    const dateMatch = norm.match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日(?:\s*[（(]\s*([月火水木金土日])\s*[）)])?/);
    const periodMatch = norm.match(/(\d+(?:・\d+)*)\s*時限/);

    if (!dateMatch) {
      return {
        raw: text,
        dateLabel: text,
        periodLabel: '',
        periods: [],
        sortKey: Number.MAX_SAFE_INTEGER - 1,
        dateSortKey: Number.MAX_SAFE_INTEGER - 1,
        month: null,
        day: null,
        weekday: ''
      };
    }

    const month = parseInt(dateMatch[1], 10);
    const day = parseInt(dateMatch[2], 10);
    const weekday = dateMatch[3] || '';
    const dateLabel = `${month}月${day}日${weekday ? `(${weekday})` : ''}`;

    const periodLabel = periodMatch ? `${periodMatch[1]}時限` : '';
    const periods = periodMatch ? periodMatch[1].split('・').map((n) => parseInt(n, 10)) : [];
    const firstPeriod = periods.length > 0 ? periods[0] : 99;

    const academicMonth = month >= 4 ? month : month + 12;
    const dateSortKey = academicMonth * 100 + day;
    const sortKey = dateSortKey * 100 + firstPeriod;

    return {
      raw: text,
      dateLabel,
      periodLabel,
      periods,
      sortKey,
      dateSortKey,
      month,
      day,
      weekday
    };
  }

  /**
   * 2次元配列（行数の制約なし）から試験レコード配列を構築
   * - 左端に空列がある場合も自動オフセット検出
   * - タイトル行・見出し行・空行を自動スキップ
   */
  function parseRowsToRecords(rows) {
    const records = [];

    let colOffset = 0;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r];
      if (!Array.isArray(row)) continue;
      const nonEmptyIndices = row
        .map((v, idx) => (cleanCell(v) !== '' ? idx : -1))
        .filter((idx) => idx >= 0);
      if (nonEmptyIndices.length >= 6) {
        colOffset = nonEmptyIndices[0];
        break;
      }
    }

    rows.forEach((row, index) => {
      if (!Array.isArray(row)) return;

      const c0 = cleanCell(row[colOffset + 0]);
      const c1 = cleanCell(row[colOffset + 1]);
      const c2 = cleanCell(row[colOffset + 2]);
      const c3 = cleanCell(row[colOffset + 3]);
      const c4 = normalizePeriodDisplay(row[colOffset + 4]);
      const c5 = cleanCell(row[colOffset + 5]);
      const c6 = cleanCell(row[colOffset + 6]);
      const c7 = cleanCell(row[colOffset + 7]);
      const c8 = cleanCell(row[colOffset + 8]);
      const c9 = cleanCell(row[colOffset + 9]);
      const c10 = cleanCell(row[colOffset + 10]);

      const nonEmptyCount = [c0, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10].filter(Boolean).length;
      if (nonEmptyCount < 3) return;

      if (
        index <= 2 &&
        isNaN(Number(c0)) &&
        (c3.includes('曜日') || c5.includes('科目') || c8.includes('試験') || c1.includes('学科'))
      ) {
        return;
      }

      const noNum = parseInt(c0, 10);
      const parsedExam = parseExamDateTime(c8);
      const regDayOrder = DAY_ORDER[c3] || 99;
      const regPeriods = parsePeriodNumbers(c4);
      const regPeriodFirst = regPeriods.length > 0 ? regPeriods[0] : 99;
      const regularSortKey = regDayOrder * 100 + regPeriodFirst;

      const stableKey = `${c2}__${normalizeText(c5)}__${normalizeText(c6)}__${c0}`;
      const searchBlob = normalizeText([c0, c1, c2, `${c3}${c4}`, c5, c6, c7, c8, c9, c10].join(' '));

      records.push({
        uid: `${index}_${c0}_${c2}`,
        stableKey,
        no: Number.isNaN(noNum) ? c0 : noNum,
        noSort: Number.isNaN(noNum) ? Number.MAX_SAFE_INTEGER : noNum,
        department: c1 || '未分類',
        classCode: c2,
        regularDay: c3,
        regularPeriod: c4,
        regularDisplay: `${c3}${c4}`,
        regularSortKey,
        subject: c5,
        subjectNorm: normalizeText(c5),
        instructor: c6,
        campus: c7,
        examRaw: c8,
        examDateLabel: parsedExam.dateLabel,
        examPeriodLabel: parsedExam.periodLabel,
        examPeriods: parsedExam.periods,
        examSortKey: parsedExam.sortKey,
        examDateSortKey: parsedExam.dateSortKey,
        examMonth: parsedExam.month,
        examDay: parsedExam.day,
        room: c9,
        roomsList: c9 ? c9.split(/[\s\u3000]+/).filter(Boolean) : [],
        status: c10 || '実施する',
        searchBlob
      });
    });

    return records;
  }

  // =========================================================================
  // 3. Application State
  // =========================================================================
  const state = {
    records: [],
    meta: null,
    pinnedKeys: new Set(),
    conflictUids: new Set(),
    viewMode: 'all', // 'all' | 'pinned' | 'timeline'
    mobileMode: false,
    mobileModeManual: false,
    mobileFiltersOpen: false,
    filters: {
      query: '',
      dept: '',
      campus: '',
      examDate: '',
      regDay: '',
      regPeriod: '',
      status: ''
    },
    sortBy: 'examDate_asc'
  };

  // =========================================================================
  // 4. DOM References
  // =========================================================================
  const dom = {
    globalDropOverlay: document.getElementById('globalDropOverlay'),
    dbMetaText: document.getElementById('dbMetaText'),
    mobileModeBtn: document.getElementById('mobileModeBtn'),
    mobileModeBtnLabel: document.getElementById('mobileModeBtnLabel'),
    uploadPanel: document.getElementById('uploadPanel'),
    toggleUploadBtn: document.getElementById('toggleUploadBtn'),
    loadSampleBtn: document.getElementById('loadSampleBtn'),
    printBtn: document.getElementById('printBtn'),
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('fileInput'),

    errorBanner: document.getElementById('errorBanner'),
    errorMessage: document.getElementById('errorMessage'),
    closeErrorBtn: document.getElementById('closeErrorBtn'),

    conflictBanner: document.getElementById('conflictBanner'),
    conflictMessage: document.getElementById('conflictMessage'),

    tabButtons: document.querySelectorAll('.tab-btn'),
    countAll: document.getElementById('countAll'),
    countPinned: document.getElementById('countPinned'),

    exportIcsBtn: document.getElementById('exportIcsBtn'),
    clearPinsBtn: document.getElementById('clearPinsBtn'),
    resetDbBtn: document.getElementById('resetDbBtn'),

    searchInput: document.getElementById('searchInput'),
    clearSearchBtn: document.getElementById('clearSearchBtn'),
    sortSelect: document.getElementById('sortSelect'),

    toggleMobileFiltersBtn: document.getElementById('toggleMobileFiltersBtn'),
    activeFilterCountBadge: document.getElementById('activeFilterCountBadge'),
    mobileFilterChevron: document.getElementById('mobileFilterChevron'),
    filterGrid: document.getElementById('filterGrid'),

    filterDept: document.getElementById('filterDept'),
    filterCampus: document.getElementById('filterCampus'),
    filterExamDate: document.getElementById('filterExamDate'),
    filterRegDay: document.getElementById('filterRegDay'),
    filterRegPeriod: document.getElementById('filterRegPeriod'),
    filterStatus: document.getElementById('filterStatus'),
    resetFiltersBtn: document.getElementById('resetFiltersBtn'),
    resultStats: document.getElementById('resultStats'),

    tableView: document.getElementById('tableView'),
    tableBody: document.getElementById('tableBody'),
    mobileCardList: document.getElementById('mobileCardList'),
    sortableHeaders: document.querySelectorAll('th.sortable'),
    emptyState: document.getElementById('emptyState'),
    emptyTitle: document.getElementById('emptyTitle'),
    emptyDesc: document.getElementById('emptyDesc'),

    timelineView: document.getElementById('timelineView'),

    printTitle: document.getElementById('printTitle'),
    printMeta: document.getElementById('printMeta'),
    toast: document.getElementById('toast')
  };

  let toastTimer = null;
  function showToast(message) {
    dom.toast.textContent = message;
    dom.toast.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      dom.toast.classList.remove('show');
    }, 3000);
  }

  function showError(message) {
    dom.errorMessage.textContent = message;
    dom.errorBanner.classList.remove('hidden');
  }

  function hideError() {
    dom.errorBanner.classList.add('hidden');
    dom.errorMessage.textContent = '';
  }

  // =========================================================================
  // 4.5 Smartphone Mode Controller
  // =========================================================================
  function applyMobileMode(enabled) {
    state.mobileMode = Boolean(enabled);
    document.body.classList.toggle('mobile-mode', state.mobileMode);

    dom.mobileModeBtn.setAttribute('aria-pressed', state.mobileMode ? 'true' : 'false');
    dom.mobileModeBtn.classList.toggle('btn-mode-active', state.mobileMode);
    dom.mobileModeBtnLabel.textContent = state.mobileMode ? 'スマホ表示: ON' : 'スマホモード';

    syncMobileFilterDrawer();
  }

  function syncMobileFilterDrawer() {
    if (!state.mobileMode) {
      dom.filterGrid.classList.remove('mobile-collapsed');
      return;
    }
    dom.filterGrid.classList.toggle('mobile-collapsed', !state.mobileFiltersOpen);
    dom.toggleMobileFiltersBtn.setAttribute('aria-expanded', state.mobileFiltersOpen ? 'true' : 'false');
    dom.mobileFilterChevron.textContent = state.mobileFiltersOpen ? '▲' : '▼';
  }

  function updateActiveFilterBadge() {
    const keys = ['dept', 'campus', 'examDate', 'regDay', 'regPeriod', 'status'];
    const activeCount = keys.filter((k) => Boolean(state.filters[k])).length;
    if (activeCount > 0) {
      dom.activeFilterCountBadge.textContent = `${activeCount}件適用中`;
      dom.activeFilterCountBadge.classList.remove('hidden');
    } else {
      dom.activeFilterCountBadge.classList.add('hidden');
    }
  }

  function initMobileMode() {
    const savedPref = localStorage.getItem(MOBILE_PREF_KEY);
    const mq = window.matchMedia('(max-width: 768px)');

    if (savedPref === 'on') {
      state.mobileModeManual = true;
      applyMobileMode(true);
    } else if (savedPref === 'off') {
      state.mobileModeManual = true;
      applyMobileMode(false);
    } else {
      state.mobileModeManual = false;
      applyMobileMode(mq.matches);
    }

    const handleMediaChange = (e) => {
      if (!state.mobileModeManual) {
        applyMobileMode(e.matches);
      }
    };

    if (typeof mq.addEventListener === 'function') {
      mq.addEventListener('change', handleMediaChange);
    } else if (typeof mq.addListener === 'function') {
      mq.addListener(handleMediaChange);
    }
  }

  // =========================================================================
  // 5. Excel File Ingestion
  // =========================================================================
  function readFileAsArrayBuffer(file) {
    if (typeof file.arrayBuffer === 'function') {
      return file.arrayBuffer();
    }
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error || new Error('ファイルの読み取りに失敗しました'));
      reader.readAsArrayBuffer(file);
    });
  }

  async function handleExcelFile(file) {
    if (!file) return;
    hideError();

    try {
      const XLSXLib = await ensureXlsxLoaded();
      const buffer = await readFileAsArrayBuffer(file);
      const workbook = XLSXLib.read(new Uint8Array(buffer), { type: 'array' });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        showError('Excelファイル内にシートが見つかりませんでした。');
        return;
      }

      let targetSheetName = workbook.SheetNames[0];
      let parsed = [];

      for (const sheetName of workbook.SheetNames) {
        const ws = workbook.Sheets[sheetName];
        const rows = XLSXLib.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const candidate = parseRowsToRecords(rows);
        if (candidate.length > 0) {
          targetSheetName = sheetName;
          parsed = candidate;
          break;
        }
      }

      if (parsed.length === 0) {
        showError(
          `「${file.name}」から有効な試験データ行を検出できませんでした。11列構成（通し番号・学科・クラス・曜日・時限・科目名・担当教員・校舎・試験日時・教室・実施有無）になっているか確認してください。`
        );
        return;
      }

      const meta = {
        fileName: file.name,
        sheetName: targetSheetName,
        totalRows: parsed.length,
        updatedAt: new Date().toLocaleString('ja-JP')
      };

      state.records = parsed;
      state.meta = meta;

      await StorageService.set('records', parsed);
      await StorageService.set('meta', meta);

      dom.uploadPanel.classList.add('hidden');
      refreshAllUI();
      showToast(`「${file.name}」から ${parsed.length} 件のデータをデータベース化しました`);
    } catch (err) {
      console.error(err);
      showError(`ファイル読み込みエラー: ${err?.message || '不明なエラーが発生しました'}`);
    }
  }

  async function loadSampleDataset() {
    hideError();
    const rawRows = window.__SAMPLE_EXAM_ROWS__ || [];
    const parsed = parseRowsToRecords(rawRows);
    const meta = {
      fileName: 'テスト一覧例エクセル.xlsx（サンプル）',
      sheetName: 'Sheet1',
      totalRows: parsed.length,
      updatedAt: new Date().toLocaleString('ja-JP')
    };

    state.records = parsed;
    state.meta = meta;

    await StorageService.set('records', parsed);
    await StorageService.set('meta', meta);

    dom.uploadPanel.classList.add('hidden');
    refreshAllUI();
    showToast(`サンプルデータ（${parsed.length}件）を読み込みました`);
  }

  // =========================================================================
  // 6. Dynamic Filter Options & Conflict Detection
  // =========================================================================
  function populateFilterDropdowns() {
    const depts = new Map();
    const campuses = new Map();
    const examDates = new Map();
    const regPeriods = new Set();

    state.records.forEach((rec) => {
      if (rec.department) {
        depts.set(rec.department, (depts.get(rec.department) || 0) + 1);
      }
      if (rec.campus) {
        campuses.set(rec.campus, (campuses.get(rec.campus) || 0) + 1);
      }
      if (rec.examDateLabel && rec.examDateLabel !== '日時未定・記載なし') {
        if (!examDates.has(rec.examDateLabel)) {
          examDates.set(rec.examDateLabel, { sortKey: rec.examDateSortKey, count: 0 });
        }
        examDates.get(rec.examDateLabel).count += 1;
      }
      if (rec.regularPeriod) {
        regPeriods.add(rec.regularPeriod);
      }
    });

    const prevDept = state.filters.dept;
    dom.filterDept.innerHTML = '<option value="">すべての学科・区分</option>';
    [...depts.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], 'ja'))
      .forEach(([dept, count]) => {
        const opt = document.createElement('option');
        opt.value = dept;
        opt.textContent = `${dept} (${count})`;
        dom.filterDept.appendChild(opt);
      });
    dom.filterDept.value = depts.has(prevDept) ? prevDept : '';
    state.filters.dept = dom.filterDept.value;

    const prevCampus = state.filters.campus;
    dom.filterCampus.innerHTML = '<option value="">すべての校舎</option>';
    [...campuses.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], 'ja'))
      .forEach(([campus, count]) => {
        const opt = document.createElement('option');
        opt.value = campus;
        opt.textContent = `${campus} (${count})`;
        dom.filterCampus.appendChild(opt);
      });
    dom.filterCampus.value = campuses.has(prevCampus) ? prevCampus : '';
    state.filters.campus = dom.filterCampus.value;

    const prevExamDate = state.filters.examDate;
    dom.filterExamDate.innerHTML = '<option value="">すべての試験日</option>';
    [...examDates.entries()]
      .sort((a, b) => a[1].sortKey - b[1].sortKey)
      .forEach(([label, info]) => {
        const opt = document.createElement('option');
        opt.value = label;
        opt.textContent = `${label} (${info.count})`;
        dom.filterExamDate.appendChild(opt);
      });
    dom.filterExamDate.value = examDates.has(prevExamDate) ? prevExamDate : '';
    state.filters.examDate = dom.filterExamDate.value;

    const prevRegPeriod = state.filters.regPeriod;
    dom.filterRegPeriod.innerHTML = '<option value="">すべて</option>';
    [...regPeriods]
      .sort((a, b) => a.localeCompare(b, 'ja', { numeric: true }))
      .forEach((p) => {
        const opt = document.createElement('option');
        opt.value = p;
        opt.textContent = `${p}時限`;
        dom.filterRegPeriod.appendChild(opt);
      });
    dom.filterRegPeriod.value = regPeriods.has(prevRegPeriod) ? prevRegPeriod : '';
    state.filters.regPeriod = dom.filterRegPeriod.value;
  }

  function detectPinnedConflicts() {
    state.conflictUids.clear();
    const pinnedRecords = state.records.filter(
      (r) => state.pinnedKeys.has(r.stableKey) && r.status !== '実施しない' && r.examMonth !== null
    );

    const conflictMessages = [];

    for (let i = 0; i < pinnedRecords.length; i++) {
      for (let j = i + 1; j < pinnedRecords.length; j++) {
        const a = pinnedRecords[i];
        const b = pinnedRecords[j];
        if (a.examDateSortKey === b.examDateSortKey) {
          const overlap = a.examPeriods.some((p) => b.examPeriods.includes(p));
          if (overlap) {
            state.conflictUids.add(a.uid);
            state.conflictUids.add(b.uid);
            conflictMessages.push(
              `【${a.examDateLabel} ${a.examPeriodLabel}】「${a.subject}」と「${b.subject}」が重複しています`
            );
          }
        }
      }
    }

    if (conflictMessages.length > 0) {
      dom.conflictBanner.classList.remove('hidden');
      dom.conflictMessage.innerHTML = `<strong>試験日時の重複警告：</strong> ${conflictMessages.join(' ／ ')}`;
    } else {
      dom.conflictBanner.classList.add('hidden');
      dom.conflictMessage.textContent = '';
    }
  }

  // =========================================================================
  // 7. Filtering & Sorting Pipeline
  // =========================================================================
  function getFilteredAndSortedRecords() {
    const qNorm = normalizeText(state.filters.query);

    const filtered = state.records.filter((rec) => {
      if (state.viewMode === 'pinned' && !state.pinnedKeys.has(rec.stableKey)) {
        return false;
      }
      if (state.filters.dept && rec.department !== state.filters.dept) {
        return false;
      }
      if (state.filters.campus && rec.campus !== state.filters.campus) {
        return false;
      }
      if (state.filters.examDate && rec.examDateLabel !== state.filters.examDate) {
        return false;
      }
      if (state.filters.regDay && rec.regularDay !== state.filters.regDay) {
        return false;
      }
      if (state.filters.regPeriod && rec.regularPeriod !== state.filters.regPeriod) {
        return false;
      }
      if (state.filters.status && rec.status !== state.filters.status) {
        return false;
      }
      if (qNorm && !rec.searchBlob.includes(qNorm)) {
        return false;
      }
      return true;
    });

    const [field, dir] = state.sortBy.split('_');
    const sign = dir === 'desc' ? -1 : 1;

    filtered.sort((a, b) => {
      let cmp = 0;
      switch (field) {
        case 'examDate':
          cmp = a.examSortKey - b.examSortKey;
          break;
        case 'subject':
          cmp = a.subjectNorm.localeCompare(b.subjectNorm, 'ja');
          break;
        case 'dept':
          cmp = a.department.localeCompare(b.department, 'ja') || a.classCode.localeCompare(b.classCode, 'ja');
          break;
        case 'code':
          cmp = a.classCode.localeCompare(b.classCode, 'ja');
          break;
        case 'regular':
          cmp = a.regularSortKey - b.regularSortKey;
          break;
        case 'instructor':
          cmp = a.instructor.localeCompare(b.instructor, 'ja');
          break;
        case 'campus':
          cmp = a.campus.localeCompare(b.campus, 'ja');
          break;
        case 'room':
          cmp = a.room.localeCompare(b.room, 'ja', { numeric: true });
          break;
        case 'status':
          cmp = a.status.localeCompare(b.status, 'ja');
          break;
        case 'no':
        default:
          cmp = a.noSort - b.noSort;
          break;
      }
      if (cmp === 0) {
        cmp = a.noSort - b.noSort;
      }
      return cmp * sign;
    });

    return filtered;
  }

  // =========================================================================
  // 8. Rendering
  // =========================================================================
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function updateHeaderAndCounts() {
    if (state.meta && state.records.length > 0) {
      dom.dbMetaText.textContent = `読込済: ${state.meta.fileName}（全 ${state.records.length} 件 / 更新: ${state.meta.updatedAt}）`;
    } else {
      dom.dbMetaText.textContent = 'Excelファイルを読み込んでデータベース化します（ブラウザ内ローカル保存）';
    }

    dom.countAll.textContent = state.records.length;
    const activePinnedCount = state.records.filter((r) => state.pinnedKeys.has(r.stableKey)).length;
    dom.countPinned.textContent = activePinnedCount;
  }

  function updateSortHeaderIndicators() {
    const [field, dir] = state.sortBy.split('_');
    dom.sortableHeaders.forEach((th) => {
      th.classList.remove('sort-asc', 'sort-desc');
      if (th.dataset.sort === field) {
        th.classList.add(dir === 'desc' ? 'sort-desc' : 'sort-asc');
      }
    });
  }

  function renderTableAndCards(records) {
    dom.tableBody.innerHTML = '';
    dom.mobileCardList.innerHTML = '';

    if (records.length === 0) {
      dom.emptyState.classList.remove('hidden');
      if (state.records.length === 0) {
        dom.emptyTitle.textContent = 'データベースにデータがありません';
        dom.emptyDesc.textContent = '上部のエリアからExcelファイルを読み込むか、「サンプル読込」を押してください。';
      } else if (state.viewMode === 'pinned') {
        dom.emptyTitle.textContent = 'マイ試験日程に登録された科目がありません';
        dom.emptyDesc.textContent = '「全科目一覧」タブで履修している科目の「☆」をタップすると登録できます。';
      } else {
        dom.emptyTitle.textContent = '条件に一致する試験が見つかりませんでした';
        dom.emptyDesc.textContent = '検索キーワードや絞り込み条件を変更してみてください。';
      }
      return;
    }

    dom.emptyState.classList.add('hidden');
    const tableFrag = document.createDocumentFragment();
    const cardFrag = document.createDocumentFragment();

    records.forEach((rec) => {
      const isPinned = state.pinnedKeys.has(rec.stableKey);
      const isCancelled = rec.status === '実施しない';
      const isConflict = state.conflictUids.has(rec.uid);

      const roomsHtml =
        rec.roomsList.length > 0
          ? rec.roomsList.map((rm) => `<span class="room-chip">${escapeHtml(rm)}</span>`).join('')
          : '<span class="text-muted">-</span>';

      const statusClass = isCancelled ? 'status-inactive' : 'status-active';

      // 1. Desktop Table Row
      const tr = document.createElement('tr');
      if (isPinned) tr.classList.add('row-pinned');
      if (isCancelled) tr.classList.add('row-cancelled');
      if (isConflict) tr.classList.add('row-conflict');

      tr.innerHTML = `
        <td class="col-pin no-print">
          <button type="button" class="pin-btn ${isPinned ? 'pinned' : ''}" data-key="${escapeHtml(rec.stableKey)}" title="${isPinned ? 'マイ日程から解除' : 'マイ日程に登録'}">
            ${isPinned ? '★' : '☆'}
          </button>
        </td>
        <td class="col-no">${escapeHtml(rec.no)}</td>
        <td class="col-dept"><span class="dept-pill">${escapeHtml(rec.department)}</span></td>
        <td class="col-code">${escapeHtml(rec.classCode)}</td>
        <td class="col-reg">${escapeHtml(rec.regularDisplay)}</td>
        <td class="col-subject">${escapeHtml(rec.subject)}</td>
        <td class="col-teacher">${escapeHtml(rec.instructor)}</td>
        <td class="col-campus">${escapeHtml(rec.campus)}</td>
        <td class="col-exam">${escapeHtml(rec.examRaw || '-')}</td>
        <td class="col-room">${roomsHtml}</td>
        <td class="col-status"><span class="status-badge ${statusClass}">${escapeHtml(rec.status)}</span></td>
      `;
      tableFrag.appendChild(tr);

      // 2. Smartphone Mode Card
      const card = document.createElement('article');
      card.className = 'exam-card';
      if (isPinned) card.classList.add('row-pinned');
      if (isCancelled) card.classList.add('row-cancelled');
      if (isConflict) card.classList.add('row-conflict');

      const conflictHtml = isConflict
        ? `<div class="exam-card-conflict-tag">⚠️ 試験日時が他の登録科目と重複しています</div>`
        : '';

      card.innerHTML = `
        <div class="exam-card-top">
          <div class="exam-card-schedule">
            <span class="exam-card-datetime">${escapeHtml(rec.examRaw || '日時未定')}</span>
            ${rec.roomsList.length > 0 ? roomsHtml : ''}
          </div>
          <div class="exam-card-actions">
            <span class="status-badge ${statusClass}">${escapeHtml(rec.status)}</span>
            <button type="button" class="pin-btn ${isPinned ? 'pinned' : ''}" data-key="${escapeHtml(rec.stableKey)}" aria-label="${isPinned ? 'マイ日程から解除' : 'マイ日程に登録'}">
              ${isPinned ? '★' : '☆'}
            </button>
          </div>
        </div>
        <h3 class="exam-card-subject">${escapeHtml(rec.subject)}</h3>
        ${conflictHtml}
        <div class="exam-card-meta">
          <span class="dept-pill">${escapeHtml(rec.department)}</span>
          <span class="exam-card-code">${escapeHtml(rec.classCode)}</span>
          <span class="exam-card-meta-item">通常: ${escapeHtml(rec.regularDisplay)}</span>
          <span class="exam-card-meta-item">担当: ${escapeHtml(rec.instructor)}</span>
          <span class="exam-card-meta-item">${escapeHtml(rec.campus)}</span>
          <span class="exam-card-no">No.${escapeHtml(rec.no)}</span>
        </div>
      `;
      cardFrag.appendChild(card);
    });

    dom.tableBody.appendChild(tableFrag);
    dom.mobileCardList.appendChild(cardFrag);
  }

  function renderTimeline(records) {
    dom.timelineView.innerHTML = '';

    if (records.length === 0) {
      dom.timelineView.innerHTML = `
        <div class="table-card">
          <div class="empty-state">
            <p class="empty-title">表示する試験日程がありません</p>
            <p class="empty-desc">データを読み込むか、絞り込み条件を解除してください。</p>
          </div>
        </div>
      `;
      return;
    }

    const sorted = [...records].sort((a, b) => a.examSortKey - b.examSortKey || a.noSort - b.noSort);
    const groups = new Map();

    sorted.forEach((rec) => {
      const key = rec.examDateLabel;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(rec);
    });

    const fragment = document.createDocumentFragment();

    groups.forEach((items, dateLabel) => {
      const card = document.createElement('article');
      card.className = 'timeline-day-card';

      const itemsHtml = items
        .map((rec) => {
          const isPinned = state.pinnedKeys.has(rec.stableKey);
          const isCancelled = rec.status === '実施しない';
          const roomsHtml = rec.roomsList
            .map((rm) => `<span class="room-chip">${escapeHtml(rm)}</span>`)
            .join('');
          const statusClass = isCancelled ? 'status-inactive' : 'status-active';

          return `
            <div class="timeline-item ${isPinned ? 'row-pinned' : ''} ${isCancelled ? 'row-cancelled' : ''}">
              <button type="button" class="pin-btn no-print ${isPinned ? 'pinned' : ''}" data-key="${escapeHtml(rec.stableKey)}">
                ${isPinned ? '★' : '☆'}
              </button>
              <div class="timeline-period">${escapeHtml(rec.examPeriodLabel || '時限未定')}</div>
              <div class="timeline-main">
                <div class="timeline-subject">${escapeHtml(rec.subject)}</div>
                <div class="timeline-submeta">
                  ${escapeHtml(rec.department)} / ${escapeHtml(rec.classCode)} ・ 担当: ${escapeHtml(rec.instructor)}（通常: ${escapeHtml(rec.regularDisplay)}）
                </div>
              </div>
              <div class="timeline-right">
                <span class="dept-pill">${escapeHtml(rec.campus)}</span>
                ${roomsHtml}
                <span class="status-badge ${statusClass}">${escapeHtml(rec.status)}</span>
              </div>
            </div>
          `;
        })
        .join('');

      card.innerHTML = `
        <div class="timeline-day-header">
          <h3 class="timeline-day-title">${escapeHtml(dateLabel)}</h3>
          <span class="timeline-day-count">${items.length} 科目</span>
        </div>
        <div class="timeline-list">${itemsHtml}</div>
      `;

      fragment.appendChild(card);
    });

    dom.timelineView.appendChild(fragment);
  }

  function renderCurrentView() {
    const filtered = getFilteredAndSortedRecords();
    dom.resultStats.innerHTML = `表示件数: <strong>${filtered.length}</strong> / 全 ${state.records.length} 件`;
    updateActiveFilterBadge();

    const modeLabel =
      state.viewMode === 'pinned'
        ? 'マイ定期試験日程表'
        : state.viewMode === 'timeline'
        ? '定期試験日程（日別タイムライン）'
        : '定期試験日程一覧';
    dom.printTitle.textContent = modeLabel;
    dom.printMeta.textContent = `出力日時: ${new Date().toLocaleString('ja-JP')} ／ 表示件数: ${filtered.length}件`;

    updateSortHeaderIndicators();

    if (state.viewMode === 'timeline') {
      dom.tableView.classList.add('hidden');
      dom.timelineView.classList.remove('hidden');
      renderTimeline(filtered);
    } else {
      dom.timelineView.classList.add('hidden');
      dom.tableView.classList.remove('hidden');
      renderTableAndCards(filtered);
    }
  }

  function refreshAllUI() {
    populateFilterDropdowns();
    detectPinnedConflicts();
    updateHeaderAndCounts();
    renderCurrentView();
  }

  // =========================================================================
  // 9. Pinning & ICS Calendar Export
  // =========================================================================
  async function togglePin(stableKey) {
    if (!stableKey) return;
    if (state.pinnedKeys.has(stableKey)) {
      state.pinnedKeys.delete(stableKey);
    } else {
      state.pinnedKeys.add(stableKey);
    }
    await StorageService.set('pinnedKeys', [...state.pinnedKeys]);
    detectPinnedConflicts();
    updateHeaderAndCounts();
    renderCurrentView();
  }

  function exportPinnedToICS() {
    const pinned = state.records.filter(
      (r) => state.pinnedKeys.has(r.stableKey) && r.examMonth !== null && r.examDay !== null
    );

    if (pinned.length === 0) {
      showToast('書き出し可能な「★マイ試験日程」がありません');
      return;
    }

    const periodTimes = {
      1: ['090000', '104000'],
      2: ['105000', '123000'],
      3: ['132000', '150000'],
      4: ['151000', '165000'],
      5: ['170000', '184000'],
      6: ['185000', '203000']
    };

    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1;

    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//UniversityExamScheduleViewer//JP',
      'CALSCALE:GREGORIAN'
    ];

    pinned.forEach((rec, idx) => {
      let year = currentYear;
      if (currentMonth >= 9 && rec.examMonth <= 3) {
        year = currentYear + 1;
      }

      const mm = String(rec.examMonth).padStart(2, '0');
      const dd = String(rec.examDay).padStart(2, '0');
      const startP = rec.examPeriods[0] || 1;
      const endP = rec.examPeriods[rec.examPeriods.length - 1] || startP;
      const startTime = (periodTimes[startP] || ['090000', '104000'])[0];
      const endTime = (periodTimes[endP] || ['090000', '104000'])[1];

      lines.push('BEGIN:VEVENT');
      lines.push(`UID:exam-${rec.no}-${idx}-${year}${mm}${dd}@local`);
      lines.push(`DTSTART:${year}${mm}${dd}T${startTime}`);
      lines.push(`DTEND:${year}${mm}${dd}T${endTime}`);
      lines.push(`SUMMARY:【試験】${rec.subject} (${rec.examPeriodLabel})`);
      lines.push(`LOCATION:${rec.campus} ${rec.room}`);
      lines.push(`DESCRIPTION:クラス: ${rec.classCode} / 担当: ${rec.instructor} / 状態: ${rec.status}`);
      lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');

    const blob = new Blob([lines.join('\r\n')], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'my_exam_schedule.ics';
    a.click();
    URL.revokeObjectURL(url);
    showToast(`${pinned.length}件のマイ試験日程を .ics 形式で書き出しました`);
  }

  // =========================================================================
  // 10. Event Listeners
  // =========================================================================
  function extractDroppedFile(dataTransfer) {
    if (!dataTransfer) return null;
    if (dataTransfer.files && dataTransfer.files.length > 0) {
      return dataTransfer.files[0];
    }
    if (dataTransfer.items && dataTransfer.items.length > 0) {
      for (const item of dataTransfer.items) {
        if (item.kind === 'file') {
          return item.getAsFile();
        }
      }
    }
    return null;
  }

  function bindEvents() {
    dom.closeErrorBtn.addEventListener('click', hideError);

    // Smartphone Mode Toggle
    dom.mobileModeBtn.addEventListener('click', () => {
      const nextMode = !state.mobileMode;
      state.mobileModeManual = true;
      localStorage.setItem(MOBILE_PREF_KEY, nextMode ? 'on' : 'off');
      applyMobileMode(nextMode);
      showToast(
        nextMode
          ? 'スマホモードをONにしました（カード表示）'
          : 'スマホモードをOFFにしました（PC一覧表示）'
      );
    });

    // Collapsible Mobile Filters Drawer
    dom.toggleMobileFiltersBtn.addEventListener('click', () => {
      state.mobileFiltersOpen = !state.mobileFiltersOpen;
      syncMobileFilterDrawer();
    });

    // Upload panel toggle
    dom.toggleUploadBtn.addEventListener('click', () => {
      dom.uploadPanel.classList.toggle('hidden');
    });

    // File input change
    dom.fileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) {
        handleExcelFile(file);
        e.target.value = '';
      }
    });

    // Dropzone direct drag-drop
    ['dragenter', 'dragover'].forEach((evtName) => {
      dom.dropzone.addEventListener(evtName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dom.dropzone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach((evtName) => {
      dom.dropzone.addEventListener(evtName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dom.dropzone.classList.remove('drag-over');
      });
    });

    dom.dropzone.addEventListener('drop', (e) => {
      dom.globalDropOverlay.classList.add('hidden');
      dragDepth = 0;
      const file = extractDroppedFile(e.dataTransfer);
      if (file) handleExcelFile(file);
    });

    // Full-window drag-and-drop support
    let dragDepth = 0;
    window.addEventListener('dragenter', (e) => {
      if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) {
        e.preventDefault();
        dragDepth += 1;
        dom.globalDropOverlay.classList.remove('hidden');
      }
    });

    window.addEventListener('dragover', (e) => {
      if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) {
        e.preventDefault();
      }
    });

    window.addEventListener('dragleave', (e) => {
      if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) {
        dragDepth = Math.max(0, dragDepth - 1);
        if (dragDepth === 0) {
          dom.globalDropOverlay.classList.add('hidden');
        }
      }
    });

    window.addEventListener('drop', (e) => {
      e.preventDefault();
      dragDepth = 0;
      dom.globalDropOverlay.classList.add('hidden');
      const file = extractDroppedFile(e.dataTransfer);
      if (file) handleExcelFile(file);
    });

    // Sample data & Print
    dom.loadSampleBtn.addEventListener('click', loadSampleDataset);
    dom.printBtn.addEventListener('click', () => window.print());

    // View mode tabs
    dom.tabButtons.forEach((btn) => {
      btn.addEventListener('click', () => {
        dom.tabButtons.forEach((b) => {
          b.classList.remove('active');
          b.setAttribute('aria-selected', 'false');
        });
        btn.classList.add('active');
        btn.setAttribute('aria-selected', 'true');
        state.viewMode = btn.dataset.view;
        renderCurrentView();
      });
    });

    // Search input (debounced)
    let searchDebounce = null;
    dom.searchInput.addEventListener('input', (e) => {
      const val = e.target.value;
      dom.clearSearchBtn.classList.toggle('hidden', val.length === 0);
      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(() => {
        state.filters.query = val;
        renderCurrentView();
      }, 120);
    });

    dom.clearSearchBtn.addEventListener('click', () => {
      dom.searchInput.value = '';
      dom.clearSearchBtn.classList.add('hidden');
      state.filters.query = '';
      renderCurrentView();
      dom.searchInput.focus();
    });

    // Dropdown filters
    const filterBindings = [
      [dom.filterDept, 'dept'],
      [dom.filterCampus, 'campus'],
      [dom.filterExamDate, 'examDate'],
      [dom.filterRegDay, 'regDay'],
      [dom.filterRegPeriod, 'regPeriod'],
      [dom.filterStatus, 'status']
    ];

    filterBindings.forEach(([el, key]) => {
      el.addEventListener('change', () => {
        state.filters[key] = el.value;
        renderCurrentView();
      });
    });

    // Reset filters
    dom.resetFiltersBtn.addEventListener('click', () => {
      state.filters = {
        query: '',
        dept: '',
        campus: '',
        examDate: '',
        regDay: '',
        regPeriod: '',
        status: ''
      };
      dom.searchInput.value = '';
      dom.clearSearchBtn.classList.add('hidden');
      dom.filterDept.value = '';
      dom.filterCampus.value = '';
      dom.filterExamDate.value = '';
      dom.filterRegDay.value = '';
      dom.filterRegPeriod.value = '';
      dom.filterStatus.value = '';
      renderCurrentView();
      showToast('絞り込み条件をリセットしました');
    });

    // Sort select & Table column headers
    dom.sortSelect.addEventListener('change', () => {
      state.sortBy = dom.sortSelect.value;
      renderCurrentView();
    });

    dom.sortableHeaders.forEach((th) => {
      th.addEventListener('click', () => {
        const field = th.dataset.sort;
        const [curField, curDir] = state.sortBy.split('_');
        const nextDir = curField === field && curDir === 'asc' ? 'desc' : 'asc';
        state.sortBy = `${field}_${nextDir}`;
        if ([...dom.sortSelect.options].some((o) => o.value === state.sortBy)) {
          dom.sortSelect.value = state.sortBy;
        }
        renderCurrentView();
      });
    });

    // Event delegation for Pin buttons
    document.addEventListener('click', (e) => {
      const pinBtn = e.target.closest('.pin-btn');
      if (pinBtn) {
        togglePin(pinBtn.dataset.key);
      }
    });

    // Export ICS / Clear Pins / Reset DB
    dom.exportIcsBtn.addEventListener('click', exportPinnedToICS);

    dom.clearPinsBtn.addEventListener('click', async () => {
      if (state.pinnedKeys.size === 0) return;
      state.pinnedKeys.clear();
      await StorageService.set('pinnedKeys', []);
      detectPinnedConflicts();
      updateHeaderAndCounts();
      renderCurrentView();
      showToast('マイ試験日程の選択をすべて解除しました');
    });

    dom.resetDbBtn.addEventListener('click', async () => {
      if (!confirm('保存されている試験日程データベースとマイ日程をすべて削除しますか？')) return;
      await StorageService.clearAll();
      state.records = [];
      state.meta = null;
      state.pinnedKeys.clear();
      dom.uploadPanel.classList.remove('hidden');
      refreshAllUI();
      showToast('データベースを初期化しました');
    });
  }

  // =========================================================================
  // 11. Initialization
  // =========================================================================
  async function init() {
    initMobileMode();
    bindEvents();

    const [savedRecords, savedMeta, savedPins] = await Promise.all([
      StorageService.get('records'),
      StorageService.get('meta'),
      StorageService.get('pinnedKeys')
    ]);

    if (Array.isArray(savedPins)) {
      state.pinnedKeys = new Set(savedPins);
    }

    if (Array.isArray(savedRecords) && savedRecords.length > 0) {
      state.records = savedRecords;
      state.meta = savedMeta;
      dom.uploadPanel.classList.add('hidden');
    } else {
      dom.uploadPanel.classList.remove('hidden');
    }

    refreshAllUI();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
