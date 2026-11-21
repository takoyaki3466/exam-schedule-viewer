/**
 * 定期試験日程データベース (script.js)
 * - ハンバーガーメニュー統合型・左スライドインドロワ（絞り込み条件＋データ管理操作）
 * - 初期並び替え：科目名順（あいうえお順 / Intl.Collator('ja') 安定ソート）
 * - 確認モーダル付き破壊的・置換操作（マイ試験日程をすべて解除 / 保存データの全削除 / サンプル読込）
 * - スクロール性能最適化（印刷用テーブルのオンデマンド生成・段階描画・ピン留め差分DOM更新）
 * - 実処理閾値連動の単色濃淡スケルトンUI
 */

(() => {
  'use strict';

  // =========================================================================
  // 0. Dynamic SheetJS Loader (Fallback if primary CDN is slow/blocked)
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
  // 2. Normalization, Japanese Collator & Excel Row Parsing
  // =========================================================================
  const DAY_ORDER = { '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6, '日': 7 };
  const ALL_WEEKDAYS = ['月', '火', '水', '木', '金', '土', '日'];
  const JA_COLLATOR = new Intl.Collator('ja', { numeric: true, sensitivity: 'base' });

  function normalizeText(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .normalize('NFKC')
        .toLowerCase()
        .replace(/[\s\u3000]+/g, '');
  }

  function makeSubjectSortKey(value) {
    if (value === null || value === undefined) return '';
    return String(value).normalize('NFKC').trim();
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

      const c0 = cleanCell(row[colOffset + 0]); // No.
      const c1 = cleanCell(row[colOffset + 1]); // 学科
      const c2 = cleanCell(row[colOffset + 2]); // 時間割コード
      const c3 = cleanCell(row[colOffset + 3]); // 曜日
      const c4 = normalizePeriodDisplay(row[colOffset + 4]); // 時限
      const c5 = cleanCell(row[colOffset + 5]); // 科目名
      const c6 = cleanCell(row[colOffset + 6]); // 担当教員
      const c7 = cleanCell(row[colOffset + 7]); // 校舎
      const c8 = cleanCell(row[colOffset + 8]); // 理解度確認テスト日程
      const c9 = cleanCell(row[colOffset + 9]); // 理解度確認テスト実施教室
      const c10 = cleanCell(row[colOffset + 10]); // 追評価の実施

      const nonEmptyCount = [c0, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10].filter(Boolean).length;
      if (nonEmptyCount < 3) return;

      if (
          index <= 2 &&
          isNaN(Number(c0)) &&
          (c3.includes('曜日') || c5.includes('科目') || c8.includes('試験') || c8.includes('テスト') || c1.includes('学科'))
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
        noSort: Number.isNaN(noNum) ? index + 100000 : noNum,
        department: c1 || '未分類',
        classCode: c2,
        regularDay: c3,
        regularPeriod: c4,
        regularDisplay: `${c3}${c4}`,
        regularSortKey,
        subject: c5,
        subjectNorm: normalizeText(c5),
        subjectSortKey: makeSubjectSortKey(c5),
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
  function createEmptyFilters() {
    return {
      query: '',
      dept: '',
      campus: '', // Single-select chip (preserved)
      examDates: new Set(), // Multi-select chips
      regDays: new Set(), // Multi-select chips (regularDay)
      regPeriods: new Set(), // Multi-select chips (regularPeriod)
      status: '' // Segmented control: '' | '実施する' | '実施しない'
    };
  }

  function cloneFilters(src) {
    return {
      query: src.query || '',
      dept: src.dept || '',
      campus: src.campus || '',
      examDates: new Set(src.examDates || []),
      regDays: new Set(src.regDays || []),
      regPeriods: new Set(src.regPeriods || []),
      status: src.status || ''
    };
  }

  const RENDER_CHUNK_SIZE = 60;

  const state = {
    records: [],
    meta: null,
    pinnedKeys: new Set(),
    conflictUids: new Set(),
    viewMode: 'all', // 'all' | 'pinned'
    appliedFilters: createEmptyFilters(),
    draftFilters: createEmptyFilters(),
    sortBy: 'subject_asc', // Default: Subject Name (Japanese Gojuon Order)
    availableOptions: {
      depts: [],
      campuses: [],
      examDates: [],
      regPeriods: []
    },
    currentRenderedRecords: [],
    renderedCount: 0,
    lastDateGroupSection: null,
    lastDateGroupKey: null,
    isDrawerOpen: false,
    isUnsavedModalOpen: false,
    isActionModalOpen: false,
    pendingConfirmAction: null,
    activeDetailRecord: null,
    lastFocusedElement: null
  };

  // =========================================================================
  // 4. DOM References
  // =========================================================================
  const dom = {
    globalDropOverlay: document.getElementById('globalDropOverlay'),
    openMenuDrawerBtn: document.getElementById('openMenuDrawerBtn'),
    activeFilterBadge: document.getElementById('activeFilterBadge'),
    dbMetaText: document.getElementById('dbMetaText'),

    uploadPanel: document.getElementById('uploadPanel'),
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('fileInput'),
    inlineSampleBtn: document.getElementById('inlineSampleBtn'),

    errorBanner: document.getElementById('errorBanner'),
    errorMessage: document.getElementById('errorMessage'),
    closeErrorBtn: document.getElementById('closeErrorBtn'),

    conflictBanner: document.getElementById('conflictBanner'),
    conflictMessage: document.getElementById('conflictMessage'),

    tabButtons: document.querySelectorAll('.tab-btn'),
    countAll: document.getElementById('countAll'),
    countPinned: document.getElementById('countPinned'),

    searchInput: document.getElementById('searchInput'),
    clearSearchBtn: document.getElementById('clearSearchBtn'),
    sortSelect: document.getElementById('sortSelect'),

    menuGuideBar: document.getElementById('menuGuideBar'),
    guideOpenMenuBtn: document.getElementById('guideOpenMenuBtn'),

    activeFiltersBar: document.getElementById('activeFiltersBar'),
    activeFilterTags: document.getElementById('activeFilterTags'),
    editFiltersFromBarBtn: document.getElementById('editFiltersFromBarBtn'),
    quickClearAllFiltersBtn: document.getElementById('quickClearAllFiltersBtn'),
    resultStats: document.getElementById('resultStats'),

    scheduleListSection: document.getElementById('scheduleListSection'),
    skeletonGrid: document.getElementById('skeletonGrid'),
    cardListContainer: document.getElementById('cardListContainer'),
    loadMoreSentinel: document.getElementById('loadMoreSentinel'),
    emptyState: document.getElementById('emptyState'),
    emptyTitle: document.getElementById('emptyTitle'),
    emptyDesc: document.getElementById('emptyDesc'),

    printTitle: document.getElementById('printTitle'),
    printMeta: document.getElementById('printMeta'),
    printTableBody: document.getElementById('printTableBody'),

    // Unified Left Slide-In Drawer
    appDrawerBackdrop: document.getElementById('appDrawerBackdrop'),
    appDrawer: document.getElementById('appDrawer'),
    closeAppDrawerBtn: document.getElementById('closeAppDrawerBtn'),
    appDrawerBody: document.getElementById('appDrawerBody'),
    drawerFilterSection: document.getElementById('drawerFilterSection'),

    drawerFilterDept: document.getElementById('drawerFilterDept'),
    drawerCampusChips: document.getElementById('drawerCampusChips'),
    drawerExamDateChips: document.getElementById('drawerExamDateChips'),
    drawerRegDayChips: document.getElementById('drawerRegDayChips'),
    drawerRegPeriodChips: document.getElementById('drawerRegPeriodChips'),
    drawerStatusSegments: document.getElementById('drawerStatusSegments'),
    drawerResetBtn: document.getElementById('drawerResetBtn'),
    drawerApplyBtn: document.getElementById('drawerApplyBtn'),

    // Drawer Management Menu Buttons
    menuExcelUploadBtn: document.getElementById('menuExcelUploadBtn'),
    menuLoadSampleBtn: document.getElementById('menuLoadSampleBtn'),
    menuExportIcsBtn: document.getElementById('menuExportIcsBtn'),
    menuPrintBtn: document.getElementById('menuPrintBtn'),
    menuClearPinsBtn: document.getElementById('menuClearPinsBtn'),
    menuResetDbBtn: document.getElementById('menuResetDbBtn'),

    // Unsaved Changes Modal
    unsavedModalBackdrop: document.getElementById('unsavedModalBackdrop'),
    unsavedConfirmModal: document.getElementById('unsavedConfirmModal'),
    unsavedApplyBtn: document.getElementById('unsavedApplyBtn'),
    unsavedDiscardBtn: document.getElementById('unsavedDiscardBtn'),

    // Destructive / Overwrite Confirmation Modal
    actionModalBackdrop: document.getElementById('actionModalBackdrop'),
    actionConfirmModal: document.getElementById('actionConfirmModal'),
    actionModalTitle: document.getElementById('actionModalTitle'),
    actionModalDesc: document.getElementById('actionModalDesc'),
    actionModalConfirmBtn: document.getElementById('actionModalConfirmBtn'),
    actionModalCancelBtn: document.getElementById('actionModalCancelBtn'),

    // Subject Detail Modal
    detailModalBackdrop: document.getElementById('detailModalBackdrop'),
    detailModal: document.getElementById('detailModal'),
    closeDetailModalBtn: document.getElementById('closeDetailModalBtn'),
    detailModalTitle: document.getElementById('detailModalTitle'),
    detailModalBody: document.getElementById('detailModalBody'),
    detailModalPinBtn: document.getElementById('detailModalPinBtn'),
    detailModalCloseFooterBtn: document.getElementById('detailModalCloseFooterBtn'),

    toast: document.getElementById('toast')
  };

  // =========================================================================
  // 5. UI Helpers: Toast, Errors & Real-Processing Skeleton Controller
  // =========================================================================
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

  let activeLoadingTokens = 0;
  function beginSkeletonLoading() {
    activeLoadingTokens += 1;
    dom.scheduleListSection.setAttribute('aria-busy', 'true');
    dom.skeletonGrid.classList.remove('hidden');
    dom.cardListContainer.classList.add('hidden');
    dom.emptyState.classList.add('hidden');
  }

  function endSkeletonLoading() {
    activeLoadingTokens = Math.max(0, activeLoadingTokens - 1);
    if (activeLoadingTokens === 0) {
      dom.scheduleListSection.setAttribute('aria-busy', 'false');
      dom.skeletonGrid.classList.add('hidden');
      dom.cardListContainer.classList.remove('hidden');
    }
  }

  function syncBodyScrollLock() {
    const hasOpenOverlay =
        state.isDrawerOpen ||
        state.isUnsavedModalOpen ||
        state.isActionModalOpen ||
        Boolean(state.activeDetailRecord);
    document.body.classList.toggle('overlay-open', hasOpenOverlay);
  }

  // =========================================================================
  // 6. Focus Trap Utility for Drawer & Modals
  // =========================================================================
  function getFocusableElements(container) {
    if (!container) return [];
    const selector =
        'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';
    return Array.from(container.querySelectorAll(selector)).filter(
        (el) => !el.classList.contains('hidden') && el.offsetParent !== null
    );
  }

  function handleContainerTabTrap(e, container) {
    if (e.key !== 'Tab') return;
    const focusables = getFocusableElements(container);
    if (focusables.length === 0) return;

    const first = focusables[0];
    const last = focusables[focusables.length - 1];

    if (e.shiftKey) {
      if (document.activeElement === first || !container.contains(document.activeElement)) {
        e.preventDefault();
        last.focus();
      }
    } else {
      if (document.activeElement === last || !container.contains(document.activeElement)) {
        e.preventDefault();
        first.focus();
      }
    }
  }

  // =========================================================================
  // 7. Excel File Ingestion & Sample Dataset
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
    beginSkeletonLoading();

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
            `「${file.name}」から有効な試験データ行を検出できませんでした。11列構成（No.・学科・時間割コード・曜日・時限・科目名・担当教員・校舎・理解度確認テスト日程・実施教室・追評価の実施）になっているか確認してください。`
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
      showToast(`「${file.name}」から ${parsed.length} 件のデータを読み込みました`);
    } catch (err) {
      console.error(err);
      showError(`ファイル読み込みエラー: ${err?.message || '不明なエラーが発生しました'}`);
    } finally {
      endSkeletonLoading();
      renderCurrentView();
    }
  }

  async function executeLoadSampleDataset() {
    hideError();
    beginSkeletonLoading();
    try {
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
    } finally {
      endSkeletonLoading();
      renderCurrentView();
    }
  }

  function requestLoadSampleDataset() {
    if (state.records.length > 0) {
      const currentName = state.meta?.fileName || '現在の試験データ';
      openActionConfirmModal({
        title: 'サンプルデータで現在のデータを置き換えますか？',
        message: `現在読み込まれている「${currentName}」（${state.records.length}件）が、サンプルデータ（37件）に置き換わります。※マイ試験日程の登録情報は保持されます。`,
        confirmLabel: '置き換えて読み込む',
        confirmVariant: 'primary',
        onConfirm: executeLoadSampleDataset
      });
    } else {
      executeLoadSampleDataset();
    }
  }

  // =========================================================================
  // 8. Build Filter Options & Conflict Detection
  // =========================================================================
  function buildAvailableFilterOptions() {
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

    state.availableOptions.depts = [...depts.entries()]
        .sort((a, b) => JA_COLLATOR.compare(a[0], b[0]))
        .map(([value, count]) => ({ value, count }));

    state.availableOptions.campuses = [...campuses.entries()]
        .sort((a, b) => JA_COLLATOR.compare(a[0], b[0]))
        .map(([value, count]) => ({ value, count }));

    state.availableOptions.examDates = [...examDates.entries()]
        .sort((a, b) => a[1].sortKey - b[1].sortKey)
        .map(([value, info]) => ({ value, count: info.count, sortKey: info.sortKey }));

    state.availableOptions.regPeriods = [...regPeriods].sort((a, b) =>
        JA_COLLATOR.compare(a, b)
    );

    if (state.appliedFilters.dept && !depts.has(state.appliedFilters.dept)) {
      state.appliedFilters.dept = '';
    }
    if (state.appliedFilters.campus && !campuses.has(state.appliedFilters.campus)) {
      state.appliedFilters.campus = '';
    }
    for (const d of [...state.appliedFilters.examDates]) {
      if (!examDates.has(d)) state.appliedFilters.examDates.delete(d);
    }
    for (const p of [...state.appliedFilters.regPeriods]) {
      if (!regPeriods.has(p)) state.appliedFilters.regPeriods.delete(p);
    }
  }

  function detectPinnedConflicts() {
    state.conflictUids.clear();
    const pinnedRecords = state.records.filter(
        (r) => state.pinnedKeys.has(r.stableKey) && r.examMonth !== null
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
  // 9. Unified Left Slide-In Drawer Rendering & State Comparison
  // =========================================================================
  function setsAreEqual(setA, setB) {
    if (setA.size !== setB.size) return false;
    for (const item of setA) {
      if (!setB.has(item)) return false;
    }
    return true;
  }

  function hasUnsavedDrawerChanges() {
    const a = state.appliedFilters;
    const d = state.draftFilters;
    if (a.dept !== d.dept) return true;
    if (a.campus !== d.campus) return true;
    if (a.status !== d.status) return true;
    if (!setsAreEqual(a.examDates, d.examDates)) return true;
    if (!setsAreEqual(a.regDays, d.regDays)) return true;
    if (!setsAreEqual(a.regPeriods, d.regPeriods)) return true;
    return false;
  }

  function renderDrawerControls() {
    const draft = state.draftFilters;

    // 1. 学科・区分 Dropdown
    dom.drawerFilterDept.innerHTML = '<option value="">すべての学科・区分</option>';
    state.availableOptions.depts.forEach(({ value, count }) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = `${value} (${count})`;
      dom.drawerFilterDept.appendChild(opt);
    });
    dom.drawerFilterDept.value = draft.dept;

    // 2. 校舎 Chips (Single-select preserved)
    dom.drawerCampusChips.innerHTML = '';
    if (state.availableOptions.campuses.length === 0) {
      dom.drawerCampusChips.innerHTML = '<span class="drawer-hint">データ読込後に表示されます</span>';
    } else {
      state.availableOptions.campuses.forEach(({ value }) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        const isSelected = draft.campus === value;
        btn.className = `chip-btn${isSelected ? ' selected' : ''}`;
        btn.setAttribute('aria-pressed', isSelected ? 'true' : 'false');
        btn.textContent = value;
        btn.addEventListener('click', () => {
          state.draftFilters.campus = state.draftFilters.campus === value ? '' : value;
          renderDrawerControls();
        });
        dom.drawerCampusChips.appendChild(btn);
      });
    }

    // 3. 試験日 Chips (Multi-select, Unified Size)
    dom.drawerExamDateChips.innerHTML = '';
    if (state.availableOptions.examDates.length === 0) {
      dom.drawerExamDateChips.innerHTML = '<span class="drawer-hint">データ読込後に表示されます</span>';
    } else {
      state.availableOptions.examDates.forEach(({ value }) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        const isSelected = draft.examDates.has(value);
        btn.className = `chip-btn${isSelected ? ' selected' : ''}`;
        btn.setAttribute('aria-pressed', isSelected ? 'true' : 'false');
        btn.textContent = value;
        btn.addEventListener('click', () => {
          if (state.draftFilters.examDates.has(value)) {
            state.draftFilters.examDates.delete(value);
          } else {
            state.draftFilters.examDates.add(value);
          }
          renderDrawerControls();
        });
        dom.drawerExamDateChips.appendChild(btn);
      });
    }

    // 4. 通常授業の曜日 Chips (Multi-select, Mon-Sun)
    dom.drawerRegDayChips.innerHTML = '';
    ALL_WEEKDAYS.forEach((dayChar) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      const isSelected = draft.regDays.has(dayChar);
      btn.className = `chip-btn${isSelected ? ' selected' : ''}`;
      btn.setAttribute('aria-pressed', isSelected ? 'true' : 'false');
      btn.textContent = `${dayChar}曜`;
      btn.addEventListener('click', () => {
        if (state.draftFilters.regDays.has(dayChar)) {
          state.draftFilters.regDays.delete(dayChar);
        } else {
          state.draftFilters.regDays.add(dayChar);
        }
        renderDrawerControls();
      });
      dom.drawerRegDayChips.appendChild(btn);
    });

    // 5. 通常授業の時限 Chips (Multi-select, Unified Size)
    dom.drawerRegPeriodChips.innerHTML = '';
    if (state.availableOptions.regPeriods.length === 0) {
      dom.drawerRegPeriodChips.innerHTML = '<span class="drawer-hint">データ読込後に表示されます</span>';
    } else {
      state.availableOptions.regPeriods.forEach((periodVal) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        const isSelected = draft.regPeriods.has(periodVal);
        btn.className = `chip-btn${isSelected ? ' selected' : ''}`;
        btn.setAttribute('aria-pressed', isSelected ? 'true' : 'false');
        btn.textContent = `${periodVal}時限`;
        btn.addEventListener('click', () => {
          if (state.draftFilters.regPeriods.has(periodVal)) {
            state.draftFilters.regPeriods.delete(periodVal);
          } else {
            state.draftFilters.regPeriods.add(periodVal);
          }
          renderDrawerControls();
        });
        dom.drawerRegPeriodChips.appendChild(btn);
      });
    }

    // 6. 追評価の実施有無 Segmented Control
    const segButtons = dom.drawerStatusSegments.querySelectorAll('.segment-btn');
    segButtons.forEach((btn) => {
      const val = btn.dataset.value || '';
      const isActive = draft.status === val;
      btn.classList.toggle('active', isActive);
      btn.setAttribute('aria-checked', isActive ? 'true' : 'false');
    });
  }

  function openAppDrawer(triggerEl) {
    state.lastFocusedElement = triggerEl || document.activeElement;
    state.draftFilters = cloneFilters(state.appliedFilters);
    renderDrawerControls();

    state.isDrawerOpen = true;
    dom.openMenuDrawerBtn.setAttribute('aria-expanded', 'true');
    dom.appDrawerBackdrop.classList.remove('hidden');
    dom.appDrawer.classList.remove('hidden');
    syncBodyScrollLock();

    requestAnimationFrame(() => {
      dom.appDrawer.classList.add('is-open');
      dom.appDrawerBody.scrollTop = 0;
      dom.closeAppDrawerBtn.focus();
    });
  }

  function closeAppDrawerImmediate() {
    state.isDrawerOpen = false;
    dom.openMenuDrawerBtn.setAttribute('aria-expanded', 'false');
    dom.appDrawer.classList.remove('is-open');
    dom.appDrawerBackdrop.classList.add('hidden');
    dom.appDrawer.classList.add('hidden');
    syncBodyScrollLock();

    if (state.lastFocusedElement && typeof state.lastFocusedElement.focus === 'function') {
      state.lastFocusedElement.focus();
    }
  }

  function requestCloseAppDrawer() {
    if (hasUnsavedDrawerChanges()) {
      openUnsavedConfirmModal();
    } else {
      closeAppDrawerImmediate();
    }
  }

  function applyDraftFiltersAndClose() {
    state.appliedFilters = cloneFilters(state.draftFilters);
    closeUnsavedConfirmModal();
    closeAppDrawerImmediate();
    renderCurrentViewWithThresholdSkeleton();
  }

  function discardDraftFiltersAndClose() {
    state.draftFilters = cloneFilters(state.appliedFilters);
    closeUnsavedConfirmModal();
    closeAppDrawerImmediate();
  }

  // =========================================================================
  // 10. Confirmation Modals (Unsaved Filters & Destructive/Overwrite Actions)
  // =========================================================================
  function openUnsavedConfirmModal() {
    state.isUnsavedModalOpen = true;
    dom.unsavedModalBackdrop.classList.remove('hidden');
    dom.unsavedConfirmModal.classList.remove('hidden');
    syncBodyScrollLock();
    dom.unsavedApplyBtn.focus();
  }

  function closeUnsavedConfirmModal() {
    state.isUnsavedModalOpen = false;
    dom.unsavedModalBackdrop.classList.add('hidden');
    dom.unsavedConfirmModal.classList.add('hidden');
    syncBodyScrollLock();
    if (state.isDrawerOpen) {
      dom.closeAppDrawerBtn.focus();
    }
  }

  function openActionConfirmModal({ title, message, confirmLabel, confirmVariant = 'danger', onConfirm }) {
    state.pendingConfirmAction = onConfirm;
    state.isActionModalOpen = true;
    dom.actionModalTitle.textContent = title;
    dom.actionModalDesc.textContent = message;
    dom.actionModalConfirmBtn.textContent = confirmLabel || '実行する';
    dom.actionModalConfirmBtn.className =
        confirmVariant === 'primary' ? 'btn btn-primary btn-block' : 'btn btn-danger btn-block';

    dom.actionModalBackdrop.classList.remove('hidden');
    dom.actionConfirmModal.classList.remove('hidden');
    syncBodyScrollLock();
    dom.actionModalCancelBtn.focus();
  }

  function closeActionConfirmModal() {
    state.isActionModalOpen = false;
    state.pendingConfirmAction = null;
    dom.actionModalBackdrop.classList.add('hidden');
    dom.actionConfirmModal.classList.add('hidden');
    syncBodyScrollLock();
    if (state.isDrawerOpen) {
      dom.closeAppDrawerBtn.focus();
    } else if (state.lastFocusedElement && typeof state.lastFocusedElement.focus === 'function') {
      state.lastFocusedElement.focus();
    }
  }

  // =========================================================================
  // 11. Active Filter Summary Tags (Outside Drawer)
  // =========================================================================
  function renderActiveFilterSummary() {
    const f = state.appliedFilters;
    const tags = [];

    if (f.dept) {
      tags.push({
        label: `学科: ${f.dept}`,
        onRemove: () => {
          state.appliedFilters.dept = '';
        }
      });
    }
    if (f.campus) {
      tags.push({
        label: `校舎: ${f.campus}`,
        onRemove: () => {
          state.appliedFilters.campus = '';
        }
      });
    }
    f.examDates.forEach((d) => {
      tags.push({
        label: `試験日: ${d}`,
        onRemove: () => {
          state.appliedFilters.examDates.delete(d);
        }
      });
    });
    f.regDays.forEach((day) => {
      tags.push({
        label: `曜日: ${day}曜`,
        onRemove: () => {
          state.appliedFilters.regDays.delete(day);
        }
      });
    });
    f.regPeriods.forEach((p) => {
      tags.push({
        label: `時限: ${p}時限`,
        onRemove: () => {
          state.appliedFilters.regPeriods.delete(p);
        }
      });
    });
    if (f.status) {
      tags.push({
        label: `追評価: ${f.status}`,
        onRemove: () => {
          state.appliedFilters.status = '';
        }
      });
    }

    if (tags.length > 0) {
      dom.activeFilterBadge.textContent = String(tags.length);
      dom.activeFilterBadge.classList.remove('hidden');
      dom.activeFiltersBar.classList.remove('hidden');
    } else {
      dom.activeFilterBadge.classList.add('hidden');
      dom.activeFiltersBar.classList.add('hidden');
    }

    dom.activeFilterTags.innerHTML = '';
    const fragment = document.createDocumentFragment();
    tags.forEach((tag) => {
      const span = document.createElement('span');
      span.className = 'active-tag';
      span.textContent = tag.label;

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'active-tag-remove';
      btn.setAttribute('aria-label', `${tag.label} を解除`);
      btn.textContent = '×';
      btn.addEventListener('click', () => {
        tag.onRemove();
        renderCurrentViewWithThresholdSkeleton();
      });

      span.appendChild(btn);
      fragment.appendChild(span);
    });
    dom.activeFilterTags.appendChild(fragment);
  }

  // =========================================================================
  // 12. Filtering & Sorting Pipeline (Non-Destructive, Japanese Gojuon Order)
  // =========================================================================
  function getFilteredAndSortedRecords() {
    const f = state.appliedFilters;
    const qNorm = normalizeText(f.query);

    const filtered = state.records.filter((rec) => {
      if (state.viewMode === 'pinned' && !state.pinnedKeys.has(rec.stableKey)) {
        return false;
      }
      if (f.dept && rec.department !== f.dept) {
        return false;
      }
      if (f.campus && rec.campus !== f.campus) {
        return false;
      }
      if (f.examDates.size > 0 && !f.examDates.has(rec.examDateLabel)) {
        return false;
      }
      if (f.regDays.size > 0 && !f.regDays.has(rec.regularDay)) {
        return false;
      }
      if (f.regPeriods.size > 0 && !f.regPeriods.has(rec.regularPeriod)) {
        return false;
      }
      if (f.status && rec.status !== f.status) {
        return false;
      }
      if (qNorm && !rec.searchBlob.includes(qNorm)) {
        return false;
      }
      return true;
    });

    const [field, dir] = state.sortBy.split('_');
    const sign = dir === 'desc' ? -1 : 1;

    // Sort a shallow-filtered array without mutating original state.records
    filtered.sort((a, b) => {
      let cmp = 0;
      switch (field) {
        case 'subject': {
          const keyA = a.subjectSortKey || makeSubjectSortKey(a.subject);
          const keyB = b.subjectSortKey || makeSubjectSortKey(b.subject);
          cmp = JA_COLLATOR.compare(keyA, keyB);
          break;
        }
        case 'examDate':
          cmp = a.examSortKey - b.examSortKey;
          break;
        case 'dept':
          cmp = JA_COLLATOR.compare(a.department, b.department) || JA_COLLATOR.compare(a.classCode, b.classCode);
          break;
        case 'regular':
          cmp = a.regularSortKey - b.regularSortKey;
          break;
        case 'room':
          cmp = JA_COLLATOR.compare(a.room, b.room);
          break;
        case 'no':
        default:
          cmp = a.noSort - b.noSort;
          break;
      }
      // Stable tie-breaker using original Excel row order
      if (cmp === 0) {
        return a.noSort - b.noSort;
      }
      return cmp * sign;
    });

    return filtered;
  }

  // =========================================================================
  // 13. Progressive Card Rendering, Differential Pin Updates & On-Demand Print
  // =========================================================================
  function escapeHtml(str) {
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
  }

  function createStarSvg(filled) {
    return filled
        ? `<svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>`
        : `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>`;
  }

  function createExamCardElement(rec) {
    const isPinned = state.pinnedKeys.has(rec.stableKey);
    const isMakeupCancelled = rec.status === '実施しない';
    const isConflict = state.conflictUids.has(rec.uid);

    const card = document.createElement('article');
    card.className = 'exam-card';
    card.dataset.uid = rec.uid;
    card.dataset.stableKey = rec.stableKey;
    card.tabIndex = 0;
    card.setAttribute('role', 'button');
    card.setAttribute('aria-label', `${rec.subject} の詳細を表示`);

    if (isPinned) card.classList.add('card-pinned');
    if (isMakeupCancelled) card.classList.add('card-cancelled');
    if (isConflict) card.classList.add('card-conflict');

    const scheduleDisplay = rec.examRaw || '試験日時未定・記載なし';
    const roomLocationDisplay =
        rec.campus && rec.room
            ? `${rec.campus} / ${rec.room}`
            : rec.room || rec.campus || '未定・記載なし';

    const makeupStatusTag = isMakeupCancelled
        ? `<span class="card-status-note">追評価: 実施しない</span>`
        : '';

    const conflictHtml = isConflict
        ? `<div class="card-conflict-notice" role="status">
           <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
             <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
             <line x1="12" y1="9" x2="12" y2="13"></line>
             <line x1="12" y1="17" x2="12.01" y2="17"></line>
           </svg>
           <span>試験日時が他の登録科目と重複しています</span>
         </div>`
        : '';

    card.innerHTML = `
      <div class="card-header-row">
        <div class="card-datetime-block">
          <span class="card-datetime">${escapeHtml(scheduleDisplay)}</span>
          ${makeupStatusTag}
        </div>
        <button
          type="button"
          class="pin-btn ${isPinned ? 'pinned' : ''}"
          data-key="${escapeHtml(rec.stableKey)}"
          aria-pressed="${isPinned ? 'true' : 'false'}"
          aria-label="${isPinned ? 'マイ試験日程から解除' : 'マイ試験日程に登録'}"
          title="${isPinned ? 'マイ試験日程から解除' : 'マイ試験日程に登録'}"
        >
          ${createStarSvg(isPinned)}
        </button>
      </div>

      <h3 class="card-subject">${escapeHtml(rec.subject || '科目名未記載')}</h3>

      ${conflictHtml}

      <div class="card-details-list">
        <div class="card-detail-row">
          <span class="card-detail-label">試験教室:</span>
          <span class="card-detail-value">${escapeHtml(roomLocationDisplay)}</span>
        </div>
        <div class="card-detail-row">
          <span class="card-detail-label">担当教員:</span>
          <span class="card-detail-teacher">${escapeHtml(rec.instructor || '未記載')}</span>
        </div>
      </div>
    `;

    const pinBtn = card.querySelector('.pin-btn');
    pinBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      togglePin(rec.stableKey);
    });

    card.addEventListener('click', () => {
      openDetailModal(rec, card);
    });

    card.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openDetailModal(rec, card);
      }
    });

    return card;
  }

  function appendNextCardChunk() {
    const records = state.currentRenderedRecords;
    if (state.renderedCount >= records.length) {
      dom.loadMoreSentinel.classList.add('hidden');
      return;
    }

    const slice = records.slice(state.renderedCount, state.renderedCount + RENDER_CHUNK_SIZE);
    const isDateSorted = state.sortBy === 'examDate_asc' || state.sortBy === 'examDate_desc';
    const fragment = document.createDocumentFragment();

    if (isDateSorted) {
      slice.forEach((rec) => {
        const dateLabel = rec.examDateLabel;
        if (state.lastDateGroupKey !== dateLabel || !state.lastDateGroupSection) {
          const totalInGroup = records.filter((r) => r.examDateLabel === dateLabel).length;
          const section = document.createElement('section');
          section.className = 'date-group-section';

          const header = document.createElement('div');
          header.className = 'date-group-header';
          header.innerHTML = `
            <h2 class="date-group-title">${escapeHtml(dateLabel)}</h2>
            <span class="date-group-count">${totalInGroup} 件</span>
          `;

          const grid = document.createElement('div');
          grid.className = 'card-grid';
          section.appendChild(header);
          section.appendChild(grid);

          fragment.appendChild(section);
          state.lastDateGroupSection = grid;
          state.lastDateGroupKey = dateLabel;
        }
        state.lastDateGroupSection.appendChild(createExamCardElement(rec));
      });
    } else {
      let grid = dom.cardListContainer.querySelector(':scope > .card-grid');
      if (!grid) {
        grid = document.createElement('div');
        grid.className = 'card-grid';
        fragment.appendChild(grid);
      }
      slice.forEach((rec) => {
        grid.appendChild(createExamCardElement(rec));
      });
    }

    dom.cardListContainer.appendChild(fragment);
    state.renderedCount += slice.length;

    if (state.renderedCount < records.length) {
      dom.loadMoreSentinel.classList.remove('hidden');
    } else {
      dom.loadMoreSentinel.classList.add('hidden');
    }
  }

  function renderCards(records) {
    dom.cardListContainer.innerHTML = '';
    state.currentRenderedRecords = records;
    state.renderedCount = 0;
    state.lastDateGroupSection = null;
    state.lastDateGroupKey = null;

    if (records.length === 0) {
      dom.loadMoreSentinel.classList.add('hidden');
      dom.emptyState.classList.remove('hidden');
      if (state.records.length === 0) {
        dom.emptyTitle.textContent = 'データがありません';
        dom.emptyDesc.textContent = '上部のエリアからExcelファイルを読み込むか、「サンプルデータを読み込む」を押してください。';
      } else if (state.viewMode === 'pinned') {
        dom.emptyTitle.textContent = 'マイ試験日程に登録された科目がありません';
        dom.emptyDesc.textContent = '「全科目」タブで履修している科目の星アイコンを押すと登録できます。';
      } else {
        dom.emptyTitle.textContent = '条件に一致する試験が見つかりませんでした';
        dom.emptyDesc.textContent = '検索キーワードやメニュー内の絞り込み条件を変更してみてください。';
      }
      return;
    }

    dom.emptyState.classList.add('hidden');
    appendNextCardChunk();
  }

  /**
   * 印刷・PDF出力時のみオンデマンドでテーブルを生成（通常スクロール時のDOM負荷を削減）
   */
  function preparePrintTable() {
    const records = state.currentRenderedRecords.length > 0
        ? state.currentRenderedRecords
        : getFilteredAndSortedRecords();

    const modeLabel = state.viewMode === 'pinned' ? 'マイ定期試験日程表' : '定期試験日程一覧';
    dom.printTitle.textContent = modeLabel;
    dom.printMeta.textContent = `出力日時: ${new Date().toLocaleString('ja-JP')} ／ 表示件数: ${records.length}件`;

    dom.printTableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();

    records.forEach((rec) => {
      const isPinned = state.pinnedKeys.has(rec.stableKey);
      const tr = document.createElement('tr');
      if (isPinned) tr.classList.add('row-pinned');

      tr.innerHTML = `
        <td>${escapeHtml(rec.no)}</td>
        <td>${escapeHtml(rec.department)}</td>
        <td>${escapeHtml(rec.classCode)}</td>
        <td>${escapeHtml(rec.regularDisplay)}</td>
        <td>${escapeHtml(rec.subject)}</td>
        <td>${escapeHtml(rec.instructor)}</td>
        <td>${escapeHtml(rec.campus)}</td>
        <td>${escapeHtml(rec.examRaw || '-')}</td>
        <td>${escapeHtml(rec.room || '-')}</td>
        <td>${escapeHtml(rec.status)}</td>
      `;
      fragment.appendChild(tr);
    });

    dom.printTableBody.appendChild(fragment);
  }

  function updateHeaderAndCounts() {
    if (state.meta && state.records.length > 0) {
      dom.dbMetaText.textContent = `読込済: ${state.meta.fileName}（全 ${state.records.length} 件 / 更新: ${state.meta.updatedAt}）`;
    } else {
      dom.dbMetaText.textContent = 'Excelファイルを読み込んでデータベース化します（ブラウザ内ローカル保存）';
    }

    dom.countAll.textContent = String(state.records.length);
    const activePinnedCount = state.records.filter((r) => state.pinnedKeys.has(r.stableKey)).length;
    dom.countPinned.textContent = String(activePinnedCount);
  }

  function renderCurrentView() {
    const filtered = getFilteredAndSortedRecords();
    dom.resultStats.innerHTML = `表示件数: <strong>${filtered.length}</strong> / 全 ${state.records.length} 件`;
    renderActiveFilterSummary();
    renderCards(filtered);
  }

  /**
   * 処理時間が閾値（80ms）を超える規模の場合のみスケルトンを表示し、
   * 瞬時に完了する通常操作ではチラつき・待ち時間を発生させない
   */
  function renderCurrentViewWithThresholdSkeleton() {
    if (state.records.length >= 400) {
      beginSkeletonLoading();
      requestAnimationFrame(() => {
        renderCurrentView();
        endSkeletonLoading();
      });
    } else {
      renderCurrentView();
    }
  }

  function refreshAllUI() {
    buildAvailableFilterOptions();
    detectPinnedConflicts();
    updateHeaderAndCounts();
    renderCurrentView();
  }

  // =========================================================================
  // 14. Subject Detail Modal (All 10 Record Fields)
  // =========================================================================
  function syncDetailModalPinButton() {
    const rec = state.activeDetailRecord;
    if (!rec) return;
    const isPinned = state.pinnedKeys.has(rec.stableKey);
    dom.detailModalPinBtn.textContent = isPinned
        ? 'マイ試験日程に登録済み（解除する）'
        : 'マイ試験日程に登録する';
    dom.detailModalPinBtn.className = isPinned ? 'btn btn-ghost' : 'btn btn-secondary';
  }

  function openDetailModal(rec, triggerEl) {
    state.activeDetailRecord = rec;
    state.lastFocusedElement = triggerEl || document.activeElement;

    dom.detailModalTitle.textContent = rec.subject || '科目名未記載';

    const isMakeupActive = rec.status !== '実施しない';
    const statusBadgeClass = isMakeupActive ? 'status-pill-active' : 'status-pill-inactive';

    const detailItems = [
      {
        label: '追評価の実施有無',
        html: `<span class="status-pill-inline ${statusBadgeClass}">${escapeHtml(rec.status || '未記載')}</span>`
      },
      {
        label: '理解度確認テスト日程',
        value: rec.examRaw || '未定・記載なし'
      },
      {
        label: '校舎',
        value: rec.campus || '未記載'
      },
      {
        label: '理解度確認テスト実施教室',
        value: rec.room || '未定・記載なし'
      },
      {
        label: '担当教員',
        value: rec.instructor || '未記載'
      },
      {
        label: '学科・区分',
        value: rec.department || '未記載'
      },
      {
        label: '時間割コード',
        value: rec.classCode || '未記載'
      },
      {
        label: '通常授業の曜日・時限',
        value: rec.regularDisplay || '未記載'
      },
      {
        label: '通し番号（No.）',
        value: rec.no !== '' && rec.no !== null && rec.no !== undefined ? String(rec.no) : '未記載'
      }
    ];

    const dl = document.createElement('dl');
    dl.className = 'detail-dl';

    detailItems.forEach((item) => {
      const wrapper = document.createElement('div');
      wrapper.className = 'detail-item';

      const dt = document.createElement('dt');
      dt.className = 'detail-dt';
      dt.textContent = item.label;

      const dd = document.createElement('dd');
      dd.className = 'detail-dd';
      if (item.html) {
        dd.innerHTML = item.html;
      } else {
        dd.textContent = item.value;
      }

      wrapper.appendChild(dt);
      wrapper.appendChild(dd);
      dl.appendChild(wrapper);
    });

    dom.detailModalBody.innerHTML = '';
    dom.detailModalBody.appendChild(dl);

    syncDetailModalPinButton();

    dom.detailModalBackdrop.classList.remove('hidden');
    dom.detailModal.classList.remove('hidden');
    syncBodyScrollLock();

    dom.closeDetailModalBtn.focus();
  }

  function closeDetailModal() {
    state.activeDetailRecord = null;
    dom.detailModalBackdrop.classList.add('hidden');
    dom.detailModal.classList.add('hidden');
    syncBodyScrollLock();

    if (state.lastFocusedElement && typeof state.lastFocusedElement.focus === 'function') {
      state.lastFocusedElement.focus();
    }
  }

  // =========================================================================
  // 15. Differential Pinning & ICS Calendar Export
  // =========================================================================
  function syncRenderedCardPinAndConflictStates() {
    const cards = dom.cardListContainer.querySelectorAll('.exam-card');
    cards.forEach((card) => {
      const uid = card.dataset.uid;
      const key = card.dataset.stableKey;
      const isPinned = state.pinnedKeys.has(key);
      const isConflict = state.conflictUids.has(uid);

      card.classList.toggle('card-pinned', isPinned);
      card.classList.toggle('card-conflict', isConflict);

      const pinBtn = card.querySelector('.pin-btn');
      if (pinBtn) {
        pinBtn.classList.toggle('pinned', isPinned);
        pinBtn.setAttribute('aria-pressed', isPinned ? 'true' : 'false');
        pinBtn.setAttribute('aria-label', isPinned ? 'マイ試験日程から解除' : 'マイ試験日程に登録');
        pinBtn.title = isPinned ? 'マイ試験日程から解除' : 'マイ試験日程に登録';
        pinBtn.innerHTML = createStarSvg(isPinned);
      }

      const existingNotice = card.querySelector('.card-conflict-notice');
      if (isConflict && !existingNotice) {
        const notice = document.createElement('div');
        notice.className = 'card-conflict-notice';
        notice.setAttribute('role', 'status');
        notice.innerHTML = `
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
            <line x1="12" y1="9" x2="12" y2="13"></line>
            <line x1="12" y1="17" x2="12.01" y2="17"></line>
          </svg>
          <span>試験日時が他の登録科目と重複しています</span>
        `;
        const detailsList = card.querySelector('.card-details-list');
        if (detailsList) {
          card.insertBefore(notice, detailsList);
        }
      } else if (!isConflict && existingNotice) {
        existingNotice.remove();
      }
    });
  }

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
    syncDetailModalPinButton();

    if (state.viewMode === 'pinned') {
      // Remove unpinned item cleanly in My Exam Schedule view
      renderCurrentView();
    } else {
      // Differential DOM update without rebuilding list or resetting scroll
      syncRenderedCardPinAndConflictStates();
    }
  }

  function exportPinnedToICS() {
    const pinned = state.records.filter(
        (r) => state.pinnedKeys.has(r.stableKey) && r.examMonth !== null && r.examDay !== null
    );

    if (pinned.length === 0) {
      showToast('書き出し可能な「マイ試験日程」がありません');
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
      lines.push(`DESCRIPTION:時間割コード: ${rec.classCode} / 担当教員: ${rec.instructor} / 追評価: ${rec.status}`);
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
  // 16. Event Listeners & Progressive Scroll Observer
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

  function setupProgressiveScrollObserver() {
    if (!('IntersectionObserver' in window)) return;
    const observer = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.isIntersecting && state.renderedCount < state.currentRenderedRecords.length) {
              appendNextCardChunk();
            }
          });
        },
        { rootMargin: '320px 0px' }
    );
    observer.observe(dom.loadMoreSentinel);
  }

  function bindEvents() {
    dom.closeErrorBtn.addEventListener('click', hideError);

    // File input change
    dom.fileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) {
        handleExcelFile(file);
        e.target.value = '';
      }
    });

    // Inline Sample button inside initial dropzone
    dom.inlineSampleBtn.addEventListener('click', requestLoadSampleDataset);

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

    // Prepare print table right before browser print
    window.addEventListener('beforeprint', preparePrintTable);

    // Two View Tabs (全科目 / マイ試験日程)
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

    // Search input with real-time debounce (keeps existing list visible to avoid layout jumps)
    let searchDebounce = null;
    dom.searchInput.addEventListener('input', (e) => {
      const val = e.target.value;
      dom.clearSearchBtn.classList.toggle('hidden', val.length === 0);

      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(() => {
        state.appliedFilters.query = val;
        renderCurrentViewWithThresholdSkeleton();
      }, 150);
    });

    dom.clearSearchBtn.addEventListener('click', () => {
      clearTimeout(searchDebounce);
      dom.searchInput.value = '';
      dom.clearSearchBtn.classList.add('hidden');
      state.appliedFilters.query = '';
      renderCurrentView();
      dom.searchInput.focus();
    });

    // Quick Clear All Filter Chips/Dropdowns outside drawer
    dom.quickClearAllFiltersBtn.addEventListener('click', () => {
      const currentQuery = state.appliedFilters.query;
      state.appliedFilters = createEmptyFilters();
      state.appliedFilters.query = currentQuery;
      renderCurrentViewWithThresholdSkeleton();
      showToast('絞り込み条件をすべて解除しました');
    });

    // Open Unified Drawer triggers
    dom.openMenuDrawerBtn.addEventListener('click', () => openAppDrawer(dom.openMenuDrawerBtn));
    dom.guideOpenMenuBtn.addEventListener('click', () => openAppDrawer(dom.guideOpenMenuBtn));
    dom.editFiltersFromBarBtn.addEventListener('click', () => openAppDrawer(dom.editFiltersFromBarBtn));

    // Sort select
    dom.sortSelect.addEventListener('change', () => {
      state.sortBy = dom.sortSelect.value;
      renderCurrentView();
    });

    // Unified Drawer Controls
    dom.closeAppDrawerBtn.addEventListener('click', requestCloseAppDrawer);
    dom.appDrawerBackdrop.addEventListener('click', requestCloseAppDrawer);

    dom.drawerFilterDept.addEventListener('change', () => {
      state.draftFilters.dept = dom.drawerFilterDept.value;
    });

    dom.drawerStatusSegments.querySelectorAll('.segment-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.draftFilters.status = btn.dataset.value || '';
        renderDrawerControls();
      });
    });

    dom.drawerResetBtn.addEventListener('click', () => {
      const keepQuery = state.draftFilters.query;
      state.draftFilters = createEmptyFilters();
      state.draftFilters.query = keepQuery;
      renderDrawerControls();
    });

    dom.drawerApplyBtn.addEventListener('click', applyDraftFiltersAndClose);

    // Drawer Management Menu Operations
    dom.menuExcelUploadBtn.addEventListener('click', () => {
      closeAppDrawerImmediate();
      dom.uploadPanel.classList.remove('hidden');
      dom.uploadPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      const fileLabel = dom.uploadPanel.querySelector('.file-select-label');
      if (fileLabel) fileLabel.focus();
    });

    dom.menuLoadSampleBtn.addEventListener('click', () => {
      closeAppDrawerImmediate();
      requestLoadSampleDataset();
    });

    dom.menuExportIcsBtn.addEventListener('click', () => {
      closeAppDrawerImmediate();
      exportPinnedToICS();
    });

    dom.menuPrintBtn.addEventListener('click', () => {
      closeAppDrawerImmediate();
      preparePrintTable();
      requestAnimationFrame(() => {
        window.print();
      });
    });

    dom.menuClearPinsBtn.addEventListener('click', () => {
      const pinnedCount = state.records.filter((r) => state.pinnedKeys.has(r.stableKey)).length;
      if (pinnedCount === 0) {
        showToast('解除するマイ試験日程が登録されていません');
        return;
      }
      openActionConfirmModal({
        title: 'マイ試験日程をすべて解除しますか？',
        message: `登録済みのマイ試験日程（${pinnedCount}件）がすべて解除されます。読み込まれている試験データ自体は削除されません。`,
        confirmLabel: 'すべて解除する',
        confirmVariant: 'danger',
        onConfirm: async () => {
          state.pinnedKeys.clear();
          await StorageService.set('pinnedKeys', []);
          detectPinnedConflicts();
          updateHeaderAndCounts();
          syncDetailModalPinButton();
          renderCurrentView();
          closeAppDrawerImmediate();
          showToast('マイ試験日程の登録をすべて解除しました');
        }
      });
    });

    dom.menuResetDbBtn.addEventListener('click', () => {
      const totalCount = state.records.length;
      const pinnedCount = state.pinnedKeys.size;
      openActionConfirmModal({
        title: '保存データをすべて削除しますか？',
        message: `ブラウザ内に保存されている試験日程データ（${totalCount}件）とマイ試験日程の登録（${pinnedCount}件）がすべて削除されます。この操作は取り消せません。`,
        confirmLabel: 'すべて削除する',
        confirmVariant: 'danger',
        onConfirm: async () => {
          await StorageService.clearAll();
          state.records = [];
          state.meta = null;
          state.pinnedKeys.clear();
          state.appliedFilters = createEmptyFilters();
          state.draftFilters = createEmptyFilters();
          dom.searchInput.value = '';
          dom.clearSearchBtn.classList.add('hidden');
          dom.uploadPanel.classList.remove('hidden');
          refreshAllUI();
          closeAppDrawerImmediate();
          showToast('保存データをすべて削除しました');
        }
      });
    });

    // Unsaved Changes Confirmation Modal Actions
    dom.unsavedApplyBtn.addEventListener('click', applyDraftFiltersAndClose);
    dom.unsavedDiscardBtn.addEventListener('click', discardDraftFiltersAndClose);
    dom.unsavedModalBackdrop.addEventListener('click', closeUnsavedConfirmModal);

    // Destructive / Overwrite Action Confirmation Modal Actions
    dom.actionModalCancelBtn.addEventListener('click', closeActionConfirmModal);
    dom.actionModalBackdrop.addEventListener('click', closeActionConfirmModal);
    dom.actionModalConfirmBtn.addEventListener('click', async () => {
      const fn = state.pendingConfirmAction;
      closeActionConfirmModal();
      if (typeof fn === 'function') {
        await fn();
      }
    });

    // Subject Detail Modal Actions
    dom.closeDetailModalBtn.addEventListener('click', closeDetailModal);
    dom.detailModalCloseFooterBtn.addEventListener('click', closeDetailModal);
    dom.detailModalBackdrop.addEventListener('click', closeDetailModal);
    dom.detailModalPinBtn.addEventListener('click', () => {
      if (state.activeDetailRecord) {
        togglePin(state.activeDetailRecord.stableKey);
      }
    });

    // Global Keyboard Navigation (Esc & Focus Traps)
    document.addEventListener('keydown', (e) => {
      if (state.isActionModalOpen) {
        if (e.key === 'Escape') {
          e.preventDefault();
          closeActionConfirmModal();
        } else {
          handleContainerTabTrap(e, dom.actionConfirmModal);
        }
        return;
      }

      if (state.isUnsavedModalOpen) {
        if (e.key === 'Escape') {
          e.preventDefault();
          closeUnsavedConfirmModal();
        } else {
          handleContainerTabTrap(e, dom.unsavedConfirmModal);
        }
        return;
      }

      if (state.activeDetailRecord) {
        if (e.key === 'Escape') {
          e.preventDefault();
          closeDetailModal();
        } else {
          handleContainerTabTrap(e, dom.detailModal);
        }
        return;
      }

      if (state.isDrawerOpen) {
        if (e.key === 'Escape') {
          e.preventDefault();
          requestCloseAppDrawer();
        } else {
          handleContainerTabTrap(e, dom.appDrawer);
        }
      }
    });
  }

  // =========================================================================
  // 17. Initialization
  // =========================================================================
  async function init() {
    bindEvents();
    setupProgressiveScrollObserver();

    let initLoaded = false;
    const slowStorageTimer = setTimeout(() => {
      if (!initLoaded) {
        beginSkeletonLoading();
      }
    }, 120);

    try {
      const [savedRecords, savedMeta, savedPins] = await Promise.all([
        StorageService.get('records'),
        StorageService.get('meta'),
        StorageService.get('pinnedKeys')
      ]);

      if (Array.isArray(savedPins)) {
        state.pinnedKeys = new Set(savedPins);
      }

      if (Array.isArray(savedRecords) && savedRecords.length > 0) {
        // Ensure backward compatibility with existing stored records
        state.records = savedRecords.map((rec) => ({
          ...rec,
          subjectSortKey: rec.subjectSortKey || makeSubjectSortKey(rec.subject)
        }));
        state.meta = savedMeta;
        dom.uploadPanel.classList.add('hidden');
      } else {
        dom.uploadPanel.classList.remove('hidden');
      }
    } finally {
      const skeletonWasShown = !initLoaded && activeLoadingTokens > 0;
      initLoaded = true;
      clearTimeout(slowStorageTimer);
      refreshAllUI();
      if (skeletonWasShown) {
        endSkeletonLoading();
      }
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();