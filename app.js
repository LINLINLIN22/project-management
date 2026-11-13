/**
 * IssueDesk - Pure Vanilla JavaScript Application Controller
 * Zero frameworks, zero third-party packages, zero API keys.
 * Works both on GitHub Pages / static HTTP servers and locally via file://.
 */

(function () {
  'use strict';

  const { buildXlsxBuffer, downloadFile, parseXlsx, computeCRC32, excelSerialToIso } = window.XlsxEngine;

  // ==========================================================================
  // Constants & Sample Data
  // ==========================================================================

  const STORAGE_KEY = 'issuedesk_vanilla_data_v2';
  const AUTO_FILE_PATHS = ['./data/issues.xlsx', './issues.xlsx'];

  const INITIAL_PROJECTS = [
    {
      id: 'proj-web',
      name: '電商官網改版 (Web Platform)',
      description: '重構前端購物流程、結帳與會員系統，提升載入速度與轉換率',
      color: '#0969da',
      createdAt: '2026-11-01T09:00:00.000Z',
    },
    {
      id: 'proj-mobile',
      name: '行動 App 2.0 (Mobile App)',
      description: '優化 iOS / Android 雙平台使用者體驗與離線購物車同步',
      color: '#8250df',
      createdAt: '2026-11-03T10:30:00.000Z',
    },
    {
      id: 'proj-infra',
      name: '核心系統優化 (Core & Infra)',
      description: 'API 回應時間最佳化、資料庫索引調整與快取架構升級',
      color: '#1a7f37',
      createdAt: '2026-11-05T14:15:00.000Z',
    },
  ];

  const INITIAL_ISSUES = [
    {
      id: 1,
      projectId: 'proj-web',
      title: '結帳頁面在手機瀏覽器偶發性卡死 (P0 緊急)',
      description: '使用者反映在 Safari iOS 18 進入多重優惠券折抵時，結帳按鈕未正常觸發驗證事件，造成畫面停止回應。\n\n重現步驟：\n1. 加入 2 件商品至購物車\n2. 輸入折扣代碼「VIP2026」\n3. 點選快速結帳按鈕',
      status: 'open',
      priority: 'p0',
      tags: ['bug', 'checkout', 'urgent', 'safari'],
      progress: 70,
      createdAt: '2026-11-10T08:30:00.000Z',
      updatedAt: '2026-11-12T04:15:00.000Z',
      updates: [
        {
          id: 'u-1',
          timestamp: '2026-11-10T10:00:00.000Z',
          percentage: 20,
          note: '已在實體 iPhone 15 Pro 上成功重現問題，確認是 Promise catch 區塊遺漏未處理微任務錯誤。',
        },
        {
          id: 'u-2',
          timestamp: '2026-11-11T16:20:00.000Z',
          percentage: 50,
          note: '修正前端驗證邏輯，重構購物車折抵計算非同步調用。正進行跨裝置單元測試。',
        },
        {
          id: 'u-3',
          timestamp: '2026-11-12T04:15:00.000Z',
          percentage: 70,
          note: 'Safari 與 Chrome 測試通過，目前正進行真實付款閘道壓力測試。',
        },
      ],
    },
    {
      id: 2,
      projectId: 'proj-web',
      title: '新增深色模式 (Dark Theme) 與系統偏好自動切換',
      description: '依照 GitHub 色票規範，實作高對比度深色主題切換功能，並支援跟隨作業系統設定。',
      status: 'open',
      priority: 'p2',
      tags: ['feature', 'ui/ux', 'frontend'],
      progress: 40,
      createdAt: '2026-11-08T11:00:00.000Z',
      updatedAt: '2026-11-11T09:20:00.000Z',
      updates: [
        {
          id: 'u-4',
          timestamp: '2026-11-09T14:00:00.000Z',
          percentage: 20,
          note: '定義 dark mode 顏色變數與色票對應表。',
        },
        {
          id: 'u-5',
          timestamp: '2026-11-11T09:20:00.000Z',
          percentage: 40,
          note: '已完成導航列與按鈕元件深色樣式。',
        },
      ],
    },
    {
      id: 3,
      projectId: 'proj-web',
      title: '商品列表圖片支援 WebP 格式與延遲載入 (Lazy Loading)',
      description: '將靜態資源全面替換為新世代圖片格式，並針對首頁與分類列表啟用 IntersectionObserver 延遲載入，減少 45% 首屏流量。',
      status: 'closed',
      priority: 'p1',
      tags: ['performance', 'optimization'],
      progress: 100,
      createdAt: '2026-11-02T10:00:00.000Z',
      updatedAt: '2026-11-06T15:30:00.000Z',
      closedAt: '2026-11-06T15:30:00.000Z',
      updates: [
        {
          id: 'u-6',
          timestamp: '2026-11-03T09:10:00.000Z',
          percentage: 30,
          note: '測試不同圖片壓縮比例與品質設定 (q=82)。',
        },
        {
          id: 'u-7',
          timestamp: '2026-11-06T15:30:00.000Z',
          percentage: 100,
          note: '驗證通過！Lighthouse 效能分數從 68 提升至 94。已發布上線並結案。',
        },
      ],
    },
    {
      id: 4,
      projectId: 'proj-mobile',
      title: '支援生物辨識登入 (Face ID / 指紋解鎖)',
      description: '利用 WebAuthn 與原生鑰匙圈 API，讓已驗證使用者下次開啟時可直接刷臉或指紋解鎖。',
      status: 'open',
      priority: 'p1',
      tags: ['security', 'feature', 'mobile'],
      progress: 85,
      createdAt: '2026-11-05T13:00:00.000Z',
      updatedAt: '2026-11-11T17:00:00.000Z',
      updates: [
        {
          id: 'u-8',
          timestamp: '2026-11-07T11:00:00.000Z',
          percentage: 40,
          note: '完成本機安全存儲整合。',
        },
        {
          id: 'u-9',
          timestamp: '2026-11-11T17:00:00.000Z',
          percentage: 85,
          note: '已實裝 Face ID 測試流程，剩下使用者拒絕授權時的 fallback 提示。',
        },
      ],
    },
    {
      id: 5,
      projectId: 'proj-mobile',
      title: '推播通知設定頁面分類開關',
      description: '讓使用者自訂欲接收之通知種類 (訂單動態、促銷優惠、系統公告)。',
      status: 'closed',
      priority: 'p3',
      tags: ['ui/ux', 'settings'],
      progress: 100,
      createdAt: '2026-11-04T09:00:00.000Z',
      updatedAt: '2026-11-07T12:00:00.000Z',
      closedAt: '2026-11-07T12:00:00.000Z',
      updates: [
        {
          id: 'u-10',
          timestamp: '2026-11-07T12:00:00.000Z',
          percentage: 100,
          note: '已整合使用者偏好儲存並部署完成。',
        },
      ],
    },
    {
      id: 6,
      projectId: 'proj-infra',
      title: '商品全文檢索回應時間優化 (目標 < 50ms)',
      description: '目前熱門搜尋時資料庫 CPU 負載較高，需建立全文檢索複合索引並導入記憶體快取。',
      status: 'open',
      priority: 'p1',
      tags: ['backend', 'database', 'performance'],
      progress: 60,
      createdAt: '2026-11-06T14:00:00.000Z',
      updatedAt: '2026-11-12T02:00:00.000Z',
      updates: [
        {
          id: 'u-11',
          timestamp: '2026-11-08T10:00:00.000Z',
          percentage: 30,
          note: '分析 slow query log，找出瓶頸在於多欄位 LIKE 模糊搜尋。',
        },
        {
          id: 'u-12',
          timestamp: '2026-11-12T02:00:00.000Z',
          percentage: 60,
          note: '已建立 pg_trgm 倒排索引，P95 回應時間降至 42ms。',
        },
      ],
    },
  ];

  // App State
  const state = {
    projects: [],
    issues: [],
    activeProjectId: 'all',
    statusFilter: 'open',
    searchQuery: '',
    priorityFilter: 'all',
    tagFilter: 'all',
    sortBy: 'newest',
    selectedIssueId: null,
    newIssueSelectedPriority: 'p2',
    selectedProjectColor: '#0969da',
    // Source tracking: 'auto-file' | 'local-cache' | 'uploaded-file' | 'empty'
    dataSourceMode: 'empty',
    detectedFilePath: null,
    lastAutoFileHash: null,
    isFileProtocolBlocked: false,
  };

  const PRIORITY_MAP = {
    p0: { label: 'P0 - 緊急', short: 'P0 緊急', cls: 'priority-p0' },
    p1: { label: 'P1 - 高優先', short: 'P1 高', cls: 'priority-p1' },
    p2: { label: 'P2 - 中優先', short: 'P2 中', cls: 'priority-p2' },
    p3: { label: 'P3 - 低優先', short: 'P3 低', cls: 'priority-p3' },
  };

  const PROJECT_PRESET_COLORS = ['#0969da', '#8250df', '#1a7f37', '#cf222e', '#bf8700', '#218bff', '#f0883e'];

  const TAG_COLORS = [
    { bg: '#ffebe9', text: '#cf222e' },
    { bg: '#ddf4ff', text: '#0969da' },
    { bg: '#dafbe1', text: '#1a7f37' },
    { bg: '#fbefff', text: '#8250df' },
    { bg: '#fff8c5', text: '#9a6700' },
    { bg: '#f3f4f6', text: '#57606a' },
  ];

  function getTagColors(tag) {
    const lower = String(tag).toLowerCase().trim();
    if (lower === 'bug' || lower === 'urgent') return TAG_COLORS[0];
    if (lower === 'feature' || lower === 'frontend') return TAG_COLORS[1];
    if (lower === 'docs' || lower === 'performance') return TAG_COLORS[2];
    if (lower === 'ui' || lower === 'ux' || lower === 'ui/ux') return TAG_COLORS[3];
    if (lower === 'security') return TAG_COLORS[4];

    let hash = 0;
    for (let i = 0; i < tag.length; i++) {
      hash = tag.charCodeAt(i) + ((hash << 5) - hash);
    }
    return TAG_COLORS[Math.abs(hash) % TAG_COLORS.length];
  }

  // ==========================================================================
  // Formatting & Normalization Helpers
  // ==========================================================================

  function formatDateTime(iso) {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      if (isNaN(d.getTime())) return String(iso);
      return d.toLocaleString('zh-TW', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hour12: false,
      });
    } catch {
      return String(iso);
    }
  }

  function formatRelativeTime(iso) {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      if (isNaN(d.getTime())) return String(iso);
      const diff = Date.now() - d.getTime();
      if (diff < 0) return '剛剛';
      const sec = Math.floor(diff / 1000);
      if (sec < 60) return '剛剛';
      const min = Math.floor(sec / 60);
      if (min < 60) return `${min} 分鐘前`;
      const hour = Math.floor(min / 60);
      if (hour < 24) return `${hour} 小時前`;
      const day = Math.floor(hour / 24);
      if (day < 30) return `${day} 天前`;
      return formatDateTime(iso);
    } catch {
      return String(iso);
    }
  }

  function normalizeDateValue(val, fallback) {
    if (val === null || val === undefined || val === '') return fallback || new Date().toISOString();
    const num = Number(val);
    if (!isNaN(num) && String(val).trim() !== '' && num > 20000 && num < 100000) {
      const iso = excelSerialToIso(num);
      if (iso) return iso;
    }
    try {
      const d = new Date(String(val).replace(/\//g, '-'));
      if (!isNaN(d.getTime())) return d.toISOString();
    } catch {
      // ignore
    }
    return fallback || new Date().toISOString();
  }

  function normalizeProgressValue(val) {
    if (val === undefined || val === null || val === '') return 0;
    const str = String(val).replace('%', '').trim();
    const num = parseFloat(str);
    if (isNaN(num)) return 0;
    if (num > 0 && num <= 1 && String(val).includes('.')) {
      return Math.min(100, Math.max(0, Math.round(num * 100)));
    }
    return Math.min(100, Math.max(0, Math.round(num)));
  }

  function normalizePriorityValue(val) {
    if (!val) return 'p2';
    const p = String(val).toLowerCase().trim();
    if (p.includes('p0') || p.includes('緊急') || p.includes('urgent') || p === '0') return 'p0';
    if (p.includes('p1') || p.includes('高') || p.includes('high') || p === '1') return 'p1';
    if (p.includes('p2') || p.includes('中') || p.includes('medium') || p === '2') return 'p2';
    if (p.includes('p3') || p.includes('低') || p.includes('low') || p === '3') return 'p3';
    return 'p2';
  }

  function normalizeStatusValue(val, progressNum) {
    if (!val) return progressNum === 100 ? 'closed' : 'open';
    const s = String(val).toLowerCase().trim();
    if (s.includes('close') || s.includes('結案') || s.includes('關閉') || s.includes('完成') || s === 'done' || s === 'resolved') {
      return 'closed';
    }
    return 'open';
  }

  function normalizeTagsValue(val) {
    if (!val) return [];
    return String(val)
      .split(/[,;\n|/]/)
      .map((t) => t.trim().toLowerCase())
      .filter(Boolean);
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function showToast(msg) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    container.className = 'toast-notice';
    container.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#3fb950" stroke-width="2.5"><polyline points="20 6 9 17 4 12"></polyline></svg>
      <span>${escapeHtml(msg)}</span>
    `;
    container.style.display = 'flex';
    clearTimeout(showToast._timer);
    showToast._timer = setTimeout(() => {
      container.style.display = 'none';
    }, 3500);
  }

  // ==========================================================================
  // LocalStorage Persistence
  // ==========================================================================

  function loadStoredState() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed.projects) && Array.isArray(parsed.issues)) {
          return {
            projects: parsed.projects,
            issues: parsed.issues,
            lastAutoFileHash: parsed.lastAutoFileHash || null,
          };
        }
      }
    } catch (e) {
      console.warn('LocalStorage read error:', e);
    }
    return null;
  }

  function saveState() {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          projects: state.projects,
          issues: state.issues,
          lastAutoFileHash: state.lastAutoFileHash,
        })
      );
    } catch (e) {
      console.error('LocalStorage save error:', e);
    }
  }

  // ==========================================================================
  // XLSX Parsing Logic (Supports Multi-Sheet & IssueDesk Master Format)
  // ==========================================================================

  async function convertXlsxBufferToData(arrayBuffer) {
    const parsedSheets = await parseXlsx(arrayBuffer);
    if (!parsedSheets || parsedSheets.length === 0) {
      throw new Error('該 Excel 檔案中沒有發現任何工作表');
    }

    const newProjects = [];
    const newIssues = [];
    const projectMapByName = new Map();
    const projectMapById = new Map();

    const registerProject = (name, id, desc, color, createdAt) => {
      const trimmedName = String(name || '未命名專案').trim();
      const targetId = id ? String(id).trim() : null;
      if (targetId && projectMapById.has(targetId)) return projectMapById.get(targetId);
      if (projectMapByName.has(trimmedName.toLowerCase())) return projectMapByName.get(trimmedName.toLowerCase());

      const proj = {
        id: targetId || `proj-${Date.now()}-${Math.random().toString(36).substring(2, 6)}`,
        name: trimmedName,
        description: desc ? String(desc) : '',
        color: color || PROJECT_PRESET_COLORS[newProjects.length % PROJECT_PRESET_COLORS.length],
        createdAt: createdAt || new Date().toISOString(),
      };
      newProjects.push(proj);
      projectMapById.set(proj.id, proj);
      projectMapByName.set(trimmedName.toLowerCase(), proj);
      return proj;
    };

    // 1. Check if Projects metadata sheet exists
    const projectsSheet = parsedSheets.find((s) => {
      const n = s.name.toLowerCase().trim();
      return n === 'projects' || n === '專案列表' || n === '專案總覽';
    });

    if (projectsSheet && projectsSheet.rows.length > 1) {
      const headers = projectsSheet.rows[0].map((h) => String(h).toLowerCase());
      const colName = headers.findIndex((h) => h.includes('name') || h.includes('專案名稱') || h.includes('名稱'));
      const colId = headers.findIndex((h) => h.includes('id') || h.includes('專案id'));
      const colDesc = headers.findIndex((h) => h.includes('desc') || h.includes('說明'));
      const colCreated = headers.findIndex((h) => h.includes('created') || h.includes('建立'));

      for (let r = 1; r < projectsSheet.rows.length; r++) {
        const row = projectsSheet.rows[r];
        const name = colName !== -1 && row[colName] ? row[colName] : row[0];
        if (!name) continue;
        const id = colId !== -1 && row[colId] ? row[colId] : null;
        const desc = colDesc !== -1 && row[colDesc] ? row[colDesc] : '';
        const createdAt = colCreated !== -1 && row[colCreated] ? normalizeDateValue(row[colCreated]) : null;
        registerProject(name, id, desc, null, createdAt);
      }
    }

    // 2. Check if Progress History sheet exists
    const progressHistoryMap = new Map();
    const historySheet = parsedSheets.find((s) => {
      const n = s.name.toLowerCase().trim();
      return n === 'progress history' || n.includes('歷程') || n.includes('history');
    });

    if (historySheet && historySheet.rows.length > 1) {
      const headers = historySheet.rows[0].map((h) => String(h).toLowerCase());
      const colIssueId = headers.findIndex((h) => h.includes('issue') || h.includes('編號') || h.includes('id'));
      const colPct = headers.findIndex((h) => h.includes('progress') || h.includes('進度'));
      const colTime = headers.findIndex((h) => h.includes('time') || h.includes('時間'));
      const colNote = headers.findIndex((h) => h.includes('note') || h.includes('說明') || h.includes('備註'));

      for (let r = 1; r < historySheet.rows.length; r++) {
        const row = historySheet.rows[r];
        const rawId = colIssueId !== -1 ? row[colIssueId] : row[0];
        if (!rawId) continue;
        const cleanId = String(rawId).replace('#', '').trim();
        const pct = normalizeProgressValue(colPct !== -1 ? row[colPct] : 0);
        const ts = normalizeDateValue(colTime !== -1 ? row[colTime] : null);
        const note = colNote !== -1 && row[colNote] ? String(row[colNote]) : '';
        if (!progressHistoryMap.has(cleanId)) progressHistoryMap.set(cleanId, []);
        progressHistoryMap.get(cleanId).push({
          id: `up-hist-${r}-${Date.now()}`,
          timestamp: ts,
          percentage: pct,
          note,
        });
      }
    }

    // 3. Check if Master Issues sheet exists
    const masterIssuesSheet = parsedSheets.find((s) => {
      const n = s.name.toLowerCase().trim();
      return n === 'issues' || n === '任務清單' || n === '問題清單';
    });

    if (masterIssuesSheet && masterIssuesSheet.rows.length > 1) {
      const headers = masterIssuesSheet.rows[0].map((h) => String(h).toLowerCase());
      const colProj = headers.findIndex((h) => h.includes('project') && !h.includes('id') || h.includes('所屬專案') || h.includes('專案名稱') || h === '專案');
      const colProjId = headers.findIndex((h) => h.includes('專案id') || h.includes('project id'));
      const colId = headers.findIndex((h) => h.includes('issue編號') || h === 'id' || h === '編號' || h.includes('issue id'));
      const colTitle = headers.findIndex((h) => h.includes('title') || h.includes('標題') || h.includes('名稱'));
      const colStatus = headers.findIndex((h) => h.includes('status') || h.includes('狀態'));
      const colPriority = headers.findIndex((h) => h.includes('priority') || h.includes('優先'));
      const colTags = headers.findIndex((h) => h.includes('tag') || h.includes('標籤'));
      const colProgress = headers.findIndex((h) => h.includes('progress') || h.includes('進度'));
      const colDesc = headers.findIndex((h) => h.includes('desc') || h.includes('詳細說明') || h === '說明' || h.includes('內容'));
      const colCreated = headers.findIndex((h) => h.includes('created') || h.includes('建立'));
      const colUpdated = headers.findIndex((h) => h.includes('updated') || h.includes('更新'));
      const colClosed = headers.findIndex((h) => h.includes('closed') || h.includes('結案'));
      const colNote = headers.findIndex((h) => h.includes('note') || h.includes('備註'));
      const colUpdatesJson = headers.findIndex((h) => h.includes('updatesdata_json'));

      for (let r = 1; r < masterIssuesSheet.rows.length; r++) {
        const row = masterIssuesSheet.rows[r];
        if (!row.some((c) => c !== '')) continue;

        const projIdCell = colProjId !== -1 && row[colProjId] ? String(row[colProjId]).trim() : '';
        const projNameCell = colProj !== -1 && row[colProj] ? String(row[colProj]).trim() : '';
        const project = (projIdCell && projectMapById.get(projIdCell))
          || registerProject(projNameCell || '預設專案', projIdCell || null);

        const rawId = colId !== -1 && row[colId] ? String(row[colId]).replace('#', '').trim() : String(r);
        const numId = parseInt(rawId, 10) || r;
        const title = colTitle !== -1 && row[colTitle] ? String(row[colTitle]) : `任務 #${numId}`;
        const progressNum = normalizeProgressValue(colProgress !== -1 ? row[colProgress] : 0);
        const status = normalizeStatusValue(colStatus !== -1 ? row[colStatus] : '', progressNum);
        const priority = normalizePriorityValue(colPriority !== -1 ? row[colPriority] : 'p2');
        const tags = normalizeTagsValue(colTags !== -1 ? row[colTags] : '');
        const desc = colDesc !== -1 && row[colDesc] ? String(row[colDesc]) : '';
        const createdAt = normalizeDateValue(colCreated !== -1 ? row[colCreated] : null);
        const updatedAt = normalizeDateValue(colUpdated !== -1 ? row[colUpdated] : null, createdAt);
        const closedAt = status === 'closed'
          ? normalizeDateValue(colClosed !== -1 ? row[colClosed] : null, updatedAt)
          : null;

        let updates = [];
        if (colUpdatesJson !== -1 && row[colUpdatesJson]) {
          try {
            const parsed = JSON.parse(String(row[colUpdatesJson]));
            if (Array.isArray(parsed)) updates = parsed;
          } catch {
            // ignore
          }
        }
        if (updates.length === 0 && progressHistoryMap.has(String(numId))) {
          updates = progressHistoryMap.get(String(numId));
        }
        if (updates.length === 0) {
          const note = colNote !== -1 && row[colNote] ? String(row[colNote]) : '';
          updates = [{
            id: `up-init-${numId}`,
            timestamp: updatedAt,
            percentage: progressNum,
            note: note || `匯入任務，當前進度 ${progressNum}%`,
          }];
        }

        newIssues.push({
          id: numId,
          projectId: project.id,
          title,
          description: desc,
          status,
          priority,
          tags,
          progress: progressNum,
          createdAt,
          updatedAt,
          closedAt,
          updates,
        });
      }
    } else {
      // Multi-sheet mode: each worksheet represents an independent Project page!
      let globalIdCounter = 1;

      parsedSheets.forEach((sheet) => {
        const lowerName = sheet.name.toLowerCase().trim();
        if (
          lowerName === 'projects' ||
          lowerName === 'progress history' ||
          lowerName.includes('summary') ||
          sheet.rows.length <= 1
        ) {
          return;
        }

        const project = registerProject(sheet.name);
        const headers = sheet.rows[0].map((h) => String(h).toLowerCase());

        const colId = headers.findIndex((h) => h === 'id' || h === '編號' || h.includes('issue編號'));
        const colTitle = headers.findIndex((h) => h.includes('title') || h.includes('標題') || h.includes('名稱') || h.includes('任務') || h.includes('issue'));
        const colStatus = headers.findIndex((h) => h.includes('status') || h.includes('狀態'));
        const colPriority = headers.findIndex((h) => h.includes('priority') || h.includes('優先'));
        const colTags = headers.findIndex((h) => h.includes('tag') || h.includes('標籤') || h.includes('分類'));
        const colProgress = headers.findIndex((h) => h.includes('progress') || h.includes('進度'));
        const colDesc = headers.findIndex((h) => h.includes('desc') || h.includes('說明') || h.includes('內容'));
        const colCreated = headers.findIndex((h) => h.includes('created') || h.includes('建立'));
        const colUpdated = headers.findIndex((h) => h.includes('updated') || h.includes('更新'));
        const colNote = headers.findIndex((h) => h.includes('note') || h.includes('備註'));

        for (let r = 1; r < sheet.rows.length; r++) {
          const row = sheet.rows[r];
          if (!row.some((c) => c !== '')) continue;

          const numId = (colId !== -1 && row[colId] ? parseInt(String(row[colId]).replace('#', ''), 10) : 0) || globalIdCounter++;
          if (numId >= globalIdCounter) globalIdCounter = numId + 1;

          const title = colTitle !== -1 && row[colTitle] ? String(row[colTitle]) : (row[0] || `任務 #${numId}`);
          const progressNum = normalizeProgressValue(colProgress !== -1 ? row[colProgress] : 0);
          const status = normalizeStatusValue(colStatus !== -1 ? row[colStatus] : '', progressNum);
          const priority = normalizePriorityValue(colPriority !== -1 ? row[colPriority] : 'p2');
          const tags = normalizeTagsValue(colTags !== -1 ? row[colTags] : '');
          const desc = colDesc !== -1 && row[colDesc] ? String(row[colDesc]) : '';
          const note = colNote !== -1 && row[colNote] ? String(row[colNote]) : '';
          const createdAt = normalizeDateValue(colCreated !== -1 ? row[colCreated] : null);
          const updatedAt = normalizeDateValue(colUpdated !== -1 ? row[colUpdated] : null, createdAt);

          newIssues.push({
            id: numId,
            projectId: project.id,
            title,
            description: desc,
            status,
            priority,
            tags,
            progress: progressNum,
            createdAt,
            updatedAt,
            closedAt: status === 'closed' ? updatedAt : null,
            updates: [{
              id: `up-sheet-${numId}`,
              timestamp: updatedAt,
              percentage: progressNum,
              note: note || `從專案工作表「${sheet.name}」匯入，進度 ${progressNum}%`,
            }],
          });
        }
      });
    }

    if (newProjects.length === 0) {
      throw new Error('無法在 Excel 檔案中解析出專案或 Issue 資料，請確認工作表至少包含標題列與資料列');
    }

    return { projects: newProjects, issues: newIssues };
  }

  // ==========================================================================
  // Auto-Load XLSX From Designated Location (./data/issues.xlsx or ./issues.xlsx)
  // ==========================================================================

  async function tryFetchDesignatedXlsx() {
    if (window.location.protocol === 'file:') {
      return { found: false, fileProtocolBlocked: true };
    }

    for (const path of AUTO_FILE_PATHS) {
      try {
        const res = await fetch(path, { cache: 'no-store' });
        if (res.ok) {
          const contentType = res.headers.get('content-type') || '';
          // Ensure server didn't return an HTML fallback page for 404
          if (contentType.includes('text/html')) continue;
          const buffer = await res.arrayBuffer();
          if (buffer && buffer.byteLength > 100) {
            const bytes = new Uint8Array(buffer);
            // Verify ZIP magic bytes (PK\x03\x04)
            if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
              const hash = `${buffer.byteLength}-${computeCRC32(bytes)}`;
              return { found: true, path, buffer, hash };
            }
          }
        }
      } catch (err) {
        // Continue trying next candidate
      }
    }

    return { found: false, fileProtocolBlocked: false };
  }

  async function reloadFromDesignatedFile(forceOverwrite = true) {
    const result = await tryFetchDesignatedXlsx();
    if (!result.found) {
      showToast('未在 ./data/issues.xlsx 偵測到有效的 XLSX 檔案');
      return false;
    }

    try {
      const parsed = await convertXlsxBufferToData(result.buffer);
      state.projects = parsed.projects;
      state.issues = parsed.issues;
      state.activeProjectId = 'all';
      state.lastAutoFileHash = result.hash;
      state.detectedFilePath = result.path;
      state.dataSourceMode = 'auto-file';
      if (forceOverwrite) {
        saveState();
      }
      showToast(`已自動解析 ${result.path}（${parsed.projects.length} 個專案、${parsed.issues.length} 筆 Issue）`);
      renderApp();
      return true;
    } catch (err) {
      alert(`自動解析 ${result.path} 失敗：${err.message}`);
      return false;
    }
  }

  // ==========================================================================
  // Rendering Views
  // ==========================================================================

  function renderSourceStatusAndOnboarding() {
    const statusContainer = document.getElementById('sourceStatusContainer');
    const onboardingContainer = document.getElementById('onboardingContainer');
    const workspaceContent = document.getElementById('workspaceContent');

    const isEmpty = state.projects.length === 0;

    // 1. Render Top Status Notice Bar
    if (state.dataSourceMode === 'auto-file' && state.detectedFilePath) {
      statusContainer.innerHTML = `
        <div class="source-notice-bar source-auto">
          <div class="source-notice-left">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polyline points="20 6 9 17 4 12"></polyline></svg>
            <span>已自動解析指定檔案 <span class="code-path">${escapeHtml(state.detectedFilePath)}</span> 為專案管理系統</span>
          </div>
          <div class="source-notice-actions">
            <button type="button" class="btn" id="btnReloadAutoFile" style="padding: 4px 10px; font-size: 12px;">重新讀取檔案</button>
          </div>
        </div>
      `;
      document.getElementById('btnReloadAutoFile').addEventListener('click', () => reloadFromDesignatedFile(true));
    } else if (!isEmpty) {
      const msg = state.isFileProtocolBlocked
        ? '目前以本機檔案模式 (file://) 開啟，您可以透過「上傳 XLSX」選擇電腦檔案，或直接於此管理並匯出覆蓋至 <span class="code-path">./data/issues.xlsx</span>。'
        : state.detectedFilePath
          ? `已載入本機最新編輯暫存（預設來源：<span class="code-path">${escapeHtml(state.detectedFilePath)}</span>）`
          : '未在 <span class="code-path">./data/issues.xlsx</span> 偵測到檔案，目前使用本機工作區資料。';

      statusContainer.innerHTML = `
        <div class="source-notice-bar ${state.detectedFilePath ? '' : 'source-missing'}">
          <div class="source-notice-left">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
            <span>${msg}</span>
          </div>
          <div class="source-notice-actions">
            ${state.detectedFilePath ? `<button type="button" class="btn" id="btnReloadAutoFile" style="padding: 4px 10px; font-size: 12px;">從 ${escapeHtml(state.detectedFilePath)} 重載</button>` : ''}
            <button type="button" class="btn" id="btnNoticeUpload" style="padding: 4px 10px; font-size: 12px;">選擇 XLSX 檔案</button>
          </div>
        </div>
      `;
      const reloadBtn = document.getElementById('btnReloadAutoFile');
      if (reloadBtn) reloadBtn.addEventListener('click', () => reloadFromDesignatedFile(true));
      document.getElementById('btnNoticeUpload').addEventListener('click', () => {
        document.getElementById('uploadPreviewSection').style.display = 'none';
        document.getElementById('modalUpload').style.display = 'flex';
      });
    } else {
      statusContainer.innerHTML = '';
    }

    // 2. Render Onboarding when workspace is completely empty (no file at ./data/issues.xlsx & no saved localStorage)
    if (isEmpty) {
      workspaceContent.style.display = 'none';
      onboardingContainer.style.display = 'block';
      onboardingContainer.innerHTML = `
        <div class="onboarding-card">
          <div class="onboarding-icon">
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="12" y1="18" x2="12" y2="12"></line>
              <line x1="9" y1="15" x2="15" y2="15"></line>
            </svg>
          </div>
          <h2 class="onboarding-title">指定位置尚無 XLSX 檔案</h2>
          <p class="onboarding-desc">
            ${state.isFileProtocolBlocked
              ? '偵測到您正在本機以 <span class="code-path">file://</span> 直接開啟網頁。請點擊下方按鈕選擇您的 <span class="code-path">issues.xlsx</span> 檔案（或直接將檔案拖曳至視窗），即可立即解析為專案管理系統！'
              : '未在預設路徑 <span class="code-path">./data/issues.xlsx</span>（或 <span class="code-path">./issues.xlsx</span>）找到 Excel 檔案。您可以將檔案放入該資料夾並重新整理，或直接使用以下方式開始：'}
          </p>
          <div class="onboarding-actions">
            <button type="button" class="btn btn-primary" id="btnOnboardingUpload">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
              <span>選擇 / 拖曳 XLSX 檔案</span>
            </button>
            <button type="button" class="btn btn-blue" id="btnOnboardingNewProject">
              <span>+ 建立空白專案</span>
            </button>
            <button type="button" class="btn" id="btnOnboardingLoadDemo">
              <span>載入範例專案資料</span>
            </button>
            <button type="button" class="btn" id="btnOnboardingDownloadTemplate">
              <span>下載範本 (issues.xlsx)</span>
            </button>
          </div>
        </div>
      `;

      document.getElementById('btnOnboardingUpload').addEventListener('click', () => {
        document.getElementById('uploadPreviewSection').style.display = 'none';
        document.getElementById('modalUpload').style.display = 'flex';
      });
      document.getElementById('btnOnboardingNewProject').addEventListener('click', () => openProjectModal(null));
      document.getElementById('btnOnboardingLoadDemo').addEventListener('click', () => {
        state.projects = JSON.parse(JSON.stringify(INITIAL_PROJECTS));
        state.issues = JSON.parse(JSON.stringify(INITIAL_ISSUES));
        state.activeProjectId = 'all';
        state.dataSourceMode = 'local-cache';
        saveState();
        showToast('已載入範例專案資料');
        renderApp();
      });
      document.getElementById('btnOnboardingDownloadTemplate').addEventListener('click', () => {
        exportProjectsAndIssuesToXlsx(INITIAL_PROJECTS, INITIAL_ISSUES, 'issues.xlsx');
      });
    } else {
      workspaceContent.style.display = 'block';
      onboardingContainer.style.display = 'none';
      onboardingContainer.innerHTML = '';
    }
  }

  function renderProjectTabs() {
    const container = document.getElementById('projectTabsContainer');
    if (!container) return;

    const totalOpen = state.issues.filter((i) => i.status === 'open').length;

    let html = `
      <button type="button" class="project-tab-btn ${state.activeProjectId === 'all' ? 'active' : ''}" data-project-id="all">
        <span>全部專案 (${state.projects.length})</span>
        <span class="tab-badge">${totalOpen}</span>
      </button>
    `;

    state.projects.forEach((proj) => {
      const projIssues = state.issues.filter((i) => i.projectId === proj.id);
      const openCount = projIssues.filter((i) => i.status === 'open').length;
      const isActive = state.activeProjectId === proj.id;

      html += `
        <button type="button" class="project-tab-btn ${isActive ? 'active' : ''}" data-project-id="${escapeHtml(proj.id)}">
          <span class="project-color-dot" style="background-color: ${proj.color || '#0969da'};"></span>
          <span>${escapeHtml(proj.name)}</span>
          <span class="tab-badge">${openCount}</span>
        </button>
      `;
    });

    container.innerHTML = html;

    container.querySelectorAll('.project-tab-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.activeProjectId = btn.getAttribute('data-project-id');
        renderApp();
      });
    });
  }

  function renderProjectBanner() {
    const container = document.getElementById('projectBanner');
    if (!container || state.projects.length === 0) return;

    if (state.activeProjectId === 'all') {
      const totalIssues = state.issues.length;
      const closedIssues = state.issues.filter((i) => i.status === 'closed').length;
      const avgProgress = totalIssues > 0
        ? Math.round(state.issues.reduce((acc, i) => acc + (i.progress || 0), 0) / totalIssues)
        : 0;

      container.innerHTML = `
        <div class="banner-header">
          <div>
            <div class="banner-title-area">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#0969da" stroke-width="2"><path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"></path></svg>
              <h2 class="banner-title">所有專案概覽 (All Projects)</h2>
            </div>
            <div class="banner-desc">共 ${state.projects.length} 個專案進行中，整體進度 ${avgProgress}%</div>
          </div>

          <div class="banner-stats">
            <div class="stat-pill">
              <div class="stat-label">待處理 Issue</div>
              <div class="stat-value" style="color: #238636;">${totalIssues - closedIssues}</div>
            </div>
            <div class="stat-pill">
              <div class="stat-label">已結案</div>
              <div class="stat-value" style="color: #8250df;">${closedIssues}</div>
            </div>
            <div class="stat-pill">
              <div class="stat-label">平均完成度</div>
              <div class="stat-value">${avgProgress}%</div>
            </div>
          </div>
        </div>
        <div class="progress-track">
          <div class="progress-fill" style="width: ${avgProgress}%;"></div>
        </div>
      `;
    } else {
      const proj = state.projects.find((p) => p.id === state.activeProjectId);
      if (!proj) {
        state.activeProjectId = 'all';
        renderProjectBanner();
        return;
      }

      const projIssues = state.issues.filter((i) => i.projectId === proj.id);
      const openCount = projIssues.filter((i) => i.status === 'open').length;
      const closedCount = projIssues.filter((i) => i.status === 'closed').length;
      const avgProgress = projIssues.length > 0
        ? Math.round(projIssues.reduce((acc, i) => acc + (i.progress || 0), 0) / projIssues.length)
        : 0;

      container.innerHTML = `
        <div class="banner-header">
          <div>
            <div class="banner-title-area">
              <span class="project-color-dot" style="width: 12px; height: 12px; background-color: ${proj.color || '#0969da'};"></span>
              <h2 class="banner-title">${escapeHtml(proj.name)}</h2>
              <span class="badge-tag">ID: ${escapeHtml(proj.id)}</span>
            </div>
            ${proj.description ? `<div class="banner-desc">${escapeHtml(proj.description)}</div>` : ''}
          </div>

          <div class="banner-stats">
            <div class="stat-pill">
              <div class="stat-label">待處理</div>
              <div class="stat-value" style="color: #238636;">${openCount}</div>
            </div>
            <div class="stat-pill">
              <div class="stat-label">已結案</div>
              <div class="stat-value" style="color: #8250df;">${closedCount}</div>
            </div>
            <div class="stat-pill">
              <div class="stat-label">達成率</div>
              <div class="stat-value">${avgProgress}%</div>
            </div>
            <button type="button" class="btn" id="btnEditCurrentProject" title="編輯專案設定" style="padding: 6px 10px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
            </button>
            <button type="button" class="btn btn-danger-outline" id="btnDeleteCurrentProject" title="刪除專案" style="padding: 6px 10px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
            </button>
          </div>
        </div>
        <div class="progress-track">
          <div class="progress-fill" style="width: ${avgProgress}%;"></div>
        </div>
      `;

      const editBtn = document.getElementById('btnEditCurrentProject');
      if (editBtn) editBtn.addEventListener('click', () => openProjectModal(proj));

      const delBtn = document.getElementById('btnDeleteCurrentProject');
      if (delBtn) {
        delBtn.addEventListener('click', () => {
          if (confirm(`確定要刪除專案「${proj.name}」嗎？該專案下的 ${projIssues.length} 個 Issue 也將被一併移除。`)) {
            state.projects = state.projects.filter((p) => p.id !== proj.id);
            state.issues = state.issues.filter((i) => i.projectId !== proj.id);
            state.activeProjectId = 'all';
            saveState();
            showToast(`已刪除專案「${proj.name}」`);
            renderApp();
          }
        });
      }
    }
  }

  function updateTagFilterDropdown() {
    const select = document.getElementById('tagFilter');
    if (!select) return;

    const scopedIssues = state.activeProjectId === 'all'
      ? state.issues
      : state.issues.filter((i) => i.projectId === state.activeProjectId);

    const tagsSet = new Set();
    scopedIssues.forEach((i) => {
      (i.tags || []).forEach((t) => tagsSet.add(t));
    });
    const tagsList = Array.from(tagsSet).sort();

    if (state.tagFilter !== 'all' && !tagsSet.has(state.tagFilter)) {
      state.tagFilter = 'all';
    }

    let html = `<option value="all">所有標籤 (${tagsList.length})</option>`;
    tagsList.forEach((t) => {
      html += `<option value="${escapeHtml(t)}" ${state.tagFilter === t ? 'selected' : ''}>🏷️ ${escapeHtml(t)}</option>`;
    });
    select.innerHTML = html;
  }

  function renderIssueList() {
    const container = document.getElementById('issueListContainer');
    if (!container || state.projects.length === 0) return;

    let list = state.activeProjectId === 'all'
      ? state.issues
      : state.issues.filter((i) => i.projectId === state.activeProjectId);

    const openCount = list.filter((i) => i.status === 'open').length;
    const closedCount = list.filter((i) => i.status === 'closed').length;

    document.getElementById('openCountText').innerText = `${openCount} 待處理 (Open)`;
    document.getElementById('closedCountText').innerText = `${closedCount} 已結案 (Closed)`;

    document.getElementById('tabOpenIssues').className = `state-tab-btn ${state.statusFilter === 'open' ? 'active' : ''}`;
    document.getElementById('tabClosedIssues').className = `state-tab-btn ${state.statusFilter === 'closed' ? 'active' : ''}`;

    list = list.filter((i) => i.status === state.statusFilter);

    if (state.searchQuery.trim()) {
      const q = state.searchQuery.toLowerCase().trim();
      list = list.filter((i) =>
        i.title.toLowerCase().includes(q) ||
        (i.description && i.description.toLowerCase().includes(q)) ||
        String(i.id).includes(q) ||
        (i.tags && i.tags.some((t) => t.toLowerCase().includes(q)))
      );
    }

    if (state.priorityFilter !== 'all') {
      list = list.filter((i) => i.priority === state.priorityFilter);
    }

    if (state.tagFilter !== 'all') {
      list = list.filter((i) => i.tags && i.tags.includes(state.tagFilter));
    }

    list = [...list].sort((a, b) => {
      if (state.sortBy === 'newest') return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
      if (state.sortBy === 'oldest') return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
      if (state.sortBy === 'progress-desc') return (b.progress || 0) - (a.progress || 0);
      if (state.sortBy === 'progress-asc') return (a.progress || 0) - (b.progress || 0);
      if (state.sortBy === 'priority') {
        const order = { p0: 0, p1: 1, p2: 2, p3: 3 };
        return (order[a.priority] ?? 2) - (order[b.priority] ?? 2);
      }
      return 0;
    });

    document.getElementById('showingCountText').innerText = `共 ${list.length} 筆`;

    if (list.length === 0) {
      container.innerHTML = `
        <li class="empty-state">
          <svg class="empty-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
            <circle cx="12" cy="12" r="10"></circle><path d="m15 9-6 6"></path><path d="m9 9 6 6"></path>
          </svg>
          <h3>尚無符合條件的 Issue</h3>
          <p>${state.statusFilter === 'open' ? '目前沒有待處理的 Issue，做得好！' : '尚未有已結案的 Issue 記錄。'}</p>
          <button type="button" class="btn btn-primary" id="btnEmptyNewIssue">+ 立即建立 Issue</button>
        </li>
      `;
      const btnEmpty = document.getElementById('btnEmptyNewIssue');
      if (btnEmpty) btnEmpty.addEventListener('click', openNewIssueModal);
      return;
    }

    let html = '';
    list.forEach((issue) => {
      const proj = state.projects.find((p) => p.id === issue.projectId);
      const prio = PRIORITY_MAP[issue.priority] || PRIORITY_MAP.p2;
      const progress = issue.progress || 0;
      const updatesCount = (issue.updates && issue.updates.length) || 0;

      let tagsHtml = '';
      (issue.tags || []).forEach((t) => {
        const col = getTagColors(t);
        tagsHtml += `
          <span class="tag-badge" style="background-color: ${col.bg}; color: ${col.text};">
            ${escapeHtml(t)}
          </span>
        `;
      });

      html += `
        <li class="issue-item" data-id="${issue.id}">
          <div class="issue-left">
            <div class="issue-status-icon">
              ${issue.status === 'open' ? `
                <svg class="icon-open" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
                  <circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="3"></circle>
                </svg>
              ` : `
                <svg class="icon-closed" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline>
                </svg>
              `}
            </div>

            <div class="issue-info">
              <div class="issue-title-line">
                <span class="issue-title">${escapeHtml(issue.title)}</span>
                <span class="priority-badge ${prio.cls}">${prio.short}</span>
                ${tagsHtml}
              </div>

              <div class="issue-meta-line">
                <span class="meta-id">#${issue.id}</span>
                ${state.activeProjectId === 'all' && proj ? `
                  <span style="display: inline-flex; align-items: center; gap: 4px; font-weight: 500;">
                    <span class="project-color-dot" style="background-color: ${proj.color || '#0969da'};"></span>
                    <span>${escapeHtml(proj.name)}</span>
                  </span>
                  <span>•</span>
                ` : ''}
                <span>由 Owner 建立於 ${formatRelativeTime(issue.createdAt)}</span>
                ${updatesCount > 0 ? `<span>• 最後更新：${formatRelativeTime(issue.updatedAt)}</span>` : ''}
              </div>
            </div>
          </div>

          <div class="issue-right">
            <div class="progress-cell">
              <div class="progress-text">${progress}%</div>
              <div class="progress-bar-mini">
                <div class="progress-bar-fill ${progress === 100 ? 'complete' : ''}" style="width: ${progress}%;"></div>
              </div>
            </div>

            <div class="comments-count" title="進度回報次數">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>
              <span>${updatesCount}</span>
            </div>
          </div>
        </li>
      `;
    });

    container.innerHTML = html;

    container.querySelectorAll('.issue-item').forEach((item) => {
      item.addEventListener('click', () => {
        const id = parseInt(item.getAttribute('data-id'), 10);
        openIssueDetail(id);
      });
    });
  }

  function renderApp() {
    renderSourceStatusAndOnboarding();
    renderProjectTabs();
    renderProjectBanner();
    updateTagFilterDropdown();
    renderIssueList();
  }

  // ==========================================================================
  // Issue Detail Modal & Progress Updates
  // ==========================================================================

  function openIssueDetail(issueId) {
    state.selectedIssueId = issueId;
    const issue = state.issues.find((i) => i.id === issueId);
    if (!issue) return;

    const modal = document.getElementById('modalIssueDetail');
    const header = document.getElementById('detailHeader');
    const main = document.getElementById('detailMainContent');
    const sidebar = document.getElementById('detailSidebarContent');

    header.innerHTML = `
      <div style="display: flex; align-items: flex-start; justify-content: space-between; gap: 16px;">
        <div style="flex: 1;">
          <div class="detail-title-line" id="detailTitleContainer">
            <h2 class="detail-title" id="detailTitleText">${escapeHtml(issue.title)}</h2>
            <span class="detail-id">#${issue.id}</span>
            <button type="button" class="detail-edit-btn" id="btnEditTitleInline" title="編輯標題">編輯</button>
          </div>
          <div class="detail-status-bar">
            <span class="state-badge-pill ${issue.status}">
              ${issue.status === 'open' ? `
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="3"></circle></svg>
                <span>Open (待處理)</span>
              ` : `
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                <span>Closed (已結案)</span>
              `}
            </span>
            <span class="detail-meta-text">由 Owner 建立於 ${formatDateTime(issue.createdAt)} • ${(issue.updates && issue.updates.length) || 0} 則進度回報</span>
          </div>
        </div>

        <div style="display: flex; align-items: center; gap: 8px;">
          <button type="button" class="btn ${issue.status === 'open' ? 'btn-purple' : 'btn-primary'}" id="btnToggleIssueStatus">
            ${issue.status === 'open' ? '關閉 Issue (Close)' : '重新開啟 Issue (Reopen)'}
          </button>
          <button type="button" class="btn-close" id="btnCloseDetailModal">&times;</button>
        </div>
      </div>
    `;

    let timelineHtml = '';
    if (issue.updates && issue.updates.length > 0) {
      issue.updates.forEach((u, idx) => {
        timelineHtml += `
          <div class="timeline-item">
            <div class="timeline-dot"></div>
            <div class="timeline-card">
              <div class="timeline-card-header">
                <span><strong>進度回報 #${idx + 1}</strong> • ${formatDateTime(u.timestamp)} (${formatRelativeTime(u.timestamp)})</span>
                <span class="priority-badge ${u.percentage === 100 ? 'priority-p2' : 'priority-p1'}">完成度 ${u.percentage}%</span>
              </div>
              <div class="timeline-note">${escapeHtml(u.note)}</div>
            </div>
          </div>
        `;
      });
    }

    main.innerHTML = `
      <div class="comment-card">
        <div class="comment-header">
          <div>
            <span class="comment-author">Owner</span>
            <span> • ${formatRelativeTime(issue.createdAt)}</span>
          </div>
          <button type="button" class="detail-edit-btn" id="btnEditDescInline">編輯說明</button>
        </div>
        <div class="comment-body" id="detailDescBody">${issue.description ? escapeHtml(issue.description) : '<span style="color: #8c959f; font-style: italic;">無詳細說明</span>'}</div>
      </div>

      <div class="timeline-section">
        <div style="font-size: 13px; font-weight: 700; color: #57606a; margin-bottom: 12px; text-transform: uppercase;">
          進度回報歷程 (${(issue.updates && issue.updates.length) || 0})
        </div>
        ${timelineHtml}
      </div>

      <div class="update-form-box">
        <div class="update-form-title">
          <h4>提交、更新當下的進度</h4>
          <span style="font-size: 13px; color: #238636; font-weight: 700;" id="updateProgressDisplay">${issue.progress || 0}%</span>
        </div>

        <div class="range-wrap">
          <input type="range" id="updateProgressRange" min="0" max="100" step="5" value="${issue.progress || 0}">
        </div>

        <div class="presets-row">
          <span style="font-size: 11px; color: #656d76; align-self: center;">快捷設定：</span>
          <button type="button" class="preset-chip" data-pct="25">25%</button>
          <button type="button" class="preset-chip" data-pct="50">50%</button>
          <button type="button" class="preset-chip" data-pct="75">75%</button>
          <button type="button" class="preset-chip" data-pct="100">100% (完成)</button>
        </div>

        <textarea id="updateProgressNote" class="form-textarea" rows="3" placeholder="簡述本次更新內容、測試結果或進度說明..."></textarea>

        <div style="display: flex; align-items: center; justify-content: space-between; margin-top: 12px; flex-wrap: wrap; gap: 8px;">
          <label style="font-size: 13px; color: #57606a; cursor: pointer; display: flex; align-items: center; gap: 6px;">
            <input type="checkbox" id="checkCloseOnUpdate" ${issue.progress === 100 ? 'checked' : ''} />
            <span>更新進度並同時結案 (Close Issue)</span>
          </label>
          <button type="button" class="btn btn-primary" id="btnSubmitProgressUpdate">
            提交進度更新
          </button>
        </div>
      </div>
    `;

    let tagsCloudHtml = '';
    (issue.tags || []).forEach((t) => {
      const col = getTagColors(t);
      tagsCloudHtml += `
        <span class="tag-removable" style="background-color: ${col.bg}; color: ${col.text};">
          <span>${escapeHtml(t)}</span>
          <span class="tag-remove-x" data-tag="${escapeHtml(t)}">&times;</span>
        </span>
      `;
    });

    let projectOptionsHtml = '';
    state.projects.forEach((p) => {
      projectOptionsHtml += `<option value="${escapeHtml(p.id)}" ${p.id === issue.projectId ? 'selected' : ''}>${escapeHtml(p.name)}</option>`;
    });

    sidebar.innerHTML = `
      <div class="sidebar-block">
        <div class="sidebar-title">優先層級 (Priority)</div>
        <div class="priority-selector">
          <button type="button" class="priority-radio-btn ${issue.priority === 'p0' ? 'selected' : ''}" data-prio="p0">P0</button>
          <button type="button" class="priority-radio-btn ${issue.priority === 'p1' ? 'selected' : ''}" data-prio="p1">P1</button>
          <button type="button" class="priority-radio-btn ${issue.priority === 'p2' ? 'selected' : ''}" data-prio="p2">P2</button>
          <button type="button" class="priority-radio-btn ${issue.priority === 'p3' ? 'selected' : ''}" data-prio="p3">P3</button>
        </div>
      </div>

      <div class="sidebar-block">
        <div class="sidebar-title">標籤 / 關鍵字 (Tags)</div>
        <div class="tags-cloud" id="sidebarTagsContainer">
          ${tagsCloudHtml || '<span style="font-size: 12px; color: #8c959f; font-style: italic;">尚無設定標籤</span>'}
        </div>
        <div class="add-tag-inline">
          <input type="text" id="inputAddTagInline" class="add-tag-input" placeholder="新增標籤 (Enter)..." />
          <button type="button" class="btn" id="btnAddTagInlineBtn" style="padding: 4px 8px; font-size: 12px;">+ 加</button>
        </div>
      </div>

      <div class="sidebar-block">
        <div class="sidebar-title">所屬專案 (Project)</div>
        <select id="selectDetailProject" class="form-select" style="font-size: 13px; padding: 6px;">
          ${projectOptionsHtml}
        </select>
      </div>

      <div class="sidebar-block">
        <div class="sidebar-title">
          <span>當前完成進度</span>
          <strong style="color: #0969da; font-size: 14px;">${issue.progress || 0}%</strong>
        </div>
        <div class="progress-track" style="margin-top: 6px;">
          <div class="progress-fill ${issue.progress === 100 ? 'complete' : ''}" style="width: ${issue.progress || 0}%;"></div>
        </div>
      </div>

      <div class="sidebar-block" style="font-size: 12px; color: #656d76; border-bottom: none;">
        <div style="margin-bottom: 4px;">建立：${formatDateTime(issue.createdAt)}</div>
        <div style="margin-bottom: 4px;">更新：${formatDateTime(issue.updatedAt)}</div>
        ${issue.closedAt ? `<div style="color: #8250df;">結案：${formatDateTime(issue.closedAt)}</div>` : ''}
      </div>

      <div>
        <button type="button" class="btn btn-danger-outline" id="btnDeleteIssue" style="width: 100%;">
          刪除此 Issue
        </button>
      </div>
    `;

    // Close Modal
    document.getElementById('btnCloseDetailModal').addEventListener('click', () => {
      modal.style.display = 'none';
    });

    // Inline Edit Title
    document.getElementById('btnEditTitleInline').addEventListener('click', () => {
      const container = document.getElementById('detailTitleContainer');
      container.innerHTML = `
        <input type="text" id="inputEditTitleInline" class="form-input" value="${escapeHtml(issue.title)}" style="flex: 1;" />
        <button type="button" class="btn btn-primary" id="btnSaveTitleInline" style="padding: 5px 10px;">儲存</button>
        <button type="button" class="btn" id="btnCancelTitleInline" style="padding: 5px 10px;">取消</button>
      `;
      const inputEl = document.getElementById('inputEditTitleInline');
      inputEl.focus();
      document.getElementById('btnSaveTitleInline').addEventListener('click', () => {
        const val = inputEl.value.trim();
        if (!val) return;
        issue.title = val;
        issue.updatedAt = new Date().toISOString();
        saveState();
        showToast('標題已更新');
        openIssueDetail(issue.id);
        renderApp();
      });
      document.getElementById('btnCancelTitleInline').addEventListener('click', () => openIssueDetail(issue.id));
    });

    // Inline Edit Description
    document.getElementById('btnEditDescInline').addEventListener('click', () => {
      const descBody = document.getElementById('detailDescBody');
      descBody.innerHTML = `
        <textarea id="textareaEditDescInline" class="form-textarea" rows="4">${escapeHtml(issue.description || '')}</textarea>
        <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 8px;">
          <button type="button" class="btn" id="btnCancelDescInline">取消</button>
          <button type="button" class="btn btn-primary" id="btnSaveDescInline">儲存說明</button>
        </div>
      `;
      document.getElementById('btnSaveDescInline').addEventListener('click', () => {
        issue.description = document.getElementById('textareaEditDescInline').value.trim();
        issue.updatedAt = new Date().toISOString();
        saveState();
        showToast('說明已更新');
        openIssueDetail(issue.id);
        renderApp();
      });
      document.getElementById('btnCancelDescInline').addEventListener('click', () => openIssueDetail(issue.id));
    });

    // Toggle Status (Close / Reopen)
    document.getElementById('btnToggleIssueStatus').addEventListener('click', () => {
      const isClosing = issue.status === 'open';
      const now = new Date().toISOString();
      issue.status = isClosing ? 'closed' : 'open';
      issue.updatedAt = now;
      issue.updates = issue.updates || [];

      if (isClosing) {
        issue.closedAt = now;
        issue.progress = 100;
        issue.updates.push({
          id: `up-${Date.now()}`,
          timestamp: now,
          percentage: 100,
          note: 'Issue 已由 Owner 結案 (Closed)。',
        });
      } else {
        issue.closedAt = null;
        issue.updates.push({
          id: `up-${Date.now()}`,
          timestamp: now,
          percentage: issue.progress || 0,
          note: 'Issue 已由 Owner 重新開啟 (Reopened)。',
        });
      }
      saveState();
      showToast(`Issue #${issue.id} 狀態已變更為 ${issue.status === 'open' ? 'Open' : 'Closed'}`);
      openIssueDetail(issue.id);
      renderApp();
    });

    // Progress slider & presets
    const range = document.getElementById('updateProgressRange');
    const disp = document.getElementById('updateProgressDisplay');
    const closeCheck = document.getElementById('checkCloseOnUpdate');

    range.addEventListener('input', () => {
      disp.innerText = `${range.value}%`;
      if (parseInt(range.value, 10) === 100) {
        closeCheck.checked = true;
      }
    });

    main.querySelectorAll('.preset-chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const pct = parseInt(chip.getAttribute('data-pct'), 10);
        range.value = pct;
        disp.innerText = `${pct}%`;
        if (pct === 100) closeCheck.checked = true;
      });
    });

    // Submit progress update
    document.getElementById('btnSubmitProgressUpdate').addEventListener('click', () => {
      const note = document.getElementById('updateProgressNote').value.trim();
      if (!note) {
        alert('請填寫進度說明與備註！');
        return;
      }
      const pct = parseInt(range.value, 10);
      const shouldClose = closeCheck.checked || pct === 100;
      const now = new Date().toISOString();

      issue.progress = shouldClose && pct < 100 ? 100 : pct;
      issue.updatedAt = now;
      if (shouldClose) {
        issue.status = 'closed';
        issue.closedAt = now;
      }

      issue.updates = issue.updates || [];
      issue.updates.push({
        id: `up-${Date.now()}`,
        timestamp: now,
        percentage: issue.progress,
        note,
      });

      saveState();
      showToast(`Issue #${issue.id} 進度已更新為 ${issue.progress}%`);
      openIssueDetail(issue.id);
      renderApp();
    });

    // Priority buttons
    sidebar.querySelectorAll('.priority-radio-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        issue.priority = btn.getAttribute('data-prio');
        issue.updatedAt = new Date().toISOString();
        saveState();
        openIssueDetail(issue.id);
        renderApp();
      });
    });

    // Remove tag
    sidebar.querySelectorAll('.tag-remove-x').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const tagToRemove = btn.getAttribute('data-tag');
        issue.tags = (issue.tags || []).filter((t) => t !== tagToRemove);
        issue.updatedAt = new Date().toISOString();
        saveState();
        openIssueDetail(issue.id);
        renderApp();
      });
    });

    // Add tag
    const addTagInput = document.getElementById('inputAddTagInline');
    const addTagAction = () => {
      const val = addTagInput.value.trim().toLowerCase();
      if (!val) return;
      issue.tags = issue.tags || [];
      if (!issue.tags.includes(val)) {
        issue.tags.push(val);
        issue.updatedAt = new Date().toISOString();
        saveState();
        openIssueDetail(issue.id);
        renderApp();
      }
      addTagInput.value = '';
    };

    document.getElementById('btnAddTagInlineBtn').addEventListener('click', addTagAction);
    addTagInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addTagAction();
      }
    });

    // Change project
    document.getElementById('selectDetailProject').addEventListener('change', (e) => {
      issue.projectId = e.target.value;
      issue.updatedAt = new Date().toISOString();
      saveState();
      showToast(`Issue #${issue.id} 已移至新專案`);
      openIssueDetail(issue.id);
      renderApp();
    });

    // Delete issue
    document.getElementById('btnDeleteIssue').addEventListener('click', () => {
      if (confirm(`確定要刪除 Issue #${issue.id}「${issue.title}」嗎？此動作不可復原。`)) {
        state.issues = state.issues.filter((i) => i.id !== issue.id);
        saveState();
        modal.style.display = 'none';
        showToast(`已刪除 Issue #${issue.id}`);
        renderApp();
      }
    });

    modal.style.display = 'flex';
  }

  // ==========================================================================
  // New Issue Modal
  // ==========================================================================

  function openNewIssueModal() {
    if (state.projects.length === 0) {
      alert('請先建立至少一個專案（或先上傳 XLSX 檔案）再新增 Issue！');
      openProjectModal(null);
      return;
    }

    const modal = document.getElementById('modalNewIssue');
    const select = document.getElementById('newIssueProject');

    let html = '';
    state.projects.forEach((p) => {
      const isSelected = state.activeProjectId !== 'all' && state.activeProjectId === p.id;
      html += `<option value="${escapeHtml(p.id)}" ${isSelected ? 'selected' : ''}>${escapeHtml(p.name)}</option>`;
    });
    select.innerHTML = html;

    document.getElementById('newIssueTitle').value = '';
    document.getElementById('newIssueDesc').value = '';
    document.getElementById('newIssueTagsInput').value = '';
    document.getElementById('newIssueProgressRange').value = 0;
    document.getElementById('newIssueProgressLabel').innerText = '0%';
    state.newIssueSelectedPriority = 'p2';

    document.querySelectorAll('#newIssuePrioritySelector .priority-radio-btn').forEach((btn) => {
      btn.className = `priority-radio-btn ${btn.getAttribute('data-priority') === 'p2' ? 'selected' : ''}`;
    });

    modal.style.display = 'flex';
    document.getElementById('newIssueTitle').focus();
  }

  function initNewIssueForm() {
    const modal = document.getElementById('modalNewIssue');
    document.getElementById('btnCloseNewIssueModal').addEventListener('click', () => (modal.style.display = 'none'));
    document.getElementById('btnCancelNewIssue').addEventListener('click', () => (modal.style.display = 'none'));

    document.querySelectorAll('#newIssuePrioritySelector .priority-radio-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#newIssuePrioritySelector .priority-radio-btn').forEach((b) => b.classList.remove('selected'));
        btn.classList.add('selected');
        state.newIssueSelectedPriority = btn.getAttribute('data-priority');
      });
    });

    const range = document.getElementById('newIssueProgressRange');
    const label = document.getElementById('newIssueProgressLabel');
    range.addEventListener('input', () => {
      label.innerText = `${range.value}%`;
    });

    document.getElementById('formNewIssue').addEventListener('submit', (e) => {
      e.preventDefault();
      const title = document.getElementById('newIssueTitle').value.trim();
      if (!title) return;

      const projectId = document.getElementById('newIssueProject').value;
      const desc = document.getElementById('newIssueDesc').value.trim();
      const tagsRaw = document.getElementById('newIssueTagsInput').value;
      const tags = tagsRaw.split(/[,;\n\s]+/).map((t) => t.trim().toLowerCase()).filter(Boolean);
      const progress = parseInt(range.value, 10);

      const maxId = state.issues.reduce((m, i) => Math.max(m, i.id || 0), 0);
      const nextId = maxId + 1;
      const now = new Date().toISOString();

      const newIssue = {
        id: nextId,
        projectId,
        title,
        description: desc,
        status: progress === 100 ? 'closed' : 'open',
        priority: state.newIssueSelectedPriority,
        tags,
        progress,
        createdAt: now,
        updatedAt: now,
        closedAt: progress === 100 ? now : null,
        updates: [
          {
            id: `up-init-${Date.now()}`,
            timestamp: now,
            percentage: progress,
            note: progress > 0 ? `初始建立 Issue，設定進度為 ${progress}%。` : 'Issue 建立完成。',
          },
        ],
      };

      state.issues.unshift(newIssue);
      if (state.dataSourceMode === 'empty') {
        state.dataSourceMode = 'local-cache';
      }
      saveState();
      modal.style.display = 'none';
      showToast(`已建立 Issue #${nextId}「${title}」`);
      renderApp();
    });
  }

  // ==========================================================================
  // Project Modal (New / Edit)
  // ==========================================================================

  function openProjectModal(existingProject = null) {
    const modal = document.getElementById('modalProject');
    const title = document.getElementById('projectModalTitle');
    const idInput = document.getElementById('editProjectId');
    const nameInput = document.getElementById('projectNameInput');
    const descInput = document.getElementById('projectDescInput');
    const picker = document.getElementById('projectColorPicker');

    if (existingProject) {
      title.innerText = '編輯專案資訊';
      idInput.value = existingProject.id;
      nameInput.value = existingProject.name;
      descInput.value = existingProject.description || '';
      state.selectedProjectColor = existingProject.color || PROJECT_PRESET_COLORS[0];
    } else {
      title.innerText = '新增專案 (Project)';
      idInput.value = '';
      nameInput.value = '';
      descInput.value = '';
      state.selectedProjectColor = PROJECT_PRESET_COLORS[state.projects.length % PROJECT_PRESET_COLORS.length];
    }

    const renderColorPicker = () => {
      let html = '';
      PROJECT_PRESET_COLORS.forEach((c) => {
        const isSelected = c === state.selectedProjectColor;
        html += `
          <div class="color-dot-choice" data-color="${c}" style="width: 26px; height: 26px; border-radius: 50%; background-color: ${c}; cursor: pointer; border: ${isSelected ? '3px solid #1f2328' : '2px solid transparent'};"></div>
        `;
      });
      picker.innerHTML = html;

      picker.querySelectorAll('.color-dot-choice').forEach((dot) => {
        dot.addEventListener('click', () => {
          state.selectedProjectColor = dot.getAttribute('data-color');
          renderColorPicker();
        });
      });
    };

    renderColorPicker();
    modal.style.display = 'flex';
    nameInput.focus();
  }

  function initProjectForm() {
    const modal = document.getElementById('modalProject');
    document.getElementById('btnCloseProjectModal').addEventListener('click', () => (modal.style.display = 'none'));
    document.getElementById('btnCancelProject').addEventListener('click', () => (modal.style.display = 'none'));

    document.getElementById('formProject').addEventListener('submit', (e) => {
      e.preventDefault();
      const id = document.getElementById('editProjectId').value;
      const name = document.getElementById('projectNameInput').value.trim();
      const desc = document.getElementById('projectDescInput').value.trim();
      if (!name) return;

      if (id) {
        const proj = state.projects.find((p) => p.id === id);
        if (proj) {
          proj.name = name;
          proj.description = desc;
          proj.color = state.selectedProjectColor;
          saveState();
          showToast(`專案「${name}」已更新`);
        }
      } else {
        const newId = `proj-${Date.now()}`;
        state.projects.push({
          id: newId,
          name,
          description: desc,
          color: state.selectedProjectColor,
          createdAt: new Date().toISOString(),
        });
        state.activeProjectId = newId;
        if (state.dataSourceMode === 'empty') {
          state.dataSourceMode = 'local-cache';
        }
        saveState();
        showToast(`已新增專案「${name}」`);
      }

      modal.style.display = 'none';
      renderApp();
    });
  }

  // ==========================================================================
  // Excel (XLSX) Export Engine
  // ==========================================================================

  function exportProjectsAndIssuesToXlsx(projectsList, issuesList, filename = 'issues.xlsx') {
    const sheets = [];

    // Sheet 1: Projects Summary
    const projHeaders = [
      '專案名稱 (Project Name)',
      '專案ID (Project ID)',
      '說明 (Description)',
      '待處理 (Open Issues)',
      '已結案 (Closed Issues)',
      '總 Issue 數 (Total)',
      '平均完成度 (Avg Progress %)',
      '建立時間 (Created At)',
    ];

    const projRows = [projHeaders];
    projectsList.forEach((p) => {
      const pIssues = issuesList.filter((i) => i.projectId === p.id);
      const open = pIssues.filter((i) => i.status === 'open').length;
      const closed = pIssues.filter((i) => i.status === 'closed').length;
      const avg = pIssues.length > 0
        ? Math.round(pIssues.reduce((acc, i) => acc + (i.progress || 0), 0) / pIssues.length)
        : 0;

      projRows.push([
        p.name,
        p.id,
        p.description || '',
        open,
        closed,
        pIssues.length,
        `${avg}%`,
        formatDateTime(p.createdAt),
      ]);
    });
    sheets.push({ name: 'Projects', rows: projRows });

    // Sheet 2: Master Issues (with lossless UpdatesData_JSON)
    const issueHeaders = [
      '所屬專案 (Project)',
      '專案ID (Project ID)',
      'Issue編號 (ID)',
      '標題 (Title)',
      '狀態 (Status)',
      '優先等級 (Priority)',
      '標籤 (Tags)',
      '進度% (Progress)',
      '詳細說明 (Description)',
      '建立時間 (Created At)',
      '最後更新 (Updated At)',
      '結案時間 (Closed At)',
      '回報次數 (Updates Count)',
      '最新進度備註 (Latest Note)',
      'UpdatesData_JSON',
    ];

    const issueRows = [issueHeaders];
    issuesList.forEach((issue) => {
      const proj = projectsList.find((p) => p.id === issue.projectId);
      const updatesArr = issue.updates || [];
      const lastUpdate = updatesArr[updatesArr.length - 1] || null;

      issueRows.push([
        proj ? proj.name : issue.projectId,
        issue.projectId,
        issue.id,
        issue.title,
        issue.status,
        (issue.priority || 'p2').toUpperCase(),
        (issue.tags || []).join(', '),
        issue.progress || 0,
        issue.description || '',
        formatDateTime(issue.createdAt),
        formatDateTime(issue.updatedAt),
        issue.closedAt ? formatDateTime(issue.closedAt) : '',
        updatesArr.length,
        lastUpdate ? lastUpdate.note : '',
        JSON.stringify(updatesArr),
      ]);
    });
    sheets.push({ name: 'Issues', rows: issueRows });

    // Sheet 3: Human-Readable Progress History
    const historyRows = [['Issue編號 (Issue ID)', '所屬專案 (Project)', 'Issue標題 (Title)', '進度% (Progress %)', '回報時間 (Timestamp)', '進度說明與備註 (Notes)']];
    issuesList.forEach((issue) => {
      const proj = projectsList.find((p) => p.id === issue.projectId);
      (issue.updates || []).forEach((u) => {
        historyRows.push([
          `#${issue.id}`,
          proj ? proj.name : issue.projectId,
          issue.title,
          `${u.percentage}%`,
          formatDateTime(u.timestamp),
          u.note || '',
        ]);
      });
    });
    sheets.push({ name: 'Progress History', rows: historyRows });

    // Sheets 4+: Individual Sheet per Project
    projectsList.forEach((proj) => {
      const pIssues = issuesList.filter((i) => i.projectId === proj.id);
      const singleSheetRows = [
        ['編號 (ID)', '標題 (Title)', '狀態 (Status)', '優先級 (Priority)', '標籤 (Tags)', '進度% (Progress)', '說明 (Description)', '更新時間', '最新備註'],
      ];
      pIssues.forEach((issue) => {
        const updatesArr = issue.updates || [];
        const lastUpdate = updatesArr[updatesArr.length - 1] || null;
        singleSheetRows.push([
          issue.id,
          issue.title,
          issue.status,
          (issue.priority || 'p2').toUpperCase(),
          (issue.tags || []).join(', '),
          issue.progress || 0,
          issue.description || '',
          formatDateTime(issue.updatedAt),
          lastUpdate ? lastUpdate.note : '',
        ]);
      });
      sheets.push({ name: proj.name, rows: singleSheetRows });
    });

    const xlsxBuffer = buildXlsxBuffer(sheets);
    downloadFile(xlsxBuffer, filename);
    showToast(`已匯出 ${filename}（可放置於 ./data/issues.xlsx 下次自動載入）`);
  }

  function handleExportXlsx() {
    if (state.projects.length === 0) {
      alert('目前工作區尚無專案可匯出，請先建立專案或載入範例資料！');
      return;
    }
    exportProjectsAndIssuesToXlsx(state.projects, state.issues, 'issues.xlsx');
  }

  // ==========================================================================
  // Manual File Upload & Global Drag-Drop
  // ==========================================================================

  async function processUploadedFile(file) {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const { projects: newProjects, issues: newIssues } = await convertXlsxBufferToData(arrayBuffer);

      const previewSection = document.getElementById('uploadPreviewSection');
      previewSection.style.display = 'block';

      let projectsPills = '';
      newProjects.forEach((p) => {
        const count = newIssues.filter((i) => i.projectId === p.id).length;
        projectsPills += `
          <div style="padding: 6px 10px; background: #ffffff; border: 1px solid var(--border-subtle); border-radius: var(--radius-sm); font-size: 12px; display: flex; align-items: center; justify-content: space-between;">
            <span><strong>${escapeHtml(p.name)}</strong></span>
            <span class="badge-tag">${count} 個 Issue</span>
          </div>
        `;
      });

      previewSection.innerHTML = `
        <div style="padding: 12px; background: #dafbe1; border: 1px solid #aceebb; border-radius: var(--radius-md); font-size: 13px; color: #1a7f37; margin-bottom: 12px;">
          <strong>檔案解析成功 (${escapeHtml(file.name)})！</strong> 偵測到 ${newProjects.length} 個專案，共 ${newIssues.length} 筆 Issue。
        </div>

        <div style="font-size: 12px; font-weight: 700; color: #57606a; margin-bottom: 6px;">解析出的專案分頁：</div>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px; max-height: 140px; overflow-y: auto; margin-bottom: 16px;">
          ${projectsPills}
        </div>

        <div style="display: flex; gap: 8px; justify-content: flex-end;">
          ${state.projects.length > 0 ? '<button type="button" class="btn" id="btnConfirmMergeImport">合併至現有資料</button>' : ''}
          <button type="button" class="btn btn-primary" id="btnConfirmReplaceImport">載入專案工作區</button>
        </div>
      `;

      document.getElementById('btnConfirmReplaceImport').addEventListener('click', () => {
        state.projects = newProjects;
        state.issues = newIssues;
        state.activeProjectId = 'all';
        state.dataSourceMode = 'uploaded-file';
        saveState();
        document.getElementById('modalUpload').style.display = 'none';
        showToast(`成功載入 ${newProjects.length} 個專案與 ${newIssues.length} 個 Issue！`);
        renderApp();
      });

      const mergeBtn = document.getElementById('btnConfirmMergeImport');
      if (mergeBtn) {
        mergeBtn.addEventListener('click', () => {
          const existingByName = new Map();
          state.projects.forEach((p) => existingByName.set(p.name.toLowerCase(), p));

          const idRemap = new Map();
          newProjects.forEach((np) => {
            const match = existingByName.get(np.name.toLowerCase());
            if (match) {
              idRemap.set(np.id, match.id);
            } else {
              state.projects.push(np);
              idRemap.set(np.id, np.id);
            }
          });

          let maxId = state.issues.reduce((m, i) => Math.max(m, i.id || 0), 0);
          newIssues.forEach((ni) => {
            ni.id = ++maxId;
            ni.projectId = idRemap.get(ni.projectId) || ni.projectId;
            state.issues.push(ni);
          });

          state.dataSourceMode = 'uploaded-file';
          saveState();
          document.getElementById('modalUpload').style.display = 'none';
          showToast(`已成功合併 ${newProjects.length} 個專案與 ${newIssues.length} 筆 Issue！`);
          renderApp();
        });
      }
    } catch (err) {
      console.error(err);
      alert(`解析 Excel 檔案失敗: ${err.message || '格式異常'}`);
    }
  }

  function initUploadModal() {
    const modal = document.getElementById('modalUpload');
    document.getElementById('btnOpenUploadModal').addEventListener('click', () => {
      document.getElementById('uploadPreviewSection').style.display = 'none';
      modal.style.display = 'flex';
    });
    document.getElementById('btnCloseUploadModal').addEventListener('click', () => {
      modal.style.display = 'none';
    });

    const dropZone = document.getElementById('modalDropZone');
    const fileInput = document.getElementById('modalFileInput');

    dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
      dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) processUploadedFile(file);
    });

    document.getElementById('btnSelectFileTrigger').addEventListener('click', () => {
      fileInput.value = '';
      fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) processUploadedFile(file);
    });

    // Global drag & drop anywhere on page
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault();
      const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (file && /\.(xlsx|xls)$/i.test(file.name)) {
        document.getElementById('uploadPreviewSection').style.display = 'none';
        modal.style.display = 'flex';
        processUploadedFile(file);
      }
    });
  }

  // ==========================================================================
  // Boot Application
  // ==========================================================================

  async function initApp() {
    // 1. Bind UI listeners immediately
    document.getElementById('searchInput').addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      renderIssueList();
    });

    document.getElementById('priorityFilter').addEventListener('change', (e) => {
      state.priorityFilter = e.target.value;
      renderIssueList();
    });

    document.getElementById('tagFilter').addEventListener('change', (e) => {
      state.tagFilter = e.target.value;
      renderIssueList();
    });

    document.getElementById('sortFilter').addEventListener('change', (e) => {
      state.sortBy = e.target.value;
      renderIssueList();
    });

    document.getElementById('tabOpenIssues').addEventListener('click', () => {
      state.statusFilter = 'open';
      renderIssueList();
    });

    document.getElementById('tabClosedIssues').addEventListener('click', () => {
      state.statusFilter = 'closed';
      renderIssueList();
    });

    document.getElementById('btnExportXlsx').addEventListener('click', handleExportXlsx);

    document.getElementById('btnDownloadSample').addEventListener('click', () => {
      const p = state.projects.length > 0 ? state.projects : INITIAL_PROJECTS;
      const i = state.projects.length > 0 ? state.issues : INITIAL_ISSUES;
      exportProjectsAndIssuesToXlsx(p, i, 'issues.xlsx');
    });

    document.getElementById('btnResetData').addEventListener('click', () => {
      if (confirm('確定要載入系統範例資料嗎？現有工作區資料將會被覆蓋。')) {
        state.projects = JSON.parse(JSON.stringify(INITIAL_PROJECTS));
        state.issues = JSON.parse(JSON.stringify(INITIAL_ISSUES));
        state.activeProjectId = 'all';
        state.dataSourceMode = 'local-cache';
        saveState();
        showToast('已載入系統範例資料！');
        renderApp();
      }
    });

    document.getElementById('btnOpenNewIssue').addEventListener('click', openNewIssueModal);
    document.getElementById('btnQuickNewIssue').addEventListener('click', openNewIssueModal);
    document.getElementById('btnOpenNewProject').addEventListener('click', () => openProjectModal(null));

    initNewIssueForm();
    initProjectForm();
    initUploadModal();

    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        document.getElementById('modalIssueDetail').style.display = 'none';
        document.getElementById('modalNewIssue').style.display = 'none';
        document.getElementById('modalProject').style.display = 'none';
        document.getElementById('modalUpload').style.display = 'none';
      }
    });

    // 2. Load saved state from localStorage first
    const stored = loadStoredState();

    // 3. Attempt auto-detection of designated XLSX file (./data/issues.xlsx or ./issues.xlsx)
    const autoCheck = await tryFetchDesignatedXlsx();
    state.isFileProtocolBlocked = !!autoCheck.fileProtocolBlocked;

    if (autoCheck.found) {
      state.detectedFilePath = autoCheck.path;
      const fileChanged = !stored || stored.lastAutoFileHash !== autoCheck.hash || stored.projects.length === 0;

      if (fileChanged) {
        try {
          const parsed = await convertXlsxBufferToData(autoCheck.buffer);
          state.projects = parsed.projects;
          state.issues = parsed.issues;
          state.lastAutoFileHash = autoCheck.hash;
          state.dataSourceMode = 'auto-file';
          saveState();
          showToast(`已自動載入並解析 ${autoCheck.path}！`);
        } catch (err) {
          console.error('Auto parse error:', err);
          if (stored && stored.projects.length > 0) {
            state.projects = stored.projects;
            state.issues = stored.issues;
            state.lastAutoFileHash = stored.lastAutoFileHash;
            state.dataSourceMode = 'local-cache';
          }
        }
      } else {
        // File hasn't changed since last load; keep user's local edits in localStorage
        state.projects = stored.projects;
        state.issues = stored.issues;
        state.lastAutoFileHash = stored.lastAutoFileHash;
        state.dataSourceMode = 'local-cache';
      }
    } else if (stored && stored.projects.length > 0) {
      // No file at designated location, restore localStorage session
      state.projects = stored.projects;
      state.issues = stored.issues;
      state.lastAutoFileHash = stored.lastAutoFileHash;
      state.dataSourceMode = 'local-cache';
    } else {
      // Neither designated file nor localStorage exists -> clean empty workspace onboarding
      state.projects = [];
      state.issues = [];
      state.dataSourceMode = 'empty';
    }

    renderApp();
  }

  window.addEventListener('DOMContentLoaded', initApp);
})();
