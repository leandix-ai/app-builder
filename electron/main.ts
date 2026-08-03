import { app, BrowserWindow, WebContentsView, ipcMain, dialog, Menu } from 'electron'
import * as path from 'path'
import * as fs from 'fs/promises'
import * as fsSync from 'fs'
import * as os from 'os'
import * as dns from 'dns'
import { execFile } from 'child_process'
import { autoUpdater } from 'electron-updater'
import { parseDocx } from '../../extension/src/extension/DocxParser'
import { serializeToDocx } from '../../extension/src/extension/DocxSerializer'
import { DEFAULT_PAGE_SETTINGS } from '../../extension/src/shared/constants'

// Renderer globals injected via contextBridge (preload.ts)
declare global {
  interface Window {
    leandixApp?: {
      dropFile: (filePath: string) => void
    }
  }
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

interface AppSettings {
  language: 'vi' | 'en'
  theme: 'dark' | 'light'
}

const DEFAULT_SETTINGS: AppSettings = { language: 'vi', theme: 'dark' }

function getSettingsPath(): string {
  return path.join(app.getPath('userData'), 'settings.json')
}

function loadSettings(): AppSettings {
  try {
    const raw = fsSync.readFileSync(getSettingsPath(), 'utf-8')
    const parsed = JSON.parse(raw) as Partial<AppSettings>
    return {
      language: parsed.language === 'en' ? 'en' : 'vi',
      theme: parsed.theme === 'light' ? 'light' : 'dark',
    }
  } catch {
    return { ...DEFAULT_SETTINGS }
  }
}

function saveSettings(settings: AppSettings): void {
  try {
    fsSync.writeFileSync(getSettingsPath(), JSON.stringify(settings, null, 2), 'utf-8')
  } catch {
    // ignore write errors
  }
}

// ---------------------------------------------------------------------------
// Recent Files
// ---------------------------------------------------------------------------

interface RecentFileEntry {
  filePath: string
  title: string
  lastOpenedAt: string // ISO string
}

const MAX_RECENT_FILES = 20

function getRecentFilesPath(): string {
  return path.join(app.getPath('userData'), 'recent-files.json')
}

function loadRecentFiles(): RecentFileEntry[] {
  try {
    const raw = fsSync.readFileSync(getRecentFilesPath(), 'utf-8')
    return JSON.parse(raw) as RecentFileEntry[]
  } catch {
    return []
  }
}

function saveRecentFiles(entries: RecentFileEntry[]): void {
  try {
    fsSync.writeFileSync(getRecentFilesPath(), JSON.stringify(entries, null, 2), 'utf-8')
  } catch {
    // ignore write errors
  }
}

function addRecentFile(filePath: string): void {
  let entries = loadRecentFiles()
  // Remove existing entry for this path
  entries = entries.filter((e) => e.filePath !== filePath)
  entries.unshift({
    filePath,
    title: path.basename(filePath),
    lastOpenedAt: new Date().toISOString(),
  })
  if (entries.length > MAX_RECENT_FILES) entries = entries.slice(0, MAX_RECENT_FILES)
  saveRecentFiles(entries)
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type Comment = {
  id: string
  author: string
  date: string
  text: string
  rangeStart: number
  rangeEnd: number
  replies?: Array<{ id: string; author: string; date: string; text: string }>
}

type FileExtras = {
  headerHtml?: string
  footerHtml?: string
  comments?: Comment[]
  pageSettings?: typeof DEFAULT_PAGE_SETTINGS
  defaultFont?: string
  defaultFontSize?: number
}

interface TabDocumentState {
  currentFilePath: string | null
  fileContentHtml: string
  fileExtras: FileExtras
  /**
   * DocxAst cached sau khi parseDocx() — dùng bởi serializeToDocx() khi save.
   * Không truyền qua IPC. Undefined nếu file chưa được parse (new/template).
   */
  originalAst?: import('../../extension/src/shared/types').DocxAst
  /**
   * Raw XML string của word/numbering.xml — inject lại khi save để preserve
   * numId và tránh numbering counter reset.
   */
  rawNumberingXml?: string
}

interface Tab {
  id: string
  type: 'editor' | 'pdf'
  title: string
  documentState: TabDocumentState
  isDirty: boolean
  view: WebContentsView
  /** HTML content to pre-load when the tab's webview sends 'ready' (templates) */
  pendingTemplateHtml?: string
}

interface TabStateSnapshot {
  id: string
  title: string
  type: 'editor' | 'pdf'
  isDirty: boolean
  isActive: boolean
}

// ---------------------------------------------------------------------------
// Globals
// ---------------------------------------------------------------------------

let homeWindow: BrowserWindow | null = null
let mainWindow: BrowserWindow | null = null
let atlasWindow: BrowserWindow | null = null

const tabs = new Map<string, Tab>()
let activeTabId: string | null = null
let tabBarView: WebContentsView | null = null

let pendingOpenFilePath: string | null = null

/** File path to open when the first editor window is created */
let pendingInitialFilePath: string | null = null

/** Pending PDF export resolve callbacks keyed by tabId */
const pendingPdfExports = new Map<string, (html: string) => void>()

const _initialSettings = loadSettings()
/** Current UI language */
let currentLanguage: 'vi' | 'en' = _initialSettings.language
/** Current UI theme */
let currentTheme: 'dark' | 'light' = _initialSettings.theme

// ---------------------------------------------------------------------------
// Layout Helpers
// ---------------------------------------------------------------------------

const TAB_BAR_HEIGHT = 36

function getTabBarBounds(win: BrowserWindow): { x: number; y: number; width: number; height: number } {
  const [width] = win.getContentSize()
  return { x: 0, y: 0, width, height: TAB_BAR_HEIGHT }
}

function getContentBounds(win: BrowserWindow): { x: number; y: number; width: number; height: number } {
  const [width, height] = win.getContentSize()
  return { x: 0, y: TAB_BAR_HEIGHT, width, height: Math.max(0, height - TAB_BAR_HEIGHT) }
}

function updateLayout(): void {
  if (!mainWindow || mainWindow.isDestroyed()) return
  if (tabBarView) tabBarView.setBounds(getTabBarBounds(mainWindow))
  if (activeTabId) {
    const tab = tabs.get(activeTabId)
    if (tab) tab.view.setBounds(getContentBounds(mainWindow))
  }
}

// ---------------------------------------------------------------------------
// Tab State
// ---------------------------------------------------------------------------

function pushTabState(): void {
  if (!tabBarView || tabBarView.webContents.isDestroyed()) return
  const snapshots: TabStateSnapshot[] = []
  for (const tab of tabs.values()) {
    snapshots.push({
      id: tab.id,
      title: tab.title,
      type: tab.type,
      isDirty: tab.isDirty,
      isActive: tab.id === activeTabId,
    })
  }
  tabBarView.webContents.send('tab-state', snapshots)
}

// ---------------------------------------------------------------------------
// Window Creation
// ---------------------------------------------------------------------------

/** One-shot DNS check — resolves quickly (typically < 2s on healthy connections). */
function checkNetwork(): Promise<boolean> {
  return new Promise((resolve) => {
    dns.lookup('atlas.leandix.com', (err) => resolve(!err))
  })
}

/**
 * Check network with a cancellable timeout.
 * Returns a Promise that resolves to `true` (online), `false` (offline/timeout),
 * or `'cancelled'` when the user hits Cancel.
 * `onCancel` is called back with a cancel function so the caller can wire it to a button.
 */
function checkNetworkCancellable(
  timeoutMs: number,
  onCancel: (cancelFn: () => void) => void,
): Promise<boolean | 'cancelled'> {
  return new Promise((resolve) => {
    let settled = false
    function finish(result: boolean | 'cancelled') {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve(result)
    }

    // Expose cancel to the caller
    onCancel(() => finish('cancelled'))

    const timer = setTimeout(() => finish(false), timeoutMs)

    dns.lookup('atlas.leandix.com', (err) => finish(!err))
  })
}

function createHomeWindow(): void {
  homeWindow = new BrowserWindow({
    width: 960,
    height: 640,
    minWidth: 720,
    minHeight: 500,
    resizable: true,
    center: true,
    title: 'Leandix Atlas',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: path.join(__dirname, 'home-preload.js'),
    },
  })

  homeWindow.loadFile(path.join(__dirname, '..', 'home-webview', 'index.html'))
  homeWindow.setMenuBarVisibility(false)

  // Home window drag & drop: dropping a file on the home screen opens the editor
  homeWindow.webContents.on('will-navigate', (event, url) => {
    if (!url.startsWith('file://')) return
    let filePath: string
    try {
      filePath = decodeURIComponent(new URL(url).pathname)
      if (process.platform === 'win32' && filePath.startsWith('/')) filePath = filePath.slice(1)
    } catch { return }
    if (!isSupportedDroppedFile(filePath)) return
    event.preventDefault()
    void handleDroppedFile(filePath)
  })

  homeWindow.webContents.once('did-finish-load', () => {
    // Send language and theme immediately — no network check needed at startup
    homeWindow?.webContents.send('home-language', currentLanguage)
    homeWindow?.webContents.send('home-theme', currentTheme)

    // Delay nhỏ để renderer kịp mount và đăng ký IPC listeners
    setTimeout(() => {
      void autoUpdater.checkForUpdates()
    }, 2000)
  })

  homeWindow.on('closed', () => {
    homeWindow = null
  })
}

function openEditorFromHome(filePath?: string): void {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.show()
    mainWindow.focus()
    homeWindow?.close()
    if (filePath) {
      const ext = path.extname(filePath).toLowerCase()
      if (ext === '.pdf') createPdfTab(filePath)
      else createEditorTab(filePath)
    }
    return
  }
  createWindow(filePath)
  homeWindow?.close()
}

function openEditorFromHomeWithTemplate(templateId: string): void {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.show()
    mainWindow.focus()
    homeWindow?.close()
    createEditorTab(undefined, templateId)
    return
  }
  // No main window yet — create it, then the first blank tab will pick up
  // the template via pendingTemplateHtml set in createEditorTab
  homeWindow?.close()
  createWindow(undefined)
  // createWindow calls createEditorTab(pendingInitialFilePath) which won't
  // have the templateId. Re-create with template once window is ready.
  // Simpler: just call createEditorTab with templateId after window exists.
  // createWindow is synchronous (window creation is sync), so tabs may already exist.
  // Close the blank tab that createWindow may have created and open a template one.
  // Actually the cleanest approach: create window normally, then createEditorTab with template.
  // The initial blank tab is created lazily by tab-bar-ready, so we just call:
  if (mainWindow && !mainWindow.isDestroyed()) {
    createEditorTab(undefined, templateId)
  }
}

function openAtlasWebWindow(): void {
  if (atlasWindow && !atlasWindow.isDestroyed()) {
    atlasWindow.show()
    atlasWindow.focus()
    homeWindow?.close()
    return
  }

  atlasWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    title: 'Atlas Web — Leandix',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      // Allow the embedded page to run normally
      webSecurity: true,
    },
  })

  atlasWindow.loadURL('https://atlas.leandix.com')
  atlasWindow.setMenuBarVisibility(false)

  const atlasMenu = Menu.buildFromTemplate([
    {
      label: 'Điều hướng',
      submenu: [
        {
          label: 'Trang chủ',
          accelerator: 'CmdOrCtrl+Shift+H',
          click: () => {
            if (homeWindow && !homeWindow.isDestroyed()) {
              homeWindow.show()
              homeWindow.focus()
            } else {
              createHomeWindow()
            }
          },
        },
        {
          label: 'Mở Editor',
          accelerator: 'CmdOrCtrl+Shift+E',
          click: () => {
            if (mainWindow && !mainWindow.isDestroyed()) {
              mainWindow.show()
              mainWindow.focus()
            } else {
              createWindow()
            }
          },
        },
        { type: 'separator' },
        {
          label: 'Thoát',
          role: 'quit',
        },
      ],
    },
  ])
  atlasWindow.setMenu(atlasMenu)
  atlasWindow.setMenuBarVisibility(true)

  atlasWindow.on('closed', () => {
    atlasWindow = null
  })

  homeWindow?.close()
}

// ---------------------------------------------------------------------------
// Template HTML helpers
// ---------------------------------------------------------------------------

/** Returns starter HTML for a named template in the current UI language. */
function getTemplateHtml(templateId: string, lang: 'vi' | 'en'): string {
  switch (templateId) {
    case 'report':
      return lang === 'vi'
        ? '<h1>Báo cáo</h1><h2>1. Giới thiệu</h2><p>Nội dung giới thiệu...</p><h2>2. Nội dung chính</h2><p>Nội dung chính...</p><h2>3. Kết luận</h2><p>Kết luận...</p>'
        : '<h1>Report</h1><h2>1. Introduction</h2><p>Introduction content...</p><h2>2. Main Content</h2><p>Main content...</p><h2>3. Conclusion</h2><p>Conclusion...</p>'
    case 'letter':
      return lang === 'vi'
        ? '<p style="text-align:right">Ngày ... tháng ... năm ...</p><p><strong>Kính gửi:</strong> ...</p><p>Nội dung thư...</p><p>Trân trọng,</p><p><strong>Người gửi</strong></p>'
        : '<p style="text-align:right">Date ... Month ... Year ...</p><p><strong>To:</strong> ...</p><p>Letter content...</p><p>Best regards,</p><p><strong>Sender</strong></p>'
    case 'notes':
      return lang === 'vi'
        ? '<h1>Ghi chú</h1><ul><li>Mục 1</li><li>Mục 2</li><li>Mục 3</li></ul>'
        : '<h1>Notes</h1><ul><li>Item 1</li><li>Item 2</li><li>Item 3</li></ul>'
    default: // 'blank' or unknown
      return '<p></p>'
  }
}

// ---------------------------------------------------------------------------
// Tab Creation / Activation
// ---------------------------------------------------------------------------

function createEditorTab(filePath?: string, templateId?: string): string {
  if (!mainWindow || mainWindow.isDestroyed()) return ''
  const id = crypto.randomUUID()
  const view = new WebContentsView({
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: path.join(__dirname, 'preload.js'),
      backgroundThrottling: false,  // keep rendering at full rate even when unfocused
      spellcheck: true,           // enable native OS spell checker
    },
  })
  view.webContents.loadFile(path.join(__dirname, '..', 'webview', 'index.html'))

  const tab: Tab = {
    id,
    type: 'editor',
    title: filePath ? path.basename(filePath) : (currentLanguage === 'vi' ? 'Tài liệu mới' : 'New Document'),
    documentState: { currentFilePath: filePath ?? null, fileContentHtml: '', fileExtras: {} },
    isDirty: false,
    view,
    pendingTemplateHtml: (!filePath && templateId) ? getTemplateHtml(templateId, currentLanguage) : undefined,
  }
  tabs.set(id, tab)
  mainWindow.contentView.addChildView(view)

  // Hide until activated
  view.setBounds({ x: 0, y: 0, width: 0, height: 0 })

  view.webContents.once('did-finish-load', () => {
    // Send current theme and language so editor starts with correct settings
    view.webContents.send('host-message', { type: 'theme', payload: { theme: currentTheme } })
    view.webContents.send('host-message', { type: 'language', payload: { language: currentLanguage } })

    // Inject drag-and-drop handler so the user can drop .docx/.pdf onto the editor
    void view.webContents.executeJavaScript(`
      (function() {
        const el = document.body;
        if (el.__leandixDropRegistered) return;
        el.__leandixDropRegistered = true;

        el.addEventListener('dragover', function(e) {
          const hasFile = e.dataTransfer && Array.from(e.dataTransfer.items).some(i => i.kind === 'file');
          if (hasFile) {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'copy';
          }
        });

        el.addEventListener('drop', function(e) {
          const files = e.dataTransfer && e.dataTransfer.files;
          if (!files || files.length === 0) return;
          const supported = Array.from(files).filter(f => {
            const name = f.name.toLowerCase();
            return name.endsWith('.docx') || name.endsWith('.pdf');
          });
          if (supported.length === 0) return;
          e.preventDefault();
          // Send each file path to main via IPC (path is available in Electron)
          supported.forEach(f => {
            if (f.path) {
              if (window.leandixApp) window.leandixApp.dropFile(f.path);
            }
          });
        });
      })();
    `).catch(() => { /* webview may be destroyed */ })

    if (filePath) {
      void loadAndSendFileToTab(id, filePath)
    }
  })

  setActiveTab(id)
  return id
}

function createPdfTab(filePath: string): string {
  if (!mainWindow || mainWindow.isDestroyed()) return ''
  const id = crypto.randomUUID()

  const pdfjsRoot = path.join(__dirname, '..', 'pdfjs')
  const viewerHtmlPath = path.join(pdfjsRoot, 'web', 'viewer.html')
  const buildBase = path.join(pdfjsRoot, 'build').replace(/\\/g, '/')
  const pdfFileUrl = `file:///${filePath.replace(/\\/g, '/')}`
  const workerSrc = `file:///${buildBase}/pdf.worker.mjs`
  const sandboxSrc = `file:///${buildBase}/pdf.sandbox.mjs`

  const view = new WebContentsView({
    webPreferences: {
      contextIsolation: false,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: false,
    },
  })
  view.webContents.loadFile(viewerHtmlPath)

  view.webContents.once('did-finish-load', () => {
    view.webContents.executeJavaScript(`
      if (typeof PDFViewerApplicationOptions !== 'undefined') {
        PDFViewerApplicationOptions.set('defaultUrl', '${pdfFileUrl.replace(/'/g, "\\'")}');
        PDFViewerApplicationOptions.set('workerSrc', '${workerSrc}');
        PDFViewerApplicationOptions.set('sandboxBundleSrc', '${sandboxSrc}');
        PDFViewerApplicationOptions.set('disablePreferences', true);
        if (typeof PDFViewerApplication !== 'undefined' && PDFViewerApplication.open) {
          PDFViewerApplication.open({ url: '${pdfFileUrl.replace(/'/g, "\\'")}' });
        }
      }
    `).catch(console.error)
  })

  const tab: Tab = {
    id,
    type: 'pdf',
    title: path.basename(filePath),
    documentState: { currentFilePath: filePath, fileContentHtml: '', fileExtras: {} },
    isDirty: false,
    view,
  }
  tabs.set(id, tab)
  mainWindow.contentView.addChildView(view)
  view.setBounds({ x: 0, y: 0, width: 0, height: 0 })

  setActiveTab(id)
  return id
}

function setActiveTab(id: string): void {
  if (!mainWindow || mainWindow.isDestroyed()) return
  const tab = tabs.get(id)
  if (!tab) return

  // Hide previous active tab
  if (activeTabId && activeTabId !== id) {
    const prev = tabs.get(activeTabId)
    if (prev) prev.view.setBounds({ x: 0, y: 0, width: 0, height: 0 })
  }

  activeTabId = id
  tab.view.setBounds(getContentBounds(mainWindow))
  pushTabState()
}

// ---------------------------------------------------------------------------
// Tab Close / Open Flow
// ---------------------------------------------------------------------------

async function confirmUnsavedChangesForTab(tab: Tab): Promise<'save' | 'discard' | 'cancel'> {
  if (!mainWindow || mainWindow.isDestroyed()) return 'discard'
  const { response } = await dialog.showMessageBox(mainWindow, {
    type: 'question',
    title: currentLanguage === 'vi' ? 'Tài liệu chưa lưu' : 'Unsaved document',
    message: currentLanguage === 'vi'
      ? `"${tab.title}" có thay đổi chưa lưu. Bạn có muốn lưu không?`
      : `"${tab.title}" has unsaved changes. Do you want to save?`,
    buttons: currentLanguage === 'vi' ? ['Lưu', 'Không lưu', 'Hủy'] : ['Save', 'Don\'t save', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  })
  if (response === 0) return 'save'
  if (response === 1) return 'discard'
  return 'cancel'
}

async function closeTab(tabId: string): Promise<void> {
  const tab = tabs.get(tabId)
  if (!tab) return

  if (tab.isDirty) {
    const action = await confirmUnsavedChangesForTab(tab)
    if (action === 'cancel') return
    if (action === 'save') await saveTabFile(tabId)
  }

  if (!mainWindow || mainWindow.isDestroyed()) return
  mainWindow.contentView.removeChildView(tab.view)
  tab.view.webContents.close()
  tabs.delete(tabId)

  if (activeTabId === tabId) {
    activeTabId = null
    const remaining = [...tabs.keys()]
    if (remaining.length > 0) {
      setActiveTab(remaining[remaining.length - 1])
    } else {
      // No tabs left — quit the app
      mainWindow?.destroy()
      return
    }
  }

  pushTabState()
}

async function openFileInTab(filePath: string, mode: 'new' | 'replace'): Promise<void> {
  const ext = path.extname(filePath).toLowerCase()

  if (mode === 'new') {
    if (ext === '.pdf') {
      createPdfTab(filePath)
    } else {
      createEditorTab(filePath)
    }
    return
  }

  // replace mode
  if (!activeTabId) {
    if (ext === '.pdf') createPdfTab(filePath)
    else createEditorTab(filePath)
    return
  }

  const activeTab = tabs.get(activeTabId)
  if (!activeTab) return

  if (activeTab.isDirty) {
    const action = await confirmUnsavedChangesForTab(activeTab)
    if (action === 'cancel') return
    if (action === 'save') await saveTabFile(activeTabId)
  }

  if (ext === '.pdf') {
    // Replace with a new pdf tab, remove old
    const oldId = activeTabId
    createPdfTab(filePath)
    const old = tabs.get(oldId)
    if (old && mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.contentView.removeChildView(old.view)
      old.view.webContents.close()
      tabs.delete(oldId)
    }
  } else {
    activeTab.title = path.basename(filePath)
    activeTab.documentState.currentFilePath = filePath
    activeTab.documentState.fileContentHtml = ''
    activeTab.documentState.fileExtras = {}
    activeTab.isDirty = false
    pushTabState()
    await loadAndSendFileToTab(activeTabId, filePath)
  }
}

// ---------------------------------------------------------------------------
// Per-tab File Loading / Saving / Message Handling
// ---------------------------------------------------------------------------

async function loadAndSendFileToTab(tabId: string, filePath: string): Promise<void> {
  const tab = tabs.get(tabId)
  if (!tab) return

  try {
    const buffer = await fs.readFile(filePath)
    const parsed = await parseDocx(buffer as unknown as Buffer)
    addRecentFile(filePath)

    tab.documentState.currentFilePath = filePath
    tab.documentState.fileContentHtml = parsed.bodyHtml
    tab.documentState.fileExtras = {
      headerHtml: parsed.headerHtml,
      footerHtml: parsed.footerHtml,
      comments: parsed.comments,
      pageSettings: parsed.pageSettings || DEFAULT_PAGE_SETTINGS,
      defaultFont: parsed.metadata?.defaultFont ?? undefined,
      defaultFontSize: parsed.metadata?.defaultFontSize ?? undefined,
    }
    // Cache AST — passed to serializeToDocx() on save to preserve round-trip fidelity
    tab.documentState.originalAst = parsed.ast
    tab.documentState.rawNumberingXml = parsed.rawNumberingXml

    tab.view.webContents.send('host-message', {
      type: 'load',
      payload: {
        html: parsed.bodyHtml,
        headerHtml: parsed.headerHtml,
        footerHtml: parsed.footerHtml,
        comments: parsed.comments,
        warnings: parsed.warnings,
        pageSettings: parsed.pageSettings || DEFAULT_PAGE_SETTINGS,
        language: currentLanguage,
        metadata: {
          defaultFont: parsed.metadata?.defaultFont ?? 'Times New Roman',
          defaultFontSize: parsed.metadata?.defaultFontSize ?? 13,
          defaultLineSpacing: parsed.metadata?.defaultLineSpacing,
          defaultSpacingAfter: parsed.metadata?.defaultSpacingAfter,
        },
        footnotes: parsed.footnotes ?? [],
      },
    })
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err)
    dialog.showErrorBox('Lỗi mở file', message)
  }
}

async function saveTabFile(tabId: string): Promise<void> {
  const tab = tabs.get(tabId)
  if (!tab || !mainWindow || mainWindow.isDestroyed()) return

  if (tab.documentState.currentFilePath === null) {
    await saveTabFileAs(tabId)
    return
  }

  try {
    const { fileExtras } = tab.documentState
    const buffer = await serializeToDocx(tab.documentState.fileContentHtml, {
      headerHtml: fileExtras.headerHtml,
      footerHtml: fileExtras.footerHtml,
      comments: fileExtras.comments,
      pageSettings: fileExtras.pageSettings,
      defaultFont: fileExtras.defaultFont,
      defaultFontSize: fileExtras.defaultFontSize,
      originalAst: tab.documentState.originalAst,
      rawNumberingXml: tab.documentState.rawNumberingXml,
    })
    await fs.writeFile(tab.documentState.currentFilePath, buffer)
    tab.isDirty = false
    tab.view.webContents.send('host-message', { type: 'saved', payload: {} })
    pushTabState()
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err)
    dialog.showErrorBox('Lỗi lưu file', message)
  }
}

async function saveTabFileAs(tabId: string): Promise<void> {
  const tab = tabs.get(tabId)
  if (!tab || !mainWindow || mainWindow.isDestroyed()) return

  const result = await dialog.showSaveDialog(mainWindow, {
    filters: [{ name: 'Word Documents', extensions: ['docx'] }],
  })
  if (result.canceled || !result.filePath) return

  try {
    const { fileExtras } = tab.documentState
    const buffer = await serializeToDocx(tab.documentState.fileContentHtml, {
      headerHtml: fileExtras.headerHtml,
      footerHtml: fileExtras.footerHtml,
      comments: fileExtras.comments,
      pageSettings: fileExtras.pageSettings,
      defaultFont: fileExtras.defaultFont,
      defaultFontSize: fileExtras.defaultFontSize,
      originalAst: tab.documentState.originalAst,
      rawNumberingXml: tab.documentState.rawNumberingXml,
    })
    await fs.writeFile(result.filePath, buffer)
    tab.documentState.currentFilePath = result.filePath
    tab.title = path.basename(result.filePath)
    tab.isDirty = false
    addRecentFile(result.filePath)
    tab.view.webContents.send('host-message', { type: 'saved', payload: {} })
    pushTabState()
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err)
    dialog.showErrorBox('Lỗi lưu file', message)
  }
}

async function handleVsCodeMessageForTab(tabId: string, message: unknown): Promise<void> {
  const tab = tabs.get(tabId)
  if (!tab) return
  if (!message || typeof message !== 'object') return
  const msg = message as { type?: unknown; payload?: unknown }
  if (typeof msg.type !== 'string') return

  switch (msg.type) {
    case 'ready': {
      if (tab.type === 'pdf') {
        const fileUrl = `file:///${tab.documentState.currentFilePath!.replace(/\\/g, '/')}`
        tab.view.webContents.send('host-message', { type: 'load-pdf-url', payload: { url: fileUrl } })
        tab.view.webContents.send('host-message', { type: 'language', payload: { language: currentLanguage } })
        break
      }
      if (tab.documentState.currentFilePath) {
        await loadAndSendFileToTab(tabId, tab.documentState.currentFilePath)
      } else {
        const templateHtml = tab.pendingTemplateHtml ?? ''
        tab.pendingTemplateHtml = undefined  // consume once

        // Serialize a blank document then parse it back so defaults (font, size,
        // spacing, pageSettings) come from the actual XML DocxSerializer generates.
        let defaultFont = 'Aptos'
        let defaultFontSize = 12
        let resolvedPageSettings = DEFAULT_PAGE_SETTINGS
        try {
          const blankBuffer = await serializeToDocx(templateHtml)
          const parsed = await parseDocx(blankBuffer as unknown as Buffer)
          if (parsed.metadata?.defaultFont) defaultFont = parsed.metadata.defaultFont
          if (parsed.metadata?.defaultFontSize) defaultFontSize = parsed.metadata.defaultFontSize
          resolvedPageSettings = parsed.pageSettings ?? DEFAULT_PAGE_SETTINGS
        } catch {
          // keep fallback values above
        }

        tab.view.webContents.send('host-message', {
          type: 'load',
          payload: {
            html: templateHtml,
            warnings: [],
            headerHtml: '',
            footerHtml: '',
            comments: [],
            pageSettings: resolvedPageSettings,
            language: currentLanguage,
            metadata: { defaultFont, defaultFontSize },
          },
        })
      }
      break
    }

    case 'update': {
      const payload = msg.payload as Record<string, unknown> | undefined
      if (!payload || typeof payload !== 'object') return
      if (
        typeof payload.html !== 'string' ||
        typeof payload.headerHtml !== 'string' ||
        typeof payload.footerHtml !== 'string' ||
        !Array.isArray(payload.comments) ||
        !payload.pageSettings ||
        typeof payload.pageSettings !== 'object'
      ) return
      tab.documentState.fileContentHtml = payload.html
      tab.documentState.fileExtras = {
        headerHtml: payload.headerHtml as string,
        footerHtml: payload.footerHtml as string,
        comments: payload.comments as Comment[],
        pageSettings: payload.pageSettings as typeof DEFAULT_PAGE_SETTINGS,
        defaultFont: typeof payload.defaultFont === 'string' ? payload.defaultFont : tab.documentState.fileExtras.defaultFont,
        defaultFontSize: typeof payload.defaultFontSize === 'number' ? payload.defaultFontSize : tab.documentState.fileExtras.defaultFontSize,
      }
      break
    }

    case 'save': {
      const payload = msg.payload as Record<string, unknown> | undefined
      if (!payload || typeof payload !== 'object') return
      if (
        typeof payload.html !== 'string' ||
        typeof payload.headerHtml !== 'string' ||
        typeof payload.footerHtml !== 'string' ||
        !Array.isArray(payload.comments) ||
        !payload.pageSettings ||
        typeof payload.pageSettings !== 'object'
      ) return
      tab.documentState.fileContentHtml = payload.html
      tab.documentState.fileExtras = {
        headerHtml: payload.headerHtml as string,
        footerHtml: payload.footerHtml as string,
        comments: payload.comments as Comment[],
        pageSettings: payload.pageSettings as typeof DEFAULT_PAGE_SETTINGS,
        defaultFont: typeof payload.defaultFont === 'string' ? payload.defaultFont : tab.documentState.fileExtras.defaultFont,
        defaultFontSize: typeof payload.defaultFontSize === 'number' ? payload.defaultFontSize : tab.documentState.fileExtras.defaultFontSize,
      }
      await saveTabFile(tabId)
      break
    }

    case 'requestLocalImage': {
      await handleRequestLocalImage(tab.view.webContents)
      break
    }

    case 'dirtyStateChanged': {
      const payload = msg.payload as Record<string, unknown> | undefined
      if (payload && typeof payload.isDirty === 'boolean') {
        tab.isDirty = payload.isDirty
        pushTabState()
      }
      break
    }

    case 'toggleLanguage': {
      currentLanguage = currentLanguage === 'vi' ? 'en' : 'vi'
      for (const t of tabs.values()) {
        if (!t.view.webContents.isDestroyed()) {
          t.view.webContents.send('host-message', { type: 'language', payload: { language: currentLanguage } })
        }
      }
      // Rebuild native app menu in new language
      createAppMenu()
      // Push language to tab bar
      if (tabBarView && !tabBarView.webContents.isDestroyed()) {
        tabBarView.webContents.send('tab-bar-language', currentLanguage)
      }
      break
    }

    case 'exportPdfHtml': {
      const payload = msg.payload as Record<string, unknown> | undefined
      const html = typeof payload?.html === 'string' ? payload.html : ''
      const resolve = pendingPdfExports.get(tabId)
      if (resolve) {
        pendingPdfExports.delete(tabId)
        resolve(html)
      }
      break
    }

    default:
      break
  }
}

function createWindow(initialFilePath?: string): void {
  pendingInitialFilePath = initialFilePath ?? null
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  })

  // Set up tab bar view
  tabBarView = new WebContentsView({
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: path.join(__dirname, 'tab-bar-preload.js'),
    },
  })
  mainWindow.contentView.addChildView(tabBarView)
  tabBarView.webContents.loadFile(path.join(__dirname, '..', 'tab-bar', 'index.html'))
  tabBarView.setBounds(getTabBarBounds(mainWindow))

  mainWindow.on('resize', updateLayout)

  // Main window drag & drop: intercept file:// navigation caused by OS drop
  mainWindow.webContents.on('will-navigate', (event, url) => {
    if (!url.startsWith('file://')) return
    let filePath: string
    try {
      filePath = decodeURIComponent(new URL(url).pathname)
      if (process.platform === 'win32' && filePath.startsWith('/')) filePath = filePath.slice(1)
    } catch { return }
    if (!isSupportedDroppedFile(filePath)) return
    event.preventDefault()
    void handleDroppedFile(filePath)
  })

  createAppMenu()

  mainWindow.on('close', (event) => {
    const dirtyTabs = [...tabs.values()].filter((t) => t.isDirty)
    if (dirtyTabs.length === 0) return
    event.preventDefault()
    void (async () => {
      for (const tab of dirtyTabs) {
        const action = await confirmUnsavedChangesForTab(tab)
        if (action === 'cancel') return
        if (action === 'save') await saveTabFile(tab.id)
      }
      mainWindow?.destroy()
    })()
  })

  mainWindow.on('closed', () => {
    mainWindow = null
    tabBarView = null
    // Quit the app when the editor window is closed
    if (!isQuitting) {
      isQuitting = true
      app.quit()
    }
  })
}

// ---------------------------------------------------------------------------
// IPC Message Dispatch (Task 3.2) — registered in app.whenReady()
// ---------------------------------------------------------------------------

function registerIpcHandlers(): void {
  ipcMain.handle('open-file-dialog-from-home', async () => {
    if (!homeWindow || homeWindow.isDestroyed()) return null
    const result = await dialog.showOpenDialog(homeWindow, {
      filters: [
        { name: currentLanguage === 'vi' ? 'Tài liệu được hỗ trợ' : 'Supported files', extensions: ['docx', 'pdf'] },
        { name: 'Word Documents', extensions: ['docx'] },
        { name: 'PDF Files', extensions: ['pdf'] },
      ],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    return result.filePaths[0]
  })

  ipcMain.handle('open-pdf-dialog-from-home', async () => {
    if (!homeWindow || homeWindow.isDestroyed()) return null
    const result = await dialog.showOpenDialog(homeWindow, {
      filters: [
        { name: 'PDF Files', extensions: ['pdf'] },
      ],
      properties: ['openFile'],
    })
    if (result.canceled || result.filePaths.length === 0) return null
    return result.filePaths[0]
  })

  ipcMain.handle('get-recent-files', () => {
    return loadRecentFiles()
  })

  ipcMain.handle('get-settings', () => {
    return { language: currentLanguage, theme: currentTheme }
  })

  ipcMain.on('apply-settings', (_event, settings: { language: 'vi' | 'en'; theme: 'dark' | 'light' }) => {
    currentLanguage = settings.language
    currentTheme = settings.theme
    saveSettings({ language: currentLanguage, theme: currentTheme })

    // Broadcast language to all surfaces
    if (homeWindow && !homeWindow.isDestroyed()) {
      homeWindow.webContents.send('home-language', currentLanguage)
      homeWindow.webContents.send('home-theme', currentTheme)
    }
    for (const t of tabs.values()) {
      if (!t.view.webContents.isDestroyed()) {
        t.view.webContents.send('host-message', { type: 'language', payload: { language: currentLanguage } })
        t.view.webContents.send('host-message', { type: 'theme', payload: { theme: currentTheme } })
      }
    }
    if (tabBarView && !tabBarView.webContents.isDestroyed()) {
      tabBarView.webContents.send('tab-bar-language', currentLanguage)
      tabBarView.webContents.send('tab-bar-theme', currentTheme)
    }
    createAppMenu()
  })

  ipcMain.on('home-action', (_event, message: { type: string; filePath?: string; templateId?: string }) => {
    if (message.type === 'open-editor') {
      openEditorFromHome(message.filePath)
    } else if (message.type === 'open-editor-new') {
      openEditorFromHome()
    } else if (message.type === 'open-editor-template') {
      openEditorFromHomeWithTemplate(message.templateId ?? 'blank')
    } else if (message.type === 'open-atlas-web') {
      openAtlasWebWindow()
    } else if (message.type === 'toggle-language') {
      currentLanguage = currentLanguage === 'vi' ? 'en' : 'vi'
      // Update home window
      if (homeWindow && !homeWindow.isDestroyed()) {
        homeWindow.webContents.send('home-language', currentLanguage)
      }
      // Update all editor tabs
      for (const t of tabs.values()) {
        if (!t.view.webContents.isDestroyed()) {
          t.view.webContents.send('host-message', { type: 'language', payload: { language: currentLanguage } })
        }
      }
      // Update tab bar
      if (tabBarView && !tabBarView.webContents.isDestroyed()) {
        tabBarView.webContents.send('tab-bar-language', currentLanguage)
      }
      // Rebuild native app menu
      createAppMenu()
    }
  })

  // Atlas Web: check network first (up to 5 min), with cancel support
  // Sends back 'atlas-web-check-result' with { status: 'online' | 'offline' | 'cancelled' }
  ipcMain.on('check-network-for-atlas', () => {
    if (!homeWindow || homeWindow.isDestroyed()) return

    const FIVE_MINUTES = 5 * 60 * 1000
    void checkNetworkCancellable(FIVE_MINUTES, (cancelFn) => {
      // Wire the cancel function so the renderer can trigger it
      ipcMain.once('cancel-atlas-web-check', () => cancelFn())
    }).then((result) => {
      if (!homeWindow || homeWindow.isDestroyed()) return
      if (result === 'cancelled') {
        homeWindow.webContents.send('atlas-web-check-result', { status: 'cancelled' })
      } else if (result === true) {
        homeWindow.webContents.send('atlas-web-check-result', { status: 'online' })
        openAtlasWebWindow()
      } else {
        homeWindow.webContents.send('atlas-web-check-result', { status: 'offline' })
      }
    })
  })

  ipcMain.on('vscode-message', (event, message) => {
    const tab = [...tabs.values()].find((t) => t.view.webContents === event.sender)
    if (tab) {
      void handleVsCodeMessageForTab(tab.id, message)
    }
  })

  ipcMain.on('tab-bar-ready', () => {
    if (tabs.size === 0) {
      const p = pendingInitialFilePath
      pendingInitialFilePath = null
      if (p) {
        const ext = path.extname(p).toLowerCase()
        if (ext === '.pdf') createPdfTab(p)
        else createEditorTab(p)
      } else {
        createEditorTab()
      }
    } else {
      pushTabState()
    }
    if (tabBarView && !tabBarView.webContents.isDestroyed()) {
      tabBarView.webContents.send('tab-bar-language', currentLanguage)
      tabBarView.webContents.send('tab-bar-theme', currentTheme)
    }
  })

  ipcMain.on('tab-create', () => {
    createEditorTab()
  })

  ipcMain.on('tab-close', (_event, { tabId }: { tabId: string }) => {
    void closeTab(tabId)
  })

  ipcMain.on('tab-switch', (_event, { tabId }: { tabId: string }) => {
    setActiveTab(tabId)
  })

  ipcMain.on('tab-open-mode', (_event, { mode }: { mode: 'new' | 'replace' }) => {
    if (pendingOpenFilePath) {
      const p = pendingOpenFilePath
      pendingOpenFilePath = null
      void openFileInTab(p, mode)
    }
  })

  // Drag & Drop file open (Feature 1.5)
  // Fired when the user drops a .docx or .pdf file onto an editor WebContentsView.
  ipcMain.on('drop-file', (_event, filePath: string) => {
    if (typeof filePath !== 'string' || !isSupportedDroppedFile(filePath)) return
    void handleDroppedFile(filePath)
  })

  // ── Auto Updater IPC ──────────────────────────────────────────────────────

  // Renderer asks: is there a pending update?
  ipcMain.handle('update-get-pending', () => pendingUpdateInfo)

  // Renderer triggers manual check
  ipcMain.on('update-check', () => {
    void autoUpdater.checkForUpdates()
  })

  // Renderer triggers download
  ipcMain.on('update-download', () => {
    void autoUpdater.downloadUpdate()
  })

  // Renderer triggers restart & install
  ipcMain.on('update-install', () => {
    autoUpdater.quitAndInstall()
  })
}


// ---------------------------------------------------------------------------
// File Open Dialog
// ---------------------------------------------------------------------------

async function showOpenFileDialog(): Promise<void> {
  if (!mainWindow || mainWindow.isDestroyed()) return

  const result = await dialog.showOpenDialog(mainWindow, {
    filters: [
      { name: 'Supported Files', extensions: ['docx', 'pdf'] },
      { name: 'Word Documents', extensions: ['docx'] },
      { name: 'PDF Files', extensions: ['pdf'] },
    ],
    properties: ['openFile'],
  })

  if (result.canceled || result.filePaths.length === 0) return

  const filePath = result.filePaths[0]

  if (tabs.size === 0) {
    const ext = path.extname(filePath).toLowerCase()
    if (ext === '.pdf') createPdfTab(filePath)
    else createEditorTab(filePath)
    return
  }

  const { response } = await dialog.showMessageBox(mainWindow, {
    type: 'question',
    title: currentLanguage === 'vi' ? 'Mở file' : 'Open file',
    message: currentLanguage === 'vi'
      ? `Mở "${path.basename(filePath)}" trong:`
      : `Open "${path.basename(filePath)}" in:`,
    buttons: currentLanguage === 'vi'
      ? ['Tab mới', 'Thay thế tab hiện tại', 'Hủy']
      : ['New tab', 'Replace current tab', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  })

  if (response === 2) return
  void openFileInTab(filePath, response === 0 ? 'new' : 'replace')
}

// ---------------------------------------------------------------------------
// Local Image Insertion
// ---------------------------------------------------------------------------

async function handleRequestLocalImage(webContents: Electron.WebContents): Promise<void> {
  if (!mainWindow || mainWindow.isDestroyed()) return

  const result = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg'] }],
    properties: ['openFile'],
  })

  if (result.canceled || result.filePaths.length === 0) return

  const filePath = result.filePaths[0]

  try {
    const buffer = await fs.readFile(filePath)
    const ext = path.extname(filePath).toLowerCase().replace('.', '')
    const mimeMap: Record<string, string> = {
      png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
      gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml',
    }
    const mime = mimeMap[ext] || 'application/octet-stream'
    webContents.send('host-message', {
      type: 'localImageResult',
      payload: { success: true, base64Data: `data:${mime};base64,${buffer.toString('base64')}`, fileName: path.basename(filePath) },
    })
  } catch {
    webContents.send('host-message', { type: 'localImageResult', payload: { success: false } })
  }
}

// ---------------------------------------------------------------------------
// Native PDF Export (Task 3.6) — headless Chrome/Edge
// ---------------------------------------------------------------------------

function findBrowserExecutable(): string | undefined {
  if (process.platform === 'win32') {
    const programFiles = process.env['ProgramFiles'] || 'C:\\Program Files'
    const programFilesX86 = process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)'
    const localAppData = process.env['LocalAppData'] || path.join(os.homedir(), 'AppData\\Local')
    const candidates = [
      path.join(programFilesX86, 'Microsoft\\Edge\\Application\\msedge.exe'),
      path.join(programFiles, 'Microsoft\\Edge\\Application\\msedge.exe'),
      path.join(programFiles, 'Google\\Chrome\\Application\\chrome.exe'),
      path.join(programFilesX86, 'Google\\Chrome\\Application\\chrome.exe'),
      path.join(localAppData, 'Google\\Chrome\\Application\\chrome.exe'),
    ]
    for (const p of candidates) {
      if (fsSync.existsSync(p)) return p
    }
  } else if (process.platform === 'darwin') {
    const candidates = [
      '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
      '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    ]
    for (const p of candidates) {
      if (fsSync.existsSync(p)) return p
    }
  } else {
    return 'chromium'
  }
  return undefined
}

function getGoogleFontsLinks(html: string): string {
  const systemFonts = new Set([
    'Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana',
    'Tahoma', 'Trebuchet MS', 'Comic Sans MS', 'Impact', 'Lucida Console',
    'Palatino Linotype', 'Segoe UI', 'Calibri', 'Cambria',
  ])
  const fontSet = new Set<string>()
  const regex = /font-family:\s*([^;"<]+)/gi
  let match: RegExpExecArray | null
  while ((match = regex.exec(html)) !== null) {
    const fontName = match[1].trim().replace(/^["']|["']$/g, '').split(',')[0].trim()
    if (fontName && !systemFonts.has(fontName)) fontSet.add(fontName)
  }
  let links = ''
  for (const fontName of fontSet) {
    const encoded = fontName.replace(/\s+/g, '+')
    links += `<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=${encoded}:ital,wght@0,100;0,300;0,400;0,500;0,700;0,900;1,100;1,300;1,400;1,500;1,700;1,900&display=swap">\n`
  }
  return links
}

async function exportToPdf(): Promise<void> {
  if (!mainWindow || mainWindow.isDestroyed()) return

  const activeTab = activeTabId ? tabs.get(activeTabId) : null
  if (!activeTab) {
    dialog.showErrorBox('Lỗi xuất PDF', 'Vui lòng mở hoặc tạo tài liệu trước khi xuất PDF.')
    return
  }

  const docState = activeTab.documentState
  if (!docState?.fileContentHtml) {
    dialog.showErrorBox('Lỗi xuất PDF', 'Vui lòng mở hoặc tạo tài liệu trước khi xuất PDF.')
    return
  }

  const defaultName = docState.currentFilePath
    ? path.parse(docState.currentFilePath).name + '.pdf'
    : 'document.pdf'
  const defaultDir = docState.currentFilePath
    ? path.dirname(docState.currentFilePath)
    : app.getPath('documents')

  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: path.join(defaultDir, defaultName),
    filters: [{ name: 'PDF Files', extensions: ['pdf'] }],
  })
  if (result.canceled || !result.filePath) return

  const browserPath = findBrowserExecutable()

  // Read the compiled main.css from the webview output
  const cssPath = path.join(__dirname, '..', 'webview', 'main.css')
  let cssContent = ''
  try {
    cssContent = await fs.readFile(cssPath, 'utf-8')
  } catch {
    // proceed without CSS
  }

  const pageSettings = docState.fileExtras.pageSettings ?? DEFAULT_PAGE_SETTINGS
  // Treat Tiptap's empty-editor placeholder ('<p></p>') the same as '' so we
  // don't reserve header/footer space (and resize content area) for empty sections.
  const rawHeaderHtml = docState.fileExtras.headerHtml ?? ''
  const rawFooterHtml = docState.fileExtras.footerHtml ?? ''
  const headerHtml = rawHeaderHtml && rawHeaderHtml.replace(/<[^>]+>/g, '').trim() ? rawHeaderHtml : ''
  const footerHtml = rawFooterHtml && rawFooterHtml.replace(/<[^>]+>/g, '').trim() ? rawFooterHtml : ''

  // Run pagination via webview — same as VS Code extension flow.
  // We send an 'exportPdf' message and wait for the webview's onExportPdf
  // handler (App.tsx) to respond with 'exportPdfHtml' containing the fully
  // paginated, header/footer-aware HTML.
  let pagesHtml = ''
  try {
    pagesHtml = await new Promise<string>((resolve, reject) => {
      const timeout = setTimeout(() => {
        pendingPdfExports.delete(activeTab.id)
        reject(new Error('Timeout waiting for exportPdfHtml from webview'))
      }, 30000)

      pendingPdfExports.set(activeTab.id, (html: string) => {
        clearTimeout(timeout)
        resolve(html)
      })

      activeTab.view.webContents.send('host-message', { type: 'exportPdf', payload: {} })
    })
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    dialog.showErrorBox('Lỗi xuất PDF', 'Không lấy được nội dung từ editor: ' + msg)
    return
  }

  if (!pagesHtml) {
    dialog.showErrorBox('Lỗi xuất PDF', 'Không tìm thấy nội dung editor. Hãy mở tài liệu trước khi xuất PDF.')
    return
  }

  const googleFontsLinks = getGoogleFontsLinks(docState.fileContentHtml)

  const fullHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Export PDF</title>
  ${googleFontsLinks}
  <style>${cssContent}</style>
  <style>
    @page {
      size: ${pageSettings.pageWidth}mm ${pageSettings.pageHeight}mm;
      margin: 0;
    }
    /* Override editor.css which sets overflow:hidden and height:100% on body —
       those styles clip content to one page when printing. */
    html, body {
      height: auto !important;
      overflow: visible !important;
      margin: 0 !important;
      padding: 0 !important;
      background: white !important;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .page-preview-scale-wrapper {
      display: none !important;
    }
    .page-preview-sheet {
      box-shadow: none !important;
      transform: none !important;
      position: relative !important;
      page-break-after: always !important;
      break-after: page !important;
      overflow: hidden !important;
    }
  </style>
</head>
<body>
${pagesHtml}
</body>
</html>`

  const tempDir = path.join(os.tmpdir(), `leandix-pdf-${Date.now()}`)
  const tempHtmlPath = path.join(tempDir, 'document.html')

  if (!browserPath) {
    // Fallback: use Electron printToPDF
    await dialog.showMessageBox(mainWindow, {
      type: 'warning',
      title: 'Không tìm thấy trình duyệt',
      message: 'Không tìm thấy Microsoft Edge hoặc Google Chrome. Sẽ xuất PDF bằng phương pháp dự phòng — kết quả có thể không chính xác.',
      buttons: ['Tiếp tục'],
    })
    try {
      fsSync.mkdirSync(tempDir, { recursive: true })
      fsSync.writeFileSync(tempHtmlPath, fullHtml, 'utf-8')
      const pdfBuffer = await mainWindow.webContents.printToPDF({
        pageSize: { width: pageSettings.pageWidth * 1000, height: pageSettings.pageHeight * 1000 },
        printBackground: true,
        margins: { marginType: 'none' },
      })
      await fs.writeFile(result.filePath, pdfBuffer)
      await dialog.showMessageBox(mainWindow, {
        type: 'info',
        title: 'Xuất PDF',
        message: 'Xuất PDF thành công!',
      })
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err)
      dialog.showErrorBox('Lỗi xuất PDF', message)
    } finally {
      try { fsSync.rmSync(tempDir, { recursive: true, force: true }) } catch { /* ignore */ }
    }
    return
  }

  try {
    fsSync.mkdirSync(tempDir, { recursive: true })
    fsSync.writeFileSync(tempHtmlPath, fullHtml, 'utf-8')

    const fileUrl = `file:///${tempHtmlPath.replace(/\\/g, '/')}`

    await new Promise<void>((resolve, reject) => {
      const args = [
        '--headless=old',
        '--disable-gpu',
        '--no-pdf-header-footer',
        `--print-to-pdf=${result.filePath}`,
        fileUrl,
      ]
      execFile(browserPath, args, (err) => {
        if (err) reject(err)
        else resolve()
      })
    })

    await dialog.showMessageBox(mainWindow!, {
      type: 'info',
      title: 'Xuất PDF',
      message: 'Xuất PDF thành công!',
    })
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err)
    dialog.showErrorBox('Lỗi xuất PDF', message)
  } finally {
    try { fsSync.rmSync(tempDir, { recursive: true, force: true }) } catch { /* ignore */ }
  }
}

// ---------------------------------------------------------------------------
// Vietnamese Application Menu (Task 3.7)
// ---------------------------------------------------------------------------

const MENU_STRINGS: Record<'vi' | 'en', Record<string, string>> = {
  vi: {
    file: 'Tệp tin',
    home: 'Trang chủ',
    new: 'Tạo mới',
    open: 'Mở file...',
    save: 'Lưu',
    saveAs: 'Lưu dưới dạng...',
    exportPdf: 'Xuất PDF...',
    quit: 'Thoát',
    edit: 'Chỉnh sửa',
    undo: 'Hoàn tác',
    redo: 'Làm lại',
    cut: 'Cắt',
    copy: 'Sao chép',
    paste: 'Dán',
    selectAll: 'Chọn tất cả',
    view: 'Xem',
    reload: 'Tải lại giao diện',
    devTools: 'Bật DevTools',
    resetZoom: 'Đặt lại Zoom',
    zoomIn: 'Phóng to',
    zoomOut: 'Thu nhỏ',
    fullscreen: 'Toàn màn hình',
  },
  en: {
    file: 'File',
    home: 'Home',
    new: 'New',
    open: 'Open...',
    save: 'Save',
    saveAs: 'Save As...',
    exportPdf: 'Export PDF...',
    quit: 'Quit',
    edit: 'Edit',
    undo: 'Undo',
    redo: 'Redo',
    cut: 'Cut',
    copy: 'Copy',
    paste: 'Paste',
    selectAll: 'Select All',
    view: 'View',
    reload: 'Reload',
    devTools: 'Toggle DevTools',
    resetZoom: 'Reset Zoom',
    zoomIn: 'Zoom In',
    zoomOut: 'Zoom Out',
    fullscreen: 'Toggle Fullscreen',
  },
}

function createAppMenu(): void {
  const m = MENU_STRINGS[currentLanguage]
  const template: Electron.MenuItemConstructorOptions[] = [
    {
      label: m.file,
      submenu: [
        {
          label: m.home,
          accelerator: 'CmdOrCtrl+Shift+H',
          click: () => {
            if (homeWindow && !homeWindow.isDestroyed()) {
              homeWindow.show()
              homeWindow.focus()
            } else {
              createHomeWindow()
            }
          },
        },
        { type: 'separator' },
        {
          label: m.new,
          accelerator: 'CmdOrCtrl+N',
          click: () => { createEditorTab() },
        },
        {
          label: m.open,
          accelerator: 'CmdOrCtrl+O',
          click: () => { void showOpenFileDialog() },
        },
        {
          label: m.save,
          accelerator: 'CmdOrCtrl+S',
          click: () => { if (activeTabId) void saveTabFile(activeTabId) },
        },
        {
          label: m.saveAs,
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => { if (activeTabId) void saveTabFileAs(activeTabId) },
        },
        { type: 'separator' },
        {
          label: m.exportPdf,
          accelerator: 'CmdOrCtrl+P',
          click: () => { void exportToPdf() },
        },
        { type: 'separator' },
        { label: m.quit, role: 'quit' },
      ],
    },
    {
      label: m.edit,
      submenu: [
        { label: m.undo, role: 'undo' },
        { label: m.redo, role: 'redo' },
        { type: 'separator' },
        { label: m.cut, role: 'cut' },
        { label: m.copy, role: 'copy' },
        { label: m.paste, role: 'paste' },
        { label: m.selectAll, role: 'selectAll' },
      ],
    },
    {
      label: m.view,
      submenu: [
        { label: m.reload, role: 'reload' },
        { label: m.devTools, role: 'toggleDevTools' },
        { type: 'separator' },
        { label: m.resetZoom, role: 'resetZoom' },
        { label: m.zoomIn, role: 'zoomIn' },
        { label: m.zoomOut, role: 'zoomOut' },
        { type: 'separator' },
        { label: m.fullscreen, role: 'togglefullscreen' },
      ],
    },
  ]

  const menu = Menu.buildFromTemplate(template)
  Menu.setApplicationMenu(menu)
}

// ---------------------------------------------------------------------------
// Drag & Drop File Open (1.5)
// ---------------------------------------------------------------------------

/**
 * Validate that a dropped file path has an extension the app can open.
 */
function isSupportedDroppedFile(filePath: string): boolean {
  const ext = path.extname(filePath).toLowerCase()
  return ext === '.docx' || ext === '.pdf'
}

/**
 * Register drag-and-drop handlers on a BrowserWindow's webContents.
 * When the user drops a .docx or .pdf file onto the window, we open it
 * in a new tab (or prompt replace/new if the editor window is ready).
 */
function registerDropHandlerOnWindow(win: BrowserWindow): void {
  // Electron fires will-navigate when the webContents would navigate to a
  // file:// URL as a result of the user dropping a file onto the window.
  win.webContents.on('will-navigate', (event, url) => {
    // Only intercept file:// drops (Electron turns a dropped file into a
    // file:// navigation request on the main webContents when there is no
    // explicit drop handler).
    if (!url.startsWith('file://')) return

    let filePath: string
    try {
      // Convert file URL → OS path
      filePath = decodeURIComponent(new URL(url).pathname)
      // Windows: strip leading slash from /C:/...
      if (process.platform === 'win32' && filePath.startsWith('/')) {
        filePath = filePath.slice(1)
      }
    } catch {
      return
    }

    if (!isSupportedDroppedFile(filePath)) return

    event.preventDefault()
    void handleDroppedFile(filePath)
  })

  // Also cover the case where Electron's default behavior is suppressed but
  // the drop event still bubbles to the native layer.  We use the
  // `webContents.on('did-finish-load')` opportunity to inject a JS handler
  // that forwards dragged file paths over IPC.
}

/**
 * Open a dropped file in the editor.  If the editor window is already open
 * we ask the user "New tab or replace current tab?" — same flow as
 * showOpenFileDialog.  If the editor window is not yet open we launch it.
 */
async function handleDroppedFile(filePath: string): Promise<void> {
  // Home window drop → open editor
  if (!mainWindow || mainWindow.isDestroyed()) {
    openEditorFromHome(filePath)
    return
  }

  mainWindow.show()
  mainWindow.focus()

  if (tabs.size === 0) {
    const ext = path.extname(filePath).toLowerCase()
    if (ext === '.pdf') createPdfTab(filePath)
    else createEditorTab(filePath)
    return
  }

  const { response } = await dialog.showMessageBox(mainWindow, {
    type: 'question',
    title: currentLanguage === 'vi' ? 'Mở file' : 'Open file',
    message: currentLanguage === 'vi'
      ? `Mở "${path.basename(filePath)}" trong:`
      : `Open "${path.basename(filePath)}" in:`,
    buttons: currentLanguage === 'vi'
      ? ['Tab mới', 'Thay thế tab hiện tại', 'Hủy']
      : ['New tab', 'Replace current tab', 'Cancel'],
    defaultId: 0,
    cancelId: 2,
  })

  if (response === 2) return
  void openFileInTab(filePath, response === 0 ? 'new' : 'replace')
}

// ---------------------------------------------------------------------------
// Auto Updater
// ---------------------------------------------------------------------------

/** Cached update info if a new version is available */
let pendingUpdateInfo: { version: string } | null = null

function setupAutoUpdater(): void {
  autoUpdater.autoDownload = false
  autoUpdater.autoInstallOnAppQuit = false

  autoUpdater.on('checking-for-update', () => {
    console.log('[updater] checking-for-update')
  })

  autoUpdater.on('update-available', (info) => {
    pendingUpdateInfo = { version: info.version }
    // Notify home window if it's open
    if (homeWindow && !homeWindow.isDestroyed()) {
      homeWindow.webContents.send('update-available', { version: info.version })
    }
  })

  autoUpdater.on('update-not-available', () => {
    pendingUpdateInfo = null
    if (homeWindow && !homeWindow.isDestroyed()) {
      homeWindow.webContents.send('update-not-available')
    }
  })

  autoUpdater.on('download-progress', (progress) => {
    if (homeWindow && !homeWindow.isDestroyed()) {
      homeWindow.webContents.send('update-download-progress', {
        percent: Math.round(progress.percent),
        transferred: progress.transferred,
        total: progress.total,
      })
    }
  })

  autoUpdater.on('update-downloaded', () => {
    if (homeWindow && !homeWindow.isDestroyed()) {
      homeWindow.webContents.send('update-downloaded')
    }
  })

  autoUpdater.on('error', (err) => {
    console.error('[updater] error:', err.message, err.stack)
    const msg = err.message + (err.stack ? '\n' + err.stack : '')
    if (homeWindow && !homeWindow.isDestroyed()) {
      homeWindow.webContents.send('update-error', { message: msg })
    }
  })
}

// ---------------------------------------------------------------------------
// App Lifecycle
// ---------------------------------------------------------------------------

let isQuitting = false

// ── GPU / rendering flags removed ─────────────────────────────────────────

/**
 * Extract a supported file path from argv.
 * Electron passes the file as the last argument when launched via file association.
 * In packaged builds argv[0] is the exe, argv[1] may be '--' or the file path.
 * In dev builds argv[0] is electron, argv[1] is the entry script, argv[2]+ are extras.
 */
function getFileFromArgv(argv: string[]): string | null {
  // Walk args in reverse, skip flags (starting with '-'), pick first real path
  for (let i = argv.length - 1; i >= 1; i--) {
    const arg = argv[i]
    if (arg.startsWith('-')) continue
    const lower = arg.toLowerCase()
    if (lower.endsWith('.docx') || lower.endsWith('.pdf')) {
      return arg
    }
  }
  return null
}

// Single-instance lock: if another instance tries to open, forward its argv here
const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  // Another instance is already running — it will receive 'second-instance' event
  app.quit()
} else {
  app.on('second-instance', (_event, argv) => {
    // A second instance was launched (e.g. user double-clicked another file)
    const filePath = getFileFromArgv(argv)
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.show()
      mainWindow.focus()
      if (filePath) {
        const ext = path.extname(filePath).toLowerCase()
        if (ext === '.pdf') createPdfTab(filePath)
        else createEditorTab(filePath)
      }
    } else if (filePath) {
      openEditorFromHome(filePath)
    } else {
      // No file and no window open — just ignore, app is idle
    }
  })

  app.whenReady().then(() => {
    setupAutoUpdater()
    registerIpcHandlers()

    // Check if launched with a file argument (double-click / file association)
    const initialFile = getFileFromArgv(process.argv)
    if (initialFile) {
      // Open directly to editor with the file, skip home screen
      createWindow(initialFile)
    } else {
      createHomeWindow()
    }

    app.on('activate', () => {
      // macOS: re-create window when dock icon is clicked and no windows exist
      if (BrowserWindow.getAllWindows().length === 0) {
        createHomeWindow()
      }
    })
  })
}

app.on('before-quit', (event) => {
  if (isQuitting) return
  const dirtyTabs = [...tabs.values()].filter((t) => t.isDirty)
  if (dirtyTabs.length === 0) return
  if (!mainWindow || mainWindow.isDestroyed()) return
  event.preventDefault()
  void (async () => {
    for (const tab of dirtyTabs) {
      const action = await confirmUnsavedChangesForTab(tab)
      if (action === 'cancel') return
      if (action === 'save') await saveTabFile(tab.id)
    }
    isQuitting = true
    app.quit()
  })()
})

app.on('window-all-closed', () => {
  // On macOS, apps typically stay active until Cmd+Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

// ---------------------------------------------------------------------------
// Exports for testing
// ---------------------------------------------------------------------------

export {
  createHomeWindow,
  createWindow,
  openEditorFromHome,
  openAtlasWebWindow,
  handleVsCodeMessageForTab,
  loadAndSendFileToTab,
  showOpenFileDialog,
  saveTabFile,
  saveTabFileAs,
  handleRequestLocalImage,
  exportToPdf,
  createPdfTab,
  createEditorTab,
  createAppMenu,
  addRecentFile,
  loadRecentFiles,
  tabs,
  activeTabId,
  homeWindow,
  mainWindow,
  atlasWindow,
  currentLanguage,
}
