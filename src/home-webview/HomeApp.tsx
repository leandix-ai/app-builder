import React, { useEffect, useState } from 'react'
import { FileText, Globe, WifiOff, Loader, Plus, Clock, Settings, Moon, Sun, BookOpen, BarChart2, Mail, NotebookPen, HelpCircle, Download, RefreshCw, X } from 'lucide-react'
import logoUrl from '../logo/logo_final.png'

type Lang = 'vi' | 'en'
type Theme = 'dark' | 'light'
/** State of the Atlas Web connection check overlay */
type AtlasCheckState = 'idle' | 'checking' | 'offline'

interface RecentFile {
  filePath: string
  title: string
  lastOpenedAt: string
}

declare global {
  interface Window {
    homeApi: {
      onNetworkStatus: (cb: (online: boolean) => void) => void
      openEditor: () => void
      openEditorNew: () => void
      openEditorWithFile: (filePath: string) => void
      openEditorFromTemplate: (templateId: string) => void
      openAtlasWeb: () => void
      checkNetworkForAtlas: () => void
      cancelAtlasWebCheck: () => void
      onAtlasWebCheckResult: (cb: (result: { status: 'online' | 'offline' | 'cancelled' }) => void) => void
      onLanguage: (cb: (lang: string) => void) => void
      onTheme: (cb: (theme: string) => void) => void
      getSettings: () => Promise<{ language: Lang; theme: Theme }>
      applySettings: (settings: { language: Lang; theme: Theme }) => void
      openFileDialog: () => Promise<string | null>
      openPdfDialog: () => Promise<string | null>
      getRecentFiles: () => Promise<RecentFile[]>
      // Auto updater
      getPendingUpdate: () => Promise<{ version: string } | null>
      checkForUpdate: () => void
      downloadUpdate: () => void
      installUpdate: () => void
      onUpdateAvailable: (cb: (info: { version: string }) => void) => void
      onUpdateNotAvailable: (cb: () => void) => void
      onUpdateDownloadProgress: (cb: (progress: { percent: number; transferred: number; total: number }) => void) => void
      onUpdateDownloaded: (cb: () => void) => void
      onUpdateError: (cb: (err: { message: string }) => void) => void
    }
  }
}

const STRINGS = {
  vi: {
    checking: 'Đang kiểm tra kết nối...',
    offline: 'Không có kết nối mạng',
    subtitle: 'Chọn phương thức làm việc',
    editorTitle: 'Editor',
    editorDesc: 'Soạn thảo tài liệu DOCX trực tiếp',
    atlasTitle: 'Atlas Web',
    atlasDesc: 'Sử dụng AI tại atlas.leandix.com',
    offlineNote: 'Chế độ offline chỉ hỗ trợ chỉnh sửa văn bản.',
    newDoc: 'Tài liệu trống',
    openDoc: 'Mở DOCX...',
    openPdf: 'Mở PDF...',
    recentDocs: 'Tài liệu gần đây',
    noRecent: 'Chưa có tài liệu nào được mở gần đây.',
    back: '← Trang chủ',
    settings: 'Cài đặt',
    settingsTheme: 'Giao diện',
    settingsDark: 'Tối',
    settingsLight: 'Sáng',
    settingsLang: 'Ngôn ngữ',
    settingsApply: 'Áp dụng',
    settingsCancel: 'Hủy',
    templates: 'Mẫu tài liệu',
    tplReport: 'Báo cáo',
    tplLetter: 'Thư / Đơn',
    tplNotes: 'Ghi chú',
    atlasCheckingTitle: 'Đang kết nối Atlas Web...',
    atlasCheckingDesc: 'Đang kiểm tra kết nối đến atlas.leandix.com',
    atlasCancel: 'Hủy',
    atlasOfflineTitle: 'Không có kết nối mạng',
    atlasOfflineDesc: 'Không thể kết nối đến atlas.leandix.com. Vui lòng kiểm tra lại kết nối.',
    atlasOfflineClose: 'Đóng',
  },
  en: {
    checking: 'Checking connection...',
    offline: 'No network connection',
    subtitle: 'Choose your workspace',
    editorTitle: 'Editor',
    editorDesc: 'Edit DOCX documents directly',
    atlasTitle: 'Atlas Web',
    atlasDesc: 'Use AI at atlas.leandix.com',
    offlineNote: 'Offline mode supports text editing only.',
    newDoc: 'Blank document',
    openDoc: 'Open DOCX...',
    openPdf: 'Open PDF...',
    recentDocs: 'Recent documents',
    noRecent: 'No documents opened recently.',
    back: '← Home',
    settings: 'Settings',
    settingsTheme: 'Theme',
    settingsDark: 'Dark',
    settingsLight: 'Light',
    settingsLang: 'Language',
    settingsApply: 'Apply',
    settingsCancel: 'Cancel',
    templates: 'Templates',
    tplReport: 'Report',
    tplLetter: 'Letter',
    tplNotes: 'Notes',
    atlasCheckingTitle: 'Connecting to Atlas Web...',
    atlasCheckingDesc: 'Checking connection to atlas.leandix.com',
    atlasCancel: 'Cancel',
    atlasOfflineTitle: 'No network connection',
    atlasOfflineDesc: 'Unable to reach atlas.leandix.com. Please check your connection.',
    atlasOfflineClose: 'Close',
  },
}

function formatRelativeDate(isoString: string, lang: Lang): string {
  const date = new Date(isoString)
  const now = new Date()
  const diffMs = now.getTime() - date.getTime()
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24))

  if (lang === 'vi') {
    if (diffDays === 0) return 'Hôm nay'
    if (diffDays === 1) return 'Hôm qua'
    if (diffDays < 7) return `${diffDays} ngày trước`
    return date.toLocaleDateString('vi-VN', { day: 'numeric', month: 'long', year: 'numeric' })
  } else {
    if (diffDays === 0) return 'Today'
    if (diffDays === 1) return 'Yesterday'
    if (diffDays < 7) return `${diffDays} days ago`
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
  }
}

function DocIcon() {
  return (
    <svg width="28" height="36" viewBox="0 0 28 36" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect width="28" height="36" rx="3" fill="#4285f4" opacity="0.12" />
      <rect x="3" y="3" width="22" height="30" rx="2" fill="#e8f0fe" />
      <path d="M18 3v7h7" fill="none" stroke="#4285f4" strokeWidth="1" opacity="0.5" />
      <path d="M18 3l7 7H18V3z" fill="#c5d8fd" />
      <rect x="6" y="14" width="16" height="1.5" rx="0.75" fill="#4285f4" opacity="0.4" />
      <rect x="6" y="18" width="16" height="1.5" rx="0.75" fill="#4285f4" opacity="0.4" />
      <rect x="6" y="22" width="11" height="1.5" rx="0.75" fill="#4285f4" opacity="0.4" />
    </svg>
  )
}

function PdfIcon() {
  return (
    <svg width="28" height="36" viewBox="0 0 28 36" fill="none" xmlns="http://www.w3.org/2000/svg">
      <rect width="28" height="36" rx="3" fill="#ef4444" opacity="0.12" />
      <rect x="3" y="3" width="22" height="30" rx="2" fill="#fee2e2" />
      <path d="M18 3v7h7" fill="none" stroke="#ef4444" strokeWidth="1" opacity="0.5" />
      <path d="M18 3l7 7H18V3z" fill="#fecaca" />
      <text x="14" y="25" textAnchor="middle" fontSize="7" fontWeight="700" fill="#ef4444" opacity="0.8" fontFamily="sans-serif">PDF</text>
    </svg>
  )
}

function AboutModal({ lang, onClose }: { lang: Lang; onClose: () => void }) {
  function handleOverlayClick(e: React.MouseEvent<HTMLDivElement>) {
    if (e.target === e.currentTarget) onClose()
  }

  return (
    <div className="settings-overlay" onClick={handleOverlayClick}>
      <div className="settings-modal about-modal">
        {/* Logo + app name */}
        <div className="about-header">
          <img src={logoUrl} alt="Leandix Atlas" className="about-logo" />
          <div>
            <div className="about-app-name">Leandix Atlas</div>
            <div className="about-version">v2.0.0</div>
          </div>
        </div>

        <div className="about-divider" />

        {/* Contact */}
        <div className="about-contact-row">
          <Mail size={15} className="about-contact-icon" />
          <div className="about-contact-info">
            <span className="about-contact-label">
              {lang === 'vi' ? 'Liên hệ hỗ trợ' : 'Support contact'}
            </span>
            <a href="mailto:atlas@leandix.vn" className="about-contact-email">
              atlas@leandix.vn
            </a>
          </div>
        </div>

        <div className="about-divider" />

        {/* Credits */}
        <div className="about-credits">
          <div className="about-company">Công Ty TNHH LEANDIX</div>
          <div className="about-team-label">
            {lang === 'vi' ? 'Đội ngũ phát triển' : 'Development team'}: Team Atlas
          </div>
          <div className="about-team-members">
            <span>Tạ Ngọc Nam</span>
            <span>Thịnh Hải Triều</span>
          </div>
        </div>

        <div className="settings-modal-actions">
          <button className="settings-btn-apply" onClick={onClose}>
            {lang === 'vi' ? 'Đóng' : 'Close'}
          </button>
        </div>
      </div>
    </div>
  )
}

interface SettingsModalProps {
  lang: Lang
  theme: Theme
  onClose: () => void
  onApply: (lang: Lang, theme: Theme) => void
  pendingUpdate: string | null
}

type UpdateState = 'idle' | 'checking' | 'available' | 'downloading' | 'downloaded' | 'error' | 'up-to-date'

function SettingsModal({ lang, theme, onClose, onApply, pendingUpdate }: SettingsModalProps) {
  const [draftLang, setDraftLang] = useState<Lang>(lang)
  const [draftTheme, setDraftTheme] = useState<Theme>(theme)
  const [updateState, setUpdateState] = useState<UpdateState>(pendingUpdate ? 'available' : 'idle')
  const [updateVersion, setUpdateVersion] = useState<string | null>(pendingUpdate)
  const [downloadPercent, setDownloadPercent] = useState(0)
  const [updateErrorMsg, setUpdateErrorMsg] = useState('')
  const s = STRINGS[lang]

  // Register update listeners when modal opens
  useEffect(() => {
    // Nếu không có pending update, tự check lại khi mở Settings
    if (!pendingUpdate) {
      setUpdateState('checking')
      window.homeApi.checkForUpdate()
    }

    window.homeApi.onUpdateAvailable((info) => {
      setUpdateVersion(info.version)
      setUpdateState('available')
    })
    window.homeApi.onUpdateNotAvailable(() => {
      setUpdateState('up-to-date')
    })
    window.homeApi.onUpdateDownloadProgress((progress) => {
      setDownloadPercent(progress.percent)
      setUpdateState('downloading')
    })
    window.homeApi.onUpdateDownloaded(() => {
      setUpdateState('downloaded')
    })
    window.homeApi.onUpdateError((err) => {
      setUpdateErrorMsg(err.message)
      setUpdateState('error')
    })
  }, [])

  function handleApply() {
    onApply(draftLang, draftTheme)
  }

  function handleOverlayClick(e: React.MouseEvent<HTMLDivElement>) {
    if (e.target === e.currentTarget) onClose()
  }

  return (
    <div className="settings-overlay" onClick={handleOverlayClick}>
      <div className="settings-modal">
        <div className="settings-modal-title">{s.settings}</div>

        <div className="settings-group">
          <div className="settings-label">{s.settingsTheme}</div>
          <div className="settings-options">
            <button
              className={`settings-option${draftTheme === 'dark' ? ' active' : ''}`}
              onClick={() => setDraftTheme('dark')}
            >
              <Moon size={13} style={{ marginRight: '0.35rem', verticalAlign: 'middle' }} />
              {s.settingsDark}
            </button>
            <button
              className={`settings-option${draftTheme === 'light' ? ' active' : ''}`}
              onClick={() => setDraftTheme('light')}
            >
              <Sun size={13} style={{ marginRight: '0.35rem', verticalAlign: 'middle' }} />
              {s.settingsLight}
            </button>
          </div>
        </div>

        <div className="settings-group">
          <div className="settings-label">{s.settingsLang}</div>
          <div className="settings-options">
            <button
              className={`settings-option${draftLang === 'vi' ? ' active' : ''}`}
              onClick={() => setDraftLang('vi')}
            >
              <span style={{ fontWeight: 700, marginRight: '0.35rem' }}>VI</span>Tiếng Việt
            </button>
            <button
              className={`settings-option${draftLang === 'en' ? ' active' : ''}`}
              onClick={() => setDraftLang('en')}
            >
              <span style={{ fontWeight: 700, marginRight: '0.35rem' }}>EN</span>English
            </button>
          </div>
        </div>

        {/* ── Update section ── */}
        <div className="settings-update-section">
          <div className="settings-label" style={{ marginBottom: '0.5rem' }}>
            {lang === 'vi' ? 'Cập nhật' : 'Updates'}
          </div>

          {updateState === 'idle' && (
            <button className="settings-update-check-btn" onClick={() => {
              setUpdateState('checking')
              window.homeApi.checkForUpdate()
            }}>
              <RefreshCw size={13} style={{ marginRight: '0.4rem' }} />
              {lang === 'vi' ? 'Kiểm tra cập nhật' : 'Check for updates'}
            </button>
          )}

          {updateState === 'checking' && (
            <div className="settings-update-status">
              <Loader size={13} className="home-spin" style={{ marginRight: '0.4rem' }} />
              {lang === 'vi' ? 'Đang kiểm tra...' : 'Checking...'}
            </div>
          )}

          {updateState === 'up-to-date' && (
            <div className="settings-update-status settings-update-ok">
              {lang === 'vi' ? '✓ Bạn đang dùng phiên bản mới nhất' : '✓ You are on the latest version'}
            </div>
          )}

          {updateState === 'available' && (
            <div className="settings-update-available">
              <span className="settings-update-available-text">
                {lang === 'vi' ? `Có phiên bản mới: v${updateVersion}` : `New version available: v${updateVersion}`}
              </span>
              <button className="settings-btn-apply settings-update-dl-btn" onClick={() => {
                window.homeApi.downloadUpdate()
                setUpdateState('downloading')
              }}>
                <Download size={13} style={{ marginRight: '0.4rem' }} />
                {lang === 'vi' ? 'Tải về' : 'Download'}
              </button>
            </div>
          )}

          {updateState === 'downloading' && (
            <div className="settings-update-progress">
              <div className="settings-update-progress-bar">
                <div className="settings-update-progress-fill" style={{ width: `${downloadPercent}%` }} />
              </div>
              <span className="settings-update-progress-text">{downloadPercent}%</span>
            </div>
          )}

          {updateState === 'downloaded' && (
            <div className="settings-update-available">
              <span className="settings-update-available-text">
                {lang === 'vi' ? 'Đã tải xong — khởi động lại để cài đặt' : 'Downloaded — restart to install'}
              </span>
              <button className="settings-btn-apply settings-update-dl-btn" onClick={() => window.homeApi.installUpdate()}>
                {lang === 'vi' ? 'Khởi động lại' : 'Restart & Install'}
              </button>
            </div>
          )}

          {updateState === 'error' && (
            <div className="settings-update-status settings-update-error">
              {lang === 'vi' ? `Lỗi: ${updateErrorMsg || 'Không thể kiểm tra cập nhật'}` : `Error: ${updateErrorMsg || 'Could not check for updates'}`}
            </div>
          )}
        </div>

        <div className="settings-modal-actions">
          <button className="settings-btn-cancel" onClick={onClose}>{s.settingsCancel}</button>
          <button className="settings-btn-apply" onClick={handleApply}>{s.settingsApply}</button>
        </div>
      </div>
    </div>
  )
}

function EditorHomeScreen({
  lang,
  theme,
  s,
  onBack,
  onOpenSettings,
}: {
  lang: Lang
  theme: Theme
  s: typeof STRINGS['vi']
  onBack: () => void
  onOpenSettings: (pendingUpdate: string | null) => void
}) {
  const [recentFiles, setRecentFiles] = useState<RecentFile[]>([])
  const [loading, setLoading] = useState(true)
  const [atlasCheck, setAtlasCheck] = useState<AtlasCheckState>('idle')
  const [showAbout, setShowAbout] = useState(false)
  const [updateToast, setUpdateToast] = useState<string | null>(null) // version string nếu có update

  // Listen for update events from main process
  useEffect(() => {
    window.homeApi.onUpdateAvailable((info) => {
      setUpdateToast(info.version)
    })
    // Check if there's already a pending update (e.g. re-render after settings close)
    window.homeApi.getPendingUpdate().then((info) => {
      if (info) setUpdateToast(info.version)
    }).catch(() => {})
  }, [])

  useEffect(() => {
    window.homeApi.getRecentFiles().then((files) => {
      setRecentFiles(files)
      setLoading(false)
    }).catch(() => setLoading(false))
  }, [])

  function handleOpenAtlasWeb() {
    setAtlasCheck('checking')
    window.homeApi.onAtlasWebCheckResult((result) => {
      if (result.status === 'online') {
        // Atlas Web window opened by main process — just close overlay
        setAtlasCheck('idle')
      } else if (result.status === 'offline') {
        setAtlasCheck('offline')
      } else {
        // cancelled by user
        setAtlasCheck('idle')
      }
    })
    window.homeApi.checkNetworkForAtlas()
  }

  function handleCancelAtlasCheck() {
    window.homeApi.cancelAtlasWebCheck()
    setAtlasCheck('idle')
  }

  return (
    <div className="editor-home-root" data-theme={theme}>
      {showAbout && <AboutModal lang={lang} onClose={() => setShowAbout(false)} />}

      {/* Update toast notification */}
      {updateToast && (
        <div className="update-toast">
          <Download size={15} className="update-toast-icon" />
          <span className="update-toast-text">
            {lang === 'vi'
              ? `Có bản cập nhật mới v${updateToast} — Vào Cài đặt để tải về`
              : `Update v${updateToast} available — Go to Settings to download`}
          </span>
          <button className="update-toast-close" onClick={() => setUpdateToast(null)} title="Đóng">
            <X size={13} />
          </button>
        </div>
      )}

      {/* Atlas Web connection check overlay */}
      {atlasCheck !== 'idle' && (
        <div className="settings-overlay">
          <div className="settings-modal atlas-check-modal">
            {atlasCheck === 'checking' ? (
              <>
                <div className="atlas-check-icon"><Loader size={28} className="home-spin" /></div>
                <div className="settings-modal-title">{s.atlasCheckingTitle}</div>
                <p className="atlas-check-desc">{s.atlasCheckingDesc}</p>
                <div className="settings-modal-actions">
                  <button className="settings-btn-cancel" onClick={handleCancelAtlasCheck}>
                    {s.atlasCancel}
                  </button>
                </div>
              </>
            ) : (
              <>
                <div className="atlas-check-icon"><WifiOff size={28} /></div>
                <div className="settings-modal-title">{s.atlasOfflineTitle}</div>
                <p className="atlas-check-desc">{s.atlasOfflineDesc}</p>
                <div className="settings-modal-actions">
                  <button className="settings-btn-apply" onClick={() => setAtlasCheck('idle')}>
                    {s.atlasOfflineClose}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      <div className="editor-home-header">
        <span className="editor-home-brand">
          <img src={logoUrl} alt="Leandix Atlas" className="editor-home-brand-logo" />
          Leandix Atlas
        </span>
        <div className="editor-home-header-actions">
          <button className="editor-home-atlas-btn" onClick={handleOpenAtlasWeb} title={s.atlasDesc}>
            <Globe size={15} style={{ marginRight: '0.4rem', verticalAlign: 'middle' }} />
            {s.atlasTitle}
          </button>
          <button className="home-settings-btn editor-home-settings" onClick={() => setShowAbout(true)} title={lang === 'vi' ? 'Giới thiệu' : 'About'}>
            <HelpCircle size={15} />
          </button>
          <button className="home-settings-btn editor-home-settings" onClick={() => onOpenSettings(updateToast)} title={s.settings}>
            <Settings size={15} />
          </button>
        </div>
      </div>

      <div className="editor-home-body">
        {/* New document */}
        <div className="editor-home-new-row">
          <button className="editor-home-new-card" onClick={() => window.homeApi.openEditorNew()}>
            <div className="editor-home-new-icon"><Plus size={22} strokeWidth={1.5} /></div>
            <span>{s.newDoc}</span>
          </button>
          <button className="editor-home-new-card editor-home-open-card" onClick={async () => {
            const filePath = await window.homeApi.openFileDialog()
            if (filePath) window.homeApi.openEditorWithFile(filePath)
          }}>
            <div className="editor-home-new-icon"><FileText size={22} strokeWidth={1.5} /></div>
            <span>{s.openDoc}</span>
          </button>
          <button className="editor-home-new-card editor-home-open-pdf-card" onClick={async () => {
            const filePath = await window.homeApi.openPdfDialog()
            if (filePath) window.homeApi.openEditorWithFile(filePath)
          }}>
            <div className="editor-home-new-icon"><BookOpen size={22} strokeWidth={1.5} /></div>
            <span>{s.openPdf}</span>
          </button>
        </div>

        {/* Templates */}
        <div className="editor-home-section-header">
          <span className="editor-home-section-label">{s.templates}</span>
        </div>
        <div className="editor-home-templates">
          {[
            { id: 'report', label: s.tplReport, icon: <BarChart2 size={22} strokeWidth={1.5} /> },
            { id: 'letter', label: s.tplLetter, icon: <Mail size={22} strokeWidth={1.5} /> },
            { id: 'notes',  label: s.tplNotes,  icon: <NotebookPen size={22} strokeWidth={1.5} /> },
          ].map(tpl => (
            <button
              key={tpl.id}
              className="editor-home-tpl-card"
              onClick={() => window.homeApi.openEditorFromTemplate(tpl.id)}
              title={tpl.label}
            >
              <span className="editor-home-tpl-icon">{tpl.icon}</span>
              <span className="editor-home-tpl-label">{tpl.label}</span>
            </button>
          ))}
        </div>

        {/* Recent files */}
        <div className="editor-home-section-header">
          <span className="editor-home-section-label">{s.recentDocs}</span>
        </div>

        {loading ? (
          <div className="editor-home-empty"><Loader size={18} className="home-spin" /></div>
        ) : recentFiles.length === 0 ? (
          <div className="editor-home-empty">
            <Clock size={28} strokeWidth={1} opacity={0.3} />
            <span>{s.noRecent}</span>
          </div>
        ) : (
          <div className="editor-home-list">
            {recentFiles.map((file) => (
              <button
                key={file.filePath}
                className="editor-home-list-item"
                onClick={() => window.homeApi.openEditorWithFile(file.filePath)}
                title={file.filePath}
              >
                <div className="editor-home-list-icon">
                  {file.filePath.toLowerCase().endsWith('.pdf') ? <PdfIcon /> : <DocIcon />}
                </div>
                <div className="editor-home-list-info">
                  <span className="editor-home-list-name">{file.title}</span>
                  <span className="editor-home-list-path">{file.filePath}</span>
                </div>
                <span className="editor-home-list-date">
                  {formatRelativeDate(file.lastOpenedAt, lang)}
                </span>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}

export default function HomeApp() {
  const [lang, setLang] = useState<Lang>('vi')
  const [theme, setTheme] = useState<Theme>('dark')
  const [showSettings, setShowSettings] = useState(false)
  const [pendingUpdate, setPendingUpdate] = useState<string | null>(null)

  const s = STRINGS[lang]

  // Apply theme to <html> so CSS variables take effect globally
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme)
  }, [theme])

  useEffect(() => {
    // Load persisted settings on mount
    window.homeApi.getSettings().then((settings) => {
      setLang(settings.language)
      setTheme(settings.theme)
    }).catch(() => {})

    window.homeApi.onLanguage((l) => {
      if (l === 'vi' || l === 'en') setLang(l)
    })
    window.homeApi.onTheme((t) => {
      if (t === 'dark' || t === 'light') setTheme(t)
    })
  }, [])

  function handleApplySettings(newLang: Lang, newTheme: Theme) {
    setLang(newLang)
    setTheme(newTheme)
    window.homeApi.applySettings({ language: newLang, theme: newTheme })
    setShowSettings(false)
  }

  const settingsModal = showSettings ? (
    <SettingsModal
      lang={lang}
      theme={theme}
      onClose={() => setShowSettings(false)}
      onApply={handleApplySettings}
      pendingUpdate={pendingUpdate}
    />
  ) : null

  // Always open directly to the editor home screen
  return (
    <>
      {settingsModal}
      <EditorHomeScreen
        lang={lang}
        theme={theme}
        s={s}
        onBack={() => {}}
        onOpenSettings={(update) => {
          setPendingUpdate(update)
          setShowSettings(true)
        }}
      />
    </>
  )
}
