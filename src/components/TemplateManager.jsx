import { useState, useEffect, useRef, useCallback, useMemo } from 'react'

// ── Helpers ───────────────────────────────────────────────────────────────────

function prettyFromKey(key) {
  return String(key || '').replace(/^[A-Z_]+__/, '').replace(/_/g, ' ').trim()
}

function displayLabel(key, meta) {
  const label = String(meta?.label || '').trim()
  const sample = String(meta?.sample || '').trim()
  if (!label) return prettyFromKey(key)
  if (sample && label === sample) return prettyFromKey(key)
  return label
}

function buildVarIndex(catalog) {
  const order = ['GLOBAL', 'LIST_NTCV', 'LIST_VAT_LIEU', 'LIST_YC_NTCV']
  const groups = {}
  for (const [key, meta] of Object.entries(catalog || {})) {
    const owner = meta.owner || 'GLOBAL'
    if (!groups[owner]) groups[owner] = []
    groups[owner].push({
      key,
      owner,
      label: displayLabel(key, meta),
      sample: String(meta?.sample || '').trim(),
    })
  }
  for (const arr of Object.values(groups)) arr.sort((a, b) => a.label.localeCompare(b.label, 'vi'))
  return order.flatMap(o => groups[o] || [])
}

function groupVarIndex(varIndex) {
  const groups = {}
  for (const v of varIndex) {
    if (!groups[v.owner]) groups[v.owner] = []
    groups[v.owner].push(v)
  }
  return groups
}

const OWNER_LABELS = {
  GLOBAL: '🌐 Trường chung',
  LIST_NTCV: '📋 Nghiệm thu công việc',
  LIST_VAT_LIEU: '📦 Nghiệm thu vật liệu',
  LIST_YC_NTCV: '📝 Yêu cầu nghiệm thu',
}

function parseVarsFromText(text = '') {
  const found = new Set()
  const re = /\{([^{}]+)\}/g
  let m
  while ((m = re.exec(text)) !== null) found.add(m[1])
  return [...found]
}

function validateTemplateContent(content, catalog) {
  const issues = []
  const lines = String(content || '').split('\n')
  if (!String(content || '').trim()) {
    issues.push({ type: 'error', msg: 'Template đang trống.' })
    return issues
  }

  const openCount = (content.match(/\{/g) || []).length
  const closeCount = (content.match(/\}/g) || []).length
  if (openCount !== closeCount) {
    issues.push({ type: 'error', msg: 'Số dấu { và } chưa cân bằng.' })
  }

  const vars = parseVarsFromText(content)
  const unknown = vars.filter(v => !catalog[v])
  if (unknown.length) {
    issues.push({ type: 'warn', msg: `Có ${unknown.length} biến chưa tồn tại trong catalog.` })
  }

  const longLines = lines.filter(l => l.length > 220).length
  if (longLines > 0) {
    issues.push({ type: 'warn', msg: `Có ${longLines} dòng rất dài (>220 ký tự), khó đọc khi chỉnh sửa.` })
  }

  return issues
}

function renderTextWithVars(text, catalog) {
  const parts = String(text || '').split(/(\{[^{}]+\})/)
  return parts.map((part, i) => {
    if (part.startsWith('{') && part.endsWith('}')) {
      const key = part.slice(1, -1)
      const meta = catalog[key]
      const label = meta ? displayLabel(key, meta) : prettyFromKey(key)
      return (
        <span key={i} className="tpl-var" title={`${key}\n${label}`}>
          {part}
          <span className="tpl-var-label">{label}</span>
        </span>
      )
    }
    return <span key={i}>{part}</span>
  })
}

// ── Template Viewer ───────────────────────────────────────────────────────────

function TemplateViewer({ templateName, catalog, onEdit, onClose }) {
  const [paragraphs, setParagraphs] = useState(null)
  const [error, setError] = useState(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    setLoading(true); setError(null)
    window.api.readTemplate({ templateName })
      .then(r => {
        if (r.error) setError(r.error)
        else setParagraphs(r.paragraphs)
      })
      .catch(e => setError(e.message))
      .finally(() => setLoading(false))
  }, [templateName])

  return (
    <div className="tplm-viewer">
      <div className="tplm-toolbar">
        <span className="tplm-fname">👁 {templateName}</span>
        <div className="tplm-toolbar-right">
          <button className="btn btn-secondary btn-sm" onClick={() => onEdit(templateName)}>✏️ Sửa template này</button>
          <button className="btn-close" onClick={onClose}>✕</button>
        </div>
      </div>
      <div className="tplm-legend">
        <span className="tpl-var-demo">{'{biến}'}</span> = placeholder sẽ được thay bằng dữ liệu — hover để xem tên trường
      </div>
      <div className="tplm-content">
        {loading && <div className="tplm-loading">Đang đọc template…</div>}
        {error && <div className="tplm-error">⚠ {error}</div>}
        {!loading && !error && paragraphs && (
          paragraphs.filter(p => p.trim()).map((p, i) => (
            <div key={i} className="tplm-para">{renderTextWithVars(p, catalog)}</div>
          ))
        )}
      </div>
    </div>
  )
}

// ── Template Editor ───────────────────────────────────────────────────────────

function TemplateEditor({ templateName, catalog, onSaved, onClose }) {
  const [name, setName] = useState(templateName === 'new' ? '' : (templateName || ''))
  const [content, setContent] = useState('')
  const [loading, setLoading] = useState(!!templateName && templateName !== 'new')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState(null)

  const [varFilter, setVarFilter] = useState('')
  const [ownerFilter, setOwnerFilter] = useState('ALL')
  const [expandedOwners, setExpandedOwners] = useState({ GLOBAL: true, LIST_NTCV: true, LIST_VAT_LIEU: true, LIST_YC_NTCV: true })
  const [showPreview, setShowPreview] = useState(true)
  const [showVarPanel, setShowVarPanel] = useState(true)

  const textareaRef = useRef(null)

  const varIndex = useMemo(() => buildVarIndex(catalog), [catalog])
  const varGroups = useMemo(() => groupVarIndex(varIndex), [varIndex])
  const validationIssues = useMemo(() => validateTemplateContent(content, catalog), [content, catalog])
  const hasError = validationIssues.some(i => i.type === 'error')

  const currentVars = useMemo(() => parseVarsFromText(content), [content])

  const previewLines = useMemo(() => {
    const lines = String(content || '').split('\n')
    // Render as page-like paragraphs (hide pure empty lines)
    return lines.filter(line => line.trim().length > 0)
  }, [content])

  useEffect(() => {
    if (!templateName || templateName === 'new') return
    setLoading(true)
    window.api.readTemplate({ templateName })
      .then(r => {
        if (r.error) { setError(r.error); return }
        setContent((r.paragraphs || []).join('\n'))
        setName(templateName)
      })
      .catch(e => setError(e.message))
      .finally(() => setLoading(false))
  }, [templateName])

  useEffect(() => {
    const onKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault()
        handleSave()
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  })

  const insertText = useCallback((insertValue) => {
    const ta = textareaRef.current
    if (!ta) return
    ta.focus()
    const start = ta.selectionStart
    const end = ta.selectionEnd
    const newVal = content.slice(0, start) + insertValue + content.slice(end)
    setContent(newVal)
    setTimeout(() => {
      ta.selectionStart = ta.selectionEnd = start + insertValue.length
      ta.focus()
    }, 0)
  }, [content])

  const insertVar = useCallback((key) => insertText(`{${key}}`), [insertText])

  const insertBlock = useCallback((kind) => {
    if (kind === 'bb') {
      insertText(
`BIÊN BẢN NGHIỆM THU\n
Hôm nay, ngày {GLOBAL__ngay_31_thang_12_nam_2025}\n
Địa điểm: {GLOBAL__so_07_phan_chu_trinh_phuong_hac_thanh_tinh_thanh_hoa}\nNội dung: {LIST_NTCV__thi_cong_cai_tao_xay_dung_va_trang_bi_co_so_vat_chat_pgd_hac_tha}\n`
      )
    } else if (kind === 'hd') {
      insertText(
`HỢP ĐỒNG\n
Số: {GLOBAL__3112_2025_h_kt_bidv_lso_bph}\nBên A: {GLOBAL__ngan_hang_tmcp_au_tu_va_phat_trien_viet_nam}\nBên B: {GLOBAL__cong_ty_co_phan_xay_dung_va_thuong_mai_bac_phu_hung}\n`
      )
    } else if (kind === 'sign') {
      insertText('\nĐẠI DIỆN CÁC BÊN\nBên A\n\n\nBên B\n')
    }
  }, [insertText])

  const handleImport = useCallback(async () => {
    const result = await window.api.importTemplate()
    if (!result) return
    setName(result.name)
    setContent((result.paragraphs || []).join('\n'))
  }, [])

  const handleSave = useCallback(async () => {
    if (!name.trim()) { setError('Nhập tên template'); return }
    if (hasError) { setError('Cần sửa lỗi trước khi lưu'); return }

    setSaving(true); setError(null)
    try {
      const r = await window.api.saveTemplate({
        templateName: name.trim(),
        content,
        baseTemplate: templateName !== 'new' ? templateName : null,
      })
      if (r.error) { setError(r.error); return }
      onSaved(r)
    } catch (e) {
      setError(e.message)
    } finally {
      setSaving(false)
    }
  }, [name, content, templateName, onSaved, hasError])

  const filteredGroups = useMemo(() => {
    return Object.entries(varGroups)
      .filter(([owner]) => ownerFilter === 'ALL' || owner === ownerFilter)
      .map(([owner, vars]) => ({
        owner,
        vars: varFilter
          ? vars.filter(v =>
              v.key.toLowerCase().includes(varFilter.toLowerCase()) ||
              v.label.toLowerCase().includes(varFilter.toLowerCase()) ||
              (v.sample || '').toLowerCase().includes(varFilter.toLowerCase())
            )
          : vars,
      }))
      .filter(g => g.vars.length > 0)
  }, [varGroups, ownerFilter, varFilter])

  const toggleOwner = (owner) => setExpandedOwners(prev => ({ ...prev, [owner]: !prev[owner] }))

  return (
    <div className="tplm-editor">
      <div className="tplm-toolbar">
        <span className="tplm-fname">{templateName === 'new' ? '✨ Tạo template mới' : `✏️ ${templateName}`}</span>
        <div className="tplm-toolbar-right">
          <button className="btn btn-secondary btn-sm" onClick={() => setShowPreview(v => !v)}>
            {showPreview ? '🙈 Ẩn preview' : '👀 Hiện preview'}
          </button>
          <button className="btn btn-secondary btn-sm" onClick={() => setShowVarPanel(v => !v)}>
            {showVarPanel ? '📚 Ẩn biến' : '📌 Hiện biến'}
          </button>
          <button className="btn btn-secondary btn-sm" onClick={handleImport}>📥 Import DOCX</button>
          <button className="btn btn-primary btn-sm" onClick={handleSave} disabled={saving || hasError}>
            {saving ? '⏳ Đang lưu…' : '💾 Lưu (Ctrl/Cmd+S)'}
          </button>
          <button className="btn-close" onClick={onClose}>✕</button>
        </div>
      </div>

      {error && <div className="tplm-error">⚠ {error}</div>}

      <div className="tplm-edit-body">
        <div className="tplm-edit-left">
          <div className="tplm-edit-namebar">
            <label>Tên file:</label>
            <input
              value={name}
              onChange={e => setName(e.target.value)}
              placeholder="VD: Biên bản nghiệm thu.docx"
              className="tplm-name-input"
            />
          </div>

          <div className="tplm-helper-row">
            <span className="tplm-helper-title">Mẫu nhanh:</span>
            <button className="btn btn-secondary btn-sm" onClick={() => insertBlock('bb')}>+ Khung biên bản</button>
            <button className="btn btn-secondary btn-sm" onClick={() => insertBlock('hd')}>+ Khung hợp đồng</button>
            <button className="btn btn-secondary btn-sm" onClick={() => insertBlock('sign')}>+ Khối ký</button>
          </div>

          <div className="tplm-issues">
            {validationIssues.length === 0 && <div className="tplm-issue-ok">✅ Template hợp lệ</div>}
            {validationIssues.map((i, idx) => (
              <div key={idx} className={`tplm-issue tplm-issue-${i.type}`}>{i.type === 'error' ? '⛔' : '⚠️'} {i.msg}</div>
            ))}
            <div className="tplm-issue-meta">Biến đã dùng: <strong>{currentVars.length}</strong></div>
          </div>

          {loading
            ? <div className="tplm-loading">Đang tải nội dung…</div>
            : <textarea
                ref={textareaRef}
                className="tplm-textarea"
                value={content}
                onChange={e => setContent(e.target.value)}
                placeholder={'Nhập nội dung template ở đây...\nVí dụ:\nBIÊN BẢN NGHIỆM THU\nNgày: {LIST_NTCV__ngay_lap_bb}\nCông việc: {LIST_NTCV__ten_cong_viec}'}
                spellCheck={false}
              />}
        </div>

        {showPreview && (
          <div className="tplm-preview-panel">
            <div className="tplm-var-header">🪞 Preview trực quan (gần giống Word)</div>
            <div className="tplm-preview-content">
              <div className="tplm-preview-page">
                {previewLines.length === 0 && (
                  <div className="tplm-preview-empty-doc">(Template trống)</div>
                )}
                {previewLines.map((line, i) => (
                  <div key={i} className="tplm-preview-line tplm-preview-line-paragraph">
                    <span className="tplm-preview-ln">{i + 1}</span>
                    <span className="tplm-preview-tx">{renderTextWithVars(line, catalog)}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {showVarPanel && <div className="tplm-var-panel">
          <div className="tplm-var-header">📌 Chọn biến để chèn</div>
          <div className="tplm-var-filters">
            <input
              className="tplm-var-search"
              placeholder="Tìm key / label / sample…"
              value={varFilter}
              onChange={e => setVarFilter(e.target.value)}
            />
            <select className="tplm-owner-select" value={ownerFilter} onChange={e => setOwnerFilter(e.target.value)}>
              <option value="ALL">Tất cả nhóm</option>
              <option value="GLOBAL">GLOBAL</option>
              <option value="LIST_NTCV">LIST_NTCV</option>
              <option value="LIST_VAT_LIEU">LIST_VAT_LIEU</option>
              <option value="LIST_YC_NTCV">LIST_YC_NTCV</option>
            </select>
          </div>

          <div className="tplm-var-list">
            {filteredGroups.map(({ owner, vars }) => (
              <div key={owner} className="tplm-var-group">
                <div className="tplm-var-group-hdr" onClick={() => toggleOwner(owner)}>
                  <span>{OWNER_LABELS[owner] || owner} ({vars.length})</span>
                  <span>{expandedOwners[owner] ? '▼' : '▶'}</span>
                </div>
                {expandedOwners[owner] && vars.map(v => (
                  <div key={v.key} className="tplm-var-item" onClick={() => insertVar(v.key)} title={v.key}>
                    <div className="tplm-var-item-key">{v.key}</div>
                    <div className="tplm-var-item-label">{v.label}</div>
                    {v.sample && <div className="tplm-var-item-sample">Ví dụ: {v.sample}</div>}
                  </div>
                ))}
              </div>
            ))}
            {filteredGroups.length === 0 && (
              <div style={{ padding: 12, color: '#9ca3af', fontSize: 12 }}>Không tìm thấy biến</div>
            )}
          </div>
        </div>}
      </div>
    </div>
  )
}

// ── Template List ─────────────────────────────────────────────────────────────

export default function TemplateManager({ catalog, templates, onClose, onCatalogUpdate, onTemplatesUpdate }) {
  const [view, setView] = useState('list')
  const [selected, setSelected] = useState(null)
  const [toast, setToast] = useState(null)
  const [search, setSearch] = useState('')

  const showToast = (msg, type = 'success') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 3000)
  }

  const filteredTemplates = useMemo(() => {
    const q = search.trim().toLowerCase()
    if (!q) return templates
    return templates.filter(t => t.toLowerCase().includes(q))
  }, [templates, search])

  const handleView = (tpl) => { setSelected(tpl); setView('viewer') }
  const handleEdit = (tpl) => { setSelected(tpl); setView('editor') }
  const handleNew = () => { setSelected('new'); setView('editor') }

  const handleDelete = async (tpl) => {
    if (!confirm(`Xóa template "${tpl}"? Không thể hoàn tác.`)) return
    const r = await window.api.deleteTemplate({ templateName: tpl })
    if (r.error) { showToast(r.error, 'error'); return }
    onTemplatesUpdate(templates.filter(t => t !== tpl))
    showToast(`Đã xóa: ${tpl}`)
  }

  const handleSaved = async ({ name, newVars }) => {
    const newTemplates = await window.api.getTemplates()
    onTemplatesUpdate(newTemplates)
    const newCatalog = await window.api.getCatalog()
    onCatalogUpdate(newCatalog)
    const msg = newVars?.length
      ? `Đã lưu "${name}". Thêm ${newVars.length} biến mới vào catalog.`
      : `Đã lưu "${name}"`
    showToast(msg)
    setView('list')
  }

  return (
    <div className="overlay" onClick={view === 'list' ? onClose : undefined}>
      <div className={`modal tplm-modal ${view !== 'list' ? 'tplm-modal-wide' : ''}`}
           onClick={e => e.stopPropagation()}>
        <div className="modal-hdr">
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {view !== 'list' && (
              <button className="btn btn-secondary btn-sm" onClick={() => setView('list')}>← Danh sách</button>
            )}
            <span>📐 Quản lý Templates</span>
          </div>
          <button className="btn-close" onClick={onClose}>✕</button>
        </div>

        {toast && (
          <div className={`tplm-toast tplm-toast-${toast.type}`}>{toast.msg}</div>
        )}

        {view === 'list' && (
          <div className="modal-body tplm-list-body">
            <div className="tplm-list-toolbar tplm-list-toolbar-2row">
              <div>
                <button className="btn btn-primary btn-sm" onClick={handleNew}>✨ Tạo template mới</button>
              </div>
              <div className="tplm-list-search-wrap">
                <input
                  className="tplm-list-search"
                  placeholder="Tìm theo tên template…"
                  value={search}
                  onChange={e => setSearch(e.target.value)}
                />
                <span className="tplm-list-count">{filteredTemplates.length}/{templates.length}</span>
              </div>
            </div>

            {filteredTemplates.length === 0 && (
              <div style={{ padding: 20, color: '#9ca3af', textAlign: 'center' }}>
                {templates.length === 0 ? 'Chưa có template nào' : 'Không tìm thấy template phù hợp'}
              </div>
            )}

            {filteredTemplates.map(tpl => (
              <div key={tpl} className="tplm-list-item">
                <span className="tplm-list-name">📄 {tpl}</span>
                <div className="tplm-list-actions">
                  <button className="btn btn-secondary btn-sm" onClick={() => handleView(tpl)}>👁 Xem</button>
                  <button className="btn btn-secondary btn-sm" onClick={() => handleEdit(tpl)}>✏️ Sửa</button>
                  <button className="btn btn-danger btn-sm" onClick={() => handleDelete(tpl)}>🗑</button>
                </div>
              </div>
            ))}
          </div>
        )}

        {view === 'viewer' && selected && (
          <TemplateViewer
            templateName={selected}
            catalog={catalog}
            onEdit={handleEdit}
            onClose={() => setView('list')}
          />
        )}

        {view === 'editor' && (
          <TemplateEditor
            templateName={selected}
            catalog={catalog}
            onSaved={handleSaved}
            onClose={() => setView('list')}
          />
        )}
      </div>
    </div>
  )
}
