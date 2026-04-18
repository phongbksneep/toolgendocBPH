import { useState, useEffect, useCallback, useMemo, useRef } from 'react'
import FieldForm from './components/FieldForm.jsx'
import TemplateManager from './components/TemplateManager.jsx'

// ── helpers ───────────────────────────────────────────────────────────────────

function extractValue(v) {
  if (v === null || v === undefined) return ''
  if (typeof v === 'object' && 'value' in v) return String(v.value ?? '')
  return String(v)
}

// ── Auto-save (file-based via IPC) ────────────────────────────────────────
function saveAutosave(project, file) {
  window.api?.saveAutosave({ project, currentFile: file, savedAt: Date.now() })
}
function clearAutosave() {
  window.api?.clearAutosave()
}

// ── Catalog-based grouping & job-field helpers ──────────────────────────

// Compute {templateToGroup, groupLabels} from catalog once
function buildCatalogGroups(catalog) {
  const templateToGroup = new Map()   // 'Template.docx' → 'LIST_XXX'
  const groupFiles = new Map()         // 'LIST_XXX' → Set<filename>

  for (const [, meta] of Object.entries(catalog)) {
    if (meta.owner && meta.owner !== 'GLOBAL' && Array.isArray(meta.files)) {
      if (!groupFiles.has(meta.owner)) groupFiles.set(meta.owner, new Set())
      for (const f of meta.files) {
        templateToGroup.set(f, meta.owner)
        groupFiles.get(meta.owner).add(f)
      }
    }
  }

  // Friendly label = longest common word-prefix of all filenames in the group
  const groupLabels = new Map()
  for (const [owner, files] of groupFiles) {
    const names = [...files].map(f => f.replace(/\.docx$/i, '').trim())
    if (!names.length) { groupLabels.set(owner, owner); continue }
    const words0 = names[0].split(/\s+/)
    let n = words0.length
    for (const nm of names.slice(1)) {
      const w = nm.split(/\s+/); let i = 0
      while (i < n && i < w.length && words0[i] === w[i]) i++
      n = i
    }
    groupLabels.set(owner, n >= 2 ? words0.slice(0, n).join(' ') + ' …' : names[0] + ' …')
  }
  return { templateToGroup, groupLabels }
}

// Build initial job.fields from catalog (fields with non-GLOBAL owner for this template)
function buildJobFields(tpl, catalog) {
  const fields = {}
  for (const [key, meta] of Object.entries(catalog)) {
    if (meta.owner !== 'GLOBAL' && Array.isArray(meta.files) && meta.files.includes(tpl))
      fields[key] = ''
  }
  return fields
}

// Numeric-prefix acronym group detector (for templates like "2.BB Ktra..." → "BB")
function numericPrefixGroup(name) {
  const bare = (name || '').replace(/\.docx$/i, '')
  const m = bare.match(/^\d+[a-z]*[.\s]+\s*(\S+)/)
  if (!m) return null
  const first = m[1]
  return first.length >= 2 && first === first.toUpperCase() &&
    /[A-Z\u0110\u01a0-\u01b0\u00c0-\u024f]/.test(first) ? first : null
}

function groupJobs(jobs, catalogGroups) {
  const { templateToGroup, groupLabels } = catalogGroups || { templateToGroup: new Map(), groupLabels: new Map() }

  // Assign effective group for each job
  const effective = jobs.map(job => {
    if (job.group) return job.group // explicit manual group
    const tpl = job.template || job.output || ''
    // 1️⃣ catalog LIST_xxx group
    const cg = templateToGroup.get(tpl)
    if (cg) return groupLabels.get(cg) || cg
    // GLOBAL templates are never auto-grouped
    return null
  })

  // Only apply numeric prefix group when ≥2 jobs share it
  const numCnt = {}
  effective.forEach((g, i) => {
    if (g && !jobs[i].group && !templateToGroup.get(jobs[i].template || '')) {
      numCnt[g] = (numCnt[g] || 0) + 1
    }
  })

  const out = [], map = new Map()
  jobs.forEach((job, idx) => {
    const tpl  = job.template || job.output || ''
    const cg   = templateToGroup.get(tpl)        // catalog group (always valid if present)
    const rawG = effective[idx]
    const useGroup = rawG && (job.group || cg || (numCnt[rawG] || 0) >= 2)
    if (useGroup) {
      if (!map.has(rawG)) {
        const e = { type: 'group', groupName: rawG, items: [] }
        map.set(rawG, e); out.push(e)
      }
      map.get(rawG).items.push({ job, idx })
    } else {
      out.push({ type: 'single', job, idx })
    }
  })
  return out
}

function listToDict(arr) {
  if (!Array.isArray(arr)) return arr || {}
  return Object.fromEntries(arr.filter(e => e && e.key).map(e => [e.key, extractValue(e.value)]))
}

function normalizeProject(raw, catalog) {
  if (!raw || typeof raw !== 'object') return buildDefault(catalog)

  const mergeJobFields = (_j, parsedFields) => {
    // Keep exactly the fields present in input JSON/autosave.
    // This avoids showing duplicate/irrelevant fields when many jobs share one template
    // but each job only needs a subset of values.
    return { ...parsedFields }
  }

  // Defaults ensure newly-added catalog GLOBAL keys are always present
  const defaultGlobal = buildDefault(catalog).global

  // v3: global_fields as array (only when the field actually exists)
  if (raw.global_fields) {
    const globalFields = { ...defaultGlobal, ...listToDict(raw.global_fields) }
    const jobs = (raw.jobs || []).map(j => ({
      ...j,
      fields: mergeJobFields(j, listToDict(j.fields || [])),
    }))
    return { meta: raw.meta || {}, global: globalFields, jobs, notes: raw.notes || {} }
  }

  // v1/v2 and internal autosave format: global as dict
  const globalRaw = raw.global || {}
  const globalFields = {
    ...defaultGlobal,
    ...Object.fromEntries(Object.entries(globalRaw).map(([k, v]) => [k, extractValue(v)])),
  }
  const jobs = (raw.jobs || []).map(j => {
    const fieldsRaw = j.fields || {}
    const parsed = Array.isArray(fieldsRaw)
      ? listToDict(fieldsRaw)
      : Object.fromEntries(Object.entries(fieldsRaw).map(([k, v]) => [k, extractValue(v)]))
    return { ...j, fields: mergeJobFields(j, parsed) }
  })
  return { meta: raw.meta || {}, global: globalFields, jobs, notes: raw.notes || {} }
}

function buildDefault(catalog) {
  const globalFields = {}
  for (const [key, meta] of Object.entries(catalog)) {
    if (meta.owner === 'GLOBAL') globalFields[key] = ''
  }
  return { meta: {}, global: globalFields, jobs: [], notes: {} }
}

const GLOBAL_ALIAS_KEYS = {
  // Địa chỉ nhà thầu (biến thể giữa HĐ chính và phụ lục)
  GLOBAL__lo_25tt1_o_thi_my_inh_me_tri_phuong_tu_liem_ha_noi: [
    'GLOBAL__so_25tt1_o_thi_my_inh_me_tri_phuong_tu_liem_thanh_pho_ha_noi',
  ],

  // Số tài khoản nhà thầu (biến thể ngân hàng)
  GLOBAL__0621101126007_mo_tai_ngan_hang_mb_cn_ien_bien_phu: [
    'GLOBAL__0621101126007_mo_tai_ngan_hang_tmcp_quan_oi_chi_nhanh_ien_bien_p',
  ],
}

// Các key phụ được auto-sync từ key chính và ẩn khỏi UI để tránh nhập trùng
const GLOBAL_ALIAS_HIDDEN_KEYS = [
  'GLOBAL__so_25tt1_o_thi_my_inh_me_tri_phuong_tu_liem_thanh_pho_ha_noi',
  'GLOBAL__0621101126007_mo_tai_ngan_hang_tmcp_quan_oi_chi_nhanh_ien_bien_p',
]

function syncGlobalAliases(fields, changedKey) {
  if (!changedKey) return fields
  const linked = GLOBAL_ALIAS_KEYS[changedKey] || []
  if (!linked.length) return fields

  const val = fields[changedKey] ?? ''
  const next = { ...fields }
  linked.forEach(k => {
    if (k in next) next[k] = val
  })
  return next
}

// ── App ───────────────────────────────────────────────────────────────────────

const isElectron = typeof window !== 'undefined' && !!window.api

export default function App() {
  const [catalog, setCatalog]               = useState({})
  const [templates, setTemplates]           = useState([])
  const [project, setProject]               = useState({ meta: {}, global: {}, jobs: [], notes: {} })
  const [currentFile, setCurrentFile]       = useState(null)
  const [selectedJobIdx, setSelectedJobIdx] = useState(null)
  const [mode, setMode]                     = useState('global') // 'global' | 'job'
  const [toast, setToast]                   = useState(null)     // {message, type, id}
  const [genResult, setGenResult]           = useState(null)
  const [collapsedGroups, setCollapsedGroups] = useState(new Set())
  const [showAddJob, setShowAddJob]         = useState(false)
  const [generating, setGenerating]         = useState(false)
  const [updateInfo, setUpdateInfo]         = useState(null)
  const [appVersion, setAppVersion]         = useState('1.0.0')
  const [platform, setPlatform]             = useState(null)  // {platform, arch}
  const [showTemplateManager, setShowTemplateManager] = useState(false)
  const [showExportMenu, setShowExportMenu]   = useState(false)
  const [showDataMenu, setShowDataMenu]       = useState(false)
  const timerRef = useRef()
  const initialLoad = useRef(true)
  const exportMenuRef = useRef(null)
  const dataMenuRef = useRef(null)

  // ── bootstrap ────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!isElectron) return
    Promise.resolve(window.api.refreshRuntimeAssets?.())
      .then(() => Promise.all([window.api.getCatalog(), window.api.getTemplates(), window.api.loadAutosave()]))
      .then(([cat, tpls, saved]) => {
        setCatalog(cat); setTemplates(tpls)
        const hasData = saved?.project?.jobs?.length > 0
          || Object.values(saved?.project?.global || {}).some(v => v)
        if (hasData) {
          const p = normalizeProject(saved.project, cat)
          setProject(p)
          setCurrentFile(saved.currentFile || null)
          if (p.jobs.length) setSelectedJobIdx(0)
          // suppress first auto-save tick (data just came from file)
          initialLoad.current = false
          setTimeout(() => { initialLoad.current = true }, 100)
        } else {
          setProject(buildDefault(cat))
        }
      })
      .catch(e => showToast('Lỗi khởi động: ' + e.message, 'error'))
    window.api.checkUpdate?.().then(info => { if (info?.hasUpdate) setUpdateInfo(info) }).catch(() => {})
    window.api.getVersion?.().then(v => { if (v) setAppVersion(v) }).catch(() => {})
    window.api.getPlatform?.().then(p => { if (p) setPlatform(p) }).catch(() => {})
  }, [])

  // auto-save debounced
  useEffect(() => {
    if (!isElectron) return
    clearTimeout(timerRef.current)
    timerRef.current = setTimeout(() => saveAutosave(project, currentFile), 1500)
    return () => clearTimeout(timerRef.current)
  }, [project, currentFile])

  // auto-dismiss toast
  useEffect(() => {
    if (!toast) return
    const t = setTimeout(() => setToast(null), toast.type === 'error' ? 7000 : 4000)
    return () => clearTimeout(t)
  }, [toast?.id])

  // close dropdown menus when clicking outside or pressing ESC
  useEffect(() => {
    if (!showExportMenu && !showDataMenu) return
    const onDown = (e) => {
      if (showExportMenu && exportMenuRef.current && !exportMenuRef.current.contains(e.target)) {
        setShowExportMenu(false)
      }
      if (showDataMenu && dataMenuRef.current && !dataMenuRef.current.contains(e.target)) {
        setShowDataMenu(false)
      }
    }
    const onEsc = (e) => {
      if (e.key !== 'Escape') return
      setShowExportMenu(false)
      setShowDataMenu(false)
    }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onEsc)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onEsc)
    }
  }, [showExportMenu, showDataMenu])

  function showToast(message, type = 'success') {
    setToast({ message, type, id: Date.now() })
  }


  // ── File operations ────────────────────────────────────────────────────────

  const newProject = useCallback(() => {
    if (!confirm('Tạo dự án mới? Dữ liệu đang nhập đã được tự động lưu.')) return
    setProject(buildDefault(catalog)); setCurrentFile(null); setSelectedJobIdx(null)
    showToast('Dự án mới đã sẵn sàng', 'info')
  }, [catalog])

  const openJson = useCallback(async () => {
    const result = await window.api.openJson()
    if (!result) return
    const normalized = normalizeProject(result.data, catalog)
    setProject(normalized); setCurrentFile(result.path)
    setSelectedJobIdx(normalized.jobs.length > 0 ? 0 : null)
    showToast(`Đã mở: ${result.path.split('/').pop()}`, 'info')
  }, [catalog])

  const loadDemoData = useCallback(async () => {
    const hasCurrentData = project.jobs.length > 0 || Object.values(project.global || {}).some(v => String(v || '').trim() !== '')
    if (hasCurrentData) {
      const ok = confirm('Hiện đang có dữ liệu đã nhập. Nạp dữ liệu demo sẽ ghi đè dữ liệu hiện tại. Tiếp tục?')
      if (!ok) return
    }

    const result = await window.api.loadDemoProject?.()
    if (!result) return
    if (result.error) {
      showToast('Không mở được dữ liệu demo: ' + result.error, 'error')
      return
    }
    const normalized = normalizeProject(result.data, catalog)
    setProject(normalized)
    setCurrentFile(result.path || 'project-data.demo.friendly.json')
    setSelectedJobIdx(normalized.jobs.length > 0 ? 0 : null)
    setMode('global')
    showToast('Đã nạp dữ liệu demo', 'success')
  }, [catalog, project])

  const saveJson = useCallback(async (saveAs = false) => {
    const path = await window.api.saveJson({ filePath: saveAs ? null : currentFile, data: project })
    if (path) { setCurrentFile(path); clearAutosave(); showToast(`Đã lưu: ${path.split('/').pop()}`) }
  }, [project, currentFile])

  const generateDocs = useCallback(async () => {
    const enabled = project.jobs.filter(j => j.enabled !== false)
    if (!enabled.length) { showToast('Không có job nào được bật để tạo file', 'error'); return }
    setGenerating(true)
    try {
      const r = await window.api.generateDocs({ jobs: project.jobs, globalFields: project.global })
      if (r.canceled) { showToast('Đã hủy', 'info'); return }
      setGenResult(r)
    } catch (e) {
      showToast('Lỗi khi tạo file: ' + e.message, 'error')
    } finally { setGenerating(false) }
  }, [project])

  const clearAll = useCallback(() => {
    if (!confirm('Xóa toàn bộ dữ liệu hiện tại? Không thể hoàn tác.')) return
    clearAutosave(); setProject(buildDefault(catalog)); setCurrentFile(null); setSelectedJobIdx(null)
    showToast('Đã xóa toàn bộ dữ liệu', 'info')
  }, [catalog])

  const checkUpdate = useCallback(async () => {
    showToast('Đang kiểm tra cập nhật…', 'info')
    try {
      const info = await window.api.checkUpdate()
      if (!info || info.error) { showToast(info?.error || 'Chưa cấu hình URL cập nhật', 'info'); return }
      if (info.hasUpdate) { setUpdateInfo(info); showToast(`Có phiên bản mới: ${info.version}`, 'info') }
      else showToast('Đang dùng phiên bản mới nhất ✓', 'success')
    } catch (e) { showToast('Không kiểm tra được: ' + e.message, 'error') }
  }, [])

  const handleRequestEditLabel = useCallback(async ({ key, currentLabel, sample }) => {
    if (!window.api?.editLabelLocal) {
      showToast('Tính năng này chỉ dùng trong bản desktop mới', 'error')
      return
    }

    const next = prompt(
      `Sửa label cho ${key}\nVí dụ hiện tại: ${sample || '(trống)'}`,
      currentLabel || ''
    )
    if (next === null) return
    if (!String(next).trim()) {
      showToast('Label không được để trống', 'error')
      return
    }

    const r = await window.api.editLabelLocal({
      key,
      label: String(next).trim(),
    })

    if (!r || r.error) {
      showToast(r?.error || 'Không sửa được label', 'error')
      return
    }

    const newCatalog = await window.api.getCatalog()
    setCatalog(newCatalog)
    showToast(`Đã cập nhật label: ${key}`, 'success')
  }, [])

  const exportLabelPack = useCallback(async () => {
    if (!window.api?.exportLabelPack) {
      showToast('Tính năng này chỉ dùng trong bản desktop mới', 'error')
      return
    }
    const r = await window.api.exportLabelPack()
    if (!r) return
    if (r.error) { showToast(r.error, 'error'); return }
    showToast('Đã xuất gói label theo phiên bản hiện tại', 'success')
  }, [])

  const importLabelPack = useCallback(async () => {
    if (!window.api?.importLabelPack) {
      showToast('Tính năng này chỉ dùng trong bản desktop mới', 'error')
      return
    }
    const ok = confirm('Import sẽ ghi đè catalog/demo label hiện tại trên máy này. Tiếp tục?')
    if (!ok) return
    const r = await window.api.importLabelPack()
    if (!r) return
    if (r.error) { showToast(r.error, 'error'); return }

    const newCatalog = await window.api.getCatalog()
    setCatalog(newCatalog)
    showToast('Đã import gói label và ghi đè local', 'success')
  }, [])

  // ── Project mutations ──────────────────────────────────────────────────────

  const setGlobalFields = useCallback(fields => {
    // Khi người dùng sửa 1 key, auto sync các key alias tương đương để tránh phải nhập lặp
    const prev = project.global || {}
    let changedKey = null
    for (const k of Object.keys(fields)) {
      if ((prev[k] ?? '') !== (fields[k] ?? '')) { changedKey = k; break }
    }
    const merged = syncGlobalAliases(fields, changedKey)
    setProject(p => ({ ...p, global: merged }))
  }, [project.global])

  const setJobFields = useCallback(fields => {
    if (selectedJobIdx === null) return
    setProject(p => { const j = [...p.jobs]; j[selectedJobIdx] = { ...j[selectedJobIdx], fields }; return { ...p, jobs: j } })
  }, [selectedJobIdx])

  const setJobProp = useCallback((prop, val) => {
    if (selectedJobIdx === null) return
    setProject(p => {
      const j = [...p.jobs]
      j[selectedJobIdx] = { ...j[selectedJobIdx], [prop]: val !== '' ? val : undefined }
      return { ...p, jobs: j }
    })
  }, [selectedJobIdx])

  const toggleJobEnabled = useCallback((idx, e) => {
    e?.stopPropagation()
    setProject(p => { const j = [...p.jobs]; j[idx] = { ...j[idx], enabled: !(j[idx].enabled !== false) }; return { ...p, jobs: j } })
  }, [])

  const toggleGroupEnabled = useCallback((groupIndices, e) => {
    e?.stopPropagation()
    setProject(p => {
      const allOn = groupIndices.every(i => p.jobs[i]?.enabled !== false)
      const j = [...p.jobs]
      groupIndices.forEach(i => { if (j[i]) j[i] = { ...j[i], enabled: !allOn } })
      return { ...p, jobs: j }
    })
  }, [])

  const importExcel = useCallback(async () => {
    const result = await window.api.importExcel()
    if (!result) return
    const p = normalizeProject(result.project, catalog)
    setProject(p)
    setMode('global')
    setSelectedJobIdx(null)
    showToast(`Đã import: ${result.counts.global} trường chung, ${result.counts.jobs} job`, 'success')
  }, [catalog])

  const cloneJob = useCallback(() => {
    if (selectedJobIdx === null) return
    const orig = project.jobs[selectedJobIdx]
    const job = { ...orig, output: (orig.output || orig.template || '').replace(/\.docx$/i, '-bản2.docx') }
    const jobs = [...project.jobs]; jobs.splice(selectedJobIdx + 1, 0, job)
    setProject(p => ({ ...p, jobs })); setSelectedJobIdx(selectedJobIdx + 1)
  }, [project.jobs, selectedJobIdx])

  const deleteJob = useCallback(() => {
    if (selectedJobIdx === null || !confirm('Xóa job đang chọn?')) return
    const jobs = [...project.jobs]; jobs.splice(selectedJobIdx, 1)
    setProject(p => ({ ...p, jobs }))
    setSelectedJobIdx(jobs.length ? Math.min(selectedJobIdx, jobs.length - 1) : null)
    if (!jobs.length) setMode('global')
  }, [project.jobs, selectedJobIdx])

  const addJob = useCallback(tpl => {
    const fields = buildJobFields(tpl, catalog)
    const jobs = [...project.jobs, { template: tpl, output: tpl, enabled: true, fields }]
    setProject(p => ({ ...p, jobs })); setSelectedJobIdx(jobs.length - 1)
    setMode('job'); setShowAddJob(false); showToast(`Đã thêm: ${tpl}`)
  }, [project.jobs, catalog])

  const toggleGroup = useCallback(name => {
    setCollapsedGroups(s => { const n = new Set(s); n.has(name) ? n.delete(name) : n.add(name); return n })
  }, [])

  // ── Render ────────────────────────────────────────────────────────────────

  const selectedJob    = selectedJobIdx !== null ? project.jobs[selectedJobIdx] : null
  const catalogGroups  = useMemo(() => buildCatalogGroups(catalog), [catalog])
  const grouped        = useMemo(() => groupJobs(project.jobs, catalogGroups), [project.jobs, catalogGroups])
  const enabledCount   = project.jobs.filter(j => j.enabled !== false).length
  const totalJobs      = project.jobs.length

  return (
    <>
      {/* Top bar */}
      <div className="topbar">
        <span className="app-title">BPHH DocGen</span>
        <span className="app-version">v{appVersion}</span>
        <div className="topbar-group">
          <button className="btn btn-secondary" onClick={newProject}>Tạo mới</button>

          <div className="topbar-dropdown" style={{ position: 'relative' }} ref={dataMenuRef}>
            <button
              className="btn btn-secondary"
              aria-haspopup="menu"
              aria-expanded={showDataMenu}
              onClick={() => setShowDataMenu(v => !v)}>
              📁 Dữ liệu
            </button>
            {showDataMenu && (
              <div className="topbar-menu" role="menu">
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowDataMenu(false); openJson() }}>
                  📂 Mở JSON
                </button>
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowDataMenu(false); loadDemoData() }}>
                  🧪 Nạp dữ liệu demo
                </button>
                <div className="topbar-menu-sep" />
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowDataMenu(false); saveJson(false) }}>
                  💾 Lưu
                </button>
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowDataMenu(false); saveJson(true) }}>
                  💾 Lưu thành…
                </button>
                <div className="topbar-menu-sep" />
                <div className="topbar-menu-section">─ Nâng cao (label pack)</div>
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowDataMenu(false); exportLabelPack() }}>
                  📦 Export label pack (kèm version)
                </button>
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowDataMenu(false); importLabelPack() }}>
                  📥 Import label pack (ghi đè)
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="topbar-group">
          {/* Export/Import dropdown */}
          <div className="topbar-dropdown" style={{ position: 'relative' }} ref={exportMenuRef}>
            <button
              className="btn btn-secondary"
              aria-haspopup="menu"
              aria-expanded={showExportMenu}
              onClick={() => setShowExportMenu(v => !v)}>
              📤 Xuất / Nhập
            </button>
            {showExportMenu && (
              <div className="topbar-menu" role="menu">
                <div className="topbar-menu-section">─ Xuất mẫu trống</div>
                <button className="topbar-menu-item" role="menuitem" onClick={async () => {
                  setShowExportMenu(false)
                  const p = await window.api.exportSampleJson()
                  if (p) showToast('Đã xuất JSON mẫu: ' + p.split(/[\/\\]/).pop(), 'success')
                }}>📄 File JSON mẫu</button>
                <button className="topbar-menu-item" role="menuitem" onClick={async () => {
                  setShowExportMenu(false)
                  const p = await window.api.exportSampleExcel()
                  if (p) showToast('Đã xuất Excel mẫu: ' + p.split(/[\/\\]/).pop(), 'success')
                }}>📊 File Excel mẫu (.xlsx)</button>
                <div className="topbar-menu-sep" />
                <div className="topbar-menu-section">─ Xuất dữ liệu hiện tại</div>
                <button className="topbar-menu-item" role="menuitem" onClick={async () => {
                  setShowExportMenu(false)
                  const p = await window.api.exportCurrentExcel({ project })
                  if (p) showToast('Đã xuất Excel dữ liệu: ' + p.split(/[\/\\]/).pop(), 'success')
                }}>📤 Export dữ liệu đang nhập → Excel</button>
                <div className="topbar-menu-sep" />
                <div className="topbar-menu-section">─ Nhập dữ liệu</div>
                <button className="topbar-menu-item" role="menuitem" onClick={() => { setShowExportMenu(false); importExcel() }}>
                  📊 Import từ Excel
                </button>
              </div>
            )}
          </div>
          <button className="btn btn-secondary" onClick={() => setShowTemplateManager(true)}>📐 Templates</button>
          <button className="btn btn-ghost-danger" onClick={clearAll} title="Xóa toàn bộ dữ liệu">🗑 Xóa</button>
          <button className="btn btn-secondary btn-sm" onClick={checkUpdate} title="Kiểm tra cập nhật">
            {updateInfo ? '↑ Có cập nhật' : '⟳ Cập nhật'}
          </button>
        </div>
        <span className="spacer" />
        <span className={`enabled-badge ${enabledCount === 0 ? 'zero' : ''}`}>
          {enabledCount}/{totalJobs} job
        </span>
        <button className="btn btn-generate" onClick={generateDocs} disabled={generating}>
          {generating ? '⏳ Đang tạo…' : '▶ Generate DOCX'}
        </button>
      </div>

      {/* Path bar */}
      <div className="path-bar">
        <span><strong>File:</strong> {currentFile || '(chưa lưu)'}</span>
        <span className="autosave-dot">● tự động lưu</span>
      </div>

      {/* Update banner */}
      {updateInfo && (
        <div className="update-banner">
          <span>
            🗒️ Phímbản mới <strong>{updateInfo.version}</strong>
            {updateInfo.changelog ? ` — ${updateInfo.changelog.slice(0, 80)}` : ''}
          </span>
          {updateInfo.installUrl && (
            <button className="btn btn-primary" style={{ padding: '3px 12px', fontSize: 12 }}
              onClick={async () => {
                showToast('Đang tải bản cài… (~vài phút)', 'info')
                const r = await window.api.installUpdate({ url: updateInfo.installUrl })
                if (r?.ok) showToast('Đã tải xong! Mở file cài đặt từ thư mục Downloads ✅', 'success')
                else showToast('Lỗi tải: ' + (r?.error || '?'), 'error')
              }}>
              📥 Tải &amp; Cài
            </button>
          )}
          <button className="btn btn-secondary" style={{ padding: '3px 10px', fontSize: 12 }}
            onClick={() => setUpdateInfo(null)}>Đóng</button>
        </div>
      )}

      {/* Main body */}
      <div className="main-body">

        {/* LEFT: job list */}
        <div className="job-panel">
          <div className="job-panel-header">
            <span>Danh sách jobs ({totalJobs})</span>
            <button className="btn-add-job" onClick={() => setShowAddJob(true)} title="Thêm job mới">＋</button>
          </div>
          <div className="job-list">
            {grouped.length === 0 && (
              <div className="job-empty">
                Nhấn <strong>＋</strong> để thêm job, hoặc <strong>Mở JSON</strong>
              </div>
            )}
            {grouped.map((entry) => {
              if (entry.type === 'single') {
                const { job, idx } = entry
                const en = job.enabled !== false
                return (
                  <div key={idx}
                    className={`job-item ${idx === selectedJobIdx ? 'selected' : ''} ${!en ? 'job-off' : ''}`}
                    onClick={() => { setSelectedJobIdx(idx); setMode('job') }}>
                    <input type="checkbox" className="job-cb" checked={en}
                      onClick={e => e.stopPropagation()} onChange={e => toggleJobEnabled(idx, e)} />
                    <div className="job-names">
                      <div className="job-output">{job.output || job.template}</div>
                      <div className="job-tpl">⟵ {job.template}</div>
                    </div>
                  </div>
                )
              }
              // group entry
              const { groupName, items } = entry
              const col    = collapsedGroups.has(groupName)
              const allOn  = items.every(({ job }) => job.enabled !== false)
              const someOn = items.some(({ job }) => job.enabled !== false)
              const cnt    = items.filter(({ job }) => job.enabled !== false).length
              return (
                <div key={groupName} className="job-group">
                  <div className="job-group-hdr" onClick={() => toggleGroup(groupName)}>
                    <input type="checkbox" className="job-cb" checked={someOn}
                      ref={el => { if (el) el.indeterminate = someOn && !allOn }}
                      onClick={e => e.stopPropagation()}
                      onChange={e => toggleGroupEnabled(items.map(({idx}) => idx), e)} />
                    <span className="group-name">{groupName}</span>
                    <span className="group-cnt">{cnt}/{items.length}</span>
                    <span className="group-chev">{col ? '▶' : '▼'}</span>
                  </div>
                  {!col && items.map(({ job, idx }) => {
                    const en = job.enabled !== false
                    return (
                      <div key={idx}
                        className={`job-item job-group-item ${idx === selectedJobIdx ? 'selected' : ''} ${!en ? 'job-off' : ''}`}
                        onClick={() => { setSelectedJobIdx(idx); setMode('job') }}>
                        <input type="checkbox" className="job-cb" checked={en}
                          onClick={e => e.stopPropagation()} onChange={e => toggleJobEnabled(idx, e)} />
                        <div className="job-names">
                          <div className="job-output">{job.output || job.template}</div>
                          <div className="job-tpl">⟵ {job.template}</div>
                        </div>
                      </div>
                    )
                  })}
                </div>
              )
            })}
          </div>
          <div className="job-actions">
            <button className="btn btn-secondary btn-sm" onClick={cloneJob} disabled={selectedJobIdx === null}>Nhân bản</button>
            <button className="btn btn-danger btn-sm" onClick={deleteJob} disabled={selectedJobIdx === null}>Xóa job</button>
          </div>
        </div>

        {/* RIGHT: form panel */}
        <div className="form-panel">
          <div className="mode-bar">
            <button className={`btn btn-sm ${mode === 'global' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setMode('global')}>
              🌐 Trường chung ({Object.keys(project.global).length})
            </button>
            <button className={`btn btn-sm ${mode === 'job' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setMode('job')}>
              📄 Trường job {selectedJob ? `(${Object.keys(selectedJob.fields || {}).length})` : ''}
            </button>
          </div>

          {mode === 'global' && (
            <FieldForm fields={project.global} catalog={catalog} onChange={setGlobalFields}
              onRequestEditLabel={handleRequestEditLabel}
              hiddenKeys={GLOBAL_ALIAS_HIDDEN_KEYS}
              hint="Các trường dùng chung cho tất cả file — điền một lần, áp dụng toàn bộ" />
          )}

          {mode === 'job' && !selectedJob && (
            <div className="empty-state">
              <div className="empty-icon">👈</div>
              <h3>Chọn một job bên trái</h3>
              <p>Sau đó điền dữ liệu riêng của job đó tại đây</p>
            </div>
          )}

          {mode === 'job' && selectedJob && (
            <>
              <div className="job-meta">
                <div className="job-meta-row">
                  <label>Template</label>
                  <span className="meta-mono">{selectedJob.template}</span>
                </div>
                <div className="job-meta-row">
                  <label>Tên file output</label>
                  <input value={selectedJob.output || ''} onChange={e => setJobProp('output', e.target.value)} />
                </div>
                <div className="job-meta-row">
                  <label>Nhóm</label>
                  <input value={selectedJob.group || ''} placeholder="(vd: Yêu cầu nghiệm thu)"
                    onChange={e => setJobProp('group', e.target.value)} />
                </div>
                <div className="job-meta-row">
                  <label>Trạng thái</label>
                  <label className="toggle-wrap">
                    <input type="checkbox" checked={selectedJob.enabled !== false}
                      onChange={() => toggleJobEnabled(selectedJobIdx)} />
                    <span className={`toggle-label ${selectedJob.enabled !== false ? 'on' : 'off'}`}>
                      {selectedJob.enabled !== false ? '✓ Bật – sẽ tạo file' : '✗ Tắt – bỏ qua'}
                    </span>
                  </label>
                </div>
              </div>
              <FieldForm fields={selectedJob.fields || {}} catalog={catalog} onChange={setJobFields}
                onRequestEditLabel={handleRequestEditLabel}
                hint="Trường riêng của job này (ghi đè trường chung nếu trùng key)" />
            </>
          )}
        </div>
      </div>

      {/* ── Add Job Modal ── */}
      {showAddJob && (
        <div className="overlay" onClick={() => setShowAddJob(false)}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-hdr">
              <span>Thêm job từ template</span>
              <button className="btn-close" onClick={() => setShowAddJob(false)}>✕</button>
            </div>
            <div className="modal-body">
              {templates.length === 0 && <p style={{ color: '#9ca3af', padding: 12 }}>Chưa có template nào</p>}
              {templates.map(tpl => (
                <div key={tpl} className="tpl-item" onClick={() => addJob(tpl)}>
                  <span>📄</span>
                  <span>{tpl}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ── Generate Result Modal ── */}
      {genResult && (
        <div className="overlay" onClick={() => setGenResult(null)}>
          <div className="result-modal" onClick={e => e.stopPropagation()}>
            <div className={`result-icon ${genResult.errors?.length ? 'warn' : 'ok'}`}>
              {genResult.errors?.length ? '⚠️' : '✅'}
            </div>
            <h2 className="result-title">
              {genResult.errors?.length ? 'Hoàn thành có cảnh báo' : 'Tạo file thành công!'}
            </h2>
            <div className="result-stats">
              <span className="stat-ok">✓ {genResult.generated} file đã tạo</span>
              {genResult.skipped > 0 && <span className="stat-skip">⊘ {genResult.skipped} bỏ qua</span>}
            </div>
            <div className="result-path-block">
              <div className="result-path-label">Thư mục output:</div>
              <div className="result-path-val">{genResult.outDir}</div>
            </div>
            {genResult.errors?.length > 0 && (
              <div className="result-errors">
                {genResult.errors.map((e, i) => <div key={i} className="result-err-row">⚠ {e}</div>)}
              </div>
            )}
            <div className="result-actions">
              <button className="btn btn-primary"
                onClick={() => { window.api.openPath(genResult.outDir); setGenResult(null) }}>
                📂 Mở thư mục output
              </button>
              <button className="btn btn-secondary" onClick={() => setGenResult(null)}>Đóng</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Template Manager ── */}
      {showTemplateManager && (
        <TemplateManager
          catalog={catalog}
          templates={templates}
          onClose={() => setShowTemplateManager(false)}
          onCatalogUpdate={newCat => {
            setCatalog(newCat)
            // Inject any new catalog fields into existing job.fields
            setProject(p => ({
              ...p,
              global: { ...buildDefault(newCat).global, ...p.global },
              jobs: p.jobs.map(j => ({
                ...j,
                fields: { ...buildJobFields(j.template || '', newCat), ...j.fields },
              })),
            }))
          }}
          onTemplatesUpdate={newTpls => setTemplates(newTpls)}
        />
      )}

      {/* ── Toast ── */}
      {toast && (
        <div className={`toast toast-${toast.type}`} onClick={() => setToast(null)}>
          {toast.message}
        </div>
      )}
    </>
  )
}
