import { useState, useCallback, useMemo } from 'react'

function prettyFromKey(key) {
  return String(key || '')
    .replace(/^[A-Z_]+__/, '')
    .replace(/_/g, ' ')
    .trim()
}

function displayLabel(key, meta) {
  const label = String(meta?.label || '').trim()
  const sample = String(meta?.sample || '').trim()
  if (!label) return sample || prettyFromKey(key)
  // If label is literally the raw sample/value, show a friendlier synthesized label
  if (sample && label === sample) return prettyFromKey(key)
  return label
}

function buildTooltip(meta) {
  const parts = []
  if (meta.sample) parts.push(`Ví dụ: "${meta.sample}"`)
  if (meta.files?.length) {
    parts.push('Dùng trong:')
    meta.files.forEach(f => parts.push(`  • ${f}`))
  }
  return parts.length ? parts.join('\n') : null
}

export default function FieldForm({ fields, catalog, onChange, hint, onRequestEditLabel, hiddenKeys = [] }) {
  const [search, setSearch] = useState('')
  const [activeTooltip, setActiveTooltip] = useState(null)

  const hiddenSet = useMemo(() => new Set(hiddenKeys), [hiddenKeys])

  const orderedKeys = useMemo(() => {
    const templateOrder = [
      'Hợp đồng TC.docx',
      '1.BB ban giao mat bang.docx',
      '2.BB Ktra Điều kiện Thi công.docx',
      '3.BB Nghiệm thu nhân lực.docx',
      '4.BB ktra máy móc.docx',
      'Nghiệm thu vật liệu.docx',
      'BB Nghiệm thu công việc.docx',
      'Yêu cầu nghiệm thu công việc.docx',
      'BBNT khối lượng hoàn thành.docx',
      'BBBG đưa vào sử dụng.docx',
    ]

    const rank = (key) => {
      const meta = catalog[key] || {}
      const files = Array.isArray(meta.files) ? meta.files : []
      let best = 999
      for (const f of files) {
        const idx = templateOrder.indexOf(f)
        if (idx !== -1 && idx < best) best = idx
      }
      const specificity = files.length || 99
      const label = displayLabel(key, meta).toLowerCase()
      return [best, specificity, label]
    }

    return Object.keys(fields)
      .filter(k => !hiddenSet.has(k))
      .sort((a, b) => {
        const ra = rank(a)
        const rb = rank(b)
        if (ra[0] !== rb[0]) return ra[0] - rb[0]
        if (ra[1] !== rb[1]) return ra[1] - rb[1]
        return ra[2].localeCompare(rb[2], 'vi')
      })
  }, [fields, catalog, hiddenSet])

  const filtered = useMemo(() => {
    if (!search.trim()) return orderedKeys
    const q = search.toLowerCase()
    return orderedKeys.filter(k => {
      const meta = catalog[k] || {}
      const label = displayLabel(k, meta).toLowerCase()
      return k.toLowerCase().includes(q) || label.includes(q)
    })
  }, [orderedKeys, catalog, search])

  const handleChange = useCallback((key, value) => {
    onChange({ ...fields, [key]: value })
  }, [fields, onChange])

  return (
    <div className="field-form">
      {hint && <div className="field-form-hint">{hint}</div>}
      <div className="search-bar">
        <label>Tìm:</label>
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Tìm theo tên trường hoặc label..."
        />
        <span className="count">{filtered.length} / {orderedKeys.length} trường</span>
      </div>

      <div className="field-rows">
        {filtered.length === 0 && (
          <div className="empty-state">
            <h3>Không có trường nào</h3>
            <p>Thử xóa bộ lọc hoặc mở file JSON</p>
          </div>
        )}
        {filtered.map(key => {
          const meta = catalog[key] || {}
          const label = displayLabel(key, meta)
          const tip = buildTooltip(meta)
          const isOpen = activeTooltip === key
          return (
            <div className="field-row" key={key}>
              <div className="field-label-col">
                <div className="field-label-wrap">
                  <span className="field-label">{label}</span>
                  {tip && (
                    <span
                      className={`field-info-btn ${isOpen ? 'active' : ''}`}
                      onClick={() => setActiveTooltip(isOpen ? null : key)}
                      title="Xem thông tin trường"
                    >ⓘ</span>
                  )}
                  {onRequestEditLabel && (
                    <span
                      className="field-info-btn"
                      onClick={() => onRequestEditLabel({ key, currentLabel: meta.label || '', sample: meta.sample || '' })}
                      title="Sửa label (admin)"
                    >✏️</span>
                  )}
                </div>
                {isOpen && tip && (
                  <div className="field-tooltip">
                    {tip.split('\n').map((line, i) => (
                      <span key={i} className={line.startsWith('  •') ? 'tip-file' : ''}>
                        {line}
                      </span>
                    ))}
                  </div>
                )}
              </div>
              <div className="field-entry">
                <input
                  value={fields[key] ?? ''}
                  onChange={e => handleChange(key, e.target.value)}
                />
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}
