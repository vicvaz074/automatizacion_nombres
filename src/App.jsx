import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import './App.css'

const FUTURA_STACK = "'Futura', 'Futura PT', 'Century Gothic', 'Arial', sans-serif"

const TEMPLATE_OPTIONS = [
  {
    id: 'png-default',
    label: 'Plantilla base (PNG)',
    front: '/Plantilla_hoja_1.png',
    back: '/Plantilla_hoja_2.png',
  },
  {
    id: 'svg-alt',
    label: 'Plantilla alternativa (SVG)',
    front: '/template-front.svg',
    back: '/template-back.svg',
  },
  {
    id: 'custom',
    label: 'Usar mi propia plantilla',
    front: '',
    back: '',
  },
]

const demoRows = [
  { company: 'davara Abogados', lastName: 'Rangel', firstName: 'María' },
  { company: 'Tsuru', lastName: 'Aguayo', firstName: 'Diego' },
  { company: 'OpenAI', lastName: 'Coder', firstName: 'GPT' },
  { company: 'Tech Partners', lastName: 'López', firstName: 'Camila' },
]

function normalizeValue(value) {
  if (value === undefined || value === null) return ''
  return String(value).trim()
}

function calculateFontSize(text, { baseSize, minSize, maxChars }) {
  const length = normalizeValue(text).length
  if (!length) return baseSize
  if (length <= maxChars) return baseSize

  const scaled = (maxChars / length) * baseSize
  return Math.max(minSize, Math.round(scaled * 10) / 10)
}

function estimateLines(text, idealLineLength = 14) {
  const words = normalizeValue(text).split(/\s+/).filter(Boolean)
  if (!words.length) return 1

  const target = Math.max(10, idealLineLength)
  let lines = 1
  let current = 0

  for (const word of words) {
    const len = word.length
    if (current === 0) {
      current = len
      continue
    }

    if (current + 1 + len > target) {
      lines += 1
      current = len
    } else {
      current += 1 + len
    }
  }

  return Math.max(lines, Math.ceil(normalizeValue(text).length / (target + 2)))
}

function getTypographyMetrics(fullName, company) {
  const normalizedName = normalizeValue(fullName) || 'Nombre Apellido'
  const normalizedCompany = normalizeValue(company) || 'Empresa'

  const nameWords = normalizedName.split(/\s+/).filter(Boolean)
  const wordCount = nameWords.length || 1
  const longestWord = nameWords.reduce((max, word) => Math.max(max, word.length), 0)
  const companyWords = normalizedCompany.split(/\s+/).filter(Boolean)
  const companyLongestWord = companyWords.reduce((max, word) => Math.max(max, word.length), 0)

  const baseNameSize = calculateFontSize(normalizedName, { baseSize: 27, minSize: 17, maxChars: 18 })
  const baseCompanySize = calculateFontSize(normalizedCompany, { baseSize: 16.6, minSize: 10.4, maxChars: 18 })
  const roomyCompanySize = calculateFontSize(normalizedCompany, { baseSize: 15.4, minSize: 10.8, maxChars: 24 })

  const density = Math.max(
    normalizedName.length / 18,
    normalizedCompany.length / 22,
    (normalizedName.length + normalizedCompany.length) / 40
  )

  const scale = density > 1 ? Math.max(0.72, 1 / density) : 1

  const estimatedNameLines = estimateLines(normalizedName, longestWord >= 12 ? 12 : 14)
  const estimatedCompanyLines = estimateLines(normalizedCompany, Math.min(20, Math.max(16, 26 - companyLongestWord)))

  const crowdedNameScale = wordCount >= 3 ? 0.88 : longestWord >= 12 ? 0.93 : 1
  const extraWordsScale = wordCount >= 6 ? 0.78 : wordCount === 5 ? 0.85 : 1
  const longWordScale = longestWord >= 14 ? 0.9 : 1
  const multilineNameScale = estimatedNameLines >= 2 ? 0.92 - Math.min(0.08, (estimatedNameLines - 2) * 0.04) : 1
  const nameFontSize = Math.max(
    15,
    Math.round(baseNameSize * scale * crowdedNameScale * multilineNameScale * extraWordsScale * longWordScale * 10) / 10
  )

  const companyDensity = normalizedCompany.length / 24
  const companyScale = companyDensity > 1 ? Math.max(0.56, 1 - (companyDensity - 1) * 0.28) : 1 + (1 - companyDensity) * 0.08
  const balancedCompanySize = Math.round(nameFontSize * 0.8 * 10) / 10
  const companyCrowdingScale = wordCount >= 3 || estimatedNameLines >= 2 ? 0.9 : 1
  const companyLongWordScale =
    companyLongestWord >= 18 ? 0.82 : companyLongestWord >= 14 ? 0.88 : companyLongestWord >= 12 ? 0.93 : 1
  const companyLengthScale =
    normalizedCompany.length >= 42
      ? 0.72
      : normalizedCompany.length >= 34
        ? 0.8
        : normalizedCompany.length >= 28
          ? 0.88
          : 1
  const companyMultilineScale =
    estimatedCompanyLines >= 2 ? 0.9 - Math.min(0.18, (estimatedCompanyLines - 2) * 0.06) : 1
  const shortCompanyBoost = normalizedCompany.length <= 8 && estimatedCompanyLines === 1 ? 1.16 : normalizedCompany.length <= 12 ? 1.08 : 1
  const tinyWordBoost = companyLongestWord <= 6 && normalizedCompany.length <= 10 ? 1.1 : 1
  const nameToCompanyBalance = Math.min(nameFontSize * 0.84, Math.max(baseCompanySize * 0.96, balancedCompanySize))
  const companyFontSize = Math.max(
    10.4,
    Math.round(
      baseCompanySize *
        scale *
        companyCrowdingScale *
        companyLongWordScale *
        companyLengthScale *
        companyMultilineScale *
        shortCompanyBoost *
        tinyWordBoost *
        10
    ) / 10,
    Math.round(
      roomyCompanySize *
        companyCrowdingScale *
        companyLongWordScale *
        companyLengthScale *
        companyMultilineScale *
        shortCompanyBoost *
        tinyWordBoost *
        10
    ) / 10,
    nameToCompanyBalance
  )

  const baseGap = 3.2
  const multilineGapBoost = estimatedNameLines >= 3 ? 1.18 : estimatedNameLines === 2 ? 1.12 : 1
  const companyLinesBoost = estimatedCompanyLines >= 2 ? 1.12 + Math.min(0.18, (estimatedCompanyLines - 2) * 0.05) : 1
  const companyLengthGapBoost = normalizedCompany.length >= 26 ? 1.14 : normalizedCompany.length >= 18 ? 1.08 : 1
  const shortCompanyRelax = normalizedCompany.length <= 10 && estimatedCompanyLines === 1 ? 0.94 : 1
  const namesGap = Math.max(baseGap * multilineGapBoost * companyLinesBoost * companyLengthGapBoost * shortCompanyRelax, 6.1)

  const baseOffset = 27.5
  const multilineOffsetBoost = estimatedNameLines > 1 ? Math.min(5.4, (estimatedNameLines - 1) * 2.35) : 0
  const companyOffsetBoost = estimatedCompanyLines > 1 ? Math.min(3.6, (estimatedCompanyLines - 1) * 1.35) : 0
  const namesOffset = Math.max(21.5, baseOffset - multilineOffsetBoost - companyOffsetBoost)

  const widthPenalty = Math.max(
    0,
    (normalizedCompany.length - 16) * 0.24,
    (companyLongestWord - 10) * 0.72,
    estimatedCompanyLines >= 2 ? 1.6 : 0
  )
  const namesWidth = Math.max(64, Math.round((74 - widthPenalty) * 10) / 10)

  return { nameFontSize, companyFontSize, namesGap, namesOffset, namesWidth }
}

function buildAttendees(rows) {
  return rows
    .map((row) => {
      const company = normalizeValue(row['Empresa'])
      const firstName = normalizeValue(row['Nombre'])
      const lastName = normalizeValue(row['Apellido'])
      if (!company && !firstName && !lastName) return null
      return {
        company,
        firstName,
        lastName,
        fullName: `${firstName}${firstName && lastName ? ' ' : ''}${lastName}`.trim(),
      }
    })
    .filter(Boolean)
}

function buildNameStyles(attendee, isBack, positionAdjustments) {
  const { fullName, company } = attendee
  const { nameFontSize, companyFontSize, namesGap, namesOffset, namesWidth } = getTypographyMetrics(fullName, company)

  const adjustedGap = namesGap + positionAdjustments.gap
  const adjustedOffset = namesOffset + positionAdjustments.vertical + (isBack ? 1.4 : 0)
  const adjustedWidth = Math.max(50, namesWidth + positionAdjustments.width - (isBack ? 2 : 0))
  const topOffset = Math.max(12.5, adjustedOffset - 8.2)
  const bottomOffset = Math.max(10, adjustedOffset - 12.2)

  return {
    fontFamily: FUTURA_STACK,
    '--name-size': `${nameFontSize}pt`,
    '--company-size': `${companyFontSize}pt`,
    '--names-gap': `${adjustedGap}mm`,
    '--names-offset': `${adjustedOffset}mm`,
    '--names-offset-top': `${topOffset}mm`,
    '--names-offset-bottom': `${bottomOffset}mm`,
    '--names-width': `${adjustedWidth}mm`,
    '--names-horizontal': `${positionAdjustments.horizontal}mm`,
  }
}

function BadgeFace({ attendees, variant = 'front', template, layoutMode, positionAdjustments }) {
  const isBack = variant === 'back'
  const templateSrc = isBack ? template.back : template.front
  const [first, second] = attendees
  const primaryStyles = buildNameStyles(first, isBack, positionAdjustments)
  const secondaryStyles = second ? buildNameStyles(second, isBack, positionAdjustments) : primaryStyles
  const isMirrorLayout = layoutMode === 'mirror'

  return (
    <section className={`badge badge--${variant}`}>
      {templateSrc ? (
        <img className="badge__template" src={templateSrc} alt={`Plantilla ${variant}`} />
      ) : (
        <div className="badge__template badge__template--placeholder">Sube tu plantilla de {variant === 'front' ? 'frente' : 'reverso'}</div>
      )}

      {isMirrorLayout ? (
        <>
          <div className="names" style={primaryStyles}>
            <p className="name">{first.fullName || 'Nombre Apellido'}</p>
            <p className="company">{first.company || 'Empresa'}</p>
          </div>

          <div className="names names--mirrored" style={primaryStyles}>
            <p className="name">{first.fullName || 'Nombre Apellido'}</p>
            <p className="company">{first.company || 'Empresa'}</p>
          </div>
        </>
      ) : (
        <>
          <div className="names names--top" style={primaryStyles}>
            <p className="name">{first.fullName || 'Nombre Apellido'}</p>
            <p className="company">{first.company || 'Empresa'}</p>
          </div>

          {second && (
            <div className="names names--bottom" style={secondaryStyles}>
              <p className="name">{second.fullName || 'Nombre Apellido'}</p>
              <p className="company">{second.company || 'Empresa'}</p>
            </div>
          )}
        </>
      )}
    </section>
  )
}

function App() {
  const [attendees, setAttendees] = useState([])
  const [error, setError] = useState('')
  const [templateId, setTemplateId] = useState(TEMPLATE_OPTIONS[0].id)
  const [customTemplate, setCustomTemplate] = useState({ front: '', back: '' })
  const [layoutMode, setLayoutMode] = useState('mirror')
  const [positionAdjustments, setPositionAdjustments] = useState({ vertical: 0, gap: 0, width: 0, horizontal: 0 })
  const [objectUrls, setObjectUrls] = useState([])
  const [isExporting, setIsExporting] = useState(false)

  useEffect(
    () => () => {
      objectUrls.forEach((url) => URL.revokeObjectURL(url))
    },
    [objectUrls]
  )

  const handleFile = async (event) => {
    setError('')
    const [file] = event.target.files || []
    if (!file) return

    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' })
      const parsed = buildAttendees(rows)
      if (!parsed.length) {
        setError('No se encontraron filas con las columnas "Empresa", "Apellido" y "Nombre".')
      }
      setAttendees(parsed)
    } catch (err) {
      console.error(err)
      setError('No pudimos leer el archivo. Asegúrate de subir un Excel (.xlsx) con las columnas esperadas.')
    }
  }

  const handleTemplateUpload = (side, fileList) => {
    const [file] = fileList || []
    if (!file) return
    const url = URL.createObjectURL(file)
    setObjectUrls((prev) => [...prev, url])
    setCustomTemplate((prev) => ({ ...prev, [side]: url }))
    setTemplateId('custom')
  }

  const loadDemo = () => {
    const sampleRows = demoRows.map((row) => ({
      Empresa: row.company,
      Apellido: row.lastName,
      Nombre: row.firstName,
    }))
    setAttendees(buildAttendees(sampleRows))
    setError('')
  }

  const updatePosition = (field, value) => {
    setPositionAdjustments((prev) => ({ ...prev, [field]: value }))
  }

  const activeTemplate = useMemo(() => {
    const selected = TEMPLATE_OPTIONS.find((option) => option.id === templateId) || TEMPLATE_OPTIONS[0]
    if (selected.id !== 'custom') return selected
    return { ...selected, ...customTemplate }
  }, [customTemplate, templateId])

  const missingCustomTemplate = useMemo(
    () => activeTemplate.id === 'custom' && (!activeTemplate.front || !activeTemplate.back),
    [activeTemplate]
  )

  const badgeGroups = useMemo(() => {
    if (layoutMode === 'mirror') return attendees.map((attendee) => [attendee])

    const groups = []
    for (let i = 0; i < attendees.length; i += 2) {
      groups.push(attendees.slice(i, i + 2))
    }
    return groups
  }, [attendees, layoutMode])

  const totalSheets = useMemo(() => badgeGroups.length, [badgeGroups])
  const totalPeople = useMemo(() => attendees.length, [attendees])

  const handleDownloadPDF = async () => {
    if (!badgeGroups.length || missingCustomTemplate) return
    setIsExporting(true)

    const badges = Array.from(document.querySelectorAll('.badge'))
    const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [100, 100] })

    for (const [index, badge] of badges.entries()) {
      // eslint-disable-next-line no-await-in-loop
      const canvas = await html2canvas(badge, { scale: 2, useCORS: true, backgroundColor: '#fff' })
      const imgData = canvas.toDataURL('image/png')
      if (index > 0) pdf.addPage()
      pdf.addImage(imgData, 'PNG', 0, 0, 100, 100)
    }

    pdf.save('gafetes.pdf')
    setIsExporting(false)
  }

  const handlePrint = () => {
    if (!badgeGroups.length || missingCustomTemplate) return
    window.print()
  }

  return (
    <div className="page">
      <header className="hero">
        <div>
          <p className="eyebrow">Plantilla 10cm x 10cm · Impresión a doble cara</p>
          <h1>Generador de gafetes</h1>
          <p className="lead">
            Elige una plantilla, sube un Excel con las columnas <strong>Empresa</strong>, <strong>Apellido</strong> y{' '}
            <strong>Nombre</strong> y personaliza la posición del texto. Puedes duplicar el mismo nombre en espejo o
            imprimir dos nombres por hoja listos para doble cara.
          </p>
        </div>
        <div className="actions">
          <label className="upload">
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} />
            <span>Subir Excel</span>
          </label>
          <button type="button" className="ghost" onClick={loadDemo}>
            Cargar ejemplo
          </button>
          <div className="inline-actions">
            <button type="button" className="primary" onClick={handlePrint} disabled={!attendees.length || missingCustomTemplate}>
              Imprimir (frente y reverso)
            </button>
            <button type="button" onClick={handleDownloadPDF} disabled={!attendees.length || missingCustomTemplate || isExporting}>
              {isExporting ? 'Generando PDF…' : 'Descargar PDF'}
            </button>
          </div>
          <p className="helper">Consejo: imprime en doble cara con unión por borde largo.</p>
        </div>
      </header>

      <section className="panel">
        <div>
          <p className="panel__title">Plantillas</p>
          <div className="controls">
            <label className="control">
              <span>Selecciona plantilla</span>
              <select value={templateId} onChange={(event) => setTemplateId(event.target.value)}>
                {TEMPLATE_OPTIONS.map((option) => (
                  <option key={option.id} value={option.id}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>

            {templateId === 'custom' && (
              <div className="control control--inline">
                <label>
                  <span>Frente</span>
                  <input type="file" accept="image/*" onChange={(event) => handleTemplateUpload('front', event.target.files)} />
                </label>
                <label>
                  <span>Reverso</span>
                  <input type="file" accept="image/*" onChange={(event) => handleTemplateUpload('back', event.target.files)} />
                </label>
              </div>
            )}

            {missingCustomTemplate && <p className="helper warning">Sube frente y reverso para usar tu plantilla personalizada.</p>}
          </div>
        </div>

        <div>
          <p className="panel__title">Modo de distribución</p>
          <div className="controls controls--inline">
            <label className="pill-option">
              <input
                type="radio"
                name="layout"
                value="mirror"
                checked={layoutMode === 'mirror'}
                onChange={(event) => setLayoutMode(event.target.value)}
              />
              <span>Nombre en espejo (1 por hoja)</span>
            </label>
            <label className="pill-option">
              <input
                type="radio"
                name="layout"
                value="paired"
                checked={layoutMode === 'paired'}
                onChange={(event) => setLayoutMode(event.target.value)}
              />
              <span>Dos nombres por hoja</span>
            </label>
          </div>
        </div>

        <div>
          <p className="panel__title">Ajuste fino de texto</p>
          <div className="controls grid">
            <label className="control">
              <span>Desplazamiento vertical</span>
              <input
                type="range"
                min="-10"
                max="10"
                step="0.5"
                value={positionAdjustments.vertical}
                onChange={(event) => updatePosition('vertical', Number(event.target.value))}
              />
              <span className="control__value">{positionAdjustments.vertical} mm</span>
            </label>
            <label className="control">
              <span>Separación nombre/empresa</span>
              <input
                type="range"
                min="-5"
                max="10"
                step="0.5"
                value={positionAdjustments.gap}
                onChange={(event) => updatePosition('gap', Number(event.target.value))}
              />
              <span className="control__value">{positionAdjustments.gap} mm</span>
            </label>
            <label className="control">
              <span>Ancho del bloque de texto</span>
              <input
                type="range"
                min="-10"
                max="10"
                step="0.5"
                value={positionAdjustments.width}
                onChange={(event) => updatePosition('width', Number(event.target.value))}
              />
              <span className="control__value">{positionAdjustments.width} mm</span>
            </label>
            <label className="control">
              <span>Desplazamiento horizontal</span>
              <input
                type="range"
                min="-10"
                max="10"
                step="0.5"
                value={positionAdjustments.horizontal}
                onChange={(event) => updatePosition('horizontal', Number(event.target.value))}
              />
              <span className="control__value">{positionAdjustments.horizontal} mm</span>
            </label>
          </div>
        </div>
      </section>

      {error && <div className="alert">{error}</div>}

      <section className="status">
        <div>
          <p className="label">Personas listas</p>
          <strong className="stat">{totalPeople}</strong>
        </div>
        <div>
          <p className="label">Hojas generadas</p>
          <strong className="stat">{totalSheets}</strong>
        </div>
        <div>
          <p className="label">Columnas esperadas</p>
          <p className="pill">Empresa · Apellido · Nombre</p>
        </div>
      </section>

      <section className="preview">
        <div className="preview-info">
          <h2>Vista previa</h2>
          <p>
            Ajusta los deslizadores hasta que el texto caiga en el lugar exacto de tu plantilla. En modo espejo la parte
            inferior rota 180° para que coincida al imprimir y doblar; en modo doble se incluyen dos nombres por hoja y
            la siguiente página repite el par para que imprimas a doble cara.
          </p>
        </div>
        {!attendees.length && <p className="empty">Sube tu Excel o usa el ejemplo para comenzar.</p>}
        {missingCustomTemplate && <p className="empty">Sube ambos lados de la plantilla personalizada para generar la vista previa.</p>}
        <div className="badge-grid">
          {badgeGroups.map((group, index) => (
            <div className="badge-pair" key={`${group.map((person) => person.fullName).join('-')}-${index}`}>
              <BadgeFace
                attendees={group}
                variant="front"
                template={activeTemplate}
                layoutMode={layoutMode}
                positionAdjustments={positionAdjustments}
              />
              <BadgeFace
                attendees={group}
                variant="back"
                template={activeTemplate}
                layoutMode={layoutMode}
                positionAdjustments={positionAdjustments}
              />
            </div>
          ))}
        </div>
      </section>
    </div>
  )
}

export default App
