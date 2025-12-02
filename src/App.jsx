import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import './App.css'

const FUTURA_STACK = "'Futura', 'Futura PT', 'Century Gothic', 'Arial', sans-serif"
const DEFAULT_TEMPLATE_PATH = encodeURI('/Plantilla_4_personas.png')
const DEFAULT_TEMPLATE_PATH_FRONT = DEFAULT_TEMPLATE_PATH
const DEFAULT_TEMPLATE_PATH_BACK = DEFAULT_TEMPLATE_PATH
const COMPANY_FONT_BOOST = 1.12

const TEMPLATE_OPTIONS = [
  {
    id: 'sheet-letter',
    label: 'Plantilla general (tamaño carta)',
    front: DEFAULT_TEMPLATE_PATH_FRONT,
    back: DEFAULT_TEMPLATE_PATH_BACK,
    layout: 'sheet',
  },
  {
    id: 'custom',
    label: 'Usar mi propia plantilla',
    front: '',
    back: '',
    layout: 'sheet',
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

  const boostedCompanyFontSize = Math.max(10.4, Math.round(companyFontSize * COMPANY_FONT_BOOST * 10) / 10)

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

  return { nameFontSize, companyFontSize: boostedCompanyFontSize, namesGap, namesOffset, namesWidth }
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

function buildNameStyles(attendee, isBack, positionAdjustments, fontScale = 1, uniformMetrics) {
  const { fullName, company } = attendee
  const { nameFontSize, companyFontSize, namesGap, namesOffset, namesWidth } = getTypographyMetrics(fullName, company)

  const adjustedGap = namesGap + positionAdjustments.gap
  const adjustedOffset = namesOffset + positionAdjustments.vertical + (isBack ? 1.4 : 0)
  const adjustedWidth = Math.max(50, namesWidth + positionAdjustments.width - (isBack ? 2 : 0))
  const mirroredOffset = Math.max(12.5, adjustedOffset - 9.8)
  const safeScale = Math.min(Math.max(fontScale, 0.6), 1.6)
  const scaledNameSize = Math.round(nameFontSize * safeScale * 10) / 10
  const scaledCompanySize = Math.round(companyFontSize * safeScale * 10) / 10
  const uniformNameSize = uniformMetrics?.nameFontSize
  const uniformCompanySize = uniformMetrics?.companyFontSize
  const uniformNamesWidth = uniformMetrics?.namesWidth

  return {
    fontFamily: FUTURA_STACK,
    '--name-size': `${uniformNameSize ?? scaledNameSize}pt`,
    '--company-size': `${uniformCompanySize ?? scaledCompanySize}pt`,
    '--names-gap': `${adjustedGap}mm`,
    '--names-offset': `${adjustedOffset}mm`,
    '--names-offset-top': `${mirroredOffset}mm`,
    '--names-offset-bottom': `${mirroredOffset}mm`,
    '--names-width': `${uniformNamesWidth ?? adjustedWidth}mm`,
    '--names-horizontal': `${positionAdjustments.horizontal}mm`,
  }
}

function BadgeFace({
  attendees,
  variant = 'front',
  template,
  positionAdjustments,
  fontScale,
  hideTemplateImage = false,
  uniformMetrics,
}) {
  const isBack = variant === 'back'
  const templateSrc = hideTemplateImage ? '' : isBack ? template.back : template.front
  const [first] = attendees
  const baseScale = isBack ? fontScale.back : fontScale.front
  const primaryScale = baseScale * (isBack ? first.fontScaleBack ?? 1 : first.fontScaleFront ?? 1)
  const primaryStyles = buildNameStyles(first, isBack, positionAdjustments, primaryScale, uniformMetrics)

  return (
    <section className={`badge badge--${variant} ${templateSrc ? '' : 'badge--sheet-template'}`}>
      {templateSrc ? (
        <img className="badge__template" src={templateSrc} alt={`Plantilla ${variant}`} />
      ) : hideTemplateImage ? null : (
        <div className="badge__template badge__template--placeholder">Sube tu plantilla de {variant === 'front' ? 'frente' : 'reverso'}</div>
      )}

      <div className="names" style={primaryStyles}>
        <p className="name">{first.fullName || 'Nombre Apellido'}</p>
        <p className="company">{first.company || 'Empresa'}</p>
      </div>
    </section>
  )
}

function preloadImage(src) {
  return new Promise((resolve) => {
    if (!src) {
      resolve()
      return
    }

    const img = new Image()
    img.onload = () => resolve()
    img.onerror = () => resolve()
    img.src = src
  })
}

const EXPORT_SCALE = Math.max(3, window.devicePixelRatio * 2)

function chunkIntoSheets(groups, perSheet = 4) {
  const sheets = []
  for (let i = 0; i < groups.length; i += perSheet) {
    sheets.push(groups.slice(i, i + perSheet))
  }
  return sheets
}

function buildSlots(sheet, variant) {
  const order = variant === 'back' ? [1, 0, 3, 2] : [0, 1, 2, 3]
  const arranged = Array(4).fill(null)

  sheet.forEach((group, index) => {
    const targetIndex = order[index] ?? index
    arranged[targetIndex] = group
  })

  return arranged
}

function buildUniformMetrics(attendees) {
  if (!attendees.length) return null

  const metrics = attendees.map((person) => getTypographyMetrics(person.fullName, person.company))
  const nameFontSize = Math.min(...metrics.map((item) => item.nameFontSize))
  const companyFontSize = Math.min(...metrics.map((item) => item.companyFontSize))
  const namesWidth = Math.max(...metrics.map((item) => item.namesWidth))

  return {
    nameFontSize,
    companyFontSize,
    namesWidth,
  }
}

function PrintSheet({
  sheet,
  variant,
  template,
  positionAdjustments,
  fontScale,
  index,
  uniformMetrics,
}) {
  const useSheetTemplate = template.layout === 'sheet'
  const sheetTemplateSrc = variant === 'back' ? template.back : template.front
  const sheetBackgroundImage = sheetTemplateSrc ? `url("${sheetTemplateSrc}")` : ''
  const sheetStyle = useSheetTemplate && sheetBackgroundImage ? { backgroundImage: sheetBackgroundImage } : undefined
  const slots = buildSlots(sheet, variant)
  const hasContent = sheet.length > 0

  return (
    <section
      className={`print-sheet ${variant === 'back' ? 'print-sheet--back' : ''} ${useSheetTemplate ? 'print-sheet--full-template' : ''}`}
      style={sheetStyle}
      data-has-content={hasContent}
    >
      <p className="print-sheet__label">
        Hoja {index + 1} · {variant === 'front' ? 'Frente' : 'Reverso'}
      </p>
      {slots.map((group, slotIndex) => (
        <div className="print-slot" key={`${variant}-${index}-${slotIndex}`}>
          {group ? (
            <BadgeFace
              attendees={group}
              variant={variant}
              template={useSheetTemplate ? { ...template, front: '', back: '' } : template}
              positionAdjustments={positionAdjustments}
              fontScale={fontScale}
              uniformMetrics={uniformMetrics}
              hideTemplateImage={useSheetTemplate}
            />
          ) : (
            <div className="print-slot__placeholder" aria-hidden>
              Carga más nombres para completar esta hoja
            </div>
          )}
        </div>
      ))}
    </section>
  )
}

function App() {
  const [attendees, setAttendees] = useState([])
  const [error, setError] = useState('')
  const [templateId, setTemplateId] = useState(TEMPLATE_OPTIONS[0].id)
  const [customTemplate, setCustomTemplate] = useState({ front: '', back: '' })
  const [positionAdjustments, setPositionAdjustments] = useState({ vertical: 0, gap: 0, width: 0, horizontal: 0 })
  const [fontScale, setFontScale] = useState({ front: 1, back: 1 })
  const [attendeeOverrides, setAttendeeOverrides] = useState({})
  const [editingIndex, setEditingIndex] = useState(0)
  const [objectUrls, setObjectUrls] = useState([])
  const [isExporting, setIsExporting] = useState(false)
  const [quickSearch, setQuickSearch] = useState('')
  const [highlightedSuggestion, setHighlightedSuggestion] = useState(null)
  const [isPrinting, setIsPrinting] = useState(false)

  useEffect(
    () => () => {
      objectUrls.forEach((url) => URL.revokeObjectURL(url))
    },
    [objectUrls]
  )

  useEffect(() => {
    setEditingIndex((prev) => {
      if (!attendees.length) return 0
      if (prev >= attendees.length) return 0
      return prev
    })
  }, [attendees.length])

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
      setAttendeeOverrides({})
      setEditingIndex(0)
      setQuickSearch('')
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
    setAttendeeOverrides({})
    setEditingIndex(0)
    setQuickSearch('')
    setError('')
  }

  const updatePosition = (field, value) => {
    setPositionAdjustments((prev) => ({ ...prev, [field]: value }))
  }

  const updateFontScale = (side, value) => {
    setFontScale((prev) => ({ ...prev, [side]: value }))
  }

  const updateAttendeeOverride = (index, field, value) => {
    setAttendeeOverrides((prev) => ({
      ...prev,
      [index]: {
        ...(prev[index] || {}),
        [field]: value,
      },
    }))
  }

  const resetAttendeeOverride = (index) => {
    setAttendeeOverrides((prev) => {
      const { [index]: removed, ...rest } = prev
      return rest
    })
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

  const decoratedAttendees = useMemo(
    () =>
      attendees.map((attendee, index) => {
        const overrides = attendeeOverrides[index] || {}
        return {
          ...attendee,
          fullName: overrides.name ?? attendee.fullName,
          company: overrides.company ?? attendee.company,
          fontScaleFront: overrides.fontScaleFront ?? 1,
          fontScaleBack: overrides.fontScaleBack ?? 1,
        }
      }),
    [attendeeOverrides, attendees]
  )

  const suggestionCatalog = useMemo(
    () =>
      decoratedAttendees.map((person, index) => ({
        index,
        label: `${person.fullName || `Persona ${index + 1}`} · ${person.company || 'Sin empresa'}`,
      })),
    [decoratedAttendees]
  )

  const filteredSuggestions = useMemo(() => {
    if (!quickSearch.trim()) return suggestionCatalog.slice(0, 8)
    const normalized = quickSearch.toLowerCase()
    return suggestionCatalog
      .filter((item) => item.label.toLowerCase().includes(normalized))
      .slice(0, 8)
  }, [quickSearch, suggestionCatalog])

  const badgeGroups = useMemo(() => decoratedAttendees.map((attendee) => [attendee]), [decoratedAttendees])

  const sheets = useMemo(() => chunkIntoSheets(badgeGroups), [badgeGroups])
  const totalSheets = useMemo(() => sheets.length, [sheets])
  const totalPeople = useMemo(() => decoratedAttendees.length, [decoratedAttendees])
  const activeSheetIndex = useMemo(() => Math.floor(editingIndex / 4), [editingIndex])
  const activeSheet = sheets[activeSheetIndex] || []

  const editingPerson = decoratedAttendees[editingIndex] || null
  const editingPositionLabel = decoratedAttendees.length ? `#${editingIndex + 1} de ${decoratedAttendees.length}` : 'Sin selección'
  const isPersonCustomized = Boolean(attendeeOverrides[editingIndex])

  const activeBadgeGroup = useMemo(
    () => (decoratedAttendees.length ? [decoratedAttendees[editingIndex] || decoratedAttendees[0]] : []),
    [decoratedAttendees, editingIndex]
  )

  const uniformMetrics = useMemo(() => buildUniformMetrics(decoratedAttendees), [decoratedAttendees])

  const handleDownloadPDF = async () => {
    if (!badgeGroups.length || missingCustomTemplate) return
    setIsExporting(true)
    setError('')

    try {
      await new Promise((resolve) => requestAnimationFrame(() => resolve()))
      await Promise.all([preloadImage(activeTemplate.front), preloadImage(activeTemplate.back)])

      const sheetsToExport = Array.from(document.querySelectorAll('.print-sheet')).filter(
        (sheet) => sheet.dataset.hasContent === 'true'
      )
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' })
      const pageWidth = pdf.internal.pageSize.getWidth()
      const pageHeight = pdf.internal.pageSize.getHeight()

      for (const [index, sheet] of sheetsToExport.entries()) {
        const { width, height } = sheet.getBoundingClientRect()
        const scaledWidth = Math.round(width)
        const scaledHeight = Math.round(height)

        // eslint-disable-next-line no-await-in-loop
        const canvas = await html2canvas(sheet, {
          scale: EXPORT_SCALE,
          useCORS: true,
          backgroundColor: '#fff',
          width: scaledWidth,
          height: scaledHeight,
          windowWidth: scaledWidth,
          windowHeight: scaledHeight,
          scrollX: 0,
          scrollY: 0,
        })
        const imgData = canvas.toDataURL('image/png', 1)

        if (index > 0) pdf.addPage()
        pdf.addImage(imgData, 'PNG', 0, 0, pageWidth, pageHeight)
      }

      pdf.save('gafetes.pdf')
    } catch (err) {
      console.error(err)
      setError('No pudimos generar el PDF. Vuelve a intentarlo o revisa que la plantilla sea una imagen.')
    } finally {
      setIsExporting(false)
    }
  }

  const handlePrint = () => {
    if (!badgeGroups.length || missingCustomTemplate) return
    setIsPrinting(true)
    setTimeout(() => {
      window.print()
      setIsPrinting(false)
    }, 50)
  }

  const handleQuickSelect = (value) => {
    setQuickSearch(value)
    setHighlightedSuggestion(null)
    if (!value.trim()) return

    const normalized = value.toLowerCase()
    const exactIndex = decoratedAttendees.findIndex((person, index) => {
      const label = `${person.fullName || `Persona ${index + 1}`} · ${person.company || 'Sin empresa'}`
      return label.toLowerCase() === normalized
    })
    if (exactIndex >= 0) {
      setEditingIndex(exactIndex)
      return
    }

    const fuzzyIndex = decoratedAttendees.findIndex(
      (person) =>
        (person.fullName || '').toLowerCase().includes(normalized) ||
        (person.company || '').toLowerCase().includes(normalized)
    )
    if (fuzzyIndex >= 0) {
      setEditingIndex(fuzzyIndex)
    }
  }

  const handleSuggestionPick = (index, label) => {
    setHighlightedSuggestion(label)
    setQuickSearch(label)
    setEditingIndex(index)
  }

  return (
    <div className={`page ${isPrinting ? 'page--printing' : ''} ${isExporting ? 'page--exporting' : ''}`}>
      <header className="hero">
        <div>
          <p className="eyebrow">Plantilla carta · 4 gafetes por hoja · Impresión a doble cara</p>
          <h1>Generador de gafetes</h1>
          <p className="lead">
            Elige la plantilla carta de cuatro espacios o carga tu diseño, sube un Excel con las columnas <strong>Empresa</strong>,{' '}
            <strong>Apellido</strong> y <strong>Nombre</strong> y personaliza la posición del texto. La hoja carta ya está dividida en
            cuatro zonas listas para imprimir frente y reverso sin ajustes adicionales.
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
          <p className="panel__title">Tamaño de letra</p>
          <div className="controls controls--inline">
            <label className="control">
              <span>Hoja 1 / Frente</span>
              <input
                type="range"
                min="0.6"
                max="1.6"
                step="0.05"
                value={fontScale.front}
                onChange={(event) => updateFontScale('front', Number(event.target.value))}
              />
              <span className="control__value">{Math.round(fontScale.front * 100)}%</span>
            </label>
            <label className="control">
              <span>Hoja 2 / Reverso</span>
              <input
                type="range"
                min="0.6"
                max="1.6"
                step="0.05"
                value={fontScale.back}
                onChange={(event) => updateFontScale('back', Number(event.target.value))}
              />
              <span className="control__value">{Math.round(fontScale.back * 100)}%</span>
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

      <section className="panel panel--stacked">
        <div className="panel__heading">
          <div>
            <p className="panel__title">Edición individual</p>
            <p className="helper">
              Ajusta un nombre o empresa de manera puntual y controla el tamaño de letra de cada hoja (frente y reverso)
              sin afectar a los demás. Recorre la lista, aplica cambios rápidos y valida en la vista previa inmediata de
              cada lado.
            </p>
          </div>
          <div className="controls controls--inline">
            <label className="control">
              <span>Selecciona una persona</span>
              <select
                value={decoratedAttendees.length ? editingIndex : ''}
                onChange={(event) => setEditingIndex(Number(event.target.value))}
                disabled={!decoratedAttendees.length}
              >
                {!decoratedAttendees.length && <option value="">Sube tu Excel para editar</option>}
                {decoratedAttendees.map((person, index) => (
                  <option key={`${person.fullName}-${index}`} value={index}>
                    {person.fullName || `Persona ${index + 1}`}
                  </option>
                ))}
              </select>
            </label>
            <label className="control control--search">
              <span>Buscar y editar rápido</span>
              <input
                type="search"
                list="people-suggestions"
                value={quickSearch}
                onChange={(event) => handleQuickSelect(event.target.value)}
                placeholder="Escribe un nombre o empresa"
                disabled={!decoratedAttendees.length}
              />
              <datalist id="people-suggestions">
                {filteredSuggestions.map((item) => (
                  <option key={item.label} value={item.label} />
                ))}
              </datalist>
              <span className="control__hint">Autocompleta y salta directo a la persona que necesitas.</span>
              {Boolean(filteredSuggestions.length) && quickSearch.trim() && (
                <div className="suggestion-grid" role="listbox" aria-label="Coincidencias rápidas">
                  {filteredSuggestions.map((item) => (
                    <button
                      type="button"
                      key={item.label}
                      className={`suggestion ${highlightedSuggestion === item.label ? 'suggestion--active' : ''}`}
                      onClick={() => handleSuggestionPick(item.index, item.label)}
                    >
                      <span className="suggestion__name">{item.label}</span>
                      <span className="pill pill--mini">Ir a #{item.index + 1}</span>
                    </button>
                  ))}
                </div>
              )}
            </label>
            <div className="inline-actions">
              <button
                type="button"
                className="ghost"
                disabled={!decoratedAttendees.length}
                onClick={() =>
                  setEditingIndex((prev) =>
                    decoratedAttendees.length ? (prev - 1 + decoratedAttendees.length) % decoratedAttendees.length : prev
                  )
                }
              >
                ← Anterior
              </button>
              <button
                type="button"
                className="ghost"
                disabled={!decoratedAttendees.length}
                onClick={() =>
                  setEditingIndex((prev) =>
                    decoratedAttendees.length ? (prev + 1) % decoratedAttendees.length : prev
                  )
                }
              >
                Siguiente →
              </button>
            </div>
            <button
              type="button"
              className="ghost"
              disabled={!decoratedAttendees.length}
              onClick={() => resetAttendeeOverride(editingIndex)}
            >
              Restablecer persona
            </button>
          </div>
        </div>

        {decoratedAttendees.length > 0 ? (
          <>
            <div className="person-rail" role="list">
              {decoratedAttendees.map((person, index) => {
                const isActive = index === editingIndex
                const hasChanges = Boolean(attendeeOverrides[index])
                return (
                  <button
                    type="button"
                    key={`${person.fullName}-${index}`}
                    className={`person-chip ${isActive ? 'person-chip--active' : ''} ${hasChanges ? 'person-chip--dirty' : ''}`}
                    onClick={() => setEditingIndex(index)}
                    aria-current={isActive}
                    role="listitem"
                  >
                    <span className="person-chip__name">{person.fullName || `Persona ${index + 1}`}</span>
                    <span className="person-chip__meta">{person.company || 'Sin empresa'}</span>
                    {hasChanges && <span className="pill pill--mini">Ajustado</span>}
                  </button>
                )
              })}
            </div>

            <div className="individual-editor">
              <div className="editor-grid">
                <div className="control control--summary">
                  <span className="label">Editando</span>
                  <strong>{editingPerson?.fullName || 'Selecciona una persona'}</strong>
                  <div className="pill pill--neutral">{editingPositionLabel}</div>
                  {isPersonCustomized ? <p className="helper">Esta persona tiene ajustes únicos.</p> : <p className="helper">Los valores se basan en el Excel original.</p>}
                </div>
                <label className="control">
                  <span>Nombre a mostrar</span>
                  <input
                    type="text"
                    value={attendeeOverrides[editingIndex]?.name ?? attendees[editingIndex]?.fullName ?? ''}
                    onChange={(event) => updateAttendeeOverride(editingIndex, 'name', event.target.value)}
                    placeholder="Nombre y apellidos"
                  />
                  <span className="control__hint">Ideal para corregir tildes o apellidos compuestos.</span>
                </label>
                <label className="control">
                  <span>Empresa a mostrar</span>
                  <input
                    type="text"
                    value={attendeeOverrides[editingIndex]?.company ?? attendees[editingIndex]?.company ?? ''}
                    onChange={(event) => updateAttendeeOverride(editingIndex, 'company', event.target.value)}
                    placeholder="Nombre de la empresa"
                  />
                  <span className="control__hint">Se adapta automáticamente a textos largos.</span>
                </label>
                <label className="control">
                  <span>Tamaño solo para hoja 1 (frente)</span>
                  <input
                    type="range"
                    min="0.6"
                    max="1.6"
                    step="0.05"
                    value={attendeeOverrides[editingIndex]?.fontScaleFront ?? 1}
                    onChange={(event) => updateAttendeeOverride(editingIndex, 'fontScaleFront', Number(event.target.value))}
                  />
                  <span className="control__value">{Math.round((attendeeOverrides[editingIndex]?.fontScaleFront ?? 1) * 100)}%</span>
                  <span className="control__hint">Usa esto cuando el nombre en el frente necesite más aire.</span>
                </label>
                <label className="control">
                  <span>Tamaño solo para hoja 2 (reverso)</span>
                  <input
                    type="range"
                    min="0.6"
                    max="1.6"
                    step="0.05"
                    value={attendeeOverrides[editingIndex]?.fontScaleBack ?? 1}
                    onChange={(event) => updateAttendeeOverride(editingIndex, 'fontScaleBack', Number(event.target.value))}
                  />
                  <span className="control__value">{Math.round((attendeeOverrides[editingIndex]?.fontScaleBack ?? 1) * 100)}%</span>
                  <span className="control__hint">Ajusta aquí si la cara posterior se ve más cargada.</span>
                </label>
              </div>

              <div className="person-preview">
                <div className="person-preview__header">
                  <p className="person-preview__label">
                    Vista previa en vivo de <strong>{editingPerson?.fullName || 'la persona seleccionada'}</strong>
                  </p>
                  <div className="pill pill--neutral">Plantilla carta · 4 gafetes</div>
                </div>
                <p className="helper">Observa cómo se verá cada lado sin salir de la edición individual.</p>
              <div className="badge-pair badge-pair--compact">
                  <div className="badge-preview">
                    <p className="badge-preview__label">Hoja activa · Frente</p>
                    <PrintSheet
                      sheet={activeSheet}
                      variant="front"
                      template={activeTemplate}
                      positionAdjustments={positionAdjustments}
                      fontScale={fontScale}
                      index={activeSheetIndex}
                      uniformMetrics={uniformMetrics}
                    />
                  </div>
                  <div className="badge-preview">
                    <p className="badge-preview__label">Hoja activa · Reverso</p>
                    <PrintSheet
                      sheet={activeSheet}
                      variant="back"
                      template={activeTemplate}
                      positionAdjustments={positionAdjustments}
                      fontScale={fontScale}
                      index={activeSheetIndex}
                      uniformMetrics={uniformMetrics}
                    />
                  </div>
                </div>
              </div>
            </div>
          </>
        ) : (
          <p className="empty">Carga nombres para habilitar la edición individual.</p>
        )}
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
            Ajusta los deslizadores hasta que el texto caiga en el lugar exacto de tu plantilla. Cada hoja carta acomoda
            4 gafetes simétricos listos para imprimir frente y reverso con la misma orientación.
          </p>
        </div>
        {!attendees.length && <p className="empty">Sube tu Excel o usa el ejemplo para comenzar.</p>}
        {missingCustomTemplate && <p className="empty">Sube ambos lados de la plantilla personalizada para generar la vista previa.</p>}
        <div className="sheet-grid">
          {sheets.map((sheet, index) => (
            <div className="sheet-pair" key={`sheet-${index}`}>
              <PrintSheet
                sheet={sheet}
                variant="front"
                template={activeTemplate}
                positionAdjustments={positionAdjustments}
                fontScale={fontScale}
                index={index}
                uniformMetrics={uniformMetrics}
              />
              <PrintSheet
                sheet={sheet}
                variant="back"
                template={activeTemplate}
                positionAdjustments={positionAdjustments}
                fontScale={fontScale}
                index={index}
                uniformMetrics={uniformMetrics}
              />
            </div>
          ))}
        </div>
      </section>
    </div>
  )
}

export default App
