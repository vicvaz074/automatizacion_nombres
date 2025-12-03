import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import './App.css'

const FUTURA_STACK = "'Futura', 'Futura PT', 'Century Gothic', 'Arial', sans-serif"
const DEFAULT_TEMPLATE_PATH = encodeURI('/Plantilla_4_personas.png')
const JORNADA_TEMPLATE_PATH = encodeURI('/Gafetes_jornada.png')
const DEFAULT_TEMPLATE_PATH_FRONT = DEFAULT_TEMPLATE_PATH
const DEFAULT_TEMPLATE_PATH_BACK = DEFAULT_TEMPLATE_PATH
const COMPANY_FONT_BOOST = 1.12

const MODES = {
  PERSONIFICADORES: 'personificadores',
  JORNADA: 'gafetes-jornada',
}

const TEMPLATE_OPTIONS = {
  [MODES.PERSONIFICADORES]: [
    {
      id: 'sheet-letter',
      label: 'Plantilla general (tamaño carta)',
      front: DEFAULT_TEMPLATE_PATH_FRONT,
      back: DEFAULT_TEMPLATE_PATH_BACK,
      layout: 'sheet',
      perSheet: 4,
      grid: { columns: 2, rows: 2, gap: '6mm' },
      orderFront: [0, 1, 2, 3],
      orderBack: [1, 0, 3, 2],
    },
    {
      id: 'custom',
      label: 'Usar mi propia plantilla',
      front: '',
      back: '',
      layout: 'sheet',
      perSheet: 4,
      grid: { columns: 2, rows: 2, gap: '6mm' },
      orderFront: [0, 1, 2, 3],
      orderBack: [1, 0, 3, 2],
    },
  ],
  [MODES.JORNADA]: [
    {
      id: 'jornada',
      label: 'Gafetes Jornada (2 columnas · 4 filas)',
      front: JORNADA_TEMPLATE_PATH,
      back: JORNADA_TEMPLATE_PATH,
      layout: 'sheet',
      perSheet: 8,
      grid: { columns: 2, rows: 4, gap: '2mm' },
      orderFront: Array.from({ length: 8 }, (_, index) => index),
      orderBack: Array.from({ length: 8 }, (_, index) => index),
    },
    {
      id: 'custom-jornada',
      label: 'Usar mi propia plantilla de jornada',
      front: '',
      back: '',
      layout: 'sheet',
      perSheet: 8,
      grid: { columns: 2, rows: 4, gap: '2mm' },
      orderFront: Array.from({ length: 8 }, (_, index) => index),
      orderBack: Array.from({ length: 8 }, (_, index) => index),
    },
  ],
}

const CUSTOM_TEMPLATE_IDS = {
  [MODES.PERSONIFICADORES]: 'custom',
  [MODES.JORNADA]: 'custom-jornada',
}

const PREVIEW_LIMIT = 10

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

function toTitleCase(value) {
  return normalizeValue(value)
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => `${word.charAt(0).toUpperCase()}${word.slice(1).toLowerCase()}`)
    .join(' ')
}

function extractCompoundLastName(rawLastName) {
  const parts = normalizeValue(rawLastName).split(/\s+/).filter(Boolean)
  if (!parts.length) return ''
  const [first, second] = parts
  if (first.length <= 3 && second) {
    return toTitleCase(`${first} ${second}`)
  }
  return toTitleCase(first)
}

function extractSingleFirstName(rawFirstName) {
  const [first] = normalizeValue(rawFirstName).split(/\s+/).filter(Boolean)
  return toTitleCase(first || '')
}

function splitIntoLines(text, wordsPerLine = 2) {
  const words = normalizeValue(text).split(/\s+/).filter(Boolean)
  if (!words.length) return ['']

  const lines = []
  for (let i = 0; i < words.length; i += wordsPerLine) {
    lines.push(words.slice(i, i + wordsPerLine).join(' '))
  }
  return lines
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
      const company = toTitleCase(row['Empresa'])
      const rawFirstName = normalizeValue(row['Nombre'])
      const rawLastName = normalizeValue(row['Apellido'])

      const firstName = extractSingleFirstName(rawFirstName)
      const lastName = extractCompoundLastName(rawLastName)
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
  rowOffset = 0,
  isJornada = false,
}) {
  const isBack = variant === 'back'
  const templateSrc = hideTemplateImage ? '' : isBack ? template.back : template.front
  const [first] = attendees
  const baseScale = isBack ? fontScale.back : fontScale.front
  const primaryScale = baseScale * (isBack ? first.fontScaleBack ?? 1 : first.fontScaleFront ?? 1)
  const mergedAdjustments = {
    ...positionAdjustments,
    vertical: positionAdjustments.vertical + rowOffset,
  }
  const primaryStyles = buildNameStyles(first, isBack, mergedAdjustments, primaryScale, uniformMetrics)
  const nameLines = splitIntoLines(first.fullName || 'Nombre Apellido')
  const companyLines = splitIntoLines(first.company || 'Empresa')

  return (
    <section className={`badge badge--${variant} ${isJornada ? 'badge--jornada' : ''} ${templateSrc ? '' : 'badge--sheet-template'}`}>
      {templateSrc ? (
        <img className="badge__template" src={templateSrc} alt={`Plantilla ${variant}`} />
      ) : hideTemplateImage ? null : (
        <div className="badge__template badge__template--placeholder">Sube tu plantilla de {variant === 'front' ? 'frente' : 'reverso'}</div>
      )}

      <div className="names" style={primaryStyles}>
        <p className="name">
          {nameLines.map((line, index) => (
            <span key={`name-${index}`}>{line}</span>
          ))}
        </p>
        <p className="company">
          {companyLines.map((line, index) => (
            <span key={`company-${index}`}>{line}</span>
          ))}
        </p>
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

const EXPORT_SCALE = Math.max(5, window.devicePixelRatio * 3)

function chunkIntoSheets(groups, perSheet = 4) {
  const sheets = []
  for (let i = 0; i < groups.length; i += perSheet) {
    sheets.push(groups.slice(i, i + perSheet))
  }
  return sheets
}

function buildSlots(sheet, variant, template) {
  const perSheet = template?.perSheet || 4
  const orderFront = template?.orderFront || Array.from({ length: perSheet }, (_, index) => index)
  const orderBack = template?.orderBack || orderFront
  const order = variant === 'back' ? orderBack : orderFront
  const arranged = Array(perSheet).fill(null)

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
  const columns = template?.grid?.columns || 2
  const rows = template?.grid?.rows || Math.ceil((template?.perSheet || sheet.length || 4) / columns)
  const isJornadaTemplate = (template?.id || '').includes('jornada') || template?.perSheet === 8
  const sheetStyle = {
    gridTemplateColumns: `repeat(${columns}, 1fr)`,
    gridTemplateRows: `repeat(${rows}, 1fr)`,
    ...(template?.grid?.gap && !useSheetTemplate ? { gap: template.grid.gap } : {}),
    ...(isJornadaTemplate
      ? { alignContent: 'center', justifyItems: 'center', padding: '16mm 14mm', rowGap: template?.grid?.gap || '6mm' }
      : {}),
    ...(useSheetTemplate && sheetBackgroundImage ? { backgroundImage: sheetBackgroundImage } : {}),
  }
  const slots = buildSlots(sheet, variant, template)
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
      {slots.map((group, slotIndex) => {
        const rowIndex = Math.floor(slotIndex / columns)
        const rowOffset = isJornadaTemplate ? (rowIndex === 0 ? 4 : rowIndex === rows - 1 ? -4 : 0) : 0

        return (
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
                rowOffset={rowOffset}
                isJornada={isJornadaTemplate}
              />
            ) : (
              <div className="print-slot__placeholder" aria-hidden>
                Carga más nombres para completar esta hoja
              </div>
            )}
          </div>
        )
      })}
    </section>
  )
}

function App() {
  const [attendees, setAttendees] = useState([])
  const [error, setError] = useState('')
  const [activeMode, setActiveMode] = useState(MODES.PERSONIFICADORES)
  const [templateSelection, setTemplateSelection] = useState({
    [MODES.PERSONIFICADORES]: TEMPLATE_OPTIONS[MODES.PERSONIFICADORES][0].id,
    [MODES.JORNADA]: TEMPLATE_OPTIONS[MODES.JORNADA][0].id,
  })
  const [customTemplates, setCustomTemplates] = useState({
    [MODES.PERSONIFICADORES]: { front: '', back: '' },
    [MODES.JORNADA]: { front: '', back: '' },
  })
  const [positionAdjustments, setPositionAdjustments] = useState({ vertical: 0, gap: 0, width: 0, horizontal: 0 })
  const [fontScale, setFontScale] = useState({ front: 1, back: 1 })
  const [attendeeOverrides, setAttendeeOverrides] = useState({})
  const [editingIndex, setEditingIndex] = useState(0)
  const [objectUrls, setObjectUrls] = useState([])
  const [isExporting, setIsExporting] = useState(false)
  const [quickSearch, setQuickSearch] = useState('')
  const [highlightedSuggestion, setHighlightedSuggestion] = useState(null)
  const [isPrinting, setIsPrinting] = useState(false)
  const [useUniformScaling, setUseUniformScaling] = useState(false)
  const [manualForm, setManualForm] = useState({ company: '', firstName: '', lastName: '' })
  const [editingManualIndex, setEditingManualIndex] = useState(null)
  const [listName, setListName] = useState('')
  const [savedLists, setSavedLists] = useState([])

  const templateOptions = TEMPLATE_OPTIONS[activeMode] || []
  const templateId = templateSelection[activeMode] || templateOptions[0]?.id || ''
  const customTemplate = customTemplates[activeMode] || { front: '', back: '' }
  const activeTemplateOption = templateOptions.find((option) => option.id === templateId) || templateOptions[0]
  const activeTemplate =
    activeTemplateOption?.id === CUSTOM_TEMPLATE_IDS[activeMode]
      ? { ...activeTemplateOption, front: customTemplate.front, back: customTemplate.back }
      : activeTemplateOption || {
          id: 'fallback',
          label: 'Plantilla base',
          front: '',
          back: '',
          layout: 'sheet',
          perSheet: 4,
          grid: { columns: 2, rows: 2, gap: '6mm' },
          orderFront: [0, 1, 2, 3],
          orderBack: [1, 0, 3, 2],
        }

  const perSheet = activeTemplate?.perSheet || 4
  const isPersonMode = activeMode === MODES.PERSONIFICADORES
  const perSheetLabel = `${perSheet} gafetes por hoja`

  useEffect(
    () => () => {
      objectUrls.forEach((url) => URL.revokeObjectURL(url))
    },
    [objectUrls]
  )

  useEffect(() => {
    try {
      const storedLists = localStorage.getItem('saved-badge-lists')
      if (storedLists) {
        setSavedLists(JSON.parse(storedLists))
      }
    } catch (err) {
      console.error(err)
    }
  }, [])

  useEffect(() => {
    setEditingIndex((prev) => {
      if (!attendees.length) return 0
      if (prev >= attendees.length) return 0
      return prev
    })
  }, [attendees.length])

  useEffect(() => {
    setEditingIndex(0)
    setQuickSearch('')
    setHighlightedSuggestion(null)
  }, [activeMode])

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
      resetManualForm()
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
    setCustomTemplates((prev) => ({
      ...prev,
      [activeMode]: { ...(prev[activeMode] || { front: '', back: '' }), [side]: url },
    }))
    setTemplateSelection((prev) => ({ ...prev, [activeMode]: CUSTOM_TEMPLATE_IDS[activeMode] }))
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
    resetManualForm()
  }

  const handleModeChange = (mode) => {
    setActiveMode(mode)
  }

  const resetManualForm = () => {
    setEditingManualIndex(null)
    setManualForm({ company: '', firstName: '', lastName: '' })
  }

  const persistSavedLists = (updater) => {
    setSavedLists((prev) => {
      const next = typeof updater === 'function' ? updater(prev) : updater
      localStorage.setItem('saved-badge-lists', JSON.stringify(next))
      return next
    })
  }

  const handleSaveList = () => {
    setError('')
    if (!attendees.length) {
      setError('Necesitas al menos una persona para guardar la lista.')
      return
    }

    const trimmedName = listName.trim()
    if (!trimmedName) {
      setError('Asigna un nombre a la lista para guardarla.')
      return
    }

    const payload = {
      id: `${activeMode}-${Date.now()}`,
      name: trimmedName,
      mode: activeMode,
      attendees: attendees.map((item) => ({ ...item })),
      overrides: { ...attendeeOverrides },
    }

    persistSavedLists((prev) => {
      const existingIndex = prev.findIndex((item) => item.name === trimmedName && item.mode === activeMode)
      if (existingIndex >= 0) {
        const updated = [...prev]
        updated[existingIndex] = { ...payload, id: prev[existingIndex].id }
        return updated
      }
      return [...prev, payload]
    })

    setListName('')
  }

  const handleLoadList = (id) => {
    const target = savedLists.find((item) => item.id === id && item.mode === activeMode)
    if (!target) return
    setAttendees(target.attendees || [])
    setAttendeeOverrides(target.overrides || {})
    setEditingIndex(0)
    setQuickSearch('')
    resetManualForm()
  }

  const handleDeleteList = (id) => {
    persistSavedLists((prev) => prev.filter((item) => item.id !== id))
  }

  const rebuildOverridesAfterRemoval = (overrides, removedIndex) => {
    const next = {}
    Object.entries(overrides).forEach(([key, value]) => {
      const numericKey = Number(key)
      if (Number.isNaN(numericKey) || numericKey === removedIndex) return
      const targetIndex = numericKey > removedIndex ? numericKey - 1 : numericKey
      next[targetIndex] = value
    })
    return next
  }

  const handleRemovePerson = (index) => {
    setEditingManualIndex((current) => (current === index ? null : current))
    setAttendees((prev) => {
      const next = prev.filter((_, idx) => idx !== index)
      setAttendeeOverrides((overrides) => rebuildOverridesAfterRemoval(overrides, index))
      setEditingIndex((current) => Math.max(0, Math.min(current, next.length - 1)))
      return next
    })
  }

  const handleManualSubmit = (event) => {
    event.preventDefault()
    setError('')

    const [entry] = buildAttendees([
      { Empresa: manualForm.company, Nombre: manualForm.firstName, Apellido: manualForm.lastName },
    ])

    if (!entry) {
      setError('Completa el nombre o la empresa para añadirlo a la lista.')
      return
    }

    setAttendees((prev) => {
      const next = [...prev]
      if (editingManualIndex !== null && editingManualIndex >= 0) {
        next[editingManualIndex] = entry
      } else {
        next.push(entry)
      }
      setEditingIndex((current) => (editingManualIndex !== null ? Math.min(current, next.length - 1) : next.length - 1))
      return next
    })

    setAttendeeOverrides((prev) => {
      if (editingManualIndex === null) return prev
      const { [editingManualIndex]: removed, ...rest } = prev
      return rest
    })

    resetManualForm()
  }

  const handleStartManualEdit = (index) => {
    const target = attendees[index]
    if (!target) return
    setEditingManualIndex(index)
    setManualForm({ company: target.company, firstName: target.firstName, lastName: target.lastName })
    setEditingIndex(index)
  }

  const updatePosition = (field, value) => {
    setPositionAdjustments((prev) => ({ ...prev, [field]: value }))
  }

  const updateFontScale = (side, value) => {
    setFontScale((prev) => ({ ...prev, [side]: value }))
  }

  const updateAttendeeOverride = (index, field, value) => {
    const normalizedValue = ['name', 'company'].includes(field) ? toTitleCase(value) : value
    setAttendeeOverrides((prev) => ({
      ...prev,
      [index]: {
        ...(prev[index] || {}),
        [field]: normalizedValue,
      },
    }))
  }

  const resetAttendeeOverride = (index) => {
    setAttendeeOverrides((prev) => {
      const { [index]: removed, ...rest } = prev
      return rest
    })
  }

  const missingCustomTemplate = useMemo(
    () =>
      activeTemplate?.id === CUSTOM_TEMPLATE_IDS[activeMode] && (!activeTemplate?.front || !activeTemplate?.back),
    [activeMode, activeTemplate]
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

  const savedListsForMode = useMemo(
    () => savedLists.filter((item) => item.mode === activeMode),
    [activeMode, savedLists]
  )

  const filteredSuggestions = useMemo(() => {
    if (!quickSearch.trim()) return suggestionCatalog.slice(0, 8)
    const normalized = quickSearch.toLowerCase()
    return suggestionCatalog
      .filter((item) => item.label.toLowerCase().includes(normalized))
      .slice(0, 8)
  }, [quickSearch, suggestionCatalog])

  const badgeGroups = useMemo(() => decoratedAttendees.map((attendee) => [attendee]), [decoratedAttendees])

  const exportSheets = useMemo(() => chunkIntoSheets(badgeGroups, perSheet), [badgeGroups, perSheet])
  const previewBadgeGroups = useMemo(() => badgeGroups.slice(0, PREVIEW_LIMIT), [badgeGroups])
  const previewSheets = useMemo(
    () => chunkIntoSheets(previewBadgeGroups, perSheet),
    [previewBadgeGroups, perSheet]
  )
  const totalSheets = useMemo(() => exportSheets.length, [exportSheets])
  const totalPeople = useMemo(() => decoratedAttendees.length, [decoratedAttendees])
  const activeSheetIndex = useMemo(() => Math.floor(editingIndex / perSheet), [editingIndex, perSheet])
  const activeSheet = exportSheets[activeSheetIndex] || []

  const editingPerson = decoratedAttendees[editingIndex] || null
  const editingPositionLabel = decoratedAttendees.length ? `#${editingIndex + 1} de ${decoratedAttendees.length}` : 'Sin selección'
  const isPersonCustomized = Boolean(attendeeOverrides[editingIndex])

  const activeBadgeGroup = useMemo(
    () => (decoratedAttendees.length ? [decoratedAttendees[editingIndex] || decoratedAttendees[0]] : []),
    [decoratedAttendees, editingIndex]
  )

  const uniformMetrics = useMemo(
    () => (useUniformScaling ? buildUniformMetrics(decoratedAttendees) : null),
    [decoratedAttendees, useUniformScaling]
  )

  const handleDownloadPDF = async () => {
    if (!badgeGroups.length || missingCustomTemplate) return
    setIsExporting(true)
    setError('')

    try {
      await new Promise((resolve) => requestAnimationFrame(() => resolve()))
      await Promise.all([preloadImage(activeTemplate.front), preloadImage(activeTemplate.back)])
      await new Promise((resolve) => setTimeout(resolve, 150))

      const sheetsToExport = Array.from(document.querySelectorAll('.sheet-grid--export .print-sheet')).filter(
        (sheet) => sheet.dataset.hasContent === 'true'
      )
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' })
      const pageWidth = pdf.internal.pageSize.getWidth()
      const pageHeight = pdf.internal.pageSize.getHeight()

      for (const [index, sheet] of sheetsToExport.entries()) {
        const scaledWidth = Math.round(sheet.offsetWidth)
        const scaledHeight = Math.round(sheet.offsetHeight)

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
          imageTimeout: 0,
          removeContainer: true,
        })
        const imgData = canvas.toDataURL('image/png', 1)
        const imgProps = pdf.getImageProperties(imgData)
        const ratioHeight = (imgProps.height * pageWidth) / imgProps.width
        const ratioWidth = (imgProps.width * pageHeight) / imgProps.height
        const finalWidth = ratioHeight > pageHeight ? ratioWidth : pageWidth
        const finalHeight = ratioHeight > pageHeight ? pageHeight : ratioHeight

        if (index > 0) pdf.addPage()
        pdf.addImage(imgData, 'PNG', 0, 0, finalWidth, finalHeight)
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
        <div className="hero__intro">
          <p className="eyebrow">
            {perSheet} gafetes por hoja · Guarda listas locales · Impresión a doble cara
          </p>
          <h1>Generador de {isPersonMode ? 'personificadores' : 'gafetes de jornada'}</h1>
          <p className="lead">
            {isPersonMode
              ? 'Migra la experiencia original de personificadores: ajusta la plantilla carta, sube tu Excel y personaliza el texto con total control.'
              : 'Usa la plantilla de Gafetes Jornada (2 columnas x 4 filas), carga tu Excel y obtén las 8 credenciales alineadas en una sola hoja.'}
          </p>

          <div className="mode-grid" role="tablist" aria-label="Tipo de gafete">
            <button
              type="button"
              className={`mode-card ${isPersonMode ? 'mode-card--active' : ''}`}
              onClick={() => handleModeChange(MODES.PERSONIFICADORES)}
              role="tab"
              aria-selected={isPersonMode}
            >
              <div className="mode-card__header">
                <span className="pill pill--neutral">4 por hoja</span>
                {isPersonMode && <span className="pill pill--accent">Activo</span>}
              </div>
              <h3>Personificadores</h3>
              <p>Todo lo que ya tenías: controles finos, plantillas personalizadas y vista previa doble.</p>
            </button>

            <button
              type="button"
              className={`mode-card ${!isPersonMode ? 'mode-card--active' : ''}`}
              onClick={() => handleModeChange(MODES.JORNADA)}
              role="tab"
              aria-selected={!isPersonMode}
            >
              <div className="mode-card__header">
                <span className="pill pill--neutral">8 por hoja</span>
                {!isPersonMode && <span className="pill pill--accent">Activo</span>}
              </div>
              <h3>Gafetes Jornada</h3>
              <p>Replica el flujo con la plantilla de 2 columnas y 4 filas usando Gafetes_jornada.png.</p>
            </button>
          </div>
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
          <div className="actions__meta">
            <p className="helper">Consejo: imprime en doble cara con unión por borde largo.</p>
            <div className="pill pill--neutral">Plantilla activa: {activeTemplate?.label}</div>
          </div>
        </div>
      </header>

      <section className="panel panel--stacked panel--soft">
        <div className="panel__heading">
          <div>
            <p className="panel__title">Captura rápida y listas locales</p>
            <p className="helper">
              Añade o edita nombres manualmente, guarda la lista en tu navegador y recupérala en cualquiera de las dos secciones
              sin volver a cargar un Excel.
            </p>
          </div>
        </div>
        <div className="list-tools">
          <form className="list-tools__block" onSubmit={handleManualSubmit}>
            <p className="block__title">Añade o edita una persona</p>
            <div className="controls controls--inline">
              <label className="control">
                <span>Nombre(s)</span>
                <input
                  type="text"
                  value={manualForm.firstName}
                  onChange={(event) => setManualForm((prev) => ({ ...prev, firstName: event.target.value }))}
                  placeholder="Ej. Ana Sofía"
                />
              </label>
              <label className="control">
                <span>Apellidos</span>
                <input
                  type="text"
                  value={manualForm.lastName}
                  onChange={(event) => setManualForm((prev) => ({ ...prev, lastName: event.target.value }))}
                  placeholder="Ej. Del Castillo"
                />
              </label>
              <label className="control">
                <span>Empresa</span>
                <input
                  type="text"
                  value={manualForm.company}
                  onChange={(event) => setManualForm((prev) => ({ ...prev, company: event.target.value }))}
                  placeholder="Nombre de la empresa"
                />
              </label>
            </div>
            <div className="inline-actions">
              <button type="submit" className="primary">
                {editingManualIndex !== null ? 'Actualizar persona' : 'Agregar a la lista'}
              </button>
              {editingManualIndex !== null && (
                <button type="button" className="ghost" onClick={resetManualForm}>
                  Cancelar edición
                </button>
              )}
            </div>
            <p className="helper">Los nombres se normalizan con mayúscula inicial y apellidos compuestos.</p>
          </form>

          <div className="list-tools__block">
            <p className="block__title">Listas guardadas para esta sección</p>
            <div className="controls">
              <label className="control">
                <span>Nombre de la lista</span>
                <input
                  type="text"
                  value={listName}
                  onChange={(event) => setListName(event.target.value)}
                  placeholder="Ej. Jornada matutina"
                />
              </label>
              <div className="inline-actions">
                <button type="button" className="primary" onClick={handleSaveList} disabled={!attendees.length}>
                  Guardar lista actual
                </button>
                <span className="control__hint">Se almacenan localmente en este navegador.</span>
              </div>
              <div className="saved-lists" role="list">
                {savedListsForMode.length ? (
                  savedListsForMode.map((item) => (
                    <div key={item.id} className="saved-list" role="listitem">
                      <div>
                        <strong>{item.name}</strong>
                        <p className="helper">{item.attendees?.length || 0} personas guardadas</p>
                      </div>
                      <div className="inline-actions">
                        <button type="button" className="ghost" onClick={() => handleLoadList(item.id)}>
                          Cargar
                        </button>
                        <button type="button" className="ghost danger" onClick={() => handleDeleteList(item.id)}>
                          Eliminar
                        </button>
                      </div>
                    </div>
                  ))
                ) : (
                  <p className="empty">Aún no hay listas guardadas para esta vista.</p>
                )}
              </div>
            </div>
          </div>
        </div>

        {attendees.length > 0 && (
          <div className="people-table" role="list">
            {attendees.map((person, index) => (
              <div className="people-row" key={`${person.fullName}-${index}`} role="listitem">
                <div>
                  <p className="people-row__name">#{index + 1} · {person.fullName || 'Sin nombre'}</p>
                  <p className="people-row__meta">{person.company || 'Sin empresa'} · Apellido base: {person.lastName || 'N/A'}</p>
                </div>
                <div className="inline-actions">
                  <button type="button" className="ghost" onClick={() => handleStartManualEdit(index)}>
                    Editar
                  </button>
                  <button type="button" className="ghost danger" onClick={() => handleRemovePerson(index)}>
                    Eliminar
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      <section className="panel">
        <div>
          <p className="panel__title">Plantillas</p>
          <div className="controls">
            <label className="control">
              <span>Selecciona plantilla</span>
              <select
                value={templateId}
                onChange={(event) => setTemplateSelection((prev) => ({ ...prev, [activeMode]: event.target.value }))}
              >
                {templateOptions.map((option) => (
                  <option key={option.id} value={option.id}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>

            {templateId === CUSTOM_TEMPLATE_IDS[activeMode] && (
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
          <label className="control control--inline control--toggle">
            <input
              type="checkbox"
              checked={useUniformScaling}
              onChange={(event) => setUseUniformScaling(event.target.checked)}
            />
            <span>
              Mantener el mismo tamaño para todos (útil cuando quieres que ningún nombre sobresalga, desactívalo para que los
              controles globales y por persona surtan efecto).
            </span>
          </label>
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
            </div>

                <div className="person-preview person-preview--stacked">
                  <div className="person-preview__header">
                    <p className="person-preview__label">
                      Vista previa en vivo de <strong>{editingPerson?.fullName || 'la persona seleccionada'}</strong>
                    </p>
                    <div className="pill pill--neutral">{perSheetLabel}</div>
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
            Ajusta los deslizadores hasta que el texto caiga en el lugar exacto de tu plantilla. Cada hoja acomoda
            {perSheetLabel} listos para imprimir frente y reverso con la misma orientación.
          </p>
          {totalPeople > PREVIEW_LIMIT && (
            <p className="helper">
              Solo se previsualizan las primeras {PREVIEW_LIMIT} personas para mantener la app ágil. La descarga incluye
              las {totalPeople} personas cargadas.
            </p>
          )}
        </div>
        {!attendees.length && <p className="empty">Sube tu Excel o usa el ejemplo para comenzar.</p>}
        {missingCustomTemplate && <p className="empty">Sube ambos lados de la plantilla personalizada para generar la vista previa.</p>}
        <div className="sheet-grid sheet-grid--preview">
          {previewSheets.map((sheet, index) => (
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
        <div className="sheet-grid sheet-grid--export" aria-hidden>
          {exportSheets.map((sheet, index) => (
            <div className="sheet-pair" key={`export-sheet-${index}`}>
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
