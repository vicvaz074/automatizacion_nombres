import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const FUTURA_STACK = "'Futura', 'Futura PT', 'Century Gothic', 'Arial', sans-serif"

const demoRows = [
  { company: 'davara Abogados', lastName: 'Rangel', firstName: 'María' },
  { company: 'Tsuru', lastName: 'Aguayo', firstName: 'Diego' },
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
  const companyLongWordScale = companyLongestWord >= 18 ? 0.82 : companyLongestWord >= 14 ? 0.88 : companyLongestWord >= 12 ? 0.93 : 1
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
        10
    ) / 10,
    Math.round(nameToCompanyBalance * 10) / 10
  )

  const baseGap = nameFontSize >= 26 ? 7.2 : nameFontSize >= 22 ? 6.7 : 6.1
  const multilineGapBoost = estimatedNameLines > 1 ? 1.25 + (estimatedNameLines - 1) * 0.22 : 1
  const companyLinesBoost =
    estimatedCompanyLines >= 3
      ? 1.38
      : estimatedCompanyLines === 2
        ? 1.22
        : 1
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

function BadgeFace({ attendee, variant = 'front' }) {
  const { fullName, company } = attendee
  const templateSrc = variant === 'front' ? '/Plantilla_hoja_1.png' : '/Plantilla_hoja_2.png'

  const { nameFontSize, companyFontSize, namesGap, namesOffset, namesWidth } = useMemo(
    () => getTypographyMetrics(fullName, company),
    [company, fullName]
  )

  const isBack = variant === 'back'
  const adjustedGap = namesGap + (isBack ? 0.6 : 0)
  const adjustedOffset = namesOffset + (isBack ? 1.4 : 0)
  const adjustedWidth = Math.max(62, namesWidth - (isBack ? 2 : 0))
  const topOffset = Math.max(15, adjustedOffset - 5.5)
  const bottomOffset = Math.max(12, adjustedOffset - 8.5)

  const namesStyles = useMemo(
    () => ({
      fontFamily: FUTURA_STACK,
      '--name-size': `${nameFontSize}pt`,
      '--company-size': `${companyFontSize}pt`,
      '--names-gap': `${adjustedGap}mm`,
      '--names-offset': `${adjustedOffset}mm`,
      '--names-offset-top': `${topOffset}mm`,
      '--names-offset-bottom': `${bottomOffset}mm`,
      '--names-width': `${adjustedWidth}mm`,
    }),
    [adjustedGap, adjustedOffset, adjustedWidth, bottomOffset, companyFontSize, nameFontSize, topOffset]
  )

  return (
    <section className={`badge badge--${variant}`}>
      <img className="badge__template" src={templateSrc} alt={`Plantilla ${variant}`} />

      <div className="names" style={namesStyles}>
        <p className="name">{fullName || 'Nombre Apellido'}</p>
        <p className="company">{company || 'Empresa'}</p>
      </div>

      <div className="names names--mirrored" style={namesStyles}>
        <p className="name">{fullName || 'Nombre Apellido'}</p>
        <p className="company">{company || 'Empresa'}</p>
      </div>
    </section>
  )
}

function App() {
  const [attendees, setAttendees] = useState([])
  const [error, setError] = useState('')

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

  const loadDemo = () => {
    const sampleRows = demoRows.map((row) => ({
      Empresa: row.company,
      Apellido: row.lastName,
      Nombre: row.firstName,
    }))
    setAttendees(buildAttendees(sampleRows))
    setError('')
  }

  const totalBadges = useMemo(() => attendees.length, [attendees])

  return (
    <div className="page">
      <header className="hero">
        <div>
          <p className="eyebrow">Plantilla 10cm x 10cm · Impresión a doble cara</p>
          <h1>Generador de gafetes</h1>
          <p className="lead">
            Sube un Excel con las columnas <strong>Empresa</strong>, <strong>Apellido</strong> y{' '}
            <strong>Nombre</strong>. El orden final será <strong>Nombre Apellido</strong>. Todo lo demás
            conserva la posición y estilos de la plantilla.
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
          <button type="button" className="primary" onClick={() => window.print()} disabled={!attendees.length}>
            Imprimir (frente y reverso)
          </button>
          <p className="helper">Consejo: imprime en doble cara con unión por borde largo.</p>
        </div>
      </header>

      {error && <div className="alert">{error}</div>}

      <section className="status">
        <div>
          <p className="label">Registros listos</p>
          <strong className="stat">{totalBadges}</strong>
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
            Cada persona genera dos páginas: frente y reverso. La tipografía usa el stack{' '}
            <code>{FUTURA_STACK}</code> con un tamaño dinámico: el nombre parte de 27pt en negrita y la
            empresa de 16pt en peso normal, ambos se ajustan automáticamente para que sigan siendo legibles
            aunque el texto sea largo.
          </p>
        </div>
        {!attendees.length && <p className="empty">Sube tu Excel o usa el ejemplo para comenzar.</p>}
        <div className="badge-grid">
          {attendees.map((attendee, index) => (
            <div className="badge-pair" key={`${attendee.fullName}-${index}`}>
              <BadgeFace attendee={attendee} variant="front" />
              <BadgeFace attendee={attendee} variant="back" />
            </div>
          ))}
        </div>
      </section>
    </div>
  )
}

export default App
