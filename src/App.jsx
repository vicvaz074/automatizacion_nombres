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
  const templateSrc = variant === 'front' ? '/template-front.svg' : '/template-back.svg'

  return (
    <section className={`badge badge--${variant}`}>
      <img className="badge__template" src={templateSrc} alt={`Plantilla ${variant}`} />

      <div className="names" style={{ fontFamily: FUTURA_STACK }}>
        <p className="name">{fullName || 'Nombre Apellido'}</p>
        <p className="company">{company || 'Empresa'}</p>
      </div>

      <div className="names names--mirrored" style={{ fontFamily: FUTURA_STACK }}>
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
            <code>{FUTURA_STACK}</code>, con tamaños: Nombre 33.3pt en negrita, Empresa 22.6pt en negrita y
            la leyenda 5.9pt sin negrita.
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
