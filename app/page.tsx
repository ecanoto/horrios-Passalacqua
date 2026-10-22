'use client'

import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { Button } from '@/components/ui/button'
import { CalendarDays, Check, ChevronDown, Cloud, Download, FileSpreadsheet, ListFilter, Pencil, Printer, Search, ShieldAlert, ShieldCheck, UserCheck, UserX, Users, X, Zap, Clock, Trash2, Plus } from 'lucide-react'

type View = 'horario' | 'docentes' | 'ausencias' | 'preceptores' | 'paro'
type Status = 'Titular' | 'Provisional' | 'Titular Interino'
type Attendance = 'normal' | 'paro' | 'licencia'
type CoberturaStatus = 'Cubierto por Decreto' | 'Sin cubrir'

type Block = { 
  id: string; 
  curso: string; 
  turno: 'Teoría' | 'Taller'; 
  tipo: 'Teoría' | 'Taller' | 'Recreo'; 
  dia: string; 
  horario: string; 
  materia: string; 
  docente: string; 
  situacionRevista: Status; 
  suplente?: string; 
  grupo?: string 
}

type AusenciaRecord = {
  id: string;
  curso: string;
  mes: string;
  fecha: string;
  materia: string;
  docenteAusente: string;
  estadoCobertura: CoberturaStatus;
  docenteDecreto: string;
}

type PreceptorAssignment = {
  curso: string;
  preceptorTitular: string;
  preceptorSuplente: string;
}

type ParoRecord = {
  id: string;
  fecha: string;
  curso: string;
  horario: string;
  materia: string;
  docente: string;
  estadoParo: Attendance;
  observaciones?: string;
}

const logo = 'https://hebbkx1anhila5yf.public.blob.vercel-storage.com/logo-M35qmFS9dpbN14cGWOmnvOvidznySB.png'
const spreadsheetApi = 'https://script.google.com/macros/s/AKfycbyzwuM5bDtNhaBZUHsgN7BzxcMwPEBqORz_PkGuzYpo7Y2jYclWMUMnqXGXN7WPZDJ7/exec'
const days = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
const courses = ['1°1°', '1°2°', '1°3°', '1°4°', '1°5°', '1°6°', '2°1°', '2°2°', '2°3°', '2°4°', '2°5°', '2°6°', '3°1°', '3°2°', '3°3°', '3°4°', '3°5°', '3°6°', '4°1°', '4°2°', '4°3°', '4°4°', '4°5°', '5°1°', '5°2°', '5°3°', '5°6°', '5°7°', '6°1°', '6°2°', '6°3°', '6°6°', '6°7°', '7°1°', '7°2°', '7°3°', '7°6°', '7°7°']

const seed: Block[] = [
  { id: '1', curso: '1°1°', turno: 'Teoría', tipo: 'Teoría', dia: 'Lunes', horario: '07:40–08:40', materia: 'Cs. Naturales', docente: 'MARTIN', situacionRevista: 'Titular' },
  { id: '2', curso: '1°1°', turno: 'Teoría', tipo: 'Recreo', dia: 'Lunes', horario: '09:40–09:50', materia: 'Recreo', docente: '', situacionRevista: 'Titular' },
  { id: '3', curso: '1°1°', turno: 'Teoría', tipo: 'Teoría', dia: 'Jueves', horario: '11:50–12:50', materia: 'Cs. Sociales', docente: 'BOGARIN', situacionRevista: 'Provisional' },
  { id: '4', curso: '1°1°', turno: 'Taller', tipo: 'Taller', dia: 'Miércoles', horario: '13:00–15:00', materia: 'Lenguajes Tecnológicos', docente: 'OSORIO MARCELA', situacionRevista: 'Provisional', grupo: 'Grupo A' },
]

const seedAusencias: AusenciaRecord[] = [
  { id: 'AUS-001', curso: '1°1°', mes: 'Abril', fecha: '2026-04-06', materia: 'Cs. Sociales', docenteAusente: 'BOGARIN', estadoCobertura: 'Cubierto por Decreto', docenteDecreto: 'ASPREA' },
]

const seedPreceptores: PreceptorAssignment[] = courses.map((c, idx) => ({
  curso: c,
  preceptorTitular: ['CAYATA', 'FEMIA', 'CEPEDA', 'NAHIA', 'ALFONSIN', 'GRISELDA'][idx % 6],
  preceptorSuplente: idx % 3 === 0 ? 'RODRIGUEZ SILVINA' : ''
}))

const statusStyle: Record<Status, string> = { Titular: 'text-blue-700 bg-blue-50 border-blue-200', Provisional: 'text-red-700 bg-red-50 border-red-200', 'Titular Interino': 'text-gray-950 bg-gray-100 border-gray-300' }

export default function Page() {
  const [view, setView] = useState<View>('horario')
  const [blocks, setBlocks] = useState(seed)
  const [ausencias, setAusencias] = useState<AusenciaRecord[]>(seedAusencias)
  const [preceptores, setPreceptores] = useState<PreceptorAssignment[]>(seedPreceptores)
  const [paroHistorial, setParoHistorial] = useState<ParoRecord[]>([])
  
  const [course, setCourse] = useState('1°1°')
  const [filter, setFilter] = useState<'Todos' | 'Teoría' | 'Taller'>('Todos')
  const [selectedDays, setSelectedDays] = useState(days)
  const [query, setQuery] = useState('')
  const [editing, setEditing] = useState<Block | null>(null)
  const [showRecreosModal, setShowRecreosModal] = useState(false)
  const [dirty, setDirty] = useState(0)
  const [saved, setSaved] = useState(false)
  const [syncing, setSyncing] = useState(false)
  const [syncMessage, setSyncMessage] = useState('Conectando con Google Drive…')

  useEffect(() => { syncSpreadsheet() }, [])

  const courseTimes = useMemo(() => {
    const courseBlocks = blocks.filter((b) => b.curso === course)
    const timesSet = new Set(courseBlocks.map((b) => b.horario))
    return Array.from(timesSet).sort((a, b) => a.localeCompare(b))
  }, [blocks, course])

  const visible = blocks.filter((b) => b.curso === course && selectedDays.includes(b.dia) && (filter === 'Todos' || b.tipo === filter || b.tipo === 'Recreo'))
  const teachers = useMemo(() => [...new Set(blocks.flatMap((b) => [b.docente, b.suplente].filter(Boolean)))], [blocks])
  const searchResults = blocks.filter((b) => (b.docente + ' ' + (b.suplente ?? '')).toLowerCase().includes(query.toLowerCase()))
  const dayBlocks = blocks.filter((b) => b.tipo !== 'Recreo')

  function updateBlock(next: Block) { 
    setBlocks((current) => current.map((b) => b.id === next.id ? next : b))
    setEditing(null)
    setDirty((n) => n + 1) 
  }
  
  function exportExcel() { 
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(blocks.map(({ id, ...r }) => ({ ...r, 'Situación Revista': r.situacionRevista }))), 'DATA')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ausencias), 'AUSENCIAS')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(preceptores), 'PRECEPTORES')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(paroHistorial), 'PARO_HISTORIAL')
    XLSX.writeFile(wb, 'gestion-escolar-eest6.xlsx') 
  }

  async function syncSpreadsheet() {
    setSyncing(true)
    try {
      const response = await fetch(spreadsheetApi, { cache: 'no-store' })
      if (!response.ok) throw new Error('No se pudo leer la planilla')
      const payload = await response.json()
      const rows = Array.isArray(payload) ? payload : payload.data ?? payload.rows ?? payload.blocks ?? payload.values
      
      const read = (row: any, ...keys: string[]) => { const found = keys.find((key) => row?.[key] !== undefined); return found ? row[found] : '' }
      
      if (Array.isArray(rows) && rows.length) {
        const normalized = rows.map((row: any, index: number) => ({ 
          id: String(read(row, 'id', 'ID') || index + 1), 
          curso: String(read(row, 'curso', 'Curso')), 
          turno: read(row, 'turno', 'Turno') || 'Teoría', 
          tipo: read(row, 'tipo', 'Tipo') || 'Teoría', 
          dia: read(row, 'dia', 'Día', 'Dia'), 
          horario: read(row, 'horario', 'Horario'), 
          materia: read(row, 'materia', 'Materia'), 
          docente: read(row, 'docente', 'Docente'), 
          situacionRevista: read(row, 'situacionRevista', 'Situación Revista', 'Situacion Revista') || 'Titular', 
          suplente: read(row, 'suplente', 'Suplente'), 
          grupo: read(row, 'grupo', 'Grupo') 
        }))
        setBlocks(normalized)
      }

      if (payload.ausencias && Array.isArray(payload.ausencias)) {
        const normAus = payload.ausencias.map((a: any, index: number) => ({
          id: String(read(a, 'id', 'ID') || `AUS-${index + 1}`),
          curso: String(read(a, 'curso', 'Curso') || '1°1°'),
          mes: String(read(a, 'mes', 'Mes') || 'Octubre'),
          fecha: String(read(a, 'fecha', 'Fecha', 'Fecha / Horario') || new Date().toISOString().split('T')[0]),
          materia: String(read(a, 'materia', 'Materia') || ''),
          docenteAusente: String(read(a, 'docenteAusente', 'Docente Ausente', 'docente') || ''),
          estadoCobertura: (read(a, 'estadoCobertura', 'Estado Cobertura') || 'Sin cubrir') as CoberturaStatus,
          docenteDecreto: String(read(a, 'docenteDecreto', 'Docente Decreto', 'PF Cubre por Decreto / Observación') || 'Sin cubrir')
        }))
        setAusencias(normAus)
      }

      if (payload.paroHistorial && Array.isArray(payload.paroHistorial)) {
        setParoHistorial(payload.paroHistorial)
      }

      setSyncMessage(`Conectado exitosamente (${courses.length} cursos cargados)`)
    } catch { 
      setSyncMessage('No se pudo conectar · usando datos locales') 
    } finally { 
      setSyncing(false) 
    }
  }

  async function save() { 
    setSyncing(true)
    try { 
      await fetch(spreadsheetApi, { 
        method: 'POST', 
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' }, 
        body: JSON.stringify({ 
          action: 'replace', 
          rows: blocks.map(({ id, ...row }) => row), 
          ausencias: ausencias,
          paroHistorial: paroHistorial 
        }) 
      }) 
      setSaved(true)
      setDirty(0)
      setSyncMessage('Guardado en Google Drive') 
      setTimeout(() => setSaved(false), 2600) 
    } catch { 
      setSyncMessage('No se pudo guardar en Google Drive') 
    } finally { 
      setSyncing(false) 
    } 
  }

  return <div className="min-h-screen bg-[#f5f8f2] text-slate-950">
    <header className="border-b border-[#d7e8c9] bg-white print:hidden">
      <div className="mx-auto flex max-w-[1500px] flex-col sm:flex-row items-center justify-between gap-3 px-4 sm:px-5 py-3">
        <div className="flex items-center gap-3 w-full sm:w-auto">
          <img src={logo} alt="Escudo E.E.S.T. N° 6" className="size-12 sm:size-14 object-contain shrink-0" />
          <div>
            <div className="text-[10px] sm:text-[11px] font-extrabold uppercase tracking-[0.22em] text-[#168000]">E.E.S.T. N° 6</div>
            <h1 className="text-base sm:text-lg font-black tracking-tight">Gestión de Horarios, Preceptores y Paros</h1>
            <p className="text-[11px] sm:text-xs text-slate-500">Banfield · Lomas de Zamora</p>
          </div>
        </div>
        <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
          <Button variant="outline" size="sm" onClick={syncSpreadsheet} disabled={syncing}><Download data-icon="inline-start" />{syncing ? 'Conectando…' : 'Sincronizar'}</Button>
          <Button variant="outline" size="sm" onClick={() => window.print()}><Printer data-icon="inline-start" />Imprimir</Button>
        </div>
      </div>
      <nav className="mx-auto flex max-w-[1500px] gap-1 overflow-x-auto px-4 sm:px-5 scrollbar-none" aria-label="Secciones">
        <Tab active={view === 'horario'} onClick={() => setView('horario')} icon={<CalendarDays className="size-4 shrink-0" />} label="Horarios" />
        <Tab active={view === 'docentes'} onClick={() => setView('docentes')} icon={<Users className="size-4 shrink-0" />} label="Docentes" />
        <Tab active={view === 'ausencias'} onClick={() => setView('ausencias')} icon={<UserX className="size-4 shrink-0" />} label="Ausencias" />
        <Tab active={view === 'preceptores'} onClick={() => setView('preceptores')} icon={<UserCheck className="size-4 shrink-0" />} label="Preceptores" />
        <Tab active={view === 'paro'} onClick={() => setView('paro')} icon={<ShieldCheck className="size-4 shrink-0" />} label="Registro de Paro" />
      </nav>
    </header>
    
    <main className="mx-auto max-w-[1500px] px-4 sm:px-5 py-4 sm:py-6">
      <div className="mb-6 print:hidden">
        <div className="flex flex-col sm:flex-row flex-wrap items-start sm:items-center justify-between gap-3">
          <div>
            <p className="mb-1 text-xs font-bold uppercase tracking-[0.15em] text-[#168000]">Panel de control</p>
            <h2 className="text-xl sm:text-2xl font-black tracking-tight">
              {view === 'horario' ? 'Horario semanal y recreos' : view === 'docentes' ? 'Buscar docentes' : view === 'ausencias' ? 'Base de datos de ausencias y cobertura' : view === 'preceptores' ? 'Asignación de preceptores por curso' : 'Historial y control diario de paro'}
            </h2>
            <p className="mt-1 text-xs sm:text-sm text-slate-500">Actualización activa · {syncMessage}</p>
          </div>
          <div className="flex flex-wrap gap-2 w-full sm:w-auto">
            <Button variant="outline" className="flex-1 sm:flex-none" onClick={exportExcel}><FileSpreadsheet data-icon="inline-start" />Descargar Excel</Button>
            <Button className="bg-[#168000] text-white hover:bg-[#106500] flex-1 sm:flex-none" onClick={save}>
              {saved ? <Check data-icon="inline-start" /> : <Cloud data-icon="inline-start" />}
              {saved ? 'Guardado' : dirty ? `Guardar (${dirty})` : 'Guardar cambios'}
            </Button>
          </div>
        </div>
      </div>

      {view === 'horario' && <ScheduleView course={course} setCourse={setCourse} filter={filter} setFilter={setFilter} selectedDays={selectedDays} setSelectedDays={setSelectedDays} visible={visible} blocks={blocks} setBlocks={setBlocks} courseTimes={courseTimes} courses={courses} setEditing={setEditing} onAdd={() => setEditing({ id: crypto.randomUUID(), curso: course, turno: 'Teoría', tipo: 'Teoría', dia: 'Lunes', horario: '07:40–08:40', materia: 'Nueva materia', docente: '', situacionRevista: 'Titular' })} onManageRecreos={() => setShowRecreosModal(true)} />}
      {view === 'docentes' && <TeacherView query={query} setQuery={setQuery} teachers={teachers} results={searchResults} setEditing={setEditing} />}
      {view === 'ausencias' && <AusenciasView ausencias={ausencias} setAusencias={setAusencias} setDirty={setDirty} courses={courses} />}
      {view === 'preceptores' && <PreceptoresView preceptores={preceptores} setPreceptores={setPreceptores} courses={courses} setDirty={setDirty} />}
      {view === 'paro' && <StrikeView dayBlocks={dayBlocks} courses={courses} paroHistorial={paroHistorial} setParoHistorial={setParoHistorial} setDirty={setDirty} />}
    </main>
    
    <footer className="mx-auto flex max-w-[1500px] flex-col sm:flex-row items-center justify-between gap-2 border-t border-[#d7e8c9] px-4 sm:px-5 py-4 text-xs text-slate-500 print:hidden text-center sm:text-left">
      <span>© 2026 E.E.S.T. N° 6 · Sistema de Gestión Escolar</span>
      <span className="flex items-center gap-1.5"><Zap className="size-3 text-[#79b82a]" />Sincronización en vivo con Drive</span>
    </footer>
    {editing && <Editor block={editing} courses={courses} onClose={() => setEditing(null)} onSave={updateBlock} onDelete={() => { setBlocks((bs) => bs.filter((b) => b.id !== editing.id)); setEditing(null); setDirty((n) => n + 1) }} />}
    {showRecreosModal && <RecreosModal course={course} blocks={blocks} setBlocks={setBlocks} onClose={() => setShowRecreosModal(false)} setDirty={setDirty} />}
  </div>
}

function Tab({ active, onClick, icon, label }: { active: boolean; onClick: () => void; icon: React.ReactNode; label: string }) { 
  return <button onClick={onClick} className={`flex items-center gap-2 border-b-2 px-3 py-3 text-sm font-bold transition whitespace-nowrap ${active ? 'border-[#168000] text-[#168000]' : 'border-transparent text-slate-500 hover:text-slate-800'}`}>{icon}<span>{label}</span></button> 
}

function ScheduleView({ course, setCourse, filter, setFilter, selectedDays, setSelectedDays, visible, blocks, setBlocks, courseTimes, courses, setEditing, onAdd, onManageRecreos }: any) { 
  return <>
    <div className="mb-4 flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-3 rounded-2xl border border-[#d7e8c9] bg-white p-3 shadow-sm">
      <label className="flex items-center justify-between sm:justify-start gap-2 text-sm font-bold">Curso 
        <select value={course} onChange={(e) => setCourse(e.target.value)} className="rounded-lg border border-slate-200 bg-white px-3 py-2 font-semibold outline-none focus:ring-2 focus:ring-[#9dcc39]">
          {courses.map((c: string) => <option key={c}>{c}</option>)}
        </select>
      </label>
      <div className="flex flex-wrap items-center gap-1 rounded-lg bg-slate-100 p-1 justify-center">
        {days.map((day) => <button key={day} onClick={() => setSelectedDays((current: string[]) => current.includes(day) ? current.filter((item) => item !== day) : [...current, day])} className={`rounded-md px-2 py-1.5 text-xs font-bold ${selectedDays.includes(day) ? 'bg-white text-[#168000] shadow-sm' : 'text-slate-400'}`}>{day.slice(0, 3)}</button>)}
      </div>
      <div className="flex rounded-lg bg-slate-100 p-1 justify-center">
        {(['Todos', 'Teoría', 'Taller'] as const).map((x) => <button key={x} onClick={() => setFilter(x)} className={`rounded-md px-3 py-1.5 text-xs font-bold ${filter === x ? 'bg-white text-[#168000] shadow-sm' : 'text-slate-500'}`}>{x === 'Todos' ? 'Todos' : `Solo ${x}`}</button>)}
      </div>
      <div className="flex gap-2 sm:ml-auto w-full sm:w-auto">
        <Button variant="outline" className="flex-1 sm:flex-none border-[#ffb700] text-[#866500] hover:bg-[#fff8d8]" onClick={onManageRecreos}><Clock className="size-4 mr-1" /> Recreos</Button>
        <Button className="bg-[#168000] text-white hover:bg-[#106500] flex-1 sm:flex-none" onClick={onAdd}><span className="text-lg leading-none">+</span> Agregar</Button>
      </div>
    </div>
    
    <div className="mb-4 flex flex-col sm:flex-row flex-wrap gap-2 text-xs font-semibold justify-between items-start sm:items-center">
      <div className="flex flex-wrap gap-2">
        <Legend color="bg-blue-600" label="Titular" />
        <Legend color="bg-red-600" label="Provisional" />
        <Legend color="bg-gray-900" label="Titular interino" />
        <Legend color="bg-green-600" label="Suplente" />
      </div>
      <span className="text-slate-500">Doble clic para editar</span>
    </div>

    <div className="overflow-x-auto rounded-2xl border border-[#d7e8c9] bg-white shadow-sm">
      <section className="min-w-[760px]">
        <div className="grid grid-cols-[130px_repeat(5,1fr)] border-b bg-[#f4faed] text-xs font-extrabold uppercase tracking-wider text-[#267018]">
          <div className="p-3">Horario</div>
          {days.map((d) => <div key={d} className="border-l border-[#d7e8c9] p-3">{d}</div>)}
        </div>
        {courseTimes.map((time: string) => {
          const isRecreoRow = blocks.some((x: Block) => x.curso === course && x.horario === time && x.tipo === 'Recreo')
          
          if (isRecreoRow) {
            const recreoBlock = blocks.find((x: Block) => x.curso === course && x.horario === time && x.tipo === 'Recreo')
            return (
              <div key={time} onDoubleClick={() => recreoBlock && setEditing(recreoBlock)} className="grid grid-cols-[130px_1fr] border-b border-slate-100 bg-[#ffb700] text-slate-950 cursor-pointer hover:bg-[#f5ad00] transition">
                <div className="flex flex-col justify-center gap-1 p-2.5 font-mono text-xs font-black uppercase tracking-wider border-r border-[#e0a200]">
                  {time}
                </div>
                <div className="flex items-center justify-between px-4 py-2 text-xs font-black uppercase tracking-[0.3em]">
                  <span>RECREO</span>
                  <span className="text-[10px] font-bold lowercase opacity-75">(Doble clic para editar)</span>
                </div>
              </div>
            )
          }

          return (
            <div key={time} className="grid min-h-[100px] grid-cols-[130px_repeat(5,1fr)] border-b border-slate-100">
              <div className="flex flex-col justify-center gap-1 p-2 sm:p-3">
                <span className="font-mono text-xs sm:text-sm font-bold text-slate-700">{time}</span>
              </div>
              {days.map((day) => { 
                const cellBlocks = visible.filter((x: Block) => x.dia === day && x.horario === time && x.tipo !== 'Recreo')
                return (
                  <div key={day} className="border-l border-slate-100 p-1.5 sm:p-2">
                    <div className={`grid gap-1 h-full ${cellBlocks.length > 1 ? 'grid-cols-2' : 'grid-cols-1'}`}>
                      {cellBlocks.map((b: Block) => (
                        <BlockCard key={b.id} block={b} onDoubleClick={() => setEditing(b)} />
                      ))}
                    </div>
                  </div>
                )
              })}
            </div>
          )
        })}
      </section>
    </div>
  </> 
}

function RecreosModal({ course, blocks, setBlocks, onClose, setDirty }: any) {
  const [nuevoHorario, setNuevoHorario] = useState('09:40–09:50')
  const recreosActuales = blocks.filter((b: Block) => b.curso === course && b.tipo === 'Recreo')

  function agregarRecreo(e: React.FormEvent) {
    e.preventDefault()
    if (!nuevoHorario) return
    const nuevo: Block = {
      id: crypto.randomUUID(), curso: course, turno: 'Teoría', tipo: 'Recreo', dia: 'Lunes', horario: nuevoHorario, materia: 'Recreo', docente: '', situacionRevista: 'Titular'
    }
    setBlocks([...blocks, nuevo])
    setDirty((n: number) => n + 1)
  }

  function eliminarRecreo(id: string) {
    setBlocks(blocks.filter((b: Block) => b.id !== id))
    setDirty((n: number) => n + 1)
  }

  return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/40 p-4" role="dialog" aria-modal="true">
    <div className="w-full max-w-md rounded-2xl bg-white shadow-2xl">
      <div className="flex items-center justify-between border-b p-5">
        <div><p className="text-xs font-bold uppercase tracking-widest text-[#866500]">Recreos</p><h2 className="text-xl font-black">Curso {course}</h2></div>
        <button onClick={onClose} aria-label="Cerrar"><X /></button>
      </div>
      <div className="p-5 flex flex-col gap-4">
        <form onSubmit={agregarRecreo} className="flex gap-2">
          <input value={nuevoHorario} onChange={(e) => setNuevoHorario(e.target.value)} placeholder="Ej. 09:40–09:50" className="flex-1 rounded-xl border border-slate-200 px-3 py-2 text-sm font-mono font-bold" required />
          <Button type="submit" className="bg-[#ffb700] text-slate-950 hover:bg-[#f5ad00] font-bold"><Plus className="size-4 mr-1" /> Sumar</Button>
        </form>
        <div className="border-t pt-3">
          <p className="text-xs font-bold text-slate-500 mb-2">Recreos actuales:</p>
          {recreosActuales.length === 0 ? <p className="text-sm text-slate-400 italic py-2">Sin recreos.</p> : (
            <div className="flex flex-col gap-2 max-h-[250px] overflow-y-auto">
              {recreosActuales.map((r: Block) => (
                <div key={r.id} className="flex items-center justify-between rounded-xl border border-[#ffb700]/50 bg-[#fff8d8] px-3 py-2">
                  <span className="font-mono font-bold text-sm text-[#866500]">{r.horario}</span>
                  <Button variant="ghost" size="sm" onClick={() => eliminarRecreo(r.id)} className="text-red-600 hover:text-red-700 h-8 px-2"><Trash2 className="size-4" /></Button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
      <div className="flex justify-end border-t bg-slate-50 p-4">
        <Button onClick={onClose} className="bg-[#168000] text-white hover:bg-[#106500]">Aceptar</Button>
      </div>
    </div>
  </div>
}

function BlockCard({ block, onDoubleClick }: { block: Block; onDoubleClick: () => void }) { 
  const statusColorClass = block.situacionRevista === 'Titular' ? 'text-blue-700' : block.situacionRevista === 'Provisional' ? 'text-red-700' : 'text-gray-950'

  return <div onDoubleClick={onDoubleClick} className={`group cursor-pointer rounded-xl border p-2 transition hover:-translate-y-0.5 hover:shadow-md flex flex-col justify-between ${block.tipo === 'Teoría' ? 'border-[#9dcc39] bg-[#eaf7d9] text-[#205d19]' : 'border-slate-300 bg-slate-50 text-slate-700'}`}>
    <div>
      <div className="mb-0.5 flex items-start justify-between gap-1">
        <p className="text-[10px] sm:text-[11px] font-extrabold leading-tight">{block.materia}</p>
        {block.grupo && <span className="rounded bg-[#168000] px-1 py-0.5 text-[8px] sm:text-[9px] font-black text-white shrink-0">{block.grupo}</span>}
      </div>
      <div className="flex flex-col text-[9px] sm:text-[10px] font-semibold">
        {block.docente ? <span className={`${statusColorClass} font-bold`}>{block.docente}</span> : <span className="text-slate-400">Sin docente</span>}
      </div>
    </div>
    {block.suplente && (
      <div className="mt-1 rounded border border-green-300 bg-green-50 px-1 py-0.5 text-[9px] font-bold text-green-800">
        Sup: {block.suplente}
      </div>
    )}
  </div> 
}

function Legend({ color, label }: { color: string; label: string }) { 
  return <span className="inline-flex items-center gap-1.5 rounded-full border border-slate-200 bg-white px-2 py-1"><span className={`size-2 rounded-full ${color}`} />{label}</span> 
}

function TeacherView({ query, setQuery, teachers, results, setEditing }: any) { 
  const [status, setStatus] = useState('Todos')
  const filtered = results.filter((b: Block) => status === 'Todos' || (status === 'Suplentes' ? b.suplente?.toLowerCase().includes(query.toLowerCase()) : b.situacionRevista === status))
  const selected = query ? query : 'Todos los docentes' 
  return <div className="grid gap-4 lg:grid-cols-[340px_1fr]">
    <aside className="rounded-2xl border border-[#d7e8c9] bg-white p-4 shadow-sm">
      <p className="mb-2 text-sm font-black">Buscar docente</p>
      <div className="relative">
        <Search className="absolute left-3 top-2.5 size-4 text-slate-400" />
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Nombre o apellido..." className="w-full rounded-xl border border-slate-200 py-2 pl-9 pr-3 text-sm outline-none focus:ring-2 focus:ring-[#9dcc39]" />
      </div>
      <div className="mt-3 flex flex-col gap-1 max-h-[300px] sm:max-h-[400px] overflow-y-auto">
        {teachers.filter((t: string) => t.toLowerCase().includes(query.toLowerCase())).map((t: string) => (
          <button key={t} onClick={() => setQuery(t)} className="flex items-center justify-between rounded-lg px-2 py-2 text-left text-sm font-semibold hover:bg-[#f4faed]">{t}<ChevronDown className="size-3 -rotate-90 text-slate-400" /></button>
        ))}
      </div>
    </aside>
    <section className="flex flex-col gap-4">
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
        <Stat label="Bloques semanales" value={filtered.length} icon={<CalendarDays />} />
        <Stat label="Cursos asignados" value={new Set(filtered.map((b: Block) => b.curso)).size} icon={<Users />} />
        <Stat label="Como titular" value={filtered.filter((b: Block) => b.docente === selected && b.situacionRevista === 'Titular').length} icon={<ShieldCheck />} />
        <Stat label="Como suplente" value={filtered.filter((b: Block) => b.suplente === selected).length} icon={<Check />} />
      </div>
      <div className="rounded-2xl border border-[#d7e8c9] bg-white shadow-sm">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b p-4">
          <div>
            <h3 className="font-black">{selected}</h3>
            <p className="text-xs text-slate-500">{filtered.length} bloques encontrados</p>
          </div>
          <div className="flex flex-wrap gap-1">
            {['Todos', 'Titular', 'Provisional', 'Titular Interino', 'Suplentes'].map((x) => (
              <button key={x} onClick={() => setStatus(x)} className={`rounded-full px-2.5 py-1 text-[11px] font-bold ${status === x ? 'bg-[#168000] text-white' : 'bg-slate-100 text-slate-500'}`}>{x}</button>
            ))}
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[720px] text-left text-sm">
            <thead className="bg-[#f4faed] text-xs uppercase tracking-wider text-[#267018]">
              <tr><th className="p-3">Curso</th><th>Materia</th><th>Día / horario</th><th>Tipo</th><th>Situación</th><th /></tr>
            </thead>
            <tbody>
              {filtered.map((b: Block) => (
                <tr key={b.id} onDoubleClick={() => setEditing(b)} className="border-t border-slate-100 hover:bg-slate-50 cursor-pointer">
                  <td className="p-3 font-bold">{b.curso} {b.grupo && <span className="text-xs font-normal text-slate-500">({b.grupo})</span>}</td>
                  <td>{b.materia}<div className="text-xs text-slate-500">{b.docente}</div></td>
                  <td>{b.dia}<div className="font-mono text-xs text-slate-500">{b.horario}</div></td>
                  <td><span className="rounded-full bg-slate-100 px-2 py-1 text-xs font-bold">{b.tipo}</span></td>
                  <td><span className={`rounded-full border px-2 py-1 text-xs font-bold ${statusStyle[b.situacionRevista]}`}>{b.suplente === selected ? 'Suplente' : b.situacionRevista}</span></td>
                  <td><Pencil className="size-4 text-slate-400" /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  </div> 
}

function AusenciasView({ ausencias, setAusencias, setDirty, courses }: any) {
  const [cursoFiltro, setCursoFiltro] = useState('Todos')
  const [docenteFiltro, setDocenteFiltro] = useState('')
  const [modoFecha, setModoFecha] = useState<'todo' | 'rango'>('todo')
  const [fechaDesde, setFechaDesde] = useState('2026-04-01')
  const [fechaHasta, setFechaHasta] = useState('2026-12-31')

  const [docenteInput, setDocenteInput] = useState('')
  const [materiaInput, setMateriaInput] = useState('')
  const [fechaInput, setFechaInput] = useState(new Date().toISOString().split('T')[0])
  const [estadoCob, setEstadoCob] = useState<CoberturaStatus>('Sin cubrir')
  const [docenteDecreto, setDocenteDecreto] = useState('')

  const ausenciasFiltradas = ausencias.filter((a: AusenciaRecord) => {
    const matchCurso = cursoFiltro === 'Todos' || a.curso === cursoFiltro
    const matchDocente = !docenteFiltro || a.docenteAusente.toLowerCase().includes(docenteFiltro.toLowerCase()) || a.docenteDecreto.toLowerCase().includes(docenteFiltro.toLowerCase())
    
    let matchFecha = true
    if (modoFecha === 'rango' && a.fecha && a.fecha.match(/^\d{4}-\d{2}-\d{2}$/)) {
      matchFecha = a.fecha >= fechaDesde && a.fecha <= fechaHasta
    }
    return matchCurso && matchDocente && matchFecha
  })

  const totalAusencias = ausenciasFiltradas.length
  const cubiertas = ausenciasFiltradas.filter((a: AusenciaRecord) => a.estadoCobertura === 'Cubierto por Decreto').length
  const sinCubrir = totalAusencias - cubiertas

  function registrarAusencia(e: React.FormEvent) {
    e.preventDefault()
    if (!docenteInput || !cursoFiltro) return
    const nueva: AusenciaRecord = {
      id: `AUS-${Date.now().toString().slice(-4)}`,
      curso: cursoFiltro === 'Todos' ? (courses[0] || '1°1°') : cursoFiltro,
      mes: 'Octubre',
      fecha: fechaInput,
      materia: materiaInput || 'Materia general',
      docenteAusente: docenteInput.toUpperCase(),
      estadoCobertura: estadoCob,
      docenteDecreto: estadoCob === 'Cubierto por Decreto' ? (docenteDecreto.toUpperCase() || 'CUBIERTO') : 'Sin cubrir'
    }
    setAusencias([nueva, ...ausencias])
    setDirty((n: number) => n + 1)
    setDocenteInput('')
    setDocenteDecreto('')
  }

  return <div className="flex flex-col gap-6">
    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
      <Stat label="Ausencias en la selección" value={totalAusencias} icon={<UserX />} />
      <Stat label="Cubiertas por decreto" value={cubiertas} icon={<ShieldCheck />} />
      <Stat label="Sin cubrir" value={sinCubrir} icon={<ShieldAlert />} />
    </div>

    <div className="rounded-2xl border border-[#d7e8c9] bg-white p-4 sm:p-5 shadow-sm">
      <h3 className="mb-3 font-black text-base flex items-center gap-2"><Search className="size-4 text-[#168000]" /> Filtrar y buscar ausencias</h3>
      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Filtrar por Curso
          <select value={cursoFiltro} onChange={(e) => setCursoFiltro(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold">
            <option value="Todos">Todos los cursos</option>
            {courses.map((c: string) => <option key={c}>{c}</option>)}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Buscar por Docente (Ausente / Cubre)
          <input value={docenteFiltro} onChange={(e) => setDocenteFiltro(e.target.value)} placeholder="Ej. BOGARIN, ASPREA..." className="rounded-lg border border-slate-200 px-3 py-2 text-sm uppercase" />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Periodo de fechas
          <select value={modoFecha} onChange={(e) => setModoFecha(e.target.value as any)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold">
            <option value="todo">Todo el historial</option>
            <option value="rango">Desde / Hasta específico</option>
          </select>
        </label>
        {modoFecha === 'rango' ? (
          <div className="grid grid-cols-2 gap-2">
            <label className="flex flex-col gap-1 text-[11px] font-bold text-slate-600">Desde <input type="date" value={fechaDesde} onChange={(e) => setFechaDesde(e.target.value)} className="rounded-lg border border-slate-200 px-2 py-2 text-xs" /></label>
            <label className="flex flex-col gap-1 text-[11px] font-bold text-slate-600">Hasta <input type="date" value={fechaHasta} onChange={(e) => setFechaHasta(e.target.value)} className="rounded-lg border border-slate-200 px-2 py-2 text-xs" /></label>
          </div>
        ) : (
          <div className="flex items-end pb-1"><span className="text-xs text-slate-400 font-medium">Historial completo sin límite de fecha.</span></div>
        )}
      </div>
    </div>

    <div className="rounded-2xl border border-[#d7e8c9] bg-white p-4 sm:p-5 shadow-sm">
      <h3 className="mb-4 font-black text-lg">Registrar nueva ausencia docente</h3>
      <form onSubmit={registrarAusencia} className="grid gap-4 sm:grid-cols-2 lg:grid-cols-6">
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Curso
          <select value={cursoFiltro === 'Todos' ? (courses[0] || '1°1°') : cursoFiltro} onChange={(e) => setCursoFiltro(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm">
            {courses.map((c: string) => <option key={c}>{c}</option>)}
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Fecha
          <input type="date" value={fechaInput} onChange={(e) => setFechaInput(e.target.value)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm" />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Docente ausente
          <input value={docenteInput} onChange={(e) => setDocenteInput(e.target.value)} placeholder="Ej. BOGARIN" className="rounded-lg border border-slate-200 px-3 py-2 text-sm uppercase" required />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Materia
          <input value={materiaInput} onChange={(e) => setMateriaInput(e.target.value)} placeholder="Ej. Matemática" className="rounded-lg border border-slate-200 px-3 py-2 text-sm" />
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Estado de Cobertura
          <select value={estadoCob} onChange={(e) => setEstadoCob(e.target.value as CoberturaStatus)} className="rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold">
            <option value="Sin cubrir">Sin cubrir</option>
            <option value="Cubierto por Decreto">Cubierto por Decreto</option>
          </select>
        </label>
        <label className="flex flex-col gap-1 text-xs font-bold text-slate-600">Profesor que cubre / Obs.
          <input value={docenteDecreto} onChange={(e) => setDocenteDecreto(e.target.value)} placeholder="Ej. ASPREA" disabled={estadoCob === 'Sin cubrir'} className="rounded-lg border border-slate-200 px-3 py-2 text-sm uppercase disabled:bg-slate-100" />
        </label>
        <div className="sm:col-span-2 lg:col-span-6 flex justify-end">
          <Button type="submit" className="bg-[#168000] text-white hover:bg-[#106500] w-full sm:w-auto">Registrar ausencia</Button>
        </div>
      </form>
    </div>

    <div className="overflow-hidden rounded-2xl border border-[#d7e8c9] bg-white shadow-sm">
      <div className="border-b p-4 flex items-center justify-between">
        <div>
          <h3 className="font-black text-lg">Historial de ausencias y cobertura</h3>
          <p className="text-xs text-slate-500">Mostrando {ausenciasFiltradas.length} registros según los filtros seleccionados</p>
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full min-w-[850px] text-left text-sm">
          <thead className="bg-[#f4faed] text-xs uppercase tracking-wider text-[#267018]">
            <tr><th className="p-3">Curso</th><th>Fecha</th><th>Docente ausente</th><th>Materia</th><th>Estado</th><th>Cubre por decreto / Observación</th></tr>
          </thead>
          <tbody>
            {ausenciasFiltradas.length === 0 ? (
              <tr><td colSpan={6} className="p-6 text-center text-slate-400 font-semibold">No se encontraron ausencias con los filtros seleccionados.</td></tr>
            ) : (
              ausenciasFiltradas.map((a: AusenciaRecord) => (
                <tr key={a.id} className="border-t border-slate-100 hover:bg-slate-50">
                  <td className="p-3 font-bold">{a.curso}</td>
                  <td>{a.fecha}</td>
                  <td className="font-bold text-slate-800">{a.docenteAusente}</td>
                  <td>{a.materia}</td>
                  <td>
                    <span className={`rounded-full px-2.5 py-1 text-xs font-bold ${a.estadoCobertura === 'Cubierto por Decreto' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}`}>
                      {a.estadoCobertura}
                    </span>
                  </td>
                  <td className="font-semibold text-slate-700">{a.docenteDecreto}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  </div>
}

function PreceptoresView({ preceptores, setPreceptores, courses, setDirty }: any) {
  const [filtroCurso, setFiltroCurso] = useState('')

  function actualizarPreceptor(curso: string, campo: 'preceptorTitular' | 'preceptorSuplente', valor: string) {
    setPreceptores(preceptores.map((p: PreceptorAssignment) => p.curso === curso ? { ...p, [campo]: valor.toUpperCase() } : p))
    setDirty((n: number) => n + 1)
  }

  const filtrados = preceptores.filter((p: PreceptorAssignment) => !filtroCurso || p.curso.toLowerCase().includes(filtroCurso.toLowerCase()) || p.preceptorTitular.toLowerCase().includes(filtroCurso.toLowerCase()) || p.preceptorSuplente.toLowerCase().includes(filtroCurso.toLowerCase()))

  return <div className="flex flex-col gap-6">
    <div className="rounded-2xl border border-[#d7e8c9] bg-white p-4 sm:p-5 shadow-sm">
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-4">
        <div>
          <h3 className="font-black text-lg">Asignación de Preceptores por Curso</h3>
          <p className="text-xs text-slate-500">Configurá el preceptor titular y suplente asignado a cada división de la escuela</p>
        </div>
        <div className="w-full sm:w-auto">
          <input value={filtroCurso} onChange={(e) => setFiltroCurso(e.target.value)} placeholder="Filtrar curso o preceptor..." className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm w-full sm:w-64" />
        </div>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full min-w-[600px] text-left text-sm">
          <thead className="bg-[#f4faed] text-xs uppercase tracking-wider text-[#267018]">
            <tr><th className="p-3">Curso / División</th><th>Preceptor Titular</th><th>Preceptor Suplente</th></tr>
          </thead>
          <tbody>
            {filtrados.map((p: PreceptorAssignment) => (
              <tr key={p.curso} className="border-t border-slate-100 hover:bg-slate-50">
                <td className="p-3 font-bold text-base text-[#168000]">{p.curso}</td>
                <td className="p-2">
                  <input value={p.preceptorTitular} onChange={(e) => actualizarPreceptor(p.curso, 'preceptorTitular', e.target.value)} placeholder="Nombre del preceptor titular" className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm uppercase font-semibold" />
                </td>
                <td className="p-2">
                  <input value={p.preceptorSuplente} onChange={(e) => actualizarPreceptor(p.curso, 'preceptorSuplente', e.target.value)} placeholder="Opcional (Suplente)" className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm uppercase font-semibold" />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  </div>
}

function StrikeView({ dayBlocks, courses, paroHistorial, setParoHistorial, setDirty }: any) { 
  const [date, setDate] = useState(new Date().toISOString().split('T')[0])
  const [cursoFiltro, setCursoFiltro] = useState('Todos')
  
  const paroDelDia = paroHistorial.filter((p: ParoRecord) => p.fecha === date)

  function cambiarEstado(b: any, estado: Attendance, obs?: string) {
    const existingIndex = paroHistorial.findIndex((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia)
    const nuevoRegistro: ParoRecord = {
      id: existingIndex >= 0 ? paroHistorial[existingIndex].id : `PARO-${Date.now().toString().slice(-4)}`,
      fecha: date,
      curso: b.curso,
      horario: b.horario,
      materia: b.materia,
      docente: b.suplente || b.docente,
      estadoParo: estado,
      observaciones: obs !== undefined ? obs : (existingIndex >= 0 ? paroHistorial[existingIndex].observaciones : '')
    }

    if (existingIndex >= 0) {
      const copy = [...paroHistorial]
      copy[existingIndex] = nuevoRegistro
      setParoHistorial(copy)
    } else {
      setParoHistorial([nuevoRegistro, ...paroHistorial])
    }
    setDirty((n: number) => n + 1)
  }

  const bloquesFiltrados = cursoFiltro === 'Todos' ? dayBlocks : dayBlocks.filter((b: Block) => b.curso === cursoFiltro)

  const totalRelevados = bloquesFiltrados.length
  const totalParo = bloquesFiltrados.filter((b: Block) => {
    const reg = paroHistorial.find((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia)
    return reg?.estadoParo === 'paro'
  }).length

  return <div className="flex flex-col gap-6">
    <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-[#d7e8c9] bg-white p-4 shadow-sm">
      <label className="text-sm font-bold">Fecha del Relevamiento <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="ml-2 rounded-lg border border-slate-200 px-3 py-2 text-sm font-semibold" /></label>
      <label className="flex items-center gap-2 text-sm font-bold sm:ml-auto">Curso 
        <select value={cursoFiltro} onChange={(e) => setCursoFiltro(e.target.value)} className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-semibold">
          <option value="Todos">Todos los cursos</option>
          {courses.map((c: string) => <option key={c}>{c}</option>)}
        </select>
      </label>
    </div>
    
    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
      <Stat label="Bloques en la fecha" value={totalRelevados} icon={<ListFilter />} />
      <Stat label="Adhesiones al paro registradas" value={totalParo} icon={<ShieldCheck />} />
      <Stat label="% acatamiento del día" value={totalRelevados ? Math.round(totalParo / totalRelevados * 100) : 0} icon={<Zap />} />
    </div>

    <div className="overflow-hidden rounded-2xl border border-[#d7e8c9] bg-white shadow-sm">
      <div className="flex items-center justify-between border-b p-4">
        <div>
          <h3 className="font-black text-lg">Registro de Paro · {date}</h3>
          <p className="text-xs text-slate-500">Los cambios quedan guardados de manera persistente al apretar Guardar en Google Drive</p>
        </div>
        <Button variant="outline" size="sm" onClick={() => window.print()}><Printer data-icon="inline-start" />Imprimir parte</Button>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full min-w-[850px] text-left text-sm">
          <thead className="bg-[#f4faed] text-xs uppercase tracking-wider text-[#267018]">
            <tr><th className="p-3">Curso / horario</th><th>Materia</th><th>Docente activo</th><th>Estado (¿Hizo Paro?)</th><th>Observaciones</th></tr>
          </thead>
          <tbody>
            {bloquesFiltrados.map((b: Block) => {
              const reg = paroHistorial.find((p: ParoRecord) => p.fecha === date && p.curso === b.curso && p.horario === b.horario && p.materia === b.materia)
              const estadoActual = reg?.estadoParo || 'normal'
              const obsActual = reg?.observaciones || ''

              return <tr key={b.id + b.curso} className="border-t border-slate-100 hover:bg-slate-50">
                <td className="p-3 font-bold">{b.curso} {b.grupo && <span className="text-xs font-normal text-slate-500">({b.grupo})</span>}<div className="font-mono text-xs font-normal text-slate-500">{b.horario}</div></td>
                <td>{b.materia}</td>
                <td>
                  <span className={b.suplente ? 'font-bold text-green-700' : 'font-semibold'}>{b.suplente || b.docente}</span>
                  {b.suplente && <div className="text-xs text-slate-500">Titular: {b.docente}</div>}
                </td>
                <td>
                  <div className="flex gap-1">
                    <button type="button" onClick={() => cambiarEstado(b, 'normal')} className={`rounded-full px-2.5 py-1 text-xs font-bold ${estadoActual === 'normal' ? 'bg-green-100 text-green-700 ring-1 ring-green-400' : 'bg-slate-100 text-slate-500'}`}>Normal</button>
                    <button type="button" onClick={() => cambiarEstado(b, 'paro')} className={`rounded-full px-2.5 py-1 text-xs font-bold ${estadoActual === 'paro' ? 'bg-red-100 text-red-700 ring-1 ring-red-400' : 'bg-slate-100 text-slate-500'}`}>Hizo Paro</button>
                    <button type="button" onClick={() => cambiarEstado(b, 'licencia')} className={`rounded-full px-2.5 py-1 text-xs font-bold ${estadoActual === 'licencia' ? 'bg-amber-100 text-amber-700 ring-1 ring-amber-400' : 'bg-slate-100 text-slate-500'}`}>Licencia</button>
                  </div>
                </td>
                <td><input value={obsActual} onChange={(e) => cambiarEstado(b, estadoActual, e.target.value)} placeholder="Observación..." className="w-full min-w-40 rounded-lg border border-slate-200 px-2 py-1.5 text-xs outline-none focus:ring-1 focus:ring-[#9dcc39]" /></td>
              </tr>
            })}
          </tbody>
        </table>
      </div>
    </div>
  </div> 
}

function Stat({ label, value, icon }: { label: string; value: number; icon: React.ReactNode }) { 
  return <div className="rounded-2xl border border-[#d7e8c9] bg-white p-3 sm:p-4 shadow-sm"><div className="mb-2 sm:mb-3 flex items-center justify-between text-[#168000]">{icon}<span className="text-xl sm:text-2xl font-black text-slate-900">{value}</span></div><p className="text-[11px] sm:text-xs font-bold text-slate-500">{label}</p></div> 
}

function Editor({ block, courses, onClose, onSave, onDelete }: { block: Block; courses: string[]; onClose: () => void; onSave: (b: Block) => void; onDelete: () => void }) { 
  const [draft, setDraft] = useState(block) 
  const set = (key: keyof Block, value: string) => setDraft((d) => ({ ...d, [key]: value })) 
  return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/40 p-4" role="dialog" aria-modal="true">
    <div className="w-full max-w-lg rounded-2xl bg-white shadow-2xl">
      <div className="flex items-center justify-between border-b p-5">
        <div><p className="text-xs font-bold uppercase tracking-widest text-[#168000]">Edición rápida</p><h2 className="text-xl font-black">{draft.materia}</h2></div>
        <button onClick={onClose} aria-label="Cerrar"><X /></button>
      </div>
      <div className="grid gap-4 p-5 sm:grid-cols-2">
        <Field label="Curso"><select value={draft.curso} onChange={(e) => set('curso', e.target.value)}>{courses.map((c: string) => <option key={c}>{c}</option>)}</select></Field>
        <Field label="Materia"><input value={draft.materia} onChange={(e) => set('materia', e.target.value)} /></Field>
        <Field label="Docente"><input value={draft.docente} onChange={(e) => set('docente', e.target.value)} /></Field>
        <Field label="Día"><select value={draft.dia} onChange={(e) => set('dia', e.target.value)}>{days.map((d) => <option key={d}>{d}</option>)}</select></Field>
        <Field label="Horario"><input value={draft.horario} onChange={(e) => set('horario', e.target.value)} placeholder="Ej. 11:50–12:50 o 09:40–09:50" /></Field>
        <Field label="Tipo de horario"><select value={draft.tipo} onChange={(e) => { const tipo = e.target.value as Block['tipo']; setDraft((d) => ({ ...d, tipo, turno: tipo === 'Taller' ? 'Taller' : 'Teoría', materia: tipo === 'Recreo' ? 'Recreo' : d.materia })) }}><option>Teoría</option><option>Taller</option><option>Recreo</option></select></Field>
        <Field label="Grupo"><select value={draft.grupo ?? ''} onChange={(e) => set('grupo', e.target.value)}><option value="">Curso completo</option><option>Grupo A</option><option>Grupo B</option><option>Ambos</option></select></Field>
        <Field label="Situación de revista"><select value={draft.situacionRevista} onChange={(e) => set('situacionRevista', e.target.value)}><option>Titular</option><option>Provisional</option><option>Titular Interino</option></select></Field>
        <Field label="Suplente" wide><input value={draft.suplente ?? ''} onChange={(e) => set('suplente', e.target.value)} placeholder="Opcional" /></Field>
      </div>
      <div className="flex items-center justify-between border-t bg-slate-50 p-4">
        <Button variant="ghost" className="text-red-600 hover:text-red-700" onClick={onDelete}>Eliminar bloque</Button>
        <div className="flex gap-2">
          <Button variant="outline" onClick={onClose}>Cancelar</Button>
          <Button className="bg-[#168000] text-white hover:bg-[#106500]" onClick={() => onSave(draft)}>Guardar bloque</Button>
        </div>
      </div>
    </div>
  </div> 
}

function Field({ label, children, wide }: { label: string; children: React.ReactNode; wide?: boolean }) { 
  return <label className={`flex flex-col gap-1.5 text-xs font-bold text-slate-600 ${wide ? 'sm:col-span-2' : ''}`}>{label}{children}</label> 
}
