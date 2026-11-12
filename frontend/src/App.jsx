import { useState, useRef, useEffect } from 'react'
import DropZone from './components/DropZone'
import FileList from './components/FileList'
import ProgressCard from './components/ProgressCard'
import ResultsTable from './components/ResultsTable'
import './index.css'

const POLL_INTERVAL = 1500

export default function App() {
  const [files, setFiles] = useState([])
  const [phase, setPhase] = useState('idle')
  const [jobId, setJobId] = useState(null)
  const [progress, setProgress] = useState(0)
  const [results, setResults] = useState([])
  const [errorMsg, setErrorMsg] = useState('')
  const [ocrAvailable, setOcrAvailable] = useState(null)
  const pollRef = useRef(null)

  useEffect(() => {
    fetch('/api/health').then(r => r.json()).then(d => setOcrAvailable(d.ocr_available)).catch(() => setOcrAvailable(false))
  }, [])

  const addFiles = newFiles => setFiles(prev => {
    const names = new Set(prev.map(f => f.name))
    return [...prev, ...newFiles.filter(f => !names.has(f.name))]
  })

  const removeFile = idx => setFiles(prev => prev.filter((_, i) => i !== idx))

  const reset = () => { clearInterval(pollRef.current); setFiles([]); setPhase('idle'); setJobId(null); setProgress(0); setResults([]); setErrorMsg('') }

  const startPolling = id => {
    clearInterval(pollRef.current)
    pollRef.current = setInterval(async () => {
      try {
        const data = await fetch(`/api/status/${id}`).then(r => r.json())
        setProgress(data.progress)
        if (data.status === 'done') { clearInterval(pollRef.current); setResults(data.results); setPhase('done') }
        else if (data.status === 'error') { clearInterval(pollRef.current); setErrorMsg(data.error || 'Erro desconhecido.'); setPhase('error') }
      } catch { clearInterval(pollRef.current); setErrorMsg('Erro de conexão com o servidor.'); setPhase('error') }
    }, POLL_INTERVAL)
  }

  const handleProcess = async () => {
    if (!files.length) return
    setPhase('uploading'); setProgress(0); setErrorMsg(''); setResults([])
    const formData = new FormData()
    files.forEach(f => formData.append('files', f))
    try {
      const res = await fetch('/api/process', { method: 'POST', body: formData })
      if (!res.ok) { const e = await res.json().catch(() => ({ detail: res.statusText })); throw new Error(e.detail) }
      const { job_id } = await res.json()
      setJobId(job_id); setPhase('processing'); setProgress(5); startPolling(job_id)
    } catch (e) { setErrorMsg(e.message); setPhase('error') }
  }

  const handleDownload = () => { const a = document.createElement('a'); a.href = `/api/download/${jobId}`; a.download = 'DS_Unificadas.zip'; a.click() }

  const busy = phase === 'uploading' || phase === 'processing'

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      <header className="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-10">
        <div className="max-w-4xl mx-auto px-4 py-3 flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-blue-600 flex items-center justify-center">
            <svg className="w-5 h-5 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-bold text-slate-800 leading-tight">DS Unificador OCR</h1>
            <p className="text-xs text-slate-500 hidden sm:block">Unificação automática de Declarações de Serviço por cliente</p>
          </div>
          {ocrAvailable !== null && (
            <div className={`ml-auto flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-full ${ocrAvailable ? 'bg-green-100 text-green-700' : 'bg-amber-100 text-amber-700'}`}>
              <span className={`w-2 h-2 rounded-full ${ocrAvailable ? 'bg-green-500' : 'bg-amber-500'}`} />
              OCR {ocrAvailable ? 'ativo' : 'indisponível'}
            </div>
          )}
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 py-8 space-y-6">
        <div className="bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 flex gap-3 text-sm text-blue-800">
          <svg className="w-5 h-5 text-blue-500 flex-shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M11.25 11.25l.041-.02a.75.75 0 011.063.852l-.708 2.836a.75.75 0 001.063.853l.041-.021M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-9-3.75h.008v.008H12V8.25z"/></svg>
          <p>Envie PDFs de DS ou um ZIP. O sistema extrai código, nome e competência (mês/ano), une as páginas de cada cliente em um único PDF e gera um ZIP renomeado automaticamente.</p>
        </div>

        {phase === 'idle' && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-4">
            <h2 className="text-base font-semibold text-slate-700">Selecionar arquivos</h2>
            <DropZone onFiles={addFiles} disabled={busy} />
            <FileList files={files} onRemove={removeFile} disabled={busy} />
            {files.length > 0 && (
              <button onClick={handleProcess} className="w-full py-3 px-6 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-xl transition-colors shadow-sm flex items-center justify-center gap-2">
                <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M5 3l14 9-14 9V3z"/></svg>
                Processar {files.length} arquivo{files.length !== 1 ? 's' : ''}
              </button>
            )}
          </div>
        )}

        {busy && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-4">
            <ProgressCard status={phase === 'uploading' ? 'pending' : 'processing'} progress={progress} />
            <p className="text-center text-sm text-slate-500">{phase === 'uploading' ? 'Enviando arquivos...' : 'Analisando PDFs e unificando por cliente...'}</p>
          </div>
        )}

        {phase === 'done' && (
          <div className="space-y-4">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6"><ProgressCard status="done" progress={100} /></div>
            <div className="bg-green-50 border border-green-200 rounded-2xl px-5 py-4 flex flex-col sm:flex-row items-start sm:items-center gap-4">
              <div className="w-12 h-12 bg-green-100 rounded-xl flex items-center justify-center flex-shrink-0">
                <svg className="w-7 h-7 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
              </div>
              <div className="flex-1">
                <p className="font-semibold text-green-800">{results.length} PDF{results.length !== 1 ? 's' : ''} gerado{results.length !== 1 ? 's' : ''} com sucesso</p>
                <p className="text-sm text-green-700 mt-0.5">Clique em "Baixar ZIP" para fazer o download.</p>
              </div>
              <button onClick={handleDownload} className="flex items-center gap-2 px-5 py-2.5 bg-green-600 hover:bg-green-700 text-white font-semibold rounded-xl transition-colors shadow-sm whitespace-nowrap">
                <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3"/></svg>
                Baixar ZIP
              </button>
            </div>
            <ResultsTable results={results} />
            <button onClick={reset} className="w-full py-2.5 border border-slate-300 hover:border-slate-400 text-slate-600 font-medium rounded-xl transition-colors text-sm">Processar novos arquivos</button>
          </div>
        )}

        {phase === 'error' && (
          <div className="space-y-4">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6"><ProgressCard status="error" progress={progress} /></div>
            <div className="bg-red-50 border border-red-200 rounded-xl px-5 py-4">
              <p className="font-semibold text-red-800 mb-1">Erro no processamento</p>
              <p className="text-sm text-red-700 font-mono break-all">{errorMsg}</p>
            </div>
            <button onClick={reset} className="w-full py-2.5 border border-slate-300 hover:border-slate-400 text-slate-600 font-medium rounded-xl transition-colors text-sm">Tentar novamente</button>
          </div>
        )}
      </main>

      <footer className="max-w-4xl mx-auto px-4 py-6 text-center text-xs text-slate-400 border-t border-slate-200 mt-8">
        DS Unificador OCR — Processamento local, seus arquivos não saem do servidor.
      </footer>
    </div>
  )
}
