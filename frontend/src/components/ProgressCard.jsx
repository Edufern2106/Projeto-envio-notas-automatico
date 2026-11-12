export default function ProgressCard({ status, progress }) {
  const labels = { pending: 'Aguardando...', processing: 'Processando PDFs...', done: 'Concluído!', error: 'Erro no processamento' }
  const isError = status === 'error', isDone = status === 'done'
  return (
    <div className={`rounded-2xl border p-6 shadow-sm ${isError ? 'bg-red-50 border-red-200' : 'bg-white border-slate-200'}`}>
      <div className="flex items-center gap-3 mb-4">
        {!isDone && !isError && <svg className="w-6 h-6 text-blue-500 animate-spin flex-shrink-0" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>}
        {isDone && <div className="w-6 h-6 rounded-full bg-green-100 flex items-center justify-center"><svg className="w-4 h-4 text-green-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M4.5 12.75l6 6 9-13.5"/></svg></div>}
        {isError && <div className="w-6 h-6 rounded-full bg-red-100 flex items-center justify-center"><svg className="w-4 h-4 text-red-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12"/></svg></div>}
        <span className={`font-semibold ${isError ? 'text-red-700' : isDone ? 'text-green-700' : 'text-slate-700'}`}>{labels[status]}</span>
        <span className="ml-auto text-sm font-medium text-slate-500">{progress}%</span>
      </div>
      <div className="h-2.5 bg-slate-100 rounded-full overflow-hidden">
        <div className={`h-full rounded-full transition-all duration-500 ${isError ? 'bg-red-400' : isDone ? 'bg-green-500' : 'bg-blue-500'}`} style={{width:`${progress}%`}}/>
      </div>
    </div>
  )
}
