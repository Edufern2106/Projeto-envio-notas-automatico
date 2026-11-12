const fmtSize = b => b < 1024 ? `${b} B` : b < 1048576 ? `${(b/1024).toFixed(1)} KB` : `${(b/1048576).toFixed(1)} MB`

export default function FileList({ files, onRemove, disabled }) {
  if (!files.length) return null
  return (
    <div className="mt-4 space-y-2">
      <p className="text-sm font-medium text-slate-600">{files.length} arquivo(s) selecionado(s):</p>
      {files.map((f, i) => (
        <div key={i} className="flex items-center gap-3 bg-white border border-slate-200 rounded-xl px-4 py-2.5 shadow-sm">
          <svg className={`w-5 h-5 flex-shrink-0 ${f.name.endsWith('.zip') ? 'text-purple-500' : 'text-red-500'}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
          </svg>
          <div className="flex-1 min-w-0">
            <p className="text-sm font-medium text-slate-700 truncate">{f.name}</p>
            <p className="text-xs text-slate-400">{fmtSize(f.size)}</p>
          </div>
          {!disabled && (
            <button onClick={() => onRemove(i)} className="p-1 rounded-lg hover:bg-red-50 text-slate-400 hover:text-red-500 transition-colors">
              <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          )}
        </div>
      ))}
    </div>
  )
}
