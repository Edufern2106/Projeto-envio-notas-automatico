import { useRef, useState } from 'react'

export default function DropZone({ onFiles, disabled }) {
  const inputRef = useRef(null)
  const [dragging, setDragging] = useState(false)

  const isValid = f => f.name.toLowerCase().endsWith('.pdf') || f.name.toLowerCase().endsWith('.zip')

  return (
    <div
      onDragOver={e => { e.preventDefault(); if (!disabled) setDragging(true) }}
      onDragLeave={() => setDragging(false)}
      onDrop={e => { e.preventDefault(); setDragging(false); if (!disabled) { const f = Array.from(e.dataTransfer.files).filter(isValid); if (f.length) onFiles(f) } }}
      onClick={() => !disabled && inputRef.current?.click()}
      className={['border-2 border-dashed rounded-2xl p-12 flex flex-col items-center gap-4 cursor-pointer transition-all',
        dragging ? 'border-blue-500 bg-blue-50' : 'border-slate-300 bg-white hover:border-blue-400 hover:bg-blue-50/30',
        disabled ? 'opacity-50 cursor-not-allowed' : ''].join(' ')}
    >
      <input ref={inputRef} type="file" multiple accept=".pdf,.zip" className="hidden"
        onChange={e => { const f = Array.from(e.target.files).filter(isValid); if (f.length) onFiles(f); e.target.value = '' }} />
      <svg className="w-14 h-14 text-blue-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
      </svg>
      <div className="text-center">
        <p className="text-lg font-semibold text-slate-700">Arraste arquivos PDF ou ZIP aqui</p>
        <p className="text-sm text-slate-500 mt-1">ou clique para selecionar arquivos</p>
      </div>
      <div className="flex gap-2 mt-1">
        <span className="px-3 py-1 bg-blue-100 text-blue-700 text-xs font-medium rounded-full">.PDF</span>
        <span className="px-3 py-1 bg-purple-100 text-purple-700 text-xs font-medium rounded-full">.ZIP</span>
      </div>
    </div>
  )
}
