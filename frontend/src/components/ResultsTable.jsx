const Badge = ({ children, color }) => {
  const c = { blue: 'bg-blue-100 text-blue-700', green: 'bg-green-100 text-green-700' }
  return <span className={`inline-block px-2.5 py-0.5 rounded-full text-xs font-medium ${c[color]}`}>{children}</span>
}

export default function ResultsTable({ results }) {
  if (!results.length) return null
  return (
    <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
      <div className="px-5 py-4 border-b border-slate-100 flex items-center gap-2">
        <svg className="w-5 h-5 text-green-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
        <span className="font-semibold text-slate-700">PDFs gerados ({results.length})</span>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="bg-slate-50 text-left">
            <th className="px-5 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wide">Arquivo</th>
            <th className="px-5 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wide">Código</th>
            <th className="px-5 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wide">Competência</th>
            <th className="px-5 py-3 text-xs font-semibold text-slate-500 uppercase tracking-wide text-center">Páginas</th>
          </tr></thead>
          <tbody className="divide-y divide-slate-100">
            {results.map((r, i) => (
              <tr key={i} className="hover:bg-slate-50 transition-colors">
                <td className="px-5 py-3 font-medium text-slate-700 max-w-xs"><span className="truncate block" title={r.filename}>{r.filename}</span></td>
                <td className="px-5 py-3"><Badge color="blue">{r.code}</Badge></td>
                <td className="px-5 py-3 text-slate-600">{r.month_name} {r.year}</td>
                <td className="px-5 py-3 text-center"><Badge color="green">{r.page_count}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}
