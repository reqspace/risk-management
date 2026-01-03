import { useState, useRef } from 'react'
import { useParams } from 'react-router-dom'
import { useApi, API_BASE } from '../hooks/useApi'
import html2pdf from 'html2pdf.js'
import {
  FileText,
  Download,
  Mail,
  AlertTriangle,
  CheckSquare,
  TrendingUp,
  TrendingDown,
  Minus,
  Loader2,
  Calendar,
  Clock,
  CalendarDays,
  CalendarRange,
  Send,
  Paperclip
} from 'lucide-react'


function TrendIcon({ trend }) {
  if (trend === 'Increasing' || trend === 'Up') {
    return <TrendingUp className="h-4 w-4 text-red-500" />
  }
  if (trend === 'Decreasing' || trend === 'Down') {
    return <TrendingDown className="h-4 w-4 text-green-500" />
  }
  return <Minus className="h-4 w-4 text-gray-400" />
}


function ReportTypeCard({ type, icon: Icon, title, description, selected, onClick }) {
  return (
    <button
      onClick={onClick}
      className={`p-4 rounded-xl border-2 text-left transition-all ${
        selected
          ? 'border-slate-800 bg-slate-50 shadow-md'
          : 'border-gray-200 bg-white hover:border-slate-400'
      }`}
    >
      <div className="flex items-start gap-3">
        <div className={`p-2 rounded-lg ${selected ? 'bg-slate-800 text-white' : 'bg-gray-100 text-gray-600'}`}>
          <Icon className="h-5 w-5" />
        </div>
        <div>
          <h3 className={`font-semibold ${selected ? 'text-slate-800' : 'text-gray-700'}`}>{title}</h3>
          <p className="text-sm text-gray-500 mt-1">{description}</p>
        </div>
      </div>
    </button>
  )
}


export default function Reports() {
  const { projectCode } = useParams()
  const project = projectCode || 'RH'
  const reportRef = useRef(null)

  const { data: statsData } = useApi(`/api/stats?project=${project}`)
  const { data: risksData } = useApi(`/api/risks?project=${project}`)
  const { data: tasksData } = useApi(`/api/tasks?project=${project}`)
  const { data: portfolioData } = useApi('/api/portfolio')
  const { data: milestonesData } = useApi(`/api/milestones?project=${project}`)

  const [reportType, setReportType] = useState('current')
  const [email, setEmail] = useState('')
  const [includeAttachments, setIncludeAttachments] = useState(true)
  const [exporting, setExporting] = useState(false)
  const [sending, setSending] = useState(false)
  const [downloading, setDownloading] = useState(false)
  const [message, setMessage] = useState(null)

  const stats = statsData?.stats || {}
  const risks = risksData?.risks || []
  const tasks = tasksData?.tasks || []
  const portfolio = portfolioData?.projects || []
  const milestones = milestonesData?.milestones || []

  const activeRisks = risks.filter(r => ['Open', 'Active', 'Escalated'].includes(r.Status))
  const highPriorityRisks = risks.filter(r =>
    (r.Probability === 'High' && r.Impact === 'High') || r.Impact === 'Critical'
  ).filter(r => r.Status !== 'Closed')
  const openTasks = tasks.filter(t => !['Completed', 'Done', 'Complete'].includes(t.Status))
  const criticalMilestones = milestones.filter(m => ['Critical', 'At Risk'].includes(m.Status))

  const today = new Date().toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  })

  const reportTypes = [
    {
      type: 'current',
      icon: Clock,
      title: 'Current Status',
      description: 'Real-time snapshot of risks and tasks'
    },
    {
      type: 'daily',
      icon: Calendar,
      title: 'Daily Digest',
      description: 'Daily summary with portfolio overview'
    },
    {
      type: 'weekly',
      icon: CalendarDays,
      title: 'Weekly Summary',
      description: 'Week-over-week changes and trends'
    },
    {
      type: 'monthly',
      icon: CalendarRange,
      title: 'Monthly Report',
      description: 'Executive summary with full details'
    },
  ]

  const handleExportPDF = async () => {
    if (!reportRef.current) return

    setExporting(true)
    setMessage(null)

    try {
      const element = reportRef.current
      const opt = {
        margin: [10, 10],
        filename: `Risk_Report_${project}_${reportType}_${new Date().toISOString().split('T')[0]}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
      }

      await html2pdf().set(opt).from(element).save()
      setMessage({ type: 'success', text: 'PDF exported successfully!' })
    } catch (err) {
      setMessage({ type: 'error', text: 'Failed to export PDF: ' + err.message })
    } finally {
      setExporting(false)
      setTimeout(() => setMessage(null), 3000)
    }
  }

  const handleDownloadServerPDF = async () => {
    setDownloading(true)
    setMessage(null)

    try {
      const response = await fetch(`${API_BASE}/api/reports/pdf?project=${project}&type=${reportType}`)
      if (!response.ok) throw new Error('Failed to generate PDF')

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `${reportType}_Report_${project}_${new Date().toISOString().split('T')[0]}.pdf`
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)

      setMessage({ type: 'success', text: 'PDF downloaded successfully!' })
    } catch (err) {
      setMessage({ type: 'error', text: 'Failed to download PDF: ' + err.message })
    } finally {
      setDownloading(false)
      setTimeout(() => setMessage(null), 3000)
    }
  }

  const handleSendEmail = async () => {
    if (!email.trim()) {
      setMessage({ type: 'error', text: 'Please enter an email address' })
      setTimeout(() => setMessage(null), 3000)
      return
    }

    setSending(true)
    setMessage(null)

    try {
      const response = await fetch(`${API_BASE}/api/reports/send`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          project,
          type: reportType,
          email: email.trim(),
          include_attachments: includeAttachments
        })
      })

      const data = await response.json()

      if (data.success) {
        setMessage({ type: 'success', text: `Report sent to ${email}!` })
        setEmail('')
      } else {
        setMessage({ type: 'error', text: data.error || 'Failed to send email' })
      }
    } catch (err) {
      setMessage({ type: 'error', text: 'Failed to send email: ' + err.message })
    } finally {
      setSending(false)
      setTimeout(() => setMessage(null), 5000)
    }
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <h2 className="text-2xl font-bold text-gray-900">Reports</h2>
      </div>

      {message && (
        <div className={`p-4 rounded-lg ${
          message.type === 'success' ? 'bg-green-50 text-green-800 border border-green-200' :
          message.type === 'error' ? 'bg-red-50 text-red-800 border border-red-200' :
          'bg-blue-50 text-blue-800 border border-blue-200'
        }`}>
          {message.text}
        </div>
      )}

      {/* Report Type Selection */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
        <h3 className="text-lg font-semibold text-gray-900 mb-4">Select Report Type</h3>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {reportTypes.map(rt => (
            <ReportTypeCard
              key={rt.type}
              {...rt}
              selected={reportType === rt.type}
              onClick={() => setReportType(rt.type)}
            />
          ))}
        </div>
      </div>

      {/* Actions Panel */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6">
        <h3 className="text-lg font-semibold text-gray-900 mb-4">Generate & Send Report</h3>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Download Section */}
          <div className="space-y-4">
            <h4 className="font-medium text-gray-700">Download PDF</h4>
            <div className="flex flex-wrap gap-3">
              <button
                onClick={handleExportPDF}
                disabled={exporting}
                className="inline-flex items-center gap-2 px-4 py-2 bg-white border border-gray-300 rounded-lg text-sm font-medium text-gray-700 hover:bg-gray-50 transition-colors disabled:opacity-50"
              >
                {exporting ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
                {exporting ? 'Exporting...' : 'Export from View'}
              </button>
              <button
                onClick={handleDownloadServerPDF}
                disabled={downloading}
                className="inline-flex items-center gap-2 px-4 py-2 bg-slate-800 rounded-lg text-sm font-medium text-white hover:bg-slate-700 transition-colors disabled:opacity-50"
              >
                {downloading ? <Loader2 className="h-4 w-4 animate-spin" /> : <FileText className="h-4 w-4" />}
                {downloading ? 'Generating...' : 'Download Full Report'}
              </button>
            </div>
          </div>

          {/* Email Section */}
          <div className="space-y-4">
            <h4 className="font-medium text-gray-700">Send via Email</h4>
            <div className="flex flex-col sm:flex-row gap-3">
              <div className="flex-1">
                <input
                  type="email"
                  placeholder="Enter email address..."
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
                />
              </div>
              <button
                onClick={handleSendEmail}
                disabled={sending || !email.trim()}
                className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 rounded-lg text-sm font-medium text-white hover:bg-blue-700 transition-colors disabled:opacity-50"
              >
                {sending ? <Loader2 className="h-4 w-4 animate-spin" /> : <Send className="h-4 w-4" />}
                {sending ? 'Sending...' : 'Send Report'}
              </button>
            </div>
            <label className="flex items-center gap-2 text-sm text-gray-600">
              <input
                type="checkbox"
                checked={includeAttachments}
                onChange={(e) => setIncludeAttachments(e.target.checked)}
                className="rounded border-gray-300 text-slate-600 focus:ring-slate-500"
              />
              <Paperclip className="h-4 w-4" />
              Attach Risk Register spreadsheets
            </label>
          </div>
        </div>
      </div>

      {/* Critical Alerts Banner */}
      {highPriorityRisks.length > 0 && (
        <div className="bg-red-600 text-white rounded-xl p-6">
          <h3 className="text-lg font-bold flex items-center gap-2 mb-3">
            <AlertTriangle className="h-5 w-5" />
            CRITICAL ALERTS ({highPriorityRisks.length})
          </h3>
          <div className="space-y-2">
            {highPriorityRisks.slice(0, 3).map(risk => (
              <div key={risk['Risk ID']} className="bg-red-700/50 rounded-lg p-3">
                <p className="font-semibold">{risk['Risk ID']}: {risk.Title}</p>
                <p className="text-sm text-red-100 mt-1">{risk.Description?.slice(0, 150)}...</p>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Report Preview */}
      <div ref={reportRef} className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        {/* Report Header */}
        <div className="bg-gradient-to-r from-slate-800 to-slate-700 px-8 py-6 text-white">
          <div className="flex items-center gap-3 mb-2">
            <FileText className="h-8 w-8" />
            <h1 className="text-2xl font-bold">
              {reportType === 'monthly' ? 'Monthly Executive Report' :
               reportType === 'weekly' ? 'Weekly Summary Report' :
               reportType === 'daily' ? 'Daily Risk Digest' :
               'Current Status Report'}
            </h1>
          </div>
          <p className="text-slate-300">{project} Project - {today}</p>
        </div>

        <div className="p-8 space-y-8">
          {/* Portfolio Summary (for daily/monthly reports) */}
          {(reportType === 'daily' || reportType === 'monthly') && portfolio.length > 0 && (
            <section>
              <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4">
                Portfolio Overview
              </h2>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                {portfolio.map(p => (
                  <div
                    key={p.code}
                    className={`rounded-lg p-4 border-l-4 ${
                      p.stats?.health === 'Critical' ? 'border-red-500 bg-red-50' :
                      p.stats?.health === 'At Risk' ? 'border-yellow-500 bg-yellow-50' :
                      'border-green-500 bg-green-50'
                    }`}
                  >
                    <div className="flex justify-between items-start">
                      <div>
                        <h3 className="font-bold text-gray-900">{p.code}</h3>
                        <p className={`text-sm font-semibold ${
                          p.stats?.health === 'Critical' ? 'text-red-600' :
                          p.stats?.health === 'At Risk' ? 'text-yellow-600' :
                          'text-green-600'
                        }`}>{p.stats?.health || 'Unknown'}</p>
                      </div>
                    </div>
                    <p className="text-sm text-gray-600 mt-2">
                      {p.stats?.active_risks || 0} risks | {p.stats?.open_tasks || 0} tasks
                    </p>
                  </div>
                ))}
              </div>
            </section>
          )}

          {/* Summary Stats */}
          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4">
              Executive Summary
            </h2>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div className="bg-red-50 rounded-lg p-4 text-center">
                <p className="text-3xl font-bold text-red-700">{stats.active_risks || activeRisks.length}</p>
                <p className="text-sm text-red-600">Active Risks</p>
              </div>
              <div className="bg-yellow-50 rounded-lg p-4 text-center">
                <p className="text-3xl font-bold text-yellow-700">{stats.watching_risks || 0}</p>
                <p className="text-sm text-yellow-600">Watching</p>
              </div>
              <div className="bg-blue-50 rounded-lg p-4 text-center">
                <p className="text-3xl font-bold text-blue-700">{stats.open_tasks || openTasks.length}</p>
                <p className="text-sm text-blue-600">Open Tasks</p>
              </div>
              <div className="bg-orange-50 rounded-lg p-4 text-center">
                <p className="text-3xl font-bold text-orange-700">{criticalMilestones.length}</p>
                <p className="text-sm text-orange-600">Critical Milestones</p>
              </div>
            </div>
          </section>

          {/* Critical Milestones */}
          {criticalMilestones.length > 0 && (
            <section>
              <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4 flex items-center gap-2">
                <Calendar className="h-5 w-5 text-orange-500" />
                Critical Milestones
              </h2>
              <div className="space-y-3">
                {criticalMilestones.map((ms, idx) => (
                  <div key={idx} className="border-l-4 border-orange-500 bg-orange-50 p-4 rounded-r-lg">
                    <p className="font-semibold text-gray-900">{ms.Milestone}</p>
                    <p className="text-sm text-gray-600 mt-1">
                      Baseline: {ms['Baseline Date']} | Current: {ms['Current Date'] || 'TBD'}
                    </p>
                    {ms.Notes && <p className="text-sm text-orange-700 mt-1">{ms.Notes}</p>}
                  </div>
                ))}
              </div>
            </section>
          )}

          {/* High Priority Risks */}
          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4 flex items-center gap-2">
              <AlertTriangle className="h-5 w-5 text-red-500" />
              High Priority Risks
            </h2>
            {highPriorityRisks.length === 0 ? (
              <p className="text-gray-500 italic">No high priority risks at this time.</p>
            ) : (
              <div className="space-y-3">
                {highPriorityRisks.slice(0, 5).map(risk => (
                  <div key={risk['Risk ID']} className="border-l-4 border-red-500 bg-red-50 p-4 rounded-r-lg">
                    <div className="flex items-start justify-between">
                      <div>
                        <p className="font-semibold text-gray-900">{risk['Risk ID']}: {risk.Title}</p>
                        <p className="text-sm text-gray-600 mt-1">{risk.Description}</p>
                        <p className="text-xs text-gray-500 mt-2">
                          Owner: {risk.Owner || 'TBD'} | {risk.Probability}/{risk.Impact}
                        </p>
                      </div>
                      <TrendIcon trend={risk.Trend} />
                    </div>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* Active Risks Table */}
          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4">
              All Active Risks
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-50">
                    <th className="text-left p-3 font-semibold">Risk ID</th>
                    <th className="text-left p-3 font-semibold">Title</th>
                    <th className="text-left p-3 font-semibold">Priority</th>
                    <th className="text-left p-3 font-semibold">Owner</th>
                    <th className="text-left p-3 font-semibold">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {activeRisks.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="p-4 text-center text-gray-500">No active risks</td>
                    </tr>
                  ) : (
                    activeRisks.slice(0, 10).map(risk => (
                      <tr key={risk['Risk ID']}>
                        <td className="p-3 font-mono text-gray-600">{risk['Risk ID']}</td>
                        <td className="p-3">{risk.Title}</td>
                        <td className="p-3">
                          <span className={`px-2 py-1 text-xs rounded-full ${
                            risk.Probability === 'High' ? 'bg-red-100 text-red-800' :
                            risk.Probability === 'Medium' ? 'bg-yellow-100 text-yellow-800' :
                            'bg-green-100 text-green-800'
                          }`}>
                            {risk.Probability}/{risk.Impact}
                          </span>
                        </td>
                        <td className="p-3 text-gray-600">{risk.Owner || 'TBD'}</td>
                        <td className="p-3">
                          <span className={`px-2 py-1 text-xs rounded-full ${
                            risk.Status === 'Active' ? 'bg-red-100 text-red-800' :
                            risk.Status === 'Escalated' ? 'bg-purple-100 text-purple-800' :
                            'bg-blue-100 text-blue-800'
                          }`}>
                            {risk.Status}
                          </span>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </section>

          {/* Open Tasks */}
          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4 flex items-center gap-2">
              <CheckSquare className="h-5 w-5 text-blue-500" />
              Open Action Items
            </h2>
            {openTasks.length === 0 ? (
              <p className="text-gray-500 italic">No open action items.</p>
            ) : (
              <ul className="space-y-2">
                {openTasks.slice(0, 10).map(task => (
                  <li key={task['Task ID']} className="flex items-start gap-3 p-3 bg-gray-50 rounded-lg">
                    <span className="font-mono text-xs text-gray-500 mt-0.5">{task['Task ID']}</span>
                    <div className="flex-1">
                      <p className="text-gray-900">{task.Task}</p>
                      <p className="text-xs text-gray-500 mt-1">
                        Owner: {task.Owner || 'TBD'}
                        {task['Due Date'] && ` | Due: ${task['Due Date']}`}
                      </p>
                    </div>
                  </li>
                ))}
              </ul>
            )}
          </section>

          {/* Recommendations */}
          <section>
            <h2 className="text-lg font-semibold text-gray-900 border-b-2 border-slate-800 pb-2 mb-4">
              Recommendations
            </h2>
            <ul className="space-y-2 text-gray-700">
              {highPriorityRisks.length > 0 && (
                <li className="flex items-start gap-2">
                  <span className="text-red-500 mt-1">•</span>
                  <span><strong>CRITICAL:</strong> Address {highPriorityRisks.length} high priority risk(s) immediately</span>
                </li>
              )}
              {criticalMilestones.length > 0 && (
                <li className="flex items-start gap-2">
                  <span className="text-orange-500 mt-1">•</span>
                  <span><strong>Schedule:</strong> {criticalMilestones.length} milestone(s) at risk - executive attention required</span>
                </li>
              )}
              {stats.overdue_tasks > 0 && (
                <li className="flex items-start gap-2">
                  <span className="text-orange-500 mt-1">•</span>
                  <span>Follow up on {stats.overdue_tasks} overdue task(s)</span>
                </li>
              )}
              {activeRisks.length > 5 && (
                <li className="flex items-start gap-2">
                  <span className="text-yellow-500 mt-1">•</span>
                  <span>Consider risk review meeting to prioritize {activeRisks.length} active risks</span>
                </li>
              )}
              {activeRisks.length === 0 && openTasks.length === 0 && (
                <li className="flex items-start gap-2">
                  <span className="text-green-500 mt-1">•</span>
                  <span>Project is in good health - continue monitoring</span>
                </li>
              )}
            </ul>
          </section>
        </div>

        {/* Report Footer */}
        <div className="bg-gray-50 px-8 py-4 border-t border-gray-200 text-center text-sm text-gray-500">
          Risk Management System | Generated {new Date().toLocaleString()}
        </div>
      </div>
    </div>
  )
}
