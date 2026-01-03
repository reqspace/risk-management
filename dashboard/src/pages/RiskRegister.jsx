import { useState, Fragment } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { useApi } from '../hooks/useApi'
import {
  TrendingUp,
  TrendingDown,
  Minus,
  ChevronUp,
  ChevronDown,
  ChevronRight,
  Filter,
  Download,
  AlertTriangle,
  Eye,
  XCircle,
  AlertOctagon,
  CheckSquare,
  Clock,
  ArrowRight
} from 'lucide-react'

function downloadCSV(risks, filename) {
  const headers = ['Project', 'Risk ID', 'Title', 'Description', 'Category', 'Status', 'Probability', 'Impact', 'Risk Score', 'Owner', 'Mitigation', 'Last Updated']
  const csvContent = [
    headers.join(','),
    ...risks.map(r => [
      r.project || '',
      r['Risk ID'] || '',
      `"${(r.Title || '').replace(/"/g, '""')}"`,
      `"${(r.Description || '').replace(/"/g, '""')}"`,
      r.Category || '',
      r.Status || '',
      r.Probability || '',
      r.Impact || '',
      r['Risk Score'] || '',
      r.Owner || '',
      `"${(r.Mitigation || '').replace(/"/g, '""')}"`,
      r['Last Updated'] || ''
    ].join(','))
  ].join('\n')

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = filename
  link.click()
}

function TrendIcon({ trend }) {
  if (trend === 'Increasing' || trend === 'Up') {
    return <TrendingUp className="h-4 w-4 text-red-500" />
  }
  if (trend === 'Decreasing' || trend === 'Down') {
    return <TrendingDown className="h-4 w-4 text-green-500" />
  }
  return <Minus className="h-4 w-4 text-gray-400" />
}

function StatusBadge({ status }) {
  const colors = {
    Open: 'bg-red-100 text-red-800 border-red-200',
    Active: 'bg-red-100 text-red-800 border-red-200',
    Escalated: 'bg-red-200 text-red-900 border-red-300',
    Watching: 'bg-yellow-100 text-yellow-800 border-yellow-200',
    Closed: 'bg-green-100 text-green-800 border-green-200',
  }
  return (
    <span className={`px-2 py-1 text-xs font-medium rounded-full border ${colors[status] || 'bg-gray-100 text-gray-800 border-gray-200'}`}>
      {status || 'Unknown'}
    </span>
  )
}

function PriorityBadge({ probability, impact }) {
  const probColor = {
    High: 'text-red-600',
    Medium: 'text-yellow-600',
    Low: 'text-green-600',
  }
  return (
    <div className="text-xs">
      <span className={probColor[probability] || 'text-gray-600'}>{probability}</span>
      <span className="text-gray-400"> / </span>
      <span className={probColor[impact] || 'text-gray-600'}>{impact}</span>
    </div>
  )
}

function RiskSummary({ risks, onFilterChange, activeFilter }) {
  const activeRisks = risks.filter(r => ['Open', 'Active', 'Escalated'].includes(r.Status)).length
  const highHighRisks = risks.filter(r => r.Probability === 'High' && r.Impact === 'High' && r.Status !== 'Closed').length
  const watchingRisks = risks.filter(r => r.Status === 'Watching').length
  const closedRisks = risks.filter(r => r.Status === 'Closed').length

  const StatButton = ({ label, value, color, filterKey, icon: Icon }) => {
    const isActive = activeFilter === filterKey
    const baseClasses = "text-center cursor-pointer rounded-xl p-4 transition-all"
    const activeClasses = isActive
      ? "bg-white/20 ring-2 ring-white/50 scale-105"
      : "hover:bg-white/10 hover:scale-105"

    return (
      <div
        className={`${baseClasses} ${activeClasses}`}
        onClick={() => onFilterChange(isActive ? 'all' : filterKey)}
      >
        <div className="flex items-center justify-center gap-2">
          <Icon className={`h-5 w-5 ${color}`} />
          <p className={`text-3xl font-bold ${color}`}>{value}</p>
        </div>
        <p className="text-sm text-slate-400 mt-1">{label}</p>
        {isActive && <p className="text-xs text-white/60 mt-1">Click to clear</p>}
      </div>
    )
  }

  return (
    <div className="bg-gradient-to-br from-slate-800 via-slate-700 to-slate-800 rounded-xl p-6 text-white mb-6 shadow-lg">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-lg font-semibold">Risk Overview</h3>
        <p className="text-sm text-slate-400">Click a stat to filter</p>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <StatButton
          label="Active Risks"
          value={activeRisks}
          color="text-red-400"
          filterKey="active"
          icon={AlertTriangle}
        />
        <StatButton
          label="High/High"
          value={highHighRisks}
          color="text-orange-400"
          filterKey="high-high"
          icon={AlertOctagon}
        />
        <StatButton
          label="Watching"
          value={watchingRisks}
          color="text-yellow-400"
          filterKey="watching"
          icon={Eye}
        />
        <StatButton
          label="Closed"
          value={closedRisks}
          color="text-green-400"
          filterKey="closed"
          icon={XCircle}
        />
      </div>
    </div>
  )
}

function isOverdue(dueDate) {
  if (!dueDate) return false
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const due = new Date(dueDate)
  return due < today
}

function formatDate(dateStr) {
  if (!dateStr) return null
  try {
    const date = new Date(dateStr)
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
  } catch {
    return dateStr
  }
}

function TaskRow({ task, onClick }) {
  const overdue = isOverdue(task['Due Date']) && !['Completed', 'Done'].includes(task.Status)
  const isComplete = ['Completed', 'Done'].includes(task.Status)
  const formattedDate = formatDate(task['Due Date'])

  return (
    <div
      onClick={onClick}
      className={`flex items-center justify-between p-3 rounded-lg cursor-pointer transition-all hover:shadow-md ${
        isComplete ? 'bg-green-50 hover:bg-green-100' :
        overdue ? 'bg-red-50 hover:bg-red-100' :
        'bg-gray-50 hover:bg-gray-100'
      }`}
    >
      <div className="flex items-center gap-3 flex-1 min-w-0">
        <CheckSquare className={`h-4 w-4 flex-shrink-0 ${isComplete ? 'text-green-600' : 'text-blue-600'}`} />
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className="text-xs font-mono text-gray-400">{task['Task ID']}</span>
            {formattedDate && (
              <span className={`text-xs px-2 py-0.5 rounded ${
                overdue ? 'bg-red-100 text-red-700 font-medium' :
                isComplete ? 'bg-green-100 text-green-700' :
                'bg-gray-200 text-gray-600'
              }`}>
                <Clock className="h-3 w-3 inline mr-1" />
                {formattedDate}
              </span>
            )}
          </div>
          <p className="text-sm font-medium text-gray-900 mt-1">{task.Task}</p>
          <p className="text-xs text-gray-500 mt-0.5">Owner: {task.Owner || 'Unassigned'}</p>
        </div>
      </div>
      <div className="flex items-center gap-2">
        <span className={`px-2 py-1 text-xs font-medium rounded-full ${
          isComplete ? 'bg-green-100 text-green-800' :
          overdue ? 'bg-red-100 text-red-800' :
          'bg-blue-100 text-blue-800'
        }`}>
          {task.Status || 'Open'}
        </span>
        <ArrowRight className="h-4 w-4 text-gray-400" />
      </div>
    </div>
  )
}

export default function RiskRegister() {
  const { projectCode } = useParams()
  const navigate = useNavigate()
  const endpoint = projectCode ? `/api/risks?project=${projectCode}` : '/api/risks'
  const tasksEndpoint = projectCode ? `/api/tasks?project=${projectCode}` : '/api/tasks'

  const { data, loading, error } = useApi(endpoint)
  const { data: tasksData } = useApi(tasksEndpoint)

  const [sortField, setSortField] = useState('Risk ID')
  const [sortDirection, setSortDirection] = useState('asc')
  const [statusFilter, setStatusFilter] = useState('all')
  const [projectFilter, setProjectFilter] = useState('all')
  const [expandedRisk, setExpandedRisk] = useState(null)

  const risks = data?.risks || []
  const allTasks = tasksData?.tasks || []
  const projects = [...new Set(risks.map(r => r.project).filter(Boolean))]

  // Get tasks linked to a specific risk
  const getTasksForRisk = (risk) => {
    const riskId = risk['Risk ID']
    const riskProject = risk.project

    return allTasks.filter(task => {
      // Match by Linked Risk field
      if (task['Linked Risk'] === riskId) return true
      // Match by Source containing the risk ID
      if (task.Source && task.Source.includes(riskId)) return true
      // Match tasks from same project (fallback - shows all project tasks for now)
      return false
    })
  }

  const handleSort = (field) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc')
    } else {
      setSortField(field)
      setSortDirection('asc')
    }
  }

  const filteredRisks = risks.filter(risk => {
    const matchesStatus =
      statusFilter === 'all' ||
      (statusFilter === 'active' && ['Open', 'Active', 'Escalated'].includes(risk.Status)) ||
      (statusFilter === 'watching' && risk.Status === 'Watching') ||
      (statusFilter === 'closed' && risk.Status === 'Closed') ||
      (statusFilter === 'high-high' && risk.Probability === 'High' && risk.Impact === 'High' && risk.Status !== 'Closed')

    const matchesProject = projectFilter === 'all' || risk.project === projectFilter

    return matchesStatus && matchesProject
  })

  const sortedRisks = [...filteredRisks].sort((a, b) => {
    let aVal = a[sortField] || ''
    let bVal = b[sortField] || ''

    if (sortField === 'Risk Score') {
      aVal = Number(aVal) || 0
      bVal = Number(bVal) || 0
    }

    if (aVal < bVal) return sortDirection === 'asc' ? -1 : 1
    if (aVal > bVal) return sortDirection === 'asc' ? 1 : -1
    return 0
  })

  const SortHeader = ({ field, children }) => (
    <th
      className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider cursor-pointer hover:bg-gray-100 select-none"
      onClick={() => handleSort(field)}
    >
      <div className="flex items-center gap-1">
        {children}
        {sortField === field && (
          sortDirection === 'asc' ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />
        )}
      </div>
    </th>
  )

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-slate-800"></div>
      </div>
    )
  }

  if (error) {
    return (
      <div className="bg-red-50 text-red-700 p-4 rounded-lg">
        Error loading risks: {error}
      </div>
    )
  }

  return (
    <div className="space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div className="flex items-center gap-4">
          <h2 className="text-2xl font-bold text-gray-900">
            {projectCode ? 'Risk Register' : 'All Risks'}
          </h2>
          <button
            onClick={() => downloadCSV(sortedRisks, `risks_${projectCode || 'all'}_${new Date().toISOString().split('T')[0]}.csv`)}
            className="flex items-center gap-2 px-3 py-1.5 text-sm bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg transition-colors"
          >
            <Download className="h-4 w-4" />
            CSV
          </button>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <Filter className="h-4 w-4 text-gray-500" />
          {!projectCode && projects.length > 1 && (
            <select
              value={projectFilter}
              onChange={(e) => setProjectFilter(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
            >
              <option value="all">All Projects</option>
              {projects.map(p => (
                <option key={p} value={p}>{p}</option>
              ))}
            </select>
          )}
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
          >
            <option value="all">All Statuses</option>
            <option value="active">Active Only</option>
            <option value="high-high">High/High Priority</option>
            <option value="watching">Watching</option>
            <option value="closed">Closed</option>
          </select>
        </div>
      </div>

      <RiskSummary
        risks={risks}
        onFilterChange={setStatusFilter}
        activeFilter={statusFilter}
      />

      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gradient-to-r from-slate-100 to-slate-50 border-b border-gray-200">
              <tr>
                {!projectCode && <SortHeader field="project">Project</SortHeader>}
                <SortHeader field="Risk ID">ID</SortHeader>
                <SortHeader field="Title">Title</SortHeader>
                <SortHeader field="Status">Status</SortHeader>
                <SortHeader field="Probability">Priority</SortHeader>
                <SortHeader field="Risk Score">Score</SortHeader>
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Trend</th>
                <SortHeader field="Owner">Owner</SortHeader>
                <SortHeader field="Last Updated">Updated</SortHeader>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {sortedRisks.length === 0 ? (
                <tr>
                  <td colSpan={projectCode ? 8 : 9} className="px-4 py-8 text-center text-gray-500">
                    No risks found
                  </td>
                </tr>
              ) : (
                sortedRisks.map((risk) => {
                  const riskKey = `${risk.project}-${risk['Risk ID']}`
                  const isExpanded = expandedRisk === riskKey
                  const linkedTasks = getTasksForRisk(risk)

                  return (
                    <Fragment key={riskKey}>
                      <tr
                        onClick={() => setExpandedRisk(isExpanded ? null : riskKey)}
                        className={`cursor-pointer transition-colors ${isExpanded ? 'bg-slate-50' : 'hover:bg-gray-50'}`}
                      >
                        {!projectCode && (
                          <td className="px-4 py-3 text-sm font-medium text-slate-600">{risk.project}</td>
                        )}
                        <td className="px-4 py-3">
                          <div className="flex items-center gap-2">
                            <ChevronRight className={`h-4 w-4 text-gray-400 transition-transform ${isExpanded ? 'rotate-90' : ''}`} />
                            <span className="text-sm font-mono text-gray-600">{risk['Risk ID']}</span>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <div className="max-w-xs">
                            <p className="text-sm font-medium text-gray-900 truncate">{risk.Title}</p>
                            <p className="text-xs text-gray-500 truncate">{risk.Category}</p>
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          <StatusBadge status={risk.Status} />
                        </td>
                        <td className="px-4 py-3">
                          <PriorityBadge probability={risk.Probability} impact={risk.Impact} />
                        </td>
                        <td className="px-4 py-3">
                          <span className={`text-sm font-semibold ${
                            risk['Risk Score'] >= 7 ? 'text-red-600' :
                            risk['Risk Score'] >= 4 ? 'text-yellow-600' : 'text-green-600'
                          }`}>
                            {risk['Risk Score'] || '-'}
                          </span>
                        </td>
                        <td className="px-4 py-3">
                          <TrendIcon trend={risk.Trend} />
                        </td>
                        <td className="px-4 py-3 text-sm text-gray-600">{risk.Owner || 'Unassigned'}</td>
                        <td className="px-4 py-3 text-sm text-gray-500">{risk['Last Updated']}</td>
                      </tr>
                      {isExpanded && (
                        <tr className="bg-slate-50">
                          <td colSpan={projectCode ? 9 : 10} className="px-6 py-4">
                            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                              {/* Risk Details */}
                              <div className="space-y-4">
                                <div>
                                  <h4 className="text-sm font-semibold text-gray-700 mb-2">Description</h4>
                                  <p className="text-sm text-gray-600 bg-white p-3 rounded-lg border border-gray-200">
                                    {risk.Description || 'No description available'}
                                  </p>
                                </div>
                                <div>
                                  <h4 className="text-sm font-semibold text-gray-700 mb-2">Mitigation Strategy</h4>
                                  <p className="text-sm text-gray-600 bg-white p-3 rounded-lg border border-gray-200">
                                    {risk.Mitigation || 'No mitigation strategy defined'}
                                  </p>
                                </div>
                              </div>

                              {/* Linked Tasks */}
                              <div>
                                <div className="flex items-center justify-between mb-2">
                                  <h4 className="text-sm font-semibold text-gray-700 flex items-center gap-2">
                                    <CheckSquare className="h-4 w-4 text-blue-600" />
                                    Related Tasks
                                  </h4>
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation()
                                      navigate(projectCode ? `/project/${projectCode}/tasks` : '/tasks')
                                    }}
                                    className="text-xs text-blue-600 hover:text-blue-800 font-medium"
                                  >
                                    View All Tasks →
                                  </button>
                                </div>
                                <div className="space-y-2 max-h-48 overflow-y-auto">
                                  {linkedTasks.length === 0 ? (
                                    <p className="text-sm text-gray-500 bg-white p-3 rounded-lg border border-gray-200 text-center">
                                      No linked tasks found
                                    </p>
                                  ) : (
                                    linkedTasks.map((task) => (
                                      <TaskRow
                                        key={task['Task ID']}
                                        task={task}
                                        onClick={(e) => {
                                          e.stopPropagation()
                                          const taskId = task['Task ID']
                                          const basePath = projectCode ? `/project/${projectCode}/tasks` : '/tasks'
                                          navigate(`${basePath}?task=${encodeURIComponent(taskId)}`)
                                        }}
                                      />
                                    ))
                                  )}
                                </div>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  )
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="text-sm text-gray-500">
        Showing {sortedRisks.length} of {risks.length} risks
      </div>
    </div>
  )
}
