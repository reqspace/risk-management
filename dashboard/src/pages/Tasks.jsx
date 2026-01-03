import { useState, useEffect } from 'react'
import { useParams, useSearchParams } from 'react-router-dom'
import { useApi } from '../hooks/useApi'
import { Filter, Clock, CheckCircle, Circle, Download, Calendar, X } from 'lucide-react'

function StatusBadge({ status }) {
  const isCompleted = status === 'Completed' || status === 'Done'
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-1 text-xs font-medium rounded-full ${
      isCompleted
        ? 'bg-green-100 text-green-800'
        : 'bg-blue-100 text-blue-800'
    }`}>
      {isCompleted ? <CheckCircle className="h-3 w-3" /> : <Circle className="h-3 w-3" />}
      {status || 'Open'}
    </span>
  )
}

function isOverdue(dueDate) {
  if (!dueDate) return false
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const due = new Date(dueDate)
  return due < today
}

function downloadCSV(tasks, filename) {
  const headers = ['Project', 'Task ID', 'Task', 'Owner', 'Due Date', 'Status', 'Source']
  const csvContent = [
    headers.join(','),
    ...tasks.map(t => [
      t.project || '',
      t['Task ID'] || '',
      `"${(t.Task || '').replace(/"/g, '""')}"`,
      t.Owner || '',
      t['Due Date'] || '',
      t.Status || '',
      t.Source || ''
    ].join(','))
  ].join('\n')

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = filename
  link.click()
}

function createCalendarEvent(task) {
  const title = `[${task.project}] ${task.Task}`
  const dueDate = task['Due Date']

  if (!dueDate) {
    alert('No due date set for this task')
    return
  }

  // Create ICS content
  const date = new Date(dueDate)
  const dateStr = date.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z'
  const endDate = new Date(date.getTime() + 60 * 60 * 1000) // 1 hour later
  const endStr = endDate.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z'

  const icsContent = `BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Risk Management//Task//EN
BEGIN:VEVENT
UID:${task['Task ID']}-${Date.now()}@riskmanager
DTSTAMP:${new Date().toISOString().replace(/[-:]/g, '').split('.')[0]}Z
DTSTART:${dateStr}
DTEND:${endStr}
SUMMARY:${title}
DESCRIPTION:Owner: ${task.Owner || 'Unassigned'}\\nStatus: ${task.Status || 'Open'}
STATUS:CONFIRMED
END:VEVENT
END:VCALENDAR`

  const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = `task_${task['Task ID']}.ics`
  link.click()
}

export default function Tasks() {
  const { projectCode } = useParams()
  const [searchParams, setSearchParams] = useSearchParams()
  const endpoint = projectCode ? `/api/tasks?project=${projectCode}` : '/api/tasks'

  const { data, loading, error } = useApi(endpoint)
  const [statusFilter, setStatusFilter] = useState('all')
  const [ownerFilter, setOwnerFilter] = useState('all')
  const [projectFilter, setProjectFilter] = useState('all')

  // Get task filter from URL
  const taskIdFilter = searchParams.get('task')

  const tasks = data?.tasks || []

  // Get unique values for filters
  const owners = [...new Set(tasks.map(t => t.Owner).filter(Boolean))]
  const projects = [...new Set(tasks.map(t => t.project).filter(Boolean))]

  // Clear task filter
  const clearTaskFilter = () => {
    searchParams.delete('task')
    setSearchParams(searchParams)
  }

  const filteredTasks = tasks.filter(task => {
    // If filtering by specific task ID, only show that task
    if (taskIdFilter) {
      return task['Task ID'] === taskIdFilter
    }

    const matchesStatus = statusFilter === 'all' ||
      (statusFilter === 'open' && !['Completed', 'Done'].includes(task.Status)) ||
      (statusFilter === 'completed' && ['Completed', 'Done'].includes(task.Status)) ||
      (statusFilter === 'overdue' && isOverdue(task['Due Date']) && !['Completed', 'Done'].includes(task.Status))

    const matchesOwner = ownerFilter === 'all' || task.Owner === ownerFilter
    const matchesProject = projectFilter === 'all' || task.project === projectFilter

    return matchesStatus && matchesOwner && matchesProject
  })

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
        Error loading tasks: {error}
      </div>
    )
  }

  return (
    <div className="space-y-4">
      {/* Task ID Filter Banner */}
      {taskIdFilter && (
        <div className="flex items-center justify-between bg-blue-50 border border-blue-200 rounded-lg px-4 py-3">
          <div className="flex items-center gap-2">
            <Filter className="h-4 w-4 text-blue-600" />
            <span className="text-sm text-blue-800">
              Showing task: <span className="font-mono font-semibold">{taskIdFilter}</span>
            </span>
          </div>
          <button
            onClick={clearTaskFilter}
            className="flex items-center gap-1 px-3 py-1 text-sm bg-blue-100 hover:bg-blue-200 text-blue-700 rounded-lg transition-colors"
          >
            <X className="h-4 w-4" />
            Show All Tasks
          </button>
        </div>
      )}

      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div className="flex items-center gap-4">
          <h2 className="text-2xl font-bold text-gray-900">
            {projectCode ? 'Tasks' : 'All Tasks'}
          </h2>
          <button
            onClick={() => downloadCSV(filteredTasks, `tasks_${projectCode || 'all'}_${new Date().toISOString().split('T')[0]}.csv`)}
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
            <option value="all">All Tasks</option>
            <option value="open">Open</option>
            <option value="completed">Completed</option>
            <option value="overdue">Overdue</option>
          </select>
          <select
            value={ownerFilter}
            onChange={(e) => setOwnerFilter(e.target.value)}
            className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
          >
            <option value="all">All Owners</option>
            {owners.map(owner => (
              <option key={owner} value={owner}>{owner}</option>
            ))}
          </select>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gradient-to-r from-slate-100 to-slate-50 border-b border-gray-200">
              <tr>
                {!projectCode && <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Project</th>}
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">ID</th>
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Task</th>
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Owner</th>
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Due Date</th>
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Status</th>
                <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider">Source</th>
                <th className="px-4 py-3 text-center text-xs font-semibold text-gray-600 uppercase tracking-wider">Cal</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filteredTasks.length === 0 ? (
                <tr>
                  <td colSpan={projectCode ? 8 : 9} className="px-4 py-8 text-center text-gray-500">
                    No tasks found
                  </td>
                </tr>
              ) : (
                filteredTasks.map((task) => {
                  const overdue = isOverdue(task['Due Date']) && !['Completed', 'Done'].includes(task.Status)
                  const isComplete = ['Completed', 'Done'].includes(task.Status)
                  return (
                    <tr key={`${task.project}-${task['Task ID']}`} className={`hover:bg-gray-50 ${overdue ? 'bg-red-50' : ''}`}>
                      {!projectCode && (
                        <td className="px-4 py-3 text-sm font-medium text-slate-600">{task.project}</td>
                      )}
                      <td className="px-4 py-3 text-sm font-mono text-gray-600">{task['Task ID']}</td>
                      <td className="px-4 py-3">
                        <p className="text-sm font-medium text-gray-900 max-w-md">{task.Task}</p>
                      </td>
                      <td className="px-4 py-3 text-sm text-gray-600">{task.Owner || 'Unassigned'}</td>
                      <td className="px-4 py-3">
                        <div className={`flex items-center gap-1 text-sm ${overdue ? 'text-red-600 font-semibold' : 'text-gray-600'}`}>
                          {overdue && <Clock className="h-4 w-4" />}
                          {task['Due Date'] || 'No date'}
                        </div>
                      </td>
                      <td className="px-4 py-3">
                        <StatusBadge status={task.Status} />
                      </td>
                      <td className="px-4 py-3 text-sm text-gray-500">{task.Source}</td>
                      <td className="px-4 py-3 text-center">
                        {!isComplete && task['Due Date'] && (
                          <button
                            onClick={() => createCalendarEvent(task)}
                            className="p-1.5 text-slate-500 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors"
                            title="Add to Calendar"
                          >
                            <Calendar className="h-4 w-4" />
                          </button>
                        )}
                      </td>
                    </tr>
                  )
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="text-sm text-gray-500">
        Showing {filteredTasks.length} of {tasks.length} tasks
      </div>
    </div>
  )
}
