import { useState } from 'react'
import { useParams } from 'react-router-dom'
import { useApi } from '../hooks/useApi'
import {
  AlertTriangle,
  Eye,
  CheckSquare,
  Clock,
  TrendingUp,
  TrendingDown,
  Minus,
  AlertCircle,
  X,
  Plus,
  Calendar,
  User,
  Loader2
} from 'lucide-react'

function StatCard({ title, value, icon: Icon, color, onClick, items }) {
  const colorClasses = {
    red: 'bg-gradient-to-br from-red-50 to-red-100/50 text-red-700 border-red-200 hover:border-red-300 hover:shadow-red-100',
    yellow: 'bg-gradient-to-br from-yellow-50 to-yellow-100/50 text-yellow-700 border-yellow-200 hover:border-yellow-300 hover:shadow-yellow-100',
    blue: 'bg-gradient-to-br from-blue-50 to-blue-100/50 text-blue-700 border-blue-200 hover:border-blue-300 hover:shadow-blue-100',
    gray: 'bg-gradient-to-br from-gray-50 to-gray-100/50 text-gray-700 border-gray-200 hover:border-gray-300',
  }

  const iconBg = {
    red: 'bg-red-100',
    yellow: 'bg-yellow-100',
    blue: 'bg-blue-100',
    gray: 'bg-gray-100',
  }

  return (
    <div
      onClick={() => items?.length > 0 && onClick?.()}
      className={`p-6 rounded-xl border-2 transition-all ${colorClasses[color] || colorClasses.gray} ${items?.length > 0 ? 'cursor-pointer hover:shadow-lg hover:-translate-y-0.5' : ''}`}
    >
      <div className="flex items-center justify-between">
        <div>
          <p className="text-sm font-medium opacity-75">{title}</p>
          <p className="text-4xl font-bold mt-1">{value}</p>
          {items?.length > 0 && <p className="text-xs opacity-60 mt-2 font-medium">Click to view details</p>}
        </div>
        <div className={`p-3 rounded-xl ${iconBg[color] || iconBg.gray}`}>
          <Icon className="h-8 w-8 opacity-70" />
        </div>
      </div>
    </div>
  )
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

function RiskBadge({ probability }) {
  const colors = {
    High: 'bg-red-100 text-red-800',
    Medium: 'bg-yellow-100 text-yellow-800',
    Low: 'bg-green-100 text-green-800',
  }
  return (
    <span className={`px-2 py-1 text-xs font-medium rounded-full ${colors[probability] || 'bg-gray-100 text-gray-800'}`}>
      {probability}
    </span>
  )
}

function DurationPicker({ isOpen, onClose, onSelect, taskTitle }) {
  const [duration, setDuration] = useState('30')
  const [date, setDate] = useState(new Date().toISOString().split('T')[0])
  const [time, setTime] = useState('09:00')

  if (!isOpen) return null

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 z-[60] flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-xl max-w-sm w-full p-6">
        <h4 className="font-semibold text-gray-900 mb-4">Schedule: {taskTitle}</h4>
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Date</label>
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
            />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Time</label>
              <input
                type="time"
                value={time}
                onChange={(e) => setTime(e.target.value)}
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Duration</label>
              <select
                value={duration}
                onChange={(e) => setDuration(e.target.value)}
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
              >
                <option value="15">15 min</option>
                <option value="30">30 min</option>
                <option value="45">45 min</option>
                <option value="60">1 hour</option>
                <option value="90">1.5 hours</option>
                <option value="120">2 hours</option>
              </select>
            </div>
          </div>
        </div>
        <div className="flex justify-end gap-2 mt-6">
          <button
            onClick={onClose}
            className="px-4 py-2 border border-gray-300 rounded-lg text-sm hover:bg-gray-50"
          >
            Cancel
          </button>
          <button
            onClick={() => onSelect({ duration: parseInt(duration), date, time })}
            className="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm hover:bg-blue-700 flex items-center gap-1"
          >
            <Calendar className="h-4 w-4" />
            Create Invite
          </button>
        </div>
      </div>
    </div>
  )
}

function DetailModal({ isOpen, onClose, title, type, items, onCreateOutlookTask }) {
  const [durationPickerOpen, setDurationPickerOpen] = useState(false)
  const [selectedTask, setSelectedTask] = useState(null)

  const handleAddToOutlook = (item) => {
    setSelectedTask(item)
    setDurationPickerOpen(true)
  }

  const handleDurationSelect = async ({ duration, date, time }) => {
    setDurationPickerOpen(false)
    if (selectedTask) {
      await onCreateOutlookTask(selectedTask, { duration, date, time })
      setSelectedTask(null)
    }
  }

  if (!isOpen) return null

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-xl max-w-2xl w-full max-h-[80vh] overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-200 flex items-center justify-between bg-slate-800 text-white">
          <h3 className="text-lg font-semibold">{title}</h3>
          <button onClick={onClose} className="p-1 hover:bg-slate-700 rounded">
            <X className="h-5 w-5" />
          </button>
        </div>
        <div className="overflow-y-auto max-h-[60vh] divide-y divide-gray-100">
          {items.length === 0 ? (
            <p className="p-6 text-gray-500 text-center">No items</p>
          ) : (
            items.map((item, idx) => (
              <div key={idx} className="p-4 hover:bg-gray-50">
                {type === 'risk' ? (
                  <div className="flex items-start justify-between">
                    <div className="flex-1">
                      <div className="flex items-center gap-2 mb-1">
                        <span className="text-xs font-mono text-gray-500">{item['Risk ID']}</span>
                        <RiskBadge probability={item.Probability} />
                        <TrendIcon trend={item.Trend} />
                        <span className={`px-2 py-0.5 text-xs rounded ${
                          item.Status === 'Escalated' ? 'bg-red-100 text-red-700' :
                          item.Status === 'Watching' ? 'bg-yellow-100 text-yellow-700' :
                          'bg-blue-100 text-blue-700'
                        }`}>
                          {item.Status}
                        </span>
                      </div>
                      <p className="font-medium text-gray-900">{item.Title}</p>
                      <p className="text-sm text-gray-500 mt-1">{item.Description}</p>
                      <p className="text-xs text-gray-400 mt-2">Owner: {item.Owner || 'Unassigned'}</p>
                    </div>
                  </div>
                ) : (
                  <div className="flex items-start justify-between">
                    <div className="flex-1">
                      <div className="flex items-center gap-2 mb-1">
                        <span className="text-xs font-mono text-gray-500">{item['Task ID']}</span>
                        {item['Due Date'] && (
                          <span className={`px-2 py-0.5 text-xs rounded flex items-center gap-1 ${
                            isOverdue(item['Due Date']) ? 'bg-red-100 text-red-700' : 'bg-gray-100 text-gray-600'
                          }`}>
                            <Clock className="h-3 w-3" />
                            {item['Due Date']}
                          </span>
                        )}
                      </div>
                      <p className="font-medium text-gray-900">{item.Task}</p>
                      <div className="flex items-center justify-between mt-2">
                        <p className="text-xs text-gray-400">Owner: {item.Owner || 'Unassigned'}</p>
                        <button
                          onClick={() => handleAddToOutlook(item)}
                          className="text-xs px-2 py-1 bg-blue-600 text-white rounded hover:bg-blue-700 flex items-center gap-1"
                        >
                          <Calendar className="h-3 w-3" />
                          Add to Calendar
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ))
          )}
        </div>
        <div className="px-6 py-3 border-t border-gray-200 bg-gray-50 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-2 bg-slate-800 text-white rounded-lg text-sm hover:bg-slate-700"
          >
            Close
          </button>
        </div>
      </div>

      <DurationPicker
        isOpen={durationPickerOpen}
        onClose={() => setDurationPickerOpen(false)}
        onSelect={handleDurationSelect}
        taskTitle={selectedTask?.Task || selectedTask?.Title || ''}
      />
    </div>
  )
}

function CreateTaskModal({ isOpen, onClose, projectCode, onTaskCreated }) {
  const [task, setTask] = useState('')
  const [owner, setOwner] = useState('')
  const [dueDate, setDueDate] = useState('')
  const [dueTime, setDueTime] = useState('09:00')
  const [duration, setDuration] = useState('30')
  const [addToOutlook, setAddToOutlook] = useState(true)
  const [creating, setCreating] = useState(false)
  const [error, setError] = useState(null)

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!task.trim()) return

    setCreating(true)
    setError(null)

    try {
      const response = await fetch('/api/calendar/event', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          project: projectCode,
          title: task.trim(),
          attendee: owner.trim() || null,
          date: dueDate || null,
          time: dueTime,
          duration: parseInt(duration),
          add_to_outlook: addToOutlook
        })
      })

      const data = await response.json()
      if (data.success) {
        onTaskCreated?.(data)
        setTask('')
        setOwner('')
        setDueDate('')
        setDueTime('09:00')
        setDuration('30')
        onClose()

        // If ICS file was created, prompt download
        if (data.download_url) {
          window.open(data.download_url, '_blank')
        }
      } else {
        setError(data.error || 'Failed to create event')
      }
    } catch (err) {
      setError('Failed to create event: ' + err.message)
    } finally {
      setCreating(false)
    }
  }

  if (!isOpen) return null

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-xl max-w-md w-full">
        <div className="px-6 py-4 border-b border-gray-200 flex items-center justify-between bg-slate-800 text-white">
          <h3 className="text-lg font-semibold flex items-center gap-2">
            <Calendar className="h-5 w-5" />
            Create Calendar Event
          </h3>
          <button onClick={onClose} className="p-1 hover:bg-slate-700 rounded">
            <X className="h-5 w-5" />
          </button>
        </div>
        <form onSubmit={handleSubmit} className="p-6 space-y-4">
          {error && (
            <div className="p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>
          )}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Event Title *</label>
            <textarea
              value={task}
              onChange={(e) => setTask(e.target.value)}
              className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
              rows={2}
              placeholder="Enter event title..."
              required
            />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                <User className="h-4 w-4 inline mr-1" />
                Invite (email)
              </label>
              <input
                type="email"
                value={owner}
                onChange={(e) => setOwner(e.target.value)}
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
                placeholder="email@example.com"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                <Calendar className="h-4 w-4 inline mr-1" />
                Date *
              </label>
              <input
                type="date"
                value={dueDate}
                onChange={(e) => setDueDate(e.target.value)}
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
                required
              />
            </div>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                <Clock className="h-4 w-4 inline mr-1" />
                Start Time
              </label>
              <input
                type="time"
                value={dueTime}
                onChange={(e) => setDueTime(e.target.value)}
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                <Clock className="h-4 w-4 inline mr-1" />
                Duration
              </label>
              <select
                value={duration}
                onChange={(e) => setDuration(e.target.value)}
                className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500 focus:border-slate-500"
              >
                <option value="15">15 minutes</option>
                <option value="30">30 minutes</option>
                <option value="45">45 minutes</option>
                <option value="60">1 hour</option>
                <option value="90">1.5 hours</option>
                <option value="120">2 hours</option>
                <option value="180">3 hours</option>
                <option value="240">4 hours</option>
              </select>
            </div>
          </div>
          <div className="flex items-center gap-2 p-3 bg-blue-50 rounded-lg">
            <input
              type="checkbox"
              id="addToOutlook"
              checked={addToOutlook}
              onChange={(e) => setAddToOutlook(e.target.checked)}
              className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
            />
            <label htmlFor="addToOutlook" className="text-sm text-gray-700 flex items-center gap-1">
              <Calendar className="h-4 w-4 text-blue-600" />
              Download .ics file for Outlook
            </label>
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 border border-gray-300 rounded-lg text-sm hover:bg-gray-50"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={creating || !task.trim() || !dueDate}
              className="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm hover:bg-blue-700 disabled:opacity-50 flex items-center gap-2"
            >
              {creating ? <Loader2 className="h-4 w-4 animate-spin" /> : <Calendar className="h-4 w-4" />}
              {creating ? 'Creating...' : 'Create Event'}
            </button>
          </div>
        </form>
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

export default function Dashboard() {
  const { projectCode } = useParams()
  const project = projectCode || 'HB'

  const { data: statsData, loading: statsLoading } = useApi(`/api/stats?project=${project}`)
  const { data: risksData, loading: risksLoading } = useApi(`/api/risks?project=${project}`)
  const { data: tasksData, loading: tasksLoading } = useApi(`/api/tasks?project=${project}`)
  const { data: updatesData, loading: updatesLoading } = useApi(`/api/updates?project=${project}&limit=5`)

  const [modalOpen, setModalOpen] = useState(false)
  const [modalTitle, setModalTitle] = useState('')
  const [modalType, setModalType] = useState('risk')
  const [modalItems, setModalItems] = useState([])
  const [createTaskOpen, setCreateTaskOpen] = useState(false)

  const stats = statsData?.stats || {}
  const risks = risksData?.risks || []
  const tasks = tasksData?.tasks || []
  const updates = updatesData?.updates || []

  // Filter data for stat cards
  const activeRisks = risks.filter(r => ['Open', 'Active', 'Escalated'].includes(r.Status))
  const watchingRisks = risks.filter(r => r.Status === 'Watching')
  const openTasks = tasks.filter(t => !['Completed', 'Done'].includes(t.Status))
  const overdueTasks = tasks.filter(t => {
    if (['Completed', 'Done'].includes(t.Status)) return false
    return isOverdue(t['Due Date'])
  })
  const itemsNotGreen = risks.filter(r => r.Status !== 'Closed' && r.Status)

  const openDetailModal = (title, type, items) => {
    setModalTitle(title)
    setModalType(type)
    setModalItems(items)
    setModalOpen(true)
  }

  const handleCreateOutlookTask = async (task, schedule) => {
    try {
      const response = await fetch('/api/calendar/event', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title: task.Task || task.Title,
          description: task.Description || '',
          date: schedule?.date || task['Due Date'],
          time: schedule?.time || '09:00',
          duration: schedule?.duration || 30,
          project: project,
          attendee: task.Owner || null
        })
      })
      const data = await response.json()
      if (data.success) {
        if (data.method === 'graph_api') {
          alert('Event created in Outlook calendar!')
        } else if (data.filename) {
          // Download ICS file
          window.location.href = `/api/calendar/download/${data.filename}`
        }
      } else {
        alert('Failed to create event: ' + (data.error || 'Unknown error'))
      }
    } catch (err) {
      alert('Failed to create event: ' + err.message)
    }
  }

  if (statsLoading || risksLoading || tasksLoading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-slate-800"></div>
      </div>
    )
  }

  return (
    <div className="space-y-6">
      {/* At A Glance */}
      <div>
        <div className="flex items-center justify-between mb-4">
          <div>
            <h3 className="text-xl font-bold text-gray-900">At A Glance</h3>
            <p className="text-sm text-gray-500">Quick overview of project status</p>
          </div>
          <button
            onClick={() => setCreateTaskOpen(true)}
            className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-lg text-sm hover:from-blue-700 hover:to-blue-800 shadow-sm hover:shadow transition-all"
          >
            <Calendar className="h-4 w-4" />
            New Event
          </button>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          <StatCard
            title="Active Risks"
            value={stats.active_risks || 0}
            icon={AlertTriangle}
            color="red"
            items={activeRisks}
            onClick={() => openDetailModal('Active Risks', 'risk', activeRisks)}
          />
          <StatCard
            title="Watching"
            value={stats.watching_risks || 0}
            icon={Eye}
            color="yellow"
            items={watchingRisks}
            onClick={() => openDetailModal('Watching Risks', 'risk', watchingRisks)}
          />
          <StatCard
            title="Open Tasks"
            value={stats.open_tasks || 0}
            icon={CheckSquare}
            color="blue"
            items={openTasks}
            onClick={() => openDetailModal('Open Tasks', 'task', openTasks)}
          />
          <StatCard
            title="Overdue Tasks"
            value={stats.overdue_tasks || 0}
            icon={Clock}
            color={stats.overdue_tasks > 0 ? 'red' : 'gray'}
            items={overdueTasks}
            onClick={() => openDetailModal('Overdue Tasks', 'task', overdueTasks)}
          />
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Items Not Green */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
          <div className="px-6 py-4 border-b border-gray-200 bg-gradient-to-r from-red-50 to-orange-50">
            <h3 className="text-lg font-semibold text-gray-900 flex items-center">
              <div className="p-1.5 bg-red-100 rounded-lg mr-2">
                <AlertCircle className="h-5 w-5 text-red-600" />
              </div>
              Items Not Green
            </h3>
            <p className="text-sm text-gray-500 mt-1">Risks requiring attention</p>
          </div>
          <div className="divide-y divide-gray-100 max-h-96 overflow-y-auto">
            {itemsNotGreen.length === 0 ? (
              <p className="p-6 text-gray-500 text-center">All risks are resolved!</p>
            ) : (
              itemsNotGreen.slice(0, 8).map((risk) => (
                <div key={risk['Risk ID']} className="p-4 hover:bg-gray-50">
                  <div className="flex items-start justify-between">
                    <div className="flex-1">
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-mono text-gray-500">{risk['Risk ID']}</span>
                        <RiskBadge probability={risk.Probability} />
                        <TrendIcon trend={risk.Trend} />
                      </div>
                      <p className="font-medium text-gray-900 mt-1">{risk.Title}</p>
                      <p className="text-sm text-gray-500 mt-1">Owner: {risk.Owner || 'Unassigned'}</p>
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>

        {/* Recent Activity */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
          <div className="px-6 py-4 border-b border-gray-200 bg-gradient-to-r from-blue-50 to-indigo-50">
            <h3 className="text-lg font-semibold text-gray-900 flex items-center">
              <div className="p-1.5 bg-blue-100 rounded-lg mr-2">
                <Clock className="h-5 w-5 text-blue-600" />
              </div>
              Recent Activity
            </h3>
            <p className="text-sm text-gray-500 mt-1">Latest updates from the system</p>
          </div>
          <div className="divide-y divide-gray-100 max-h-96 overflow-y-auto">
            {updatesLoading ? (
              <p className="p-6 text-gray-500 text-center">Loading...</p>
            ) : updates.length === 0 ? (
              <p className="p-6 text-gray-500 text-center">No recent activity</p>
            ) : (
              updates.map((update, idx) => (
                <div key={idx} className="p-4 hover:bg-gray-50">
                  <div className="flex items-start gap-3">
                    <div className="flex-shrink-0 w-2 h-2 mt-2 rounded-full bg-blue-500"></div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-gray-900">{update.Source}</p>
                      <p className="text-xs text-gray-500">{update['Source Type']} - {update.Timestamp}</p>
                      <p className="text-sm text-gray-600 mt-1 line-clamp-2">{update['Changes Made']}</p>
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* Detail Modal */}
      <DetailModal
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        title={modalTitle}
        type={modalType}
        items={modalItems}
        onCreateOutlookTask={handleCreateOutlookTask}
      />

      {/* Create Task Modal */}
      <CreateTaskModal
        isOpen={createTaskOpen}
        onClose={() => setCreateTaskOpen(false)}
        projectCode={project}
        onTaskCreated={() => window.location.reload()}
      />
    </div>
  )
}
