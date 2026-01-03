import { useNavigate } from 'react-router-dom'
import { useApi } from '../hooks/useApi'
import {
  AlertTriangle,
  CheckSquare,
  TrendingUp,
  ArrowRight,
  Building2,
  Sun,
  Home,
  Building,
  Zap,
  Battery,
  Shield
} from 'lucide-react'

// Brand Colors (customize during setup)
const BRAND_PRIMARY = '#33A9DC'
const BRAND_SECONDARY = '#58595B'

// Project icons - add your projects during setup
// Available icons: Sun, Battery, Shield, Building2, Building, Zap, Home
const projectIcons = {
  // Example: PROJECT1: Battery, PROJECT2: Shield
}

// Project display names - add your projects during setup
const projectNames = {
  // Example: PROJECT1: 'Project One', PROJECT2: 'Project Two'
}

function HealthBadge({ health }) {
  const colors = {
    Critical: 'bg-red-100 text-red-800 border-red-300',
    'At Risk': 'bg-orange-100 text-orange-800 border-orange-300',
    Caution: 'bg-yellow-100 text-yellow-800 border-yellow-300',
    Healthy: 'bg-green-100 text-green-800 border-green-300'
  }

  return (
    <span className={`px-3 py-1 text-xs font-semibold rounded-full border ${colors[health] || colors.Caution}`}>
      {health}
    </span>
  )
}

function ProjectCard({ project, onClick }) {
  const Icon = projectIcons[project.code] || Building2
  const stats = project.stats
  const name = projectNames[project.code] || project.code

  const healthColors = {
    Critical: 'border-l-red-500 hover:border-l-red-600',
    'At Risk': 'border-l-orange-500 hover:border-l-orange-600',
    Caution: 'border-l-yellow-500 hover:border-l-yellow-600',
    Healthy: 'border-l-green-500 hover:border-l-green-600'
  }

  const iconBgColors = {
    Critical: 'bg-red-100',
    'At Risk': 'bg-orange-100',
    Caution: 'bg-yellow-100',
    Healthy: 'bg-green-100'
  }

  return (
    <div
      onClick={onClick}
      className={`bg-white rounded-xl shadow-sm border border-gray-200 border-l-4 ${healthColors[stats.health]} p-6 cursor-pointer hover:shadow-lg transition-all hover:-translate-y-0.5`}
    >
      <div className="flex items-start justify-between mb-4">
        <div className="flex items-center gap-3">
          <div className={`p-2.5 rounded-xl ${iconBgColors[stats.health] || 'bg-slate-100'}`}>
            <Icon className="h-6 w-6 text-slate-700" />
          </div>
          <div>
            <h3 className="font-semibold text-gray-900">{name}</h3>
            <p className="text-xs text-gray-500">{project.code}</p>
          </div>
        </div>
        <HealthBadge health={stats.health} />
      </div>

      <div className="grid grid-cols-2 gap-3 mb-4">
        <div className="text-center p-3 bg-gradient-to-br from-red-50 to-red-100/50 rounded-xl border border-red-100">
          <div className="flex items-center justify-center gap-1.5 text-red-600">
            <AlertTriangle className="h-4 w-4" />
            <span className="text-2xl font-bold">{stats.active_risks}</span>
          </div>
          <p className="text-xs text-red-600/70 mt-1 font-medium">Active Risks</p>
        </div>
        <div className="text-center p-3 bg-gradient-to-br from-blue-50 to-blue-100/50 rounded-xl border border-blue-100">
          <div className="flex items-center justify-center gap-1.5 text-blue-600">
            <CheckSquare className="h-4 w-4" />
            <span className="text-2xl font-bold">{stats.open_tasks}</span>
          </div>
          <p className="text-xs text-blue-600/70 mt-1 font-medium">Open Tasks</p>
        </div>
      </div>

      <div className="flex items-center justify-between text-sm pt-3 border-t border-gray-100">
        <div className="flex items-center gap-4">
          <span className="flex items-center gap-1 text-yellow-600">
            <span className="w-2 h-2 bg-yellow-400 rounded-full"></span>
            {stats.watching_risks} watching
          </span>
          {stats.overdue_tasks > 0 && (
            <span className="flex items-center gap-1 text-orange-600 font-medium">
              <span className="w-2 h-2 bg-orange-500 rounded-full animate-pulse"></span>
              {stats.overdue_tasks} overdue
            </span>
          )}
        </div>
        <ArrowRight className="h-5 w-5 text-slate-400 group-hover:text-slate-600" />
      </div>
    </div>
  )
}

function PortfolioSummary({ projects, onNavigate }) {
  const totals = projects.reduce((acc, p) => ({
    active_risks: acc.active_risks + p.stats.active_risks,
    watching_risks: acc.watching_risks + p.stats.watching_risks,
    open_tasks: acc.open_tasks + p.stats.open_tasks,
    overdue_tasks: acc.overdue_tasks + p.stats.overdue_tasks,
    high_priority: acc.high_priority + (p.stats.high_priority || 0)
  }), { active_risks: 0, watching_risks: 0, open_tasks: 0, overdue_tasks: 0, high_priority: 0 })

  const criticalProjects = projects.filter(p => p.stats.health === 'Critical').length
  const atRiskProjects = projects.filter(p => p.stats.health === 'At Risk').length

  return (
    <div className="bg-gradient-to-br from-slate-800 via-slate-700 to-slate-800 rounded-xl p-6 text-white mb-6 shadow-lg">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-3">
          <img src="/logo.png" alt="Logo" className="h-10" onError={(e) => e.target.style.display='none'} />
          <div>
            <h2 className="text-xl font-bold">Portfolio Summary</h2>
            <p className="text-xs text-slate-400">{projects.length} active projects</p>
          </div>
        </div>
        {(criticalProjects > 0 || atRiskProjects > 0) && (
          <div className="flex items-center gap-3">
            {criticalProjects > 0 && (
              <span className="flex items-center gap-1.5 px-3 py-1.5 bg-red-500/20 text-red-300 rounded-full text-sm font-medium">
                <span className="w-2 h-2 bg-red-400 rounded-full animate-pulse"></span>
                {criticalProjects} Critical
              </span>
            )}
            {atRiskProjects > 0 && (
              <span className="flex items-center gap-1.5 px-3 py-1.5 bg-orange-500/20 text-orange-300 rounded-full text-sm font-medium">
                <span className="w-2 h-2 bg-orange-400 rounded-full"></span>
                {atRiskProjects} At Risk
              </span>
            )}
          </div>
        )}
      </div>
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div
          className="text-center cursor-pointer hover:bg-white/10 rounded-xl p-4 transition-all hover:scale-105"
          onClick={() => onNavigate('/risks')}
        >
          <p className="text-4xl font-bold text-red-400">{totals.active_risks}</p>
          <p className="text-sm text-slate-400 mt-1">Active Risks</p>
        </div>
        <div
          className="text-center cursor-pointer hover:bg-white/10 rounded-xl p-4 transition-all hover:scale-105"
          onClick={() => onNavigate('/risks')}
        >
          <p className="text-4xl font-bold text-yellow-400">{totals.high_priority}</p>
          <p className="text-sm text-slate-400 mt-1">High Priority</p>
        </div>
        <div
          className="text-center cursor-pointer hover:bg-white/10 rounded-xl p-4 transition-all hover:scale-105"
          onClick={() => onNavigate('/tasks')}
        >
          <p className="text-4xl font-bold text-blue-400">{totals.open_tasks}</p>
          <p className="text-sm text-slate-400 mt-1">Open Tasks</p>
        </div>
        <div
          className="text-center cursor-pointer hover:bg-white/10 rounded-xl p-4 transition-all hover:scale-105"
          onClick={() => onNavigate('/tasks')}
        >
          <p className="text-4xl font-bold text-orange-400">{totals.overdue_tasks}</p>
          <p className="text-sm text-slate-400 mt-1">Overdue</p>
        </div>
        <div
          className="text-center cursor-pointer hover:bg-white/10 rounded-xl p-4 transition-all hover:scale-105"
          onClick={() => onNavigate('/risks')}
        >
          <p className="text-4xl font-bold text-emerald-400">{totals.watching_risks}</p>
          <p className="text-sm text-slate-400 mt-1">Watching</p>
        </div>
      </div>
    </div>
  )
}

export default function Portfolio() {
  const navigate = useNavigate()
  const { data, loading, error } = useApi('/api/portfolio')

  const projects = data?.projects || []

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-slate-600"></div>
      </div>
    )
  }

  if (error) {
    return (
      <div className="bg-red-50 text-red-700 p-4 rounded-lg">
        Error loading portfolio: {error}
      </div>
    )
  }

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-gray-900">Project Portfolio</h1>
        <p className="text-gray-500">Risk Management Dashboard</p>
      </div>

      <PortfolioSummary projects={projects} onNavigate={navigate} />

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {projects.map(project => (
          <ProjectCard
            key={project.code}
            project={project}
            onClick={() => navigate(`/project/${project.code}`)}
          />
        ))}
      </div>
    </div>
  )
}
