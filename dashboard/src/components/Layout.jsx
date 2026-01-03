import { useState, useEffect } from 'react'
import { Link, useLocation, useParams } from 'react-router-dom'
import {
  LayoutDashboard,
  AlertTriangle,
  CheckSquare,
  FileText,
  Menu,
  X,
  FolderKanban,
  ChevronLeft,
  Settings
} from 'lucide-react'

// Brand Colors (customize during setup)
const BRAND_PRIMARY = '#33A9DC'
const BRAND_SECONDARY = '#58595B'

export default function Layout({ children }) {
  const location = useLocation()
  const { projectCode } = useParams()
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [lastUpdated, setLastUpdated] = useState(null)

  const isProjectView = !!projectCode

  const navigation = isProjectView
    ? [
        { name: 'Dashboard', href: `/project/${projectCode}`, icon: LayoutDashboard },
        { name: 'Risk Register', href: `/project/${projectCode}/risks`, icon: AlertTriangle },
        { name: 'Tasks', href: `/project/${projectCode}/tasks`, icon: CheckSquare },
        { name: 'Reports', href: `/project/${projectCode}/reports`, icon: FileText },
      ]
    : [
        { name: 'Portfolio', href: '/', icon: FolderKanban },
        { name: 'All Risks', href: '/risks', icon: AlertTriangle },
        { name: 'All Tasks', href: '/tasks', icon: CheckSquare },
        { name: 'Reports', href: '/reports', icon: FileText },
        { name: 'Settings', href: '/settings', icon: Settings },
      ]

  useEffect(() => {
    const endpoint = projectCode ? `/api/stats?project=${projectCode}` : '/api/portfolio'
    fetch(endpoint)
      .then(res => res.json())
      .then(data => {
        if (data.success) {
          setLastUpdated(data.last_updated || new Date().toLocaleString())
        }
      })
      .catch(console.error)
  }, [location, projectCode])

  const projectName = projectCode || 'Portfolio'

  return (
    <div className="min-h-screen" style={{ backgroundColor: '#f5f7fa' }}>
      {/* Mobile sidebar backdrop */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 bg-gray-600 bg-opacity-75 z-20 lg:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Sidebar */}
      <div
        className={`fixed inset-y-0 left-0 z-30 w-64 bg-slate-800 transform transition-transform duration-300 ease-in-out lg:translate-x-0 ${sidebarOpen ? 'translate-x-0' : '-translate-x-full'}`}
      >
        <div className="flex items-center justify-between h-20 px-4 bg-slate-900">
          <Link to="/" className="flex items-center gap-3">
            <img src="/logo.png" alt="Logo" className="h-12" onError={(e) => e.target.style.display='none'} />
            <div>
              <span className="text-white font-semibold text-sm">Risk Management</span>
              <span className="block text-xs" style={{ color: BRAND_PRIMARY }}>Dashboard</span>
            </div>
          </Link>
          <button
            className="lg:hidden text-white"
            onClick={() => setSidebarOpen(false)}
          >
            <X className="h-6 w-6" />
          </button>
        </div>

        {isProjectView && (
          <div className="px-3 py-4 border-b border-slate-700">
            <Link
              to="/"
              className="flex items-center gap-2 text-sm mb-2 hover:opacity-80"
              style={{ color: BRAND_PRIMARY }}
            >
              <ChevronLeft className="h-4 w-4" />
              Back to Portfolio
            </Link>
            <p className="text-white font-semibold">{projectName}</p>
            <p className="text-xs text-slate-400">{projectCode}</p>
          </div>
        )}

        <nav className="mt-4 px-3">
          {navigation.map((item) => {
            const isActive = location.pathname === item.href
            return (
              <Link
                key={item.name}
                to={item.href}
                onClick={() => setSidebarOpen(false)}
                className={`flex items-center px-4 py-3 mb-1 rounded-lg text-sm font-medium transition-colors ${
                  isActive
                    ? 'bg-slate-700 text-white'
                    : 'text-slate-300 hover:bg-slate-700 hover:text-white'
                }`}
                style={isActive ? { borderLeft: `3px solid ${BRAND_PRIMARY}` } : {}}
              >
                <item.icon className="h-5 w-5 mr-3" style={isActive ? { color: BRAND_PRIMARY } : {}} />
                {item.name}
              </Link>
            )
          })}
        </nav>

        {/* Footer */}
        <div className="absolute bottom-0 left-0 right-0 p-4 text-center border-t border-slate-700">
          <p className="text-xs text-slate-500">Risk Management</p>
          <p className="text-xs text-slate-600 mt-1">v1.0.3</p>
        </div>
      </div>

      {/* Main content */}
      <div className="lg:pl-64">
        {/* Header */}
        <header className="bg-gradient-to-r from-slate-800 via-slate-700 to-slate-800 shadow-lg">
          <div className="flex items-center justify-between h-16 px-4 lg:px-8">
            <div className="flex items-center">
              <button
                className="lg:hidden text-white mr-4"
                onClick={() => setSidebarOpen(true)}
              >
                <Menu className="h-6 w-6" />
              </button>
              <div>
                <h2 className="text-xl font-semibold text-white">{projectName}</h2>
                {lastUpdated && (
                  <p className="text-xs text-slate-400">Last updated: {lastUpdated}</p>
                )}
              </div>
            </div>
            <div className="flex items-center space-x-4">
              <img src="/logo.png" alt="Logo" className="h-8 hidden lg:block opacity-90" onError={(e) => e.target.style.display='none'} />
              <span className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-green-400 bg-green-900/30 rounded-full border border-green-800/30">
                <span className="w-2 h-2 bg-green-400 rounded-full animate-pulse"></span>
                System Online
              </span>
            </div>
          </div>
        </header>

        {/* Page content */}
        <main className="p-4 lg:p-8">
          {children}
        </main>
      </div>
    </div>
  )
}
