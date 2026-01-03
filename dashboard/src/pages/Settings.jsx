import { useState, useEffect } from 'react'
import { Save, Key, Mail, Cloud, FolderOpen, ExternalLink, CheckCircle, AlertCircle } from 'lucide-react'
import { API_BASE } from '../hooks/useApi'

export default function Settings() {
  const [settings, setSettings] = useState({})
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [message, setMessage] = useState(null)

  useEffect(() => {
    fetchSettings()
  }, [])

  const fetchSettings = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/settings`)
      const data = await response.json()
      if (data.success) {
        setSettings(data.settings)
      }
    } catch (err) {
      console.error('Failed to load settings:', err)
    } finally {
      setLoading(false)
    }
  }

  const handleChange = (key, value) => {
    setSettings(prev => ({ ...prev, [key]: value }))
  }

  const handleSave = async () => {
    setSaving(true)
    setMessage(null)
    try {
      const response = await fetch(`${API_BASE}/api/settings`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(settings)
      })
      const data = await response.json()
      if (data.success) {
        setMessage({ type: 'success', text: data.message })
        fetchSettings() // Reload to get masked values
      } else {
        setMessage({ type: 'error', text: data.error || 'Failed to save' })
      }
    } catch (err) {
      setMessage({ type: 'error', text: 'Failed to save settings' })
    } finally {
      setSaving(false)
    }
  }

  if (loading) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-cyan-500"></div>
      </div>
    )
  }

  return (
    <div className="max-w-4xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold text-gray-800">Settings</h1>
        <button
          onClick={handleSave}
          disabled={saving}
          className="flex items-center gap-2 px-4 py-2 bg-cyan-600 text-white rounded-lg hover:bg-cyan-700 disabled:opacity-50"
        >
          <Save className="h-4 w-4" />
          {saving ? 'Saving...' : 'Save Settings'}
        </button>
      </div>

      {message && (
        <div className={`flex items-center gap-2 p-4 rounded-lg ${
          message.type === 'success' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'
        }`}>
          {message.type === 'success' ? <CheckCircle className="h-5 w-5" /> : <AlertCircle className="h-5 w-5" />}
          {message.text}
        </div>
      )}

      {/* Required Settings */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-cyan-600 to-cyan-700 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Key className="h-5 w-5" />
            Required Settings
          </h2>
        </div>
        <div className="p-6 space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Anthropic API Key
              <a href="https://console.anthropic.com/" target="_blank" rel="noopener noreferrer"
                 className="ml-2 text-cyan-600 hover:text-cyan-700">
                <ExternalLink className="h-3 w-3 inline" /> Get one
              </a>
            </label>
            <input
              type="password"
              value={settings.ANTHROPIC_API_KEY || ''}
              onChange={(e) => handleChange('ANTHROPIC_API_KEY', e.target.value)}
              placeholder="sk-ant-..."
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Project Names (comma-separated)
            </label>
            <input
              type="text"
              value={settings.PROJECT_NAMES || ''}
              onChange={(e) => handleChange('PROJECT_NAMES', e.target.value)}
              placeholder="RH, NSD, HERC, BRGO, HB"
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
            <p className="text-xs text-gray-500 mt-1">
              Folders will be created: 0 - Reports, 1 - [Project], 2 - [Project], etc.
            </p>
          </div>
        </div>
      </div>

      {/* ngrok Settings */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-purple-600 to-purple-700 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Cloud className="h-5 w-5" />
            Remote Access (ngrok)
          </h2>
        </div>
        <div className="p-6">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              ngrok Auth Token
              <a href="https://dashboard.ngrok.com/get-started/your-authtoken" target="_blank" rel="noopener noreferrer"
                 className="ml-2 text-cyan-600 hover:text-cyan-700">
                <ExternalLink className="h-3 w-3 inline" /> Get token
              </a>
            </label>
            <input
              type="password"
              value={settings.NGROK_AUTH_TOKEN || ''}
              onChange={(e) => handleChange('NGROK_AUTH_TOKEN', e.target.value)}
              placeholder="Optional - for Power Automate integration"
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
        </div>
      </div>

      {/* Email Settings */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-blue-600 to-blue-700 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Mail className="h-5 w-5" />
            Email Reports (Daily Digest)
          </h2>
        </div>
        <div className="p-6 grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">SMTP Server</label>
            <input
              type="text"
              value={settings.SMTP_SERVER || ''}
              onChange={(e) => handleChange('SMTP_SERVER', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">SMTP Port</label>
            <input
              type="text"
              value={settings.SMTP_PORT || ''}
              onChange={(e) => handleChange('SMTP_PORT', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">From Email</label>
            <input
              type="email"
              value={settings.EMAIL_FROM || ''}
              onChange={(e) => handleChange('EMAIL_FROM', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">To Email(s)</label>
            <input
              type="text"
              value={settings.EMAIL_TO || ''}
              onChange={(e) => handleChange('EMAIL_TO', e.target.value)}
              placeholder="recipient@example.com"
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div className="col-span-2">
            <label className="block text-sm font-medium text-gray-700 mb-1">Email Password / App Password</label>
            <input
              type="password"
              value={settings.EMAIL_PASSWORD || ''}
              onChange={(e) => handleChange('EMAIL_PASSWORD', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
        </div>
      </div>

      {/* Microsoft Graph Settings */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-orange-500 to-orange-600 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <FolderOpen className="h-5 w-5" />
            Microsoft Graph API (Outlook Calendar)
          </h2>
        </div>
        <div className="p-6 grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Client ID
              <a href="https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"
                 target="_blank" rel="noopener noreferrer" className="ml-2 text-cyan-600 hover:text-cyan-700">
                <ExternalLink className="h-3 w-3 inline" /> Azure Portal
              </a>
            </label>
            <input
              type="text"
              value={settings.MS_CLIENT_ID || ''}
              onChange={(e) => handleChange('MS_CLIENT_ID', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Tenant ID</label>
            <input
              type="text"
              value={settings.MS_TENANT_ID || ''}
              onChange={(e) => handleChange('MS_TENANT_ID', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Client Secret</label>
            <input
              type="password"
              value={settings.MS_CLIENT_SECRET || ''}
              onChange={(e) => handleChange('MS_CLIENT_SECRET', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">User Email</label>
            <input
              type="email"
              value={settings.MS_USER_EMAIL || ''}
              onChange={(e) => handleChange('MS_USER_EMAIL', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
        </div>
      </div>

      {/* IMAP Settings */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-green-600 to-green-700 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Mail className="h-5 w-5" />
            Email Reading (IMAP)
          </h2>
        </div>
        <div className="p-6 grid grid-cols-2 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">IMAP Server</label>
            <input
              type="text"
              value={settings.IMAP_SERVER || ''}
              onChange={(e) => handleChange('IMAP_SERVER', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">IMAP Port</label>
            <input
              type="text"
              value={settings.IMAP_PORT || ''}
              onChange={(e) => handleChange('IMAP_PORT', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
          <div className="col-span-2">
            <label className="block text-sm font-medium text-gray-700 mb-1">IMAP Password</label>
            <input
              type="password"
              value={settings.IMAP_PASSWORD || ''}
              onChange={(e) => handleChange('IMAP_PASSWORD', e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
          </div>
        </div>
      </div>

      {/* Save Button at bottom */}
      <div className="flex justify-end">
        <button
          onClick={handleSave}
          disabled={saving}
          className="flex items-center gap-2 px-6 py-3 bg-cyan-600 text-white rounded-lg hover:bg-cyan-700 disabled:opacity-50"
        >
          <Save className="h-5 w-5" />
          {saving ? 'Saving...' : 'Save All Settings'}
        </button>
      </div>
    </div>
  )
}
