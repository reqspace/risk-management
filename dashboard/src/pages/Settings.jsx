import { useState, useEffect, useRef } from 'react'
import { Save, Key, Mail, Cloud, FolderOpen, ExternalLink, CheckCircle, AlertCircle, RefreshCw, Loader2, Folder, Upload, Image, Plus, X, FolderPlus } from 'lucide-react'
import { API_BASE } from '../hooks/useApi'

export default function Settings() {
  const [settings, setSettings] = useState({})
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [message, setMessage] = useState(null)
  const [dataPath, setDataPath] = useState(null)
  const [testResults, setTestResults] = useState({})
  const [testing, setTesting] = useState({})
  const [logoPreview, setLogoPreview] = useState(null)
  const [uploadingLogo, setUploadingLogo] = useState(false)
  const [projects, setProjects] = useState([''])
  const [msAccount, setMsAccount] = useState(null)
  const [msSigningIn, setMsSigningIn] = useState(false)
  const fileInputRef = useRef(null)

  // Check if running in Electron
  const isElectron = typeof window !== 'undefined' && window.electronAPI?.isElectron

  useEffect(() => {
    fetchSettings()
    fetchDataPath()
    // Try to load existing logo
    setLogoPreview('/logo.png?' + Date.now())
    // Check Microsoft account status
    if (isElectron) {
      checkMicrosoftAccount()
    }
  }, [])

  // Check if signed in to Microsoft
  const checkMicrosoftAccount = async () => {
    try {
      const result = await window.electronAPI.getMicrosoftAccount()
      if (result.success) {
        setMsAccount(result.account)
      }
    } catch (err) {
      console.error('Error checking Microsoft account:', err)
    }
  }

  // Sign in with Microsoft
  const handleMicrosoftSignIn = async () => {
    setMsSigningIn(true)
    try {
      const result = await window.electronAPI.signInWithMicrosoft()
      if (result.success) {
        setMsAccount(result.account)
        setMessage({ type: 'success', text: `Signed in as ${result.account.username}` })
      } else {
        setMessage({ type: 'error', text: result.error || 'Sign in failed' })
      }
    } catch (err) {
      setMessage({ type: 'error', text: 'Sign in failed: ' + err.message })
    } finally {
      setMsSigningIn(false)
    }
  }

  // Sign out from Microsoft
  const handleMicrosoftSignOut = async () => {
    try {
      await window.electronAPI.signOutMicrosoft()
      setMsAccount(null)
      setMessage({ type: 'success', text: 'Signed out from Microsoft' })
    } catch (err) {
      setMessage({ type: 'error', text: 'Sign out failed: ' + err.message })
    }
  }

  // Sync projects array with PROJECT_NAMES setting
  useEffect(() => {
    if (settings.PROJECT_NAMES) {
      const names = settings.PROJECT_NAMES.split(',').map(n => n.trim()).filter(n => n)
      setProjects(names.length > 0 ? names : [''])
    }
  }, [settings.PROJECT_NAMES])

  // Helper functions for project management
  const addProject = () => {
    setProjects([...projects, ''])
  }

  const removeProject = (index) => {
    if (projects.length > 1) {
      const newProjects = projects.filter((_, i) => i !== index)
      setProjects(newProjects)
      // Update settings
      handleChange('PROJECT_NAMES', newProjects.filter(p => p.trim()).join(', '))
    }
  }

  const updateProject = (index, value) => {
    const newProjects = [...projects]
    newProjects[index] = value
    setProjects(newProjects)
    // Update settings
    handleChange('PROJECT_NAMES', newProjects.filter(p => p.trim()).join(', '))
  }

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

  const fetchDataPath = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/data-path`)
      const data = await response.json()
      if (data.success) {
        setDataPath(data)
      }
    } catch (err) {
      console.error('Failed to get data path:', err)
    }
  }

  const handleChange = (key, value) => {
    setSettings(prev => ({ ...prev, [key]: value }))
    // Clear test result when value changes
    setTestResults(prev => ({ ...prev, [key]: null }))
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
        fetchDataPath() // Refresh paths
      } else {
        setMessage({ type: 'error', text: data.error || 'Failed to save' })
      }
    } catch (err) {
      setMessage({ type: 'error', text: 'Failed to save settings' })
    } finally {
      setSaving(false)
    }
  }

  const testConnection = async (service) => {
    setTesting(prev => ({ ...prev, [service]: true }))
    try {
      const response = await fetch(`${API_BASE}/api/test-connection`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ service })
      })
      const data = await response.json()
      setTestResults(prev => ({ ...prev, [service]: data }))
    } catch (err) {
      setTestResults(prev => ({ ...prev, [service]: { success: false, error: err.message } }))
    } finally {
      setTesting(prev => ({ ...prev, [service]: false }))
    }
  }

  const handleLogoUpload = async (event) => {
    const file = event.target.files[0]
    if (!file) return

    // Validate file type
    if (!file.type.startsWith('image/')) {
      setMessage({ type: 'error', text: 'Please select an image file (PNG, JPG, etc.)' })
      return
    }

    // Show preview immediately
    const reader = new FileReader()
    reader.onload = (e) => setLogoPreview(e.target.result)
    reader.readAsDataURL(file)

    // Upload to server
    setUploadingLogo(true)
    try {
      const formData = new FormData()
      formData.append('logo', file)

      const response = await fetch(`${API_BASE}/api/upload-logo`, {
        method: 'POST',
        body: formData
      })
      const data = await response.json()

      if (data.success) {
        setMessage({ type: 'success', text: 'Logo uploaded successfully!' })
        // Force refresh logo in other components
        setLogoPreview(`/logo.png?${Date.now()}`)
      } else {
        setMessage({ type: 'error', text: data.error || 'Failed to upload logo' })
      }
    } catch (err) {
      setMessage({ type: 'error', text: 'Failed to upload logo: ' + err.message })
    } finally {
      setUploadingLogo(false)
    }
  }

  const TestButton = ({ service, label }) => (
    <button
      onClick={() => testConnection(service)}
      disabled={testing[service]}
      className="flex items-center gap-1 px-3 py-1 text-sm bg-gray-100 hover:bg-gray-200 rounded-lg disabled:opacity-50"
    >
      {testing[service] ? <Loader2 className="h-3 w-3 animate-spin" /> : <RefreshCw className="h-3 w-3" />}
      Test
    </button>
  )

  const TestResult = ({ service }) => {
    const result = testResults[service]
    if (!result) return null
    return (
      <div className={`flex items-center gap-1 text-sm mt-1 ${result.success ? 'text-green-600' : 'text-red-600'}`}>
        {result.success ? <CheckCircle className="h-4 w-4" /> : <AlertCircle className="h-4 w-4" />}
        {result.message || result.error}
      </div>
    )
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

      {/* Data Location */}
      {dataPath && (
        <div className="bg-blue-50 border border-blue-200 rounded-xl p-4">
          <h3 className="font-semibold text-blue-800 flex items-center gap-2 mb-2">
            <Folder className="h-5 w-5" />
            Data Location
          </h3>
          <p className="text-sm text-blue-700 font-mono break-all">{dataPath.path}</p>
          <div className="mt-3 grid grid-cols-1 gap-2 text-xs text-blue-600">
            <div className="flex items-start gap-2">
              <span className="font-medium">Project Files:</span>
              <span className="font-mono break-all">{dataPath.projects_path}</span>
            </div>
            <div className="flex items-start gap-2">
              <span className="font-medium">Reports:</span>
              <span className="font-mono break-all">{dataPath.reports_path}</span>
            </div>
          </div>
          <p className="text-xs text-blue-500 mt-3 italic">
            Tip: Drop transcript files (.txt, .docx) into project "Meetings" folders to process them automatically.
          </p>
        </div>
      )}

      {/* Logo Upload */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-slate-600 to-slate-700 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Image className="h-5 w-5" />
            Company Logo
          </h2>
        </div>
        <div className="p-6">
          <div className="flex items-center gap-6">
            <div className="flex-shrink-0">
              {logoPreview ? (
                <img
                  src={logoPreview}
                  alt="Logo preview"
                  className="h-16 w-auto max-w-32 object-contain bg-gray-100 rounded-lg p-2"
                  onError={() => setLogoPreview(null)}
                />
              ) : (
                <div className="h-16 w-32 bg-gray-100 rounded-lg flex items-center justify-center text-gray-400">
                  <Image className="h-8 w-8" />
                </div>
              )}
            </div>
            <div className="flex-1">
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleLogoUpload}
                accept="image/*"
                className="hidden"
              />
              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={uploadingLogo}
                className="flex items-center gap-2 px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg transition-colors disabled:opacity-50"
              >
                {uploadingLogo ? (
                  <Loader2 className="h-4 w-4 animate-spin" />
                ) : (
                  <Upload className="h-4 w-4" />
                )}
                {uploadingLogo ? 'Uploading...' : 'Upload Logo'}
              </button>
              <p className="text-xs text-gray-500 mt-2">
                Recommended: PNG with transparent background, 200x50px or similar aspect ratio
              </p>
            </div>
          </div>
        </div>
      </div>

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
            <div className="flex items-center justify-between mb-1">
              <label className="block text-sm font-medium text-gray-700">
                Anthropic API Key
                <a href="https://console.anthropic.com/" target="_blank" rel="noopener noreferrer"
                   className="ml-2 text-cyan-600 hover:text-cyan-700">
                  <ExternalLink className="h-3 w-3 inline" /> Get one
                </a>
              </label>
              <TestButton service="anthropic" />
            </div>
            <input
              type="password"
              value={settings.ANTHROPIC_API_KEY || ''}
              onChange={(e) => handleChange('ANTHROPIC_API_KEY', e.target.value)}
              placeholder="sk-ant-..."
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
            <TestResult service="anthropic" />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              <FolderPlus className="h-4 w-4 inline mr-1" />
              Project Names
            </label>
            <div className="space-y-2">
              {projects.map((project, index) => (
                <div key={index} className="flex items-center gap-2">
                  <span className="w-8 text-center text-sm font-medium text-gray-400">{index + 1}.</span>
                  <input
                    type="text"
                    value={project}
                    onChange={(e) => updateProject(index, e.target.value)}
                    placeholder={`Project ${index + 1} name (e.g., RH, NSD, HERC)`}
                    className="flex-1 px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
                  />
                  {projects.length > 1 && (
                    <button
                      type="button"
                      onClick={() => removeProject(index)}
                      className="p-2 text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                      title="Remove project"
                    >
                      <X className="h-4 w-4" />
                    </button>
                  )}
                </div>
              ))}
            </div>
            <button
              type="button"
              onClick={addProject}
              className="mt-3 flex items-center gap-2 px-4 py-2 text-sm text-cyan-600 hover:bg-cyan-50 rounded-lg transition-colors"
            >
              <Plus className="h-4 w-4" />
              Add Another Project
            </button>
            <p className="text-xs text-gray-500 mt-2">
              Folders will be created: 0 - Reports, 1 - [Project], 2 - [Project], etc.
              <br />Each project folder will contain: 1 - Schedules, 2 - Meetings, 3 - Deliverables
            </p>
          </div>
        </div>
      </div>

      {/* ngrok Settings */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-purple-600 to-purple-700 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <Cloud className="h-5 w-5" />
            Remote Access (ngrok) - For Power Automate
          </h2>
        </div>
        <div className="p-6">
          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="block text-sm font-medium text-gray-700">
                ngrok Auth Token
                <a href="https://dashboard.ngrok.com/get-started/your-authtoken" target="_blank" rel="noopener noreferrer"
                   className="ml-2 text-cyan-600 hover:text-cyan-700">
                  <ExternalLink className="h-3 w-3 inline" /> Get token
                </a>
              </label>
              <TestButton service="ngrok" />
            </div>
            <input
              type="password"
              value={settings.NGROK_AUTH_TOKEN || ''}
              onChange={(e) => handleChange('NGROK_AUTH_TOKEN', e.target.value)}
              placeholder="Optional - enables Power Automate to send files"
              className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-cyan-500"
            />
            <TestResult service="ngrok" />
            <p className="text-xs text-gray-500 mt-2">
              Required for Power Automate integration. After setting this, run ngrok manually or the app will start it automatically.
            </p>
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
        <div className="p-6 space-y-4">
          <div className="flex justify-end">
            <TestButton service="email" label="Test Email Connection" />
          </div>
          <TestResult service="email" />
          <div className="grid grid-cols-2 gap-4">
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
      </div>

      {/* Microsoft Sign In */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div className="px-6 py-4 bg-gradient-to-r from-blue-500 to-blue-600 text-white">
          <h2 className="text-lg font-semibold flex items-center gap-2">
            <svg className="h-5 w-5" viewBox="0 0 23 23" fill="currentColor">
              <path d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z"/>
            </svg>
            Microsoft Account (Calendar & Email)
          </h2>
        </div>
        <div className="p-6">
          {isElectron ? (
            msAccount ? (
              // Signed in state
              <div className="space-y-4">
                <div className="flex items-center justify-between p-4 bg-green-50 border border-green-200 rounded-lg">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-blue-500 rounded-full flex items-center justify-center text-white font-semibold">
                      {msAccount.name?.charAt(0) || msAccount.username?.charAt(0) || 'M'}
                    </div>
                    <div>
                      <p className="font-medium text-gray-900">{msAccount.name || 'Microsoft Account'}</p>
                      <p className="text-sm text-gray-500">{msAccount.username}</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    <CheckCircle className="h-5 w-5 text-green-500" />
                    <span className="text-sm text-green-700 font-medium">Connected</span>
                  </div>
                </div>
                <div className="flex items-center justify-between">
                  <p className="text-sm text-gray-600">
                    Access to: Calendar, Email
                  </p>
                  <button
                    onClick={handleMicrosoftSignOut}
                    className="px-4 py-2 text-sm text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                  >
                    Sign Out
                  </button>
                </div>
              </div>
            ) : (
              // Not signed in state
              <div className="space-y-4">
                <p className="text-gray-600">
                  Sign in with your Microsoft account to enable calendar integration and email features.
                </p>
                <button
                  onClick={handleMicrosoftSignIn}
                  disabled={msSigningIn}
                  className="flex items-center gap-3 px-6 py-3 bg-[#2F2F2F] hover:bg-[#1F1F1F] text-white rounded-lg transition-colors disabled:opacity-50"
                >
                  {msSigningIn ? (
                    <Loader2 className="h-5 w-5 animate-spin" />
                  ) : (
                    <svg className="h-5 w-5" viewBox="0 0 23 23" fill="currentColor">
                      <path d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z"/>
                    </svg>
                  )}
                  {msSigningIn ? 'Signing in...' : 'Sign in with Microsoft'}
                </button>
                <p className="text-xs text-gray-500">
                  A browser window will open to complete sign-in securely with Microsoft.
                </p>
              </div>
            )
          ) : (
            // Not in Electron - show manual config
            <div className="space-y-4">
              <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
                <p className="text-sm text-yellow-800">
                  Sign in with Microsoft is only available in the desktop app.
                  For web usage, configure the settings below manually.
                </p>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Client ID</label>
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
          )}
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
      <div className="flex justify-end pb-8">
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
