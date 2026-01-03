import { BrowserRouter as Router, Routes, Route } from 'react-router-dom'
import Layout from './components/Layout'
import Portfolio from './pages/Portfolio'
import Dashboard from './pages/Dashboard'
import RiskRegister from './pages/RiskRegister'
import Tasks from './pages/Tasks'
import Reports from './pages/Reports'
import Settings from './pages/Settings'

function App() {
  return (
    <Router>
      <Layout>
        <Routes>
          <Route path="/" element={<Portfolio />} />
          <Route path="/settings" element={<Settings />} />
          <Route path="/project/:projectCode" element={<Dashboard />} />
          <Route path="/project/:projectCode/risks" element={<RiskRegister />} />
          <Route path="/project/:projectCode/tasks" element={<Tasks />} />
          <Route path="/project/:projectCode/reports" element={<Reports />} />
          <Route path="/risks" element={<RiskRegister />} />
          <Route path="/tasks" element={<Tasks />} />
          <Route path="/reports" element={<Reports />} />
        </Routes>
      </Layout>
    </Router>
  )
}

export default App
