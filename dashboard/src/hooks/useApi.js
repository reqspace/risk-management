import { useState, useEffect } from 'react'

// API base URL - in Electron production, Python runs on port 5001
export const API_BASE = window.location.protocol === 'file:'
  ? 'http://localhost:5001' // Electron production (file:// protocol)
  : '' // Dev mode with Vite proxy or web server

export function useApi(endpoint, options = {}) {
  const [data, setData] = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)

  useEffect(() => {
    const fetchData = async () => {
      try {
        setLoading(true)
        const response = await fetch(API_BASE + endpoint)
        const json = await response.json()

        if (json.success) {
          setData(json)
        } else {
          setError(json.error || 'Unknown error')
        }
      } catch (err) {
        setError(err.message)
      } finally {
        setLoading(false)
      }
    }

    fetchData()
  }, [endpoint])

  const refetch = async () => {
    setLoading(true)
    try {
      const response = await fetch(API_BASE + endpoint)
      const json = await response.json()
      if (json.success) {
        setData(json)
        setError(null)
      } else {
        setError(json.error || 'Unknown error')
      }
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  return { data, loading, error, refetch }
}
