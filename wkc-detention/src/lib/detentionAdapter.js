export const DEMO_ROSTER = [
  { id: 'WKC-1042', name: 'Amelia Chen', initials: 'AC', reason: 'Late arrival', checkedIn: true, checkedInAt: Date.now() - 8 * 60000, color: 'coral' },
  { id: 'WKC-1088', name: 'Noah Williams', initials: 'NW', reason: 'Uniform', checkedIn: true, checkedInAt: Date.now() - 5 * 60000, color: 'blue' },
  { id: 'WKC-1114', name: 'Sofia Martinez', initials: 'SM', reason: 'Missed homework', checkedIn: false, checkedInAt: null, color: 'green' },
  { id: 'WKC-1176', name: 'Ethan Okafor', initials: 'EO', reason: 'Late arrival', checkedIn: false, checkedInAt: null, color: 'purple' },
  { id: 'WKC-1201', name: 'Mia Thompson', initials: 'MT', reason: 'Disruption', checkedIn: false, checkedInAt: null, color: 'gold' },
]

const config = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  databaseId: import.meta.env.VITE_FIREBASE_DATABASE_ID || '(default)',
  collection: import.meta.env.VITE_FIREBASE_COLLECTION || 'detention',
  webhookUrl: import.meta.env.VITE_DETENTION_WEBHOOK_URL,
}

const remoteEnabled = Boolean(config.apiKey && config.projectId)
let localRoster = structuredClone(DEMO_ROSTER)

export function createDataAdapter() {
  let authToken = null
  const fallbackWrite = async (operation) => {
    try {
      await operation()
    } catch (error) {
      console.warn('Remote detention adapter unavailable; keeping local demo state.', error)
    }
  }

  return {
    isRemote: remoteEnabled,
    async loadRoster() {
      if (!remoteEnabled) return structuredClone(localRoster)
      try {
        authToken = await signInAnonymously()
        const response = await fetch(firestoreDocumentUrl('roster/current'), {
          headers: { Authorization: `Bearer ${authToken}` },
        })
        if (!response.ok) throw new Error('Roster document unavailable')
        const data = await response.json()
        return data.fields?.students?.arrayValue?.values?.map(fromFirestoreValue) || structuredClone(localRoster)
      } catch (error) {
        console.warn('Firebase roster unavailable; using local demo roster.', error)
        return structuredClone(localRoster)
      }
    },
    async saveSession(session) {
      localRoster = localRoster.map((student) => ({ ...student, checkedIn: false, checkedInAt: null }))
      if (remoteEnabled) await fallbackWrite(() => writeFirestore('sessions/current', { ...session, updatedAt: Date.now() }, authToken))
    },
    async resetRoster() {
      localRoster = structuredClone(DEMO_ROSTER).map((student) => ({ ...student, checkedIn: false, checkedInAt: null }))
      if (remoteEnabled) await fallbackWrite(() => writeFirestore('roster/current', { students: localRoster }, authToken))
    },
    async checkIn(id, checkedInAt) {
      localRoster = localRoster.map((student) => student.id === id ? { ...student, checkedIn: true, checkedInAt } : student)
      if (remoteEnabled) {
        await fallbackWrite(() => writeFirestore(`checkins/${id}`, { studentId: id, checkedInAt }, authToken))
        if (config.webhookUrl) {
          await fetch(config.webhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ event: 'student.check_in', studentId: id, checkedInAt }),
          }).catch((error) => console.warn('Detention webhook unavailable.', error))
        }
      }
    },
  }
}

function firestoreDocumentUrl(path) {
  return `https://firestore.googleapis.com/v1/projects/${config.projectId}/databases/${config.databaseId}/documents/${config.collection}/${path}`
}

async function writeFirestore(path, data, authToken) {
  const fields = Object.fromEntries(Object.entries(data).map(([key, value]) => [key, toFirestoreValue(value)]))
  const headers = { 'Content-Type': 'application/json' }
  if (authToken) headers.Authorization = `Bearer ${authToken}`
  const response = await fetch(`${firestoreDocumentUrl(path)}?key=${config.apiKey}`, {
    method: 'PATCH',
    headers,
    body: JSON.stringify({ fields }),
  })
  if (!response.ok) throw new Error('Firestore write failed')
}

async function signInAnonymously() {
  const response = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signUp?key=${config.apiKey}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ returnSecureToken: true }),
  })
  if (!response.ok) throw new Error('Anonymous Firebase Auth is not enabled')
  const data = await response.json()
  return data.idToken
}

function toFirestoreValue(value) {
  if (typeof value === 'string') return { stringValue: value }
  if (typeof value === 'number') return { integerValue: String(value) }
  if (typeof value === 'boolean') return { booleanValue: value }
  if (Array.isArray(value)) return { arrayValue: { values: value.map(toFirestoreValue) } }
  if (value && typeof value === 'object') return { mapValue: { fields: Object.fromEntries(Object.entries(value).map(([key, nested]) => [key, toFirestoreValue(nested)])) } }
  return { nullValue: 'NULL_VALUE' }
}

function fromFirestoreValue(value) {
  if (value.mapValue) return Object.fromEntries(Object.entries(value.mapValue.fields || {}).map(([key, nested]) => [key, fromFirestoreValue(nested)]))
  if (value.stringValue !== undefined) return value.stringValue
  if (value.integerValue !== undefined) return Number(value.integerValue)
  if (value.booleanValue !== undefined) return value.booleanValue
  if (value.arrayValue) return (value.arrayValue.values || []).map(fromFirestoreValue)
  return null
}

export function getSessionCode(session) {
  return session?.id ? `WKC-${session.id.slice(-4).toUpperCase()}` : 'WKC-2025'
}

export function formatTimestamp(timestamp) {
  if (!timestamp) return '—'
  return new Intl.DateTimeFormat('en', { hour: 'numeric', minute: '2-digit' }).format(timestamp)
}
