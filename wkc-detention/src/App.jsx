import { useEffect, useMemo, useState } from 'react'
import {
  DEMO_ROSTER,
  createDataAdapter,
  formatTimestamp,
  getSessionCode,
} from './lib/detentionAdapter'
import './App.css'

const SESSION_MINUTES = 30

function App() {
  const adapter = useMemo(() => createDataAdapter(), [])
  const [roster, setRoster] = useState(DEMO_ROSTER)
  const [session, setSession] = useState(null)
  const [view, setView] = useState('teacher')
  const [studentId, setStudentId] = useState('')
  const [toast, setToast] = useState(null)
  const [loading, setLoading] = useState(true)
  const [syncing, setSyncing] = useState(false)

  const notify = (message, tone = 'success') => {
    setToast({ message, tone })
    window.setTimeout(() => setToast(null), 3200)
  }

  useEffect(() => {
    let active = true
    adapter.loadRoster().then((items) => {
      if (active) {
        setRoster(items)
        setLoading(false)
      }
    })
    return () => {
      active = false
    }
  }, [adapter])

  useEffect(() => {
    if (!session?.id) return undefined
    const timer = window.setInterval(() => {
      setSession((current) => {
        if (!current) return current
        const remaining = Math.max(0, current.endsAt - Date.now())
        return { ...current, remaining }
      })
    }, 1000)
    return () => window.clearInterval(timer)
  }, [session?.id])

  const checkedInCount = roster.filter((student) => student.checkedIn).length
  const remainingMinutes = session ? Math.ceil(session.remaining / 60000) : 0
  const sessionActive = Boolean(session && session.remaining > 0)
  const sessionCode = session ? getSessionCode(session) : 'WKC-2025'

  const startSession = async () => {
    const nextSession = {
      id: `session-${Date.now()}`,
      startedAt: Date.now(),
      endsAt: Date.now() + SESSION_MINUTES * 60 * 1000,
      remaining: SESSION_MINUTES * 60 * 1000,
    }
    setSession(nextSession)
    setRoster((items) => items.map((student) => ({ ...student, checkedIn: false, checkedInAt: null })))
    setSyncing(true)
    await adapter.saveSession(nextSession)
    setSyncing(false)
    notify('Session started. Students can check in now.')
  }

  const resetSession = async () => {
    setSession(null)
    setRoster((items) => items.map((student) => ({ ...student, checkedIn: false, checkedInAt: null })))
    await adapter.resetRoster()
    notify('Session reset. The roster is ready for a new group.', 'info')
  }

  const checkIn = async (event) => {
    event.preventDefault()
    const normalizedId = studentId.trim().toUpperCase()
    const student = roster.find((item) => item.id === normalizedId)
    if (!student) {
      notify('We could not find that student ID. Please try again.', 'error')
      return
    }
    if (!sessionActive) {
      notify('Ask your teacher to start the detention session first.', 'error')
      return
    }
    const checkedInAt = Date.now()
    setRoster((items) =>
      items.map((item) =>
        item.id === normalizedId ? { ...item, checkedIn: true, checkedInAt } : item,
      ),
    )
    setStudentId('')
    await adapter.checkIn(normalizedId, checkedInAt)
    notify(`${student.name} is checked in.`)
  }

  const qrData = `${window.location.origin}/?session=${sessionCode}`
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=220x220&margin=12&data=${encodeURIComponent(qrData)}`

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark" aria-hidden="true">W</div>
          <div>
            <p className="eyebrow">WKC student services</p>
            <h1>Detention desk</h1>
          </div>
        </div>
        <div className="topbar-actions">
          <div className={`connection-pill ${adapter.isRemote ? 'remote' : 'demo'}`}>
            <span className="status-dot" aria-hidden="true" />
            {adapter.isRemote ? 'Firestore connected' : 'Local demo mode'}
          </div>
          <button className="avatar" type="button" aria-label="Account menu">TS</button>
        </div>
      </header>

      <main className="main-content">
        <section className="hero-row">
          <div>
            <p className="eyebrow accent">Thursday · 14 November 2025</p>
            <h2>Keep every session<br /><span>calm & accountable.</span></h2>
            <p className="hero-copy">A simple check-in desk for teachers, with a clear path for every student.</p>
          </div>
          <div className="view-switcher" role="tablist" aria-label="Choose portal view">
            <button type="button" className={view === 'teacher' ? 'active' : ''} onClick={() => setView('teacher')} role="tab" aria-selected={view === 'teacher'}>Teacher desk</button>
            <button type="button" className={view === 'student' ? 'active' : ''} onClick={() => setView('student')} role="tab" aria-selected={view === 'student'}>Student check-in</button>
          </div>
        </section>

        <div className="demo-banner" role="status">
          <span className="sparkle" aria-hidden="true">✦</span>
          <span><strong>Local demo form fallback</strong> — changes stay in this browser when Firebase credentials are not configured.</span>
          <button type="button" onClick={() => notify('Connect Firebase by adding VITE_FIREBASE_API_KEY and VITE_FIREBASE_PROJECT_ID.', 'info')}>How to connect <span aria-hidden="true">↗</span></button>
        </div>

        {view === 'teacher' ? (
          <TeacherView
            roster={roster}
            checkedInCount={checkedInCount}
            session={session}
            sessionActive={sessionActive}
            remainingMinutes={remainingMinutes}
            sessionCode={sessionCode}
            qrUrl={qrUrl}
            loading={loading}
            syncing={syncing}
            onStart={startSession}
            onReset={resetSession}
            onStudentView={() => setView('student')}
          />
        ) : (
          <StudentView
            studentId={studentId}
            setStudentId={setStudentId}
            onSubmit={checkIn}
            session={session}
            sessionActive={sessionActive}
            remainingMinutes={remainingMinutes}
            sessionCode={sessionCode}
            checkedInCount={checkedInCount}
            roster={roster}
          />
        )}
      </main>
      <footer className="footer">
        <span>WKC detention desk</span>
        <span>Built for quiet accountability · v1.0</span>
      </footer>
      {toast && <div className={`toast ${toast.tone}`} role="alert"><span aria-hidden="true">{toast.tone === 'error' ? '!' : '✓'}</span>{toast.message}</div>}
    </div>
  )
}

function TeacherView({ roster, checkedInCount, session, sessionActive, remainingMinutes, sessionCode, qrUrl, loading, syncing, onStart, onReset, onStudentView }) {
  return (
    <>
      <section className="stats-grid" aria-label="Session overview">
        <div className={`stat-card primary-stat ${sessionActive && remainingMinutes <= 5 ? 'urgent-stat' : ''}`}>
          <div className="stat-label">Today&apos;s session</div>
          <div className="stat-value">{sessionActive ? `${String(Math.floor(session.remaining / 60000)).padStart(2, '0')}:${String(Math.floor((session.remaining % 60000) / 1000)).padStart(2, '0')}` : '—'}</div>
          <div className="stat-caption">{sessionActive ? 'remaining · live timer' : 'No active session'}</div>
          <span className={`session-status ${sessionActive ? 'live' : ''}`}><span className="status-dot" />{sessionActive ? 'In progress' : 'Ready to start'}</span>
        </div>
        <div className="stat-card">
          <div className="stat-label">Students present</div>
          <div className="stat-value">{checkedInCount}<small>/{roster.length}</small></div>
          <div className="progress-track"><span style={{ width: `${roster.length ? (checkedInCount / roster.length) * 100 : 0}%` }} /></div>
          <div className="stat-caption">{checkedInCount === roster.length && roster.length ? 'Everyone is checked in' : `${roster.length - checkedInCount} still to check in`}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Session code</div>
          <div className="stat-value code-value">{sessionCode}</div>
          <div className="stat-caption">Share with students at the door</div>
        </div>
      </section>

      <section className="workspace-grid">
        <div className="panel roster-panel">
          <div className="panel-heading">
            <div><p className="eyebrow">Live roster</p><h3>Who&apos;s in the room?</h3></div>
            <span className="live-tag"><span className="status-dot" />Live</span>
          </div>
          <div className="roster-list">
            {loading ? <div className="empty-state">Loading roster…</div> : roster.map((student) => (
              <div className="student-row" key={student.id}>
                <div className={`student-avatar avatar-${student.color}`}>{student.initials}</div>
                <div className="student-details"><strong>{student.name}</strong><span>{student.id} · {student.reason}</span></div>
                <div className="check-status">{student.checkedIn ? <><span className="check-icon">✓</span><span>{formatTimestamp(student.checkedInAt)}</span></> : <span className="not-here">Not checked in</span>}</div>
              </div>
            ))}
          </div>
          <button className="text-button" type="button" onClick={onStudentView}>Open student check-in <span aria-hidden="true">→</span></button>
        </div>
        <aside className="side-column">
          <div className="panel control-panel">
            <div className="panel-heading compact"><div><p className="eyebrow">Session control</p><h3>{sessionActive ? 'Session is live' : 'Start a session'}</h3></div><span className="control-icon" aria-hidden="true">◷</span></div>
            <p className="panel-copy">{sessionActive ? 'The countdown is synchronized locally and will highlight when five minutes remain.' : 'Open the door when you are ready. A session runs for 30 minutes.'}</p>
            <button className="primary-button full" type="button" onClick={onStart} disabled={syncing}>{syncing ? 'Saving…' : sessionActive ? 'Restart 30-minute session' : 'Start 30-minute session'} <span aria-hidden="true">→</span></button>
            {session && <button className="secondary-button full" type="button" onClick={onReset}>Reset session</button>}
          </div>
          <div className="panel qr-panel">
            <div><p className="eyebrow">Quick check-in</p><h3>Scan to check in</h3></div>
            <img className="qr-code" src={qrUrl} alt={`QR code for session ${sessionCode}`} />
            <p>Students scan this code, then enter their ID. <strong>{sessionCode}</strong></p>
          </div>
        </aside>
      </section>
    </>
  )
}

function StudentView({ studentId, setStudentId, onSubmit, session, sessionActive, remainingMinutes, sessionCode, checkedInCount, roster }) {
  const isUrgent = sessionActive && remainingMinutes <= 5
  const checkedInStudent = roster.find((student) => student.id === studentId.trim().toUpperCase() && student.checkedIn)
  return (
    <section className="student-layout">
      <div className="student-checkin-card panel">
        <div className="student-card-top"><span className="round-icon">✓</span><span className="eyebrow accent">Student portal</span></div>
        <h3>Welcome in.</h3>
        <p className="student-intro">Enter your student ID to record your arrival. Your teacher will see your check-in immediately.</p>
        <form onSubmit={onSubmit} className="checkin-form">
          <label htmlFor="student-id">Student ID</label>
          <div className="input-wrap"><span aria-hidden="true">#</span><input id="student-id" value={studentId} onChange={(event) => setStudentId(event.target.value)} placeholder="e.g. WKC-1042" autoComplete="off" /></div>
          <button className="primary-button full" type="submit" disabled={!sessionActive}>Check me in <span aria-hidden="true">→</span></button>
        </form>
        {!sessionActive && <div className="notice warning"><span aria-hidden="true">!</span><span>Your teacher hasn&apos;t started a session yet. Please wait for the session code.</span></div>}
        {checkedInStudent && <div className="notice success"><span aria-hidden="true">✓</span><span>You&apos;re checked in, {checkedInStudent.name.split(' ')[0]}. Find a quiet seat.</span></div>}
        <p className="privacy-note">Your check-in is shared only with your teacher for this session.</p>
      </div>
      <div className="student-side">
        <div className={`countdown-card ${isUrgent ? 'urgent' : ''}`}>
          <p className="eyebrow">Session {sessionCode}</p>
          <div className="countdown">{sessionActive ? `${String(Math.floor(session.remaining / 60000)).padStart(2, '0')}:${String(Math.floor((session.remaining % 60000) / 1000)).padStart(2, '0')}` : '—'}</div>
          <p>{isUrgent ? 'Five minutes left · finish strong' : sessionActive ? 'Time remaining · stay focused' : 'Waiting for teacher to begin'}</p>
        </div>
        <div className="panel student-progress"><p className="eyebrow">Room progress</p><h3>{checkedInCount} students checked in</h3><div className="progress-track"><span style={{ width: `${roster.length ? (checkedInCount / roster.length) * 100 : 0}%` }} /></div><p className="muted">Your teacher sees this roster live.</p></div>
      </div>
    </section>
  )
}

export default App
