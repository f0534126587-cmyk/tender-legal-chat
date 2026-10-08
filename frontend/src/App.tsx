import { useState, useRef, useEffect, useCallback } from 'react'
import ReactMarkdown from 'react-markdown'
import './App.css'

interface Message {
  id: string
  role: 'user' | 'assistant'
  content: string
  isStreaming?: boolean
}

const WELCOME_MESSAGE: Message = {
  id: 'welcome',
  role: 'assistant',
  content: `שלום! אני עוזר משפטי המתמחה בהכנת מכרזים לרשויות מקומיות בישראל.

אני כאן לסייע לך בהכנת מכרז מקצועי, מקיף ומדויק מבחינה משפטית, בהתאם לחוק הרשויות המקומיות ותקנות המכרזים.

כדי להתחיל, אנא ספר לי:
- **סוג המכרז** - מכרז פומבי, מכרז זוטא, או פטור ממכרז?
- **הרשות המקומית** - מי מוציאה את המכרז?
- **תחום השירות** - מהו נושא המכרז?`,
}

const API_BASE = 'http://localhost:8001'

export default function App() {
  const [messages, setMessages] = useState<Message[]>([WELCOME_MESSAGE])
  const [input, setInput] = useState('')
  const [sessionId, setSessionId] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  const sendMessage = useCallback(async () => {
    if (!input.trim() || isLoading) return

    const userContent = input.trim()
    const userMsgId = `user-${Date.now()}`
    const assistantMsgId = `asst-${Date.now()}`

    setMessages(prev => [
      ...prev,
      { id: userMsgId, role: 'user', content: userContent },
      { id: assistantMsgId, role: 'assistant', content: '', isStreaming: true },
    ])
    setInput('')
    setIsLoading(true)
    setError(null)

    // Auto-resize textarea back to default
    if (textareaRef.current) {
      textareaRef.current.style.height = 'auto'
    }

    try {
      const response = await fetch(`${API_BASE}/api/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: sessionId, message: userContent }),
      })

      if (!response.ok) {
        throw new Error(`שגיאת שרת: ${response.status}`)
      }
      if (!response.body) {
        throw new Error('לא התקבלה תגובה מהשרת')
      }

      const reader = response.body.getReader()
      const decoder = new TextDecoder()
      let buffer = ''

      while (true) {
        const { done, value } = await reader.read()
        if (done) break

        buffer += decoder.decode(value, { stream: true })
        const lines = buffer.split('\n')
        buffer = lines.pop() ?? ''

        for (const line of lines) {
          if (!line.startsWith('data: ')) continue
          try {
            const data = JSON.parse(line.slice(6))

            if (data.type === 'session_id') {
              setSessionId(data.session_id)
            } else if (data.type === 'text') {
              setMessages(prev =>
                prev.map(msg =>
                  msg.id === assistantMsgId
                    ? { ...msg, content: msg.content + data.content }
                    : msg,
                ),
              )
            } else if (data.type === 'done') {
              setMessages(prev =>
                prev.map(msg =>
                  msg.id === assistantMsgId ? { ...msg, isStreaming: false } : msg,
                ),
              )
            } else if (data.type === 'error') {
              setMessages(prev =>
                prev.map(msg =>
                  msg.id === assistantMsgId
                    ? { ...msg, content: `שגיאה: ${data.message}`, isStreaming: false }
                    : msg,
                ),
              )
            }
          } catch {
            // Skip malformed SSE lines
          }
        }
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'שגיאה לא ידועה'
      setError(msg)
      setMessages(prev =>
        prev.map(msg =>
          msg.id === assistantMsgId
            ? { ...msg, content: `⚠️ שגיאה בחיבור לשרת: ${msg.content || 'אנא בדוק שהשרת פועל ונסה שנית.'}`, isStreaming: false }
            : msg,
        ),
      )
    } finally {
      setIsLoading(false)
    }
  }, [input, isLoading, sessionId])

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      sendMessage()
    }
  }

  const handleTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value)
    // Auto-resize
    e.target.style.height = 'auto'
    e.target.style.height = `${Math.min(e.target.scrollHeight, 160)}px`
  }

  const startNewChat = async () => {
    if (sessionId) {
      try {
        await fetch(`${API_BASE}/api/session/${sessionId}`, { method: 'DELETE' })
      } catch {
        // Ignore cleanup errors
      }
    }
    setSessionId(null)
    setMessages([WELCOME_MESSAGE])
    setError(null)
    setInput('')
  }

  const exportToWord = async () => {
    if (!sessionId) return

    try {
      const response = await fetch(`${API_BASE}/api/export`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: sessionId }),
      })

      if (!response.ok) {
        const err = await response.json().catch(() => ({ detail: 'שגיאה בייצוא' }))
        throw new Error(err.detail)
      }

      const blob = await response.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'tender_draft.docx'
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'שגיאה בייצוא'
      alert(`שגיאה בייצוא: ${msg}`)
    }
  }

  const hasConversation = messages.length > 1

  return (
    <div className="app" dir="rtl">
      {/* Header */}
      <header className="header">
        <div className="header-inner">
          <div className="header-brand">
            <span className="brand-icon">⚖️</span>
            <div className="brand-text">
              <h1>עוזר משפטי - מכרזים</h1>
              <p>הכנת מכרזים לרשויות מקומיות בישראל</p>
            </div>
          </div>
          <div className="header-actions">
            {hasConversation && sessionId && (
              <button className="btn btn-export" onClick={exportToWord} title="ייצא את השיחה כקובץ Word">
                <span>📄</span> ייצוא Word
              </button>
            )}
            <button className="btn btn-new" onClick={startNewChat} title="התחל שיחה חדשה">
              <span>+</span> שיחה חדשה
            </button>
          </div>
        </div>
      </header>

      {/* Error banner */}
      {error && (
        <div className="error-banner">
          <span>⚠️ {error}</span>
          <button onClick={() => setError(null)}>✕</button>
        </div>
      )}

      {/* Messages */}
      <main className="messages-area">
        <div className="messages-list">
          {messages.map(msg => (
            <div key={msg.id} className={`message message--${msg.role}`}>
              <div className="message-avatar">
                {msg.role === 'user' ? '👤' : '⚖️'}
              </div>
              <div className="message-body">
                <div className="message-role">
                  {msg.role === 'user' ? 'עורך דין' : 'עוזר משפטי'}
                </div>
                <div className="message-content">
                  {msg.role === 'assistant' ? (
                    <ReactMarkdown>{msg.content}</ReactMarkdown>
                  ) : (
                    <p>{msg.content}</p>
                  )}
                  {msg.isStreaming && <span className="typing-cursor">▊</span>}
                </div>
              </div>
            </div>
          ))}

          {isLoading && messages[messages.length - 1]?.content === '' && (
            <div className="typing-indicator">
              <span></span><span></span><span></span>
            </div>
          )}

          <div ref={messagesEndRef} />
        </div>
      </main>

      {/* Input */}
      <footer className="input-area">
        <div className="input-wrapper">
          <textarea
            ref={textareaRef}
            value={input}
            onChange={handleTextareaChange}
            onKeyDown={handleKeyDown}
            placeholder="הקלד את שאלתך כאן... (Enter לשליחה  •  Shift+Enter לשורה חדשה)"
            disabled={isLoading}
            rows={2}
            dir="rtl"
          />
          <button
            className="send-btn"
            onClick={sendMessage}
            disabled={isLoading || !input.trim()}
            title="שלח הודעה"
          >
            {isLoading ? (
              <span className="spinner" />
            ) : (
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="22" y1="2" x2="11" y2="13" />
                <polygon points="22 2 15 22 11 13 2 9 22 2" />
              </svg>
            )}
          </button>
        </div>
        <p className="input-hint">מודל: claude-opus-4-8 &nbsp;|&nbsp; תגובות בעברית בלבד</p>
      </footer>
    </div>
  )
}
