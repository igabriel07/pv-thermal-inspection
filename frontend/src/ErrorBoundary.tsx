import React from 'react'

type Props = {
  children: React.ReactNode
}

type State = {
  hasError: boolean
  message: string
  stack: string
}

export default class ErrorBoundary extends React.Component<Props, State> {
  state: State = { hasError: false, message: '', stack: '' }

  static getDerivedStateFromError(error: unknown): Partial<State> {
    const message = error instanceof Error ? error.message : String(error)
    return { hasError: true, message }
  }

  componentDidCatch(error: unknown) {
    const stack = error instanceof Error ? error.stack ?? '' : ''
    this.setState((prev) => ({ ...prev, stack }))
  }

  render() {
    if (!this.state.hasError) return this.props.children

    return (
      <div style={{ padding: 16, fontFamily: 'ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial' }}>
        <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 8 }}>The app crashed</div>
        <div style={{ marginBottom: 12, color: '#b91c1c' }}>{this.state.message}</div>
        {this.state.stack && (
          <pre style={{ whiteSpace: 'pre-wrap', fontSize: 12, padding: 12, borderRadius: 10, border: '1px solid #e5e7eb', background: '#f8fafc' }}>
            {this.state.stack}
          </pre>
        )}
        <div style={{ marginTop: 12, fontSize: 12, color: '#475569' }}>
          Reload the page after fixing the error.
        </div>
      </div>
    )
  }
}
