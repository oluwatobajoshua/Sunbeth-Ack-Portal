import React from 'react';
import { error } from './logger';

interface State { hasError: boolean; error?: Error | null }

export class ErrorBoundary extends React.Component<React.PropsWithChildren<{}>, State> {
  constructor(props: any) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(err: Error) {
    return { hasError: true, error: err };
  }

  componentDidCatch(err: Error, info: any) {
    error('Unhandled render error', { err, info });
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 20, background: '#fff0f0', color: '#600' }}>
          <h2>Application error</h2>
          <pre style={{ whiteSpace: 'pre-wrap' }}>{this.state.error?.message}</pre>
        </div>
      );
    }
    return this.props.children;
  }
}
