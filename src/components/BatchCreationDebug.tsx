import React, { useState, useEffect, useRef } from 'react';
import { BatchLogEntry, batchLogger } from '../utils/batchLogger';

interface BatchCreationDebugProps {
  isVisible: boolean;
  onClose: () => void;
}

const BatchCreationDebug: React.FC<BatchCreationDebugProps> = ({ isVisible, onClose }) => {
  const [logs, setLogs] = useState<BatchLogEntry[]>([]);
  const [filter, setFilter] = useState<'all' | 'info' | 'warn' | 'error' | 'success' | 'debug'>('all');
  const [stepFilter, setStepFilter] = useState<string>('all');
  const [autoScroll, setAutoScroll] = useState(true);
  const logsEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleBatchLog = (event: CustomEvent<BatchLogEntry>) => {
      setLogs(prevLogs => [...prevLogs, event.detail]);
    };

    window.addEventListener('sunbeth:batchLog', handleBatchLog as EventListener);
    
    // Load existing logs
    setLogs(batchLogger.getLogs());

    return () => {
      window.removeEventListener('sunbeth:batchLog', handleBatchLog as EventListener);
    };
  }, []);

  useEffect(() => {
    if (autoScroll && logsEndRef.current) {
      logsEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs, autoScroll]);

  const filteredLogs = logs.filter(log => {
    if (filter !== 'all' && log.level !== filter) return false;
    if (stepFilter !== 'all' && log.step !== stepFilter) return false;
    return true;
  });

  const uniqueSteps = Array.from(new Set(logs.map(log => log.step))).sort();

  const getLevelColor = (level: BatchLogEntry['level']) => {
    switch (level) {
      case 'error': return '#ff4444';
      case 'warn': return '#ffaa00';
      case 'success': return '#44aa44';
      case 'debug': return '#888888';
      default: return '#4488ff';
    }
  };

  const getLevelIcon = (level: BatchLogEntry['level']) => {
    switch (level) {
      case 'error': return '❌';
      case 'warn': return '⚠️';
      case 'success': return '✅';
      case 'debug': return '🔍';
      default: return 'ℹ️';
    }
  };

  const clearLogs = () => {
    setLogs([]);
    batchLogger.clearLogs();
  };

  const exportLogs = () => {
    const dataStr = JSON.stringify(logs, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `batch-creation-logs-${new Date().toISOString()}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };

  if (!isVisible) return null;

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(0,0,0,0.7)',
      zIndex: 10000,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '20px'
    }}>
      <div style={{
        backgroundColor: '#1e1e1e',
        color: '#ffffff',
        borderRadius: '8px',
        padding: '20px',
        width: '90%',
        maxWidth: '1200px',
        height: '80%',
        display: 'flex',
        flexDirection: 'column',
        boxShadow: '0 4px 20px rgba(0,0,0,0.5)'
      }}>
        {/* Header */}
        <div style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: '20px',
          paddingBottom: '10px',
          borderBottom: '1px solid #333'
        }}>
          <h2 style={{ margin: 0, color: '#ffffff' }}>Batch Creation Debug Console</h2>
          <div style={{ display: 'flex', gap: '10px' }}>
            <button
              onClick={clearLogs}
              style={{
                padding: '8px 16px',
                backgroundColor: '#ff4444',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer'
              }}
            >
              Clear Logs
            </button>
            <button
              onClick={exportLogs}
              style={{
                padding: '8px 16px',
                backgroundColor: '#4488ff',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer'
              }}
            >
              Export Logs
            </button>
            <button
              onClick={onClose}
              style={{
                padding: '8px 16px',
                backgroundColor: '#666',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer'
              }}
            >
              Close
            </button>
          </div>
        </div>

        {/* Filters */}
        <div style={{
          display: 'flex',
          gap: '15px',
          marginBottom: '15px',
          alignItems: 'center'
        }}>
          <div>
            <label style={{ marginRight: '8px' }}>Level:</label>
            <select
              value={filter}
              onChange={(e) => setFilter(e.target.value as any)}
              style={{
                padding: '4px 8px',
                backgroundColor: '#333',
                color: 'white',
                border: '1px solid #555',
                borderRadius: '4px'
              }}
            >
              <option value="all">All</option>
              <option value="info">Info</option>
              <option value="success">Success</option>
              <option value="warn">Warning</option>
              <option value="error">Error</option>
              <option value="debug">Debug</option>
            </select>
          </div>

          <div>
            <label style={{ marginRight: '8px' }}>Step:</label>
            <select
              value={stepFilter}
              onChange={(e) => setStepFilter(e.target.value)}
              style={{
                padding: '4px 8px',
                backgroundColor: '#333',
                color: 'white',
                border: '1px solid #555',
                borderRadius: '4px'
              }}
            >
              <option value="all">All Steps</option>
              {uniqueSteps.map(step => (
                <option key={step} value={step}>{step}</option>
              ))}
            </select>
          </div>

          <div style={{ marginLeft: 'auto' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <input
                type="checkbox"
                checked={autoScroll}
                onChange={(e) => setAutoScroll(e.target.checked)}
              />
              Auto-scroll
            </label>
          </div>

          <div style={{ color: '#ccc' }}>
            Showing {filteredLogs.length} of {logs.length} logs
          </div>
        </div>

        {/* Logs Container */}
        <div style={{
          flex: 1,
          backgroundColor: '#0d1117',
          border: '1px solid #333',
          borderRadius: '4px',
          padding: '10px',
          overflow: 'auto',
          fontFamily: 'Monaco, Consolas, "Courier New", monospace',
          fontSize: '13px',
          lineHeight: '1.4'
        }}>
          {filteredLogs.length === 0 ? (
            <div style={{ color: '#888', textAlign: 'center', padding: '20px' }}>
              No logs to display
            </div>
          ) : (
            filteredLogs.map((log, index) => (
              <div
                key={index}
                style={{
                  marginBottom: '8px',
                  padding: '8px',
                  backgroundColor: log.level === 'error' ? 'rgba(255,68,68,0.1)' : 
                                  log.level === 'warn' ? 'rgba(255,170,0,0.1)' : 
                                  log.level === 'success' ? 'rgba(68,170,68,0.1)' : 
                                  'rgba(68,136,255,0.05)',
                  border: `1px solid ${getLevelColor(log.level)}20`,
                  borderRadius: '4px',
                  borderLeft: `4px solid ${getLevelColor(log.level)}`
                }}
              >
                <div style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'flex-start',
                  marginBottom: '4px'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span>{getLevelIcon(log.level)}</span>
                    <span style={{ color: getLevelColor(log.level), fontWeight: 'bold' }}>
                      {log.level.toUpperCase()}
                    </span>
                    <span style={{ color: '#888', fontSize: '11px' }}>
                      [{log.step}]
                    </span>
                    {log.duration && (
                      <span style={{ color: '#888', fontSize: '11px' }}>
                        ({log.duration}ms)
                      </span>
                    )}
                  </div>
                  <span style={{ color: '#666', fontSize: '11px' }}>
                    {new Date(log.timestamp).toLocaleTimeString()}
                  </span>
                </div>
                <div style={{ color: '#e6e6e6', marginBottom: '4px' }}>
                  {log.message}
                </div>
                {log.data && (
                  <details style={{ marginTop: '8px' }}>
                    <summary style={{ 
                      color: '#888', 
                      cursor: 'pointer', 
                      fontSize: '11px',
                      userSelect: 'none'
                    }}>
                      Show Data
                    </summary>
                    <pre style={{
                      marginTop: '8px',
                      padding: '8px',
                      backgroundColor: '#161b22',
                      border: '1px solid #333',
                      borderRadius: '4px',
                      overflow: 'auto',
                      fontSize: '11px',
                      color: '#c9d1d9'
                    }}>
                      {JSON.stringify(log.data, null, 2)}
                    </pre>
                  </details>
                )}
              </div>
            ))
          )}
          <div ref={logsEndRef} />
        </div>
      </div>
    </div>
  );
};

export default BatchCreationDebug;