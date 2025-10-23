/**
 * Batch Creation Logging Utility
 * Provides structured logging for batch creation process with different log levels
 */

export interface BatchLogEntry {
  timestamp: string;
  level: 'info' | 'warn' | 'error' | 'debug' | 'success';
  step: string;
  message: string;
  data?: any;
  duration?: number;
}

export class BatchLogger {
  private logs: BatchLogEntry[] = [];
  private stepStartTimes: Map<string, number> = new Map();

  startStep(step: string, message: string, data?: any): void {
    this.stepStartTimes.set(step, Date.now());
    this.log('info', step, `Started: ${message}`, data);
  }

  endStep(step: string, message: string, data?: any): void {
    const startTime = this.stepStartTimes.get(step);
    const duration = startTime ? Date.now() - startTime : undefined;
    this.stepStartTimes.delete(step);
    this.log('success', step, `Completed: ${message}`, data, duration);
  }

  log(level: BatchLogEntry['level'], step: string, message: string, data?: any, duration?: number): void {
    const entry: BatchLogEntry = {
      timestamp: new Date().toISOString(),
      level,
      step,
      message,
      data,
      duration
    };
    
    this.logs.push(entry);
    
    // Console logging with appropriate levels
    const consoleMessage = `[${level.toUpperCase()}] ${step}: ${message}`;
    
    switch (level) {
      case 'error':
        console.error(consoleMessage, data);
        break;
      case 'warn':
        console.warn(consoleMessage, data);
        break;
      case 'debug':
        console.debug(consoleMessage, data);
        break;
      case 'success':
        console.info(`✅ ${consoleMessage}${duration ? ` (${duration}ms)` : ''}`, data);
        break;
      default:
        console.info(consoleMessage, data);
    }
    
    // Dispatch event for UI components to listen to
    try {
      window.dispatchEvent(new CustomEvent('sunbeth:batchLog', { 
        detail: entry 
      }));
    } catch (e) {
      // Ignore if window is not available
    }
  }

  error(step: string, message: string, error?: any): void {
    this.log('error', step, message, { error: error?.message || error });
  }

  warn(step: string, message: string, data?: any): void {
    this.log('warn', step, message, data);
  }

  info(step: string, message: string, data?: any): void {
    this.log('info', step, message, data);
  }

  debug(step: string, message: string, data?: any): void {
    this.log('debug', step, message, data);
  }

  success(step: string, message: string, data?: any): void {
    this.log('success', step, message, data);
  }

  getLogs(): BatchLogEntry[] {
    return [...this.logs];
  }

  getLogsForStep(step: string): BatchLogEntry[] {
    return this.logs.filter(log => log.step === step);
  }

  clearLogs(): void {
    this.logs = [];
    this.stepStartTimes.clear();
  }

  exportLogs(): string {
    return JSON.stringify(this.logs, null, 2);
  }
}

// Export singleton instance
export const batchLogger = new BatchLogger();

// Helper function to show user-friendly toast messages
import { showToast } from './alerts';

export const showBatchToast = (message: string, level: 'info' | 'success' | 'error' = 'info') => {
  try {
    showToast(message, level as any);
  } catch (e) {
    console.warn('Failed to show toast:', e);
  }
};
