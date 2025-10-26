import React, { useEffect, useMemo, useRef } from 'react';
import ReactDOM from 'react-dom';

type ModalProps = {
  open: boolean;
  onClose: () => void;
  title?: string;
  children?: React.ReactNode;
  width?: number | string;
  maxWidth?: number | string;
};

const ensurePortalRoot = () => {
  try {
    let el = document.getElementById('sunbeth-modal-root');
    if (!el) {
      el = document.createElement('div');
      el.id = 'sunbeth-modal-root';
      document.body.appendChild(el);
    }
    return el;
  } catch {
    return null;
  }
};

// eslint-disable-next-line max-lines-per-function
const Modal: React.FC<ModalProps> = ({ open, onClose, title, children, width = 640, maxWidth = '90%' }) => {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const isSmall = useMemo(() => (typeof window !== 'undefined' ? window.innerWidth < 768 : false), []);
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open, onClose]);

  // Simple focus trap within the modal content
  useEffect(() => {
    if (!open) return;
    const root = containerRef.current;
    if (!root) return;
    const focusable = root.querySelectorAll<HTMLElement>(
      'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
    );
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (first) first.focus();
    const handler = (e: KeyboardEvent) => {
      if (e.key !== 'Tab') return;
      if (focusable.length === 0) return;
      if (e.shiftKey) {
        if (document.activeElement === first) {
          e.preventDefault();
          last?.focus();
        }
      } else {
        if (document.activeElement === last) {
          e.preventDefault();
          first?.focus();
        }
      }
    };
    root.addEventListener('keydown', handler as any);
    return () => root.removeEventListener('keydown', handler as any);
  }, [open]);

  // Prevent background scroll while modal is open
  useEffect(() => {
    if (!open) return;
    const originalOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => { document.body.style.overflow = originalOverflow; };
  }, [open]);

  if (!open) return null;
  const portalRoot = ensurePortalRoot();
  const overlay = (
    <div
      role="button"
      aria-label="Close modal backdrop"
      style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.35)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000 }}
      tabIndex={0}
      onKeyDown={(e) => { if (e.key === 'Escape' || e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onClose(); } }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      {/* eslint-disable-next-line jsx-a11y/no-noninteractive-element-interactions, jsx-a11y/no-static-element-interactions */}
      <div
        ref={containerRef}
        className="card"
        style={{
          width: isSmall ? '90vw' : width,
          maxWidth: isSmall ? '90vw' : maxWidth,
          height: isSmall ? '95vh' : 'auto',
          maxHeight: '95vh',
          overflowY: 'auto',
          padding: 16,
          borderRadius: 12
        }}
        role="dialog"
        aria-modal="true"
        aria-label={title || 'Modal'}
        tabIndex={-1}
        onMouseDown={e => e.stopPropagation()}
        onClick={e => e.stopPropagation()}
        onKeyDown={e => e.stopPropagation()}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
          <h3 style={{ margin: 0 }}>{title || 'Modal'}</h3>
          <button className="btn ghost sm" onClick={onClose} aria-label="Close">Close</button>
        </div>
        <div>{children}</div>
      </div>
    </div>
  );

  return portalRoot ? ReactDOM.createPortal(overlay, portalRoot) : overlay;
};

export default Modal;
