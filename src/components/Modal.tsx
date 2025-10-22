import React, { useEffect, useMemo, useRef } from 'react';

type ModalProps = {
  open: boolean;
  onClose: () => void;
  title?: string;
  children?: React.ReactNode;
  width?: number | string;
  maxWidth?: number | string;
};

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

  if (!open) return null;
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.35)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000 }} onClick={onClose}>
      <div
        ref={containerRef}
        className="card"
        style={{
          width: isSmall ? '100%' : width,
          maxWidth: isSmall ? '100%' : maxWidth,
          height: isSmall ? '95vh' : 'auto',
          maxHeight: '95vh',
          overflowY: 'auto',
          padding: 16,
        }}
        onClick={e => e.stopPropagation()}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
          <h3 style={{ margin: 0 }}>{title || 'Modal'}</h3>
          <button className="btn ghost sm" onClick={onClose} aria-label="Close">Close</button>
        </div>
        <div>{children}</div>
      </div>
    </div>
  );
};

export default Modal;
