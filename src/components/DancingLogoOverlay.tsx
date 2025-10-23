import React, { useEffect, useRef, useState } from 'react';
import { getBrandLogoUrl, getBrandName, getBrandPrimaryColor, getBusyOverlayShowDelayMs, getBusyOverlayMinVisibleMs } from '../utils/runtimeConfig';

/**
 * Full-screen overlay with a "dancing" brand logo used to entertain users during long operations.
 * Visibility is controlled globally via window events:
 *  - sunbeth:busy:push (detail: { label?: string })
 *  - sunbeth:busy:pop
 *  - sunbeth:busy:reset
 */
const DancingLogoOverlay: React.FC = () => {
  const [count, setCount] = useState(0);
  const [label, setLabel] = useState<string>('');
  const [visible, setVisible] = useState(false);
  const startedAtRef = useRef<number | null>(null);
  const showTimerRef = useRef<number | null>(null);
  const hideTimerRef = useRef<number | null>(null);
  const showDelayMs = getBusyOverlayShowDelayMs();      // configurable: avoid flash for very fast ops
  const minVisibleMs = getBusyOverlayMinVisibleMs();    // configurable: keep visible briefly so users notice it
  const logo = getBrandLogoUrl();
  const brand = getBrandName();
  const primary = getBrandPrimaryColor();

  useEffect(() => {
    const onPush = ((e: Event) => {
      try {
        const detail = (e as CustomEvent).detail || {};
        setLabel(String(detail.label || ''));
      } catch {}
      setCount(c => {
        const next = c + 1;
        if (next === 1) {
          // first push -> start show delay and mark start time
          startedAtRef.current = Date.now();
          // clear any pending hide
          if (hideTimerRef.current) { window.clearTimeout(hideTimerRef.current); hideTimerRef.current = null; }
          const t = window.setTimeout(() => {
            setVisible(true);
          }, showDelayMs);
          showTimerRef.current = t as any;
        }
        return next;
      });
    }) as EventListener;
    const onPop = (() => setCount(c => {
      const next = Math.max(0, c - 1);
      if (next === 0) {
        // compute how long it has been visible; delay hide if below minimum
        const since = startedAtRef.current || Date.now();
        const elapsed = Date.now() - since;
        const wait = Math.max(0, minVisibleMs - elapsed);
        // clear any pending show
        if (showTimerRef.current) { window.clearTimeout(showTimerRef.current); showTimerRef.current = null; }
        if (wait === 0) setVisible(false);
        else {
          const t = window.setTimeout(() => setVisible(false), wait);
          hideTimerRef.current = t as any;
        }
      }
      return next;
    })) as EventListener;
    const onReset = (() => setCount(0)) as EventListener;
    window.addEventListener('sunbeth:busy:push', onPush);
    window.addEventListener('sunbeth:busy:pop', onPop);
    window.addEventListener('sunbeth:busy:reset', onReset);
    return () => {
      window.removeEventListener('sunbeth:busy:push', onPush);
      window.removeEventListener('sunbeth:busy:pop', onPop);
      window.removeEventListener('sunbeth:busy:reset', onReset);
      // cleanup timers
      if (showTimerRef.current) window.clearTimeout(showTimerRef.current);
      if (hideTimerRef.current) window.clearTimeout(hideTimerRef.current);
    };
  }, []);

  if (!visible) return null;

  return (
    <div style={{ position: 'fixed', inset: 0, zIndex: 2000, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(2px)' }}>
      <style>{`
        @keyframes sunbeth-dance {
          0% { transform: translateY(0) rotate(0deg) scale(1); }
          20% { transform: translateY(-6px) rotate(-6deg) scale(1.02); }
          40% { transform: translateY(0) rotate(0deg) scale(1); }
          60% { transform: translateY(-6px) rotate(6deg) scale(1.02); }
          80% { transform: translateY(0) rotate(0deg) scale(1); }
          100% { transform: translateY(0) rotate(0deg) scale(1); }
        }
        @keyframes sunbeth-pulse {
          0% { box-shadow: 0 0 0 0 rgba(0,0,0,0.2); }
          70% { box-shadow: 0 0 0 20px rgba(0,0,0,0); }
          100% { box-shadow: 0 0 0 0 rgba(0,0,0,0); }
        }
      `}</style>
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
  <div style={{ width: 120, height: 120, borderRadius: 20, background: '#fff', display: 'grid', placeItems: 'center', animation: 'sunbeth-pulse 1.8s ease-out infinite', border: `6px solid ${primary}` }}>
          {logo ? (
            <img src={logo} alt={brand} style={{ maxWidth: '80%', maxHeight: '80%', animation: 'sunbeth-dance 1.6s ease-in-out infinite' }} />
          ) : (
            <div style={{ fontWeight: 800, color: primary, fontSize: 18, animation: 'sunbeth-dance 1.6s ease-in-out infinite' }}>{brand}</div>
          )}
        </div>
        <div style={{ color: '#fff', fontWeight: 600 }}>{label || 'Please wait...'}</div>
      </div>
    </div>
  );
};

export default DancingLogoOverlay;
